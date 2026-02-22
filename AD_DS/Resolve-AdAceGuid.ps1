<#
.SYNOPSIS
    Resolves GUIDs found on Active Directory ACEs to their friendly attribute, class,
    or extended right names.

.DESCRIPTION
    Active Directory Access Control Entries (ACEs) reference schema attributes, classes,
    extended rights, property sets, and validated writes by GUID. These GUIDs are opaque
    and not human-readable, making ACL auditing difficult.

    This script provides two functions:

        Build-GuidMap             Queries the AD schema and Extended-Rights container once,
                                  returning a hashtable that maps every known GUID to its
                                  friendly name. Only two LDAP queries are made regardless
                                  of how many GUIDs you need to resolve.

        Convert-GuidToAttribute   Performs a local hashtable lookup to resolve a single GUID
                                  to its friendly name. No LDAP round-trip per call.

    Typical workflow:
        1. Dot-source this script to load both functions
        2. Call Build-GuidMap once to create the lookup table
        3. Use Convert-GuidToAttribute in calculated properties when examining ACEs

    Three GUID sources are covered:
        - Schema attributes (schemaIDGUID on attributeSchema objects)
        - Schema classes (schemaIDGUID on classSchema objects)
        - Extended rights, property sets, and validated writes (rightsGuid in CN=Extended-Rights)

.PARAMETER GUID
    (Convert-GuidToAttribute) The GUID string to resolve, typically from the ObjectType
    or InheritedObjectType property on an ACE.

.PARAMETER GuidMap
    (Convert-GuidToAttribute) A pre-built hashtable from Build-GuidMap mapping GUID strings
    to friendly names.

.EXAMPLE
    . .\Resolve-AdAceGuid.ps1
    $map = Build-GuidMap
    Convert-GuidToAttribute -GUID '5b47d60f-6090-40b2-9f37-2a4de88f3063' -GuidMap $map

    Dot-source the script, build the map, and resolve a single GUID.

.EXAMPLE
    . .\Resolve-AdAceGuid.ps1
    $map = Build-GuidMap

    $Props = @(
        '*'
        @{n = 'ObjectType_Friendly';         e = { (Convert-GuidToAttribute $_.ObjectType          $map).Name }}
        @{n = 'InheritedObjectType_Friendly'; e = { (Convert-GuidToAttribute $_.InheritedObjectType $map).Name }}
    )

    $acl = Get-Acl "AD:\OU=MyOU,DC=corp,DC=example,DC=com"
    $acl.Access | Where-Object ActiveDirectoryRights -eq WriteProperty | Select-Object $Props

    Resolve ObjectType and InheritedObjectType GUIDs on WriteProperty ACEs for an OU.

.EXAMPLE
    . .\Resolve-AdAceGuid.ps1
    $map = Build-GuidMap

    $acl = Get-Acl "AD:\OU=Servers,DC=corp,DC=example,DC=com"
    $acl.Access |
        Select-Object IdentityReference, ActiveDirectoryRights,
            @{n = 'ObjectType_Friendly'; e = { (Convert-GuidToAttribute $_.ObjectType $map).Name }} |
        Where-Object IdentityReference -match 'ServiceAccounts' |
        Format-Table -AutoSize

    Audit which attributes a specific group can write to on the Servers OU.

.NOTES
    Author: Mike Crowley
    https://mikecrowley.us

    Requires: ActiveDirectory module (RSAT)

    The all-zeros GUID (00000000-0000-0000-0000-000000000000) means "applies to everything"
    and is returned as "(All)" rather than querying the schema.

    Performance:
        Build-GuidMap makes exactly two LDAP queries (schema + extended rights), regardless
        of tenant size. Convert-GuidToAttribute is a local hashtable lookup with no network
        calls. This design avoids the common pitfall of making an LDAP round-trip per ACE,
        which becomes very slow when resolving ACLs with dozens or hundreds of entries.

    Output object properties:
        Name         - Friendly name of the schema attribute, class, or extended right
        SchemaIDGUID - Normalized GUID string (lowercase, with dashes)

.LINK
    https://learn.microsoft.com/en-us/windows/win32/adschema/attributes-all

.LINK
    https://github.com/Mike-Crowley/Public-Scripts
#>

#Requires -Modules ActiveDirectory

function Convert-GuidToAttribute {
    <#
        .SYNOPSIS
            Resolves GUIDs found on AD ACEs to their friendly attribute, class, or extended right names.
            Covers three sources:
              - Schema attributes and classes (schemaIDGUID)
              - Extended rights, property sets, and validated writes (rightsGuid in CN=Extended-Rights)
            Builds an in-memory hashtable on first call, then does local lookups on subsequent calls.
            This avoids an LDAP round-trip per GUID, which matters when resolving ACLs with dozens of ACEs.

        .PARAMETER GUID
            The GUID string to resolve (typically from ObjectType or InheritedObjectType on an ACE).

        .PARAMETER GuidMap
            A pre-built hashtable mapping GUID strings to friendly names.
            Pass this explicitly to avoid rebuilding the map on every call.

        .EXAMPLE
            $map = Build-GuidMap
            Convert-GuidToAttribute -GUID '5b47d60f-6090-40b2-9f37-2a4de88f3063' -GuidMap $map

        .EXAMPLE
            $map = Build-GuidMap
            $aces | Select-Object *, @{n='ObjectType_Friendly'; e={ Convert-GuidToAttribute $_.ObjectType $map }}
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory, Position = 0)]
        [string]$GUID,

        [Parameter(Mandatory, Position = 1)]
        [hashtable]$GuidMap
    )

    # Normalize to consistent string format (lowercase, dashes)
    $normalizedGuid = ([guid]$GUID).Guid

    # The all-zeros GUID means "applies to everything" — not a real schema object
    if ($normalizedGuid -eq '00000000-0000-0000-0000-000000000000') {
        return [pscustomobject]@{
            Name         = '(All)'
            SchemaIDGUID = $normalizedGuid
        }
    }

    $friendlyName = $GuidMap[$normalizedGuid]

    [pscustomobject]@{
        Name         = if ($friendlyName) { $friendlyName } else { '(Unresolved)' }
        SchemaIDGUID = $normalizedGuid
    }
}

function Build-GuidMap {
    <#
        .SYNOPSIS
            Queries the schema and Extended-Rights containers once, returning a hashtable
            that maps every known GUID to its friendly name.

        .DESCRIPTION
            Two sources are merged:
              1. Schema naming context — attributes and object classes (schemaIDGUID, stored as byte[])
              2. CN=Extended-Rights in the Configuration partition — extended rights,
                 property sets, and validated writes (rightsGuid, stored as a string)

        .EXAMPLE
            $map = Build-GuidMap
            $map['5b47d60f-6090-40b2-9f37-2a4de88f3063']
            # returns: ms-DS-Key-Credential-Link
    #>
    [CmdletBinding()]
    param ()

    $rootDSE = Get-ADRootDSE
    $guidMap = @{}

    # --- Schema attributes and classes (schemaIDGUID is a byte array) ---
    $schemaSearchBase = $rootDSE.schemaNamingContext
    Write-Verbose "Querying schema at $schemaSearchBase ..."

    Get-ADObject -SearchBase $schemaSearchBase -Filter * -Properties name, schemaIDGUID |
        Where-Object { $_.schemaIDGUID } | ForEach-Object {
            $key = ([guid]$_.schemaIDGUID).Guid
            $guidMap[$key] = $_.name
        }

    Write-Verbose "Schema entries loaded: $($guidMap.Count)"

    # --- Extended rights, property sets, validated writes (rightsGuid is already a string) ---
    $extendedRightsBase = "CN=Extended-Rights," + $rootDSE.configurationNamingContext
    Write-Verbose "Querying extended rights at $extendedRightsBase ..."

    $beforeCount = $guidMap.Count

    Get-ADObject -SearchBase $extendedRightsBase -Filter * -Properties name, rightsGuid |
        Where-Object { $_.rightsGuid } | ForEach-Object {
            # rightsGuid is a string, but normalize it through [guid] for consistent casing
            $key = ([guid]$_.rightsGuid).Guid
            $guidMap[$key] = $_.name
        }

    Write-Verbose "Extended rights entries loaded: $($guidMap.Count - $beforeCount)"
    Write-Verbose "Total GUID map entries: $($guidMap.Count)"

    $guidMap
}
