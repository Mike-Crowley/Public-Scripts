<#
.SYNOPSIS
    Reports on Active Directory integrated DNS records by reading dnsNode objects straight from
    the directory, flagging stale dynamic records, orphaned or mismatched owners, and the
    permission problems that quietly break DNS scavenging or leave a name open to hijacking.
    Read-only, and needs no access to the DNS server itself.

.DESCRIPTION
    Every AD integrated DNS zone lives in the directory as a container of dnsNode objects, and
    each dnsNode carries its resource records in the multi-valued binary attribute dnsRecord.
    That blob holds the record type, the TTL, and the aging timestamp the DNS server uses to
    decide whether a record is stale. The directory also holds two things the DNS Server API
    never exposes: the security descriptor on each individual record, and the SID of the
    principal that owns it.

    This script reads all of that straight from the directory. It does not talk to the DNS
    Server RPC/WMI interface and does not need the DnsServer RSAT module, which matters in
    practice: read access to Active Directory is common, while access to the DNS server
    management interface usually is not. The tradeoff is that anything the DNS service computes at
    runtime (effective scavenging state, zone aging configuration, non-AD-integrated zones,
    the server cache) is out of scope here. Use Get-DnsServerResourceRecord when you do have
    server access and only need record data.

    For each record the report answers three questions:

        Is it stale?        Dynamically registered records carry a timestamp that the DNS
                            server refreshes as the host re-registers. A record whose
                            timestamp has not moved in months is almost always a host that no
                            longer exists. Static records carry a timestamp of 0 and are never
                            scavenged, so they are excluded by default (-IncludeStatic).

        Who owns it?        Under secure dynamic update the computer that first registers a
                            name becomes the owner of the dnsNode object. When the owner is
                            not a computer account matching the record name (because the
                            record was created by hand, created by DHCP, or the original host
                            was rebuilt), the current host generally cannot refresh or delete
                            its own record.

        Can the owner
        still write to it?  An owner with no write rights on its own record cannot update it,
                            so the record ages out or goes stale in place. Rights are read
                            from the object's security descriptor and OR-combined across every
                            Allow ACE granted to the owner SID.

    Records are flagged in RequiresReview with a human-readable ReviewReasons string, and the
    Excel report color-codes those rows and adds a small dashboard.

.PARAMETER ZoneName
    One or more zone names to report on. Wildcards are supported, so -ZoneName * reports on
    every AD integrated zone the target DC hosts. If omitted, the script prompts with a picker
    (Out-GridView where available, otherwise a numbered console menu). In a non-interactive
    session, omitting this parameter is an error rather than a hang.

.PARAMETER ListZones
    List the AD integrated zones that can be found and exit without querying any records.
    Use this to discover zone names for -ZoneName.

.PARAMETER Server
    Domain controller to query, optionally as host:port. Defaults to whatever the
    ActiveDirectory module discovers.

    Do NOT point this at a global catalog port (3268/3269). dnsRecord and nTSecurityDescriptor
    are not in the partial attribute set, and DNS application partitions are not replicated to
    global catalogs, so a GC returns incomplete data or nothing at all.

.PARAMETER Credential
    Credential for the directory queries. Defaults to the current user.

.PARAMETER RecordType
    Resource record types to include. Defaults to A, which is where dynamic registration and
    scavenging problems actually show up. Use All to report on every type in the zone.

.PARAMETER StaleThresholdDays
    A dynamic record whose timestamp is older than this many days is marked stale. Default 180.

    Pick this deliberately. A record is only genuinely scavengeable after the no-refresh
    interval plus the refresh interval have both elapsed (7 + 7 = 14 days with the defaults),
    so anything beyond a few weeks is conservative. 180 days is a "this host is gone" bar
    rather than a "this is scavengeable" bar.

.PARAMETER IncludeStatic
    Include statically created records (timestamp 0). These are never scavenged and are
    excluded by default, but they are worth reviewing if you are hunting for records that
    should have been dynamic.

.PARAMETER IncludeTombstoned
    Include dnsNode objects marked dNSTombstoned. These are deleted records awaiting cleanup
    and are excluded by default. A tombstone is stored as record type 0 with a zero timestamp,
    so when this switch is set those nodes bypass the -RecordType and -IncludeStatic filters
    that would otherwise discard every one of them.

.PARAMETER MaxRecordsPerZone
    Cap the number of dnsNode objects retrieved per zone. 0 (the default) means no cap. Useful
    for a quick look at a very large zone; the script warns when the cap truncates results.

.PARAMETER OutputFolder
    Folder for the report file. Defaults to the current user's Desktop, resolved via
    [Environment]::GetFolderPath so it follows OneDrive Known Folder redirection.

.PARAMETER Format
    Excel (default), Csv, or None. Excel requires the ImportExcel module; if that module is not
    installed the script warns and falls back to Csv rather than failing at the last step.

.PARAMETER PassThru
    Emit the report objects to the pipeline in addition to writing the report file. Combine
    with -Format None to skip the file entirely.

.EXAMPLE
    .\Get-AdDnsNodeReport.ps1 -ListZones

    Show every AD integrated DNS zone the discovered DC hosts, with its partition.

.EXAMPLE
    .\Get-AdDnsNodeReport.ps1 -ZoneName corp.example.com

    Report on dynamic A records in one zone and write an Excel report to the Desktop.

.EXAMPLE
    .\Get-AdDnsNodeReport.ps1 -ZoneName * -StaleThresholdDays 90 -Server dc01.corp.example.com -Verbose

    Every zone on a specific DC, with a 90 day stale bar.

.EXAMPLE
    $stale = .\Get-AdDnsNodeReport.ps1 -ZoneName corp.example.com -Format None -PassThru |
             Where-Object { $_.IsStale -and -not $_.OwnerIsMatchingComputer }

    Pipeline use with no file output: records that are both stale and owned by something other
    than a matching computer account. Review these before deleting anything.

.EXAMPLE
    .\Get-AdDnsNodeReport.ps1 -ZoneName 10.in-addr.arpa -RecordType PTR -Format Csv

    Reverse zone PTR records to CSV.

.NOTES
    Author: Mike Crowley
    https://mikecrowley.us

    Requires: ActiveDirectory module (RSAT). ImportExcel is optional and only needed for
    -Format Excel.

    Privilege: read-only, and Authenticated Users can read dnsNode objects and their security
    descriptors by default, so no elevation is normally required. This script never writes to
    the directory or deletes a record; treat its output as a review list, not an action list.

    Connectivity: every directory query goes through the ActiveDirectory module, which reaches
    the DC exclusively over AD Web Services (TCP 9389). This script opens no LDAP connection of
    its own on TCP 389. When -Server is omitted, initial DC discovery additionally uses DNS and
    the DC locator's CLDAP ping (UDP 389). Well-known SIDs are translated locally. No DNS
    server RPC access is used.

    Timestamps: DnsTimestampUtc comes from the dnsRecord blob and is UTC, expressed in whole
    hours. The DNS server only records aging to the hour, so there is no finer resolution to
    be had. NodeCreated / NodeChanged are the directory object's own whenCreated / whenChanged
    and are unrelated to DNS aging.

    Owner resolution is scoped to the domain of the queried DC. In a multi-domain forest, a
    record owned by a computer from another domain resolves to a SID with no friendly name.

    The dnsRecord parser here is a self-contained implementation of the layout in [MS-DNSP].
    Earlier revisions of this script pulled an ADDnsNode class from a public gist at runtime
    via Invoke-Expression; that is both a supply chain risk and a single point of failure if
    the gist ever disappears. Credit to Dave Carroll, whose gist first showed the approach:
    https://gist.github.com/thedavecarroll/ea2b4f0ba7527bd469e59e1aabf3e0b0

.LINK
    https://github.com/Mike-Crowley/Public-Scripts

.LINK
    https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-dnsp/ac793981-1c60-43b8-be59-cdbb5c4ecb8a

.LINK
    https://learn.microsoft.com/en-us/previous-versions/windows/it-pro/windows-server-2003/cc757041(v=ws.10)
#>

#Requires -Version 5.1
#Requires -Modules ActiveDirectory

[CmdletBinding(DefaultParameterSetName = 'Report')]
param(
    [Parameter(ParameterSetName = 'Report', Position = 0)]
    [SupportsWildcards()]
    [ValidateNotNullOrEmpty()]
    [string[]]$ZoneName,

    [Parameter(ParameterSetName = 'ListZones', Mandatory)]
    [switch]$ListZones,

    [ValidateNotNullOrEmpty()]
    [string]$Server,

    [System.Management.Automation.PSCredential]
    [System.Management.Automation.Credential()]
    $Credential = [System.Management.Automation.PSCredential]::Empty,

    [Parameter(ParameterSetName = 'Report')]
    [ValidateSet('A', 'AAAA', 'CNAME', 'PTR', 'NS', 'SOA', 'MX', 'SRV', 'TXT', 'All')]
    [string[]]$RecordType = 'A',

    [Parameter(ParameterSetName = 'Report')]
    [ValidateRange(1, 36500)]
    [int]$StaleThresholdDays = 180,

    [Parameter(ParameterSetName = 'Report')]
    [switch]$IncludeStatic,

    [Parameter(ParameterSetName = 'Report')]
    [switch]$IncludeTombstoned,

    [Parameter(ParameterSetName = 'Report')]
    [ValidateRange(0, [int]::MaxValue)]
    [int]$MaxRecordsPerZone = 0,

    [Parameter(ParameterSetName = 'Report')]
    [ValidateNotNullOrEmpty()]
    [string]$OutputFolder,

    [Parameter(ParameterSetName = 'Report')]
    [ValidateSet('Excel', 'Csv', 'None')]
    [string]$Format = 'Excel',

    [Parameter(ParameterSetName = 'Report')]
    [switch]$PassThru
)

$ErrorActionPreference = 'Stop'

#region Constants

# dnsRecord timestamps are whole HOURS elapsed since this instant (MS-DNSP 2.3.2.2).
$script:DnsEpochUtc = [datetime]::SpecifyKind([datetime]'1601-01-01T00:00:00', [System.DateTimeKind]::Utc)


# Resource record type numbers as they appear in the dnsRecord blob (MS-DNSP 2.2.2.1.1).
# Type 0 is the tombstone/ZERO record Windows writes in place of a deleted record.
$script:DnsRecordTypeName = @{
    0x0000 = 'TOMBSTONE'; 0x0001 = 'A'; 0x0002 = 'NS'; 0x0003 = 'MD'
    0x0004 = 'MF'; 0x0005 = 'CNAME'; 0x0006 = 'SOA'; 0x0007 = 'MB'
    0x0008 = 'MG'; 0x0009 = 'MR'; 0x000A = 'NULL'; 0x000B = 'WKS'
    0x000C = 'PTR'; 0x000D = 'HINFO'; 0x000E = 'MINFO'; 0x000F = 'MX'
    0x0010 = 'TXT'; 0x0011 = 'RP'; 0x0012 = 'AFSDB'; 0x0013 = 'X25'
    0x0014 = 'ISDN'; 0x0015 = 'RT'; 0x0018 = 'SIG'; 0x0019 = 'KEY'
    0x001C = 'AAAA'; 0x001D = 'LOC'; 0x001E = 'NXT'; 0x0021 = 'SRV'
    0x0022 = 'ATMA'; 0x0023 = 'NAPTR'; 0x0027 = 'DNAME'; 0x002B = 'DS'
    0x002E = 'RRSIG'; 0x002F = 'NSEC'; 0x0030 = 'DNSKEY'; 0x0031 = 'DHCID'
    0x0032 = 'NSEC3'; 0x0033 = 'NSEC3PARAM'; 0x0034 = 'TLSA'; 0x0040 = 'SVCB'
    0x0041 = 'HTTPS'; 0xFF01 = 'WINS'; 0xFF02 = 'WINSR'
}

# Types whose payload is a single DNS_COUNT_NAME and nothing else.
$script:DnsSingleNameTypes = @('NS', 'CNAME', 'PTR', 'DNAME', 'MB', 'MD', 'MF', 'MG', 'MR')

# Types whose payload is one or more length-prefixed character strings. Per the MS-DNSP
# DNS_RPC_RECORD_STRING table this is exactly HINFO, TXT, X25 and ISDN. LOC is deliberately
# excluded: it carries binary RFC 1876 wire data, so it falls through to the hex dump instead
# of being rendered as garbage text.
$script:DnsStringTypes = @('TXT', 'X25', 'HINFO', 'ISDN')

$script:AdRights = [System.DirectoryServices.ActiveDirectoryRights]

# ActiveDirectoryRights is a [Flags] enum, but the "generic" members are composites of the
# primitive bits, not bits of their own: GenericAll is 0xF01FF and GenericWrite is 0x20028,
# so GenericWrite -band GenericAll is non-zero. Testing membership with those composites
# therefore classifies everything as full control. The masks below use primitive bits only,
# and full control is tested as an exact superset match instead.
$script:GenericAllMask = [int]$script:AdRights::GenericAll

$script:WriteRightsMask = [int](
    $script:AdRights::CreateChild -bor $script:AdRights::DeleteChild -bor
    $script:AdRights::Self -bor $script:AdRights::WriteProperty -bor
    $script:AdRights::DeleteTree -bor $script:AdRights::Delete -bor
    $script:AdRights::WriteDacl -bor $script:AdRights::WriteOwner
)

$script:ReadRightsMask = [int](
    $script:AdRights::ListChildren -bor $script:AdRights::ReadProperty -bor
    $script:AdRights::ListObject -bor $script:AdRights::ReadControl
)

#endregion Constants

#region Binary helpers

# The dnsRecord blob mixes little-endian and big-endian fields, so read every integer with an
# explicit width and byte order instead of trusting [BitConverter], which follows host
# endianness. Multiplication rather than -shl: [byte]255 -shl 24 overflows [int] and goes
# negative on Windows PowerShell.
function Get-UInt16LittleEndian {
    param([byte[]]$Bytes, [int]$Offset)
    [uint16](([uint32]$Bytes[$Offset + 1] * 256) + [uint32]$Bytes[$Offset])
}

function Get-UInt16BigEndian {
    param([byte[]]$Bytes, [int]$Offset)
    [uint16](([uint32]$Bytes[$Offset] * 256) + [uint32]$Bytes[$Offset + 1])
}

function Get-UInt32LittleEndian {
    param([byte[]]$Bytes, [int]$Offset)
    [uint32](([uint64]$Bytes[$Offset + 3] * 16777216) + ([uint64]$Bytes[$Offset + 2] * 65536) +
             ([uint64]$Bytes[$Offset + 1] * 256) + [uint64]$Bytes[$Offset])
}

function Get-UInt32BigEndian {
    param([byte[]]$Bytes, [int]$Offset)
    [uint32](([uint64]$Bytes[$Offset] * 16777216) + ([uint64]$Bytes[$Offset + 1] * 65536) +
             ([uint64]$Bytes[$Offset + 2] * 256) + [uint64]$Bytes[$Offset + 3])
}

<#
    Decode a DNS_COUNT_NAME (MS-DNSP 2.2.2.2.3), the name encoding used inside dnsRecord:

        [Length][LabelCount][len][label][len][label]...[00]

    Length counts RawName INCLUDING the trailing null, which is why the next field cannot be
    found by simply walking the labels: that lands on the null, one byte short. SOA carries
    two names back to back, so getting this wrong reads the second one from the wrong offset.

    The label walk is driven by LabelCount rather than Length because that is what the DNS
    server itself and every independent implementation do, and it survives the occasional
    off-by-one Length seen in the wild.
#>
function ConvertFrom-DnsCountName {
    param([byte[]]$Bytes, [int]$Offset = 0)

    if ($null -eq $Bytes -or $Bytes.Length -lt ($Offset + 2)) {
        return [pscustomobject]@{ Name = $null; NextOffset = $Offset }
    }

    $declaredLength = [int]$Bytes[$Offset]
    $labelCount = [int]$Bytes[$Offset + 1]
    $position = $Offset + 2
    $labels = New-Object System.Collections.Generic.List[string]

    for ($i = 0; $i -lt $labelCount; $i++) {
        if ($position -ge $Bytes.Length) { break }
        $length = [int]$Bytes[$position]
        $position++
        if (($position + $length) -gt $Bytes.Length) { break }
        if ($length -gt 0) {
            $labels.Add([System.Text.Encoding]::UTF8.GetString($Bytes, $position, $length))
        }
        $position += $length
    }

    # Prefer the declared length (it accounts for the null terminator); fall back to the
    # walked position, plus the null, if it is absent or would run past the buffer.
    $next = $Offset + 2 + $declaredLength
    if ($declaredLength -le 0 -or $next -gt $Bytes.Length) {
        $next = [Math]::Min($position + 1, $Bytes.Length)
    }

    [pscustomobject]@{
        Name       = if ($labels.Count -gt 0) { ($labels -join '.') + '.' } else { '.' }
        NextOffset = $next
    }
}

# Decode a run of DNS_RPC_NAME character strings (one length byte then that many bytes).
function ConvertFrom-DnsCharacterString {
    param([byte[]]$Bytes)

    if ($null -eq $Bytes -or $Bytes.Length -eq 0) { return $null }

    $position = 0
    $strings = New-Object System.Collections.Generic.List[string]
    while ($position -lt $Bytes.Length) {
        $length = [int]$Bytes[$position]
        $position++
        if (($position + $length) -gt $Bytes.Length) { break }
        if ($length -gt 0) {
            $strings.Add([System.Text.Encoding]::UTF8.GetString($Bytes, $position, $length))
        }
        $position += $length
    }

    if ($strings.Count -gt 0) { $strings -join ' ' } else { $null }
}

<#
.SYNOPSIS
    Parse one value of the dnsRecord attribute into a structured object.

.DESCRIPTION
    Layout of the fixed 24 byte header, per MS-DNSP 2.3.2.2:

        Offset  Size  Field        Byte order
        0       2     DataLength   little
        2       2     Type         little
        4       1     Version      n/a
        5       1     Rank         n/a
        6       2     Flags        little
        8       4     Serial       little
        12      4     TtlSeconds   BIG
        16      4     Reserved     little
        20      4     Timestamp    little
        24      *     Data         type specific

    Timestamp is whole hours since 1601-01-01T00:00:00Z, and 0 means the record was created
    statically and is exempt from scavenging.
#>
function ConvertFrom-DnsRecordBlob {
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [byte[]]$Bytes
    )

    if ($null -eq $Bytes -or $Bytes.Length -lt 24) {
        Write-Verbose "Skipping a dnsRecord value shorter than the 24 byte header."
        return
    }

    $type = Get-UInt16LittleEndian -Bytes $Bytes -Offset 2
    $timestampHours = Get-UInt32LittleEndian -Bytes $Bytes -Offset 20

    # An exact length of 24 means there is no payload. Guard it: $Bytes[24..23] does not
    # produce an empty array in PowerShell, it silently reverses and returns element 23.
    $data = if ($Bytes.Length -gt 24) { [byte[]]$Bytes[24..($Bytes.Length - 1)] } else { [byte[]]@() }

    $typeName = if ($script:DnsRecordTypeName.ContainsKey([int]$type)) {
        $script:DnsRecordTypeName[[int]$type]
    }
    else {
        'TYPE{0}' -f $type
    }

    $recordData = switch ($typeName) {
        'A' {
            if ($data.Length -ge 4) { ([ipaddress]::new([byte[]]$data[0..3])).IPAddressToString }
            break
        }
        'AAAA' {
            if ($data.Length -ge 16) { ([ipaddress]::new([byte[]]$data[0..15])).IPAddressToString }
            break
        }
        'MX' {
            if ($data.Length -ge 4) {
                $preference = Get-UInt16BigEndian -Bytes $data -Offset 0
                '{0} {1}' -f $preference, (ConvertFrom-DnsCountName -Bytes $data -Offset 2).Name
            }
            break
        }
        'SRV' {
            if ($data.Length -ge 8) {
                '{0} {1} {2} {3}' -f (Get-UInt16BigEndian -Bytes $data -Offset 0),
                                     (Get-UInt16BigEndian -Bytes $data -Offset 2),
                                     (Get-UInt16BigEndian -Bytes $data -Offset 4),
                                     (ConvertFrom-DnsCountName -Bytes $data -Offset 6).Name
            }
            break
        }
        'SOA' {
            if ($data.Length -ge 22) {
                $primary = ConvertFrom-DnsCountName -Bytes $data -Offset 20
                $admin = ConvertFrom-DnsCountName -Bytes $data -Offset $primary.NextOffset
                '{0} {1} serial={2}' -f $primary.Name, $admin.Name, (Get-UInt32BigEndian -Bytes $data -Offset 0)
            }
            break
        }
        'TOMBSTONE' {
            # EntombedTime: a 64-bit FILETIME, little-endian, 100ns ticks since 1601.
            if ($data.Length -ge 8) {
                $high = [uint64](Get-UInt32LittleEndian -Bytes $data -Offset 4)
                $low = [uint64](Get-UInt32LittleEndian -Bytes $data -Offset 0)
                $entombed = ($high * [uint64]4294967296) + $low
                try { ([datetime]::FromFileTimeUtc([long]$entombed)).ToString('u') } catch { $null }
            }
            break
        }
        { $script:DnsSingleNameTypes -contains $_ } {
            (ConvertFrom-DnsCountName -Bytes $data).Name
            break
        }
        { $script:DnsStringTypes -contains $_ } {
            ConvertFrom-DnsCharacterString -Bytes $data
            break
        }
        default {
            if ($data.Length -gt 0) { ($data | ForEach-Object { $_.ToString('x2') }) -join '' }
            break
        }
    }

    # The timestamp is an unvalidated uint32 straight off the wire, and uint32 reaches roughly
    # 490,000 years past the 1601 epoch while DateTime stops at the year 9999. AddHours throws
    # on anything past that, which would abort an entire run over one corrupt record, and this
    # is not merely theoretical: Authenticated Users can create dnsNode objects in a default
    # AD-integrated zone. Representability is tested directly rather than against a derived
    # bound, because the exact cutoff lands where a double no longer has the significant digits
    # to distinguish it. The raw hours are still reported so the corruption stays visible.
    $timestampUtc = $null
    if ($timestampHours -gt 0) {
        try { $timestampUtc = $script:DnsEpochUtc.AddHours($timestampHours) }
        catch { Write-Verbose "Ignoring an out-of-range dnsRecord timestamp ($timestampHours hours); the stored value is corrupt." }
    }

    [pscustomobject]@{
        RecordType     = $typeName
        RecordTypeId   = [int]$type
        RecordData     = $recordData
        TtlSeconds     = Get-UInt32BigEndian -Bytes $Bytes -Offset 12
        Serial         = Get-UInt32LittleEndian -Bytes $Bytes -Offset 8
        Rank           = [int]$Bytes[5]
        Version        = [int]$Bytes[4]
        IsStatic       = ($timestampHours -eq 0)
        TimestampHours = $timestampHours
        TimestampUtc   = $timestampUtc
    }
}

#endregion Binary helpers

#region Directory helpers

<#
    Find every AD integrated DNS zone the target DC can serve.

    Zones live in up to three places, and a domain can use all three at once:
        CN=MicrosoftDNS,DC=DomainDnsZones,<domain NC>   domain-wide replication (default)
        CN=MicrosoftDNS,DC=ForestDnsZones,<forest root> forest-wide replication
        CN=MicrosoftDNS,CN=System,<domain NC>           legacy pre-Windows 2003 location

    namingContexts on the RootDSE lists only the partitions this DC actually hosts, which is
    exactly the set we can read, so it is the right source. A partition that fails to search
    is a warning rather than a fatal error: one unreadable partition should not lose the zones
    in the others.
#>
function Get-AdIntegratedDnsZone {
    [CmdletBinding()]
    param([hashtable]$AdParameters)

    $rootDse = Get-ADRootDSE @AdParameters
    $skip = @($rootDse.schemaNamingContext, $rootDse.configurationNamingContext)
    $zones = New-Object System.Collections.Generic.List[object]
    $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)

    foreach ($namingContext in $rootDse.namingContexts) {
        if ($skip -contains $namingContext) { continue }

        foreach ($container in @("CN=MicrosoftDNS,$namingContext", "CN=MicrosoftDNS,CN=System,$namingContext")) {
            try {
                $found = Get-ADObject -SearchBase $container -SearchScope OneLevel `
                    -LDAPFilter '(objectClass=dnsZone)' -Properties whenCreated @AdParameters
            }
            catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
                # Container does not exist in this partition. Normal, not an error.
                Write-Verbose "No DNS container at $container"
                continue
            }
            catch {
                Write-Warning "Could not search $container : $($_.Exception.Message)"
                continue
            }

            foreach ($zone in $found) {
                if (-not $seen.Add($zone.DistinguishedName)) { continue }
                $zones.Add([pscustomobject]@{
                    Name              = $zone.Name
                    Partition         = $namingContext
                    DistinguishedName = $zone.DistinguishedName
                    WhenCreated       = $zone.whenCreated
                })
            }
        }
    }

    if (-not ($rootDse.namingContexts -match '^DC=ForestDnsZones,')) {
        Write-Verbose 'This DC does not host the ForestDnsZones partition; forest-wide zones will not appear. Query a DC that hosts it to see them.'
    }

    $zones | Sort-Object Name
}

<#
    Turn an object's nTSecurityDescriptor into an ActiveDirectorySecurity regardless of whether
    the ActiveDirectory module handed it back already deserialized or as a raw byte array.
#>
function ConvertTo-ActiveDirectorySecurity {
    param($SecurityDescriptor)

    if ($null -eq $SecurityDescriptor) { return $null }
    if ($SecurityDescriptor -is [System.DirectoryServices.ActiveDirectorySecurity]) { return $SecurityDescriptor }

    if ($SecurityDescriptor -is [byte[]]) {
        $security = New-Object System.DirectoryServices.ActiveDirectorySecurity
        $security.SetSecurityDescriptorBinaryForm($SecurityDescriptor)
        return $security
    }

    return $null
}

# Describe a rights bitmask in the terms an admin cares about for a DDNS record.
function Get-RightsClassification {
    param([int]$RightsValue)

    if ($RightsValue -eq 0) { return 'None' }
    if (($RightsValue -band $script:GenericAllMask) -eq $script:GenericAllMask) { return 'FullControl' }
    if ($RightsValue -band $script:WriteRightsMask) { return 'Modify' }
    if ($RightsValue -band $script:ReadRightsMask) { return 'ReadOnly' }
    'Other'
}

#endregion Directory helpers

#region Setup

$adParameters = @{}
if ($PSBoundParameters.ContainsKey('Server')) { $adParameters['Server'] = $Server }
if ($Credential -ne [System.Management.Automation.PSCredential]::Empty) { $adParameters['Credential'] = $Credential }

if ($Server -match ':(3268|3269)$') {
    Write-Warning "$Server looks like a global catalog port. dnsRecord and nTSecurityDescriptor are not replicated to global catalogs, so this will return incomplete data. Use a standard LDAP endpoint instead."
}

Write-Verbose 'Enumerating AD integrated DNS zones.'
$allZones = @(Get-AdIntegratedDnsZone -AdParameters $adParameters)

if ($allZones.Count -eq 0) {
    throw 'No AD integrated DNS zones were found. Confirm the target DC hosts a DNS application partition, and that you are not pointed at a global catalog port.'
}

if ($ListZones) {
    $allZones | Select-Object Name, Partition, WhenCreated, DistinguishedName
    return
}

#endregion Setup

#region Zone selection

if ($ZoneName) {
    $selectedZones = @($allZones | Where-Object {
        $zone = $_
        $ZoneName | Where-Object { $zone.Name -like $_ }
    })

    if ($selectedZones.Count -eq 0) {
        throw "No zone matched: $($ZoneName -join ', '). Run with -ListZones to see what is available."
    }
}
elseif (-not [Environment]::UserInteractive) {
    throw 'No -ZoneName was supplied and this session is not interactive. Pass -ZoneName (wildcards allowed, * for all) or run with -ListZones first.'
}
else {
    $selectedZones = @()
    $usedPicker = $false

    # Out-GridView exists on Windows PowerShell and on PowerShell 7 for Windows, but it cannot
    # draw a window over a remoting session, so treat a failure as "no picker" and drop to the
    # console list rather than dying on zone selection.
    if (Get-Command -Name Out-GridView -ErrorAction SilentlyContinue) {
        try {
            Write-Host 'Select the DNS zone(s) to report on...' -ForegroundColor Cyan
            $selectedZones = @($allZones | Out-GridView -Title 'Select DNS zone(s)' -OutputMode Multiple)
            $usedPicker = $true
        }
        catch {
            Write-Verbose "Out-GridView unavailable here ($($_.Exception.Message)); falling back to a console list."
            $usedPicker = $false
        }
    }

    if (-not $usedPicker) {
        Write-Host 'Available DNS zones:' -ForegroundColor Cyan
        for ($i = 0; $i -lt $allZones.Count; $i++) {
            Write-Host ('  [{0,3}] {1}' -f $i, $allZones[$i].Name)
        }
        $answer = Read-Host 'Enter zone number(s), comma separated'
        $selectedZones = @(
            $answer -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -match '^\d+$' } |
                ForEach-Object { [int]$_ } | Where-Object { $_ -ge 0 -and $_ -lt $allZones.Count } |
                ForEach-Object { $allZones[$_] }
        )
    }

    if ($selectedZones.Count -eq 0) {
        Write-Host 'No zone selected; nothing to do.' -ForegroundColor Yellow
        return
    }
}

Write-Verbose "Selected $($selectedZones.Count) zone(s): $(($selectedZones.Name) -join ', ')"

#endregion Zone selection

#region Collect records

$nodeProperties = @('name', 'dnsRecord', 'dNSTombstoned', 'whenCreated', 'whenChanged', 'nTSecurityDescriptor')
$wantAllTypes = $RecordType -contains 'All'
$staleCutoffUtc = (Get-Date).ToUniversalTime().AddDays(-$StaleThresholdDays)
$scriptStart = Get-Date

$records = New-Object System.Collections.Generic.List[object]

foreach ($zone in $selectedZones) {
    Write-Verbose "Querying dnsNode objects in $($zone.Name)."

    # Filter tombstoned nodes server side rather than fetching and discarding them. The
    # negated clause also matches nodes where the attribute is simply absent, which is the
    # normal case for a live record.
    $nodeFilter = if ($IncludeTombstoned) { '(objectClass=dnsNode)' }
    else { '(&(objectClass=dnsNode)(!(dNSTombstoned=TRUE)))' }

    $query = @{
        SearchBase  = $zone.DistinguishedName
        SearchScope = 'Subtree'
        LDAPFilter  = $nodeFilter
        Properties  = $nodeProperties
    }

    try {
        # The cap is applied client side rather than with -ResultSetSize. When the server hits
        # a ResultSetSize cap the AD cmdlet writes a "size limit exceeded" error, and
        # $ErrorActionPreference = 'Stop' promotes that to terminating, so the catch below
        # would discard the entire zone in precisely the case the cap exists to handle.
        # Select-Object -First stops the upstream pipeline, so nothing extra is fetched.
        # One extra object is requested so truncation can be detected exactly rather than
        # guessed at from a count that merely equals the cap.
        $nodes = if ($MaxRecordsPerZone -gt 0) {
            @(Get-ADObject @query @adParameters | Select-Object -First ($MaxRecordsPerZone + 1))
        }
        else {
            @(Get-ADObject @query @adParameters)
        }
    }
    catch {
        Write-Warning "Could not query $($zone.Name): $($_.Exception.Message)"
        continue
    }

    if ($MaxRecordsPerZone -gt 0 -and $nodes.Count -gt $MaxRecordsPerZone) {
        $nodes = @($nodes | Select-Object -First $MaxRecordsPerZone)
        Write-Warning "$($zone.Name): stopped at -MaxRecordsPerZone $MaxRecordsPerZone. The report is a sample, not the whole zone."
    }

    Write-Verbose "  $($nodes.Count) dnsNode object(s) returned."

    foreach ($node in $nodes) {
        $isTombstoned = [bool]$node.dNSTombstoned
        if ($isTombstoned -and -not $IncludeTombstoned) { continue }
        if (-not $node.dnsRecord) { continue }

        # One dnsNode can carry several records (round-robin A records, a name with both an
        # A and a TXT, and so on), so every value has to be parsed.
        foreach ($blob in $node.dnsRecord) {
            $parsed = ConvertFrom-DnsRecordBlob -Bytes ([byte[]]$blob)
            if (-not $parsed) { continue }

            # A tombstone is record type 0 carrying a zero timestamp, so both filters below
            # would discard it and -IncludeTombstoned would be a switch that does nothing.
            # If the caller explicitly asked for tombstoned nodes, show them.
            if (-not $isTombstoned) {
                if (-not $wantAllTypes -and $RecordType -notcontains $parsed.RecordType) { continue }
                if ($parsed.IsStatic -and -not $IncludeStatic) { continue }
            }

            $records.Add([pscustomobject]@{
                Zone         = $zone.Name
                Node         = $node
                Parsed       = $parsed
                IsTombstoned = $isTombstoned
            })
        }
    }
}

if ($records.Count -eq 0) {
    Write-Warning 'No records matched the current filters. Try -RecordType All, -IncludeStatic, or a different zone.'
    return
}

Write-Verbose "$($records.Count) record(s) to analyze."

#endregion Collect records

#region Owner index

<#
    Owner SIDs are looked up once each and cached. For a large result set, one bulk pass over
    the domain's computer objects is dramatically cheaper than thousands of single-object
    round trips; for a small one, the bulk pass costs more than it saves. The threshold below
    is where the two cross over in practice.
#>
$ownerCache = @{}
$bulkIndexThreshold = 250

if ($records.Count -ge $bulkIndexThreshold) {
    Write-Verbose 'Building a SID index of domain computer objects (one query).'
    try {
        Get-ADComputer -Filter * -Properties whenChanged @adParameters | ForEach-Object {
            $ownerCache[$_.SID.Value] = [pscustomobject]@{
                Name              = $_.Name
                ObjectClass       = 'computer'
                DistinguishedName = $_.DistinguishedName
                WhenChanged       = $_.whenChanged
                Enabled           = $_.Enabled
            }
        }
        Write-Verbose "  indexed $($ownerCache.Count) computer object(s)."
    }
    catch {
        Write-Warning "Could not build the computer index, falling back to per-record lookups: $($_.Exception.Message)"
    }
}

function Resolve-OwnerPrincipal {
    param([System.Security.Principal.SecurityIdentifier]$Sid)

    if ($null -eq $Sid) { return $null }
    if ($ownerCache.ContainsKey($Sid.Value)) { return $ownerCache[$Sid.Value] }

    $resolved = $null

    if ($Sid.Value -notmatch '^S-1-5-21-') {
        # Not a domain SID, so it is well-known or built-in (SYSTEM, BUILTIN\Administrators,
        # Enterprise Domain Controllers). Those are not directory objects and an LDAP lookup
        # will never find them, so translate locally instead.
        try {
            $resolved = [pscustomobject]@{
                Name              = $Sid.Translate([System.Security.Principal.NTAccount]).Value
                ObjectClass       = 'wellKnown'
                DistinguishedName = $null
                WhenChanged       = $null
                Enabled           = $null
            }
        }
        catch {
            $resolved = $null
        }
    }
    else {
        # Get-ADObject -Identity accepts only a DN or a GUID, NOT a SID (Get-ADComputer does,
        # but the owner is not always a computer), so filter on objectSid instead.
        try {
            $object = Get-ADObject -LDAPFilter "(objectSid=$($Sid.Value))" `
                -Properties objectClass, whenChanged, userAccountControl @adParameters |
                Select-Object -First 1

            if ($object) {
                $resolved = [pscustomobject]@{
                    Name              = $object.Name
                    ObjectClass       = $object.ObjectClass
                    DistinguishedName = $object.DistinguishedName
                    WhenChanged       = $object.whenChanged
                    # ADS_UF_ACCOUNTDISABLE is 0x2.
                    Enabled           = if ($null -ne $object.userAccountControl) {
                        -not ($object.userAccountControl -band 2)
                    }
                    else { $null }
                }
            }
        }
        catch {
            $resolved = $null
        }
    }

    $ownerCache[$Sid.Value] = $resolved
    $resolved
}

#endregion Owner index

#region Build report

$report = New-Object System.Collections.Generic.List[object]
$counter = 0
$total = $records.Count
$progressEvery = [Math]::Max(1, [int]($total / 100))

foreach ($record in $records) {
    $counter++
    if (($counter % $progressEvery) -eq 0 -or $counter -eq $total) {
        Write-Progress -Activity 'Analyzing DNS record ownership and permissions' `
            -Status "$counter of $total" -PercentComplete (($counter / $total) * 100)
    }

    $node = $record.Node
    $parsed = $record.Parsed

    $security = ConvertTo-ActiveDirectorySecurity -SecurityDescriptor $node.nTSecurityDescriptor
    $ownerSid = $null
    $ownerRightsValue = 0

    if ($security) {
        # GetOwner with an explicit SecurityIdentifier target skips name resolution entirely,
        # so an orphaned or cross-domain owner returns a SID instead of throwing.
        try { $ownerSid = $security.GetOwner([System.Security.Principal.SecurityIdentifier]) } catch { $ownerSid = $null }

        if ($ownerSid) {
            # Ask for the ACEs as SIDs and compare SIDs. The previous approach matched the
            # record name against IdentityReference as a regex, which both mis-fires on names
            # containing regex metacharacters and matches any principal whose name merely
            # contains the record name.
            foreach ($ace in $security.GetAccessRules($true, $true, [System.Security.Principal.SecurityIdentifier])) {
                if ($ace.AccessControlType -ne [System.Security.AccessControl.AccessControlType]::Allow) { continue }
                if ($ace.IdentityReference.Value -ne $ownerSid.Value) { continue }
                # Rights accumulate across ACEs, so OR them together rather than testing each.
                $ownerRightsValue = $ownerRightsValue -bor [int]$ace.ActiveDirectoryRights
            }
        }
    }

    $owner = if ($ownerSid) { Resolve-OwnerPrincipal -Sid $ownerSid } else { $null }
    $ownerRights = Get-RightsClassification -RightsValue $ownerRightsValue

    $ownerIsMatchingComputer = ($null -ne $owner) -and
                               ($owner.ObjectClass -eq 'computer') -and
                               ($owner.Name -eq $node.Name)

    $fqdn = if ($node.Name -eq '@') { $record.Zone } else { '{0}.{1}' -f $node.Name, $record.Zone }

    $ageDays = if ($parsed.TimestampUtc) {
        [Math]::Round(((Get-Date).ToUniversalTime() - $parsed.TimestampUtc).TotalDays, 1)
    }
    else { $null }

    $isStale = ($null -ne $parsed.TimestampUtc) -and ($parsed.TimestampUtc -lt $staleCutoffUtc)

    # Escaped commas are legal inside a DN component, so split on commas that are not escaped.
    $ownerDnParts = if ($owner -and $owner.DistinguishedName) { $owner.DistinguishedName -split '(?<!\\),' } else { @() }
    $rootOu = if ($ownerDnParts.Count -gt 1) {
        ($ownerDnParts | Select-Object -Skip 1 | Where-Object { $_ -notmatch '^DC=' } | Select-Object -Last 1)
    }
    else { $null }
    $parentOu = if ($ownerDnParts.Count -gt 1) { ($ownerDnParts | Select-Object -Skip 1) -join ',' } else { $null }

    $reasons = New-Object System.Collections.Generic.List[string]
    if ($isStale) { $reasons.Add("Stale: not refreshed in $StaleThresholdDays days") }
    if (-not $ownerSid) { $reasons.Add('No readable owner on the record') }
    elseif (-not $owner) { $reasons.Add('Owner SID does not resolve to an existing object (orphaned)') }
    elseif (-not $ownerIsMatchingComputer) { $reasons.Add("Owner is not a computer account matching the record name (owner: $($owner.Name))") }
    if ($ownerSid -and $ownerRights -notin @('FullControl', 'Modify')) {
        $reasons.Add("Owner cannot update its own record (rights: $ownerRights)")
    }
    if ($record.IsTombstoned) { $reasons.Add('Record is tombstoned and awaiting cleanup') }

    $report.Add([pscustomobject]@{
        RequiresReview          = ($reasons.Count -gt 0)
        ReviewReasons           = $reasons -join '; '
        Zone                    = $record.Zone
        RecordName              = $node.Name
        Fqdn                    = $fqdn
        RecordType              = $parsed.RecordType
        RecordData              = $parsed.RecordData
        IsStale                 = $isStale
        AgeDays                 = $ageDays
        DnsTimestampUtc         = $parsed.TimestampUtc
        IsStatic                = $parsed.IsStatic
        IsTombstoned            = $record.IsTombstoned
        TtlSeconds              = $parsed.TtlSeconds
        OwnerIsMatchingComputer = $ownerIsMatchingComputer
        OwnerName               = if ($owner) { $owner.Name } else { $null }
        OwnerObjectClass        = if ($owner) { $owner.ObjectClass } else { $null }
        OwnerEnabled            = if ($owner) { $owner.Enabled } else { $null }
        OwnerRights             = $ownerRights
        OwnerRightsDetail       = ([System.DirectoryServices.ActiveDirectoryRights]$ownerRightsValue).ToString()
        OwnerSid                = if ($ownerSid) { $ownerSid.Value } else { $null }
        OwnerDistinguishedName  = if ($owner) { $owner.DistinguishedName } else { $null }
        OwnerWhenChanged        = if ($owner) { $owner.WhenChanged } else { $null }
        RootOu                  = $rootOu
        ParentOu                = $parentOu
        NodeCreated             = $node.whenCreated
        NodeChanged             = $node.whenChanged
        NodeDistinguishedName   = $node.DistinguishedName
    })
}

Write-Progress -Activity 'Analyzing DNS record ownership and permissions' -Completed

$elapsed = (Get-Date) - $scriptStart
Write-Verbose ('Analyzed {0} record(s) in {1:n1} minute(s).' -f $report.Count, $elapsed.TotalMinutes)

$staleCount = @($report | Where-Object IsStale).Count
$reviewCount = @($report | Where-Object RequiresReview).Count

Write-Host ''
Write-Host ("{0,-32}{1}" -f 'Records analyzed:', $report.Count) -ForegroundColor Cyan
Write-Host ("{0,-32}{1}" -f "Stale (>$StaleThresholdDays days):", $staleCount) -ForegroundColor Cyan
Write-Host ("{0,-32}{1}" -f 'Flagged for review:', $reviewCount) -ForegroundColor Cyan
Write-Host ''

#endregion Build report

#region Output

if ($Format -ne 'None') {
    if (-not $OutputFolder) {
        $desktop = [Environment]::GetFolderPath('Desktop')
        $OutputFolder = if ($desktop) { $desktop } else { $PWD.Path }
    }
    if (-not (Test-Path -LiteralPath $OutputFolder)) {
        New-Item -Path $OutputFolder -ItemType Directory -Force | Out-Null
    }

    if ($Format -eq 'Excel' -and -not (Get-Module -ListAvailable -Name ImportExcel)) {
        Write-Warning 'The ImportExcel module is not installed (Install-Module ImportExcel -Scope CurrentUser). Falling back to CSV.'
        $Format = 'Csv'
    }

    # Seconds, not just minutes: Export-Excel MERGES into an existing workbook rather than
    # replacing it, so two runs sharing a filename would leave the first run's rows sitting
    # below the second run's data (and stack a duplicate set of charts on the dashboard).
    $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $baseName = "dns-node-report-$stamp"
}

if ($Format -eq 'Csv') {
    $reportPath = Join-Path $OutputFolder "$baseName.csv"
    $report | Export-Csv -LiteralPath $reportPath -NoTypeInformation -Encoding UTF8
    Write-Host "Report: $reportPath" -ForegroundColor Green
}
elseif ($Format -eq 'Excel') {
    $reportPath = Join-Path $OutputFolder "$baseName.xlsx"
    Import-Module ImportExcel -ErrorAction Stop

    # Belt and braces against the merge behavior noted above, in case a caller passes an
    # -OutputFolder holding a report from the same second.
    if (Test-Path -LiteralPath $reportPath) { Remove-Item -LiteralPath $reportPath -Force }

    $excelPackage = $report | Export-Excel -Path $reportPath -WorksheetName 'Records' `
        -BoldTopRow -FreezeTopRow -AutoFilter -AutoSize -PassThru

    try {
        $worksheet = $excelPackage.Workbook.Worksheets['Records']
        $lastRow = $worksheet.Dimension.End.Row
        $lastColumn = $worksheet.Dimension.End.Column

        # Resolve columns by header text rather than by assumed position, so reordering or
        # adding a property cannot silently color the wrong column.
        $columnOf = @{}
        for ($column = 1; $column -le $lastColumn; $column++) {
            $header = $worksheet.Cells[1, $column].Text
            if ($header) { $columnOf[$header] = $column }
        }

        if ($lastRow -ge 2) {
            $lastColumnLetter = (Get-ExcelColumnName $lastColumn).ColumnName

            # Export-Excel writes a PowerShell [bool] as a real Excel boolean, not the text
            # "TRUE". A cellIs/Equal rule is emitted as ="TRUE" and would never match one, so
            # these have to be Expression rules that compare against the boolean itself.
            #
            # Order matters: the two column-scoped rules are added before the whole-row tint
            # so they keep the higher priority, and therefore the visible fill, where the
            # ranges overlap.
            $highlights = @(
                @{ Header = 'IsStale'; Test = 'TRUE'; Color = 'Gold'; WholeRow = $false }
                @{ Header = 'OwnerIsMatchingComputer'; Test = 'FALSE'; Color = 'Wheat'; WholeRow = $false }
                @{ Header = 'RequiresReview'; Test = 'TRUE'; Color = 'MistyRose'; WholeRow = $true }
            )

            foreach ($highlight in $highlights) {
                if (-not $columnOf.ContainsKey($highlight.Header)) { continue }
                $letter = (Get-ExcelColumnName $columnOf[$highlight.Header]).ColumnName
                $address = if ($highlight.WholeRow) {
                    'A2:{0}{1}' -f $lastColumnLetter, $lastRow
                }
                else {
                    '{0}2:{0}{1}' -f $letter, $lastRow
                }
                Add-ConditionalFormatting -Worksheet $worksheet -Address $address `
                    -RuleType Expression -ConditionValue ('${0}2={1}' -f $letter, $highlight.Test) `
                    -BackgroundColor $highlight.Color
            }
        }

        # Dashboard layout, kept deliberately non-overlapping:
        #   rows 1-15   four charts, anchored at row 0 in columns 0/7/14/21
        #   rows 17-23  run summary
        #   rows 26+    the summary tables the charts are plotted from
        # Each summary is written before its chart so the chart range can follow the real row
        # count. The previous version hardcoded rows 21:22 and 21:46, which silently plotted
        # blank cells or clipped categories as soon as the data was a different size.
        $summaryHeaderRow = 26
        $summaries = @(
            @{
                Title = 'Stale records'; Chart = 'Pie'; DataColumn = 1; ChartColumn = 0
                Data  = @($report | Group-Object IsStale -NoElement | Select-Object Name, Count | Sort-Object Count -Descending)
            }
            @{
                Title = 'Owner matches record name'; Chart = 'Doughnut'; DataColumn = 5; ChartColumn = 7
                Data  = @($report | Group-Object OwnerIsMatchingComputer -NoElement | Select-Object Name, Count | Sort-Object Count -Descending)
            }
            @{
                Title = 'Owner rights on own record'; Chart = 'BarClustered'; DataColumn = 9; ChartColumn = 14
                Data  = @($report | Group-Object OwnerRights -NoElement | Select-Object Name, Count | Sort-Object Count -Descending)
            }
            @{
                Title = 'Top OUs holding stale record owners'; Chart = 'BarClustered'; DataColumn = 13; ChartColumn = 21
                Data  = @($report | Where-Object { $_.IsStale -and $_.RootOu } | Group-Object RootOu -NoElement |
                          Select-Object Name, Count | Sort-Object Count -Descending | Select-Object -First 15)
            }
        )

        foreach ($summary in $summaries) {
            if ($summary.Data.Count -eq 0) { continue }

            $excelPackage = $summary.Data | Export-Excel -ExcelPackage $excelPackage -WorksheetName 'Dashboard' `
                -StartRow $summaryHeaderRow -StartColumn $summary.DataColumn -BoldTopRow -PassThru

            $nameLetter = (Get-ExcelColumnName $summary.DataColumn).ColumnName
            $countLetter = (Get-ExcelColumnName ($summary.DataColumn + 1)).ColumnName
            $firstDataRow = $summaryHeaderRow + 1
            $lastDataRow = $summaryHeaderRow + $summary.Data.Count

            $chartParameters = @{
                Title     = $summary.Title
                ChartType = $summary.Chart
                XRange    = 'Dashboard!{0}{1}:{0}{2}' -f $nameLetter, $firstDataRow, $lastDataRow
                YRange    = 'Dashboard!{0}{1}:{0}{2}' -f $countLetter, $firstDataRow, $lastDataRow
                Row       = 0
                Column    = $summary.ChartColumn
                Width     = 400
                Height    = 300
            }
            if ($summary.Chart -in @('Pie', 'Doughnut')) { $chartParameters['ShowPercent'] = $true }

            $excelPackage = Export-Excel -ExcelPackage $excelPackage -WorksheetName 'Dashboard' `
                -ExcelChartDefinition (New-ExcelChartDefinition @chartParameters) -PassThru
        }

        $dashboard = $excelPackage.Workbook.Worksheets['Dashboard']
        if ($dashboard) {
            $facts = [ordered]@{
                'Generated'          = (Get-Date).ToString('u')
                'Zones'              = ($selectedZones.Name -join ', ')
                'Record types'       = ($RecordType -join ', ')
                'Stale threshold'    = "$StaleThresholdDays days"
                'Records analyzed'   = $report.Count
                'Stale records'      = $staleCount
                'Flagged for review' = $reviewCount
            }
            $row = 17
            foreach ($fact in $facts.GetEnumerator()) {
                $dashboard.Cells[$row, 1].Value = $fact.Key
                $dashboard.Cells[$row, 1].Style.Font.Bold = $true
                $dashboard.Cells[$row, 2].Value = "$($fact.Value)"
                $row++
            }
            $dashboard.Column(1).Width = 24
            $dashboard.Column(2).Width = 42
            $excelPackage = Export-Excel -ExcelPackage $excelPackage -WorksheetName 'Dashboard' -MoveToStart -PassThru
        }

        Close-ExcelPackage -ExcelPackage $excelPackage
        $excelPackage = $null
        Write-Host "Report: $reportPath" -ForegroundColor Green
    }
    finally {
        # Release the file handle even if the formatting or dashboard step throws, so a partial
        # failure does not leave a locked, half-written workbook behind.
        if ($excelPackage) { $excelPackage.Dispose() }
    }
}

if ($PassThru) { $report }

#endregion Output
