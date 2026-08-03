<#
.SYNOPSIS
    Queries Entra ID credential type and OpenID configuration for a given UPN (unauthenticated OSINT).

.DESCRIPTION
    Calls the Entra ID GetCredentialType and OpenID configuration endpoints to gather
    tenant information for a given email address. Returns user existence, domain type,
    federation status, tenant GUID, region, and preferred credential type.

    No authentication is required. This uses public Microsoft endpoints.

.PARAMETER Upn
    An email address (UPN) or a bare domain. Positional, so the parameter name can be
    omitted. Accepts multiple values and pipeline input.

    Passing a domain answers the domain half of the question only: tenant GUID, region,
    and managed vs federated. The user-existence fields come back $null rather than
    INVALID_USER, because there is no user to report on. InputType records which form was
    used.

.EXAMPLE
    .\Get-EntraCredentialInfo.ps1 user@example.com

    Runs the query directly, no dot-sourcing needed.

.EXAMPLE
    .\Get-EntraCredentialInfo.ps1 example.com

    Domain only: is this an Entra tenant, what is its GUID and region, and is it federated?

.EXAMPLE
    Import-Module .\Get-EntraCredentialInfo.ps1
    Get-EntraCredentialInfo -Upn user@example.com

    Import (or dot-source) with no arguments to load the function for interactive use.

.EXAMPLE
    Get-Content .\addresses.txt | .\Get-EntraCredentialInfo.ps1

    Pipes several addresses or domains through in one pass.

.NOTES
    Author: Mike Crowley
    https://mikecrowley.us

.LINK
    https://mikecrowley.us
#>

# Plain param block on purpose. Adding [CmdletBinding()] or a [Parameter()] attribute makes
# this an advanced script, which then REFUSES pipeline input unless a parameter declares
# ValueFromPipeline, and honouring that requires a process{} block. A process{} block in turn
# breaks Import-Module, because Import-Module runs a .ps1 by calling Invoke() on it as a
# script block and Invoke() rejects blocks with more than one clause. Plain param plus $input
# is the only shape that serves positional, named, pipeline, dot-source, and Import-Module at
# once. The real work lives in the function below, which is a proper advanced function.
param (
    [string[]]$Upn
)

function Get-EntraCredentialInfo {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [ValidateNotNullOrEmpty()]
        [Alias('Domain', 'Identity')]
        [string]$Upn
    )

    process {

    # A bare domain is a legitimate question ("is crowley.us an Entra tenant, and is it
    # federated?"), so it is not treated as a malformed UPN. GetCredentialType keys its
    # DomainType off the domain rather than the user, so a throwaway local part still
    # answers the domain half; the user-existence half is reported as $null rather than as
    # a misleading INVALID_USER.
    $IsDomainOnly = $Upn -notmatch '@'
    if ($IsDomainOnly) {
        $Domain = $Upn.Trim()
        if ($Domain -notmatch '\.') {
            Write-Warning "'$Domain' has no @ and no dot, so it is neither a UPN nor a domain."
            return
        }
        $ProbeUser = "$([guid]::NewGuid().ToString('N'))@$Domain"
    }
    else {
        $Domain = ($Upn -split '@')[1]
        $ProbeUser = $Upn
    }

    $Body = @{
        username            = $ProbeUser
        isOtherIdpSupported = $true
    }
    $Body = $Body | ConvertTo-Json -Compress

    try {
        $CredentialResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/common/GetCredentialType" -Method Post -Body $Body -ContentType "application/json" -ErrorAction Stop
    }
    catch {
        Write-Warning "Failed to retrieve credential type for $Upn : $($_.Exception.Message)"
        return
    }

    try {
        $OpenidResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$Domain/.well-known/openid-configuration" -ErrorAction Stop
    }
    catch {
        Write-Warning "Failed to retrieve OpenID configuration for $Domain : $($_.Exception.Message)"
        $OpenidResponse = $null
    }

    $Output = [pscustomobject]@{
        Username                  = if ($IsDomainOnly) { $null } else { $CredentialResponse.Username }
        Domain                    = $Domain
        InputType                 = if ($IsDomainOnly) { 'Domain' } else { 'UPN' }
        UserFound                 = if ($IsDomainOnly) { $null } else { $CredentialResponse.IfExistsResult -ne 1 }

        #IfExistsResult = $CredentialResponse.IfExistsResult
        IfExistsResultDescription = if ($IsDomainOnly) { $null } else {
            switch ($CredentialResponse.IfExistsResult) {
                "-1" { "UNKNOWN" }
                "0" { "VALID_USER" }
                "1" { "INVALID_USER" }
                "2" { "THROTTLE" }
                "4" { "ERROR" }
                "5" { "VALID_USER-DIFFERENT_IDP" }
                "6" { "VALID_USER-ExistsBoth_IDP" } # causes pidpdisambiguation / accountpicker
                default { $CredentialResponse.IfExistsResult }
            } # https://github.com/BarrelTit0r/o365enum/blob/master/o365enum.py
        }

        #PrefCredential            = $CredentialResponse.Credentials.PrefCredential
        PrefCredentialDescription = if ($IsDomainOnly) { $null } else {
            switch ($CredentialResponse.Credentials.PrefCredential) {
                "0" { "0" }
                "1" { "1" }
                "2" { "2" }
                "3" { "3" }
                default { $CredentialResponse.Credentials.PrefCredential }
            } # TO DO - https://learn.microsoft.com/en-us/entra/identity/authentication/concept-system-preferred-multifactor-authentication#how-does-system-preferred-mfa-determine-the-most-secure-method
        }

        FederatedDomain           = $null -ne $CredentialResponse.Credentials.FederationRedirectUrl

        #DomainType              = $CredentialResponse.EstsProperties.DomainType
        DomainTypeDescription     = switch ($CredentialResponse.EstsProperties.DomainType) {
            '1' { "UNKNOWN" }
            '2' { "COMMERCIAL" }
            '3' { "MANAGED" }
            '4' { "FEDERATED" }
            '5' { "CLOUD_FEDERATED" }
            default { $CredentialResponse.EstsProperties.DomainType }
        }

        #DesktopSsoEnabled       = $CredentialResponse.EstsProperties.DesktopSsoEnabled
        #UserTenantBranding      = $CredentialResponse.EstsProperties.UserTenantBranding
        # Pulled out of the issuer by pattern rather than by trimming known authority
        # prefixes, so US Gov and 21Vianet issuers resolve too.
        TenantGuid                = if ("$($OpenidResponse.issuer)" -match '([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})') { $Matches[1] } else { $null }
        tenant_region_scope       = if ($null -ne $OpenidResponse) { $OpenidResponse.tenant_region_scope } else {}
        tenant_region_sub_scope   = if ($null -eq $OpenidResponse.tenant_region_sub_scope) { "WW" } else { $OpenidResponse.tenant_region_sub_scope }
        #CredentialResponse        = if ($null -ne $OpenidResponse) { $OpenidResponse.cloud_instance_name } else {}
        FederationRedirectUrl     = $CredentialResponse.Credentials.FederationRedirectUrl
    }

    $Output

    if ($Output.DomainTypeDescription -eq "FEDERATED") {
        $Who = if ($IsDomainOnly) { $Domain } else { $Output.Username }
        Write-Warning "[$Who] All users in a FEDERATED domain return VALID_USER by this endpoint. You must confirm with the system referenced in the FederationRedirectUrl.`n"
    }

    } # process
}

# Defining the function is all this file used to do, which is why invoking it with a UPN
# printed nothing. The call below fixes that.
#
# Deliberately NOT wrapped in begin/process/end: Import-Module runs a .ps1 by invoking it as
# a script block, and Invoke() rejects a block with more than one clause. Pipeline input is
# read from $input instead, which keeps this file single-clause and therefore importable.
$Queue = @()
if ($Upn) { $Queue += $Upn }
$Queue += @($input)
foreach ($Entry in $Queue) {
    if ("$Entry".Trim()) { Get-EntraCredentialInfo -Upn "$Entry".Trim() }
}
