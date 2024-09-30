<#
.SYNOPSIS
    Misc OSINT

.EXAMPLE

    Get-EntraCredentialInfo -Upn user@example.com

.LINK

    https://mikecrowley.us
#>

function Get-EntraCredentialInfo {
    param (
        [parameter(Mandatory = $true)][string]
        $Upn
    )

    $Domain = ($Upn -split '@')[1]
    $Body = @{
        username            = $Upn
        isOtherIdpSupported = $true
    }
    $Body = $Body | ConvertTo-Json -Compress

    $ErrorActionPreference = "SilentlyContinue"
    $CredentialResponse = Invoke-RestMethod "https://login.microsoftonline.com/common/GetCredentialType" -Method Post -Body $Body
    $OpenidResponse = Invoke-WebRequest "https://login.microsoftonline.com/$domain/.well-known/openid-configuration" | ConvertFrom-Json
    $ErrorActionPreference = "Continue"

    $Output = [pscustomobject]@{
        Username                  = $CredentialResponse.Username
        Domain                    = $Domain
        UserFound                 = $CredentialResponse.IfExistsResult -ne 1

        #IfExistsResult = $CredentialResponse.IfExistsResult
        IfExistsResultDescription = switch ($CredentialResponse.IfExistsResult) {
            "-1" { "UNKNOWN" }
            "0" { "VALID_USER" }
            "1" { "INVALID_USER" }
            "2" { "THROTTLE" }
            "4" { "ERROR" }
            "5" { "VALID_USER-DIFFERENT_IDP" }
            "6" { "VALID_USER-ExistsBoth_IDP" } # causes pidpdisambiguation / accountpicker
            default { $CredentialResponse.IfExistsResult }
        } # https://github.com/BarrelTit0r/o365enum/blob/master/o365enum.py

        #PrefCredential            = $CredentialResponse.Credentials.PrefCredential
        PrefCredentialDescription = switch ($CredentialResponse.Credentials.PrefCredential) {
            "0" { "0" }
            "1" { "1" }
            "2" { "2" }
            "3" { "3" }
            default { $CredentialResponse.Credentials.PrefCredential }
        } # TO DO - https://learn.microsoft.com/en-us/entra/identity/authentication/concept-system-preferred-multifactor-authentication#how-does-system-preferred-mfa-determine-the-most-secure-method

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
        TenantGuid                = if ($null -ne $OpenidResponse) { $OpenidResponse.userinfo_endpoint -replace 'https://login.microsoftonline.com/' -replace 'https://login.microsoftonline.us/' -replace '/openid/userinfo' } else {}
        tenant_region_scope       = if ($null -ne $OpenidResponse) { $OpenidResponse.tenant_region_scope } else {}
        tenant_region_sub_scope   = if ($null -eq $OpenidResponse.tenant_region_sub_scope) { "WW" } else { $OpenidResponse.tenant_region_sub_scope }
        #CredentialResponse        = if ($null -ne $OpenidResponse) { $OpenidResponse.cloud_instance_name } else {}
        FederationRedirectUrl     = $CredentialResponse.Credentials.FederationRedirectUrl
    }

    $Output

    if ($Output.DomainTypeDescription -eq "FEDERATED") {
        Write-Warning "[$($Output.Username)] All users in a FEDERATED domain return VALID_USER by this endpoint. You must confirm with the system referenced in the FederationRedirectUrl.`n"
    }
}
