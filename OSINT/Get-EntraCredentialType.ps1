<#
.SYNOPSIS
    Misc OSINT
    Superseded by Get-EntraCredentialInfo 12 Apr 2024

.EXAMPLE

    Get-EntraCredentialType -Upn user1@domain.com

.LINK

    https://mikecrowley.us
#>
Function Get-EntraCredentialType {
    param (
        [parameter(Mandatory = $true)][string]
        $Upn
    )
    $Body = @{
        username            = $Upn
        isOtherIdpSupported = $true
    }
    $Body = $Body | ConvertTo-Json -Compress
    $Response = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/common/GetCredentialType" -Body $Body
    
    [pscustomobject]@{
        Username              = $Response.Username
        PrefCredential        = $Response.Credentials.PrefCredential
        DomainFound           = $Response.IfExistsResult -eq 0
        FederatedDomain       = $null -ne $Response.Credentials.FederationRedirectUrl
        FederationRedirectUrl = $Response.Credentials.FederationRedirectUrl
        DesktopSsoEnabled     = $Response.EstsProperties.DesktopSsoEnabled
        UserTenantBranding    = $Response.EstsProperties.UserTenantBranding
    }
    
}
