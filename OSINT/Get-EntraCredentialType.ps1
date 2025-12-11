<#
.SYNOPSIS
    Misc OSINT
    Superseded by Get-EntraCredentialInfo 12 Apr 2024
    https://github.com/Mike-Crowley/Public-Scripts/blob/main/OSINT/Get-EntraCredentialInfo.ps1

.EXAMPLE

    Get-EntraCredentialType -Upn user1@domain.com

.LINK

    https://mikecrowley.us
#>
function Get-EntraCredentialType {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Upn
    )
    $Body = @{
        username            = $Upn
        isOtherIdpSupported = $true
    }
    $Body = $Body | ConvertTo-Json -Compress

    try {
        $Response = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/common/GetCredentialType" -Body $Body -ContentType "application/json" -ErrorAction Stop
    }
    catch {
        Write-Warning "Failed to retrieve credential type for $Upn : $($_.Exception.Message)"
        return
    }

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
