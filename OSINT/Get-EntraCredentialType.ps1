<#
.SYNOPSIS
    [Deprecated] Queries Entra ID credential type for a given UPN. Use Get-EntraCredentialInfo instead.

.DESCRIPTION
    This function is superseded by Get-EntraCredentialInfo (April 2024), which returns
    richer tenant and federation details. This script is retained for backward compatibility.

    See: https://github.com/Mike-Crowley/Public-Scripts/blob/main/OSINT/Get-EntraCredentialInfo.ps1

.PARAMETER Upn
    The email address (UPN) to query.

.EXAMPLE
    Get-EntraCredentialType -Upn user1@example.com

.NOTES
    Author: Mike Crowley
    https://mikecrowley.us

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
