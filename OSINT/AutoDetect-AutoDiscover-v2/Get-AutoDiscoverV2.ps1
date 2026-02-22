<#
.SYNOPSIS
    Queries the Exchange Online AutoDiscover V2 endpoint for a given email address.

.DESCRIPTION
    Sends a request to the Office 365 AutoDiscover V2 JSON endpoint and returns the
    protocol configuration for the specified email address. Useful for OSINT and
    troubleshooting Exchange Online connectivity.

.PARAMETER Upn
    The email address (UPN) to query. Must be a valid SMTP address.

.EXAMPLE
    . .\Get-AutoDiscoverV2.ps1
    Get-AutoDiscoverV2 -Upn user@example.com

.NOTES
    Author: Mike Crowley
    https://mikecrowley.us

.LINK
    https://github.com/Mike-Crowley/Public-Scripts
#>

function Get-AutoDiscoverV2 {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Upn
    )

    try {
        $Response = Invoke-WebRequest -Uri "https://outlook.office365.com/autodiscover/autodiscover.json/v1.0/$($Upn)?Protocol=activesync" -UseBasicParsing -ErrorAction Stop  # or change Protocol to ews
    }
    catch {
        Write-Warning "Failed to retrieve autodiscover information for $Upn : $($_.Exception.Message)"
        return
    }

    $Response.Content | ConvertFrom-Json
}

# Get-AutoDiscoverV2 -Upn user1@mikecrowley.us # must be a valid smtp address