<#
.SYNOPSIS
    Queries the Outlook Mobile AutoDetect endpoint to discover mail protocols for a given address.

.DESCRIPTION
    Sends a request to the Outlook Mobile AutoDetect service
    (prod-autodetect.outlookmobile.com) and returns the available mail protocols
    for the specified email address. Useful for OSINT and troubleshooting mail configuration.

.PARAMETER Upn
    The email address (UPN) to query. Must be a valid SMTP address.

.EXAMPLE
    . .\Get-AutoDetect.ps1
    Get-AutoDetect -Upn user@example.com

.NOTES
    Author: Mike Crowley
    https://mikecrowley.us

.LINK
    https://github.com/Mike-Crowley/Public-Scripts
#>

function Get-AutoDetect {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Upn
    )

    try {
        $Response = Invoke-WebRequest -Uri "https://prod-autodetect.outlookmobile.com/autodetect/detect" -Headers @{"X-Email" = $Upn } -UseBasicParsing -ErrorAction Stop
    }
    catch {
        Write-Warning "Failed to retrieve autodetect information for $Upn : $($_.Exception.Message)"
        return
    }

    $Response.Content | ConvertFrom-Json | Select-Object -ExpandProperty protocols
}

# Get-AutoDetect -Upn user1@mikecrowley.us  # must be a valid smtp address