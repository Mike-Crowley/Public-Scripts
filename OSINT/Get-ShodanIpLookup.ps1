<#
.SYNOPSIS
    Queries the Shodan database for information about an IP address.

.DESCRIPTION
    Performs IP address lookups using the Shodan internet search engine. Without an API key,
    queries the free InternetDB service (https://internetdb.shodan.io/) which returns hostnames,
    open ports, tags, and known vulnerabilities. With a paid API key, queries the full Shodan
    host API for additional detail including geolocation, ISP, ASN, and service banners.

.PARAMETER Ip
    The IP address to look up in Shodan.

.PARAMETER ApiKey
    A Shodan API key for paid service access. If omitted, the free InternetDB endpoint is used.

.EXAMPLE
    . .\Get-ShodanIpLookup.ps1
    Get-ShodanIpLookup -Ip 93.184.216.34

.EXAMPLE
    Get-ShodanIpLookup -Ip 93.184.216.34 -ApiKey "YourApiKeyHere"

.NOTES
    Author: Mike Crowley
    https://mikecrowley.us

.LINK
    https://internetdb.shodan.io/
.LINK
    https://developer.shodan.io/api
#>

function Get-ShodanIpLookup {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [ipaddress]$Ip,

        [Parameter()]
        [string]$ApiKey = ""
    )

    if ($ApiKey -eq "") {
        Invoke-RestMethod -Uri "https://internetdb.shodan.io/$Ip"
    }
    else {
        Invoke-RestMethod -Uri "https://api.shodan.io/shodan/host/$($Ip)?key=$ApiKey"
    }
}

# Get-ShodanIpLookup -Ip 93.184.216.34
