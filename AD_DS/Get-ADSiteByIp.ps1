<#
.SYNOPSIS
    Determines which Active Directory site an IP address belongs to.

.DESCRIPTION
    Queries the AD Subnets container via LDAP and uses the IP-Calc script from the
    PowerShell Gallery to match the provided IP address against defined site subnets.
    Returns the matching site name, subnet CIDR, and input IP. When multiple subnets
    match (overlapping ranges), the most specific (longest prefix) is returned.

    Requires the IP-Calc script from the PowerShell Gallery. Use -CheckForScript $true
    (the default) to auto-install it if missing.

.PARAMETER IP
    The IP address to look up against AD site subnets.

.PARAMETER CheckForScript
    When $true (default), checks for and installs the IP-Calc dependency from the
    PowerShell Gallery if not already present. Set to $false to skip this check.

.EXAMPLE
    . .\Get-ADSiteByIp.ps1
    Get-ADSiteByIp -IP 10.1.1.1

.EXAMPLE
    Get-ADSiteByIp -IP 10.1.1.1 -CheckForScript $false

.NOTES
    Author: Mike Crowley
    https://mikecrowley.us

    Requires: IP-Calc (https://www.powershellgallery.com/packages/IP-Calc)

.LINK
    https://github.com/Mike-Crowley/Public-Scripts
#>

Function Get-ADSiteByIp {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [ipaddress]$IP,

        [Parameter()]
        [bool]$CheckForScript = $true
    )

    if ($CheckForScript -eq $true) {
        try { Get-InstalledScript IP-Calc | Out-Null }
        catch {
            Install-Script -Name IP-Calc -Scope CurrentUser -Force -Confirm:$false
            Write-Output "Installing IP-Calc.ps1 from https://www.powershellgallery.com/packages/IP-Calc"
            try { Get-InstalledScript IP-Calc | Out-Null }
            catch {
                Write-Error "Unable to install required IP-Calc script"
                throw
            }
        }
    }

    $domainDN = ($env:USERDNSDOMAIN -split '\.' | % { 'dc=' + $_ } ) -join ','
    $ConfigDN = (([adsi]"LDAP://$DomainDN").objectcategory -split 'Schema,')[-1] # forest-wide support
    $Subnets = [ADSI]"LDAP://CN=Subnets,CN=Sites,$ConfigDN"

    $selectFilter = @(
        @{n = "SubnetFromSite"; e = { $_.name } }
        @{n = "Site"; e = { ($_.siteObject -split ',' -replace 'CN=', '')[0] } }
        @{n = "PrefixLength"; e = { (IP-Calc.ps1 -CIDR $_.name).PrefixLength } }
    )
    $siteSubnetDetail = $Subnets.Children | Select-Object $selectFilter

    ($siteSubnetDetail | Where-Object { (IP-Calc.ps1 $_.SubnetFromSite).Compare($IP) } | Sort-Object PrefixLength)[-1] |
        Select-Object @{n = "InputIP"; e = { $IP } }, Site, SubnetFromSite
}

# Get-ADSiteByIp -IP 10.1.1.1
