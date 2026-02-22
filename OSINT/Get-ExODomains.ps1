<#
.SYNOPSIS
    Enumerates the domains associated with a Microsoft 365 tenant via Exchange federation metadata.

.DESCRIPTION
    Queries the Exchange Online GetFederationInformation SOAP endpoint to retrieve all domains
    associated with a tenant. This is an unauthenticated OSINT technique that uses the
    Exchange AutoDiscover service.

    Reported to Microsoft in July 2018 who confirmed this is by-design behavior, not a
    security vulnerability.

.PARAMETER Domain
    A domain known to belong to the target Microsoft 365 tenant.

.EXAMPLE
    . .\Get-ExODomains.ps1
    Get-ExODomains -Domain example.com

.NOTES
    Author: Mike Crowley
    https://mikecrowley.us

    Related:
        https://github.com/Mike-Crowley/Public-Scripts/blob/main/Exchange/Get-AlternateMailboxes.ps1

.LINK
    https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/getfederationinformation-operation-soap
#>

function Get-ExODomains {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Domain
    )
    # https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/getfederationinformation-operation-soap
    $body = @"
        <soap:Envelope xmlns:exm="http://schemas.microsoft.com/exchange/services/2006/messages"
            xmlns:ext="http://schemas.microsoft.com/exchange/services/2006/types"
            xmlns:a="http://www.w3.org/2005/08/addressing"
            xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
            xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
            xmlns:xsd="http://www.w3.org/2001/XMLSchema">
            <soap:Header>
                <a:Action soap:mustUnderstand="1">http://schemas.microsoft.com/exchange/2010/Autodiscover/Autodiscover/GetFederationInformation</a:Action>
                <a:To soap:mustUnderstand="1">https://autodiscover-s.outlook.com/autodiscover/autodiscover.svc</a:To>
            </soap:Header>
            <soap:Body>
                <GetFederationInformationRequestMessage xmlns="http://schemas.microsoft.com/exchange/2010/Autodiscover">
                    <Request>
                        <Domain>$Domain</Domain>
                    </Request>
                </GetFederationInformationRequestMessage>
            </soap:Body>
        </soap:Envelope>
"@
    $headers = @{
        "SOAPAction" = '"http://schemas.microsoft.com/exchange/2010/Autodiscover/Autodiscover/GetFederationInformation"'
    }

    try {
        $response = Invoke-RestMethod -Method Post -Uri "https://autodiscover-s.outlook.com/autodiscover/autodiscover.svc" -Body $body -Headers $headers -UserAgent "AutodiscoverClient" -ContentType "text/xml; charset=utf-8" -ErrorAction Stop
    }
    catch {
        Write-Warning "Failed to retrieve federation information for $Domain : $($_.Exception.Message)"
        return
    }

    $response.Envelope.body.GetFederationInformationResponseMessage.response.Domains.Domain | Sort-Object
}

# Get-ExODomains -Domain example.com