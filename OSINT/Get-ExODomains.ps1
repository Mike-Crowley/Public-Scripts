# Similar to: 
#  https://github.com/Mike-Crowley/Public-Scripts/blob/main/Get-AlternateMailboxes.ps1
#  "Get-TenantDomains" in this script: https://github.com/Gerenios/AADInternals/blob/ad071b736e2eb0f5f2b35df2920bb515b8c29a8c/AccessToken_utils.ps1
# Reported to Microsoft in 26Jul2018 who replied to say this is not a security issue, but a feature :)

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


Get-ExODomains -Domain example.com