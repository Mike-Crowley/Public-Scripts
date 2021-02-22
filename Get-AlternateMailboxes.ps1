<#

.SYNOPSIS 
    This function queries the AlternateMailboxes node within a user's AutoDiscover response. See the link for details.

    Version: December 8a, 2017


.DESCRIPTION
    This function queries the AlternateMailboxes node within a user's AutoDiscover response. See the link for details.
  
    Author:
    Mike Crowley
    https://BaselineTechnologies.com

 .EXAMPLE
 
    Get-AlternateMailboxes -SMTPAddress mike@contoso.com -Credential (Get-Credential)
 
.LINK
    https://mikecrowley.us/2017/12/08/querying-msexchdelegatelistlink-in-exchange-online-with-powershell/

#>

Function Get-AlternateMailboxes {

    Param(
        [parameter(Mandatory=$true)]
        $SMTPAddress,
        [parameter(Mandatory=$true)]
        $Credential
        )

        $AutoDiscoverRequest = @"
        <soap:Envelope xmlns:a="http://schemas.microsoft.com/exchange/2010/Autodiscover" 
                xmlns:wsa="http://www.w3.org/2005/08/addressing" 
                xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
                xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
          <soap:Header>
            <a:RequestedServerVersion>Exchange2013</a:RequestedServerVersion>
            <wsa:Action>http://schemas.microsoft.com/exchange/2010/Autodiscover/Autodiscover/GetUserSettings</wsa:Action>
            <wsa:To>https://autodiscover.exchange.microsoft.com/autodiscover/autodiscover.svc</wsa:To>
          </soap:Header>
          <soap:Body>
            <a:GetUserSettingsRequestMessage xmlns:a="http://schemas.microsoft.com/exchange/2010/Autodiscover">
              <a:Request>
                <a:Users>
                  <a:User>
                    <a:Mailbox>$SMTPAddress</a:Mailbox>
                  </a:User>
                </a:Users>
                <a:RequestedSettings>
                  <a:Setting>UserDisplayName</a:Setting>
                  <a:Setting>UserDN</a:Setting>
                  <a:Setting>UserDeploymentId</a:Setting>
                  <a:Setting>MailboxDN</a:Setting>
                  <a:Setting>AlternateMailboxes</a:Setting>
                </a:RequestedSettings>
              </a:Request>
            </a:GetUserSettingsRequestMessage>
          </soap:Body>
        </soap:Envelope>
"@
    #Other attributes available here: https://msdn.microsoft.com/en-us/library/microsoft.exchange.webservices.autodiscover.usersettingname(v=exchg.80).aspx

    $WebResponse = Invoke-WebRequest https://autodiscover-s.outlook.com/autodiscover/autodiscover.svc -Credential $Credential -Method Post -Body $AutoDiscoverRequest -ContentType 'text/xml; charset=utf-8'
    [System.Xml.XmlDocument]$XMLResponse = $WebResponse.Content
    $RequestedSettings = $XMLResponse.Envelope.Body.GetUserSettingsResponseMessage.Response.UserResponses.UserResponse.UserSettings.UserSetting
    return $RequestedSettings.AlternateMailboxes.AlternateMailbox
}


Get-AlternateMailboxes 