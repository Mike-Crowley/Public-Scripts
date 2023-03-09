<#

.SYNOPSIS 
  This function queries the AlternateMailboxes node within a user's AutoDiscover response. This version now supports Modern Auth. For the basic Auth version of this script, use  Get-AlternateMailboxes_BasicAuth.ps1.
  
  Requirements:

  1) Install the MSAL.PS PowerShell Module (Install-Module MSAL.PS)
  2) Register an app in the target tenant
  3) Configure the API permissions on the app you just created
    a) Go to your app registration in the portal
    b) Click API permissions on the left
    c) Click Add permission
    d) Click "APIs my organization uses" (NOT GRAPH!)
    e) Type "Office 365 Exchange Online" in the search box
    f) Select the following permission:        
        User.Read.All
  4) Optionally: Use a certificate for application-based authentication, which is what the example below uses. Otherwise, you can use the different auth mentioned by Microsoft in the links below.

  Further reading:

    Use app-only authentication with the Microsoft Graph PowerShell SDK
    https://learn.microsoft.com/en-us/powershell/microsoftgraph/app-only?tabs=azure-portal&view=graph-powershell-1.0

    Create a self-signed public certificate to authenticate your application
    https://learn.microsoft.com/en-us/azure/active-directory/develop/howto-create-self-signed-certificate
    

  Version: March 9, 2023


.DESCRIPTION
  This function queries the AlternateMailboxes node within a user's AutoDiscover response. See the link for details.

  Author:
  Mike Crowley
  https://BaselineTechnologies.com

.EXAMPLE

  $TokenParams = @{    
    ClientId          = '656d524e-fe4a-407a-9579-7e2be1a74a3c'
    TenantId          = 'example.com'
    ClientCertificate = Get-Item Cert:\CurrentUser\My\<Your Cert Thumbprint>
    CorrelationId     = New-Guid
    Scopes            = 'https://outlook.office365.com/.default'
  }

  $MsalToken = Get-MsalToken @TokenParams 

  Get-AlternateMailboxes -SMTPAddress mike@example.com -MsalToken $MsalToken

.LINK
  https://mikecrowley.us/2017/12/08/querying-msexchdelegatelistlink-in-exchange-online-with-powershell/

#>

Function Get-AlternateMailboxes {

  Param(
    [parameter(Mandatory = $true)][string]
    [string]$SMTPAddress,
    [parameter(Mandatory = $true)][Microsoft.Identity.Client.AuthenticationResult]
    $MsalToken
  )
  try {
    Get-Module MSAL.PS -ListAvailable
  }
  catch {
    Write-Error "You must first install the MSAL.PS module (Install-Module MSAL.PS)."
    throw
  }
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
  # Other attributes available here: https://learn.microsoft.com/en-us/dotnet/api/microsoft.exchange.webservices.autodiscover.usersettingname?view=exchange-ews-api

  $Headers = @{
    'X-AnchorMailbox' = $SMTPAddress
    'Authorization'   = "Bearer $($MsalToken.AccessToken)"
  }

  $WebResponse = Invoke-WebRequest https://autodiscover-s.outlook.com/autodiscover/autodiscover.svc -Credential $Credential -Method Post -Body $AutoDiscoverRequest -ContentType 'text/xml; charset=utf-8' -Headers $Headers
  [System.Xml.XmlDocument]$XMLResponse = $WebResponse.Content
  $RequestedSettings = $XMLResponse.Envelope.Body.GetUserSettingsResponseMessage.Response.UserResponses.UserResponse.UserSettings.UserSetting
  return $RequestedSettings.AlternateMailboxes.AlternateMailbox
}


########## Example ##########

$TokenParams = @{    
  ClientId          = '656d524e-fe4a-407a-9579-7e2be1a74a3c'
  TenantId          = 'example.com'
  ClientCertificate = Get-Item Cert:\CurrentUser\My\<Your Cert Thumbprint>
  CorrelationId     = New-Guid
  Scopes            = 'https://outlook.office365.com/.default'
}
$MsalToken = Get-MsalToken @TokenParams 

Get-AlternateMailboxes -SMTPAddress 'mike@example.com' -MsalToken $MsalToken

########## Example ##########

