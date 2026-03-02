<#
.SYNOPSIS
    Queries the AlternateMailboxes node from a user's Exchange AutoDiscover response (Modern Auth).

.DESCRIPTION
    Queries the AlternateMailboxes node within a user's AutoDiscover response using Modern Auth
    via MSAL.PS. For the Basic Auth version, see Get-AlternateMailboxes_BasicAuth.ps1.

    Requirements:
        1) Install the MSAL.PS module (Install-Module MSAL.PS)
        2) Register an app in the target tenant
        3) Configure API permissions: Office 365 Exchange Online > User.Read.All
           (under "APIs my organization uses", NOT Microsoft Graph)
        4) Optionally use a certificate for app-only authentication

.PARAMETER SMTPAddress
    The SMTP email address of the user to query.

.PARAMETER MsalToken
    An MSAL token object from Get-MsalToken with Exchange Online scopes.

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

.NOTES
    Author: Mike Crowley
    https://mikecrowley.us

    Requires: MSAL.PS module

.LINK
    https://mikecrowley.us/2017/12/08/querying-msexchdelegatelistlink-in-exchange-online-with-powershell/
#>

Function Get-AlternateMailboxes {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$SMTPAddress,

        [Parameter(Mandatory = $true)]
        [Microsoft.Identity.Client.AuthenticationResult]$MsalToken
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

  $WebResponse = Invoke-WebRequest https://autodiscover-s.outlook.com/autodiscover/autodiscover.svc -Method Post -Body $AutoDiscoverRequest -ContentType 'text/xml; charset=utf-8' -Headers $Headers
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

