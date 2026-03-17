<#
.SYNOPSIS
    Sends email via Microsoft Graph using Send-MgUserMail with app-only authentication.

.DESCRIPTION
    Sends email with the Microsoft Graph PowerShell SDK (Send-MgUserMail),
    including HTML body content and file attachments. Uses certificate-based app-only
    authentication.

.PARAMETER EmailRecipients
    An array of recipient email addresses.

.PARAMETER EmailSender
    The sender email address (must have Send.Mail permission in the app registration).

.PARAMETER ClientId
    The Application (client) ID of the Entra ID app registration.

.PARAMETER TenantId
    The Tenant ID for the Entra ID tenant.

.PARAMETER CertificateThumbprint
    The certificate thumbprint for app-only authentication.

.PARAMETER AttachmentDirectory
    Directory path containing files to attach. Defaults to 'C:\tmp'.

.PARAMETER HtmlBodyPath
    Path to the HTML file used as the email body. Defaults to 'C:\tmp\HelloWorld.htm'.

.EXAMPLE
    .\MgUserMail.ps1 -EmailRecipients 'user1@domain.com','user2@domain.biz' -EmailSender 'me@domain.info' -ClientId '<appId>' -TenantId '<tenantId>' -CertificateThumbprint '<thumbprint>'

    Sends an email using the default attachment directory (C:\tmp) and HTML body path.

.EXAMPLE
    $params = @{
        EmailRecipients      = 'user1@domain.com', 'user2@domain.biz'
        EmailSender          = 'me@domain.info'
        ClientId             = '656d524e-fe4a-407a-9579-7e2be1a74a3c'
        TenantId             = 'contoso.onmicrosoft.com'
        CertificateThumbprint = 'A1B2C3D4E5F6...'
        AttachmentDirectory  = 'C:\Reports'
        HtmlBodyPath         = 'C:\Templates\Newsletter.htm'
    }
    .\MgUserMail.ps1 @params

    Sends an email with custom attachment directory and HTML body using splatting.

.NOTES
    Author: Mike Crowley
    https://mikecrowley.us

    Requires: Microsoft.Graph.Users.Actions module

.LINK
    https://mikecrowley.us/2021/10/27/sending-email-with-send-mgusermail-microsoft-graph-powershell

.LINK
    https://learn.microsoft.com/en-us/graph/api/user-sendmail
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string[]]$EmailRecipients,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$EmailSender,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$ClientId,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$TenantId,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$CertificateThumbprint,

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$AttachmentDirectory = 'C:\tmp',

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$HtmlBodyPath = 'C:\tmp\HelloWorld.htm'
)

#region 1: Setup

$emailSubject = "Sample Email | " + (Get-Date -UFormat %e%b%Y)

$MgConnectParams = @{
    ClientId              = $ClientId
    TenantId              = $TenantId
    CertificateThumbprint = $CertificateThumbprint
}

Function ConvertTo-IMicrosoftGraphRecipient {
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string[]]$SmtpAddresses
    )
    foreach ($address in $SmtpAddresses) {
        @{
            emailAddress = @{address = $address }
        }
    }
}

Function ConvertTo-IMicrosoftGraphAttachment {
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$UploadDirectory
    )
    $directoryContents = Get-ChildItem $UploadDirectory -Attributes !Directory -Recurse
    foreach ($file in $directoryContents) {
        $encodedAttachment = [convert]::ToBase64String((Get-Content $file.FullName -Encoding byte))
        @{
            "@odata.type" = "#microsoft.graph.fileAttachment"
            name          = ($File.FullName -split '\\')[-1]
            contentBytes  = $encodedAttachment
        }
    }
}

#endregion 1


#region 2: Run

[array]$toRecipients = ConvertTo-IMicrosoftGraphRecipient -SmtpAddresses $EmailRecipients

$attachments = ConvertTo-IMicrosoftGraphAttachment -UploadDirectory $AttachmentDirectory

$emailBody = @{
    ContentType = 'html'
    Content     = Get-Content $HtmlBodyPath
}

Connect-Graph @MgConnectParams
Select-MgProfile v1.0

$body = @{
    subject      = $emailSubject
    toRecipients = $toRecipients
    attachments  = $attachments
    body         = $emailBody
}

$bodyParameter = @{
    'message'         = $body
    'saveToSentItems' = $false
}

Send-MgUserMail -UserId $EmailSender -BodyParameter $bodyParameter

#endregion 2
