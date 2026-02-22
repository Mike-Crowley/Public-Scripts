<#
.SYNOPSIS
    Sends email via Microsoft Graph using Send-MgUserMail with app-only authentication.

.DESCRIPTION
    Demonstrates sending email with the Microsoft Graph PowerShell SDK (Send-MgUserMail),
    including HTML body content and file attachments. Uses certificate-based app-only
    authentication.

    Edit the configuration variables at the top of the script before running.

.EXAMPLE
    .\MgUserMail.ps1

    Sends an email after editing the recipient, sender, app registration, and attachment
    variables at the top of the script.

.NOTES
    Author: Mike Crowley
    https://mikecrowley.us

    Requires: Microsoft.Graph.Users.Actions module

.LINK
    https://mikecrowley.us/2021/10/27/sending-email-with-send-mgusermail-microsoft-graph-powershell

.LINK
    https://learn.microsoft.com/en-us/graph/api/user-sendmail
#>

#region 1: Setup

$emailRecipients = @(
    'user1@domain.com'
    'user2@domain.biz'
)
$emailSender = 'me@domain.info'

$emailSubject = "Sample Email | " + (Get-Date -UFormat %e%b%Y)

$MgConnectParams = @{
    ClientId              = '<your app>'
    TenantId              = '<your tenant id>'
    CertificateThumbprint = '<your thumbprint>'
}

Function ConvertTo-IMicrosoftGraphRecipient {
    [cmdletbinding()]
    Param(
        [array]$SmtpAddresses
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

[array]$toRecipients = ConvertTo-IMicrosoftGraphRecipient -SmtpAddresses $emailRecipients

$attachments = ConvertTo-IMicrosoftGraphAttachment -UploadDirectory C:\tmp

$emailBody = @{
    ContentType = 'html'
    Content     = Get-Content 'C:\tmp\HelloWorld.htm'
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

Send-MgUserMail -UserId $emailSender -BodyParameter $bodyParameter

#endregion 2
