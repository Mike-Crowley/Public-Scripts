#region 1: Setup
    $emailRecipients   = @(
        'user1@domain.com'
        'user2@domain.biz'
    )
    $emailSender  = 'me@domain.info'

    $emailSubject = "Sample Email | " + (Get-Date -UFormat %e%b%Y)

    $MgConnectParams = @{
        ClientId              = '<your app>'
        TenantId              = '<your tenant id>'
        CertificateThumbprint = '<your thumbprint>'
    }

    Function ConvertTo-IMicrosoftGraphRecipient {
        [cmdletbinding()]
        Param(
            [array]$smtpAddresses        
        )
        foreach ($address in $smtpAddresses) {
            @{
                emailAddress = @{address = $address}
            }    
        }    
    }

    Function ConvertTo-IMicrosoftGraphAttachment {
        [cmdletbinding()]
        Param(
            [array]$UploadDirectory        
        )
        $DirectoryContents = Get-ChildItem c:\tmp -Attributes !Directory -Recurse
        foreach ($File in $DirectoryContents) {
            $EncodedAttachment = [convert]::ToBase64String((Get-Content $File.FullName -Encoding byte))
            @{
                "@odata.type"= "#microsoft.graph.fileAttachment"
                name = ($File.FullName -split '\\')[-1]
                contentBytes = $EncodedAttachment
            }   
        }    
    }

#endregion 1

#region 2: Run
    $toRecipients = ConvertTo-IMicrosoftGraphRecipient -smtpAddresses $emailRecipients 

    $attachments = ConvertTo-IMicrosoftGraphAttachment -UploadDirectory C:\tmp

    $emailBody  = @{
        ContentType = 'html'
        Content = Get-Content 'C:\tmp\HelloWorld.htm'    
    }

    Connect-Graph @MgConnectParams
    Select-MgProfile v1.0
    
    $body += @{subject      = $emailSubject}
    $body += @{Body         = $emailBody}
    $body += @{toRecipients = $toRecipients}
    $body += @{attachments  = $attachments}

    $BodyParameter += @{'message' = $body}
    $BodyParameter += @{'saveToSentItems' = $false}

    Send-MgUserMail -UserId $EmailSender -BodyParameter $BodyParameter
#endregion 2


