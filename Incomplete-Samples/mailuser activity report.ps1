Connect-ExchangeOnline -UserPrincipalName <user>
Connect-MgGraph -TenantId <tenant>
Select-MgProfile beta

$Counter = 1

$MailUsers = Get-MailUser -filter {recipienttypedetails -eq 'MailUser'} -ResultSize unlimited
$ReportUsers = $MailUsers | foreach {
    $MailUser = $_   
    
    $percentComplete = (($Counter / $MailUsers.count) * 100)
    Write-Progress -Activity "Getting MG Onjects" -PercentComplete $percentComplete -Status "$percentComplete% Complete:"
    $Counter ++
       
       #to do - organize properties better
    Get-MgUser -UserId $MailUser.ExternalDirectoryObjectId -ConsistencyLevel eventual -Property @(
        'UserPrincipalName'
        'SignInActivity'
        'CreatedDateTime'        
        'DisplayName'
        'Mail'
        'OnPremisesImmutableId'
        'OnPremisesDistinguishedName'
        'OnPremisesLastSyncDateTime'
        'SignInSessionsValidFromDateTime'
        'RefreshTokensValidFromDateTime'
        'id'
    ) | select @(
        'UserPrincipalName'        
        'CreatedDateTime'        
        'DisplayName'
        'Mail'
        'OnPremisesImmutableId'
        'OnPremisesDistinguishedName'
        'OnPremisesLastSyncDateTime'
        'SignInSessionsValidFromDateTime'
        'RefreshTokensValidFromDateTime'
        'id'        
        @{n='PrimarySmtpAddress'; e={$MailUser.PrimarySmtpAddress}}
        @{n='ExternalEmailAddress'; e={$MailUser.ExternalEmailAddress}}
        @{n='LastSignInDateTime'; e={[datetime]$_.SignInActivity.LastSignInDateTime}}
        @{n='lastNonInteractiveSignInDateTime'; e={[datetime]$_.SignInActivity.AdditionalProperties.lastNonInteractiveSignInDateTime}}           
    )
}

$Common_ExportExcelParams = @{
    # PassThru     = $true
    BoldTopRow   = $true
    AutoSize     = $true
    AutoFilter   = $true
    FreezeTopRow = $true
}

$FileDate = Get-Date -Format yyyyMMddTHHmmss

$ReportUsers | sort DisplayName_Agency, UserPrincipalName | Export-Excel @Common_ExportExcelParams -Path ("c:\tmp\" + $filedate + "_report.xlsx") -WorksheetName report



