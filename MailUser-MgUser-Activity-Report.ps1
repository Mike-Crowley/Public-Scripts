Connect-ExchangeOnline -UserPrincipalName <user>
Connect-MgGraph -TenantId <tenant>
Select-MgProfile beta

$MailUsers = Get-MailUser -filter {recipienttypedetails -eq 'MailUser'} -ResultSize unlimited
$Counter = 0
$ReportUsers = $MailUsers | ForEach-Object {
    $MailUser = $_

    $Counter++
    $percentComplete = (($Counter / $MailUsers.count) * 100)
    Write-Progress -Activity "Getting MG Objects" -PercentComplete $percentComplete -Status "$percentComplete% Complete:"    
       
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
    ) | Select-Object @(
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

$ReportUsers | Sort-Object UserPrincipalName | Export-Excel @Common_ExportExcelParams -Path ("c:\tmp\" + $filedate + "_report.xlsx") -WorksheetName report
