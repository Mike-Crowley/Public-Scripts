#3Jan2018

#related to script #2 here:
# https://mikecrowley.us/2012/05/14/converting-smtp-proxy-addresses-to-lowercase/

#Connect to Exchange 2013+ Server
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://exchange-admin/PowerShell/ -Authentication Kerberos
Import-PSSession $Session
Set-ADServerSettings -ViewEntireForest $true


$TargetObjects = Get-RemoteMailbox -ResultSize Unlimited | Where {$_.UserPrincipalName.ToLower() -cne $_.UserPrincipalName}

Write-Host $TargetObjects.count "Remote mailboxes have one or more uppercase characters." -ForegroundColor Cyan

#Backup First
Function Get-FileFriendlyDate {Get-Date -format ddMMMyyyy_HHmm.s}
$DesktopPath = ([Environment]::GetFolderPath("Desktop") + '\')
$LogPath = ($DesktopPath + (Get-FileFriendlyDate) + "-UppercaseBackup.xml")

try {
    $TargetObjects | Select-Object DistinguishedName, PrimarySMTPAddress, UserPrincipalName | Export-Clixml $LogPath -ErrorAction Stop
    Write-Host "A backup XML has been placed here:" $LogPath -ForegroundColor Cyan
}
catch {
    Write-Host "Failed to create backup file: $_" -ForegroundColor Red
    Write-Host "Exiting to prevent data loss." -ForegroundColor Red
    return
}
Write-Host

$Counter = $TargetObjects.Count

foreach ($RemoteMailbox in $TargetObjects) {
    
    Write-Host "Setting: " -ForegroundColor DarkCyan -NoNewline
    Write-Host $RemoteMailbox.PrimarySmtpAddress -ForegroundColor Cyan
    Write-Host "Remaining: " -ForegroundColor DarkCyan -NoNewline
    Write-Host $Counter -ForegroundColor Cyan

    Set-RemoteMailbox $RemoteMailbox.Identity -UserPrincipalName ("TMP-Rename-" + $RemoteMailbox.UserPrincipalName)
    Set-RemoteMailbox $RemoteMailbox.Identity -UserPrincipalName $RemoteMailbox.UserPrincipalName.ToLower()
    

    $Counter --    
}

Write-Host
Write-Host "Done." -ForegroundColor DarkCyan


#End