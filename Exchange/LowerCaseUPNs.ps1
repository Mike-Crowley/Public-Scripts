<#
.SYNOPSIS
    Converts uppercase UserPrincipalName values on remote mailboxes to lowercase.

.DESCRIPTION
    Scans all remote mailboxes in an Exchange on-premises organization and converts any UPN
    containing uppercase characters to lowercase. Creates an XML backup of affected objects
    on the desktop before making changes.

    Connects to an Exchange 2013+ server via remote PowerShell.

.PARAMETER ExchangeFQDN
    The fully qualified domain name of the Exchange server to connect to via remote PowerShell.

.EXAMPLE
    .\LowerCaseUPNs.ps1 -ExchangeFQDN exchange-admin.contoso.com

    Connects to the specified Exchange server, identifies remote mailboxes with
    uppercase UPN characters, backs them up, and converts each UPN to lowercase.

.NOTES
    Author: Mike Crowley
    https://mikecrowley.us

.LINK
    https://mikecrowley.us/2012/05/14/converting-smtp-proxy-addresses-to-lowercase/
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, HelpMessage = 'Type the FQDN of an Exchange server')]
    [ValidateNotNullOrEmpty()]
    [string]$ExchangeFQDN
)

#Connect to Exchange 2013+ Server
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$ExchangeFQDN/PowerShell/" -Authentication Kerberos
Import-PSSession $Session
Set-ADServerSettings -ViewEntireForest $true


$TargetObjects = Get-RemoteMailbox -ResultSize Unlimited | Where-Object { $_.UserPrincipalName.ToLower() -cne $_.UserPrincipalName }

Write-Host $TargetObjects.count "Remote mailboxes have one or more uppercase characters." -ForegroundColor Cyan

#Backup First
Function Get-FileFriendlyDate { Get-Date -format ddMMMyyyy_HHmm.s }
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