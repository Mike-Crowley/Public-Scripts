<#
    Features:
        1) This script Creates a TXT and CSV file with the following information:
            a) TXT file: Recipient Address Statistics
            b) CSV file: Output of everyone's SMTP proxy addresses.

    Instructions:
        1) Run this from "regular" PowerShell.  Exchange Management Shell may cause problems, especially in Exchange 2010, due to PSv2.
        2) Usage: RecipientReportv5.ps1 server5.domain.local

    Requirements:
        1) Exchange 2010 or 2013
        2) PowerShell 4.0

    
    April 4 2015
    Mike Crowley
    
    https://BaselineTechnologies.com    
#>

param(
    [parameter(
        Position = 0,
        Mandatory = $true,
        ValueFromPipeline = $false,
        HelpMessage = 'Type the name of a Client Access Server'
    )]
    [string]$ExchangeFQDN
)

if ($host.version.major -lt 3) {
    Write-Host ""
    Write-Host "This script requires PowerShell 3.0 or later." -ForegroundColor Red
    Write-Host "Note: Exchange 2010's EMC always runs as version 2.  Perhaps try launching PowerShell normally." -ForegroundColor Red
    Write-Host ""
    Write-Host "Exiting..." -ForegroundColor Red
    Start-Sleep 3
    Exit
}

if ((Test-Connection $ExchangeFQDN -Count 1 -Quiet) -ne $true) {
    Write-Host ""
    Write-Host ("Cannot connect to: " + $ExchangeFQDN) -ForegroundColor Red
    Write-Host ""
    Write-Host "Exiting..." -ForegroundColor Red
    Start-Sleep 3
    Exit
}

Clear-Host

# Misc variables
$ReportTimeStamp = (Get-Date -Format s) -replace ":", "."
$TxtFile = "$env:USERPROFILE\Desktop\" + $ReportTimeStamp + "_RecipientAddressReport_Part_1of2.txt"
$CsvFile = "$env:USERPROFILE\Desktop\" + $ReportTimeStamp + "_RecipientAddressReport_Part_2of2.csv"

# Connect to Exchange
Write-Host ("Connecting to " + $ExchangeFQDN + "...") -ForegroundColor Cyan
Get-PSSession | Where-Object { $_.ConfigurationName -eq 'Microsoft.Exchange' } | Remove-PSSession
$Session = @{
    ConfigurationName = 'Microsoft.Exchange'
    ConnectionUri     = 'http://' + $ExchangeFQDN + '/PowerShell/?SerializationLevel=Full' 
    Authentication    = 'Kerberos'
}
Import-PSSession (New-PSSession @Session) 

# Get Data
Write-Host "Getting data from Exchange..." -ForegroundColor Cyan
$AcceptedDomains = Get-AcceptedDomain
$InScopeRecipients = @(
    'DynamicDistributionGroup'
    'UserMailbox'
    'MailUniversalDistributionGroup'
    'MailUniversalSecurityGroup'
    'MailNonUniversalGroup'
    'PublicFolder'
)
$AllRecipients = Get-Recipient -recipienttype $InScopeRecipients -ResultSize unlimited | Select-Object name, emailaddresses, RecipientType
$UniqueRecipientDomains = ($AllRecipients.emailaddresses | Where-Object { $_ -like 'smtp*' }) -split '@' | Where-Object { $_ -NotLike 'smtp:*' } | Select-Object -Unique

Write-Host "Preparing Output 1 of 2..." -ForegroundColor Cyan
# Output address stats
$TextBlock = @(
    "Total Number of Recipients: " + $AllRecipients.Count
    "Number of Dynamic Distribution Groups: " + ($AllRecipients | Where-Object { $_.RecipientType -eq 'DynamicDistributionGroup' }).Count
    "Number of User Mailboxes: " + ($AllRecipients | Where-Object { $_.RecipientType -eq 'UserMailbox' }).Count
    "Number of Mail-Universal Distribution Groups: " + ($AllRecipients | Where-Object { $_.RecipientType -eq 'MailUniversalDistributionGroup' }).Count
    "Number of Mail-UniversalSecurity Groups: " + ($AllRecipients | Where-Object { $_.RecipientType -eq 'MailUniversalSecurityGroup' }).Count
    "Number of Mail-NonUniversal Groups: " + ($AllRecipients | Where-Object { $_.RecipientType -eq 'MailNonUniversalGroup' }).Count
    "Number of Public Folders: " + ($AllRecipients | Where-Object { $_.RecipientType -eq 'PublicFolder' }).Count
    ""
    "Number of Accepted Domains: " + $AcceptedDomains.count 
    ""
    "Number of domains found on recipients: " + $UniqueRecipientDomains.count 
    ""
    $DomainComparrison = Compare-Object $AcceptedDomains.DomainName $UniqueRecipientDomains
    "These domains have been assigned to recipients, but are not Accepted Domains in the Exchange Organization:"
    ($DomainComparrison | Where-Object { $_.SideIndicator -eq '=>' }).InputObject 
    ""
    "These Accepted Domains are not assigned to any recipients:" 
    ($DomainComparrison | Where-Object { $_.SideIndicator -eq '<=' }).InputObject
    ""
    "See this CSV for a complete listing of all addresses: " + $CsvFile
)

Write-Host "Preparing Output 2 of 2..." -ForegroundColor Cyan

$RecipientsAndSMTPProxies = @()
$CounterWatermark = 1
 
$AllRecipients | ForEach-Object {
    
    # Create a new placeholder object
    $RecipientOutputObject = New-Object PSObject -Property @{
        Name          = $_.Name
        RecipientType = $_.RecipientType
        SMTPAddress0  = ($_.emailaddresses | Where-Object { $_ -clike 'SMTP:*' } ) -replace "SMTP:"
    }    
    
    # If applicable, get a list of other addresses for the recipient
    if (($_.emailaddresses).count -gt '1') {       
        $OtherAddresses = @()
        $OtherAddresses = ($_.emailaddresses | Where-Object { $_ -clike 'smtp:*' } ) -replace "smtp:"
        
        $Counter = $OtherAddresses.count
        if ($Counter -gt $CounterWatermark) { $CounterWatermark = $Counter }
        $OtherAddresses | ForEach-Object {
            $RecipientOutputObject | Add-Member -MemberType NoteProperty -Name (“SmtpAddress” + $Counter) -Value ($_ -replace "smtp:")
            $Counter--
        }
    }
    $RecipientsAndSMTPProxies += $RecipientOutputObject
}
  
$AttributeList = @(
    'Name'
    'RecipientType'
)
$AttributeList += 0..$CounterWatermark | ForEach-Object { "SMTPAddress" + $_ }


Write-Host "Saving report files to your desktop:" -ForegroundColor Green
Write-Host ""
Write-Host $TxtFile -ForegroundColor Green
Write-Host $CsvFile -ForegroundColor Green

$TextBlock | Out-File $TxtFile
$RecipientsAndSMTPProxies | Select-Object $AttributeList | Sort-Object RecipientType, Name | Export-CSV $CsvFile -NoTypeInformation

Write-Host ""
Write-Host ""
Write-Host "Report Complete!" -ForegroundColor Green