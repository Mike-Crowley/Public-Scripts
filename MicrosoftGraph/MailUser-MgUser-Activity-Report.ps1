<#
.SYNOPSIS
    Generates a sign-in activity report for Exchange Online mail users via Microsoft Graph.

.DESCRIPTION
    Queries Exchange Online for MailUser recipients, then retrieves sign-in activity and
    profile details from Microsoft Graph (beta endpoint) for each user. Exports the results
    to an Excel file.

    Connects to both Exchange Online and Microsoft Graph automatically.

.PARAMETER ExchangeUpn
    The UPN to authenticate with for Connect-ExchangeOnline.

.PARAMETER TenantId
    The Tenant ID or domain for Connect-MgGraph.

.PARAMETER OutputDirectory
    Directory path for the output Excel report. Defaults to 'c:\tmp'.

.EXAMPLE
    .\MailUser-MgUser-Activity-Report.ps1 -ExchangeUpn admin@example.com -TenantId 'contoso.onmicrosoft.com'

    Connects to Exchange Online and Microsoft Graph, then generates the activity report
    in the default output directory (c:\tmp).

.EXAMPLE
    .\MailUser-MgUser-Activity-Report.ps1 -ExchangeUpn admin@example.com -TenantId 'contoso.onmicrosoft.com' -OutputDirectory 'C:\Reports'

    Same report with a custom output directory.

.NOTES
    Author: Mike Crowley
    https://mikecrowley.us

    Requires: ExchangeOnlineManagement, Microsoft.Graph.Users, ImportExcel modules
    Permissions: User.Read.All (Graph beta)

.LINK
    https://github.com/Mike-Crowley/Public-Scripts
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$ExchangeUpn,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$TenantId,

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$OutputDirectory = 'c:\tmp'
)

Connect-ExchangeOnline -UserPrincipalName $ExchangeUpn
Connect-MgGraph -TenantId $TenantId
Select-MgProfile beta

$MailUsers = Get-MailUser -filter { recipienttypedetails -eq 'MailUser' } -ResultSize unlimited
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
        @{n = 'PrimarySmtpAddress'; e = { $MailUser.PrimarySmtpAddress } }
        @{n = 'ExternalEmailAddress'; e = { $MailUser.ExternalEmailAddress } }
        @{n = 'LastSignInDateTime'; e = { [datetime]$_.SignInActivity.LastSignInDateTime } }
        @{n = 'lastNonInteractiveSignInDateTime'; e = { [datetime]$_.SignInActivity.AdditionalProperties.lastNonInteractiveSignInDateTime } }
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

$ReportUsers | Sort-Object UserPrincipalName | Export-Excel @Common_ExportExcelParams -Path ("$OutputDirectory\" + $filedate + "_report.xlsx") -WorksheetName report
