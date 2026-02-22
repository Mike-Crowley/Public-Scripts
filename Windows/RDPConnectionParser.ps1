<#
.SYNOPSIS
    Parses RDP and console session events from remote servers into a CSV report.

.DESCRIPTION
    Reads the "Microsoft-Windows-TerminalServices-LocalSessionManager/Operational" event log
    from one or more servers and outputs logon, logoff, disconnect, and reconnection events
    to a CSV file on the desktop.

    Despite the log's name, it includes both RDP and regular console logins.

.PARAMETER ServersToQuery
    An array of server names to query. Defaults to the local hostname.

.PARAMETER StartTime
    The earliest event timestamp to include. Defaults to January 1, 1970 (all events).

.EXAMPLE
    .\RDPConnectionParser.ps1 -ServersToQuery Server1, Server2 -StartTime "November 1"

.NOTES
    Author: Mike Crowley
    https://mikecrowley.us

.LINK
    https://mikecrowley.us/tag/powershell
#>

Param(
    [array]$ServersToQuery = (hostname),
    [datetime]$StartTime = "January 1, 1970"
)

foreach ($Server in $ServersToQuery) {

    $LogFilter = @{
        LogName   = 'Microsoft-Windows-TerminalServices-LocalSessionManager/Operational'
        ID        = 21, 23, 24, 25
        StartTime = $StartTime
    }

    $AllEntries = Get-WinEvent -FilterHashtable $LogFilter -ComputerName $Server

    $AllEntries | ForEach-Object {
        $entry = [xml]$_.ToXml()
        [array]$Output += New-Object PSObject -Property @{
            TimeCreated = $_.TimeCreated
            User        = $entry.Event.UserData.EventXML.User
            IPAddress   = $entry.Event.UserData.EventXML.Address
            EventID     = $entry.Event.System.EventID
            ServerName  = $Server
        }
    }

}

$FilteredOutput += $Output | Select-Object TimeCreated, User, ServerName, IPAddress, @{Name = 'Action'; Expression = {
        if ($_.EventID -eq '21') { "logon" }
        if ($_.EventID -eq '22') { "Shell start" }
        if ($_.EventID -eq '23') { "logoff" }
        if ($_.EventID -eq '24') { "disconnected" }
        if ($_.EventID -eq '25') { "reconnection" }
    }
}

$Date = Get-Date -Format 'yyyyMMdd_HHmmss'
$FilePath = "$env:USERPROFILE\Desktop\$Date`_RDP_Report.csv"

try {
    $FilteredOutput | Sort-Object TimeCreated | Export-Csv $FilePath -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
    Write-Host "Writing File: $FilePath" -ForegroundColor Cyan
    Write-Host "Done!" -ForegroundColor Cyan
}
catch {
    Write-Host "Failed to write report file: $_" -ForegroundColor Red
}


#End
