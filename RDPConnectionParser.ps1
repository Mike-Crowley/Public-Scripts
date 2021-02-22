<#

.SYNOPSIS 
    This script reads the event log "Microsoft-Windows-TerminalServices-LocalSessionManager/Operational" from 
    multiple servers and outputs the human-readable results to a CSV.  This data is not filterable in the native 
    Windows Event Viewer.

    Version: November 9, 2016


.DESCRIPTION
    This script reads the event log "Microsoft-Windows-TerminalServices-LocalSessionManager/Operational" from 
    multiple servers and outputs the human-readable results to a CSV.  This data is not filterable in the native 
    Windows Event Viewer.

    NOTE: Despite this log's name, it includes both RDP logins as well as regular console logins too.
    
    Author:
    Mike Crowley
    https://BaselineTechnologies.com

 .EXAMPLE
 
    .\RDPConnectionParser.ps1 -ServersToQuery Server1, Server2 -StartTime "November 1"
 
.LINK
    https://MikeCrowley.us/tag/powershell

#>

Param(
    [array]$ServersToQuery = (hostname),
    [datetime]$StartTime = "January 1, 1970"
)

    foreach ($Server in $ServersToQuery) {

        $LogFilter = @{
            LogName = 'Microsoft-Windows-TerminalServices-LocalSessionManager/Operational'
            ID = 21, 23, 24, 25
            StartTime = $StartTime
            }

        $AllEntries = Get-WinEvent -FilterHashtable $LogFilter -ComputerName $Server

        $AllEntries | Foreach { 
            $entry = [xml]$_.ToXml()
            [array]$Output += New-Object PSObject -Property @{
                TimeCreated = $_.TimeCreated
                User = $entry.Event.UserData.EventXML.User
                IPAddress = $entry.Event.UserData.EventXML.Address
                EventID = $entry.Event.System.EventID
                ServerName = $Server
                }        
            } 

    }

    $FilteredOutput += $Output | Select TimeCreated, User, ServerName, IPAddress, @{Name='Action';Expression={
                if ($_.EventID -eq '21'){"logon"}
                if ($_.EventID -eq '22'){"Shell start"}
                if ($_.EventID -eq '23'){"logoff"}
                if ($_.EventID -eq '24'){"disconnected"}
                if ($_.EventID -eq '25'){"reconnection"}
                }
            }

    $Date = (Get-Date -Format s) -replace ":", "."
    $FilePath = "$env:USERPROFILE\Desktop\$Date`_RDP_Report.csv"
    $FilteredOutput | Sort TimeCreated | Export-Csv $FilePath -NoTypeInformation

Write-host "Writing File: $FilePath" -ForegroundColor Cyan
Write-host "Done!" -ForegroundColor Cyan


#End
