#Requires -Modules PnP.PowerShell, ImportExcel

<#
.SYNOPSIS
    Restores files from a SharePoint Online site's recycle bin after a given date, with logging.

.DESCRIPTION
    Enumerates the contents of a SPO site's first-stage recycle bin and restores files deleted
    after the specified date. Designed to handle a large number of items more reliably than the
    browser-based restore.

    Features:
        - Exports a list of files to restore before starting
        - Displays progress during the run
        - Writes a log file with success/failure status for each file

    The most common failure is when a file already exists at the original location.

    Requirements:
        - SharePoint Online (not SharePoint Server)
        - PnP.PowerShell and ImportExcel modules
        - PowerShell 5.0+
        - Permission to restore items from the site
        - Interactive session (credentials entered interactively)

.PARAMETER SiteUrl
    The full URL of the SharePoint Online site.

.PARAMETER RestoreDate
    Files deleted after this date will be restored.

.PARAMETER LogDirectory
    Directory path where the pre-restore inventory and restore log files are saved.

.EXAMPLE
    Restore-FromRecycleBin -SiteUrl https://MySpoSite.sharepoint.com/sites/Site123 -RestoreDate 3/23/2021 -LogDirectory C:\Logs

.NOTES
    Author: Mike Crowley, Jhon Ramirez
    https://mikecrowley.us

    Requires: PnP.PowerShell, ImportExcel modules

.LINK
    https://github.com/Mike-Crowley/Public-Scripts
#>


function Restore-FromRecycleBin {
    [cmdletbinding()]
    Param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$SiteUrl,

        [Parameter(Mandatory = $true)]
        [datetime]$RestoreDate,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$LogDirectory
    )

    $VerbosePreference = 'Continue'
    $ScriptDate = Get-Date -Format 'ddMMMyyyy_HHmm.s'
    $ExportParams =
    @{
        AutoSize     = $true
        AutoFilter   = $true
        BoldTopRow   = $true
        FreezeTopRow = $true
    }

    $SpoConnection = Connect-PnPOnline -Url $SiteUrl -Interactive -ReturnConnection

    mkdir $LogDirectory -Force


    $RecycleBinFiles = Get-PnPRecycleBinItem -Connection $SpoConnection -FirstStage | Where-Object { [datetime]$_.DeletedDateLocalFormatted -gt $RestoreDate } | Where-Object { $_.ItemType -eq "File" }
    Write-Verbose ("Found " + ($RecycleBinFiles.count) + " files.")
    $RecycleBinFiles | Export-Excel @ExportParams -Path ("$LogDirectory\RecycleBinFiles\" + "RecycleBinFiles--" + $ScriptDate + ".xlsx")

    $LogFile = @()
    $LoopCounter = 1

    foreach ($File in $RecycleBinFiles) {
        Write-Verbose ("Attempting to Restore: " + $File.Title + " to: " + $File.DirName + "`n")
        Write-Verbose ("$LoopCounter" + " of " + $RecycleBinFiles.Count + "`n")
        $LoopCounter ++

        $RestoreSucceeded = $true
        try { Restore-PnpRecycleBinItem -Force -Identity $File.Id.Guid -ErrorAction Stop }
        catch { $RestoreSucceeded = $false }

        $LogFile += '' | Select-Object @(
            @{N = "RestoreAttempt"; e = { Get-Date -UFormat "%D %r" } }
            @{N = "RestoreSucceeded"; e = { $RestoreSucceeded } }
            @{N = "FileName"; e = { $File.Title } }
            @{N = "DirName"; e = { $File.DirName } }
            @{N = "OriginalDeletionTime"; e = { $File.DeletedDateLocalFormatted } }
            @{N = "Id"; e = { $File.Id } }
        )
        switch ($RestoreSucceeded) {
            $true { Write-Verbose ("Restored: " + ($File.Title)) }
            $false { Write-Verbose ("ERROR: " + $Error[0].ErrorDetails) }
        }
    }

    $LogFile | Export-Excel @ExportParams -Path ("$LogDirectory\RecycleBinFiles\" + "RestoreLog--" + $ScriptDate + ".xlsx")
}
