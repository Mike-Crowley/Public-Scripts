#Requires -Modules PnP.PowerShell, ImportExcel

#DRAFT VERSION!!

<#
.SYNOPSIS 

    Restore-FromRecycleBin enumerates the contents of a SPO site's recycle bin and restores items after a given date, logging the results.
    This script was designed to handle a large number of items and aims to be more reliable than the native browser-based restore functionality.

.DESCRIPTION

    Restore-FromRecycleBin enumerates the contents of a SPO site's recycle bin and restores items after a given date, logging the results.
    This script was designed to handle a large number of items and aims to be more reliable than the native browser-based restore functionality.

    Features:

    - Before it starts, the script writes a file to disk containing all of the files it is about to restore. 
    - During the run, there is a counter and other output describing the progress.
    - When complete, a log file is written, which contains every file attempted and a true/false for its success.
    
    The most common failure we've seen is due to the fact a file already exists, which can be confirmed with use of the log file.

Requirements:
    1) SharePoint Online. This was not written for use with SharePoint Server.
    2) PnP.PowerShell and ImportExcel modules
    3) PowerShell 5.0+
    4) Permission to restore items from the site
    5) Interactive session. This is written to require an administrator to be present and enter the credentials interactivly.

 
March 22 2021
-Mike Crowley
-Jhon Ramirez 

.EXAMPLE

    Restore-FromRecycleBin -SiteUrl https://MySpoSite.sharepoint.com/sites/Site123 -RestoreDate 3/23/2021 -LogDirectory C:\Logs

.LINK 

    http://BaselineTechnologies.com
 
#>
 

function Restore-FromRecycleBin
{
    [cmdletbinding()]
    Param
    (
        [Parameter(Mandatory=$true)] [string] $SiteUrl,
        [Parameter(Mandatory=$true)] [datetime] $RestoreDate,
        [Parameter(Mandatory=$true)] [string] $LogDirectory
    )

    $VerbosePreference = 'Continue'
    $ScriptDate        = Get-Date -format ddMMMyyyy_HHmm.s
    $ExportParams =
    @{
        AutoSize     = $true
        AutoFilter   = $true
        BoldTopRow   = $true
        FreezeTopRow = $true
    }

    $SpoConnection = Connect-PnPOnline -Url $SiteUrl -Interactive -ReturnConnection

    md $LogDirectory -Force

    
    $RecycleBinFiles = Get-PnPRecycleBinItem -Connection $SpoConnection -FirstStage | Where {[datetime]$_.DeletedDateLocalFormatted -gt $RestoreDate} | Where {$_.ItemType -eq "File"}
    Write-Verbose ("Found " + ($RecycleBinFiles.count) + " files.")
    $RecycleBinFiles | Export-Excel @ExportParams -Path ("$LogDirectory\RecycleBinFiles\" + "RecycleBinFiles--" + $ScriptDate + ".xlsx")
        
    $LogFile = @()
    $LoopCounter = 1
 
    foreach ($File in $RecycleBinFiles)
    {
        Write-Verbose ("Attempting to Restore: " + $File.Title + " to: " + $File.DirName + "`n")
        Write-Verbose ("$LoopCounter" + " of " + $RecycleBinFiles.Count + "`n")
        $LoopCounter ++     
 
        $RestoreSucceeded = $true
        try {Restore-PnpRecycleBinItem -Force -Identity $File.Id.Guid -ErrorAction Stop}
        catch {$RestoreSucceeded = $false }       
 
        $LogFile += '' | select @(
            @{N="RestoreAttempt"; e={Get-Date -UFormat "%D %r"}}        
            @{N="RestoreSucceeded"; e={$RestoreSucceeded}}
            @{N="FileName"; e={$File.Title}}
            @{N="DirName"; e={$File.DirName}}
            @{N="OriginalDeletionTime"; e={$File.DeletedDateLocalFormatted}}
            @{N="Id"; e={$File.Id}}        
        ) 
        switch ($RestoreSucceeded)
        {
            $true {Write-Verbose ("Restored: " + ($File.Title))}
            $false {Write-Verbose ("ERROR: " + $Error[0].ErrorDetails)}
        }   
    }

    $LogFile | Export-Excel @ExportParams -Path ("$LogDirectory\RecycleBinFiles\" + "RestoreLog--" + $ScriptDate + ".xlsx")
}



 
#
