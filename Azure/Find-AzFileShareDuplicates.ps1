<#
.SYNOPSIS
    Finds duplicate files in an Azure File Share by comparing MD5 hashes.

.DESCRIPTION
    Find-AzFileShareDuplicates enumerates files in an Azure File Share, identifies
    duplicate groups, and reports results to the desktop, the pipeline, or both.

    Duplicates are identified using two methods:
    - High Confidence: Files with matching MD5 content hashes
    - Low Confidence: Files without hashes that share the same filename

    Can be run directly as a script or dot-sourced to load the function.

.PARAMETER StorageAccountName
    The name of the Azure storage account containing the file share.

.PARAMETER ShareName
    The name of the file share to scan.

.PARAMETER RootPath
    The starting directory path within the file share. Defaults to root.

.PARAMETER NoRecursion
    When specified, only scans the immediate directory without descending into subfolders.

.PARAMETER OutputStyle
    Controls output behavior:
        Report           - Saves CSV/JSON to desktop only
        PassThru         - Returns objects to pipeline only
        ReportAndPassThru - Both (default)

.PARAMETER ResultSize
    Maximum number of files to evaluate. Defaults to 1000. A warning is displayed
    if more files exist beyond this limit.

.PARAMETER Silent
    Suppresses console progress output.

.PARAMETER TenantId
    Azure AD tenant ID. If provided and no Azure context exists, triggers device authentication.

.EXAMPLE
    .\Find-AzFileShareDuplicates.ps1 -StorageAccountName "mystorageacct" -ShareName "myshare" -TenantId "xxx-xxx"

    # Direct invocation with device auth.

.EXAMPLE
    . .\Find-AzFileShareDuplicates.ps1
    Find-AzFileShareDuplicates -StorageAccountName "mystorageacct" -ShareName "myshare"

    # Dot-source to load function, then call (assumes existing Azure context).

.EXAMPLE
    .\Find-AzFileShareDuplicates.ps1 -StorageAccountName "mystorageacct" -ShareName "myshare" `
        -RootPath "Projects/2025" -OutputStyle Report -ResultSize 1000 -TenantId "xxx-xxx"

    # Scans specific folder, limits to 1000 files, generates desktop report only.

.NOTES
    Requires: Az.Accounts, Az.Storage, Az.Resources modules
    Author: Based on patterns from Mike Crowley's OneDrive duplicate finder

.LINK
    https://learn.microsoft.com/en-us/powershell/module/az.storage/get-azstoragefile
#>

[CmdletBinding()]
param(
    [string]$StorageAccountName,
    [string]$ShareName,
    [string]$RootPath = "",
    [switch]$NoRecursion,
    [ValidateSet("Report", "PassThru", "ReportAndPassThru")]
    [string]$OutputStyle = "ReportAndPassThru",
    [int32]$ResultSize = 1000,
    [switch]$Silent,
    [string]$TenantId
)

function Find-AzFileShareDuplicates {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$StorageAccountName,

        [Parameter(Mandatory)]
        [string]$ShareName,

        [string]$RootPath = "",

        [switch]$NoRecursion,

        [ValidateSet("Report", "PassThru", "ReportAndPassThru")]
        [string]$OutputStyle = "ReportAndPassThru",

        [int32]$ResultSize = 1000,

        [switch]$Silent,

        [string]$TenantId
    )

    $StartTime = Get-Date

    # Pre-flight checks
    $requiredModules = @('Az.Accounts', 'Az.Storage', 'Az.Resources')
    foreach ($module in $requiredModules) {
        if (-not (Get-Module -ListAvailable -Name $module)) {
            throw "Required module '$module' not found. Install via: Install-Module $module"
        }
    }

    # Authentication handling
    $azContext = Get-AzContext
    if (-not $azContext) {
        if ($TenantId) {
            if (-not $Silent) { Write-Host "No Azure context found. Initiating device authentication..." -ForegroundColor Yellow }
            $connectParams = @{
                UseDeviceAuthentication = $true
                TenantId                = $TenantId
                WarningAction           = 'SilentlyContinue'
            }
            $null = Connect-AzAccount @connectParams
        }
        else {
            throw "No Azure context found. Either run Connect-AzAccount first, or provide -TenantId to trigger device authentication."
        }
    }

    # Obtain storage context
    if (-not $Silent) { Write-Host "Connecting to storage account: " -ForegroundColor Cyan -NoNewline }
    if (-not $Silent) { Write-Host $StorageAccountName -ForegroundColor DarkCyan }

    $storageResource = Get-AzResource -Name $StorageAccountName -ResourceType "Microsoft.Storage/storageAccounts" -ErrorAction Stop
    if (-not $storageResource) {
        throw "Storage account '$StorageAccountName' not found."
    }

    $keyParams = @{
        ResourceGroupName = $storageResource.ResourceGroupName
        Name              = $StorageAccountName
    }
    $keys = Get-AzStorageAccountKey @keyParams

    $ctxParams = @{
        StorageAccountName = $StorageAccountName
        StorageAccountKey  = $keys[0].Value
    }
    $ctx = New-AzStorageContext @ctxParams

    if (-not $Silent) { Write-Host "Share: " -ForegroundColor Cyan -NoNewline }
    if (-not $Silent) { Write-Host $ShareName -ForegroundColor DarkCyan }
    if (-not $Silent) { Write-Host "" }

    # File collection
    $FileList = [Collections.Generic.List[PSCustomObject]]::new()
    $script:MoreFilesExist = $false  # Track if we hit ResultSize limit

    # Helper to write progress
    function Write-Progress {
        param([string]$Path)
        if (-not $Silent) {
            Write-Host "Searching: " -ForegroundColor Cyan -NoNewline
            Write-Host $(if ($Path) { $Path } else { "/" }) -ForegroundColor DarkCyan
            if ($FileList.Count -gt 0) {
                Write-Host "Files Found: " -ForegroundColor Cyan -NoNewline
                Write-Host $FileList.Count -ForegroundColor DarkCyan
            }
        }
    }

    # Helper to extract file properties into consistent object
    function ConvertTo-FileRecord {
        param($Item, [string]$ParentPath)

        $fullPath = if ($ParentPath) { "$ParentPath/$($Item.Name)" } else { $Item.Name }

        [PSCustomObject]@{
            Name                 = $Item.Name
            Path                 = $fullPath
            LastModifiedDateTime = $Item.FileProperties.LastModified.DateTime
            MD5Hash              = [Convert]::ToHexString($Item.FileProperties.ContentHash)
            Size                 = $Item.FileProperties.ContentLength
            Url                  = $Item.ShareFileClient.Uri.AbsoluteUri
        }
    }

    if ($NoRecursion) {
        # Non-recursive: single directory scan
        Write-Progress -Path $RootPath

        $getFileParams = @{
            ShareName = $ShareName
            Context   = $ctx
        }
        if ($RootPath) {
            $getFileParams.Path = $RootPath
        }

        $items = Get-AzStorageFile @getFileParams
        
        # If we got a directory back, enumerate its contents
        if ($items.GetType().Name -eq 'AzureStorageFileDirectory' -or $items.ShareDirectoryClient) {
            $items = $items | Get-AzStorageFile
        }

        foreach ($item in $items) {
            if ($FileList.Count -ge $ResultSize) {
                $script:MoreFilesExist = $true
                break
            }
            
            # Skip directories - we only want files
            if ($item.ShareFileClient -and $item.FileProperties.ContentHash) {
                $FileList.Add((ConvertTo-FileRecord -Item $item -ParentPath $RootPath))
            }
            elseif ($item.ShareFileClient) {
                # File without hash - still add it
                $FileList.Add((ConvertTo-FileRecord -Item $item -ParentPath $RootPath))
            }
        }
    }
    else {
        # Recursive scan using internal function
        function Get-FilesRecursively {
            param(
                [string]$Path,
                [Collections.Generic.List[PSCustomObject]]$AllFiles
            )

            if ($AllFiles.Count -ge $ResultSize) {
                $script:MoreFilesExist = $true
                return
            }

            Write-Progress -Path $Path

            $getFileParams = @{
                ShareName = $ShareName
                Context   = $ctx
            }
            if ($Path) {
                $getFileParams.Path = $Path
            }

            try {
                $items = Get-AzStorageFile @getFileParams
                
                # Handle case where we get a single directory object back
                if ($items -and ($items.GetType().Name -eq 'AzureStorageFileDirectory' -or 
                    ($items.ShareDirectoryClient -and -not $items.ShareFileClient))) {
                    $items = $items | Get-AzStorageFile
                }
            }
            catch {
                if (-not $Silent) {
                    Write-Warning "Failed to access path '$Path': $_"
                }
                return
            }

            foreach ($item in $items) {
                if ($AllFiles.Count -ge $ResultSize) {
                    $script:MoreFilesExist = $true
                    return
                }

                # Check if it's a directory (has ShareDirectoryClient but no ShareFileClient)
                $isDirectory = $item.ShareDirectoryClient -and -not $item.ShareFileClient

                if ($isDirectory) {
                    $subPath = if ($Path) { "$Path/$($item.Name)" } else { $item.Name }
                    Get-FilesRecursively -Path $subPath -AllFiles $AllFiles
                }
                elseif ($item.ShareFileClient) {
                    # Add all files, with or without hash
                    $AllFiles.Add((ConvertTo-FileRecord -Item $item -ParentPath $Path))
                }
            }
        }

        Get-FilesRecursively -Path $RootPath -AllFiles $FileList
    }

    # Separate files with and without hashes
    $filesWithHash = $FileList | Where-Object { $_.MD5Hash -and $_.MD5Hash -ne "" }
    $filesWithoutHash = $FileList | Where-Object { -not $_.MD5Hash -or $_.MD5Hash -eq "" }

    # Group by hash for high-confidence duplicates
    $hashBasedDupes = $filesWithHash | 
        Group-Object MD5Hash | 
        Where-Object Count -ge 2

    # Group by filename for files without hashes (lower confidence)
    $nameBasedDupes = $filesWithoutHash | 
        Group-Object Name | 
        Where-Object Count -ge 2

    # Build output objects
    $Output = @()

    # High confidence (hash-based) duplicates
    foreach ($Group in $hashBasedDupes) {
        $fileGroupSize = ($Group.Group.Size | Measure-Object -Sum).Sum
        $singleFileSize = $Group.Group[0].Size

        $Output += [PSCustomObject]@{
            Confidence      = "High"
            MatchType       = "MD5Hash"
            MatchValue      = $Group.Name
            NumberOfFiles   = $Group.Count
            FileSizeKB      = [math]::Round($singleFileSize / 1KB, 2)
            FileGroupSizeKB = [math]::Round($fileGroupSize / 1KB, 2)
            PossibleWasteKB = [math]::Round(($fileGroupSize - $singleFileSize) / 1KB, 2)
            FileNames       = ($Group.Group.Name | Sort-Object -Unique) -join ";"
            Paths           = ($Group.Group.Path | Sort-Object) -join ";"
            Urls            = ($Group.Group.Url) -join ";"
        }
    }

    # Low confidence (name-based) duplicates for files without hashes
    foreach ($Group in $nameBasedDupes) {
        $fileGroupSize = ($Group.Group.Size | Measure-Object -Sum).Sum
        $singleFileSize = $Group.Group[0].Size

        $Output += [PSCustomObject]@{
            Confidence      = "Low"
            MatchType       = "FileName"
            MatchValue      = $Group.Name
            NumberOfFiles   = $Group.Count
            FileSizeKB      = [math]::Round($singleFileSize / 1KB, 2)
            FileGroupSizeKB = [math]::Round($fileGroupSize / 1KB, 2)
            PossibleWasteKB = [math]::Round(($fileGroupSize - $singleFileSize) / 1KB, 2)
            FileNames       = $Group.Name
            Paths           = ($Group.Group.Path | Sort-Object) -join ";"
            Urls            = ($Group.Group.Url) -join ";"
        }
    }

    # Pipeline output
    if ($OutputStyle -in @("PassThru", "ReportAndPassThru")) {
        $Output
    }

    # Report output
    if ($OutputStyle -in @("Report", "ReportAndPassThru")) {
        $fileDate = Get-Date -Format 'ddMMMyyyy_HHmm.s'
        $desktop = [Environment]::GetFolderPath("Desktop")
        $baseName = "$StorageAccountName-$ShareName-DupeReport-$fileDate"
        $csvPath = "$desktop\$baseName.csv"
        $jsonPath = "$desktop\$baseName.json"

        try {
            $Output | Export-Csv $csvPath -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
            $Output | ConvertTo-Json -Depth 3 | Out-File $jsonPath -Encoding UTF8 -ErrorAction Stop
            
            if (-not $Silent) {
                Write-Host "`nReports saved:" -ForegroundColor Cyan
                Write-Host " CSV:  " -ForegroundColor Cyan -NoNewline
                Write-Host $csvPath -ForegroundColor DarkCyan
                Write-Host " JSON: " -ForegroundColor Cyan -NoNewline
                Write-Host $jsonPath -ForegroundColor DarkCyan
            }
        }
        catch {
            Write-Warning "Failed to save report files: $_"
        }
    }

    # Summary
    $EndTime = Get-Date
    $highConfCount = ($Output | Where-Object Confidence -eq "High").Count
    $lowConfCount = ($Output | Where-Object Confidence -eq "Low").Count

    if (-not $Silent) {
        Write-Host "`n--- Job Summary ---" -ForegroundColor Cyan
        Write-Host " Duration:        " -ForegroundColor Cyan -NoNewline
        Write-Host "$([math]::Ceiling(($EndTime - $StartTime).TotalSeconds)) seconds" -ForegroundColor DarkCyan
        Write-Host " Files Evaluated: " -ForegroundColor Cyan -NoNewline
        Write-Host $FileList.Count -ForegroundColor DarkCyan
        Write-Host " Files w/ Hash:   " -ForegroundColor Cyan -NoNewline
        Write-Host $filesWithHash.Count -ForegroundColor DarkCyan
        Write-Host " Files w/o Hash:  " -ForegroundColor Cyan -NoNewline
        Write-Host $filesWithoutHash.Count -ForegroundColor $(if ($filesWithoutHash.Count -gt 0) { "Yellow" } else { "DarkCyan" })
        Write-Host " Duplicate Groups:" -ForegroundColor Cyan -NoNewline
        Write-Host " $($Output.Count) " -ForegroundColor DarkCyan -NoNewline
        Write-Host "($highConfCount high / $lowConfCount low confidence)" -ForegroundColor DarkGray
        Write-Host " Potential Waste: " -ForegroundColor Cyan -NoNewline
        Write-Host "$([math]::Round(($Output.PossibleWasteKB | Measure-Object -Sum).Sum / 1024, 2)) MB" -ForegroundColor DarkCyan
        Write-Host " Recursion:       " -ForegroundColor Cyan -NoNewline
        Write-Host $(-not $NoRecursion) -ForegroundColor DarkCyan

        if ($script:MoreFilesExist) {
            Write-Host "`n WARNING: " -ForegroundColor Yellow -NoNewline
            Write-Host "ResultSize limit ($ResultSize) reached. More files exist but were not evaluated." -ForegroundColor Yellow
            Write-Host "          Use -ResultSize to increase the limit." -ForegroundColor DarkYellow
        }
    }
}

# Allow direct script invocation (not just dot-sourcing)
if ($MyInvocation.InvocationName -ne '.') {
    # Script was called directly - pass bound parameters to the function
    $scriptParams = @{}
    foreach ($key in $PSBoundParameters.Keys) {
        $scriptParams[$key] = $PSBoundParameters[$key]
    }
    if ($scriptParams.Count -gt 0) {
        Find-AzFileShareDuplicates @scriptParams
    }
}
