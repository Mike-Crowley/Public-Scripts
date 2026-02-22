<#
.SYNOPSIS
    Finds duplicate files in an Azure File Share by comparing MD5 hashes.

.DESCRIPTION
    Find-AzFileShareDuplicates enumerates files in an Azure File Share, identifies
    duplicate groups, and generates reports (HTML, CSV, JSON) on the desktop.

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
    .\Find-AzFileShareDuplicates.ps1 -StorageAccountName "mystorageacct" -ShareName "myshare" `
        -RootPath "Projects/2025" -ResultSize 5000 -TenantId "xxx-xxx"

    # Scans specific folder with larger result set.

.NOTES
    Author: Mike Crowley
    Requires: Az.Accounts, Az.Storage, Az.Resources modules

.LINK
    https://learn.microsoft.com/en-us/powershell/module/az.storage/get-azstoragefile
#>

[CmdletBinding()]
param(
    [string]$StorageAccountName,
    [string]$ShareName,
    [string]$RootPath = "",
    [switch]$NoRecursion,
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
    $script:MoreFilesExist = $false

    if ($NoRecursion) {
        # Non-recursive: single directory scan
        $FileList = [System.Collections.Generic.List[PSCustomObject]]::new()

        $getFileParams = @{
            ShareName = $ShareName
            Context   = $ctx
        }
        if ($RootPath) { $getFileParams.Path = $RootPath }

        $items = Get-AzStorageFile @getFileParams
        if ($items -and ($items.GetType().Name -eq 'AzureStorageFileDirectory' -or $items.ShareDirectoryClient)) {
            $items = $items | Get-AzStorageFile
        }

        foreach ($item in $items) {
            if ($FileList.Count -ge $ResultSize) {
                $script:MoreFilesExist = $true
                break
            }
            if ($item.ShareFileClient) {
                $fullPath = if ($RootPath) { "$RootPath/$($item.Name)" } else { $item.Name }
                $FileList.Add([PSCustomObject]@{
                        Name                 = $item.Name
                        Path                 = $fullPath
                        LastModifiedDateTime = $item.FileProperties.LastModified.DateTime
                        MD5Hash              = if ($item.FileProperties.ContentHash) { [Convert]::ToHexString($item.FileProperties.ContentHash) } else { "" }
                        Size                 = $item.FileProperties.ContentLength
                        Url                  = $item.ShareFileClient.Uri.AbsoluteUri
                    })
            }
        }
        $FileList = $FileList.ToArray()
    }
    else {
        # Recursive scan with progress
        function Get-FilesRecursively {
            param(
                [string]$Path,
                [System.Collections.Generic.List[PSCustomObject]]$AllFiles
            )

            if ($AllFiles.Count -ge $ResultSize) {
                $script:MoreFilesExist = $true
                return
            }

            $getFileParams = @{ ShareName = $ShareName; Context = $ctx }
            if ($Path) { $getFileParams.Path = $Path }

            try {
                $items = Get-AzStorageFile @getFileParams
                if ($items -and ($items.GetType().Name -eq 'AzureStorageFileDirectory' -or
                        ($items.ShareDirectoryClient -and -not $items.ShareFileClient))) {
                    $items = $items | Get-AzStorageFile
                }

                foreach ($item in $items) {
                    if ($AllFiles.Count -ge $ResultSize) {
                        $script:MoreFilesExist = $true
                        return
                    }

                    if ($item.ShareDirectoryClient -and -not $item.ShareFileClient) {
                        # Recurse into subdirectory
                        $subPath = if ($Path) { "$Path/$($item.Name)" } else { $item.Name }
                        Get-FilesRecursively -Path $subPath -AllFiles $AllFiles
                    }
                    elseif ($item.ShareFileClient) {
                        $fullPath = if ($Path) { "$Path/$($item.Name)" } else { $item.Name }
                        $AllFiles.Add([PSCustomObject]@{
                                Name                 = $item.Name
                                Path                 = $fullPath
                                LastModifiedDateTime = $item.FileProperties.LastModified.DateTime
                                MD5Hash              = if ($item.FileProperties.ContentHash) { [Convert]::ToHexString($item.FileProperties.ContentHash) } else { "" }
                                Size                 = $item.FileProperties.ContentLength
                                Url                  = $item.ShareFileClient.Uri.AbsoluteUri
                            })

                        # Progress every 100 files
                        if (-not $Silent -and ($AllFiles.Count % 100 -eq 0)) {
                            Write-Host "`rScanning... $($AllFiles.Count) files found" -ForegroundColor Cyan -NoNewline
                        }
                    }
                }
            }
            catch {
                if (-not $Silent) { Write-Warning "Failed to access '$Path': $_" }
            }
        }

        $FileList = [System.Collections.Generic.List[PSCustomObject]]::new()
        if (-not $Silent) { Write-Host "Scanning... " -ForegroundColor Cyan -NoNewline }
        Get-FilesRecursively -Path $RootPath -AllFiles $FileList
        if (-not $Silent) { Write-Host "`rScanned $($FileList.Count) files.                    " -ForegroundColor Cyan }
        $FileList = $FileList.ToArray()
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

    # Report output
    $fileDate = Get-Date -Format 'yyyyMMdd_HHmmss'
    $desktop = [Environment]::GetFolderPath("Desktop")
    $folderName = "AzFileDupes_$($StorageAccountName)_$($ShareName)_$fileDate"
    $outputFolder = Join-Path $desktop $folderName

    # Create output folder
    $null = New-Item -ItemType Directory -Path $outputFolder -Force

    $csvPath = Join-Path $outputFolder "DuplicateReport.csv"
    $jsonPath = Join-Path $outputFolder "DuplicateReport.json"
    $htmlPath = Join-Path $outputFolder "DuplicateReport.html"

    # Calculate metrics for HTML report
    $totalBytesEvaluated = ($FileList.Size | Measure-Object -Sum).Sum
    $totalWasteBytes = ($Output.PossibleWasteKB | Measure-Object -Sum).Sum * 1KB

    # Top 10 waste by total bytes (size * count)
    $topWaste = $Output |
    Sort-Object { $_.FileSizeKB * $_.NumberOfFiles } -Descending |
    Select-Object -First 10

    # Generate HTML report
    $htmlContent = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Azure File Share Duplicate Report</title>
    <style>
        * { box-sizing: border-box; }
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            margin: 0; padding: 20px;
            background: #1a1a2e; color: #eee;
        }
        .container { max-width: 1200px; margin: 0 auto; }
        h1 { color: #00d4ff; margin-bottom: 5px; }
        .subtitle { color: #888; margin-bottom: 30px; }
        .metrics {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 15px; margin-bottom: 30px;
        }
        .metric {
            background: #16213e; padding: 20px; border-radius: 8px;
            border-left: 4px solid #00d4ff;
        }
        .metric-value { font-size: 28px; font-weight: bold; color: #00d4ff; }
        .metric-label { color: #888; font-size: 14px; margin-top: 5px; }
        .buttons { margin-bottom: 30px; }
        .btn {
            display: inline-block; padding: 12px 24px; margin-right: 10px;
            background: #00d4ff; color: #1a1a2e; text-decoration: none;
            border-radius: 6px; font-weight: bold; cursor: pointer; border: none;
        }
        .btn:hover { background: #00a8cc; }
        .btn-secondary { background: #4a4a6a; color: #eee; }
        .btn-secondary:hover { background: #5a5a7a; }
        table { width: 100%; border-collapse: collapse; margin-top: 20px; }
        th, td { padding: 12px; text-align: left; border-bottom: 1px solid #2a2a4a; }
        th { background: #16213e; color: #00d4ff; }
        tr:hover { background: #1f1f3a; }
        .confidence-high { color: #4ade80; }
        .confidence-low { color: #fbbf24; }
        .truncate { max-width: 300px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
        .section-title { color: #00d4ff; margin-top: 30px; margin-bottom: 15px; }
    </style>
</head>
<body>
    <div class="container">
        <h1>Azure File Share Duplicate Report</h1>
        <p class="subtitle">
            <strong>Storage:</strong> $StorageAccountName |
            <strong>Share:</strong> $ShareName |
            <strong>Path:</strong> $(if ($RootPath) { $RootPath } else { "/" }) |
            <strong>Generated:</strong> $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        </p>

        <div class="metrics">
            <div class="metric">
                <div class="metric-value">$($FileList.Count.ToString('N0'))</div>
                <div class="metric-label">Files Evaluated</div>
            </div>
            <div class="metric">
                <div class="metric-value">$([math]::Round($totalBytesEvaluated / 1GB, 2)) GB</div>
                <div class="metric-label">Total Size Evaluated</div>
            </div>
            <div class="metric">
                <div class="metric-value">$($Output.Count)</div>
                <div class="metric-label">Duplicate Groups</div>
            </div>
            <div class="metric">
                <div class="metric-value">$([math]::Round($totalWasteBytes / 1MB, 2)) MB</div>
                <div class="metric-label">Potential Waste</div>
            </div>
            <div class="metric">
                <div class="metric-value">$($filesWithHash.Count)</div>
                <div class="metric-label">Files with Hash</div>
            </div>
            <div class="metric">
                <div class="metric-value">$($filesWithoutHash.Count)</div>
                <div class="metric-label">Files without Hash</div>
            </div>
        </div>

        <div class="buttons">
            <a href="DuplicateReport.csv" class="btn">Open CSV Report</a>
            <a href="DuplicateReport.json" class="btn btn-secondary">Open JSON Report</a>
        </div>

        <h2 class="section-title">Top 10 by Wasted Space</h2>
        <table>
            <thead>
                <tr>
                    <th>Confidence</th>
                    <th>Match Type</th>
                    <th>Files</th>
                    <th>File Size</th>
                    <th>Total Waste</th>
                    <th>File Names</th>
                </tr>
            </thead>
            <tbody>
$(foreach ($item in $topWaste) {
    $confClass = if ($item.Confidence -eq 'High') { 'confidence-high' } else { 'confidence-low' }
    "                <tr>
                    <td class='$confClass'>$($item.Confidence)</td>
                    <td>$($item.MatchType)</td>
                    <td>$($item.NumberOfFiles)</td>
                    <td>$([math]::Round($item.FileSizeKB / 1024, 2)) MB</td>
                    <td>$([math]::Round($item.PossibleWasteKB / 1024, 2)) MB</td>
                    <td class='truncate' title='$($item.FileNames)'>$($item.FileNames)</td>
                </tr>"
})
            </tbody>
        </table>

        $(if ($script:MoreFilesExist) {
            "<p style='color: #fbbf24; margin-top: 30px;'><strong>⚠ Warning:</strong> ResultSize limit ($ResultSize) reached. More files exist but were not evaluated.</p>"
        })
    </div>
</body>
</html>
"@

    try {
        $Output | Export-Csv $csvPath -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
        $Output | ConvertTo-Json -Depth 3 | Out-File $jsonPath -Encoding UTF8 -ErrorAction Stop
        $htmlContent | Out-File $htmlPath -Encoding UTF8 -ErrorAction Stop

        # Open the HTML report
        Start-Process $htmlPath

        if (-not $Silent) {
            $EndTime = Get-Date
            Write-Host "`nComplete. " -ForegroundColor Green -NoNewline
            Write-Host "$($FileList.Count) files scanned in $([math]::Ceiling(($EndTime - $StartTime).TotalSeconds))s. " -NoNewline
            Write-Host "Found $($Output.Count) duplicate groups." -ForegroundColor Cyan
            Write-Host "Report: $outputFolder" -ForegroundColor DarkGray
            if ($script:MoreFilesExist) {
                Write-Host "Warning: ResultSize limit ($ResultSize) reached. Use -ResultSize to scan more." -ForegroundColor Yellow
            }
        }
    }
    catch {
        Write-Warning "Failed to save report files: $_"
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
