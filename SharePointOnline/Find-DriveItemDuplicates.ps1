<#
.SYNOPSIS
    Identifies duplicate files across OneDrive and SharePoint Online document libraries using
    the Microsoft Graph API.

.DESCRIPTION
    Find-DriveItemDuplicates enumerates files in a OneDrive or SharePoint Online document library
    and groups them into duplicate sets. Two confidence levels are reported:

        High Confidence  -  Files whose quickXorHash content hashes match.
        Low  Confidence  -  Files that lack a hash but share an identical filename.

    Output is written to a timestamped folder on the desktop containing an interactive HTML
    dashboard, a CSV export, and a JSON export. Results may also be returned to the pipeline,
    controlled by -OutputStyle. The HTML report includes summary metrics, a top-10 wasted-space
    table, and links to the companion CSV and JSON files.

    The script can be dot-sourced to load the function, or invoked directly.

    Two operating modes are available:

        UPN Mode (default)  -  Scans a single user's OneDrive, identified by -Upn.
        Site Picker Mode    -  Presents an interactive Out-GridView workflow that lets an
                               administrator browse SharePoint sites (and optionally OneDrive
                               personal sites), select a target site, and choose a document
                               library when more than one exists.

    A Microsoft Graph session must be established before running this function:

        Connect-MgGraph -Scopes Files.Read                                # UPN mode
        Connect-MgGraph -Scopes Sites.Read.All                            # Site picker
        Connect-MgGraph -Scopes Sites.Read.All, User.Read.All             # Site picker + OneDrive

    To scan drives belonging to other users or application-only scenarios, register an
    Entra ID application with the appropriate permissions:
        https://learn.microsoft.com/en-us/powershell/microsoftgraph/app-only?view=graph-powershell-1.0

.PARAMETER Upn
    User principal name of the OneDrive owner to scan. When omitted, defaults to the
    Graph session account (Get-MgContext).Account, falling back to whoami /upn if
    unavailable. Mutually exclusive with -SitePicker.

.PARAMETER SitePicker
    Enables an interactive site-selection workflow powered by Out-GridView. The administrator
    is first prompted to include or exclude OneDrive personal sites, then presented with a
    filterable list of sites. If the selected SharePoint site contains multiple document
    libraries, a follow-up grid allows library selection. Mutually exclusive with -Upn.

.PARAMETER SiteCount
    Maximum number of sites to retrieve and display in the site picker grid. Accepts values
    from 1 to 1000. Default: 100. Ignored when -SitePicker is not specified.

.PARAMETER RootPath
    Starting directory path within the selected drive. Use forward-slash notation
    (e.g., "Documents/Projects"). Default: "/" (drive root).

.PARAMETER NoRecursion
    Limits the scan to the immediate contents of -RootPath. Subfolders are not traversed.

.PARAMETER OutputStyle
    Determines how results are delivered:
        Report            -  Generates desktop report files only.
        PassThru          -  Returns duplicate group objects to the pipeline only.
        ReportAndPassThru -  Generates reports and returns objects to the pipeline (default).

.PARAMETER ResultSize
    Maximum number of files to evaluate. Default: 32767 ([int16]::MaxValue). When the limit
    is reached, a warning is displayed in the console and embedded in the HTML report.

.PARAMETER Silent
    Suppresses all console progress and summary output.

.EXAMPLE
    Find-DriveItemDuplicates

    Scans the current user's entire OneDrive with default settings (recursive, up to 32 767
    files, reports and pipeline output, console progress enabled).

.EXAMPLE
    Find-DriveItemDuplicates -RootPath "Desktop" -OutputStyle Report -ResultSize 500

    Evaluates up to 500 files in the current user's Desktop folder and subfolders.
    Generates the HTML/CSV/JSON reports without returning objects to the pipeline.

.EXAMPLE
    Find-DriveItemDuplicates -Upn user1@example.com -RootPath "Desktop/DupesDirectory" -OutputStyle Report -NoRecursion

    Scans user1's Desktop\DupesDirectory without descending into subfolders and writes
    report files to the desktop.

.EXAMPLE
    Find-DriveItemDuplicates -SitePicker

    Opens the interactive site picker. An Out-GridView prompt asks whether to include
    OneDrive personal sites, followed by a site-selection grid and (if applicable) a
    document-library selection grid.

.EXAMPLE
    Find-DriveItemDuplicates -SitePicker -SiteCount 250 -ResultSize 5000

    Retrieves up to 250 sites for the picker grid and evaluates up to 5000 files on the
    selected site.

.NOTES
    Author: Mike Crowley
    https://mikecrowley.us

    Prerequisites
        Module:       Microsoft.Graph.Authentication
        Permissions:  Files.Read or Sites.Read (UPN mode)
                      Sites.Read.All (Site Picker mode)
                      Sites.Read.All, User.Read.All (Site Picker with OneDrive option)

    Testing
        The following snippet creates a set of duplicate and unique files suitable for
        validating this function. If Known Folder Move (KFM) is enabled, allow the
        OneDrive sync client to finish uploading before running Find-DriveItemDuplicates.

        # Setup
        $Desktop = [Environment]::GetFolderPath("Desktop")
        $TestDir = mkdir $Desktop\DupesDirectory -Force
        $TestLogFile = (Invoke-WebRequest "https://gist.githubusercontent.com/Mike-Crowley/d4275d6abd78ad8d19a6f1bcf9671ec4/raw/66fe537cfe8e58b1a5eb1c1336c4fdf6a9f05145/log.log.log").content

        # Duplicate files (identical content, different names)
        1..25 | ForEach-Object { $TestLogFile    | Out-File "$TestDir\$(Get-Random).log" }
        1..25 | ForEach-Object { "Hello World 1" | Out-File "$TestDir\$(Get-Random).log" }
        1..25 | ForEach-Object { "Hello World 2" | Out-File "$TestDir\$(Get-Random).log" }

        # Unique files (random content)
        1..25 | ForEach-Object { Get-Random | Out-File "$TestDir\$(Get-Random).log" }

.LINK
    https://mikecrowley.us/2024/04/20/onedrive-and-sharepoint-online-file-deduplication-report-microsoft-graph-api
#>

[CmdletBinding(DefaultParameterSetName = 'UPN')]
param(
    [Parameter(ParameterSetName = 'UPN')]
    [string]$Upn,

    [Parameter(Mandatory, ParameterSetName = 'SitePicker')]
    [switch]$SitePicker,

    [Parameter(ParameterSetName = 'SitePicker')]
    [ValidateRange(1, 1000)]
    [int32]$SiteCount = 100,

    [Parameter(ParameterSetName = 'SitePicker')]
    [switch]$IncludeStorageMetrics,

    [string]$RootPath = "/",
    [switch]$NoRecursion,

    [ValidateSet("Report", "PassThru", "ReportAndPassThru")]
    [string]$OutputStyle = "ReportAndPassThru",

    [int32]$ResultSize = [int16]::MaxValue,
    [switch]$Silent
)

function Find-DriveItemDuplicates {
    [CmdletBinding(DefaultParameterSetName = 'UPN')]
    param(
        [Parameter(ParameterSetName = 'UPN')]
        [string]$Upn,

        [Parameter(Mandatory, ParameterSetName = 'SitePicker')]
        [switch]$SitePicker,

        [Parameter(ParameterSetName = 'SitePicker')]
        [ValidateRange(1, 1000)]
        [int32]$SiteCount = 100,

        [Parameter(ParameterSetName = 'SitePicker')]
        [switch]$IncludeStorageMetrics,

        [string]$RootPath = "/",
        [switch]$NoRecursion,

        [ValidateSet("Report", "PassThru", "ReportAndPassThru")]
        [string]$OutputStyle = "ReportAndPassThru",

        [int32]$ResultSize = [int16]::MaxValue,
        [switch]$Silent
    )

    $StartTime = Get-Date
    $script:MoreFilesExist = $false

    #region Pre-flight checks
    if ($null -eq (Get-Command Invoke-MgGraphRequest -ErrorAction SilentlyContinue)) {
        throw "Invoke-MgGraphRequest cmdlet not found. Install the Microsoft.Graph.Authentication PowerShell module.`nhttps://learn.microsoft.com/en-us/graph/sdks/sdk-installation#install-the-microsoft-graph-powershell-sdk"
    }
    if ($null -eq (Get-MgContext)) {
        throw "No Graph context found. Please call Connect-MgGraph."
    }

    $scopes = (Get-MgContext).Scopes | Out-String
    if ($PSCmdlet.ParameterSetName -eq 'SitePicker') {
        if ($scopes -notlike '*Sites.Read.All*') {
            Write-Warning "Sites.Read.All scope may be required for site picker.`nhttps://learn.microsoft.com/en-us/graph/api/site-search"
        }
    }
    else {
        if ($scopes -notlike '*Files.Read*' -and $scopes -notlike '*Sites.Read*') {
            Write-Warning "Permission scope may be missing.`nhttps://learn.microsoft.com/en-us/graph/api/driveitem-list-children?view=graph-rest-beta&tabs=http#permissions"
        }
    }
    #endregion

    #region Drive resolution
    if ($PSCmdlet.ParameterSetName -eq 'SitePicker') {
        # Step A: Ask about OneDrive personal sites
        $SiteTypeOptions = @(
            [PSCustomObject]@{ Choice = "SharePoint sites only" }
            [PSCustomObject]@{ Choice = "SharePoint + OneDrive personal sites" }
        )
        $SiteTypeChoice = $SiteTypeOptions | Out-GridView -Title "Include OneDrive personal sites?" -OutputMode Single

        if (-not $SiteTypeChoice) {
            throw "No selection made. Exiting."
        }

        $includeOneDrive = $SiteTypeChoice.Choice -eq "SharePoint + OneDrive personal sites"

        # Step B: Fetch sites via Graph
        $SiteList = [Collections.Generic.List[Object]]::new()

        # Derive the tenant prefix from the Graph session account (e.g., "ussilica" from "user@ussilica.com")
        # SharePoint site URLs all contain {tenant}.sharepoint.com, so this matches broadly.
        $searchTerm = ((Get-MgContext).Account -split '@')[-1] -split '\.' | Select-Object -First 1
        if (-not $Silent) { Write-Host "Fetching SharePoint sites (search: $searchTerm)..." -ForegroundColor Cyan }
        $siteUri = "v1.0/sites?search=$searchTerm&`$top=$SiteCount&`$select=id,displayName,webUrl"
        do {
            $siteResponse = Invoke-MgGraphRequest -Uri $siteUri
            foreach ($site in $siteResponse.value) {
                if ($SiteList.Count -ge $SiteCount) { break }
                $SiteList.Add([PSCustomObject]@{
                    Type        = "SharePoint"
                    DisplayName = $site.displayName
                    WebUrl      = $site.webUrl
                    SiteId      = $site.id
                    Upn         = ""
                })
            }
            $siteUri = $siteResponse.'@odata.nextLink'
        } until (-not $siteUri -or $SiteList.Count -ge $SiteCount)

        if ($includeOneDrive) {
            if (-not $Silent) { Write-Host "Fetching OneDrive users..." -ForegroundColor Cyan }
            $userUri = "v1.0/users?`$top=100&`$select=id,displayName,userPrincipalName"
            do {
                $userResponse = Invoke-MgGraphRequest -Uri $userUri
                foreach ($user in $userResponse.value) {
                    if ($SiteList.Count -ge $SiteCount) { break }
                    $SiteList.Add([PSCustomObject]@{
                        Type        = "OneDrive"
                        DisplayName = "$($user.displayName) (OneDrive)"
                        WebUrl      = $user.userPrincipalName
                        SiteId      = ""
                        Upn         = $user.userPrincipalName
                    })
                }
                $userUri = $userResponse.'@odata.nextLink'
            } until (-not $userUri -or $SiteList.Count -ge $SiteCount)
        }

        if ($SiteList.Count -eq 0) {
            throw "No sites found. Verify your permissions (Sites.Read.All)."
        }

        if (-not $Silent) { Write-Host "Found $($SiteList.Count) sites." -ForegroundColor Cyan }

        # Step C: Present site picker
        $SelectedSite = $SiteList |
            Select-Object Type, DisplayName, WebUrl |
            Out-GridView -Title "Select a site to scan for duplicates ($($SiteList.Count) sites)" -OutputMode Single

        if (-not $SelectedSite) {
            throw "No site selected. Exiting."
        }

        # Map back to full object
        $SelectedSite = $SiteList | Where-Object {
            $_.WebUrl -eq $SelectedSite.WebUrl -and $_.Type -eq $SelectedSite.Type
        } | Select-Object -First 1

        if (-not $Silent) { Write-Host "Selected: $($SelectedSite.DisplayName)" -ForegroundColor Green }

        # Step D: Resolve the selected site's drive
        if ($SelectedSite.Type -eq "OneDrive") {
            $DriveResponse = Invoke-MgGraphRequest -Uri "beta/users/$($SelectedSite.Upn)/drive"
            $DriveId = $DriveResponse.id
            $DriveWebUrl = $DriveResponse.webUrl
            $reportIdentifier = $SelectedSite.Upn
        }
        else {
            $drivesResponse = Invoke-MgGraphRequest -Uri "beta/sites/$($SelectedSite.SiteId)/drives"
            $drives = $drivesResponse.value

            if (-not $drives -or $drives.Count -eq 0) {
                throw "No document libraries found for site '$($SelectedSite.DisplayName)'."
            }
            elseif ($drives.Count -eq 1) {
                $DriveId = $drives[0].id
                $DriveWebUrl = $drives[0].webUrl
            }
            else {
                $driveChoices = foreach ($d in $drives) {
                    [PSCustomObject]@{
                        Name    = $d.name
                        WebUrl  = $d.webUrl
                        DriveId = $d.id
                    }
                }
                $SelectedDrive = $driveChoices |
                    Select-Object Name, WebUrl |
                    Out-GridView -Title "Select a document library ($($drives.Count) libraries)" -OutputMode Single

                if (-not $SelectedDrive) {
                    throw "No document library selected. Exiting."
                }

                $matchedDrive = $driveChoices | Where-Object { $_.WebUrl -eq $SelectedDrive.WebUrl } | Select-Object -First 1
                $DriveId = $matchedDrive.DriveId
                $DriveWebUrl = $matchedDrive.WebUrl
            }
            $reportIdentifier = $SelectedSite.DisplayName -replace '[\\/:*?"<>|]', '_'
        }
    }
    else {
        # UPN mode (default) — prefer the Graph session account over the local Windows identity
        if (-not $Upn) {
            $mgAccount = (Get-MgContext).Account
            if ($mgAccount) {
                $Upn = $mgAccount
            }
            else {
                $Upn = whoami /upn
            }
        }
        if ($Upn -notmatch "^\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*$") {
            throw "Invalid UPN: $Upn"
        }

        try {
            $DriveResponse = Invoke-MgGraphRequest -Uri "beta/users/$Upn/drive" -ErrorAction Stop
        }
        catch {
            $msg = "$($_.Exception.Message) $($_.ErrorDetails.Message)"

            if ($msg -match '403|Forbidden|notAllowed|provisioningNotAllowed|not have a valid license') {
                throw "OneDrive is not provisioned or licensed for '$Upn'. Use -Upn to specify a licensed user, or use -SitePicker to browse SharePoint sites interactively."
            }
            elseif ($msg -match '404|Not Found|ResourceNotFound|User not found') {
                throw "User '$Upn' was not found in the connected tenant ($((Get-MgContext).TenantId)). Verify the UPN or use -Upn to specify a different user."
            }
            else {
                throw "Failed to retrieve drive for '$Upn': $($_.Exception.Message)"
            }
        }

        if ($DriveResponse.webUrl -notlike 'https*') {
            throw "Drive not found for '$Upn'. Use -Upn to specify a licensed user, or use -SitePicker to browse SharePoint sites interactively."
        }

        $DriveId = $DriveResponse.id
        $DriveWebUrl = $DriveResponse.webUrl
        $reportIdentifier = $Upn
    }

    if (-not $Silent) { Write-Host "`nFound Drive: " -ForegroundColor Cyan -NoNewline }
    if (-not $Silent) { Write-Host $DriveWebUrl -ForegroundColor DarkCyan }
    if (-not $Silent) { Write-Host "" }

    $baseUri = "beta/drives/$DriveId/root"
    #endregion

    #region File collection
    $FileList = [Collections.Generic.List[Object]]::new()

    if ($NoRecursion) {
        $uri = "beta/drives/$DriveId/root:/$($RootPath):/children"
        $RawItems = [Collections.Generic.List[Object]]::new()

        do {
            if (-not $Silent) { Write-Host "Searching: " -ForegroundColor Cyan -NoNewline }
            if (-not $Silent) { Write-Host $uri -ForegroundColor DarkCyan }
            $PageResults = Invoke-MgGraphRequest -Uri $uri
            if ($PageResults.value) {
                $RawItems.AddRange($PageResults.value)
            }
            else {
                $RawItems.Add($PageResults)
            }
            $uri = $PageResults.'@odata.nextlink'
        } until (-not $uri -or $RawItems.Count -ge $ResultSize)

        if ($uri -and $RawItems.Count -ge $ResultSize) {
            $script:MoreFilesExist = $true
        }

        $FileItems = $RawItems | Where-Object { $null -ne $_.file }
        $limit = [math]::Min($ResultSize, $FileItems.Count)

        foreach ($DriveItem in $FileItems[0..($limit - 1)]) {
            $FileList.Add([PSCustomObject]@{
                Name                 = $DriveItem.name
                LastModifiedDateTime = $DriveItem.lastModifiedDateTime
                QuickXorHash         = $DriveItem.file.hashes.quickXorHash
                Size                 = $DriveItem.size
                WebUrl               = $DriveItem.webUrl
            })
        }
    }
    else {
        # Internal function to recursively get files
        function Get-FolderItemsRecursively {
            [CmdletBinding()]
            param(
                [string]$Path,
                [Collections.Generic.List[Object]]$AllFiles
            )

            if ($AllFiles.Count -ge $ResultSize) {
                $script:MoreFilesExist = $true
                return
            }

            if ($Path -eq "/") {
                $uri = "$baseUri/children"
            }
            else {
                $uri = "$baseUri`:/$Path`:/children"
            }

            do {
                if (-not $Silent) { Write-Host "Searching: " -ForegroundColor Cyan -NoNewline }
                if (-not $Silent) { Write-Host $Path -ForegroundColor DarkCyan }

                if ($AllFiles.Count -gt 1) {
                    if (-not $Silent) { Write-Host "Files Evaluated: " -ForegroundColor Cyan -NoNewline }
                    if (-not $Silent) { Write-Host $AllFiles.Count -ForegroundColor DarkCyan }
                }

                $Response = Invoke-MgGraphRequest -Uri $Uri

                foreach ($Item in $Response.value) {
                    if ($null -ne $Item.folder) {
                        $FolderPath = "$Path/$($Item.name)"
                        Get-FolderItemsRecursively -Path $FolderPath -AllFiles $AllFiles
                    }
                    elseif ($null -ne $Item.file) {
                        $AllFiles.Add(
                            [PSCustomObject]@{
                                Name                 = $Item.name
                                LastModifiedDateTime = $Item.lastModifiedDateTime
                                QuickXorHash         = $Item.file.hashes.quickXorHash
                                Size                 = $Item.size
                                WebUrl               = $Item.webUrl
                            }
                        )
                    }
                }
                $Uri = $Response.'@odata.nextLink'
            } until ((-not $Uri) -or ($AllFiles.Count -ge $ResultSize))

            if ($AllFiles.Count -ge $ResultSize) {
                $script:MoreFilesExist = $true
            }
        }

        Get-FolderItemsRecursively -Path $RootPath -AllFiles $FileList
    }
    #endregion

    #region Duplicate detection (dual confidence)
    $filesWithHash = @($FileList | Where-Object { $_.QuickXorHash })
    $filesWithoutHash = @($FileList | Where-Object { -not $_.QuickXorHash })

    # High confidence: hash-based duplicates
    $hashBasedDupes = $filesWithHash |
        Group-Object QuickXorHash |
        Where-Object Count -ge 2

    # Low confidence: filename-based duplicates (for files without hashes)
    $nameBasedDupes = $filesWithoutHash |
        Group-Object Name |
        Where-Object Count -ge 2

    $Output = @()

    foreach ($Group in $hashBasedDupes) {
        $fileGroupSize = ($Group.Group.Size | Measure-Object -Sum).Sum
        $singleFileSize = $Group.Group[0].Size
        $Output += [PSCustomObject]@{
            Confidence      = "High"
            MatchType       = "QuickXorHash"
            MatchValue      = $Group.Name
            NumberOfFiles   = $Group.Count
            FileSizeKB      = [math]::Round($singleFileSize / 1KB, 2)
            FileGroupSizeKB = [math]::Round($fileGroupSize / 1KB, 2)
            PossibleWasteKB = [math]::Round(($fileGroupSize - $singleFileSize) / 1KB, 2)
            FileNames       = ($Group.Group.Name | Sort-Object -Unique) -join ";"
            WebLocalPaths   = ($Group.Group.WebUrl | ForEach-Object { ([uri]$_).LocalPath }) -join ";"
        }
    }

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
            WebLocalPaths   = ($Group.Group.WebUrl | ForEach-Object { ([uri]$_).LocalPath }) -join ";"
        }
    }
    #endregion

    #region Pipeline output
    if ($OutputStyle -eq "PassThru" -or $OutputStyle -eq "ReportAndPassThru") {
        $Output
    }
    #endregion

    #region Report generation (HTML, CSV, JSON)
    if ($OutputStyle -eq "Report" -or $OutputStyle -eq "ReportAndPassThru") {
        $fileDate = Get-Date -Format 'yyyyMMdd_HHmmss'
        $desktop = [Environment]::GetFolderPath("Desktop")
        $folderName = "SPODupes_$($reportIdentifier -replace '[\\/:*?"<>|@]', '_')_$fileDate"
        $outputFolder = Join-Path $desktop $folderName
        $null = New-Item -ItemType Directory -Path $outputFolder -Force

        $csvPath = Join-Path $outputFolder "DuplicateReport.csv"
        $jsonPath = Join-Path $outputFolder "DuplicateReport.json"
        $htmlPath = Join-Path $outputFolder "DuplicateReport.html"

        # Calculate metrics
        $totalBytesEvaluated = ($FileList.Size | Measure-Object -Sum).Sum
        $totalWasteBytes = ($Output.PossibleWasteKB | Measure-Object -Sum).Sum * 1KB

        $topWaste = $Output |
            Sort-Object { $_.FileSizeKB * $_.NumberOfFiles } -Descending |
            Select-Object -First 10

        $htmlContent = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>SharePoint / OneDrive Duplicate Report</title>
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
        <h1>SharePoint / OneDrive Duplicate Report</h1>
        <p class="subtitle">
            <strong>Source:</strong> $reportIdentifier |
            <strong>Drive:</strong> $DriveWebUrl |
            <strong>Path:</strong> $(if ($RootPath -and $RootPath -ne '/') { $RootPath } else { '/' }) |
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
            "<p style='color: #fbbf24; margin-top: 30px;'><strong>Warning:</strong> ResultSize limit ($ResultSize) reached. More files exist but were not evaluated. Use -ResultSize to increase.</p>"
        })
    </div>
</body>
</html>
"@

        try {
            $Output | Export-Csv $csvPath -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
            $Output | ConvertTo-Json -Depth 3 | Out-File $jsonPath -Encoding UTF8 -ErrorAction Stop
            $htmlContent | Out-File $htmlPath -Encoding UTF8 -ErrorAction Stop

            Start-Process $htmlPath
        }
        catch {
            Write-Warning "Failed to save report files: $_"
        }
    }
    #endregion

    #region Console summary
    if (-not $Silent) {
        $EndTime = Get-Date
        $duration = [math]::Ceiling(($EndTime - $StartTime).TotalSeconds)
        Write-Host "`nComplete. " -ForegroundColor Green -NoNewline
        Write-Host "$($FileList.Count) files scanned in ${duration}s. " -NoNewline
        Write-Host "Found $($Output.Count) duplicate groups." -ForegroundColor Cyan
        if ($OutputStyle -eq "Report" -or $OutputStyle -eq "ReportAndPassThru") {
            Write-Host "Report: $outputFolder" -ForegroundColor DarkGray
        }
        if ($script:MoreFilesExist) {
            Write-Host "Warning: ResultSize limit ($ResultSize) reached. Use -ResultSize to scan more." -ForegroundColor Yellow
        }
    }
    #endregion
}

# Allow direct script invocation (not just dot-sourcing)
if ($MyInvocation.InvocationName -ne '.') {
    $scriptParams = @{}
    foreach ($key in $PSBoundParameters.Keys) {
        $scriptParams[$key] = $PSBoundParameters[$key]
    }
    Find-DriveItemDuplicates @scriptParams
}
