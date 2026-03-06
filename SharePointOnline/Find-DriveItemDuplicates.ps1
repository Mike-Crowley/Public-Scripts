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

    Three operating modes are available:

        UPN Mode (default)  -  Scans a single user's OneDrive, identified by -Upn.
        Site Picker Mode    -  Presents an interactive Out-GridView workflow that lets an
                               administrator browse SharePoint sites (and optionally OneDrive
                               personal sites), select a target site, and choose a document
                               library when more than one exists.
        App Auth Mode       -  Connects with application (client credential) permissions for
                               full read access to all sites and drives. Use -RegisterApp once
                               to create the required Entra ID app, then supply -ClientId and
                               -TenantId for subsequent runs. Defaults to site picker behavior;
                               add -Upn to scan a specific user's OneDrive instead.

    Delegated auth (UPN and Site Picker modes) requires a pre-existing Graph session:

        Connect-MgGraph -Scopes Files.Read                                # UPN mode
        Connect-MgGraph -Scopes Sites.Read.All                            # Site picker
        Connect-MgGraph -Scopes Sites.Read.All, User.Read.All             # Site picker + OneDrive
        Connect-MgGraph -Scopes Sites.Read.All, Reports.Read.All          # Site picker + storage metrics

    App auth handles its own connection — no prior Connect-MgGraph is needed.

.PARAMETER RegisterApp
    Creates an Entra ID app registration with Sites.Read.All and Reports.Read.All
    application permissions, a client secret (90-day expiry), a service principal,
    and grants admin consent. Run this once per tenant before using app auth.
    Requires Application.ReadWrite.All and AppRoleAssignment.ReadWrite.All
    permissions (typically Application Administrator or Global Administrator).
    Returns an object with ClientId and ClientSecret for subsequent runs.

.PARAMETER ClientId
    The Application (client) ID of the Entra ID app registration created by
    -RegisterApp. Required for app auth. When used without -Upn, enters site picker
    mode. When used with -Upn, scans the specified user's OneDrive.

.PARAMETER ClientSecret
    The client secret for the app registration. If omitted, you will be prompted
    via Get-Credential. Accepts the plain-text secret returned by -RegisterApp,
    e.g. -ClientSecret $app.ClientSecret.

.PARAMETER TenantId
    The tenant ID or domain name for the Microsoft 365 tenant. Required for
    -RegisterApp and app auth.

.PARAMETER Upn
    User principal name of the OneDrive owner to scan. When omitted in UPN mode,
    defaults to the Graph session account (Get-MgContext).Account, falling back to
    whoami /upn. When used with -ClientId, scans that user's OneDrive with app
    permissions. Mutually exclusive with -SitePicker.

.PARAMETER SitePicker
    Enables an interactive site-selection workflow powered by Out-GridView. The administrator
    is first prompted to include or exclude OneDrive personal sites, then presented with a
    filterable list of sites. If the selected SharePoint site contains multiple document
    libraries, a follow-up grid allows library selection. Mutually exclusive with -Upn.

.PARAMETER SiteCount
    Maximum number of sites to retrieve for the site picker grid. When -IncludeStorageMetrics
    is used, this limits the number of site collections shown (subsites are excluded
    automatically). Accepts values from 1 to 5000. Default: 500. Ignored when -SitePicker
    is not specified.

.PARAMETER IncludeStorageMetrics
    When used with -SitePicker, retrieves site-collection-level storage usage from the
    Graph Reports API (getSharePointSiteUsageDetail). Adds SiteCollGB and SiteCollFiles
    columns to the site picker grid. These values represent the entire site collection —
    subsites share their parent collection's totals. Requires Reports.Read.All scope.
    Data may be up to 48 hours old. Ignored without -SitePicker.

.PARAMETER RootPath
    Starting directory path within the selected drive. Use forward-slash notation
    (e.g., "Documents/Projects"). Default: "/" (drive root).

.PARAMETER NoRecursion
    Limits the scan to the immediate contents of -RootPath. Subfolders are not traversed.

.PARAMETER OutputStyle
    Determines how results are delivered:
        Report            -  Generates desktop report files only (default).
        PassThru          -  Returns duplicate group objects to the pipeline only.
        ReportAndPassThru -  Generates reports and returns objects to the pipeline.

.PARAMETER ResultSize
    Maximum number of files to evaluate. Default: 32767 ([int16]::MaxValue). When the limit
    is reached, a warning is displayed in the console and embedded in the HTML report.

.PARAMETER Silent
    Suppresses all console progress and summary output.

.EXAMPLE
    Find-DriveItemDuplicates

    Scans the current user's entire OneDrive with default settings (recursive, up to 32 767
    files, report output only, console progress enabled).

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
    Find-DriveItemDuplicates -SitePicker -IncludeStorageMetrics

    Site picker with site-collection-level storage columns (SiteCollGB, SiteCollFiles) from
    the Graph Reports API. Requires Reports.Read.All scope. Data may be up to 48 hours old.

.EXAMPLE
    Find-DriveItemDuplicates -SitePicker -SiteCount 1000 -ResultSize 5000

    Retrieves up to 1000 sites for the picker grid and evaluates up to 5000 files on the
    selected site.

.EXAMPLE
    $app = Find-DriveItemDuplicates -RegisterApp -TenantId "contoso.com"

    Creates an Entra ID app registration with Sites.Read.All and Reports.Read.All
    permissions. Save the returned ClientId and ClientSecret for subsequent runs.
    The secret cannot be retrieved later.

.EXAMPLE
    Find-DriveItemDuplicates -ClientId $app.ClientId -ClientSecret $app.ClientSecret -TenantId "contoso.com"

    Uses application permissions to browse all SharePoint sites with full file access.
    No personal site membership required. Defaults to site picker mode.

.EXAMPLE
    Find-DriveItemDuplicates -ClientId $app.ClientId -ClientSecret $app.ClientSecret -TenantId "contoso.com" -IncludeStorageMetrics

    App auth site picker with storage metrics. Because app permissions include both
    Sites.Read.All and Reports.Read.All, all sites and storage data are accessible.

.EXAMPLE
    Find-DriveItemDuplicates -ClientId "00000000-..." -TenantId "contoso.com" -Upn user@contoso.com

    Uses application permissions to scan a specific user's OneDrive. The -ClientSecret
    is omitted, so you will be prompted via Get-Credential.

.NOTES
    Author: Mike Crowley
    https://mikecrowley.us

    Prerequisites
        Module:       Microsoft.Graph.Authentication
        Permissions:  Files.Read or Sites.Read (UPN mode, delegated)
                      Sites.Read.All (Site Picker mode, delegated)
                      Sites.Read.All, User.Read.All (Site Picker with OneDrive, delegated)
                      Reports.Read.All (IncludeStorageMetrics)
                      Application.ReadWrite.All (RegisterApp only)
                      Sites.Read.All, Reports.Read.All (App Auth — granted automatically)

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
    [Parameter(Mandatory, ParameterSetName = 'RegisterApp')]
    [switch]$RegisterApp,

    [Parameter(Mandatory, ParameterSetName = 'AppAuth')]
    [ValidateNotNullOrEmpty()]
    [string]$ClientId,

    [Parameter(ParameterSetName = 'AppAuth')]
    [string]$ClientSecret,

    [Parameter(Mandatory, ParameterSetName = 'RegisterApp')]
    [Parameter(Mandatory, ParameterSetName = 'AppAuth')]
    [ValidateNotNullOrEmpty()]
    [string]$TenantId,

    [Parameter(ParameterSetName = 'UPN')]
    [Parameter(ParameterSetName = 'AppAuth')]
    [string]$Upn,

    [Parameter(Mandatory, ParameterSetName = 'SitePicker')]
    [switch]$SitePicker,

    [Parameter(ParameterSetName = 'SitePicker')]
    [Parameter(ParameterSetName = 'AppAuth')]
    [ValidateRange(1, 5000)]
    [int32]$SiteCount = 500,

    [Parameter(ParameterSetName = 'SitePicker')]
    [Parameter(ParameterSetName = 'AppAuth')]
    [switch]$IncludeStorageMetrics,

    [string]$RootPath = "/",
    [switch]$NoRecursion,

    [ValidateSet("Report", "PassThru", "ReportAndPassThru")]
    [string]$OutputStyle = "Report",

    [int32]$ResultSize = [int16]::MaxValue,
    [switch]$Silent
)

function Find-DriveItemDuplicates {
    [CmdletBinding(DefaultParameterSetName = 'UPN')]
    param(
        [Parameter(Mandatory, ParameterSetName = 'RegisterApp')]
        [switch]$RegisterApp,

        [Parameter(Mandatory, ParameterSetName = 'AppAuth')]
        [ValidateNotNullOrEmpty()]
        [string]$ClientId,

        [Parameter(ParameterSetName = 'AppAuth')]
        [string]$ClientSecret,

        [Parameter(Mandatory, ParameterSetName = 'RegisterApp')]
        [Parameter(Mandatory, ParameterSetName = 'AppAuth')]
        [ValidateNotNullOrEmpty()]
        [string]$TenantId,

        [Parameter(ParameterSetName = 'UPN')]
        [Parameter(ParameterSetName = 'AppAuth')]
        [string]$Upn,

        [Parameter(Mandatory, ParameterSetName = 'SitePicker')]
        [switch]$SitePicker,

        [Parameter(ParameterSetName = 'SitePicker')]
        [Parameter(ParameterSetName = 'AppAuth')]
        [ValidateRange(1, 5000)]
        [int32]$SiteCount = 500,

        [Parameter(ParameterSetName = 'SitePicker')]
        [Parameter(ParameterSetName = 'AppAuth')]
        [switch]$IncludeStorageMetrics,

        [string]$RootPath = "/",
        [switch]$NoRecursion,

        [ValidateSet("Report", "PassThru", "ReportAndPassThru")]
        [string]$OutputStyle = "Report",

        [int32]$ResultSize = [int16]::MaxValue,
        [switch]$Silent
    )

    $StartTime = Get-Date
    $script:MoreFilesExist = $false

    #region Module check
    if ($null -eq (Get-Command Invoke-MgGraphRequest -ErrorAction SilentlyContinue)) {
        throw "Invoke-MgGraphRequest cmdlet not found. Install the Microsoft.Graph.Authentication PowerShell module.`nhttps://learn.microsoft.com/en-us/graph/sdks/sdk-installation#install-the-microsoft-graph-powershell-sdk"
    }
    #endregion

    #region Register App
    if ($PSCmdlet.ParameterSetName -eq 'RegisterApp') {
        Write-Host "Connecting to Microsoft Graph to register the application..." -ForegroundColor Green
        Connect-MgGraph -Scopes "Application.ReadWrite.All", "AppRoleAssignment.ReadWrite.All" -TenantId $TenantId -ContextScope "process" -NoWelcome

        # Look up the Sites.Read.All and Reports.Read.All app roles from the Graph service principal
        $GraphAppId = "00000003-0000-0000-c000-000000000000"
        $GraphSPResponse = Invoke-MgGraphRequest -Uri "v1.0/servicePrincipals?`$filter=appId eq '$GraphAppId'" -OutputType PSObject
        $GraphSP = $GraphSPResponse.value[0]
        $SitesReadAllRole = $GraphSP.appRoles | Where-Object { $_.value -eq "Sites.Read.All" }
        $ReportsReadAllRole = $GraphSP.appRoles | Where-Object { $_.value -eq "Reports.Read.All" }

        if (-not $SitesReadAllRole) {
            Write-Error "Could not find the Sites.Read.All app role on the Microsoft Graph service principal."
            return
        }

        # Build required permissions
        $resourceAccess = @(
            @{ id = $SitesReadAllRole.id; type = "Role" }
        )
        if ($ReportsReadAllRole) {
            $resourceAccess += @{ id = $ReportsReadAllRole.id; type = "Role" }
        }

        # Create the app registration
        $AppBody = @{
            displayName            = "Find-DriveItemDuplicates"
            signInAudience         = "AzureADMyOrg"
            requiredResourceAccess = @(
                @{
                    resourceAppId  = $GraphAppId
                    resourceAccess = $resourceAccess
                }
            )
        } | ConvertTo-Json -Depth 5

        $App = Invoke-MgGraphRequest -Method POST -Uri "v1.0/applications" -Body $AppBody -ContentType "application/json" -OutputType PSObject

        # Create a client secret (90-day expiration)
        $SecretBody = @{
            passwordCredential = @{
                displayName = "Find-DriveItemDuplicates secret"
                endDateTime = (Get-Date).AddDays(90).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
            }
        } | ConvertTo-Json -Depth 3

        $Secret = Invoke-MgGraphRequest -Method POST -Uri "v1.0/applications/$($App.id)/addPassword" -Body $SecretBody -ContentType "application/json" -OutputType PSObject

        # Create the service principal so the app is usable for client credential auth
        $SPBody = @{ appId = $App.appId } | ConvertTo-Json
        $SP = Invoke-MgGraphRequest -Method POST -Uri "v1.0/servicePrincipals" -Body $SPBody -ContentType "application/json" -OutputType PSObject

        # Grant admin consent for each app role
        $ConsentGranted = $false
        try {
            foreach ($role in $resourceAccess) {
                $ConsentBody = @{
                    principalId = $SP.id
                    resourceId  = $GraphSP.id
                    appRoleId   = $role.id
                } | ConvertTo-Json
                $null = Invoke-MgGraphRequest -Method POST -Uri "v1.0/servicePrincipals/$($SP.id)/appRoleAssignments" -Body $ConsentBody -ContentType "application/json" -OutputType PSObject
            }
            $ConsentGranted = $true
        }
        catch {
            Write-Warning "Could not grant admin consent automatically. An admin must grant consent manually:`n  Azure Portal > App Registrations > Find-DriveItemDuplicates > API Permissions > Grant admin consent"
        }

        $permNames = @('Sites.Read.All')
        if ($ReportsReadAllRole) { $permNames += 'Reports.Read.All' }

        Write-Host "`nApp Registration Created Successfully" -ForegroundColor Green
        Write-Host "  Display Name : $($App.displayName)" -ForegroundColor Cyan
        Write-Host "  Client ID    : $($App.appId)" -ForegroundColor Cyan
        Write-Host "  Client Secret: $($Secret.secretText)" -ForegroundColor Cyan
        Write-Host "  Secret Expiry: $($Secret.endDateTime)" -ForegroundColor Cyan
        Write-Host "  Permissions  : $($permNames -join ', ')" -ForegroundColor Cyan
        if ($ConsentGranted) {
            Write-Host "  Admin Consent: Granted" -ForegroundColor Green
        }
        Write-Host "`n  Save the Client ID and Secret now. The secret value cannot be retrieved later." -ForegroundColor Yellow

        return [PSCustomObject]@{
            DisplayName    = $App.displayName
            ClientId       = $App.appId
            ObjectId       = $App.id
            ClientSecret   = $Secret.secretText
            SecretExpiry   = $Secret.endDateTime
            ConsentGranted = $ConsentGranted
        }
    }
    #endregion

    #region App auth connection
    $IsAppAuth = $PSCmdlet.ParameterSetName -eq 'AppAuth'
    if ($IsAppAuth) {
        $null = Disconnect-MgGraph -ErrorAction SilentlyContinue

        if ($ClientSecret) {
            $SecureSecret = ConvertTo-SecureString $ClientSecret -AsPlainText -Force
            $ClientSecretCredential = [PSCredential]::new($ClientId, $SecureSecret)
        }
        else {
            $ClientSecretCredential = Get-Credential -Credential $ClientId
        }

        try {
            Connect-MgGraph -ClientSecretCredential $ClientSecretCredential -TenantId $TenantId -ContextScope "process" -NoWelcome -ErrorAction Stop
        }
        catch {
            Write-Error "App authentication failed. If the app was just registered, wait a minute for propagation and verify an admin has granted consent. $_"
            return
        }

        if (-not $Silent) { Write-Host "Connected with application permissions." -ForegroundColor Green }
    }
    #endregion

    #region Pre-flight checks
    if ($null -eq (Get-MgContext)) {
        throw "No Graph context found. Please call Connect-MgGraph, or use -ClientId/-TenantId for app auth."
    }

    # Determine routing
    $useSitePicker = ($PSCmdlet.ParameterSetName -eq 'SitePicker') -or
                     ($IsAppAuth -and -not $Upn)

    if (-not $IsAppAuth) {
        $scopes = (Get-MgContext).Scopes | Out-String
        if ($useSitePicker) {
            if ($scopes -notlike '*Sites.Read.All*') {
                Write-Warning "Sites.Read.All scope may be required for site picker.`nhttps://learn.microsoft.com/en-us/graph/api/site-search"
            }
            if ($IncludeStorageMetrics -and $scopes -notlike '*Reports.Read.All*') {
                Write-Warning "Reports.Read.All scope is required for -IncludeStorageMetrics.`nhttps://learn.microsoft.com/en-us/graph/api/reportroot-getsharepointsiteusagedetail"
            }
        }
        else {
            if ($scopes -notlike '*Files.Read*' -and $scopes -notlike '*Sites.Read*') {
                Write-Warning "Permission scope may be missing.`nhttps://learn.microsoft.com/en-us/graph/api/driveitem-list-children?view=graph-rest-1.0&tabs=http#permissions"
            }
        }
    }
    #endregion

    #region Drive resolution
    if ($useSitePicker) {
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

        # Three enumeration strategies, tried in order:
        #   1. getAllSites          — complete inventory (needs admin-consented Sites.Read.All)
        #   2. Reports API + batch  — complete inventory from getSharePointSiteUsageDetail,
        #                             resolving display names via JSON-batched site lookups
        #                             (needs Reports.Read.All; only when -IncludeStorageMetrics)
        #   3. Search API           — partial results (standard delegated permissions)
        if (-not $Silent) { Write-Host "Fetching SharePoint sites..." -ForegroundColor Cyan }
        $storageAlreadyMerged = $false

        # Probe whether getAllSites is accessible
        $getAllSitesOk = $true
        try {
            $null = Invoke-MgGraphRequest -Uri "v1.0/sites/getAllSites?`$top=1&`$select=id" -ErrorAction Stop
        }
        catch {
            if ($_.Exception.Message -match '403|Forbidden|accessDenied') { $getAllSitesOk = $false }
            else { throw }
        }

        if ($getAllSitesOk) {
            # Path 1: getAllSites — complete enumeration
            if (-not $Silent) { Write-Host "Using getAllSites for complete enumeration." -ForegroundColor DarkGray }
            $siteUri = "v1.0/sites/getAllSites?`$top=999&`$select=id,displayName,webUrl,createdDateTime,lastModifiedDateTime"
            $siteCollCounter = 0
            do {
                $siteResponse = Invoke-MgGraphRequest -Uri $siteUri
                foreach ($site in $siteResponse.value) {
                    if ($IncludeStorageMetrics) {
                        $pathParts = ([uri]$site.webUrl).AbsolutePath.Trim('/') -split '/' | Where-Object { $_ }
                        $isSiteCollection = ($pathParts.Count -eq 0) -or
                            ($pathParts.Count -eq 2 -and $pathParts[0] -in @('sites', 'teams'))
                        if (-not $isSiteCollection) { continue }
                        $siteCollCounter++
                        if ($siteCollCounter -gt $SiteCount) { break }
                    }
                    else {
                        if ($SiteList.Count -ge $SiteCount) { break }
                    }
                    $SiteList.Add([PSCustomObject]@{
                        Type          = "SharePoint"
                        DisplayName   = $site.displayName
                        WebUrl        = $site.webUrl
                        SiteCollGB    = ''
                        SiteCollFiles = ''
                        Created       = if ($site.createdDateTime) { ([datetime]$site.createdDateTime).ToString('yyyy-MM-dd') } else { '' }
                        SiteId        = $site.id
                        Upn           = ""
                    })
                }
                $siteUri = $siteResponse.'@odata.nextLink'
                $limitReached = if ($IncludeStorageMetrics) { $siteCollCounter -ge $SiteCount } else { $SiteList.Count -ge $SiteCount }
            } until (-not $siteUri -or $limitReached)
        }
        elseif ($IncludeStorageMetrics) {
            # Path 2: Reports API + JSON batching — complete enumeration with storage.
            # The Reports API lists every site collection with storage data; batched
            # GET /sites/{hostname},{siteId} calls resolve display names and URLs.
            if (-not $Silent) {
                Write-Warning "sites/getAllSites requires admin consent. Building site list from Reports API."
            }

            # Derive the SharePoint hostname from the root site
            $rootSite = Invoke-MgGraphRequest -Uri "v1.0/sites/root?`$select=webUrl" -ErrorAction Stop
            $spHostname = ([uri]$rootSite.webUrl).Host

            # Fetch the full usage report
            $tempCsv = Join-Path ([System.IO.Path]::GetTempPath()) "spo_report_$(Get-Random).csv"
            Invoke-MgGraphRequest -Uri "v1.0/reports/getSharePointSiteUsageDetail(period='D7')" -OutputFilePath $tempCsv -ErrorAction Stop
            $reportData = @(Import-Csv $tempCsv)
            Remove-Item $tempCsv -Force -ErrorAction SilentlyContinue

            if (-not $Silent) { Write-Host "Report returned $($reportData.Count) site rows. Resolving details via JSON batching..." -ForegroundColor Cyan }

            # Sort by storage descending and cap at SiteCount
            $reportData = @($reportData |
                Sort-Object { if ($_.'Storage Used (Byte)') { [double]$_.'Storage Used (Byte)' } else { 0 } } -Descending |
                Select-Object -First $SiteCount)

            # Batch-resolve site details (20 per batch — Graph maximum)
            $batchSize = 20
            for ($i = 0; $i -lt $reportData.Count; $i += $batchSize) {
                $end = [math]::Min($i + $batchSize - 1, $reportData.Count - 1)
                $chunk = @($reportData[$i..$end])
                $batchBody = @{
                    requests = @(foreach ($row in $chunk) {
                        @{
                            id     = $row.'Site Id'
                            method = "GET"
                            url    = "/sites/$spHostname,$($row.'Site Id')?`$select=id,displayName,webUrl,createdDateTime,lastModifiedDateTime"
                        }
                    })
                }
                $batchJson = $batchBody | ConvertTo-Json -Depth 4 -Compress
                $batchResponse = Invoke-MgGraphRequest -Uri "v1.0/`$batch" -Method POST -Body $batchJson -ContentType "application/json"

                foreach ($resp in $batchResponse.responses) {
                    if ($resp.status -ne 200) { continue }
                    $site = $resp.body

                    # Filter to site collections (managed paths only)
                    $pathParts = ([uri]$site.webUrl).AbsolutePath.Trim('/') -split '/' | Where-Object { $_ }
                    $isSiteCollection = ($pathParts.Count -eq 0) -or
                        ($pathParts.Count -eq 2 -and $pathParts[0] -in @('sites', 'teams'))
                    if (-not $isSiteCollection) { continue }

                    # Attach storage metrics from the matching report row
                    $reportRow = $chunk | Where-Object { $_.'Site Id' -eq $resp.id } | Select-Object -First 1
                    $storageGB = ''
                    $fileCount = ''
                    if ($reportRow) {
                        if ($reportRow.'Storage Used (Byte)' -and $reportRow.'Storage Used (Byte)' -ne '') {
                            $storageGB = [math]::Round([double]$reportRow.'Storage Used (Byte)' / 1GB, 2)
                        }
                        if ($reportRow.'File Count' -and $reportRow.'File Count' -ne '') {
                            $fileCount = [int]$reportRow.'File Count'
                        }
                    }

                    $SiteList.Add([PSCustomObject]@{
                        Type          = "SharePoint"
                        DisplayName   = $site.displayName
                        WebUrl        = $site.webUrl
                        SiteCollGB    = $storageGB
                        SiteCollFiles = $fileCount
                        Created       = if ($site.createdDateTime) { ([datetime]$site.createdDateTime).ToString('yyyy-MM-dd') } else { '' }
                        SiteId        = $site.id
                        Upn           = ""
                    })
                }

                if (-not $Silent) {
                    $pct = [math]::Min(100, [int](($i + $batchSize) / $reportData.Count * 100))
                    Write-Progress -Id 2 -Activity "Resolving site details" -Status "$($SiteList.Count) site collections found" -PercentComplete $pct
                }
            }
            if (-not $Silent) { Write-Progress -Id 2 -Activity "Resolving site details" -Completed }
            $storageAlreadyMerged = $true
        }
        else {
            # Path 3: Search API fallback (partial results)
            if (-not $Silent) {
                Write-Warning "sites/getAllSites requires admin consent. Falling back to search API."
                Write-Host "  Tip: Use -IncludeStorageMetrics for complete site enumeration, or reconnect with admin consent." -ForegroundColor DarkGray
            }
            $searchTerm = ((Get-MgContext).Account -split '@')[-1] -split '\.' | Select-Object -First 1
            if (-not $Silent) { Write-Host "Searching SharePoint sites ($searchTerm)..." -ForegroundColor Cyan }
            $siteUri = "v1.0/sites?search=$searchTerm&`$top=$SiteCount&`$select=id,displayName,webUrl,createdDateTime,lastModifiedDateTime"
            do {
                $siteResponse = Invoke-MgGraphRequest -Uri $siteUri
                foreach ($site in $siteResponse.value) {
                    if ($SiteList.Count -ge $SiteCount) { break }
                    $SiteList.Add([PSCustomObject]@{
                        Type          = "SharePoint"
                        DisplayName   = $site.displayName
                        WebUrl        = $site.webUrl
                        SiteCollGB    = ''
                        SiteCollFiles = ''
                        Created       = if ($site.createdDateTime) { ([datetime]$site.createdDateTime).ToString('yyyy-MM-dd') } else { '' }
                        SiteId        = $site.id
                        Upn           = ""
                    })
                }
                $siteUri = $siteResponse.'@odata.nextLink'
            } until (-not $siteUri -or $SiteList.Count -ge $SiteCount)
        }

        if ($includeOneDrive) {
            if (-not $Silent) { Write-Host "Fetching OneDrive users..." -ForegroundColor Cyan }
            $userUri = "v1.0/users?`$top=100&`$select=id,displayName,userPrincipalName"
            do {
                $userResponse = Invoke-MgGraphRequest -Uri $userUri
                foreach ($user in $userResponse.value) {
                    if ($SiteList.Count -ge $SiteCount) { break }
                    $SiteList.Add([PSCustomObject]@{
                        Type          = "OneDrive"
                        DisplayName   = "$($user.displayName) (OneDrive)"
                        WebUrl        = $user.userPrincipalName
                        SiteCollGB = ''
                        SiteCollFiles     = ''
                        Created       = ''
                        SiteId        = ""
                        Upn           = $user.userPrincipalName
                    })
                }
                $userUri = $userResponse.'@odata.nextLink'
            } until (-not $userUri -or $SiteList.Count -ge $SiteCount)
        }

        if ($SiteList.Count -eq 0) {
            throw "No sites found. Verify your permissions (Sites.Read.All)."
        }

        if (-not $Silent) { Write-Host "Found $($SiteList.Count) sites." -ForegroundColor Cyan }

        # Step B2 (optional): Fetch storage metrics from Graph Reports API
        # Skipped when Path 2 (Reports API + batch) already merged storage data.
        if ($IncludeStorageMetrics -and -not $storageAlreadyMerged) {
            if (-not $Silent) { Write-Host "Fetching storage metrics from Reports API..." -ForegroundColor Cyan }

            try {
                # getSharePointSiteUsageDetail returns a 302 redirect to a CSV download.
                # -OutputFilePath writes the response directly to disk, avoiding JSON parse issues.
                $tempCsv = Join-Path ([System.IO.Path]::GetTempPath()) "spo_report_$(Get-Random).csv"
                Invoke-MgGraphRequest -Uri "v1.0/reports/getSharePointSiteUsageDetail(period='D7')" -OutputFilePath $tempCsv -ErrorAction Stop

                # Verify the file was created and contains CSV (not a JSON error)
                if (-not (Test-Path $tempCsv)) {
                    Write-Warning "-IncludeStorageMetrics: Report file was not created."
                }
                else {
                    $firstLine = Get-Content $tempCsv -TotalCount 1
                    if (-not $Silent) { Write-Host "Report columns: $firstLine" -ForegroundColor DarkGray }

                    if ($firstLine -match '^\s*\{') {
                        Write-Warning "-IncludeStorageMetrics: Report API returned JSON instead of CSV (possible auth error).`n$firstLine"
                    }
                    else {
                        $reportData = @(Import-Csv $tempCsv)
                        Remove-Item $tempCsv -Force -ErrorAction SilentlyContinue

                        if (-not $Silent) { Write-Host "Report returned $($reportData.Count) site rows." -ForegroundColor Cyan }

                        if ($reportData.Count -eq 0) {
                            Write-Warning "-IncludeStorageMetrics: Reports API returned no data rows."
                        }
                        else {
                            # Determine matching strategy. When the tenant conceals user details
                            # in reports, Site URL is blank but Site Id (a GUID) is still present.
                            # Graph sites API returns IDs as "hostname,siteCollectionGuid,webGuid".
                            $sampleUrl = ($reportData | Select-Object -First 1).'Site URL'
                            $useUrlMatch = [bool]$sampleUrl

                            if ($useUrlMatch) {
                                # Build a case-insensitive lookup by Site URL for exact and prefix matching.
                                # The Reports API reports at site-collection level, but Graph search returns
                                # subsites too (e.g., /sites/Parent/teams/child). For subsites, we fall back
                                # to the parent site-collection's metrics.
                                if (-not $Silent) { Write-Host "Matching by Site URL..." -ForegroundColor DarkGray }
                                $storageLookup = [System.Collections.Generic.Dictionary[string,object]]::new([System.StringComparer]::OrdinalIgnoreCase)
                                foreach ($row in $reportData) {
                                    $siteUrl = $row.'Site URL'
                                    if ($siteUrl) {
                                        $storageLookup[$siteUrl.TrimEnd('/')] = $row
                                    }
                                }
                            }
                            else {
                                # Concealed-data fallback: match by Site Id GUID.
                                if (-not $Silent) { Write-Host "Site URLs are concealed. Matching by Site Id..." -ForegroundColor DarkGray }
                                $storageLookup = [System.Collections.Generic.Dictionary[string,object]]::new([System.StringComparer]::OrdinalIgnoreCase)
                                foreach ($row in $reportData) {
                                    $siteId = $row.'Site Id'
                                    if ($siteId) {
                                        $storageLookup[$siteId] = $row
                                    }
                                }
                            }

                            # The Reports API only has site-collection-level data. Subsites
                            # inherit their parent collection's totals. The column names
                            # (SiteCollGB, SiteCollFiles) make this aggregation explicit.
                            $matchCount = 0
                            foreach ($site in $SiteList) {
                                if ($site.Type -eq 'OneDrive') { continue }

                                $reportRow = $null
                                if ($useUrlMatch) {
                                    # Exact URL match; subsites try prefix match to parent collection
                                    $lookupKey = $site.WebUrl.TrimEnd('/')
                                    if ($storageLookup.ContainsKey($lookupKey)) {
                                        $reportRow = $storageLookup[$lookupKey]
                                    }
                                    else {
                                        $bestMatch = $null
                                        $bestLen = 0
                                        foreach ($reportUrl in $storageLookup.Keys) {
                                            if ($lookupKey.StartsWith($reportUrl, [System.StringComparison]::OrdinalIgnoreCase) -and $reportUrl.Length -gt $bestLen) {
                                                $bestMatch = $reportUrl
                                                $bestLen = $reportUrl.Length
                                            }
                                        }
                                        if ($bestMatch) { $reportRow = $storageLookup[$bestMatch] }
                                    }
                                }
                                else {
                                    # GUID match — use site-collection GUID (middle part of compound ID)
                                    $idParts = $site.SiteId -split ','
                                    if ($idParts.Count -ge 2) {
                                        $siteCollectionGuid = $idParts[1]
                                        if ($storageLookup.ContainsKey($siteCollectionGuid)) {
                                            $reportRow = $storageLookup[$siteCollectionGuid]
                                        }
                                    }
                                }

                                if ($reportRow) {
                                    $storageBytes = $reportRow.'Storage Used (Byte)'
                                    if ($storageBytes -and $storageBytes -ne '') {
                                        $site.SiteCollGB = [math]::Round([double]$storageBytes / 1GB, 2)
                                    }
                                    $fileCount = $reportRow.'File Count'
                                    if ($fileCount -and $fileCount -ne '') {
                                        $site.SiteCollFiles = [int]$fileCount
                                    }
                                    $matchCount++
                                }
                            }

                            if (-not $Silent) {
                                $spSites = @($SiteList | Where-Object { $_.Type -ne 'OneDrive' })
                                Write-Host "Storage metrics: matched $matchCount of $($spSites.Count) SharePoint sites." -ForegroundColor Cyan
                            }

                            if ($matchCount -eq 0) {
                                Write-Warning "-IncludeStorageMetrics: Could not match any sites. Report has $($reportData.Count) rows but no matches found against $($spSites.Count) sites."
                            }
                        }
                    }
                }
            }
            catch {
                $msg = "$($_.Exception.Message) $($_.ErrorDetails.Message)"
                if ($msg -match '403|Forbidden|Authorization') {
                    Write-Warning "Reports.Read.All permission is required for -IncludeStorageMetrics.`nhttps://learn.microsoft.com/en-us/graph/api/reportroot-getsharepointsiteusagedetail"
                }
                else {
                    Write-Warning "-IncludeStorageMetrics failed: $($_.Exception.Message)"
                }
            }
        }

        # Step C: Present site picker (sort by SiteCollGB descending when available)
        # When storage metrics are shown, filter to site collections only (exclude subsites)
        # so each row has unique storage values. In SharePoint Online, site collections
        # exist only at managed paths: /sites/{Name} or /teams/{Name}. Everything else
        # (e.g., /Purchasing, /Safety/Ottawa) is a subsite of the root site collection.
        $pickerList = if ($IncludeStorageMetrics) {
            @($SiteList | Where-Object {
                if ($_.Type -eq 'OneDrive') { $true }
                else {
                    $pathParts = ([uri]$_.WebUrl).AbsolutePath.Trim('/') -split '/' | Where-Object { $_ }
                    # Root site (0 segments) or /sites/{name} or /teams/{name} (exactly 2 segments)
                    $pathParts.Count -eq 0 -or
                    ($pathParts.Count -eq 2 -and $pathParts[0] -in @('sites', 'teams'))
                }
            })
        } else {
            $SiteList
        }

        $gridColumns = @('Type', 'DisplayName', 'WebUrl')
        if ($IncludeStorageMetrics) { $gridColumns += 'SiteCollGB', 'SiteCollFiles' }
        $gridColumns += 'Created'

        $sortedSiteList = if ($IncludeStorageMetrics) {
            $pickerList | Sort-Object { if ($_.SiteCollGB -is [double]) { $_.SiteCollGB } else { -1 } } -Descending
        } else {
            $pickerList
        }

        if (-not $Silent -and $pickerList.Count -ne $SiteList.Count) {
            Write-Host "Showing $($pickerList.Count) site collections (filtered $($SiteList.Count - $pickerList.Count) subsites)." -ForegroundColor DarkGray
        }

        $SelectedSite = $sortedSiteList |
            Select-Object $gridColumns |
            Out-GridView -Title "Select a site to scan for duplicates ($($pickerList.Count) site collections)" -OutputMode Single

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
            $DriveResponse = Invoke-MgGraphRequest -Uri "beta/users/$($SelectedSite.Upn)/drive?`$select=id,webUrl"
            $DriveId = $DriveResponse.id
            $DriveWebUrl = $DriveResponse.webUrl
            $reportIdentifier = $SelectedSite.Upn
        }
        else {
            $drivesResponse = Invoke-MgGraphRequest -Uri "beta/sites/$($SelectedSite.SiteId)/drives?`$select=id,name,webUrl"

            # Filter out system/publishing libraries that wouldn't contain user documents
            $systemLibraryNames = [Collections.Generic.HashSet[string]]::new(
                [string[]]@('Pages', 'Images', 'Slideshow', 'Site Pages', 'Style Library',
                    'Form Templates', 'FormServerTemplates', 'Site Collection Documents',
                    'Site Collection Images', 'PublishingImages', 'SiteCollectionImages',
                    '_catalogs', 'Converted Forms', 'appdata', 'appfiles'),
                [StringComparer]::OrdinalIgnoreCase
            )
            $drives = @($drivesResponse.value | Where-Object { -not $systemLibraryNames.Contains($_.name) })

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
            $DriveResponse = Invoke-MgGraphRequest -Uri "beta/users/$Upn/drive?`$select=id,webUrl" -ErrorAction Stop
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
    $script:AccessDeniedPaths = [Collections.Generic.List[string]]::new()

    # Determine whether we can show a percentage (only when ResultSize is not the default)
    $showPercent = $ResultSize -lt [int16]::MaxValue
    $progressId = 1
    $childSelect = '?$select=id,name,size,webUrl,lastModifiedDateTime,file,folder'

    if ($NoRecursion) {
        $uri = if ($RootPath -eq "/") {
            "beta/drives/$DriveId/root/children$childSelect"
        } else {
            "beta/drives/$DriveId/root:/$($RootPath):/children$childSelect"
        }
        $RawItems = [Collections.Generic.List[Object]]::new()

        do {
            if (-not $Silent) {
                $progressParams = @{
                    Id              = $progressId
                    Activity        = "Scanning files"
                    Status          = "$($RawItems.Count) items retrieved"
                    CurrentOperation = $RootPath
                }
                if ($showPercent) { $progressParams.PercentComplete = [math]::Min(100, [int]($RawItems.Count / $ResultSize * 100)) }
                Write-Progress @progressParams
            }
            try {
                $PageResults = Invoke-MgGraphRequest -Uri $uri -ErrorAction Stop
            }
            catch {
                $errMsg = "$($_.Exception.Message) $($_.ErrorDetails.Message)"
                if ($errMsg -match '403|Forbidden|accessDenied') {
                    $script:AccessDeniedPaths.Add($RootPath)
                    break
                }
                else {
                    Write-Warning "Failed to read '$RootPath': $($_.Exception.Message)"
                    break
                }
            }
            if ($PageResults.value) {
                $RawItems.AddRange($PageResults.value)
            }
            else {
                $RawItems.Add($PageResults)
            }
            $uri = $PageResults.'@odata.nextlink'
        } until (-not $uri -or $RawItems.Count -ge $ResultSize)

        if (-not $Silent) { Write-Progress -Id $progressId -Activity "Scanning files" -Completed }

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
                $uri = "$baseUri/children$childSelect"
            }
            else {
                $uri = "$baseUri`:/$Path`:/children$childSelect"
            }

            do {
                if (-not $Silent) {
                    $progressParams = @{
                        Id              = $progressId
                        Activity        = "Scanning files"
                        Status          = "$($AllFiles.Count) files found"
                        CurrentOperation = $Path
                    }
                    if ($showPercent) { $progressParams.PercentComplete = [math]::Min(100, [int]($AllFiles.Count / $ResultSize * 100)) }
                    Write-Progress @progressParams
                }

                try {
                    $Response = Invoke-MgGraphRequest -Uri $Uri -ErrorAction Stop
                }
                catch {
                    $errMsg = "$($_.Exception.Message) $($_.ErrorDetails.Message)"
                    if ($errMsg -match '403|Forbidden|accessDenied') {
                        $script:AccessDeniedPaths.Add($Path)
                        return   # skip this folder
                    }
                    else {
                        Write-Warning "Failed to read '$Path': $($_.Exception.Message)"
                        return
                    }
                }

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
        if (-not $Silent) { Write-Progress -Id $progressId -Activity "Scanning files" -Completed }
    }

    # Access-denied diagnostic — warn when the scan returned suspiciously few files
    if ($script:AccessDeniedPaths.Count -gt 0 -and -not $Silent) {
        Write-Host ""
        Write-Warning "Access denied on $($script:AccessDeniedPaths.Count) path(s). Your account may not have Read access to this site's content."
        Write-Host "  Tip: Even with Sites.Read.All, delegated Graph sessions can only see files" -ForegroundColor DarkGray
        Write-Host "  your account can access in each site. To scan all files, either:" -ForegroundColor DarkGray
        Write-Host "    1. Grant your account Read access on the target site (SharePoint Admin Center)" -ForegroundColor DarkGray
        Write-Host "    2. Use application-only auth (Connect-MgGraph -ClientId ... -CertificateThumbprint ...)" -ForegroundColor DarkGray
        Write-Host ""
    }
    elseif ($useSitePicker -and $FileList.Count -le 5 -and -not $Silent) {
        Write-Host ""
        Write-Warning "Only $($FileList.Count) file(s) found. Your account may not have Read access to this site's content."
        Write-Host "  Tip: Delegated Graph sessions can only see files your account can access in" -ForegroundColor DarkGray
        Write-Host "  each site. To scan all files, grant your account Read access on the target" -ForegroundColor DarkGray
        Write-Host "  site, or use application-only auth." -ForegroundColor DarkGray
        Write-Host ""
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
        $(if ($script:AccessDeniedPaths.Count -gt 0) {
            "<p style='color: #ff6b6b; margin-top: 30px;'><strong>Access Denied:</strong> $($script:AccessDeniedPaths.Count) path(s) could not be read. Your account may not have Read access to this site's content. Even with Sites.Read.All, delegated Graph sessions can only read files your account is permitted to access in each site. Grant your account Read access on the target site, or use application-only authentication.</p>"
        })
        $(if ($useSitePicker -and $FileList.Count -le 5 -and $script:AccessDeniedPaths.Count -eq 0) {
            "<p style='color: #fbbf24; margin-top: 30px;'><strong>Low file count:</strong> Only $($FileList.Count) file(s) found. Your account may not have full Read access to this site. Delegated Graph sessions can only read files your account is permitted to access.</p>"
        })
    </div>
</body>
</html>
"@

        try {
            if ($Output.Count -gt 0) {
                $Output | Export-Csv $csvPath -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
                $Output | ConvertTo-Json -Depth 3 | Out-File $jsonPath -Encoding UTF8 -ErrorAction Stop
            }
            else {
                # Write header-only CSV and empty JSON array so files aren't 0 bytes
                'Confidence,MatchType,MatchValue,NumberOfFiles,FileSizeKB,FileGroupSizeKB,PossibleWasteKB,FileNames,WebLocalPaths' |
                    Out-File $csvPath -Encoding UTF8 -ErrorAction Stop
                '[]' | Out-File $jsonPath -Encoding UTF8 -ErrorAction Stop
            }
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
        if ($script:AccessDeniedPaths.Count -gt 0) {
            Write-Host "Warning: Access denied on $($script:AccessDeniedPaths.Count) path(s). Results may be incomplete." -ForegroundColor Yellow
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
