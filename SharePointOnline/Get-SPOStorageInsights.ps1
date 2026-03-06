<#
.SYNOPSIS
    Analyzes SharePoint Online and OneDrive storage for duplicate files, version history
    bloat, and preservation hold libraries using the Microsoft Graph API.

.DESCRIPTION
    Get-SPOStorageInsights enumerates files in SharePoint Online and OneDrive document libraries
    and performs three categories of storage analysis:

        Duplicate Detection (dual confidence)
            High Confidence  -  Files whose quickXorHash content hashes match.
            Low  Confidence  -  Files that lack a hash but share an identical filename.

        Version History Analysis
            Fetches version history for the largest files and reports total version storage
            consumed. Controlled by -VersionCheckCount (default 100). Set to 0 to skip.

        Preservation Hold Library Detection
            Identifies sites that have a Preservation Hold Library and reports its size.

    Output is written to a timestamped folder on the desktop containing an interactive HTML
    dashboard and separate CSV/JSON exports for each analysis type. Results may also be
    returned to the pipeline, controlled by -OutputStyle.

    The site picker supports multi-select: choose one or more sites to analyze in a single
    run. Results are combined into a single report with a SiteName column.

    The script can be dot-sourced to load the function, or invoked directly.

    Two operating modes are available:

        Register App Mode  -  Use -RegisterApp once per tenant to create an Entra ID app
                              registration with the required permissions and a client secret.
        App Auth Mode      -  Connects with application (client credential) permissions for
                              full read access to all sites and drives. Supply -ClientId and
                              -TenantId (plus optionally -ClientSecret). Presents an interactive
                              site picker (Out-GridView) to browse and multi-select target sites.

    App auth handles its own connection -- no prior Connect-MgGraph is needed.

.PARAMETER RegisterApp
    Creates an Entra ID app registration with Sites.Read.All and Reports.Read.All
    application permissions, a client secret (90-day expiry), a service principal,
    and grants admin consent. Run this once per tenant before using app auth.
    Requires Application.ReadWrite.All and AppRoleAssignment.ReadWrite.All
    permissions (typically Application Administrator or Global Administrator).
    Returns an object with ClientId and ClientSecret for subsequent runs.

.PARAMETER ClientId
    The Application (client) ID of the Entra ID app registration created by
    -RegisterApp. Required for scanning.

.PARAMETER ClientSecret
    The client secret for the app registration. If omitted, you will be prompted
    via Get-Credential. Accepts the plain-text secret returned by -RegisterApp,
    e.g. -ClientSecret $app.ClientSecret.

.PARAMETER TenantId
    The tenant ID or domain name for the Microsoft 365 tenant. Required for both
    -RegisterApp and scanning.

.PARAMETER SiteCount
    Maximum number of sites to retrieve for the site picker grid. When -IncludeStorageMetrics
    is used, this limits the number of site collections shown (subsites are excluded
    automatically). Accepts values from 1 to 5000. Default: 500.

.PARAMETER IncludeStorageMetrics
    Retrieves site-collection-level storage usage from the Graph Reports API
    (getSharePointSiteUsageDetail). Adds SiteCollGB and SiteCollFiles columns to the
    site picker grid. These values represent the entire site collection -- subsites share
    their parent collection's totals. Data may be up to 48 hours old.

.PARAMETER RootPath
    Starting directory path within each selected drive. Use forward-slash notation
    (e.g., "Documents/Projects"). Default: "/" (drive root).

.PARAMETER NoRecursion
    Limits the scan to the immediate contents of -RootPath. Subfolders are not traversed.

.PARAMETER OutputStyle
    Determines how results are delivered:
        Report            -  Generates desktop report files only (default).
        PassThru          -  Returns a results object to the pipeline only.
        ReportAndPassThru -  Generates reports and returns the results object.

.PARAMETER ResultSize
    Maximum number of files to evaluate per site. Default: 32767 ([int16]::MaxValue).
    When the limit is reached, a warning is displayed in the console and embedded in
    the HTML report.

.PARAMETER VersionCheckCount
    Number of largest files (by size) to check for version history. Default: 100.
    Set to 0 to skip version history analysis entirely. Accepts values from 0 to 5000.
    Files are selected across all scanned sites combined.

.PARAMETER SkipDuplicates
    Skips duplicate detection entirely. Use when you only want version history
    and/or preservation hold analysis.

.PARAMETER Silent
    Suppresses all console progress and summary output.

.EXAMPLE
    $app = Get-SPOStorageInsights -RegisterApp -TenantId "contoso.com"

    Creates an Entra ID app registration with Sites.Read.All and Reports.Read.All
    permissions. Save the returned ClientId and ClientSecret for subsequent runs.

.EXAMPLE
    Get-SPOStorageInsights -ClientId $app.ClientId -ClientSecret $app.ClientSecret -TenantId "contoso.com"

    Uses application permissions to browse all SharePoint sites with full file access.
    Multi-select enabled in the site picker.

.EXAMPLE
    Get-SPOStorageInsights -ClientId $app.ClientId -ClientSecret $app.ClientSecret -TenantId "contoso.com" -IncludeStorageMetrics

    Site picker with storage metrics columns in the grid.

.EXAMPLE
    Get-SPOStorageInsights -ClientId $app.ClientId -ClientSecret $app.ClientSecret -TenantId "contoso.com" -SkipDuplicates -VersionCheckCount 500

    Skips duplicate detection. Only performs version history analysis on the top 500
    largest files and checks for preservation hold libraries.

.EXAMPLE
    Get-SPOStorageInsights -ClientId $app.ClientId -ClientSecret $app.ClientSecret -TenantId "contoso.com" -VersionCheckCount 0

    Multi-site scan with duplicate detection only. Version history analysis is skipped,
    but preservation hold libraries are still detected.

.NOTES
    Author: Mike Crowley
    https://mikecrowley.us

    Prerequisites
        Module:       Microsoft.Graph.Authentication
        Permissions:  Sites.Read.All, Reports.Read.All (granted automatically by -RegisterApp)
                      Application.ReadWrite.All (RegisterApp only)

    No additional permissions are needed for version history or preservation hold detection
    beyond what is already required for file enumeration.

.LINK
    https://mikecrowley.us/2024/04/20/onedrive-and-sharepoint-online-file-deduplication-report-microsoft-graph-api
#>

[CmdletBinding(DefaultParameterSetName = 'AppAuth')]
param(
    [Parameter(Mandatory, ParameterSetName = 'RegisterApp')]
    [switch]$RegisterApp,

    [Parameter(Mandatory, ParameterSetName = 'AppAuth')]
    [ValidateNotNullOrEmpty()]
    [string]$ClientId,

    [Parameter(ParameterSetName = 'AppAuth')]
    [string]$ClientSecret,

    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$TenantId,

    [ValidateRange(1, 5000)]
    [int32]$SiteCount = 500,

    [switch]$IncludeStorageMetrics,

    [string]$RootPath = "/",
    [switch]$NoRecursion,

    [ValidateSet("Report", "PassThru", "ReportAndPassThru")]
    [string]$OutputStyle = "Report",

    [int32]$ResultSize = [int16]::MaxValue,

    [ValidateRange(0, 5000)]
    [int32]$VersionCheckCount = 100,

    [switch]$SkipDuplicates,
    [switch]$Silent
)

function Get-SPOStorageInsights {
    [CmdletBinding(DefaultParameterSetName = 'AppAuth')]
    param(
        [Parameter(Mandatory, ParameterSetName = 'RegisterApp')]
        [switch]$RegisterApp,

        [Parameter(Mandatory, ParameterSetName = 'AppAuth')]
        [ValidateNotNullOrEmpty()]
        [string]$ClientId,

        [Parameter(ParameterSetName = 'AppAuth')]
        [string]$ClientSecret,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$TenantId,

        [ValidateRange(1, 5000)]
        [int32]$SiteCount = 500,

        [switch]$IncludeStorageMetrics,

        [string]$RootPath = "/",
        [switch]$NoRecursion,

        [ValidateSet("Report", "PassThru", "ReportAndPassThru")]
        [string]$OutputStyle = "Report",

        [int32]$ResultSize = [int16]::MaxValue,

        [ValidateRange(0, 5000)]
        [int32]$VersionCheckCount = 100,

        [switch]$SkipDuplicates,
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
            displayName            = "Get-SPOStorageInsights"
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
                displayName = "Get-SPOStorageInsights secret"
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
            Write-Warning "Could not grant admin consent automatically. An admin must grant consent manually:`n  Azure Portal > App Registrations > Get-SPOStorageInsights > API Permissions > Grant admin consent"
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
    #endregion

    #region Initialize combined collections
    $FileList = [Collections.Generic.List[Object]]::new()
    $script:AccessDeniedPaths = [Collections.Generic.List[string]]::new()
    $PreservationHoldFindings = [Collections.Generic.List[Object]]::new()
    $sitesScanned = 0
    #endregion

    #region Site enumeration and selection
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
    #   1. getAllSites          -- complete inventory (needs admin-consented Sites.Read.All)
    #   2. Reports API + batch  -- complete inventory from getSharePointSiteUsageDetail,
    #                             resolving display names via JSON-batched site lookups
    #                             (needs Reports.Read.All; only when -IncludeStorageMetrics)
    #   3. Search API           -- partial results (fallback)
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
        # Path 1: getAllSites -- complete enumeration
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
        # Path 2: Reports API + JSON batching -- complete enumeration with storage.
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

        # Batch-resolve site details (20 per batch -- Graph maximum)
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
            Write-Host "  Tip: Use -IncludeStorageMetrics for complete site enumeration." -ForegroundColor DarkGray
        }
        $searchTerm = ($TenantId -split '\.')[0]
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
            $tempCsv = Join-Path ([System.IO.Path]::GetTempPath()) "spo_report_$(Get-Random).csv"
            Invoke-MgGraphRequest -Uri "v1.0/reports/getSharePointSiteUsageDetail(period='D7')" -OutputFilePath $tempCsv -ErrorAction Stop

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
                        $sampleUrl = ($reportData | Select-Object -First 1).'Site URL'
                        $useUrlMatch = [bool]$sampleUrl

                        if ($useUrlMatch) {
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
                            if (-not $Silent) { Write-Host "Site URLs are concealed. Matching by Site Id..." -ForegroundColor DarkGray }
                            $storageLookup = [System.Collections.Generic.Dictionary[string,object]]::new([System.StringComparer]::OrdinalIgnoreCase)
                            foreach ($row in $reportData) {
                                $siteId = $row.'Site Id'
                                if ($siteId) {
                                    $storageLookup[$siteId] = $row
                                }
                            }
                        }

                        $matchCount = 0
                        foreach ($site in $SiteList) {
                            if ($site.Type -eq 'OneDrive') { continue }

                            $reportRow = $null
                            if ($useUrlMatch) {
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

    # Step C: Present site picker (multi-select enabled)
    $pickerList = if ($IncludeStorageMetrics) {
        @($SiteList | Where-Object {
            if ($_.Type -eq 'OneDrive') { $true }
            else {
                $pathParts = ([uri]$_.WebUrl).AbsolutePath.Trim('/') -split '/' | Where-Object { $_ }
                $pathParts.Count -eq 0 -or
                ($pathParts.Count -eq 2 -and $pathParts[0] -in @('sites', 'teams'))
            }
        })
    } else {
        $SiteList
    }

    $gridColumns = if ($IncludeStorageMetrics) {
        @('Type', 'DisplayName', 'WebUrl', 'SiteCollGB', 'SiteCollFiles', 'Created')
    } else {
        @('Type', 'DisplayName', 'WebUrl', 'Created')
    }

    $sortedSiteList = if ($IncludeStorageMetrics) {
        $pickerList | Sort-Object { if ($_.SiteCollGB -is [double]) { $_.SiteCollGB } else { -1 } } -Descending
    } else {
        $pickerList
    }

    if (-not $Silent -and $pickerList.Count -ne $SiteList.Count) {
        Write-Host "Showing $($pickerList.Count) site collections (filtered $($SiteList.Count - $pickerList.Count) subsites)." -ForegroundColor DarkGray
    }

    $SelectedSites = @($sortedSiteList |
        Select-Object $gridColumns |
        Out-GridView -Title "Select one or more sites to scan ($($pickerList.Count) site collections) - use Ctrl+Click for multiple" -OutputMode Multiple)

    if (-not $SelectedSites -or $SelectedSites.Count -eq 0) {
        throw "No site selected. Exiting."
    }

    # Map back to full objects
    $SelectedSites = @(foreach ($sel in $SelectedSites) {
        $SiteList | Where-Object {
            $_.WebUrl -eq $sel.WebUrl -and $_.Type -eq $sel.Type
        } | Select-Object -First 1
    })

    if (-not $Silent) {
        if ($SelectedSites.Count -eq 1) {
            Write-Host "Selected: $($SelectedSites[0].DisplayName)" -ForegroundColor Green
        } else {
            Write-Host "Selected $($SelectedSites.Count) sites:" -ForegroundColor Green
            foreach ($s in $SelectedSites) { Write-Host "  - $($s.DisplayName)" -ForegroundColor DarkCyan }
        }
    }

    # Build report identifier
    if ($SelectedSites.Count -eq 1) {
        $reportIdentifier = if ($SelectedSites[0].Type -eq 'OneDrive') { $SelectedSites[0].Upn } else { $SelectedSites[0].DisplayName -replace '[\\/:*?"<>|]', '_' }
    } else {
        $first = $SelectedSites[0].DisplayName -replace '[\\/:*?"<>|]', '_'
        $reportIdentifier = "${first}_plus$($SelectedSites.Count - 1)sites"
    }

    # Hoist loop invariants
    $systemLibraryNames = [System.Collections.Generic.HashSet[string]]::new(
        [string[]]@('Pages', 'Images', 'Slideshow', 'Site Pages', 'Style Library',
            'Form Templates', 'FormServerTemplates', 'Site Collection Documents',
            'Site Collection Images', 'PublishingImages', 'SiteCollectionImages',
            '_catalogs', 'Converted Forms', 'appdata', 'appfiles',
            'Preservation Hold Library'),
        [System.StringComparer]::OrdinalIgnoreCase
    )
    $showPercent = $ResultSize -lt [int16]::MaxValue
    $progressId = 1
    $childSelect = '?$select=id,name,size,webUrl,lastModifiedDateTime,file,folder'

    # Internal function to recursively get files (defined once, called per site)
    function Get-FolderItemsRecursively {
        [CmdletBinding()]
        param(
            [string]$Path,
            [Collections.Generic.List[Object]]$AllFiles,
            [string]$FnBaseUri,
            [string]$FnDriveId,
            [string]$FnSiteName
        )

        if ($AllFiles.Count -ge $ResultSize) {
            $script:MoreFilesExist = $true
            return
        }

        if ($Path -eq "/") {
            $uri = "$FnBaseUri/children$childSelect"
        }
        else {
            $uri = "$FnBaseUri`:/$Path`:/children$childSelect"
        }

        do {
            if (-not $Silent) {
                $progressParams = @{
                    Id              = $progressId
                    Activity        = "Scanning files - $FnSiteName"
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
                    $script:AccessDeniedPaths.Add("$FnSiteName`:$Path")
                    return   # skip this folder
                }
                else {
                    Write-Warning "Failed to read '$Path' on '$FnSiteName': $($_.Exception.Message)"
                    return
                }
            }

            foreach ($Item in $Response.value) {
                if ($null -ne $Item.folder) {
                    $FolderPath = "$Path/$($Item.name)"
                    Get-FolderItemsRecursively -Path $FolderPath -AllFiles $AllFiles -FnBaseUri $FnBaseUri -FnDriveId $FnDriveId -FnSiteName $FnSiteName
                }
                elseif ($null -ne $Item.file) {
                    $AllFiles.Add(
                        [PSCustomObject]@{
                            Name                 = $Item.name
                            LastModifiedDateTime = $Item.lastModifiedDateTime
                            QuickXorHash         = $Item.file.hashes.quickXorHash
                            Size                 = $Item.size
                            WebUrl               = $Item.webUrl
                            ItemId               = $Item.id
                            DriveId              = $FnDriveId
                            SiteName             = $FnSiteName
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

    # Step D: Loop through selected sites -- resolve drives and collect files
    foreach ($CurrentSite in $SelectedSites) {
        $sitesScanned++
        $siteName = $CurrentSite.DisplayName

        if (-not $Silent -and $SelectedSites.Count -gt 1) {
            Write-Host "`nProcessing site $sitesScanned of $($SelectedSites.Count): $siteName" -ForegroundColor Green
        }

        if ($CurrentSite.Type -eq "OneDrive") {
            $DriveResponse = Invoke-MgGraphRequest -Uri "beta/users/$($CurrentSite.Upn)/drive"
            $DriveId = $DriveResponse.id
            $DriveWebUrl = $DriveResponse.webUrl
        }
        else {
            $drivesResponse = Invoke-MgGraphRequest -Uri "beta/sites/$($CurrentSite.SiteId)/drives?`$select=id,name,webUrl,quota"

            # Preservation Hold Library detection
            $preservationHoldDrive = $drivesResponse.value | Where-Object { $_.name -eq 'Preservation Hold Library' }
            if ($preservationHoldDrive) {
                $phlUsedBytes = 0
                $phlItemCount = $null

                if ($preservationHoldDrive.quota -and $preservationHoldDrive.quota.used) {
                    $phlUsedBytes = $preservationHoldDrive.quota.used
                }

                # Optional: get top-level child count
                try {
                    $phlRoot = Invoke-MgGraphRequest -Uri "beta/drives/$($preservationHoldDrive.id)/root?`$select=folder" -ErrorAction Stop
                    $phlItemCount = $phlRoot.folder.childCount
                }
                catch { }

                $PreservationHoldFindings.Add([PSCustomObject]@{
                    SiteName    = $siteName
                    SiteUrl     = $CurrentSite.WebUrl
                    LibraryName = 'Preservation Hold Library'
                    UsedGB      = [math]::Round($phlUsedBytes / 1GB, 2)
                    ItemCount   = $phlItemCount
                    DriveId     = $preservationHoldDrive.id
                })

                if (-not $Silent) {
                    Write-Host "  Preservation Hold Library detected: $([math]::Round($phlUsedBytes / 1GB, 2)) GB" -ForegroundColor Yellow
                }
            }

            # Filter out system/publishing libraries (HashSet lookup is O(1) vs O(n) for -notin)
            $drives = @($drivesResponse.value | Where-Object { -not $systemLibraryNames.Contains($_.name) })

            if (-not $drives -or $drives.Count -eq 0) {
                if (-not $Silent) { Write-Warning "No document libraries found for site '$siteName'. Skipping." }
                continue
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
                    Out-GridView -Title "Select a document library for '$siteName' ($($drives.Count) libraries)" -OutputMode Single

                if (-not $SelectedDrive) {
                    if (-not $Silent) { Write-Warning "No library selected for '$siteName'. Skipping." }
                    continue
                }

                $matchedDrive = $driveChoices | Where-Object { $_.WebUrl -eq $SelectedDrive.WebUrl } | Select-Object -First 1
                $DriveId = $matchedDrive.DriveId
                $DriveWebUrl = $matchedDrive.WebUrl
            }
        }

        if (-not $Silent) { Write-Host "  Drive: $DriveWebUrl" -ForegroundColor DarkCyan }

        $baseUri = "beta/drives/$DriveId/root"

        # File collection for this site
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
                        Activity        = "Scanning files - $siteName"
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
                        $script:AccessDeniedPaths.Add("$siteName`:$RootPath")
                        break
                    }
                    else {
                        Write-Warning "Failed to read '$RootPath' on '$siteName': $($_.Exception.Message)"
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

            if (-not $Silent) { Write-Progress -Id $progressId -Activity "Scanning files - $siteName" -Completed }

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
                    ItemId               = $DriveItem.id
                    DriveId              = $DriveId
                    SiteName             = $siteName
                })
            }
        }
        else {
            Get-FolderItemsRecursively -Path $RootPath -AllFiles $FileList -FnBaseUri $baseUri -FnDriveId $DriveId -FnSiteName $siteName
            if (-not $Silent) { Write-Progress -Id $progressId -Activity "Scanning files - $siteName" -Completed }
        }
    }

    # Access-denied diagnostic
    if ($script:AccessDeniedPaths.Count -gt 0 -and -not $Silent) {
        Write-Host ""
        Write-Warning "Access denied on $($script:AccessDeniedPaths.Count) path(s). Results may be incomplete."
        Write-Host "  Tip: Verify that admin consent has been granted for Sites.Read.All." -ForegroundColor DarkGray
        Write-Host ""
    }
    #endregion

    #region Version history analysis
    $VersionAnalysis = [Collections.Generic.List[Object]]::new()

    if ($VersionCheckCount -gt 0 -and $FileList.Count -gt 0) {
        $largestFiles = @($FileList |
            Sort-Object Size -Descending |
            Select-Object -First $VersionCheckCount)

        if (-not $Silent) { Write-Host "`nFetching version history for $($largestFiles.Count) largest files..." -ForegroundColor Cyan }

        # Build a lookup table for quick file reference by batch request id
        $fileLookup = @{}
        for ($idx = 0; $idx -lt $largestFiles.Count; $idx++) {
            $fileLookup["$idx"] = $largestFiles[$idx]
        }

        # Batch version requests (20 per batch -- Graph maximum)
        $batchSize = 20
        for ($i = 0; $i -lt $largestFiles.Count; $i += $batchSize) {
            $end = [math]::Min($i + $batchSize - 1, $largestFiles.Count - 1)

            if (-not $Silent) {
                $pct = [math]::Min(100, [int](($i + $batchSize) / $largestFiles.Count * 100))
                Write-Progress -Id 3 -Activity "Fetching version history" `
                    -Status "Batch $([math]::Floor($i / $batchSize) + 1) of $([math]::Ceiling($largestFiles.Count / $batchSize))" `
                    -PercentComplete $pct
            }

            $batchBody = @{
                requests = @(for ($j = $i; $j -le $end; $j++) {
                    $f = $largestFiles[$j]
                    @{
                        id     = "$j"
                        method = "GET"
                        url    = "/drives/$($f.DriveId)/items/$($f.ItemId)/versions?`$select=id,size,lastModifiedDateTime"
                    }
                })
            }

            try {
                $batchJson = $batchBody | ConvertTo-Json -Depth 4 -Compress
                $batchResponse = Invoke-MgGraphRequest -Uri "v1.0/`$batch" -Method POST -Body $batchJson -ContentType "application/json" -ErrorAction Stop

                foreach ($resp in $batchResponse.responses) {
                    if ($resp.status -ne 200) { continue }
                    $file = $fileLookup[$resp.id]
                    $versions = $resp.body.value
                    if (-not $versions -or $versions.Count -le 1) { continue }

                    # First version is the current version; skip it for "version storage"
                    $totalVersionBytes = [long]0
                    $allDates = [Collections.Generic.List[datetime]]::new()
                    for ($v = 1; $v -lt $versions.Count; $v++) {
                        $ver = $versions[$v]
                        if ($ver.size) { $totalVersionBytes += $ver.size }
                        if ($ver.lastModifiedDateTime) { $allDates.Add([datetime]$ver.lastModifiedDateTime) }
                    }
                    $versionCount = $versions.Count - 1

                    # Handle pagination -- files with many versions need individual follow-up calls
                    $nextLink = $resp.body.'@odata.nextLink'
                    while ($nextLink) {
                        try {
                            $pageResponse = Invoke-MgGraphRequest -Uri $nextLink -ErrorAction Stop
                            foreach ($ver in $pageResponse.value) {
                                if ($ver.size) { $totalVersionBytes += $ver.size }
                                if ($ver.lastModifiedDateTime) { $allDates.Add([datetime]$ver.lastModifiedDateTime) }
                            }
                            $versionCount += $pageResponse.value.Count
                            $nextLink = $pageResponse.'@odata.nextLink'
                        }
                        catch {
                            if (-not $Silent) { Write-Warning "Pagination failed for '$($file.Name)': $($_.Exception.Message)" }
                            break
                        }
                    }

                    # Add entry after all pages collected (avoids linear search to update)
                    $allDates.Sort()
                    $VersionAnalysis.Add([PSCustomObject]@{
                        SiteName             = $file.SiteName
                        FileName             = $file.Name
                        FileSizeMB           = [math]::Round($file.Size / 1MB, 2)
                        WebUrl               = $file.WebUrl
                        VersionCount         = $versionCount
                        TotalVersionSizeMB   = [math]::Round($totalVersionBytes / 1MB, 2)
                        CurrentPlusVersionMB = [math]::Round(($file.Size + $totalVersionBytes) / 1MB, 2)
                        OldestVersionDate    = if ($allDates.Count -gt 0) { $allDates[0].ToString('yyyy-MM-dd') } else { '' }
                        NewestVersionDate    = if ($allDates.Count -gt 0) { $allDates[$allDates.Count - 1].ToString('yyyy-MM-dd') } else { '' }
                        DriveId              = $file.DriveId
                        ItemId               = $file.ItemId
                    })
                }
            }
            catch {
                # Batch call failed -- fall back to individual calls for this chunk
                if (-not $Silent) { Write-Warning "Batch request failed, falling back to individual calls: $($_.Exception.Message)" }
                for ($j = $i; $j -le $end; $j++) {
                    $file = $largestFiles[$j]
                    try {
                        $versionsUri = "v1.0/drives/$($file.DriveId)/items/$($file.ItemId)/versions?`$select=id,size,lastModifiedDateTime"
                        $versions = [Collections.Generic.List[Object]]::new()
                        do {
                            $versionResponse = Invoke-MgGraphRequest -Uri $versionsUri -ErrorAction Stop
                            if ($versionResponse.value) { $versions.AddRange($versionResponse.value) }
                            $versionsUri = $versionResponse.'@odata.nextLink'
                        } until (-not $versionsUri)

                        if ($versions.Count -le 1) { continue }

                        $totalVersionBytes = 0
                        $allDates = [Collections.Generic.List[datetime]]::new()
                        for ($v = 1; $v -lt $versions.Count; $v++) {
                            $ver = $versions[$v]
                            if ($ver.size) { $totalVersionBytes += $ver.size }
                            if ($ver.lastModifiedDateTime) { $allDates.Add([datetime]$ver.lastModifiedDateTime) }
                        }
                        $allDates.Sort()

                        $VersionAnalysis.Add([PSCustomObject]@{
                            SiteName             = $file.SiteName
                            FileName             = $file.Name
                            FileSizeMB           = [math]::Round($file.Size / 1MB, 2)
                            WebUrl               = $file.WebUrl
                            VersionCount         = $versions.Count - 1
                            TotalVersionSizeMB   = [math]::Round($totalVersionBytes / 1MB, 2)
                            CurrentPlusVersionMB = [math]::Round(($file.Size + $totalVersionBytes) / 1MB, 2)
                            OldestVersionDate    = if ($allDates.Count -gt 0) { $allDates[0].ToString('yyyy-MM-dd') } else { '' }
                            NewestVersionDate    = if ($allDates.Count -gt 0) { $allDates[$allDates.Count - 1].ToString('yyyy-MM-dd') } else { '' }
                            DriveId              = $file.DriveId
                            ItemId               = $file.ItemId
                        })
                    }
                    catch {
                        if (-not $Silent) { Write-Warning "Could not fetch versions for '$($file.Name)': $($_.Exception.Message)" }
                    }
                }
            }
        }

        if (-not $Silent) { Write-Progress -Id 3 -Activity "Fetching version history" -Completed }
    }
    #endregion

    #region Duplicate detection (dual confidence)
    $Output = [Collections.Generic.List[Object]]::new()

    if (-not $SkipDuplicates) {
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

        foreach ($Group in $hashBasedDupes) {
            $fileGroupSize = ($Group.Group.Size | Measure-Object -Sum).Sum
            $singleFileSize = $Group.Group[0].Size
            $Output.Add([PSCustomObject]@{
                Confidence      = "High"
                MatchType       = "QuickXorHash"
                MatchValue      = $Group.Name
                NumberOfFiles   = $Group.Count
                FileSizeKB      = [math]::Round($singleFileSize / 1KB, 2)
                FileGroupSizeKB = [math]::Round($fileGroupSize / 1KB, 2)
                PossibleWasteKB = [math]::Round(($fileGroupSize - $singleFileSize) / 1KB, 2)
                FileNames       = ($Group.Group.Name | Sort-Object -Unique) -join ";"
                WebLocalPaths   = ($Group.Group.WebUrl | ForEach-Object { ([uri]$_).LocalPath }) -join ";"
                SiteName        = ($Group.Group.SiteName | Sort-Object -Unique) -join ";"
            })
        }

        foreach ($Group in $nameBasedDupes) {
            $fileGroupSize = ($Group.Group.Size | Measure-Object -Sum).Sum
            $singleFileSize = $Group.Group[0].Size
            $Output.Add([PSCustomObject]@{
                Confidence      = "Low"
                MatchType       = "FileName"
                MatchValue      = $Group.Name
                NumberOfFiles   = $Group.Count
                FileSizeKB      = [math]::Round($singleFileSize / 1KB, 2)
                FileGroupSizeKB = [math]::Round($fileGroupSize / 1KB, 2)
                PossibleWasteKB = [math]::Round(($fileGroupSize - $singleFileSize) / 1KB, 2)
                FileNames       = $Group.Name
                WebLocalPaths   = ($Group.Group.WebUrl | ForEach-Object { ([uri]$_).LocalPath }) -join ";"
                SiteName        = ($Group.Group.SiteName | Sort-Object -Unique) -join ";"
            })
        }
    }
    #endregion

    #region Pipeline output
    if ($OutputStyle -eq "PassThru" -or $OutputStyle -eq "ReportAndPassThru") {
        [PSCustomObject]@{
            Duplicates        = $Output
            VersionAnalysis   = $VersionAnalysis
            PreservationHolds = $PreservationHoldFindings
        }
    }
    #endregion

    #region Report generation (HTML, CSV, JSON)
    if ($OutputStyle -eq "Report" -or $OutputStyle -eq "ReportAndPassThru") {
        $fileDate = Get-Date -Format 'yyyyMMdd_HHmmss'
        $desktop = [Environment]::GetFolderPath("Desktop")
        $folderName = "SPOInsights_$($reportIdentifier -replace '[\\/:*?"<>|@]', '_')_$fileDate"
        $outputFolder = Join-Path $desktop $folderName
        $null = New-Item -ItemType Directory -Path $outputFolder -Force

        $dupeCsvPath  = Join-Path $outputFolder "DuplicateReport.csv"
        $dupeJsonPath = Join-Path $outputFolder "DuplicateReport.json"
        $verCsvPath   = Join-Path $outputFolder "VersionReport.csv"
        $verJsonPath  = Join-Path $outputFolder "VersionReport.json"
        $phlCsvPath   = Join-Path $outputFolder "PreservationHoldReport.csv"
        $phlJsonPath  = Join-Path $outputFolder "PreservationHoldReport.json"
        $htmlPath     = Join-Path $outputFolder "StorageInsightsReport.html"

        # Calculate metrics
        $totalBytesEvaluated = ($FileList.Size | Measure-Object -Sum).Sum
        $totalWasteBytes = ($Output.PossibleWasteKB | Measure-Object -Sum).Sum * 1KB
        $totalVersionStorageMB = ($VersionAnalysis.TotalVersionSizeMB | Measure-Object -Sum).Sum
        $totalPhlGB = ($PreservationHoldFindings.UsedGB | Measure-Object -Sum).Sum

        $topWaste = $Output |
            Sort-Object { $_.FileSizeKB * $_.NumberOfFiles } -Descending |
            Select-Object -First 10

        $topVersions = $VersionAnalysis |
            Sort-Object TotalVersionSizeMB -Descending |
            Select-Object -First 20

        # Build navigation links
        $navLinks = [Collections.Generic.List[string]]::new()
        if (-not $SkipDuplicates) { $navLinks.Add("<a href='#duplicates' class='btn btn-nav'>Duplicates</a>") }
        if ($VersionCheckCount -gt 0) { $navLinks.Add("<a href='#versions' class='btn btn-nav'>Version History</a>") }
        $navLinks.Add("<a href='#preservation' class='btn btn-nav'>Preservation Holds</a>")

        $htmlContent = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>SharePoint / OneDrive Storage Insights Report</title>
    <style>
        * { box-sizing: border-box; }
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            margin: 0; padding: 20px;
            background: #1a1a2e; color: #eee;
        }
        .container { max-width: 1400px; margin: 0 auto; }
        h1 { color: #00d4ff; margin-bottom: 5px; }
        h2 { margin-top: 40px; margin-bottom: 15px; padding-top: 10px; }
        .subtitle { color: #888; margin-bottom: 30px; }
        .metrics {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
            gap: 15px; margin-bottom: 30px;
        }
        .metric {
            background: #16213e; padding: 20px; border-radius: 8px;
            border-left: 4px solid #00d4ff;
        }
        .metric-value { font-size: 28px; font-weight: bold; color: #00d4ff; }
        .metric-label { color: #888; font-size: 14px; margin-top: 5px; }
        .metric-version { border-left-color: #a855f7; }
        .metric-version .metric-value { color: #a855f7; }
        .metric-hold { border-left-color: #f97316; }
        .metric-hold .metric-value { color: #f97316; }
        .buttons { margin-bottom: 20px; }
        .btn {
            display: inline-block; padding: 10px 20px; margin-right: 8px; margin-bottom: 8px;
            background: #00d4ff; color: #1a1a2e; text-decoration: none;
            border-radius: 6px; font-weight: bold; cursor: pointer; border: none;
            font-size: 13px;
        }
        .btn:hover { background: #00a8cc; }
        .btn-secondary { background: #4a4a6a; color: #eee; }
        .btn-secondary:hover { background: #5a5a7a; }
        .btn-nav { background: #2a2a4a; color: #eee; font-size: 12px; padding: 8px 16px; }
        .btn-nav:hover { background: #3a3a5a; }
        .btn-version { background: #a855f7; }
        .btn-version:hover { background: #9333ea; }
        .btn-hold { background: #f97316; color: #1a1a2e; }
        .btn-hold:hover { background: #ea580c; }
        table { width: 100%; border-collapse: collapse; margin-top: 20px; }
        th, td { padding: 12px; text-align: left; border-bottom: 1px solid #2a2a4a; }
        th { background: #16213e; color: #00d4ff; position: sticky; top: 0; }
        tr:hover { background: #1f1f3a; }
        .confidence-high { color: #4ade80; }
        .confidence-low { color: #fbbf24; }
        .truncate { max-width: 300px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
        .section-title { color: #00d4ff; margin-top: 40px; margin-bottom: 15px; }
        .section-title-version { color: #a855f7; }
        .section-title-hold { color: #f97316; }
        .empty-state { color: #666; font-style: italic; padding: 20px; }
    </style>
</head>
<body>
    <div class="container">
        <h1>SharePoint / OneDrive Storage Insights Report</h1>
        <p class="subtitle">
            <strong>Source:</strong> $reportIdentifier |
            <strong>Sites Scanned:</strong> $sitesScanned |
            <strong>Path:</strong> $(if ($RootPath -and $RootPath -ne '/') { $RootPath } else { '/' }) |
            <strong>Generated:</strong> $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        </p>

        <div class="buttons">
            $($navLinks -join "`n            ")
        </div>

        <div class="metrics">
            <div class="metric">
                <div class="metric-value">$sitesScanned</div>
                <div class="metric-label">Sites Scanned</div>
            </div>
            <div class="metric">
                <div class="metric-value">$($FileList.Count.ToString('N0'))</div>
                <div class="metric-label">Files Evaluated</div>
            </div>
            <div class="metric">
                <div class="metric-value">$([math]::Round($totalBytesEvaluated / 1GB, 2)) GB</div>
                <div class="metric-label">Total Size Evaluated</div>
            </div>
$(if (-not $SkipDuplicates) { @"
            <div class="metric">
                <div class="metric-value">$($Output.Count)</div>
                <div class="metric-label">Duplicate Groups</div>
            </div>
            <div class="metric">
                <div class="metric-value">$([math]::Round($totalWasteBytes / 1MB, 2)) MB</div>
                <div class="metric-label">Duplicate Waste</div>
            </div>
"@ })
$(if ($VersionCheckCount -gt 0) { @"
            <div class="metric metric-version">
                <div class="metric-value">$($VersionAnalysis.Count)</div>
                <div class="metric-label">Files with Versions</div>
            </div>
            <div class="metric metric-version">
                <div class="metric-value">$([math]::Round($totalVersionStorageMB / 1024, 2)) GB</div>
                <div class="metric-label">Version Storage</div>
            </div>
"@ })
            <div class="metric metric-hold">
                <div class="metric-value">$($PreservationHoldFindings.Count)</div>
                <div class="metric-label">Preservation Hold Libraries</div>
            </div>
$(if ($PreservationHoldFindings.Count -gt 0) { @"
            <div class="metric metric-hold">
                <div class="metric-value">$([math]::Round($totalPhlGB, 2)) GB</div>
                <div class="metric-label">Preservation Hold Size</div>
            </div>
"@ })
        </div>

        <div class="buttons">
$(if (-not $SkipDuplicates) { @"
            <a href="DuplicateReport.csv" class="btn">Duplicate CSV</a>
            <a href="DuplicateReport.json" class="btn btn-secondary">Duplicate JSON</a>
"@ })
$(if ($VersionCheckCount -gt 0) { @"
            <a href="VersionReport.csv" class="btn btn-version">Version CSV</a>
            <a href="VersionReport.json" class="btn btn-secondary">Version JSON</a>
"@ })
$(if ($PreservationHoldFindings.Count -gt 0) { @"
            <a href="PreservationHoldReport.csv" class="btn btn-hold">Preservation Hold CSV</a>
            <a href="PreservationHoldReport.json" class="btn btn-secondary">Preservation Hold JSON</a>
"@ })
        </div>

$(if (-not $SkipDuplicates) { @"
        <h2 id="duplicates" class="section-title">Top 10 Duplicates by Wasted Space</h2>
$(if ($Output.Count -gt 0) { @"
        <table>
            <thead>
                <tr>
                    <th>Confidence</th>
                    <th>Match Type</th>
                    <th>Files</th>
                    <th>File Size</th>
                    <th>Total Waste</th>
                    <th>File Names</th>
                    <th>Site</th>
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
                    <td class='truncate' title='$($item.SiteName)'>$($item.SiteName)</td>
                </tr>"
})
            </tbody>
        </table>
"@ } else { "        <p class='empty-state'>No duplicate files found.</p>" })
"@ })

$(if ($VersionCheckCount -gt 0) { @"
        <h2 id="versions" class="section-title section-title-version">Version History Analysis (Top $($VersionCheckCount) Largest Files)</h2>
$(if ($VersionAnalysis.Count -gt 0) { @"
        <table>
            <thead>
                <tr>
                    <th>Site</th>
                    <th>File Name</th>
                    <th>Current Size (MB)</th>
                    <th>Versions</th>
                    <th>Version Storage (MB)</th>
                    <th>Total (MB)</th>
                    <th>Oldest Version</th>
                    <th>Newest Version</th>
                </tr>
            </thead>
            <tbody>
$(foreach ($v in $topVersions) {
    "                <tr>
                    <td class='truncate' title='$($v.SiteName)'>$($v.SiteName)</td>
                    <td class='truncate' title='$($v.FileName)'>$($v.FileName)</td>
                    <td>$($v.FileSizeMB)</td>
                    <td>$($v.VersionCount)</td>
                    <td>$($v.TotalVersionSizeMB)</td>
                    <td>$($v.CurrentPlusVersionMB)</td>
                    <td>$($v.OldestVersionDate)</td>
                    <td>$($v.NewestVersionDate)</td>
                </tr>"
})
            </tbody>
        </table>
"@ } else { "        <p class='empty-state'>No version history found in the $($VersionCheckCount) largest files checked.</p>" })
"@ })

        <h2 id="preservation" class="section-title section-title-hold">Preservation Hold Libraries</h2>
$(if ($PreservationHoldFindings.Count -gt 0) { @"
        <table>
            <thead>
                <tr>
                    <th>Site Name</th>
                    <th>Site URL</th>
                    <th>Size (GB)</th>
                    <th>Item Count</th>
                </tr>
            </thead>
            <tbody>
$(foreach ($phl in $PreservationHoldFindings) {
    "                <tr>
                    <td>$($phl.SiteName)</td>
                    <td class='truncate' title='$($phl.SiteUrl)'>$($phl.SiteUrl)</td>
                    <td>$($phl.UsedGB)</td>
                    <td>$(if ($null -ne $phl.ItemCount) { $phl.ItemCount } else { 'N/A' })</td>
                </tr>"
})
            </tbody>
        </table>
"@ } else { @"
        <p class='empty-state'>No Preservation Hold Libraries detected.</p>
"@ })

        $(if ($script:MoreFilesExist) {
            "<p style='color: #fbbf24; margin-top: 30px;'><strong>Warning:</strong> ResultSize limit ($ResultSize) reached on one or more sites. More files exist but were not evaluated. Use -ResultSize to increase.</p>"
        })
        $(if ($script:AccessDeniedPaths.Count -gt 0) {
            "<p style='color: #ff6b6b; margin-top: 30px;'><strong>Access Denied:</strong> $($script:AccessDeniedPaths.Count) path(s) could not be read. Verify that admin consent has been granted for Sites.Read.All.</p>"
        })
    </div>
</body>
</html>
"@

        try {
            # Duplicate reports
            if (-not $SkipDuplicates) {
                if ($Output.Count -gt 0) {
                    $Output | Export-Csv $dupeCsvPath -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
                    $Output | ConvertTo-Json -Depth 3 | Out-File $dupeJsonPath -Encoding UTF8 -ErrorAction Stop
                }
                else {
                    'Confidence,MatchType,MatchValue,NumberOfFiles,FileSizeKB,FileGroupSizeKB,PossibleWasteKB,FileNames,WebLocalPaths,SiteName' |
                        Out-File $dupeCsvPath -Encoding UTF8 -ErrorAction Stop
                    '[]' | Out-File $dupeJsonPath -Encoding UTF8 -ErrorAction Stop
                }
            }

            # Version reports
            if ($VersionCheckCount -gt 0) {
                if ($VersionAnalysis.Count -gt 0) {
                    $VersionAnalysis | Export-Csv $verCsvPath -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
                    $VersionAnalysis | ConvertTo-Json -Depth 3 | Out-File $verJsonPath -Encoding UTF8 -ErrorAction Stop
                }
                else {
                    'SiteName,FileName,FileSizeMB,WebUrl,VersionCount,TotalVersionSizeMB,CurrentPlusVersionMB,OldestVersionDate,NewestVersionDate,DriveId,ItemId' |
                        Out-File $verCsvPath -Encoding UTF8 -ErrorAction Stop
                    '[]' | Out-File $verJsonPath -Encoding UTF8 -ErrorAction Stop
                }
            }

            # Preservation hold reports
            if ($PreservationHoldFindings.Count -gt 0) {
                $PreservationHoldFindings | Export-Csv $phlCsvPath -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
                $PreservationHoldFindings | ConvertTo-Json -Depth 3 | Out-File $phlJsonPath -Encoding UTF8 -ErrorAction Stop
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
        Write-Host "$($FileList.Count) files scanned across $sitesScanned site(s) in ${duration}s." -ForegroundColor White

        if (-not $SkipDuplicates) {
            Write-Host "  Duplicates: " -NoNewline
            Write-Host "$($Output.Count) duplicate groups, $([math]::Round($totalWasteBytes / 1MB, 2)) MB potential waste." -ForegroundColor Cyan
        }
        if ($VersionCheckCount -gt 0) {
            Write-Host "  Versions: " -NoNewline
            Write-Host "$($largestFiles.Count) files checked, $($VersionAnalysis.Count) with history, $([math]::Round($totalVersionStorageMB, 2)) MB in version storage." -ForegroundColor Magenta
        }
        Write-Host "  Preservation Hold: " -NoNewline
        Write-Host "$($PreservationHoldFindings.Count) libraries found$(if ($PreservationHoldFindings.Count -gt 0) { ", $([math]::Round($totalPhlGB, 2)) GB total" })." -ForegroundColor Yellow
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
    Get-SPOStorageInsights @scriptParams
}
