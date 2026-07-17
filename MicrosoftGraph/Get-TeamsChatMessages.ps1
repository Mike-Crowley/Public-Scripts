#Requires -Modules Microsoft.Graph.Authentication

function Get-TeamsChatMessages {
    <#
    .SYNOPSIS
        Retrieves and processes Microsoft Teams chat messages using Microsoft Graph API.

    .DESCRIPTION
        Connects to Microsoft Graph and fetches chat messages within a date range using
        pagination, with automatic retry on throttling (429) and transient server errors.
        Extracts message properties including reactions, replies, attachments, and body
        content (converted from HTML to plain text).

        Chats are listed most-recently-active first (server-side ordering by last message)
        and include a preview of the most recent message. The picker supports selecting
        one or more chats (Out-GridView where available, with a numbered console menu
        fallback on platforms without it). When multiple chats are selected, their first
        pages of messages are requested in a single JSON batch request. Use -ChatId to
        skip the picker entirely for scripted or headless runs.

        By default, uses delegated authentication (interactive sign-in with Chat.Read scope)
        and queries the signed-in user's own chats via the /me endpoint. The default date
        range mode is Relative, specified with -DaysAgo.

        Use -StartDate/-EndDate for a specific date range, or -WindowStart/-WindowEnd for a
        sliding window. These date parameters are mutually exclusive.

        Use the -ClientId, -TenantId, and -UserId parameters to authenticate with
        application permissions (Chat.Read.All) instead, which allows querying any user's
        chats. Run with -RegisterApp -TenantId first to create the required app registration.

        System event messages (members added, call events, etc.) and deleted messages are
        excluded from output by default; use -IncludeSystemMessages to include the events.

    .PARAMETER RegisterApp
        Creates an Entra ID app registration with the Chat.Read.All application permission,
        a client secret, service principal, and grants admin consent. Run this once before
        using app auth. If an app named Get-TeamsChatMessages already exists, it is reused
        and a new secret is added rather than creating a duplicate registration. Requires
        Application.ReadWrite.All and AppRoleAssignment.ReadWrite.All permissions
        (typically Application Administrator or Global Administrator).

    .PARAMETER ClientId
        The Application (client) ID of the Entra ID app registration. Required for app auth.

    .PARAMETER ClientSecret
        The client secret for the app registration. If omitted, you will be prompted via
        Get-Credential. Accepts the plain-text secret returned by -RegisterApp, e.g.
        -ClientSecret $app.ClientSecret.

    .PARAMETER TenantId
        The tenant ID or domain name for the Microsoft 365 tenant. Required for -RegisterApp
        and app auth.

    .PARAMETER UserId
        The user ID (GUID) or UPN of the user whose chats to retrieve. Required for app auth.

    .PARAMETER UseDeviceCode
        If specified, uses device code flow for delegated authentication instead of
        interactive browser sign-in. Useful for remote sessions or environments without
        a browser.

    .PARAMETER DaysAgo
        Number of days back from now. This is the default date range mode.

    .PARAMETER StartDate
        Start date for the query range. Accepts any format PowerShell can parse as a
        DateTime (e.g. "2025-06-01", "June 1, 2025"). Must be used with -EndDate.

    .PARAMETER EndDate
        End date for the query range (inclusive). Must be used with -StartDate.

    .PARAMETER WindowStart
        Start of window as days ago. Must be used with -WindowEnd.

    .PARAMETER WindowEnd
        End of window as days ago (must be less than WindowStart). Must be used with -WindowStart.

    .PARAMETER ChatId
        One or more chat IDs (e.g. '19:...@thread.v2') to retrieve messages from, bypassing
        the interactive chat picker. Chat IDs appear in the picker's ChatId column and in
        the objects this function returns. Required for fully unattended runs.

    .PARAMETER ChatCount
        Number of most recent chats to retrieve for the picker. Default is 100.
        Use 0 or "Unlimited" to retrieve all chats. Ignored when -ChatId is specified.

    .PARAMETER PageSize
        Number of messages per Graph API page. The chat messages endpoint allows a
        maximum of 50, which is the default.

    .PARAMETER IncludeSystemMessages
        Includes system event messages (members added/removed, call started, etc.) in the
        output. These are excluded by default. Deleted messages are always excluded and
        reported in the summary.

    .PARAMETER ExportFormat
        Output format for non-interactive mode: JSON, CSV, HTML, or None. HTML produces a
        readable chat transcript. When specified, skips the interactive output picker and
        uses this format directly.

    .PARAMETER ExportToFile
        Saves output to a file instead of copying to clipboard. Used with -ExportFormat.

    .PARAMETER ExportPath
        File path prefix when using -ExportToFile. A timestamp and extension are appended.
        Default is '.\TeamsChat_'.

    .PARAMETER ShowInConsole
        Displays a sample of recent messages in the console.

    .PARAMETER ShowLastNMessages
        Number of recent messages to display in console with -ShowInConsole. Default is 10.

    .PARAMETER ShowFullBody
        Shows the full message body instead of truncating.

    .EXAMPLE
        Get-TeamsChatMessages -RegisterApp -TenantId "contoso.com"

        Creates (or reuses) an Entra ID app registration with Chat.Read.All permission in
        the specified tenant and outputs the Client ID and secret. An admin must grant
        consent before the app can be used if automatic consent fails.

    .EXAMPLE
        Get-TeamsChatMessages -DaysAgo 7

        Signs in interactively and retrieves your own Teams chat messages from the last 7 days.

    .EXAMPLE
        Get-TeamsChatMessages -DaysAgo 7 -UseDeviceCode

        Uses device code flow (prints a code to enter at https://microsoft.com/devicelogin)
        instead of opening a browser window. Useful for remote/SSH sessions.

    .EXAMPLE
        Get-TeamsChatMessages -StartDate "2025-06-01" -EndDate "2025-06-15"

        Retrieves your messages from June 1-15 and prompts for output options.

    .EXAMPLE
        Get-TeamsChatMessages -WindowStart 14 -WindowEnd 7

        Retrieves your messages from 14 to 7 days ago and prompts for output options.

    .EXAMPLE
        $app = Get-TeamsChatMessages -RegisterApp -TenantId "contoso.com" -UseDeviceCode
        Get-TeamsChatMessages -ClientId $app.ClientId -ClientSecret $app.ClientSecret -TenantId "contoso.com" -UserId "user@contoso.com" -DaysAgo 7

        Registers an app, then uses it to retrieve another user's messages from the last 7 days.
        The -UserId parameter also accepts a user's object ID (GUID). If -ClientSecret is omitted,
        you will be prompted for the secret via Get-Credential.

    .EXAMPLE
        Get-TeamsChatMessages -DaysAgo 1 -ChatId "19:2da4c29f6d7041eca70b638b43d45437@thread.v2" -ExportFormat CSV -ExportToFile -ExportPath "C:\Reports\TeamsChat_"

        Skips the chat picker and exports yesterday's messages from a specific chat to CSV.
        Combined with app auth, this form is suitable for scheduled tasks.

    .EXAMPLE
        Get-TeamsChatMessages -DaysAgo 30 -ExportFormat HTML -ExportToFile

        Saves the last 30 days of the selected chat as a readable HTML transcript.

    .NOTES
        Author: Mike Crowley
        https://mikecrowley.us

        Requires the Microsoft.Graph.Authentication module.
        Delegated auth requires the Chat.Read scope.
        App auth requires Chat.Read.All application permissions on the app registration.

        Throttled requests (429) and transient server errors are retried automatically
        with backoff, honoring the Retry-After header. When multiple chats are selected,
        first pages are fetched with a single JSON batch ($batch) request.

        TIP: To read another user's chats, you need an app registration with Chat.Read.All.
             Run 'Get-TeamsChatMessages -RegisterApp -TenantId <tenant>' to create one, then
             have a tenant admin grant consent in the Azure portal before using
             -ClientId/-TenantId/-UserId.

    .LINK
        https://github.com/Mike-Crowley/Public-Scripts
    #>

    [CmdletBinding(DefaultParameterSetName = "DelegatedRelative")]
    param(
        [Parameter(Mandatory, ParameterSetName = "RegisterApp")]
        [switch]$RegisterApp,

        # App auth parameters
        [Parameter(Mandatory, ParameterSetName = "AppAuthRelative")]
        [Parameter(Mandatory, ParameterSetName = "AppAuthSpecific")]
        [Parameter(Mandatory, ParameterSetName = "AppAuthWindow")]
        [ValidateNotNullOrEmpty()]
        [string]$ClientId,

        [Parameter(ParameterSetName = "AppAuthRelative")]
        [Parameter(ParameterSetName = "AppAuthSpecific")]
        [Parameter(ParameterSetName = "AppAuthWindow")]
        [string]$ClientSecret,

        [Parameter(Mandatory, ParameterSetName = "RegisterApp")]
        [Parameter(Mandatory, ParameterSetName = "AppAuthRelative")]
        [Parameter(Mandatory, ParameterSetName = "AppAuthSpecific")]
        [Parameter(Mandatory, ParameterSetName = "AppAuthWindow")]
        [ValidateNotNullOrEmpty()]
        [string]$TenantId,

        [Parameter(Mandatory, ParameterSetName = "AppAuthRelative")]
        [Parameter(Mandatory, ParameterSetName = "AppAuthSpecific")]
        [Parameter(Mandatory, ParameterSetName = "AppAuthWindow")]
        [ValidateNotNullOrEmpty()]
        [string]$UserId,

        # Device code flow for delegated and RegisterApp auth
        [Parameter(ParameterSetName = "RegisterApp")]
        [Parameter(ParameterSetName = "DelegatedRelative")]
        [Parameter(ParameterSetName = "DelegatedSpecific")]
        [Parameter(ParameterSetName = "DelegatedWindow")]
        [switch]$UseDeviceCode,

        # Relative date range (default)
        [Parameter(Mandatory, ParameterSetName = "DelegatedRelative")]
        [Parameter(Mandatory, ParameterSetName = "AppAuthRelative")]
        [ValidateRange(1, 36500)]
        [int]$DaysAgo,

        # Specific date range
        [Parameter(Mandatory, ParameterSetName = "DelegatedSpecific")]
        [Parameter(Mandatory, ParameterSetName = "AppAuthSpecific")]
        [datetime]$StartDate,

        [Parameter(Mandatory, ParameterSetName = "DelegatedSpecific")]
        [Parameter(Mandatory, ParameterSetName = "AppAuthSpecific")]
        [datetime]$EndDate,

        # Window date range
        [Parameter(Mandatory, ParameterSetName = "DelegatedWindow")]
        [Parameter(Mandatory, ParameterSetName = "AppAuthWindow")]
        [ValidateRange(1, 36500)]
        [int]$WindowStart,

        [Parameter(Mandatory, ParameterSetName = "DelegatedWindow")]
        [Parameter(Mandatory, ParameterSetName = "AppAuthWindow")]
        [ValidateRange(0, 36500)]
        [int]$WindowEnd,

        # Chat targeting (skips the interactive picker)
        [Parameter(ParameterSetName = "DelegatedRelative")]
        [Parameter(ParameterSetName = "DelegatedSpecific")]
        [Parameter(ParameterSetName = "DelegatedWindow")]
        [Parameter(ParameterSetName = "AppAuthRelative")]
        [Parameter(ParameterSetName = "AppAuthSpecific")]
        [Parameter(ParameterSetName = "AppAuthWindow")]
        [ValidateNotNullOrEmpty()]
        [string[]]$ChatId,

        # Shared query/export parameters
        [Parameter(ParameterSetName = "DelegatedRelative")]
        [Parameter(ParameterSetName = "DelegatedSpecific")]
        [Parameter(ParameterSetName = "DelegatedWindow")]
        [Parameter(ParameterSetName = "AppAuthRelative")]
        [Parameter(ParameterSetName = "AppAuthSpecific")]
        [Parameter(ParameterSetName = "AppAuthWindow")]
        [ValidateScript({ $_ -eq "Unlimited" -or $null -ne ($_ -as [int]) })]
        [string]$ChatCount = "100",

        [Parameter(ParameterSetName = "DelegatedRelative")]
        [Parameter(ParameterSetName = "DelegatedSpecific")]
        [Parameter(ParameterSetName = "DelegatedWindow")]
        [Parameter(ParameterSetName = "AppAuthRelative")]
        [Parameter(ParameterSetName = "AppAuthSpecific")]
        [Parameter(ParameterSetName = "AppAuthWindow")]
        [ValidateRange(1, 50)]
        [int]$PageSize = 50,

        [Parameter(ParameterSetName = "DelegatedRelative")]
        [Parameter(ParameterSetName = "DelegatedSpecific")]
        [Parameter(ParameterSetName = "DelegatedWindow")]
        [Parameter(ParameterSetName = "AppAuthRelative")]
        [Parameter(ParameterSetName = "AppAuthSpecific")]
        [Parameter(ParameterSetName = "AppAuthWindow")]
        [switch]$IncludeSystemMessages,

        # Non-interactive export parameters (skip the output picker when specified)
        [Parameter(ParameterSetName = "DelegatedRelative")]
        [Parameter(ParameterSetName = "DelegatedSpecific")]
        [Parameter(ParameterSetName = "DelegatedWindow")]
        [Parameter(ParameterSetName = "AppAuthRelative")]
        [Parameter(ParameterSetName = "AppAuthSpecific")]
        [Parameter(ParameterSetName = "AppAuthWindow")]
        [ValidateSet("JSON", "CSV", "HTML", "None")]
        [string]$ExportFormat,

        [Parameter(ParameterSetName = "DelegatedRelative")]
        [Parameter(ParameterSetName = "DelegatedSpecific")]
        [Parameter(ParameterSetName = "DelegatedWindow")]
        [Parameter(ParameterSetName = "AppAuthRelative")]
        [Parameter(ParameterSetName = "AppAuthSpecific")]
        [Parameter(ParameterSetName = "AppAuthWindow")]
        [switch]$ExportToFile,

        [Parameter(ParameterSetName = "DelegatedRelative")]
        [Parameter(ParameterSetName = "DelegatedSpecific")]
        [Parameter(ParameterSetName = "DelegatedWindow")]
        [Parameter(ParameterSetName = "AppAuthRelative")]
        [Parameter(ParameterSetName = "AppAuthSpecific")]
        [Parameter(ParameterSetName = "AppAuthWindow")]
        [string]$ExportPath = ".\TeamsChat_",

        [Parameter(ParameterSetName = "DelegatedRelative")]
        [Parameter(ParameterSetName = "DelegatedSpecific")]
        [Parameter(ParameterSetName = "DelegatedWindow")]
        [Parameter(ParameterSetName = "AppAuthRelative")]
        [Parameter(ParameterSetName = "AppAuthSpecific")]
        [Parameter(ParameterSetName = "AppAuthWindow")]
        [switch]$ShowInConsole,

        [Parameter(ParameterSetName = "DelegatedRelative")]
        [Parameter(ParameterSetName = "DelegatedSpecific")]
        [Parameter(ParameterSetName = "DelegatedWindow")]
        [Parameter(ParameterSetName = "AppAuthRelative")]
        [Parameter(ParameterSetName = "AppAuthSpecific")]
        [Parameter(ParameterSetName = "AppAuthWindow")]
        [ValidateRange(1, 2147483647)]
        [int]$ShowLastNMessages = 10,

        [Parameter(ParameterSetName = "DelegatedRelative")]
        [Parameter(ParameterSetName = "DelegatedSpecific")]
        [Parameter(ParameterSetName = "DelegatedWindow")]
        [Parameter(ParameterSetName = "AppAuthRelative")]
        [Parameter(ParameterSetName = "AppAuthSpecific")]
        [Parameter(ParameterSetName = "AppAuthWindow")]
        [switch]$ShowFullBody
    )

    # Determine auth and date mode from parameter set name
    $IsAppAuth = $PSCmdlet.ParameterSetName -like "AppAuth*"
    $DateMode = ($PSCmdlet.ParameterSetName -replace '^(Delegated|AppAuth)', '')

    # Validate date combinations before any network calls
    if ($DateMode -eq "Specific" -and $EndDate -lt $StartDate) {
        Write-Error "EndDate ($($EndDate.ToShortDateString())) must be on or after StartDate ($($StartDate.ToShortDateString()))."
        return
    }
    if ($DateMode -eq "Window" -and $WindowEnd -ge $WindowStart) {
        Write-Error "WindowEnd ($WindowEnd) must be less than WindowStart ($WindowStart). Example: -WindowStart 14 -WindowEnd 7"
        return
    }

    # Clear any existing Graph session to avoid MSAL cache conflicts between auth methods
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null

    #region Helpers
    # Culture-invariant UTC formatter. ':' is a culture-sensitive placeholder in ToString(),
    # which can produce invalid OData datetimes on systems with non-standard time separators.
    function Format-GraphDate {
        param([datetime]$Date)
        $Date.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffZ", [System.Globalization.CultureInfo]::InvariantCulture)
    }

    # Wrapper around Invoke-MgGraphRequest that retries throttled (429) and transient
    # server errors (5xx) with backoff, honoring the Retry-After header when present.
    function Invoke-GraphRequestWithRetry {
        param(
            [string]$Uri,
            [string]$Method = "GET",
            $Body,
            [int]$MaxRetries = 5
        )
        $Attempt = 0
        while ($true) {
            try {
                $RequestParams = @{
                    Uri         = $Uri
                    Method      = $Method
                    OutputType  = "PSObject"
                    ErrorAction = "Stop"
                }
                if ($Body) {
                    $RequestParams.Body = $Body
                    $RequestParams.ContentType = "application/json"
                }
                return Invoke-MgGraphRequest @RequestParams
            }
            catch {
                $Attempt++
                $StatusCode = $null
                try { $StatusCode = [int]$_.Exception.Response.StatusCode } catch { }
                if (-not $StatusCode -and $_.Exception.Message -match 'TooManyRequests|throttl') { $StatusCode = 429 }
                $Retryable = $StatusCode -in 429, 500, 502, 503, 504
                if (-not $Retryable -or $Attempt -gt $MaxRetries) { throw }

                $Delay = [int][math]::Min(60, [math]::Pow(2, $Attempt))
                try {
                    $RetryAfter = $_.Exception.Response.Headers.GetValues("Retry-After") | Select-Object -First 1
                    if ($RetryAfter) { $Delay = [int]$RetryAfter }
                }
                catch { }
                Write-Warning "Graph returned $StatusCode. Retrying in $Delay seconds (attempt $Attempt of $MaxRetries)..."
                Start-Sleep -Seconds $Delay
            }
        }
    }

    # Converts a Teams HTML message body to readable plain text
    function ConvertTo-PlainText {
        param([string]$Html)
        if ([string]::IsNullOrEmpty($Html)) { return "" }
        $Text = $Html -replace '(?i)<br\s*/?>', ' ' -replace '(?i)</(p|div)>', ' '
        $Text = $Text -replace '<[^>]+>', ''
        $Text = [System.Net.WebUtility]::HtmlDecode($Text)
        ($Text -replace '\s+', ' ').Trim()
    }

    # Truncates a string for display
    function Get-Snippet {
        param([string]$Text, [int]$Length = 80)
        if ([string]::IsNullOrEmpty($Text)) { return "" }
        if ($Text.Length -le $Length) { return $Text }
        $Text.Substring(0, $Length) + "..."
    }

    # Presents a picker: Out-GridView where available, otherwise a numbered console menu
    # (Out-GridView does not exist on non-Windows platforms)
    function Select-FromMenu {
        param(
            $Items,
            [string]$Title,
            [switch]$Multiple,
            [scriptblock]$LineFormat
        )
        $OutputMode = "Single"
        if ($Multiple) { $OutputMode = "Multiple" }
        if (Get-Command Out-GridView -ErrorAction SilentlyContinue) {
            return $Items | Out-GridView -Title $Title -OutputMode $OutputMode
        }

        Write-Host "`n$Title" -ForegroundColor Cyan
        $IndexedItems = @($Items)
        for ($i = 0; $i -lt $IndexedItems.Count; $i++) {
            $Line = "$($IndexedItems[$i])"
            if ($LineFormat) { $Line = & $LineFormat $IndexedItems[$i] }
            Write-Host ("  [{0}] {1}" -f ($i + 1), $Line)
        }
        $PromptText = "Enter selection number (blank to cancel)"
        if ($Multiple) { $PromptText = "Enter selection number(s), comma-separated (blank to cancel)" }
        $Answer = Read-Host $PromptText
        if ([string]::IsNullOrWhiteSpace($Answer)) { return $null }

        $Selected = foreach ($Token in ($Answer -split ',')) {
            $Index = 0
            if ([int]::TryParse($Token.Trim(), [ref]$Index) -and $Index -ge 1 -and $Index -le $IndexedItems.Count) {
                $IndexedItems[$Index - 1]
            }
        }
        if (-not $Multiple) { return $Selected | Select-Object -First 1 }
        return $Selected
    }

    # Connect-MgGraph -UseDeviceCode writes the code via Console.WriteLine(), which is
    # invisible in many PowerShell hosts. This handles the flow directly via Write-Host.
    function Connect-GraphDeviceCode {
        param([string]$Tenant, [string]$Scope)
        $GraphPSClientId = "14d82eec-204b-4c2f-b7e8-296a70dab67e"
        $DeviceCode = Invoke-RestMethod -Method POST `
            -Uri "https://login.microsoftonline.com/$Tenant/oauth2/v2.0/devicecode" `
            -Body @{ client_id = $GraphPSClientId; scope = "https://graph.microsoft.com/$Scope" }
        Write-Host "`n$($DeviceCode.message)" -ForegroundColor Yellow
        $TokenBody = @{
            grant_type  = "urn:ietf:params:oauth:grant-type:device_code"
            client_id   = $GraphPSClientId
            device_code = $DeviceCode.device_code
        }
        while ($true) {
            Start-Sleep -Seconds $DeviceCode.interval
            try {
                $Token = Invoke-RestMethod -Method POST `
                    -Uri "https://login.microsoftonline.com/$Tenant/oauth2/v2.0/token" `
                    -Body $TokenBody
                break
            }
            catch {
                $ErrorResponse = $null
                if ($_.ErrorDetails.Message) {
                    try { $ErrorResponse = $_.ErrorDetails.Message | ConvertFrom-Json } catch { }
                }
                if ($ErrorResponse.error -eq "authorization_pending") { continue }
                if ($ErrorResponse.error -eq "authorization_declined") { throw "Authentication was declined." }
                if ($ErrorResponse.error -eq "expired_token") { throw "Device code expired. Run the command again." }
                throw
            }
        }
        $SecureToken = ConvertTo-SecureString $Token.access_token -AsPlainText -Force
        Connect-MgGraph -AccessToken $SecureToken
    }
    #endregion

    #region Register App
    if ($PSCmdlet.ParameterSetName -eq "RegisterApp") {
        if ($UseDeviceCode) {
            Connect-GraphDeviceCode -Tenant $TenantId -Scope "Application.ReadWrite.All AppRoleAssignment.ReadWrite.All"
        }
        else {
            Write-Host "Connecting to Microsoft Graph to register the application..." -ForegroundColor Green
            Connect-MgGraph -Scopes "Application.ReadWrite.All", "AppRoleAssignment.ReadWrite.All" -TenantId $TenantId -ContextScope "process" -NoWelcome
        }

        # Look up the Chat.Read.All app role from the Microsoft Graph service principal
        $GraphAppId = "00000003-0000-0000-c000-000000000000"
        $GraphSPResponse = Invoke-GraphRequestWithRetry -Uri "v1.0/servicePrincipals?`$filter=appId eq '$GraphAppId'"
        $GraphSP = $GraphSPResponse.value[0]
        $ChatReadAllRole = $GraphSP.appRoles | Where-Object { $_.value -eq "Chat.Read.All" }

        if (-not $ChatReadAllRole) {
            Write-Error "Could not find the Chat.Read.All app role on the Microsoft Graph service principal."
            return
        }

        # Reuse an existing registration from a previous run rather than creating duplicates
        $AppDisplayName = "Get-TeamsChatMessages"
        $ExistingApp = (Invoke-GraphRequestWithRetry -Uri "v1.0/applications?`$filter=displayName eq '$AppDisplayName'").value | Select-Object -First 1

        if ($ExistingApp) {
            $App = $ExistingApp
            Write-Host "Found existing app registration '$AppDisplayName' ($($App.appId)). Reusing it and adding a new client secret." -ForegroundColor Yellow
        }
        else {
            # Create the app registration with the required permission
            $AppBody = @{
                displayName            = $AppDisplayName
                signInAudience         = "AzureADMyOrg"
                requiredResourceAccess = @(
                    @{
                        resourceAppId  = $GraphAppId
                        resourceAccess = @(
                            @{
                                id   = $ChatReadAllRole.id
                                type = "Role"
                            }
                        )
                    }
                )
            } | ConvertTo-Json -Depth 5

            $App = Invoke-GraphRequestWithRetry -Method POST -Uri "v1.0/applications" -Body $AppBody
        }

        # Create a client secret (90-day expiration)
        $SecretBody = @{
            passwordCredential = @{
                displayName = "Get-TeamsChatMessages secret"
                endDateTime = (Get-Date).AddDays(90).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ", [System.Globalization.CultureInfo]::InvariantCulture)
            }
        } | ConvertTo-Json -Depth 3

        $Secret = Invoke-GraphRequestWithRetry -Method POST -Uri "v1.0/applications/$($App.id)/addPassword" -Body $SecretBody

        # Ensure the service principal exists so the app is usable for client credential auth
        $SP = (Invoke-GraphRequestWithRetry -Uri "v1.0/servicePrincipals?`$filter=appId eq '$($App.appId)'").value | Select-Object -First 1
        if (-not $SP) {
            $SPBody = @{ appId = $App.appId } | ConvertTo-Json
            $SP = Invoke-GraphRequestWithRetry -Method POST -Uri "v1.0/servicePrincipals" -Body $SPBody
        }

        # Grant admin consent for Chat.Read.All (idempotent: an existing grant counts as granted)
        $ConsentGranted = $false
        try {
            $ConsentBody = @{
                principalId = $SP.id
                resourceId  = $GraphSP.id
                appRoleId   = $ChatReadAllRole.id
            } | ConvertTo-Json
            Invoke-GraphRequestWithRetry -Method POST -Uri "v1.0/servicePrincipals/$($SP.id)/appRoleAssignments" -Body $ConsentBody | Out-Null
            $ConsentGranted = $true
        }
        catch {
            if ($_.Exception.Message -match "already exists") {
                $ConsentGranted = $true
            }
            else {
                Write-Warning "Could not grant admin consent automatically. An admin must grant consent manually:`n  Azure Portal > App Registrations > Get-TeamsChatMessages > API Permissions > Grant admin consent"
            }
        }

        Write-Host "`nApp Registration Ready" -ForegroundColor Green
        Write-Host "  Display Name : $($App.displayName)" -ForegroundColor Cyan
        Write-Host "  Client ID    : $($App.appId)" -ForegroundColor Cyan
        Write-Host "  Client Secret: $($Secret.secretText)" -ForegroundColor Cyan
        Write-Host "  Secret Expiry: $($Secret.endDateTime)" -ForegroundColor Cyan
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

    #region Authentication
    if ($IsAppAuth) {
        if ($ClientSecret) {
            $SecureSecret = ConvertTo-SecureString $ClientSecret -AsPlainText -Force
            $ClientSecretCredential = [PSCredential]::new($ClientId, $SecureSecret)
        }
        else {
            $ClientSecretCredential = Get-Credential -Credential $ClientId
        }
        $ConnectParams = @{
            ClientSecretCredential = $ClientSecretCredential
            ContextScope           = "process"
            NoWelcome              = $true
            TenantId               = $TenantId
        }
        try {
            Connect-MgGraph @ConnectParams -ErrorAction Stop
        }
        catch {
            Write-Error "App authentication failed. If the app was just registered, wait a minute for propagation and verify an admin has granted consent. $_"
            return
        }
    }
    else {
        try {
            if ($UseDeviceCode) {
                Connect-GraphDeviceCode -Tenant "organizations" -Scope "Chat.Read"
            }
            else {
                Connect-MgGraph -Scopes "Chat.Read" -ContextScope "process" -NoWelcome -ErrorAction Stop
            }
        }
        catch {
            Write-Error "Delegated authentication failed. $_"
            return
        }
    }

    # Path roots differ by auth mode. Chat lookups/listing use the user-scoped path;
    # message URIs are kept version-relative so they can also be used in $batch requests.
    $ChatBasePath = "/me/chats"
    $MessagePathPrefix = "/me/chats"
    if ($IsAppAuth) {
        $ChatBasePath = "/users/$UserId/chats"
        $MessagePathPrefix = "/chats"
    }
    #endregion

    #region Resolve Target Chats
    $TargetChats = [Collections.Generic.List[Object]]::new()

    if ($ChatId) {
        # Look up each specified chat for a friendly label; fall back to the raw ID
        foreach ($Id in $ChatId) {
            $Label = $Id
            try {
                $Chat = Invoke-GraphRequestWithRetry -Uri "v1.0${ChatBasePath}/${Id}?`$expand=members"
                if ($Chat.topic) { $Label = $Chat.topic }
                elseif ($Chat.members) { $Label = ($Chat.members.displayName -join "; ") }
            }
            catch {
                Write-Warning "Could not look up chat '$Id' for a display name: $($_.Exception.Message)"
            }
            $TargetChats.Add([PSCustomObject]@{ ChatId = $Id; Label = $Label })
        }
    }
    else {
        $ChatLimit = if ($ChatCount -eq "Unlimited" -or $ChatCount -eq "0") { 0 } else { [int]$ChatCount }
        $InitialTop = 50
        if ($ChatLimit -gt 0 -and $ChatLimit -lt 50) { $InitialTop = $ChatLimit }

        # Server-side ordering by most recent message so "top N" means the N most recent
        # chats, and lastMessagePreview so the picker can show what each chat was about
        $ChatsPageUri = "v1.0${ChatBasePath}?`$top=$InitialTop&`$expand=members,lastMessagePreview&`$orderby=lastMessagePreview/createdDateTime desc"
        $AllChats = [Collections.Generic.List[Object]]::new()
        try {
            while ($ChatsPageUri -and ($ChatLimit -eq 0 -or $AllChats.Count -lt $ChatLimit)) {
                $ChatsResponse = Invoke-GraphRequestWithRetry -Uri $ChatsPageUri
                if ($ChatsResponse.value) {
                    $AllChats.AddRange(@($ChatsResponse.value))
                }
                $ProgressParams = @{
                    Activity = "Retrieving chats"
                    Status   = "$($AllChats.Count) chats retrieved"
                }
                if ($ChatLimit -gt 0) {
                    $ProgressParams.PercentComplete = [math]::Min(100, [int]($AllChats.Count / $ChatLimit * 100))
                }
                Write-Progress @ProgressParams
                $ChatsPageUri = $ChatsResponse.'@odata.nextlink'
            }
        }
        catch {
            Write-Progress -Activity "Retrieving chats" -Completed
            Write-Error "Failed to retrieve chats. If the app was just registered, the permission grant may still be propagating -- wait 30 seconds and try again. $_"
            return
        }
        Write-Progress -Activity "Retrieving chats" -Completed
        if ($ChatLimit -gt 0 -and $AllChats.Count -gt $ChatLimit) {
            $AllChats = $AllChats.GetRange(0, $ChatLimit)
        }
        Write-Host "Retrieved $($AllChats.Count) chats" -ForegroundColor Green

        $ChatList = foreach ($Chat in $AllChats) {
            $LastActivity = $null
            if ($Chat.lastMessagePreview.createdDateTime) { $LastActivity = [datetime]$Chat.lastMessagePreview.createdDateTime }
            elseif ($Chat.lastUpdatedDateTime) { $LastActivity = [datetime]$Chat.lastUpdatedDateTime }

            [PSCustomObject]@{
                ChatType     = $Chat.chatType
                LastActivity = $LastActivity
                Topic        = $Chat.topic
                MemberCount  = @($Chat.members).Count
                Members      = ($Chat.members.displayName -join "; ")
                LastMessage  = Get-Snippet (ConvertTo-PlainText $Chat.lastMessagePreview.body.content) 60
                ChatId       = $Chat.id
            }
        }

        $SelectedChats = Select-FromMenu -Items ($ChatList | Sort-Object LastActivity -Descending) `
            -Title "Select one or more Teams chats (Ctrl+Click for multiple), then click OK" -Multiple -LineFormat {
            param($Item)
            $Name = $Item.Topic
            if (-not $Name) { $Name = $Item.Members }
            "{0,-9} {1:yyyy-MM-dd HH:mm}  {2}" -f $Item.ChatType, $Item.LastActivity, (Get-Snippet $Name 60)
        }

        if (-not $SelectedChats) {
            Write-Error "No chat selected. Exiting."
            return
        }

        foreach ($Chat in @($SelectedChats)) {
            $Label = $Chat.Topic
            if (-not $Label) { $Label = $Chat.Members }
            if (-not $Label) { $Label = $Chat.ChatId }
            $TargetChats.Add([PSCustomObject]@{ ChatId = $Chat.ChatId; Label = $Label })
        }
    }

    $LabelById = @{}
    foreach ($Chat in $TargetChats) { $LabelById[$Chat.ChatId] = $Chat.Label }

    Write-Host "Selected $($TargetChats.Count) chat(s):" -ForegroundColor Cyan
    foreach ($Chat in $TargetChats) {
        Write-Host "  $(Get-Snippet $Chat.Label 70)  [$($Chat.ChatId)]" -ForegroundColor DarkCyan
    }
    #endregion

    #region Date Filter
    switch ($DateMode) {
        "Relative" {
            $Start = Format-GraphDate (Get-Date).AddDays(-$DaysAgo)
            $End = Format-GraphDate (Get-Date)
            Write-Host "Date range: Last $DaysAgo days" -ForegroundColor Yellow
        }
        "Specific" {
            $Start = Format-GraphDate $StartDate
            $End = Format-GraphDate $EndDate.AddDays(1)
            Write-Host "Date range: $StartDate to $EndDate" -ForegroundColor Yellow
        }
        "Window" {
            $Start = Format-GraphDate (Get-Date).AddDays(-$WindowStart)
            $End = Format-GraphDate (Get-Date).AddDays(-$WindowEnd)
            Write-Host "Date range: $WindowStart to $WindowEnd days ago" -ForegroundColor Yellow
        }
    }

    $Filter = "lastModifiedDateTime gt $Start and lastModifiedDateTime lt $End"
    $EncodedFilter = $Filter -replace ' ', '%20'
    #endregion

    #region Retrieve Messages (Batched + Paginated)
    $AllChatMessages = [Collections.Generic.List[Object]]::new()

    # Track the next URI to fetch per chat; each starts at its first page.
    # Initial URIs are version-relative ("/chats/...") for use inside $batch.
    $PendingUris = @{}
    foreach ($Chat in $TargetChats) {
        $PendingUris[$Chat.ChatId] = "$MessagePathPrefix/$($Chat.ChatId)/messages?`$top=$PageSize&`$filter=$EncodedFilter"
    }

    # Phase 1: when multiple chats are selected, fetch each chat's first page in a single
    # JSON batch request (up to 20 requests per batch) instead of one round trip per chat
    if ($TargetChats.Count -gt 1) {
        Write-Progress -Activity "Retrieving messages" -Status "Batching first page of $($TargetChats.Count) chats"
        $BatchIds = @($PendingUris.Keys)
        for ($Offset = 0; $Offset -lt $BatchIds.Count; $Offset += 20) {
            $SliceEnd = [math]::Min($Offset + 19, $BatchIds.Count - 1)
            $Slice = @($BatchIds[$Offset..$SliceEnd])
            $Requests = for ($i = 0; $i -lt $Slice.Count; $i++) {
                @{ id = "$i"; method = "GET"; url = $PendingUris[$Slice[$i]] }
            }
            $BatchBody = @{ requests = @($Requests) } | ConvertTo-Json -Depth 4
            $BatchResponse = Invoke-GraphRequestWithRetry -Uri 'v1.0/$batch' -Method POST -Body $BatchBody

            foreach ($Response in $BatchResponse.responses) {
                $ChatKey = $Slice[[int]$Response.id]
                if ($Response.status -eq 200) {
                    if ($Response.body.value) {
                        $AllChatMessages.AddRange(@($Response.body.value))
                    }
                    if ($Response.body.'@odata.nextlink') {
                        $PendingUris[$ChatKey] = $Response.body.'@odata.nextlink'
                    }
                    else {
                        $PendingUris.Remove($ChatKey)
                    }
                }
                elseif ($Response.status -in 429, 500, 502, 503, 504) {
                    # Leave the first-page URI in place; Phase 2 retries it with backoff
                }
                else {
                    Write-Warning "Skipping chat '$($LabelById[$ChatKey])': HTTP $($Response.status) $($Response.body.error.message)"
                    $PendingUris.Remove($ChatKey)
                }
            }
        }
    }

    # Phase 2: follow each chat's remaining pages (nextLink URLs are absolute)
    $RequestCount = 0
    while ($PendingUris.Count -gt 0) {
        foreach ($ChatKey in @($PendingUris.Keys)) {
            $RequestCount++
            $PageUri = $PendingUris[$ChatKey]
            if ($PageUri -notmatch '^https://') { $PageUri = "v1.0$PageUri" }
            try {
                $PageResults = Invoke-GraphRequestWithRetry -Uri $PageUri
            }
            catch {
                Write-Warning "Failed retrieving messages for '$($LabelById[$ChatKey])'. Results may be partial. $($_.Exception.Message)"
                $PendingUris.Remove($ChatKey)
                continue
            }
            if ($PageResults.value) {
                $AllChatMessages.AddRange(@($PageResults.value))
            }
            if ($PageResults.'@odata.nextlink') {
                $PendingUris[$ChatKey] = $PageResults.'@odata.nextlink'
            }
            else {
                $PendingUris.Remove($ChatKey)
            }
            Write-Progress -Activity "Retrieving messages" -Status "$($AllChatMessages.Count) messages retrieved (request $RequestCount) - $(Get-Snippet $LabelById[$ChatKey] 40)"
        }
    }

    Write-Progress -Activity "Retrieving messages" -Completed
    Write-Host "Retrieved $($AllChatMessages.Count) messages from $($TargetChats.Count) chat(s)" -ForegroundColor Green
    #endregion

    #region Process Messages
    $DeletedCount = 0
    $SystemCount = 0

    $ProcessedMessages = foreach ($Message in $AllChatMessages) {
        if ($Message.deletedDateTime) {
            $DeletedCount++
            continue
        }
        $IsSystemEvent = ($Message.messageType -ne "message") -or $Message.eventDetail
        if ($IsSystemEvent -and -not $IncludeSystemMessages) {
            $SystemCount++
            continue
        }

        $From = $Message.from.user.displayName
        if (-not $From) { $From = $Message.from.application.displayName }
        if (-not $From) { $From = "(system)" }

        $BodyHtml = $Message.body.content
        $BodyText = ConvertTo-PlainText $BodyHtml
        if ($IsSystemEvent -and -not $BodyText) {
            $EventType = "systemEvent"
            if ($Message.eventDetail.'@odata.type') {
                $EventType = $Message.eventDetail.'@odata.type' -replace '^#microsoft\.graph\.', '' -replace 'EventMessageDetail$', ''
            }
            $BodyText = "(event: $EventType)"
        }

        $ReactionSummary = ""
        if ($Message.reactions) {
            $ReactionSummary = (@($Message.reactions) | Group-Object reactionType | ForEach-Object {
                    if ($_.Count -gt 1) { "$($_.Name) x$($_.Count)" } else { $_.Name }
                }) -join ", "
        }

        $ReplyToId = $null
        $ReplyToPreview = $null
        $MessageRef = $Message.attachments | Where-Object contentType -eq "messageReference" | Select-Object -First 1
        if ($MessageRef.content) {
            try {
                $RefContent = $MessageRef.content | ConvertFrom-Json
                $ReplyToId = $RefContent.messageId
                $ReplyToPreview = Get-Snippet (ConvertTo-PlainText $RefContent.messagePreview) 60
            }
            catch { }
        }

        $AttachmentNames = (@($Message.attachments) |
            Where-Object { $_.contentType -ne "messageReference" -and $_.name } |
            ForEach-Object { $_.name }) -join "; "

        [PSCustomObject]@{
            chat             = $LabelById[$Message.chatId]
            from             = $From
            createdDateTime  = [datetime]$Message.createdDateTime
            bodyText         = $BodyText
            reactions        = $ReactionSummary
            replyToPreview   = $ReplyToPreview
            attachments      = $AttachmentNames
            isEdited         = [bool]$Message.lastEditedDateTime
            replyToMessageId = $ReplyToId
            id               = $Message.id
            chatId           = $Message.chatId
            bodyHtml         = $BodyHtml
        }
    }

    $ProcessedMessages = @($ProcessedMessages | Sort-Object chat, createdDateTime)
    #endregion

    #region Summary
    Write-Host "`nMessage Summary:" -ForegroundColor Magenta
    if ($ProcessedMessages.Count -gt 0) {
        $DateRange = $ProcessedMessages | Measure-Object -Property createdDateTime -Minimum -Maximum
        Write-Host ("  Total: {0}  |  {1:yyyy-MM-dd HH:mm} to {2:yyyy-MM-dd HH:mm}" -f $ProcessedMessages.Count, $DateRange.Minimum, $DateRange.Maximum)
        Write-Host "  With reactions: $(@($ProcessedMessages | Where-Object reactions).Count)  |  Replies: $(@($ProcessedMessages | Where-Object replyToMessageId).Count)  |  Edited: $(@($ProcessedMessages | Where-Object isEdited).Count)"
        $TopSenders = $ProcessedMessages | Group-Object from | Sort-Object Count -Descending | Select-Object -First 3
        Write-Host "  Top senders: $(@($TopSenders | ForEach-Object { "$($_.Name) ($($_.Count))" }) -join ', ')"
    }
    else {
        Write-Host "  No messages found in the selected date range." -ForegroundColor Yellow
    }
    if ($DeletedCount -or $SystemCount) {
        $SkippedParts = @()
        if ($DeletedCount) { $SkippedParts += "$DeletedCount deleted" }
        if ($SystemCount) { $SkippedParts += "$SystemCount system events (use -IncludeSystemMessages to include)" }
        Write-Host "  Skipped: $($SkippedParts -join ', ')" -ForegroundColor DarkGray
    }
    #endregion

    if ($ProcessedMessages.Count -eq 0) {
        return $ProcessedMessages
    }

    #region Output
    # Renders messages as a readable HTML transcript, grouped by chat and day
    function ConvertTo-ChatTranscriptHtml {
        param($Messages)
        $Sb = [System.Text.StringBuilder]::new()
        [void]$Sb.AppendLine('<!DOCTYPE html><html><head><meta charset="utf-8"><title>Teams Chat Export</title><style>')
        [void]$Sb.AppendLine('body{font-family:"Segoe UI",Arial,sans-serif;background:#f5f5f5;margin:0;padding:24px;color:#242424}')
        [void]$Sb.AppendLine('.chat{max-width:760px;margin:0 auto 32px}')
        [void]$Sb.AppendLine('h1{font-size:18px;border-bottom:2px solid #6264a7;padding-bottom:8px}')
        [void]$Sb.AppendLine('.day{text-align:center;color:#888;font-size:12px;margin:16px 0 8px}')
        [void]$Sb.AppendLine('.msg{background:#fff;border-radius:8px;padding:8px 12px;margin:6px 0;box-shadow:0 1px 2px rgba(0,0,0,.08)}')
        [void]$Sb.AppendLine('.sender{font-weight:600;font-size:13px}')
        [void]$Sb.AppendLine('.time{color:#888;font-size:11px;margin-left:8px}')
        [void]$Sb.AppendLine('.body{font-size:14px;margin-top:2px;white-space:pre-wrap}')
        [void]$Sb.AppendLine('.meta{color:#6264a7;font-size:12px;margin-top:4px}')
        [void]$Sb.AppendLine('.reply{border-left:3px solid #6264a7;color:#666;font-size:12px;margin:4px 0;padding:2px 8px;background:#fafafa}')
        [void]$Sb.AppendLine('</style></head><body>')

        foreach ($ChatGroup in ($Messages | Group-Object chat)) {
            [void]$Sb.AppendLine('<div class="chat">')
            [void]$Sb.AppendLine("<h1>$([System.Net.WebUtility]::HtmlEncode($ChatGroup.Name))</h1>")
            foreach ($DayGroup in ($ChatGroup.Group | Group-Object { $_.createdDateTime.ToString("yyyy-MM-dd (dddd)") })) {
                [void]$Sb.AppendLine("<div class=""day"">$([System.Net.WebUtility]::HtmlEncode($DayGroup.Name))</div>")
                foreach ($Msg in $DayGroup.Group) {
                    [void]$Sb.AppendLine('<div class="msg">')
                    [void]$Sb.AppendLine("<span class=""sender"">$([System.Net.WebUtility]::HtmlEncode($Msg.from))</span><span class=""time"">$($Msg.createdDateTime.ToString('HH:mm'))$(if ($Msg.isEdited) { ' (edited)' })</span>")
                    if ($Msg.replyToPreview) {
                        [void]$Sb.AppendLine("<div class=""reply"">$([System.Net.WebUtility]::HtmlEncode($Msg.replyToPreview))</div>")
                    }
                    [void]$Sb.AppendLine("<div class=""body"">$([System.Net.WebUtility]::HtmlEncode($Msg.bodyText))</div>")
                    $MetaParts = @()
                    if ($Msg.reactions) { $MetaParts += "Reactions: $($Msg.reactions)" }
                    if ($Msg.attachments) { $MetaParts += "Attachments: $($Msg.attachments)" }
                    if ($MetaParts) {
                        [void]$Sb.AppendLine("<div class=""meta"">$([System.Net.WebUtility]::HtmlEncode(($MetaParts -join '  |  ')))</div>")
                    }
                    [void]$Sb.AppendLine('</div>')
                }
            }
            [void]$Sb.AppendLine('</div>')
        }
        [void]$Sb.AppendLine('</body></html>')
        $Sb.ToString()
    }

    $Timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'

    $CsvData = {
        $ProcessedMessages | ForEach-Object {
            [PSCustomObject]@{
                Chat            = $_.chat
                From            = $_.from
                CreatedDateTime = $_.createdDateTime
                Body            = $_.bodyText
                Reactions       = $_.reactions
                Attachments     = $_.attachments
                Edited          = $_.isEdited
                ReplyToPreview  = $_.replyToPreview
                ReplyToId       = $_.replyToMessageId
                MessageId       = $_.id
                ChatId          = $_.chatId
            }
        }
    }

    $NonInteractive = $PSBoundParameters.ContainsKey('ExportFormat') -or
                      $PSBoundParameters.ContainsKey('ExportToFile') -or
                      $PSBoundParameters.ContainsKey('ShowInConsole')

    if ($NonInteractive) {
        # Parameter-driven export (non-interactive / scripted)
        if ($ExportFormat -and $ExportFormat -ne "None") {
            switch ($ExportFormat) {
                "JSON" {
                    $ExportContent = $ProcessedMessages | ConvertTo-Json -Depth 4
                    $FileExtension = "json"
                }
                "CSV" {
                    $ExportContent = (& $CsvData) | ConvertTo-Csv -NoTypeInformation
                    $FileExtension = "csv"
                }
                "HTML" {
                    $ExportContent = ConvertTo-ChatTranscriptHtml $ProcessedMessages
                    $FileExtension = "html"
                }
            }

            if ($ExportToFile) {
                $FileName = "$ExportPath$Timestamp.$FileExtension"
                $ExportContent | Out-File -FilePath $FileName -Encoding UTF8
                Write-Host "  Saved to: $((Resolve-Path $FileName).Path)" -ForegroundColor Green
            }
            else {
                $ExportContent | Set-Clipboard
                Write-Host "  $ExportFormat copied to clipboard" -ForegroundColor Green
            }
        }

        if ($ShowInConsole) {
            $BodyColumn = @{ n = "body"; e = { Get-Snippet $_.bodyText 100 } }
            if ($ShowFullBody) {
                $BodyColumn = @{ n = "body"; e = { $_.bodyText } }
            }
            $ProcessedMessages | Select-Object -Last $ShowLastNMessages |
                Format-Table from, createdDateTime, $BodyColumn, reactions, replyToMessageId -Wrap
        }
    }
    else {
        # Interactive output picker
        $OutputOptions = @(
            [PSCustomObject]@{ Action = "Copy JSON to clipboard" }
            [PSCustomObject]@{ Action = "Copy CSV to clipboard" }
            [PSCustomObject]@{ Action = "Save JSON file" }
            [PSCustomObject]@{ Action = "Save CSV file" }
            [PSCustomObject]@{ Action = "Save HTML transcript to Desktop" }
            [PSCustomObject]@{ Action = "View messages in grid view" }
            [PSCustomObject]@{ Action = "Show last 10 in console" }
        )

        $SelectedOutputs = Select-FromMenu -Items $OutputOptions `
            -Title "Select Output Options (select one or more, then click OK)" -Multiple -LineFormat {
            param($Item)
            $Item.Action
        }

        foreach ($Output in @($SelectedOutputs)) {
            switch ($Output.Action) {
                "Copy JSON to clipboard" {
                    $ProcessedMessages | ConvertTo-Json -Depth 4 | Set-Clipboard
                    Write-Host "  JSON copied to clipboard" -ForegroundColor Green
                }
                "Copy CSV to clipboard" {
                    (& $CsvData) | ConvertTo-Csv -NoTypeInformation | Set-Clipboard
                    Write-Host "  CSV copied to clipboard" -ForegroundColor Green
                }
                "Save JSON file" {
                    $FileName = ".\TeamsChat_$Timestamp.json"
                    $ProcessedMessages | ConvertTo-Json -Depth 4 | Out-File -FilePath $FileName -Encoding UTF8
                    Write-Host "  Saved to: $((Resolve-Path $FileName).Path)" -ForegroundColor Green
                }
                "Save CSV file" {
                    $FileName = ".\TeamsChat_$Timestamp.csv"
                    (& $CsvData) | Export-Csv -Path $FileName -NoTypeInformation -Encoding UTF8
                    Write-Host "  Saved to: $((Resolve-Path $FileName).Path)" -ForegroundColor Green
                }
                "Save HTML transcript to Desktop" {
                    # GetFolderPath resolves the real Desktop even when OneDrive redirects it
                    $FileName = Join-Path ([Environment]::GetFolderPath("Desktop")) "TeamsChat_$Timestamp.html"
                    ConvertTo-ChatTranscriptHtml $ProcessedMessages | Out-File -FilePath $FileName -Encoding UTF8
                    Write-Host "  Saved to: $FileName" -ForegroundColor Green
                }
                "View messages in grid view" {
                    $GridProperties = "chat", "from", "createdDateTime", "bodyText", "reactions", "attachments", "isEdited", "replyToPreview"
                    if (Get-Command Out-GridView -ErrorAction SilentlyContinue) {
                        $ProcessedMessages | Select-Object $GridProperties |
                            Out-GridView -Title "Teams Chat Messages ($($ProcessedMessages.Count) messages)"
                    }
                    else {
                        $ProcessedMessages | Select-Object $GridProperties | Format-Table -Wrap
                    }
                }
                "Show last 10 in console" {
                    Write-Host ""
                    $ProcessedMessages | Select-Object -Last 10 |
                        Format-Table from, createdDateTime, @{ n = "body"; e = { Get-Snippet $_.bodyText 100 } }, reactions -Wrap
                }
            }
        }
    }
    #endregion

    return $ProcessedMessages
}
