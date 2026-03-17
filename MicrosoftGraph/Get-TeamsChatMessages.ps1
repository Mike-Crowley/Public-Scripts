function Get-TeamsChatMessages {
    <#
    .SYNOPSIS
        Retrieves and processes Microsoft Teams chat messages using Microsoft Graph API.

    .DESCRIPTION
        Connects to Microsoft Graph and fetches all chat messages within a date range using
        pagination. Extracts message properties including reactions, replies, and body content.
        Presents an Out-GridView picker for chat selection.

        By default, uses delegated authentication (interactive sign-in with Chat.Read scope)
        and queries the signed-in user's own chats via the /me endpoint. The default date
        range mode is Relative, specified with -DaysAgo.

        Use -StartDate/-EndDate for a specific date range, or -WindowStart/-WindowEnd for a
        sliding window. These date parameters are mutually exclusive.

        Use the -ClientId, -TenantId, and -UserId parameters to authenticate with
        application permissions (Chat.Read.All) instead, which allows querying any user's
        chats. Run with -RegisterApp -TenantId first to create the required app registration.

    .PARAMETER RegisterApp
        Creates an Entra ID app registration with the Chat.Read.All application permission,
        a client secret, service principal, and grants admin consent. Run this once before
        using app auth. Requires Application.ReadWrite.All and AppRoleAssignment.ReadWrite.All
        permissions (typically Application Administrator or Global Administrator).

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

    .PARAMETER ChatCount
        Number of most recent chats to retrieve for the picker. Default is 100.
        Use 0 or "Unlimited" to retrieve all chats.

    .PARAMETER PageSize
        Number of messages per Graph API page. Maximum is 50. Default is 50.

    .PARAMETER ExportFormat
        Output format for non-interactive mode: JSON, CSV, or None. When specified, skips
        the interactive output picker and uses this format directly.

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

        Creates an Entra ID app registration with Chat.Read.All permission in the specified
        tenant and outputs the Client ID and secret. An admin must grant consent before the
        app can be used.

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

    .NOTES
        Author: Mike Crowley
        https://mikecrowley.us

        Requires the Microsoft.Graph.Authentication module.
        Delegated auth requires the Chat.Read scope.
        App auth requires Chat.Read.All application permissions on the app registration.

        TIP: To read another user's chats, you need an app registration with Chat.Read.All.
             Run 'Get-TeamsChatMessages -RegisterApp -TenantId <tenant>' to create one, then
             have a tenant admin grant consent in the Azure portal before using
             -ClientId/-TenantId/-UserId.

    .LINK
        https://github.com/Mike-Crowley/Public-Scripts
    #>

    #Requires -Modules Microsoft.Graph.Authentication

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
        [int]$WindowStart,

        [Parameter(Mandatory, ParameterSetName = "DelegatedWindow")]
        [Parameter(Mandatory, ParameterSetName = "AppAuthWindow")]
        [int]$WindowEnd,

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

        # Non-interactive export parameters (skip the output picker when specified)
        [Parameter(ParameterSetName = "DelegatedRelative")]
        [Parameter(ParameterSetName = "DelegatedSpecific")]
        [Parameter(ParameterSetName = "DelegatedWindow")]
        [Parameter(ParameterSetName = "AppAuthRelative")]
        [Parameter(ParameterSetName = "AppAuthSpecific")]
        [Parameter(ParameterSetName = "AppAuthWindow")]
        [ValidateSet("JSON", "CSV", "None")]
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

    # Clear any existing Graph session to avoid MSAL cache conflicts between auth methods
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null

    #region Device Code Helper
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
                $err = $_.ErrorDetails.Message | ConvertFrom-Json
                if ($err.error -eq "authorization_pending") { continue }
                if ($err.error -eq "authorization_declined") { throw "Authentication was declined." }
                if ($err.error -eq "expired_token") { throw "Device code expired. Run the command again." }
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
        $GraphSPResponse = Invoke-MgGraphRequest -Uri "v1.0/servicePrincipals?`$filter=appId eq '$GraphAppId'" -OutputType PSObject
        $GraphSP = $GraphSPResponse.value[0]
        $ChatReadAllRole = $GraphSP.appRoles | Where-Object { $_.value -eq "Chat.Read.All" }

        if (-not $ChatReadAllRole) {
            Write-Error "Could not find the Chat.Read.All app role on the Microsoft Graph service principal."
            return
        }

        # Create the app registration with the required permission
        $AppBody = @{
            displayName            = "Get-TeamsChatMessages"
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

        $App = Invoke-MgGraphRequest -Method POST -Uri "v1.0/applications" -Body $AppBody -ContentType "application/json" -OutputType PSObject

        # Create a client secret (90-day expiration)
        $SecretBody = @{
            passwordCredential = @{
                displayName = "Get-TeamsChatMessages secret"
                endDateTime = (Get-Date).AddDays(90).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
            }
        } | ConvertTo-Json -Depth 3

        $Secret = Invoke-MgGraphRequest -Method POST -Uri "v1.0/applications/$($App.id)/addPassword" -Body $SecretBody -ContentType "application/json" -OutputType PSObject

        # Create the service principal so the app is usable for client credential auth
        $SPBody = @{ appId = $App.appId } | ConvertTo-Json
        $SP = Invoke-MgGraphRequest -Method POST -Uri "v1.0/servicePrincipals" -Body $SPBody -ContentType "application/json" -OutputType PSObject

        # Grant admin consent for Chat.Read.All
        $ConsentGranted = $false
        try {
            $ConsentBody = @{
                principalId = $SP.id
                resourceId  = $GraphSP.id
                appRoleId   = $ChatReadAllRole.id
            } | ConvertTo-Json
            Invoke-MgGraphRequest -Method POST -Uri "v1.0/servicePrincipals/$($SP.id)/appRoleAssignments" -Body $ConsentBody -ContentType "application/json" -OutputType PSObject | Out-Null
            $ConsentGranted = $true
        }
        catch {
            Write-Warning "Could not grant admin consent automatically. An admin must grant consent manually:`n  Azure Portal > App Registrations > Get-TeamsChatMessages > API Permissions > Grant admin consent"
        }

        Write-Host "`nApp Registration Created Successfully" -ForegroundColor Green
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
        $ChatsUri = "v1.0/users/$UserId/chats?`$top=50&`$expand=members"
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
        $ChatsUri = "v1.0/me/chats?`$top=50&`$expand=members"
    }
    #endregion

    #region Chat Selection
    $ChatLimit = if ($ChatCount -eq "Unlimited" -or $ChatCount -eq "0") { 0 } else { [int]$ChatCount }
    $AllChats = [Collections.Generic.List[Object]]::new()
    $ChatsPageUri = $ChatsUri
    try {
        while ($ChatsPageUri -and ($ChatLimit -eq 0 -or $AllChats.Count -lt $ChatLimit)) {
            $ChatsResponse = Invoke-MgGraphRequest -Uri $ChatsPageUri -OutputType PSObject -ErrorAction Stop
            if ($ChatsResponse.value) {
                $AllChats.AddRange($ChatsResponse.value)
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
        Write-Error "Failed to retrieve chats. If the app was just registered, the permission grant may still be propagating — wait 30 seconds and try again. $_"
        return
    }
    Write-Progress -Activity "Retrieving chats" -Completed
    Write-Host "Retrieved $($AllChats.Count) chats" -ForegroundColor Green

    $ChatList = $AllChats | ForEach-Object {
        [PSCustomObject]@{
            ChatType     = $_.chatType
            LastActivity = $_.lastUpdatedDateTime
            Topic        = $_.topic
            MemberCount  = $_.members.Count
            Members      = ($_.members.displayName -join "; ")
            ChatId       = $_.id
        }
    }

    $SelectedChat = $ChatList |
        Sort-Object { [datetime]$_.LastActivity } -Descending |
        Out-GridView -Title "Select a Teams Chat" -OutputMode Single

    if (-not $SelectedChat) {
        Write-Error "No chat selected. Exiting."
        return
    }

    $ChatId = $SelectedChat.ChatId
    Write-Host "Selected chat: $($SelectedChat.Members)" -ForegroundColor Cyan
    Write-Host "ChatId: $ChatId" -ForegroundColor Yellow
    #endregion

    #region Date Filter
    $DateFormat = "yyyy-MM-ddTHH:mm:ss.fffZ"

    switch ($DateMode) {
        "Relative" {
            $Start = (Get-Date).AddDays(-$DaysAgo).ToUniversalTime().ToString($DateFormat)
            $End = (Get-Date).ToUniversalTime().ToString($DateFormat)
            Write-Host "Date range: Last $DaysAgo days" -ForegroundColor Yellow
        }
        "Specific" {
            $Start = $StartDate.ToUniversalTime().ToString($DateFormat)
            $End = $EndDate.AddDays(1).ToUniversalTime().ToString($DateFormat)
            Write-Host "Date range: $StartDate to $EndDate" -ForegroundColor Yellow
        }
        "Window" {
            $Start = (Get-Date).AddDays(-$WindowStart).ToUniversalTime().ToString($DateFormat)
            $End = (Get-Date).AddDays(-$WindowEnd).ToUniversalTime().ToString($DateFormat)
            Write-Host "Date range: $WindowStart to $WindowEnd days ago" -ForegroundColor Yellow
        }
    }

    $Filter = "lastModifiedDateTime gt $Start and lastModifiedDateTime lt $End"
    #endregion

    #region Retrieve Messages (Paginated)
    if ($IsAppAuth) {
        $MessagesUri = "v1.0/chats/$ChatId/messages?`$top=$PageSize&`$filter=$Filter"
    }
    else {
        $MessagesUri = "v1.0/me/chats/$ChatId/messages?`$top=$PageSize&`$filter=$Filter"
    }

    $AllChatMessages = [Collections.Generic.List[Object]]::new()
    $PageCount = 0

    do {
        $PageCount++
        $PageResults = Invoke-MgGraphRequest -Uri $MessagesUri -OutputType PSObject

        if ($PageResults.value) {
            $AllChatMessages.AddRange($PageResults.value)
        }
        else {
            $AllChatMessages.Add($PageResults)
        }

        Write-Progress -Activity "Retrieving messages" -Status "$($AllChatMessages.Count) messages retrieved (page $PageCount)"
        $MessagesUri = $PageResults.'@odata.nextlink'
    } until (-not $MessagesUri)

    Write-Progress -Activity "Retrieving messages" -Completed
    Write-Host "Retrieved $($AllChatMessages.Count) messages" -ForegroundColor Green
    #endregion

    #region Process Messages
    $SelectProperties = @{
        Property = @(
            @{n = "reactions"; e = { $_.reactions.reactionType } }
            "id"
            @{n = "replyToMessageId"; e = {
                    $messageRef = $_.attachments | Where-Object contentType -eq "messageReference"
                    if ($messageRef) {
                        $refContent = $messageRef.content | ConvertFrom-Json
                        $refContent.messageId
                    }
                }
            }
            @{n = "replyToPreview"; e = {
                    $messageRef = $_.attachments | Where-Object contentType -eq "messageReference"
                    if ($messageRef) {
                        $refContent = $messageRef.content | ConvertFrom-Json
                        $refContent.messagePreview[0..19] -join ""
                    }
                }
            }
            @{n = "body"; e = { $_.body.content } }
            @{n = "createdDateTime"; e = { [datetime]$_.createdDateTime } }
            @{n = "from"; e = { $_.from.user.displayName } }
        )
    }

    $ProcessedMessages = $AllChatMessages | Select-Object @SelectProperties | Sort-Object createdDateTime
    #endregion

    #region Summary
    Write-Host "`nMessage Summary:" -ForegroundColor Magenta
    Write-Host "  Total: $($ProcessedMessages.Count)  |  With reactions: $(($ProcessedMessages | Where-Object reactions).Count)  |  Replies: $(($ProcessedMessages | Where-Object replyToMessageId).Count)"
    #endregion

    #region Output
    $Timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'

    $CsvData = {
        $ProcessedMessages | ForEach-Object {
            [PSCustomObject]@{
                From            = $_.from
                CreatedDateTime = $_.createdDateTime
                Body            = $_.body -replace "`n", " " -replace "`r", " "
                Reactions       = ($_.reactions -join ", ")
                ReplyToId       = $_.replyToMessageId
                ReplyToPreview  = $_.replyToPreview
                MessageId       = $_.id
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
                    $ExportContent = $ProcessedMessages | ConvertTo-Json -Depth 3
                    $FileExtension = "json"
                }
                "CSV" {
                    $ExportContent = (& $CsvData) | ConvertTo-Csv -NoTypeInformation
                    $FileExtension = "csv"
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
            $DisplayProperties = @{
                Property = if ($ShowFullBody) {
                    "from", "createdDateTime", "body", "reactions", "replyToMessageId"
                }
                else {
                    "from", "createdDateTime", @{n = "body"; e = { ($_.body -replace '<[^>]+>')[0..80] -join "" } }, "reactions", "replyToMessageId"
                }
            }
            $ProcessedMessages | Select-Object -Last $ShowLastNMessages |
                Format-Table @DisplayProperties -Wrap
        }
    }
    else {
        # Interactive output picker
        $OutputOptions = @(
            [PSCustomObject]@{ Action = "Copy JSON to clipboard"; Order = 1 }
            [PSCustomObject]@{ Action = "Copy CSV to clipboard"; Order = 2 }
            [PSCustomObject]@{ Action = "Save JSON file"; Order = 3 }
            [PSCustomObject]@{ Action = "Save CSV file"; Order = 4 }
            [PSCustomObject]@{ Action = "View messages in Out-GridView"; Order = 5 }
            [PSCustomObject]@{ Action = "Show last 10 in console"; Order = 6 }
        )

        $SelectedOutputs = $OutputOptions |
            Sort-Object Order |
            Select-Object Action |
            Out-GridView -Title "Select Output Options (select one or more, then click OK)" -OutputMode Multiple

        foreach ($Output in $SelectedOutputs) {
            switch ($Output.Action) {
                "Copy JSON to clipboard" {
                    $ProcessedMessages | ConvertTo-Json -Depth 3 | Set-Clipboard
                    Write-Host "  JSON copied to clipboard" -ForegroundColor Green
                }
                "Copy CSV to clipboard" {
                    (& $CsvData) | ConvertTo-Csv -NoTypeInformation | Set-Clipboard
                    Write-Host "  CSV copied to clipboard" -ForegroundColor Green
                }
                "Save JSON file" {
                    $FileName = ".\TeamsChat_$Timestamp.json"
                    $ProcessedMessages | ConvertTo-Json -Depth 3 | Out-File -FilePath $FileName -Encoding UTF8
                    Write-Host "  Saved to: $((Resolve-Path $FileName).Path)" -ForegroundColor Green
                }
                "Save CSV file" {
                    $FileName = ".\TeamsChat_$Timestamp.csv"
                    (& $CsvData) | Export-Csv -Path $FileName -NoTypeInformation -Encoding UTF8
                    Write-Host "  Saved to: $((Resolve-Path $FileName).Path)" -ForegroundColor Green
                }
                "View messages in Out-GridView" {
                    $ProcessedMessages | Out-GridView -Title "Teams Chat Messages ($($ProcessedMessages.Count) messages)"
                }
                "Show last 10 in console" {
                    Write-Host ""
                    $ProcessedMessages | Select-Object -Last 10 |
                        Format-Table from, createdDateTime, @{n = "body"; e = { ($_.body -replace '<[^>]+>')[0..80] -join "" } }, reactions -Wrap
                }
            }
        }
    }
    #endregion

    return $ProcessedMessages
}
