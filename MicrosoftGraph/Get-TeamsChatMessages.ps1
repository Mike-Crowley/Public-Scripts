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
        Creates an Entra ID app registration with the Chat.Read.All application permission
        and a client secret. Run this once before using app auth. Requires permission to
        register apps in the tenant. An admin must grant consent after registration.

    .PARAMETER ClientId
        The Application (client) ID of the Azure AD app registration. Required for app auth.

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
        Start date in yyyy-MM-dd format. Must be used with -EndDate.

    .PARAMETER EndDate
        End date in yyyy-MM-dd format (inclusive). Must be used with -StartDate.

    .PARAMETER WindowStart
        Start of window as days ago. Must be used with -WindowEnd.

    .PARAMETER WindowEnd
        End of window as days ago (must be less than WindowStart). Must be used with -WindowStart.

    .PARAMETER PageSize
        Number of messages per Graph API page. Maximum is 50. Default is 50.

    .PARAMETER ExportFormat
        Output format: JSON, CSV, or None. Default is JSON.

    .PARAMETER ExportToFile
        If specified, saves output to a file instead of copying to clipboard.

    .PARAMETER ExportPath
        File path prefix when using -ExportToFile. A timestamp and extension are appended.
        Default is '.\TeamsChat_'.

    .PARAMETER ShowInConsole
        If specified, displays a sample of recent messages in the console.

    .PARAMETER ShowLastNMessages
        Number of recent messages to display in console. Default is 10.

    .PARAMETER ShowFullBody
        If specified, shows the full message body instead of truncating to 50 characters.

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
        Get-TeamsChatMessages -StartDate "2025-06-01" -EndDate "2025-06-15" -ExportToFile

        Retrieves your messages from June 1-15 using delegated auth and exports to a file.

    .EXAMPLE
        Get-TeamsChatMessages -WindowStart 14 -WindowEnd 7 -ExportFormat CSV

        Retrieves your messages from 14 to 7 days ago and copies CSV to clipboard.

    .EXAMPLE
        Get-TeamsChatMessages -ClientId "abc-123" -TenantId "contoso.com" -UserId "user@contoso.com" -DaysAgo 7

        Uses app auth with a client secret to retrieve another user's messages from the last 7 days.
        The -UserId parameter also accepts a user's object ID (GUID).

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
        [string]$ClientId,

        [Parameter(Mandatory, ParameterSetName = "RegisterApp")]
        [Parameter(Mandatory, ParameterSetName = "AppAuthRelative")]
        [Parameter(Mandatory, ParameterSetName = "AppAuthSpecific")]
        [Parameter(Mandatory, ParameterSetName = "AppAuthWindow")]
        [string]$TenantId,

        [Parameter(Mandatory, ParameterSetName = "AppAuthRelative")]
        [Parameter(Mandatory, ParameterSetName = "AppAuthSpecific")]
        [Parameter(Mandatory, ParameterSetName = "AppAuthWindow")]
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
        [string]$StartDate,

        [Parameter(Mandatory, ParameterSetName = "DelegatedSpecific")]
        [Parameter(Mandatory, ParameterSetName = "AppAuthSpecific")]
        [string]$EndDate,

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
        [ValidateRange(1, 50)]
        [int]$PageSize = 50,

        [Parameter(ParameterSetName = "DelegatedRelative")]
        [Parameter(ParameterSetName = "DelegatedSpecific")]
        [Parameter(ParameterSetName = "DelegatedWindow")]
        [Parameter(ParameterSetName = "AppAuthRelative")]
        [Parameter(ParameterSetName = "AppAuthSpecific")]
        [Parameter(ParameterSetName = "AppAuthWindow")]
        [ValidateSet("JSON", "CSV", "None")]
        [string]$ExportFormat = "JSON",

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

    #region Register App
    if ($PSCmdlet.ParameterSetName -eq "RegisterApp") {
        Write-Host "Connecting to Microsoft Graph to register the application..." -ForegroundColor Green
        $RegisterParams = @{
            Scopes       = "Application.ReadWrite.All"
            TenantId     = $TenantId
            ContextScope = "process"
            NoWelcome    = $true
        }
        if ($UseDeviceCode) { $RegisterParams.UseDeviceCode = $true }
        Connect-MgGraph @RegisterParams

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

        # Create a client secret
        $SecretBody = @{
            passwordCredential = @{
                displayName = "Get-TeamsChatMessages secret"
            }
        } | ConvertTo-Json -Depth 3

        $Secret = Invoke-MgGraphRequest -Method POST -Uri "v1.0/applications/$($App.id)/addPassword" -Body $SecretBody -ContentType "application/json" -OutputType PSObject

        Write-Host "`nApp Registration Created Successfully" -ForegroundColor Green
        Write-Host "  Display Name : $($App.displayName)" -ForegroundColor Cyan
        Write-Host "  Client ID    : $($App.appId)" -ForegroundColor Cyan
        Write-Host "  Client Secret: $($Secret.secretText)" -ForegroundColor Cyan
        Write-Host "  Secret Expiry: $($Secret.endDateTime)" -ForegroundColor Cyan
        Write-Host "`n  IMPORTANT: A tenant admin must grant consent for Chat.Read.All before this app can be used." -ForegroundColor Yellow
        Write-Host "  Azure Portal > App Registrations > Get-TeamsChatMessages > API Permissions > Grant admin consent" -ForegroundColor Yellow
        Write-Host "`n  Save the Client ID and Secret now. The secret value cannot be retrieved later." -ForegroundColor Yellow

        return [PSCustomObject]@{
            DisplayName  = $App.displayName
            ClientId     = $App.appId
            ObjectId     = $App.id
            ClientSecret = $Secret.secretText
            SecretExpiry = $Secret.endDateTime
        }
    }
    #endregion

    #region Authentication
    if ($IsAppAuth) {
        $ClientSecretCredential = Get-Credential -Credential $ClientId
        $ConnectParams = @{
            ClientSecretCredential = $ClientSecretCredential
            ContextScope           = "process"
            NoWelcome              = $true
            TenantId               = $TenantId
        }
        Connect-MgGraph @ConnectParams
        $ChatsUri = "v1.0/users/$UserId/chats?`$top=$PageSize&`$expand=members"
    }
    else {
        $DelegatedParams = @{
            Scopes       = "Chat.Read"
            ContextScope = "process"
            NoWelcome    = $true
        }
        if ($UseDeviceCode) { $DelegatedParams.UseDeviceCode = $true }
        Connect-MgGraph @DelegatedParams
        $ChatsUri = "v1.0/me/chats?`$top=$PageSize&`$expand=members"
    }
    #endregion

    #region Chat Selection
    Write-Host "Retrieving available chats..." -ForegroundColor Green
    $ChatsParams = @{
        OutputType = "PSObject"
        Uri        = $ChatsUri
    }
    $ChatsResponse = Invoke-MgGraphRequest @ChatsParams

    $ChatList = $ChatsResponse.value | ForEach-Object {
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
            $Start = ([DateTime]$StartDate).ToUniversalTime().ToString($DateFormat)
            $End = ([DateTime]$EndDate).AddDays(1).ToUniversalTime().ToString($DateFormat)
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

    Write-Host "Retrieving messages..." -ForegroundColor Green
    do {
        $RequestParams = @{
            OutputType = "PSObject"
            Uri        = $MessagesUri
        }

        $PageResults = Invoke-MgGraphRequest @RequestParams

        if ($PageResults.value) {
            Write-Host "  Retrieved $($PageResults.value.Count) messages from current page" -ForegroundColor Yellow
            $AllChatMessages.AddRange($PageResults.value)
        }
        else {
            $AllChatMessages.Add($PageResults)
        }

        $MessagesUri = $PageResults.'@odata.nextlink'
    } until (-not $MessagesUri)

    Write-Host "Total messages retrieved: $($AllChatMessages.Count)" -ForegroundColor Green
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
    Write-Host "- Total messages processed: $($ProcessedMessages.Count)"
    Write-Host "- Messages with reactions: $(($ProcessedMessages | Where-Object reactions).Count)"
    Write-Host "- Reply messages: $(($ProcessedMessages | Where-Object replyToMessageId).Count)"
    #endregion

    #region Export
    if ($ExportFormat -ne "None") {
        switch ($ExportFormat) {
            "JSON" {
                $ExportContent = $ProcessedMessages | ConvertTo-Json -Depth 3
                $FileExtension = "json"
            }
            "CSV" {
                $ExportContent = $ProcessedMessages | ForEach-Object {
                    [PSCustomObject]@{
                        From            = $_.from
                        CreatedDateTime = $_.createdDateTime
                        Body            = $_.body -replace "`n", " " -replace "`r", " "
                        Reactions       = ($_.reactions -join ", ")
                        ReplyToId       = $_.replyToMessageId
                        ReplyToPreview  = $_.replyToPreview
                        MessageId       = $_.id
                    }
                } | ConvertTo-Csv -NoTypeInformation
                $FileExtension = "csv"
            }
        }

        if ($ExportToFile) {
            $FileName = "$ExportPath$(Get-Date -Format 'yyyyMMdd_HHmmss').$FileExtension"
            $ExportContent | Out-File -FilePath $FileName -Encoding UTF8
            Write-Host "`nMessages exported to: $FileName" -ForegroundColor Green
        }
        else {
            $ExportContent | Set-Clipboard
            Write-Host "`nMessages copied to clipboard as $ExportFormat" -ForegroundColor Green
        }
    }
    #endregion

    #region Console Display
    if ($ShowInConsole) {
        Write-Host "`nLast $ShowLastNMessages messages:" -ForegroundColor Cyan

        $DisplayProperties = @{
            Property = if ($ShowFullBody) {
                "from", "createdDateTime", "body", "reactions", "replyToMessageId"
            }
            else {
                "from", "createdDateTime", @{n = "body"; e = { $_.body[0..50] -join "" } }, "reactions", "replyToMessageId"
            }
        }

        $ProcessedMessages |
            Select-Object -Last $ShowLastNMessages |
            Format-Table @DisplayProperties -Wrap
    }
    #endregion

    return $ProcessedMessages
}
