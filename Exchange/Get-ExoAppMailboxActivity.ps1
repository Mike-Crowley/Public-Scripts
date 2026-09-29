#Requires -Version 5.1
#Requires -Modules @{ ModuleName = 'Microsoft.Graph.Authentication'; ModuleVersion = '2.4.0' }

<#
.SYNOPSIS
    Reports which mailboxes each Exchange-enabled app actually touches, grouped by activity
    type, and builds the RBAC for Applications scope commands from what it found.
.DESCRIPTION
    Companion to Audit-ExoAppAccessPolicies.ps1. That report tells you which apps hold
    tenant-wide Exchange application permissions; this one tells you what they do with them,
    so an admin can judge whether the activity is legitimate and what a scope should contain.

    Source: the unified audit log, read through the Purview audit search Graph API. One
    asynchronous search per app, run through a sliding window of concurrent searches (the
    service throttles past roughly ten per tenant), collected as each finishes. Works in any
    tenant with mailbox auditing on (the default); no Log Analytics required.

    Where the app id sits in a mailbox audit record depends on the access path. For Graph access
    the app is in ClientAppId and AppAccessContext.ClientAppId while AppId holds Microsoft Graph's
    own id. For EWS the app is in AppId, ClientAppId and AppAccessContext.ClientAppId, with EWS's
    own id in AppAccessContext.APIId; a September 2026 test found that for app-only access
    (full_access_as_app) as well as delegated. EWS records written before Microsoft's November 2024
    change (MC909164) named only EWS itself, but no window this script can search reaches back that
    far. Some records also name the app as "[AppId=...]" inside ActorInfoString or
    ClientInfoString. A match in any of those places counts.

    Records whose token belonged to a user rather than to the app's own service principal (the
    record's TokenObjectId) came from a signed-in user working through the app; they are listed
    separately and left out of the scope, since a management scope constrains app-only access,
    not the user's. Session ids do not tell the two apart: app-only EWS records carry
    AppAccessContext.AADSessionId too.

    Output: an HTML report in the audit report's style, one section per app with its
    grants, the operations observed and the mailboxes under each, and a commands block that
    builds the management scope from those mailboxes. A companion CSV carries every
    app/operation/mailbox row for apps whose lists are too long for a page.
.PARAMETER TenantId
    The tenant to report on, as its id or a verified domain (contoso.com). The Graph sign-in
    is pinned to it and scoped to this PowerShell process.
.PARAMETER AppId
    One or more application (client) ids to analyze. Without -AppId or -All, every service
    principal holding Exchange application permissions is offered in Out-GridView.
.PARAMETER All
    Analyze every service principal holding Exchange application permissions, no picker.
.PARAMETER Days
    Audit window, 1 to 180 (Audit Standard retention).
.PARAMETER UserPrincipalName
    Pre-fills the account on the Graph sign-in prompt. A hint, not a lock. Ignored with
    -UseDeviceCode, which has no prompt to pre-fill.
.PARAMETER UseDeviceCode
    Device-code sign-in instead of the Windows broker.
.PARAMETER TimeoutMinutes
    How long to wait for the audit searches. 0 (the default) scales with the number of apps:
    60 minutes or 20 plus 6 per app, whichever is larger. Searches still running at the
    deadline keep running server side and are reused by a later run.
.PARAMETER OutDir
    Where the HTML report and CSV go. Defaults to the AppAccessPolicyMigration folder on the
    desktop, where the audit script writes its report.
.EXAMPLE
    .\Get-ExoAppMailboxActivity.ps1 -TenantId contoso.com
    Enumerates the apps, opens the picker, reports the selection over the last 30 days.
.EXAMPLE
    .\Get-ExoAppMailboxActivity.ps1 -TenantId contoso.com -All -Days 90 -UseDeviceCode
.LINK
    https://mikecrowley.us/2026/09/28/exchange-app-mailbox-activity/
.LINK
    https://mikecrowley.us/2026/07/10/exchange-app-access-policy-rbac-migration/
.NOTES
    Not available in GCC High / DoD (the audit search API is commercial-cloud only); use
    Search-UnifiedAuditLog -FreeText <appId> there.
    Author: Mike Crowley  https://mikecrowley.us
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$TenantId,
    [string[]]$AppId,
    [switch]$All,
    [ValidateRange(1, 180)][int]$Days = 30,
    [string]$UserPrincipalName,
    [switch]$UseDeviceCode,
    # Searches in flight at once. Purview returned 429 at the 11th concurrent submit in a large
    # tenant; five leaves room for whatever else the tenant's admins are running.
    [ValidateRange(1, 10)][int]$MaxConcurrent = 5,
    # Seconds between status checks on the running searches.
    [int]$PollSeconds = 30,
    # 0 = scale with the app count (see .PARAMETER TimeoutMinutes).
    [int]$TimeoutMinutes = 0,
    # At or below this many distinct mailboxes the app gets scope commands built from them.
    [int]$MaxScopeMailboxes = 25,
    # Per operation: mailboxes shown inline.
    [int]$MaxInline = 10,
    # Per operation: mailboxes in the collapsed full list. Beyond that, only the CSV has them.
    [int]$MaxCollapsed = 500,
    # Searches for the same app AND the same -Days window submitted by this script in the last day
    # are reused (they are still running or already done server side). A reused search's window
    # ends when it was submitted, so up to 24 hours of the newest activity can be missing; the
    # report marks those apps. Set this to always submit fresh ones.
    [switch]$NoReuse,
    # Records whose token belonged to a user rather than the app (TokenObjectId is not the app's own
    # service principal) came from a signed-in user working through the app (delegated access).
    # By default they are listed in their own muted rows under the
    # app's table and left out of the scope, because an RBAC scope constrains the app's own access,
    # not a user's. Set this to merge them into the app's lists and its scope. The CSV marks them
    # Delegated either way.
    [switch]$IncludeDelegated,
    [string]$OutDir = (Join-Path ([Environment]::GetFolderPath('Desktop')) 'AppAccessPolicyMigration')
)

$PostUrl = 'https://mikecrowley.us/2026/09/28/exchange-app-mailbox-activity/'                  # this script
$BlogUrl = 'https://mikecrowley.us/2026/07/10/exchange-app-access-policy-rbac-migration/'      # the migration
$AuditScriptUrl = 'https://github.com/Mike-Crowley/Public-Scripts/blob/main/Exchange/Audit-ExoAppAccessPolicies.ps1'
$clock = [System.Diagnostics.Stopwatch]::StartNew()

#region Sign-in
$params = @{
    TenantId     = $TenantId
    ContextScope = 'Process'
    # Least privilege: audit search for Exchange, app inventory, and basic user lookup (mailbox UPN -> id).
    Scopes       = @('AuditLogsQuery-Exchange.Read.All', 'Application.Read.All', 'User.ReadBasic.All')
    NoWelcome    = $true
}
if ($UseDeviceCode) {
    $params['UseDeviceCode'] = $true
    if ($UserPrincipalName) { Write-Warning '-UserPrincipalName is ignored with -UseDeviceCode: the device-code page has no account to pre-fill.' }
}
elseif ($UserPrincipalName) {
    # -LoginHint arrived in Microsoft.Graph.Authentication 2.39.0 (August 2026); older builds just prompt.
    if ((Get-Command Connect-MgGraph).Parameters.ContainsKey('LoginHint')) { $params['LoginHint'] = $UserPrincipalName }
    else { Write-Warning 'This Microsoft.Graph.Authentication build (before 2.39.0) has no -LoginHint; pick the account at the prompt.' }
}
# A cancelled or failed sign-in must stop here, or the searches below are submitted with no session.
# Graph credentials acquire tokens lazily, so Connect-MgGraph can "succeed" while the first real
# call fails; the probe below surfaces that here. With -UseDeviceCode the SDK sometimes builds the
# first call's credential with an empty token cache and throws "DeviceCodeCredential authentication
# failed: Object reference not set" (msgraph-sdk-powershell issue 3495, opened against 2.34, closed
# April 2026 with a fix in 2.37.0; the same error was reported on 2.37 afterwards and reproduced
# here on 2.40.0). The reporter saw it only when the account signing in was not the one logged in
# to Windows, which is the normal case for an admin working in a customer's tenant. A disconnect
# and a second sign-in usually clears it, so device-code sign-in gets one retry. Browser and broker
# sign-ins do not: the bug is in the device-code credential, and the same text from another flow
# is a different problem that a second prompt would not fix.
$attempt = 0
while ($true) {
    $attempt++
    try {
        Connect-MgGraph @params -ErrorAction Stop
    }
    catch {
        throw "Microsoft Graph sign-in failed: $_"
    }
    try {
        $null = Invoke-MgGraphRequest -Uri 'v1.0/me?$select=id' -Verbose:$false -ErrorAction Stop
        break
    }
    catch {
        if ($UseDeviceCode -and $attempt -lt 2 -and "$_" -match 'Object reference not set') {
            Write-Warning 'The Graph session signed in but cannot issue tokens (SDK device-code bug, issue 3495). Signing in once more.'
            Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
            continue
        }
        throw ("Microsoft Graph signed in but the first request failed: $_ " +
               $(if ($UseDeviceCode -and "$_" -match 'Object reference not set') { 'This is the SDK device-code token bug; rerun without -UseDeviceCode (browser or broker sign-in) if it persists.' } else { '' }))
    }
}
Write-Verbose "Signed in as $((Get-MgContext).Account) after $($clock.Elapsed.ToString('mm\:ss'))"
#endregion Sign-in

#region Pre-flight
# Fail here, with a reason, rather than ten minutes into a search that was never going to work.
$ctx = Get-MgContext
$RequiredScopes = @('AuditLogsQuery-Exchange.Read.All', 'Application.Read.All', 'User.ReadBasic.All')
# 1. Cloud. The Purview audit search API exists in the commercial cloud only (GCC included); GCC High,
#    DoD and 21Vianet have no endpoint for it, and Search-UnifiedAuditLog is the fallback there.
if ("$($ctx.Environment)" -ne 'Global') {
    throw "This session is connected to the '$($ctx.Environment)' cloud. The Purview audit search API is commercial-cloud only (GCC included); in GCC High, DoD or 21Vianet use Search-UnifiedAuditLog -FreeText <appId> in Exchange Online PowerShell instead."
}
# 2. Scopes. Consent can be declined, or need an admin the signed-in user is not; the token then lacks
#    the scope and every call fails with a 403 that says nothing about why.
$missingScopes = @($RequiredScopes | Where-Object { @($ctx.Scopes) -notcontains $_ })
if ($missingScopes.Count -gt 0) {
    throw "The Graph session lacks $($missingScopes -join ', '). These permissions need admin consent for 'Microsoft Graph Command Line Tools' in this tenant; sign in as an admin, or have one grant them, then rerun."
}
# 3. Audit search itself. A listing call is cheap and fails for tenants whose licensing has no Purview
#    Audit (some small-business and Exchange-only plans), tenants where unified audit logging was turned
#    off, and accounts without an audit role.
try {
    $null = Invoke-MgGraphRequest -Uri 'v1.0/security/auditLog/queries?$top=1' -Verbose:$false -ErrorAction Stop
}
catch {
    throw ("Purview audit search is not available to this session ($_). Likely causes, in order: the tenant's plans do not include Purview Audit; " +
           'unified audit logging is off (check with Get-AdminAuditLogConfig | Select-Object UnifiedAuditLogIngestionEnabled in Exchange Online PowerShell); ' +
           'or this account lacks an audit role (Audit Reader in Purview, or View-Only Audit Logs / Audit Logs in Exchange Online).')
}
# 4. Picker. Out-GridView is Windows-only; elsewhere the apps have to be named.
$onWindows = ($PSVersionTable.PSEdition -ne 'Core') -or [bool]$IsWindows
if (-not $onWindows -and -not $All -and -not $AppId) {
    throw 'Out-GridView is not available on this platform. Pass -AppId (one or more) or -All.'
}
Write-Verbose 'Pre-flight passed: commercial cloud, scopes granted, audit search reachable'
#endregion Pre-flight

#region Helpers
function EscSq { param([string]$s) if ($null -eq $s) { '' } else { $s.Replace("'", "''") } }
function HtmlEnc { param([string]$s) [System.Net.WebUtility]::HtmlEncode("$s") }

# Graph call with 429 handling. The SDK retries three times on its own and then throws; when that
# happens the tenant is saturated and the right move is to wait a full minute, not hammer it.
function Invoke-GraphCall {
    param([string]$Method = 'GET', [string]$Uri, [string]$Body, [int]$Attempts = 5)
    for ($try = 1; $try -le $Attempts; $try++) {
        try {
            if ($Method -eq 'POST') { return Invoke-MgGraphRequest -Method POST -Uri $Uri -Body $Body -ContentType 'application/json' -Verbose:$false -ErrorAction Stop }
            return Invoke-MgGraphRequest -Uri $Uri -Verbose:$false -ErrorAction Stop
        }
        catch {
            if ("$_" -notmatch 'TooManyRequests|\b429\b' -or $try -eq $Attempts) { throw }
            Write-Verbose "Throttled on $Method $Uri; waiting 60 s (attempt $try of $Attempts)"
            Start-Sleep -Seconds 60
        }
    }
}
function Get-GraphPages {
    param([string]$Uri)
    $out = @()
    while ($Uri) {
        $page = Invoke-GraphCall -Uri $Uri
        $out += @($page.value)
        $Uri = $page.'@odata.nextLink'
    }
    $out
}

# Permission -> RBAC role, per the Supported Application Roles table at
# https://learn.microsoft.com/en-us/exchange/permissions-exo/application-rbac
$GraphAppId = '00000003-0000-0000-c000-000000000000'
$ExoAppId = '00000002-0000-0ff1-ce00-000000000000'
$GraphRoleMap = @{
    'Mail.Read'                    = 'Application Mail.Read'
    'Mail.ReadBasic'               = 'Application Mail.ReadBasic'
    'Mail.ReadBasic.All'           = 'Application Mail.ReadBasic'
    'Mail.ReadWrite'               = 'Application Mail.ReadWrite'
    'Mail.Send'                    = 'Application Mail.Send'
    'MailboxSettings.Read'         = 'Application MailboxSettings.Read'
    'MailboxSettings.ReadWrite'    = 'Application MailboxSettings.ReadWrite'
    'Calendars.Read'               = 'Application Calendars.Read'
    'Calendars.ReadWrite'          = 'Application Calendars.ReadWrite'
    'Contacts.Read'                = 'Application Contacts.Read'
    'Contacts.ReadWrite'           = 'Application Contacts.ReadWrite'
    'MailboxFolder.Read.All'       = 'Application MailboxFolder.Read'
    'MailboxFolder.ReadWrite.All'  = 'Application MailboxFolder.ReadWrite'
    'MailboxItem.Read.All'         = 'Application MailboxItem.Read'
    'MailboxItem.ReadWrite.All'    = 'Application MailboxItem.ReadWrite'
    'MailboxItem.Export.All'       = 'Application MailboxItem.Export'
    'MailboxItem.ImportExport.All' = 'Application MailboxItem.ImportExport'
    'MailboxConfigItem.Read'       = 'Application MailboxConfigItem.Read'
    'MailboxConfigItem.ReadWrite'  = 'Application MailboxConfigItem.ReadWrite'
    'MailTips.ReadBasic.All'       = 'Application MailTips.ReadBasic.All'
    'Mail-Advanced.ReadWrite.All'  = 'Application Mail-Advanced.ReadWrite.All'   # keeps its .All, unlike the rest
}
# The table also lists two composite roles (Application Mail Full Access, Application Exchange Full
# Access) that bundle several permissions; this map stays one permission to one role.
$ExoRoleMap = @{
    'full_access_as_app' = 'Application EWS.AccessAsApp'
    'SMTP.SendAsApp'     = 'Application SMTP.SendAsApp'
}
$ExoRelevant = @('full_access_as_app', 'IMAP.AccessAsApp', 'POP.AccessAsApp', 'SMTP.SendAsApp')
$MsTenantIds = @('f8cdef31-a31e-4b4a-93e4-5f571e91255a', '72f988bf-86f1-41af-91ab-2d7cd011db47')

# Audit operation -> the kind of access it proves. Used to say which grants an app actually
# exercises. Send is unambiguous; reads and writes could be mail, calendar or contacts, since
# the record's folder path is the only clue and it is not reliable enough to act on.
function Get-OperationKind {
    param([string]$Operation)
    switch -Regex ($Operation) {
        '^(Send|SendAs|SendOnBehalf)$' { 'send'; break }
        '^(MailItemsAccessed|FolderBind|MessageBind|SearchQueryInitiated|AttachmentAccess)$' { 'read'; break }
        '^(Create|Update|Move|MoveToDeletedItems|SoftDelete|HardDelete|Copy|UpdateInboxRules|UpdateFolderPermissions|UpdateCalendarDelegation|ApplyRecord|RecordDelete)$' { 'write'; break }
        default { 'other' }
    }
}
function Get-GrantKinds {
    # Which access kinds a grant covers, for matching against observed operations. The legacy
    # Exchange Online resource spells some of these with .All (Calendars.ReadWrite.All), hence the
    # optional suffix. IMAP/POP and anything unrecognized come back 'other': no mailbox audit record
    # this report can match proves them, so they are reported as not measured rather than used.
    param([string]$Perm)
    switch -Regex ($Perm) {
        '^Mail\.Send$|^SMTP\.SendAsApp$' { @('send'); break }
        '^Mail\.ReadBasic|^Mail\.Read$|^MailboxItem\.Read\.All$|^MailboxFolder\.Read\.All$|^Calendars\.Read(Basic)?(\.All)?$|^Contacts\.Read(\.All)?$|^MailboxSettings\.Read$|^MailboxConfigItem\.Read$|^MailTips' { @('read'); break }
        '^Mail\.ReadWrite(\.All)?$|^Mail-Advanced\.ReadWrite|^MailboxFolder\.ReadWrite|^MailboxItem\.(ReadWrite|Export|ImportExport)|^Calendars\.ReadWrite(\.All)?$|^Contacts\.ReadWrite(\.All)?$|^MailboxSettings\.ReadWrite$|^MailboxConfigItem\.ReadWrite$' { @('read', 'write'); break }
        '^full_access_as_app$' { @('read', 'write', 'send'); break }
        default { @('other') }
    }
}
function ConvertTo-SafeName { param([string]$s) ($s -replace '[^\w\- ]', '').Trim() -replace '\s+', ' ' }
#endregion Helpers

#region Which apps
# Every non-Microsoft service principal holding Exchange application permissions, the same
# way the audit script finds them: appRoleAssignedTo on the Microsoft Graph and Exchange
# Online resource principals, mailbox-data roles only.
Write-Host 'Enumerating service principals with Exchange application permissions...'
$Apps = @{}   # appId -> @{ AppId; Name; SpObjectId; Grants = [perm display]; Roles = [rbac role]; Enabled; IsHomeTenant }
$holders = @{}  # sp object id -> @{ Perms = [] }
foreach ($resAppId in @($GraphAppId, $ExoAppId)) {
    $resSp = @(Get-GraphPages "v1.0/servicePrincipals?`$filter=appId eq '$resAppId'&`$select=id,appRoles")
    if ($resSp.Count -eq 0) { continue }
    $roleNameById = @{}
    foreach ($ar in $resSp[0].appRoles) { $roleNameById["$($ar.id)"] = "$($ar.value)" }
    $prefix = if ($resAppId -eq $GraphAppId) { 'Graph' } else { 'EXO' }
    foreach ($a in Get-GraphPages "v1.0/servicePrincipals/$($resSp[0].id)/appRoleAssignedTo?`$top=999") {
        if ("$($a.principalType)" -ne 'ServicePrincipal') { continue }
        $perm = $roleNameById["$($a.appRoleId)"]
        if (-not $perm) { continue }
        $relevant = if ($resAppId -eq $GraphAppId) { $GraphRoleMap.ContainsKey($perm) -or $perm -eq 'Calendars.ReadBasic' }
                    else { $ExoRelevant -contains $perm -or $perm -match '^(Mail|Calendars|Contacts|MailboxSettings)\.' }
        if (-not $relevant) { continue }
        $pid_ = "$($a.principalId)"
        if (-not $holders.ContainsKey($pid_)) { $holders[$pid_] = @{ Name = "$($a.principalDisplayName)"; Perms = @() } }
        $holders[$pid_].Perms += "$prefix`:$perm"
    }
}
foreach ($spId in @($holders.Keys)) {
    $sp = Invoke-GraphCall -Uri "v1.0/servicePrincipals/$spId`?`$select=appId,displayName,appOwnerOrganizationId,accountEnabled"
    if ($MsTenantIds -contains "$($sp.appOwnerOrganizationId)") { continue }
    $perms = @($holders[$spId].Perms | Select-Object -Unique)
    $roles = @($perms | ForEach-Object { $bare = $_ -replace '^(Graph|EXO):', ''; if ($_ -like 'Graph:*') { $GraphRoleMap[$bare] } else { $ExoRoleMap[$bare] } } | Where-Object { $_ } | Select-Object -Unique)
    $Apps["$($sp.appId)"] = @{
        AppId        = "$($sp.appId)"
        Name         = $(if ($sp.displayName) { "$($sp.displayName)".Trim() } else { $holders[$spId].Name })
        SpObjectId   = $spId
        Grants       = $perms
        Roles        = $roles
        Enabled      = $sp.accountEnabled
        IsHomeTenant = ("$($sp.appOwnerOrganizationId)" -eq "$((Get-MgContext).TenantId)")
    }
}
Write-Host "  $($Apps.Count) apps hold Exchange application permissions"

# Which of them to analyze: explicit ids (looked up if they hold no mail grant, e.g. an app
# already moved to RBAC), everything, or a picker.
$Selected = @()
if ($AppId) {
    # Graph returns app ids in lower case; normalize so the search names (and their reuse) match.
    foreach ($id in ($AppId | ForEach-Object { "$_".Trim().ToLowerInvariant() } | Select-Object -Unique)) {
        if ($Apps.ContainsKey($id)) { $Selected += $Apps[$id]; continue }
        $sp = @(Get-GraphPages "v1.0/servicePrincipals?`$filter=appId eq '$id'&`$select=id,appId,displayName,appOwnerOrganizationId,accountEnabled")
        $Selected += @{
            AppId = $id; Name = $(if ($sp.Count -gt 0 -and $sp[0].displayName) { "$($sp[0].displayName)" } else { $id })
            SpObjectId = $(if ($sp.Count -gt 0) { "$($sp[0].id)" } else { '' }); Grants = @(); Roles = @()
            Enabled = $(if ($sp.Count -gt 0) { $sp[0].accountEnabled } else { $null })
            IsHomeTenant = $(if ($sp.Count -gt 0) { "$($sp[0].appOwnerOrganizationId)" -eq "$((Get-MgContext).TenantId)" } else { $false })
        }
    }
}
elseif ($All) { $Selected = @($Apps.Values) }
else {
    $picked = @($Apps.Values | Sort-Object { $_.Name } | ForEach-Object { [pscustomobject]@{ Name = $_.Name; AppId = $_.AppId; Grants = ($_.Grants -join ', '); Enabled = $_.Enabled } } | Out-GridView -Title "Apps with Exchange application permissions - select the ones to analyze ($Days days)" -PassThru)
    if ($picked.Count -eq 0) { throw 'Nothing selected. Pass -AppId, use -All, or pick at least one app.' }
    $Selected = @($picked | ForEach-Object { $Apps[$_.AppId] })
}
$Selected = @($Selected | Sort-Object { $_.Name })
Write-Host "  Analyzing $($Selected.Count) app(s) over the last $Days days"
# Searches come back one every few minutes, not all at once: four apps took 32 minutes, and a
# 23-app run still had half its searches queued at 45. The default deadline scales with the count.
if ($TimeoutMinutes -le 0) { $TimeoutMinutes = [Math]::Max(60, 20 + 6 * $Selected.Count) }
Write-Verbose "Deadline for the audit searches: $TimeoutMinutes minutes"
#endregion Which apps

#region Audit searches (sliding window)
$namePrefix = 'app-activity'
$start = (Get-Date).ToUniversalTime().AddDays(-$Days).ToString('o')
$end = (Get-Date).ToUniversalTime().ToString('o')
$queries = @{}   # appId -> query id
$todo = [System.Collections.Generic.Queue[string]]::new()
$Selected | ForEach-Object { $todo.Enqueue($_.AppId) }

# Reuse this script's own recent searches for the same apps and the same window: a throttled or
# cancelled earlier run leaves them running server side, and resubmitting only deepens the queue.
# The window is part of the search name, so a 3-day probe is never mistaken for the 30-day run.
# A reused search's window ends when it was submitted, not now; the report marks those apps.
$ReusedAt = @{}   # appId -> submission time of the reused search
if (-not $NoReuse) {
    $recent = @{}
    foreach ($q in Get-GraphPages 'v1.0/security/auditLog/queries') {
        $m = [regex]::Match("$($q.displayName)", "^$namePrefix ([0-9a-f-]{36}) (\d+)d (\d{8}-\d{4})$", 'IgnoreCase')
        if (-not $m.Success -or "$($q.status)" -in 'failed', 'cancelled') { continue }
        if ([int]$m.Groups[2].Value -ne $Days) { continue }
        $stamp = [datetime]::ParseExact($m.Groups[3].Value, 'yyyyMMdd-HHmm', $null)
        if ($stamp -lt (Get-Date).AddHours(-24)) { continue }
        $id = $m.Groups[1].Value.ToLowerInvariant()
        if (-not $recent.ContainsKey($id) -or $stamp -gt $recent[$id].Stamp) { $recent[$id] = @{ QueryId = "$($q.id)"; Stamp = $stamp } }
    }
    $fresh = [System.Collections.Generic.Queue[string]]::new()
    foreach ($id in $todo) {
        if ($recent.ContainsKey($id)) {
            $queries[$id] = $recent[$id].QueryId
            $ReusedAt[$id] = $recent[$id].Stamp
            Write-Host "Reusing $id -> query $($recent[$id].QueryId) (submitted $($recent[$id].Stamp.ToString('HH:mm')); its window ends then)"
        }
        else { $fresh.Enqueue($id) }
    }
    $todo = $fresh
}

$active = [System.Collections.Generic.HashSet[string]]::new([string[]]@($queries.Keys))
$Results = @{}   # appId -> @{ Hits; Truncated; Failed }
$deadline = (Get-Date).AddMinutes($TimeoutMinutes)
function Submit-Next {
    while ($active.Count -lt $MaxConcurrent -and $todo.Count -gt 0) {
        $id = $todo.Peek()
        $body = @{
            '@odata.type'       = '#microsoft.graph.security.auditLogQuery'
            displayName         = "$namePrefix $id ${Days}d $(Get-Date -Format yyyyMMdd-HHmm)"
            filterStartDateTime = $start
            filterEndDateTime   = $end
            keywordFilter       = $id
            serviceFilter       = 'Exchange'
            recordTypeFilters   = @('exchangeItem', 'exchangeItemGroup', 'exchangeItemAggregated')
        } | ConvertTo-Json
        $q = $null
        # No in-call retry on submit: a 429 here means the window is full server side, and the loop
        # will come back to this app once a running search finishes.
        try { $q = Invoke-GraphCall -Method POST -Uri 'v1.0/security/auditLog/queries' -Body $body -Attempts 1 }
        catch {
            if ("$_" -match 'TooManyRequests|\b429\b') { Write-Warning "Submit for $id throttled; will retry after the window drains."; return }
            [void]$todo.Dequeue(); Write-Warning "Submit for $id failed: $_"; $Results[$id] = @{ Hits = @(); Truncated = $false; Failed = "$_" }; continue
        }
        [void]$todo.Dequeue()
        if (-not $q.id) { Write-Warning "Submit for $id returned no query id; skipped."; $Results[$id] = @{ Hits = @(); Truncated = $false; Failed = 'no query id' }; continue }
        $queries[$id] = "$($q.id)"
        [void]$active.Add($id)
        Write-Host "Submitted $id -> query $($q.id)"
    }
}
Submit-Next
Write-Verbose "$($active.Count) searches in flight, $($todo.Count) queued, after $($clock.Elapsed.ToString('mm\:ss'))"
while ($active.Count -gt 0 -or $todo.Count -gt 0) {
    if ((Get-Date) -gt $deadline) {
        Write-Warning "Gave up after $TimeoutMinutes min. Still running: $(@($active) -join ', '). Never submitted: $(@($todo) -join ', '). The running searches keep going server side; a rerun within 24 hours with the same -Days reuses every one that has finished instead of resubmitting, or pass -TimeoutMinutes to wait longer."
        foreach ($id in @($active) + @($todo)) { $Results[$id] = @{ Hits = @(); Truncated = $false; Failed = 'timed out' } }
        break
    }
    if ($active.Count -eq 0) { Submit-Next; if ($active.Count -eq 0) { Start-Sleep -Seconds $PollSeconds }; continue }
    Start-Sleep -Seconds $PollSeconds
    foreach ($id in @($active)) {
        # beta, because only beta reports whether the search stopped at its per-search record limit.
        # A truncated query still says 'succeeded', so the flag is the only reliable signal.
        try { $q = Invoke-GraphCall -Uri "beta/security/auditLog/queries/$($queries[$id])" } catch { Write-Verbose "Status check for $id failed, will retry: $_"; continue }
        $state = "$($q.status)"
        if ($state -notin 'succeeded', 'failed', 'cancelled') { continue }
        [void]$active.Remove($id)
        Write-Host "$id : $state after $($clock.Elapsed.ToString('mm\:ss'))"
        if ($state -ne 'succeeded') { Write-Warning "Search for $id ended $state."; $Results[$id] = @{ Hits = @(); Truncated = $false; Failed = $state }; continue }
        $records = @(Get-GraphPages "v1.0/security/auditLog/queries/$($queries[$id])/records?`$top=1000")
        $hits = @($records | Where-Object {
            $d = $_.auditData
            # The structured fields first; then the "[AppId=<id>]" tag that ActorInfoString and
            # ClientInfoString carry for some clients, which names the actor's own app.
            "$($d.AppId)" -eq $id -or "$($d.ClientAppId)" -eq $id -or "$($d.AppAccessContext.ClientAppId)" -eq $id -or
            "$($d.ActorInfoString) $($d.ClientInfoString)" -match "\[AppId=$([regex]::Escape($id))\]"
        })
        Write-Host "$id : $($records.Count) records matched the keyword, $($hits.Count) name the app"
        $Results[$id] = @{ Hits = $hits; Truncated = [bool]$q.isRecordCountLimitExceeded; Failed = '' }
        if ($q.isRecordCountLimitExceeded) { Write-Warning "Search for $id hit the per-search record limit ($($q.recordCountLimit)); its lists are INCOMPLETE. Rerun this app with a smaller -Days." }
    }
    Submit-Next
}
Write-Verbose "Searches done at $($clock.Elapsed.ToString('mm\:ss'))"
#endregion Audit searches

#region Roll-up and scope commands
$userIdCache = @{}
function Get-MailboxObjectId {
    param([string]$Upn)
    if ($userIdCache.ContainsKey($Upn)) { return $userIdCache[$Upn] }
    $oid = ''
    try {
        $u = @(Get-GraphPages "v1.0/users?`$filter=userPrincipalName eq '$(EscSq $Upn)' or mail eq '$(EscSq $Upn)'&`$select=id")
        if ($u.Count -gt 0) { $oid = "$($u[0].id)" }
    }
    catch { $oid = '' }
    $userIdCache[$Upn] = $oid
    $oid
}

# One row per operation, with its mailboxes ordered by event count (name breaks ties).
function Get-OpsFromRecords {
    param([object[]]$Records)
    $rows = foreach ($g in (@($Records) | Group-Object { "$($_.operation)" } | Sort-Object Count -Descending)) {
        $byMbx = @($g.Group | Group-Object { "$($_.auditData.MailboxOwnerUPN)".ToLowerInvariant() } | Sort-Object @{ Expression = 'Count'; Descending = $true }, Name)
        $stamps = @($g.Group | ForEach-Object { [datetime]$_.createdDateTime })
        @{
            Operation = $g.Name
            Kind      = Get-OperationKind $g.Name
            Events    = $g.Count
            Mailboxes = @($byMbx | ForEach-Object { @{ Mailbox = $_.Name; Events = $_.Count } })
            FirstSeen = ($stamps | Measure-Object -Minimum).Minimum
            LastSeen  = ($stamps | Measure-Object -Maximum).Maximum
        }
    }
    @($rows)
}

# Delegated records (a signed-in user working through the app) carry the app id in the same fields
# as app-only ones. What separates them is who held the token: Exchange writes the token's oid claim,
# "the verified identity of the user or service principal", to TokenObjectId
# (https://learn.microsoft.com/en-us/entra/identity/authentication/how-to-authentication-track-linkable-identifiers).
# An app-only token carries the app's own service principal; a delegated one carries the user.
# Session ids do not separate them. Until 2026-09-29 this script took AppAccessContext.AADSessionId
# as the delegated marker, but a probe on 2026-09-28 found app-only EWS records (TokenType
# V1AppOnly) that carry one, and Microsoft's own clients fill SessionId and AADSessionId
# inconsistently. The same Entra page notes the identifiers "aren't available in the Exchange Online
# audit logs on some aggregated log entries, or logs generated from background processes": a record
# without TokenObjectId, or an app with no service principal in this tenant, counts as app-only. An
# app-only record classed as delegated drops a mailbox the app needs from its scope, which the
# cutover's one-mailbox check would not catch, so every doubt falls to app-only.
function Test-DelegatedRecord {
    param($Record, [string]$SpObjectId)
    $holder = "$($Record.auditData.TokenObjectId)"
    [bool]($holder -and $SpObjectId -and $holder -ne $SpObjectId)
}

$Report = foreach ($app in $Selected) {
    $r = $Results[$app.AppId]
    if (-not $r) { $r = @{ Hits = @(); Truncated = $false; Failed = 'not run' } }
    $appSpId = "$($app.SpObjectId)"
    $allHits = @($r.Hits)
    $delegated = @($allHits | Where-Object { Test-DelegatedRecord $_ $appSpId })
    $hits = if ($IncludeDelegated) { $allHits } else { @($allHits | Where-Object { -not (Test-DelegatedRecord $_ $appSpId) }) }
    $ops = Get-OpsFromRecords $hits
    # Excluded delegated records stay visible (their own rows under the table) so a misclassified
    # record can be spotted; with -IncludeDelegated they are already merged into $ops.
    $delegatedOps = if ($IncludeDelegated) { @() } else { Get-OpsFromRecords $delegated }
    # The CSV ledger is built from every record, whatever the switch, with the access type on
    # each row: the page applies the delegated split, the CSV shows what was split.
    $ledger = foreach ($g in ($allHits | Group-Object { "$($_.operation)|$("$($_.auditData.MailboxOwnerUPN)".ToLowerInvariant())|$(if (Test-DelegatedRecord $_ $appSpId) { 'Delegated' } else { 'AppOnly' })" })) {
        $parts = $g.Name -split '\|', 3
        [pscustomobject]@{ Operation = $parts[0]; Mailbox = $parts[1]; Access = $parts[2]; Events = $g.Count; UsedForScope = ($parts[2] -eq 'AppOnly' -or [bool]$IncludeDelegated) }
    }
    $ledger = @($ledger | Sort-Object Operation, @{ Expression = 'Events'; Descending = $true }, Mailbox)
    $allMbx = @($hits | Group-Object { "$($_.auditData.MailboxOwnerUPN)".ToLowerInvariant() } | Sort-Object Count -Descending)
    $kinds = @($ops | ForEach-Object { $_.Kind } | Select-Object -Unique)
    $distinct = $allMbx.Count
    $shape = if ($r.Failed) { 'Not analyzed' }
             elseif ($distinct -eq 0) { 'No activity' }
             elseif ($distinct -eq 1) { 'Single mailbox' }
             elseif ($distinct -le $MaxScopeMailboxes) { 'Scope candidate' }
             else { 'Broad' }

    # Grants: exercised (an observed operation kind matches), unproven, or idle in the window.
    # A mailbox audit record says read or write, not which folder type, so a read or a write
    # exercises Mail.ReadWrite, Calendars.ReadWrite and Contacts.ReadWrite alike. Send proves
    # Mail.Send, and reads or writes prove a Mail.* grant (the mail folders are where almost all of
    # that traffic lands); a calendar, contacts or settings grant matched only through a generic
    # read or write is marked unproven, not exercised, so it is not mistaken for confirmed.
    # MailTips calls leave no mailbox audit record at all, so a read can only ever make them unproven.
    # 'other' (IMAP/POP, unrecognized) is never matched against observed operations: an unrelated
    # uncategorized operation is not evidence, so those grants are reported as not measured.
    $grantRows = foreach ($perm in $app.Grants) {
        $bare = $perm -replace '^(Graph|EXO):', ''
        $gk = Get-GrantKinds $bare
        $used = @($kinds | Where-Object { $_ -ne 'other' -and $gk -contains $_ })
        $role = if ($perm -like 'Graph:*') { $GraphRoleMap[$bare] } else { $ExoRoleMap[$bare] }
        $status = if ($gk -contains 'other') { 'unmeasured' }
                  elseif ($used.Count -eq 0) { 'idle' }
                  elseif ($bare -match '^(Calendars|Contacts|MailboxSettings|MailboxConfigItem|MailTips)\.') { 'unproven' }
                  else { 'used' }
        @{ Perm = $perm; Role = $role; Used = ($status -in 'used', 'unproven'); Status = $status; Kinds = $gk }
    }
    $grantRows = @($grantRows)

    # Commands: built from the mailboxes observed. Everything here is additive (steps 1-4 of the
    # migration); the revoke of the tenant-wide grant stays with the audit script, whose report
    # generates a verified cutover for exactly this situation.
    # Object names come from the display name, which Entra does not keep unique, so each one carries
    # the first 8 characters of the app id: two apps both called "Postman" must never share a group or
    # a scope. The base is cut to 40 characters so the group name and alias stay under Exchange's 64.
    $safe = ConvertTo-SafeName $app.Name
    if (-not $safe) { $safe = 'App' }
    if ($safe.Length -gt 40) { $safe = $safe.Substring(0, 40).Trim() }
    $tag = $app.AppId.Substring(0, [Math]::Min(8, $app.AppId.Length))
    $scopeName = "$safe-$tag-Scope"
    $cmd = @()
    $cmd += "# ==== $($app.Name) ($($app.AppId)) - RBAC scope from observed activity, last $Days days ===="
    if ($r.Failed) {
        $cmd += "# The audit search for this app did not complete ($($r.Failed)); rerun before scoping."
    }
    elseif ($distinct -eq 0 -and $delegated.Count -gt 0 -and -not $IncludeDelegated) {
        $cmd += "# No app-only mailbox activity in $Days days, but $($delegated.Count) delegated record(s) show signed-in users"
        $cmd += "# working through the app. The app is in use; its APPLICATION grants are what is idle. Those are"
        $cmd += "# candidates for revoking (the audit script's report handles that), not for a scope; the users' access"
        $cmd += "# runs on delegated permissions, which an RBAC for Applications scope does not constrain."
        $cmd += "# Details: $PostUrl"
    }
    elseif ($distinct -eq 0) {
        $cmd += "# No mailbox activity in $Days days. Before scoping, ask whether the app is still in use: an idle app"
        $cmd += "# is a candidate for revoking its grants outright (the audit script's report handles that), not for a scope."
        $cmd += "# If it is seasonal or new, rerun with a longer -Days (up to 180) before deciding. It can also look"
        $cmd += "# idle when the mailboxes it uses do not audit MailItemsAccessed and Send: a mailbox audit list"
        $cmd += "# customized years ago lacks both (the report footer shows how to check and fix that)."
        $cmd += "# Details: $PostUrl"
    }
    elseif ($distinct -gt $MaxScopeMailboxes) {
        $cmd += "# Observed on $distinct mailboxes ($(($ops | ForEach-Object { "$($_.Operation): $($_.Mailboxes.Count)" }) -join '; '))."
        $cmd += "# A scope built from a list this size is not a control. Decide about the app instead: is org-wide"
        $cmd += "# reach its job (backup, security, journaling), or has it drifted? If a subset is legitimate, build"
        $cmd += "# the group from a source of truth (department, license, OU), not from this list. Full list: CSV."
        $cmd += "# Details: $PostUrl"
    }
    else {
        $upns = @($allMbx | ForEach-Object { $_.Name } | Where-Object { $_ -match '@' })
        $cmd += "# Observed: " + (($ops | ForEach-Object { "$($_.Operation) on $($_.Mailboxes.Count) mailbox(es)" }) -join '; ')
        $cmd += "# Steps 1-4 are additive. Verify, then revoke the tenant-wide grant with the audit script's generated"
        $cmd += "# cutover: $AuditScriptUrl"
        $cmd += "# How this report works and its caveats: $PostUrl"
        $cmd += "# The migration itself (scopes, roles, cutover): $BlogUrl"
        $cmd += ''
        $cmd += '# 1. Exchange pointer to the Entra service principal (idempotent)'
        $spOid = if ($app.SpObjectId) { $app.SpObjectId } else { '<enterprise app object id>' }
        $cmd += "if (-not (Get-ServicePrincipal -Identity '$($app.AppId)' -ErrorAction SilentlyContinue)) {"
        $cmd += "    New-ServicePrincipal -AppId '$($app.AppId)' -ObjectId '$spOid' -DisplayName '$(EscSq $app.Name)'"
        $cmd += '}'
        $cmd += ''
        if ($distinct -eq 1) {
            $oid = Get-MailboxObjectId $upns[0]
            $cmd += "# 2. Scope: the one mailbox observed ($($upns[0]))"
            $cmd += "New-ManagementScope -Name '$(EscSq $scopeName)' -RecipientRestrictionFilter `"ExternalDirectoryObjectId -eq '$(if ($oid) { $oid } else { '<object id of ' + $upns[0] + '>' })'`""
        }
        else {
            $groupName = "$safe $tag Mailboxes"
            $alias = ($safe -replace '[^\w]', '') + $tag + 'Mailboxes'
            $cmd += "# 2. Scope: a mail-enabled security group holding the $($upns.Count) mailboxes observed (DIRECT members only;"
            $cmd += '#    membership is the control from here on, so own it like one)'
            $cmd += "New-DistributionGroup -Name '$(EscSq $groupName)' -Alias '$alias' -Type Security -Members $(($upns | ForEach-Object { "'$(EscSq $_)'" }) -join ', ')"
            $cmd += "`$dn = (Get-DistributionGroup -Identity '$(EscSq $groupName)').DistinguishedName"
            $cmd += "New-ManagementScope -Name '$(EscSq $scopeName)' -RecipientRestrictionFilter `"MemberOfGroup -eq '`$dn'`""
        }
        $cmd += ''
        $cmd += '# 3. Role assignments: one per RBAC role the grants map to. Grants with no matching activity in the'
        $cmd += '#    window are commented out - add them back only if the app needs them, or drop the grant instead.'
        $cmd += '#    "unproven" means reads or writes were seen but the record cannot say they were calendar, contacts'
        $cmd += '#    or settings rather than mail; keep the role if the app needs it, drop the grant if its job is mail.'
        $seenRoles = @{}
        foreach ($gr in $grantRows) {
            if (-not $gr.Role -or $seenRoles.ContainsKey($gr.Role)) { continue }
            $seenRoles[$gr.Role] = $true
            $line = "New-ManagementRoleAssignment -App '$($app.AppId)' -Role '$($gr.Role)' -CustomResourceScope '$(EscSq $scopeName)'"
            $cmd += switch ($gr.Status) {
                'used' { $line }
                'unproven' { "$line   # $($gr.Perm): unproven" }
                default { "# $line   # $($gr.Perm): no matching activity observed" }
            }
        }
        $noRole = @($grantRows | Where-Object { -not $_.Role } | ForEach-Object { $_.Perm })
        if ($noRole.Count -gt 0) {
            $imapNote = if (($noRole -join ' ') -match '(IMAP|POP)\.AccessAsApp') { ' IMAP/POP reach is per mailbox via Add-MailboxPermission.' } else { '' }
            $cmd += "# No RBAC role exists for $($noRole -join ', '): nothing to scope them to; they stay tenant-wide until revoked.$imapNote"
        }
        if ($grantRows.Count -eq 0) { $cmd += "# (no Exchange grants found in Entra for this app - it may already be on RBAC; see Get-ManagementRoleAssignment)" }
        $cmd += ''
        $cmd += '# 4. Verify: expect InScope = True, then test the app itself'
        $cmd += "Test-ServicePrincipalAuthorization -Identity '$($app.AppId)' -Resource '$(EscSq $upns[0])' | Format-Table"
    }

    [pscustomobject]@{
        App        = $app
        Shape      = $shape
        Failed     = $r.Failed
        Truncated  = $r.Truncated
        Events     = $hits.Count
        Delegated  = $delegated.Count
        ReusedAt   = $ReusedAt[$app.AppId]
        Distinct   = $distinct
        Ops        = $ops
        DelegatedOps = $delegatedOps
        Ledger     = $ledger
        GrantRows  = $grantRows
        Commands   = ($cmd -join "`n")
        FirstSeen  = if ($hits.Count) { ($ops | ForEach-Object { $_.FirstSeen } | Measure-Object -Minimum).Minimum } else { $null }
        LastSeen   = if ($hits.Count) { ($ops | ForEach-Object { $_.LastSeen } | Measure-Object -Maximum).Maximum } else { $null }
    }
}
$Report = @($Report)
Write-Verbose "Roll-up done at $($clock.Elapsed.ToString('mm\:ss'))"
#endregion Roll-up and scope commands

#region Output
if (-not (Test-Path $OutDir)) { $null = New-Item -ItemType Directory -Path $OutDir }
$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$tenantTag = ($TenantId -replace '[^\w\-.]', '')   # display: keep the domain's dots
$fileTag = ($tenantTag -replace '[^\w\-]', '')      # file name: no dots
$HtmlPath = Join-Path $OutDir "AppMailboxActivity_$fileTag`_$stamp.html"
$CsvPath = Join-Path $OutDir "AppMailboxActivity_$fileTag`_$stamp.csv"

# CSV: every app/operation/mailbox/access row, so nothing is lost to the page's limits. Access says
# whether the records were app-only or delegated; UsedForScope says whether the page and the
# scope commands counted them (delegated rows only with -IncludeDelegated).
$csvRows = foreach ($item in $Report) {
    foreach ($row in $item.Ledger) {
        [pscustomobject]@{ AppName = $item.App.Name; AppId = $item.App.AppId; Access = $row.Access; Operation = $row.Operation; Mailbox = $row.Mailbox; Events = $row.Events; UsedForScope = $row.UsedForScope }
    }
}
@($csvRows) | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8

$shapeBadge = @{
    'No activity'    = "<span class='status-badge delete'>No activity</span>"
    'Single mailbox' = "<span class='status-badge ready'>&#10003; Single mailbox</span>"
    'Scope candidate' = "<span class='status-badge ready'>&#10003; Scope candidate</span>"
    'Broad'          = "<span class='status-badge unconstrained'>&#9888; Broad</span>"
    'Not analyzed'   = "<span class='status-badge error'>&#8252; Not analyzed</span>"
}
# One table row for an operation: top $MaxInline mailboxes inline, up to $MaxCollapsed behind a
# disclosure, the rest in the CSV. -Delegated renders the muted, labelled variant for records a
# signed-in user made through the app, which are shown but not counted toward the scope.
function Get-OpRowHtml {
    param($Op, [switch]$Delegated)
    $n = $Op.Mailboxes.Count
    $inline = (@($Op.Mailboxes | Select-Object -First $MaxInline | ForEach-Object { "$(HtmlEnc $_.Mailbox) <span class='email'>($($_.Events))</span>" }) -join '<br>')
    $more = ''
    if ($n -gt $MaxInline) {
        $rest = @($Op.Mailboxes | Select-Object -First $MaxCollapsed | ForEach-Object { "$($_.Mailbox)  ($($_.Events))" }) -join "`n"
        $tail = if ($n -gt $MaxCollapsed) { " <span class='cmd-meta'>first $MaxCollapsed shown; all $n in the CSV</span>" } else { '' }
        $more = "<details class='cmd' style='margin-top:0.4rem'><summary>All $n mailboxes$tail</summary><pre class='commands'>$(HtmlEnc $rest)</pre></details>"
    }
    $kindBadge = switch ($Op.Kind) { 'send' { "<span class='badge warn'>send</span>" } 'read' { "<span class='badge info'>read</span>" } 'write' { "<span class='badge purp'>write</span>" } default { "<span class='badge neut'>other</span>" } }
    $dates = "$($Op.FirstSeen.ToString('yyyy-MM-dd')) to $($Op.LastSeen.ToString('yyyy-MM-dd'))"
    if ($Delegated) {
        return "<tr class='delegated-row'><td><strong>$(HtmlEnc $Op.Operation)</strong>$kindBadge<span class='badge neut' title='A signed-in user working through the app (the token belonged to a user, not to the app itself). Not counted toward the scope.'>delegated</span><br><span class='email'>$dates &middot; not in scope</span></td><td>$($Op.Events)</td><td>$n</td><td>$inline$more</td></tr>"
    }
    "<tr><td><strong>$(HtmlEnc $Op.Operation)</strong>$kindBadge<br><span class='email'>$dates</span></td><td>$($Op.Events)</td><td>$n</td><td>$inline$more</td></tr>"
}

$AppSections = foreach ($item in $Report) {
    $app = $item.App
    $links = @()
    if ($app.IsHomeTenant -and $app.AppId) { $links += "<a class='plink' href='https://entra.microsoft.com/#view/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/~/Overview/appId/$($app.AppId)' target='_blank' rel='noopener'>App registration &#8599;</a>" }
    if ($app.SpObjectId) { $links += "<a class='plink' href='https://entra.microsoft.com/#view/Microsoft_AAD_IAM/ManagedAppMenuBlade/~/Overview/objectId/$($app.SpObjectId)/appId/$($app.AppId)' target='_blank' rel='noopener'>Enterprise app &#8599;</a>" }
    $grantChips = if ($item.GrantRows.Count -gt 0) {
        (@($item.GrantRows | ForEach-Object {
            $gr = $_   # switch rebinds $_ to the value under test, so hold the row first
            switch ($gr.Status) {
                'used' { "<span class='perm migrate' title='Activity of this kind observed'>$(HtmlEnc $gr.Perm)</span>" }
                'unproven' { "<span class='perm migrate' title='Reads or writes were observed, but the record does not say whether they were mail, calendar, contacts or settings'>$(HtmlEnc $gr.Perm) &#183; unproven</span>" }
                'unmeasured' { "<span class='perm keep' title='No mailbox audit record this report can match proves or disproves this kind of access (IMAP, POP, or an unrecognized permission)'>$(HtmlEnc $gr.Perm) &#183; not measured</span>" }
                default { "<span class='perm keep' title='No matching activity in the window'>$(HtmlEnc $gr.Perm) &#183; idle</span>" }
            }
        }) -join '')
    } else { "<span class='none'>No Exchange grants in Entra (already on RBAC, or none)</span>" }
    $flags = ''
    if ($item.Truncated) { $flags += " <span class='badge crit' title='The audit search stopped at its record limit'>truncated</span>" }
    if ($item.Delegated -gt 0) {
        $flags += if ($IncludeDelegated) { " <span class='badge purp' title='Records whose token belonged to a user rather than the app: a signed-in user working through the app. Merged into the lists and the scope because -IncludeDelegated was set.'>$($item.Delegated) delegated, included</span>" }
                  else { " <span class='badge purp' title='Records whose token belonged to a user rather than the app: a signed-in user working through the app. Listed in their own rows under the table and left out of the scope; rerun with -IncludeDelegated to merge them.'>$($item.Delegated) delegated, excluded</span>" }
    }
    if ($app.Enabled -eq $false) { $flags += " <span class='badge neut'>sign-in disabled</span>" }
    if ($item.ReusedAt) { $flags += " <span class='badge neut' title='This app&#39;s audit search was submitted at $($item.ReusedAt.ToString('yyyy-MM-dd HH:mm')) and reused, so its window ends then: activity after that time is not included. Rerun with -NoReuse for a window that ends now.'>window ends $($item.ReusedAt.ToString('MMM d HH:mm'))</span>" }
    $opsRows = @()
    if ($item.Ops.Count -eq 0) {
        $msg = if ($item.Failed) { "Search did not complete: $(HtmlEnc $item.Failed)" }
               elseif ($item.Delegated -gt 0 -and -not $IncludeDelegated) { "No app-only mailbox activity in the last $Days days. $($item.Delegated) delegated record(s) show signed-in users working through the app, so the app is in use; its application grants are what is idle." }
               else { "No mailbox activity attributed to this app in the last $Days days." }
        $opsRows += "<tr><td colspan='4' class='empty-state'>$msg</td></tr>"
    }
    foreach ($op in $item.Ops) { $opsRows += Get-OpRowHtml $op }
    # Delegated rows last, muted and labelled, so the split is visible and can be second-guessed.
    foreach ($op in $item.DelegatedOps) { $opsRows += Get-OpRowHtml $op -Delegated }
    $lineCount = @($item.Commands -split "`n").Count
    @"
        <h2>$(HtmlEnc $app.Name) $($shapeBadge[$item.Shape])$flags</h2>
        <p class="section-sub"><span class="app-id">$(HtmlEnc $app.AppId)</span> &nbsp; $($links -join ' ') <br>
        Entra grants: $grantChips</p>
        <table>
            <thead><tr><th>Operation</th><th>Events</th><th>Mailboxes</th><th>Mailboxes by activity (top $MaxInline inline)</th></tr></thead>
            <tbody>
                $($opsRows -join "`n")
            </tbody>
        </table>
        <div style="margin:0.6rem 0 0 0"><details class='cmd'><summary>PowerShell <span class='cmd-meta'>$lineCount lines</span><button type='button' class='copy-btn'>Copy</button></summary><pre class='commands'>$(HtmlEnc $item.Commands)</pre></details></div>
"@
}

$countBy = { param($s) @($Report | Where-Object { $_.Shape -eq $s }).Count }
$DelegatedNote = if ($IncludeDelegated) { 'this run was made with -IncludeDelegated, so they are merged into the app''s lists and its scope; the CSV still marks them Delegated.' }
                 else { 'they are counted on the app''s heading, listed in muted rows marked <em>delegated</em> under its table, marked Delegated in the CSV, and left out of its scope, because an RBAC scope constrains the app''s own access, not the user''s.' }
$Html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Exchange App Mailbox Activity - $(HtmlEnc $tenantTag)</title>
    <style>
        /* Semantic palette - status colors are reserved for state, never decoration.
           All text/background pairs validated >= 4.5:1 (muted ink reserved for
           non-essential de-emphasis). */
        :root {
            --page: #f2f1ee;
            --surface: #fcfcfb;
            --ink: #0b0b0b;
            --ink-2: #52514e;
            --ink-3: #898781;
            --hairline: #e1e0d9;
            --link: #1c5cab;
            --good-bg: #e2f5e2;  --good-fg: #006300;
            --warn-bg: #fdf0d5;  --warn-fg: #704d00;
            --crit-bg: #fbe3e3;  --crit-fg: #8f1f1f;  --crit-solid: #b22727;
            --err-bg:  #fde8df;  --err-fg:  #93381b;
            --info-bg: #ddebfa;  --info-fg: #14477f;
            --purp-bg: #e6e2f7;  --purp-fg: #38307e;
            --neut-bg: #eceae5;  --neut-fg: #52514e;
            --code-bg: #232322;  --code-fg: #e8e6df;
        }
        * { box-sizing: border-box; }
        body {
            font-family: system-ui, -apple-system, "Segoe UI", Roboto, sans-serif;
            background: var(--page);
            color: var(--ink);
            line-height: 1.5;
            margin: 0;
            padding: 2rem;
            min-height: 100vh;
        }
        .container { max-width: 1600px; margin: 0 auto; }
        h1 { font-weight: 600; margin-bottom: 0.5rem; color: var(--ink); }
        h2 { font-weight: 600; margin: 2.5rem 0 0.5rem 0; color: var(--ink); }
        .subtitle, .section-sub { color: var(--ink-2); margin-bottom: 1.5rem; }
        .summary {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
            gap: 1rem;
            margin-bottom: 2rem;
        }
        .stat {
            background: var(--surface);
            border-radius: 8px;
            padding: 1rem;
            border: 1px solid var(--hairline);
        }
        .stat-value { font-size: 1.75rem; font-weight: 600; color: var(--ink); }
        .stat-label { color: var(--ink-2); font-size: 0.8rem; }
        .stat.warning .stat-value { color: var(--warn-fg); }
        .stat.danger .stat-value { color: var(--crit-fg); }
        .stat.success .stat-value { color: var(--good-fg); }
        .migration-note {
            background: var(--surface);
            border-left: 4px solid var(--info-fg);
            padding: 1rem;
            margin-bottom: 2rem;
            border-radius: 0 8px 8px 0;
            border-top: 1px solid var(--hairline);
            border-right: 1px solid var(--hairline);
            border-bottom: 1px solid var(--hairline);
        }
        .migration-note h3 { margin: 0 0 0.5rem 0; color: var(--info-fg); }
        .migration-note p { margin: 0 0 0.5rem 0; color: var(--ink-2); }
        .migration-note p:last-child { margin-bottom: 0; }
        .migration-note a { color: var(--link); }
        table {
            width: 100%;
            background: var(--surface);
            border-radius: 8px;
            border-collapse: collapse;
            border: 1px solid var(--hairline);
        }
        th {
            text-align: left;
            padding: 0.75rem;
            background: var(--page);
            border-bottom: 2px solid var(--hairline);
            font-weight: 600;
            font-size: 0.75rem;
            text-transform: uppercase;
            letter-spacing: 0.025em;
            color: var(--ink-2);
        }
        td {
            padding: 0.75rem;
            border-bottom: 1px solid var(--hairline);
            vertical-align: top;
            font-size: 0.875rem;
        }
        tr:last-child td { border-bottom: none; }
        /* Delegated records: shown for review, not counted toward the scope. */
        tr.delegated-row td { background: var(--neut-bg); color: var(--ink-2); }
        .app-id { font-family: ui-monospace, "Cascadia Mono", Consolas, monospace; font-size: 0.7rem; color: var(--ink-3); }
        a.plink { color: var(--link); font-size: 0.75rem; text-decoration: none; }
        a.plink:hover { text-decoration: underline; }
        .status-badge {
            display: inline-block;
            font-size: 0.7rem;
            font-weight: 600;
            padding: 0.25rem 0.5rem;
            border-radius: 4px;
            text-transform: uppercase;
            vertical-align: middle;
            margin-left: 0.5rem;
        }
        .status-badge.ready { background: var(--good-bg); color: var(--good-fg); }
        .status-badge.error { background: var(--err-bg); color: var(--err-fg); }
        .status-badge.delete { background: var(--neut-bg); color: var(--neut-fg); }
        .status-badge.unconstrained { background: var(--crit-fg); color: #ffffff; }
        .badge {
            display: inline-block;
            font-size: 0.65rem;
            font-weight: 600;
            padding: 0.15rem 0.4rem;
            border-radius: 4px;
            margin-left: 0.25rem;
            vertical-align: middle;
            text-transform: uppercase;
        }
        .badge.crit { background: var(--crit-bg); color: var(--crit-fg); }
        .badge.warn { background: var(--warn-bg); color: var(--warn-fg); }
        .badge.info { background: var(--info-bg); color: var(--info-fg); }
        .badge.purp { background: var(--purp-bg); color: var(--purp-fg); }
        .badge.neut { background: var(--neut-bg); color: var(--neut-fg); }
        .email { color: var(--ink-3); font-size: 0.75rem; }
        .perm {
            display: inline-block;
            padding: 0.1rem 0.35rem;
            border-radius: 3px;
            margin: 0.1rem;
            font-family: ui-monospace, "Cascadia Mono", Consolas, monospace;
            font-size: 0.7rem;
        }
        .perm.migrate { background: var(--info-bg); color: var(--info-fg); }
        .perm.keep { background: var(--neut-bg); color: var(--neut-fg); border: 1px dashed var(--ink-3); }
        .none { color: var(--ink-3); font-style: italic; font-size: 0.75rem; }
        .empty-state { padding: 2rem 1rem; text-align: center; color: var(--ink-2); }
        .commands {
            background: var(--code-bg);
            color: var(--code-fg);
            padding: 1rem;
            border-radius: 6px;
            font-family: ui-monospace, "Cascadia Mono", Consolas, monospace;
            font-size: 0.75rem;
            white-space: pre-wrap;
            overflow-x: auto;
            margin: 0;
        }
        /* Collapsed-by-default disclosure, shared by command blocks and long mailbox lists */
        details.cmd > summary {
            cursor: pointer;
            list-style: none;
            display: inline-flex;
            align-items: center;
            gap: 0.5rem;
            font-size: 0.72rem;
            font-weight: 600;
            color: var(--ink-2);
            background: var(--surface);
            border: 1px solid var(--hairline);
            border-radius: 999px;
            padding: 0.3rem 0.85rem;
            user-select: none;
            transition: border-color 0.15s ease, color 0.15s ease;
        }
        details.cmd > summary:hover { color: var(--link); border-color: var(--link); }
        details.cmd > summary::-webkit-details-marker { display: none; }
        details.cmd > summary::before {
            content: '';
            width: 0.42em;
            height: 0.42em;
            border-right: 2px solid currentColor;
            border-bottom: 2px solid currentColor;
            transform: rotate(-45deg);
            transition: transform 0.15s ease;
            flex: none;
        }
        details.cmd[open] > summary::before { transform: rotate(45deg); }
        details.cmd > pre.commands { margin-top: 0.5rem; }
        .cmd-meta { font-weight: 400; color: var(--ink-3); }
        .copy-btn {
            font: inherit;
            font-weight: 600;
            color: var(--link);
            background: none;
            border: none;
            border-left: 1px solid var(--hairline);
            padding: 0 0 0 0.6rem;
            cursor: pointer;
        }
        .copy-btn:hover { text-decoration: underline; }
        .footer {
            margin-top: 2rem;
            padding-top: 1rem;
            border-top: 1px solid var(--hairline);
            color: var(--ink-2);
            font-size: 0.8rem;
        }
        .footer code {
            background: var(--neut-bg);
            padding: 0.1rem 0.3rem;
            border-radius: 3px;
            font-size: 0.75rem;
        }
        .footer a { color: var(--link); }
    </style>
</head>
<body>
    <div class="container">
        <h1>Exchange App Mailbox Activity</h1>
        <p class="subtitle"><strong>$(HtmlEnc $tenantTag)</strong> &middot; $(Get-Date -Format 'MMMM d, yyyy \a\t h:mm:ss tt') &middot; last $Days days of mailbox audit records</p>
        <div class="migration-note">
            <h3>What this shows</h3>
            <p>For each app, the mailboxes it actually touched, grouped by the operation the audit log recorded
            (Send, MailItemsAccessed, Update...). Use it to judge whether the activity matches the app's job, and
            to build the RBAC for Applications scope from real mailboxes instead of guesses. Each app's PowerShell
            block does exactly that. The audit of which apps are unconstrained, and the verified cutover that revokes the
            tenant-wide grant afterwards, live in <a href="$AuditScriptUrl">Audit-ExoAppAccessPolicies.ps1</a>.
            How this report works: <a href="$PostUrl">How to Identify Which Mailbox(es) an Entra App Is Using</a>;
            the migration itself: <a href="$BlogUrl">Migrating Exchange Apps to RBAC for Applications</a>.</p>
            <p>Source: the unified audit log via the Purview audit search API. The apps listed are the ones holding
            <em>application</em> permissions. Records whose token belonged to a user rather than to the app's own service
            principal (the record's TokenObjectId) came from a signed-in user working through the app (delegated
            access); $DelegatedNote A record with no TokenObjectId counts as the app's own. Every app/operation/mailbox
            row is in the companion CSV next to this file.</p>
        </div>
        <div class="summary">
            <div class="stat"><div class="stat-value">$($Report.Count)</div><div class="stat-label">Apps analyzed</div></div>
            <div class="stat success"><div class="stat-value">$((& $countBy 'Single mailbox') + (& $countBy 'Scope candidate'))</div><div class="stat-label">Scope candidates</div></div>
            <div class="stat danger"><div class="stat-value">$(& $countBy 'Broad')</div><div class="stat-label">Broad (over $MaxScopeMailboxes mailboxes)</div></div>
            <div class="stat"><div class="stat-value">$(& $countBy 'No activity')</div><div class="stat-label">No activity</div></div>
            <div class="stat warning"><div class="stat-value">$(@($Report | Where-Object { $_.Truncated }).Count)</div><div class="stat-label">Truncated searches</div></div>
            <div class="stat"><div class="stat-value">$(@($Report | Where-Object { $_.Failed }).Count)</div><div class="stat-label">Not analyzed</div></div>
        </div>
$($AppSections -join "`n")
        <div class="footer">
            <strong>Reading the grants:</strong> a grant is marked <em>idle</em> when no operation of its kind (send,
            read, write) was observed in the window. Send is unambiguous. Reads and writes could be mail, calendar,
            contacts or settings, which the audit record does not reliably distinguish, so one observed write matches
            Mail.ReadWrite, Calendars.ReadWrite and Contacts.ReadWrite alike. A calendar, contacts, settings or MailTips
            grant matched that way is marked <em>unproven</em>: keep its role if the app needs it, drop the grant if the
            app's job is mail. IMAP, POP and unrecognized grants are marked <em>not measured</em>, because no mailbox
            audit record this report can match proves or disproves them.<br><br>
            <strong>What hides records:</strong>
            The organization switch first: <code>Get-OrganizationConfig | Format-List AuditDisabled</code>. True means
            nothing is audited, whatever the mailboxes say, and every app comes back idle; the pre-flight cannot see it.
            An audit bypass (<code>Get-MailboxAuditBypassAssociation</code>) is set on an account, not a mailbox, and
            suppresses that account's actions in every mailbox it touches. A mailbox whose audit actions were customized
            still logs what is on its list, but Microsoft stops adding new default actions to it, so an owner set
            customized before Send and MailItemsAccessed joined the defaults (2024) lacks them: check
            <code>Get-Mailbox &lt;mailbox&gt; | Format-List DefaultAuditSet, AuditOwner, AuditDelegate, AuditAdmin</code>.
            To fix one, restore Microsoft's defaults with
            <code>Set-Mailbox &lt;mailbox&gt; -DefaultAuditSet Admin,Delegate,Owner</code> (it drops any extra actions
            someone added, and new defaults then arrive on their own). Adding the two actions to the customized list
            instead fails while it still holds the retired MessageBind.
            During that 2024 rollout Microsoft also told Audit Standard tenants to run
            <code>Set-Mailbox -AuditEnabled `$true</code> on every mailbox regardless of its current value, so a mailbox
            nobody reran it on can still be missing Send and MailItemsAccessed.
            Resource mailboxes are not covered by default mailbox auditing, so a room-booking app can look idle.<br><br>
            <strong>Caveats:</strong>
            Audit Standard keeps 180 days. Read counts are a floor: reads inside a short window are folded into one
            MailItemsAccessed record. EWS records written since Microsoft's November 2024 change name the app: a
            September 2026 test found an app-only EWS app in AppId, ClientAppId and AppAccessContext.ClientAppId. Older
            ones named only EWS itself, but no window this report can search (180 days at most) reaches back that far.
            A search that stopped at the service's per-search record limit is marked
            <em>truncated</em>; rerun that app with a shorter window. Tenants with Microsoft Graph activity logs in
            Log Analytics have a second, more complete view of sends and can cross-check.
            Not available in GCC High or DoD, where <code>Search-UnifiedAuditLog -FreeText &lt;appId&gt;</code> is the fallback.
        </div>
    </div>
    <script>
    document.addEventListener('click', function (e) {
        var btn = e.target.closest('.copy-btn');
        if (!btn) { return; }
        e.preventDefault();
        var pre = btn.closest('details').querySelector('pre');
        var done = function () {
            btn.textContent = 'Copied';
            setTimeout(function () { btn.textContent = 'Copy'; }, 1500);
        };
        var fallback = function () {
            var ta = document.createElement('textarea');
            ta.value = pre.textContent;
            document.body.appendChild(ta);
            ta.select();
            try { document.execCommand('copy'); } catch (err) { }
            document.body.removeChild(ta);
            done();
        };
        if (navigator.clipboard && navigator.clipboard.writeText) {
            navigator.clipboard.writeText(pre.textContent).then(done, fallback);
        } else {
            fallback();
        }
    });
    </script>
</body>
</html>
"@
$Html | Out-File -FilePath $HtmlPath -Encoding UTF8
Write-Host "Report saved to: $HtmlPath" -ForegroundColor Green
Write-Host "Mailbox detail saved to: $CsvPath" -ForegroundColor Green
Write-Verbose "Total run time $($clock.Elapsed.ToString('mm\:ss'))"
Start-Process $HtmlPath
#endregion Output
