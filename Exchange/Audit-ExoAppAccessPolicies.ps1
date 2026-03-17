#Requires -Modules ExchangeOnlineManagement, Microsoft.Graph.Authentication

<#
.SYNOPSIS
    Inventories Exchange Application Access Policies and generates RBAC migration guidance.

.DESCRIPTION
    Application Access Policies in Exchange Online are deprecated and being replaced by
    RBAC for Applications. This script audits all existing policies and produces an HTML
    report with:

      - Per-policy migration status (Ready, Blocked, Review, Delete Only)
      - Entra ID app registration details and granted permissions
      - Target scope resolution (groups, individual mailboxes, orphaned targets)
      - Exchange Service Principal existence check
      - Ready-to-run PowerShell commands for each migration step
      - Summary statistics with visual dashboard

    Migration steps generated per policy:
      1. New-ServicePrincipal       — Create Exchange pointer to the Entra service principal
      2. New-ManagementScope        — Define mailbox scope via MemberOfGroup filter
      3. New-ManagementRoleAssignment — Assign RBAC roles with the custom scope
      4. Remove Entra API permissions (manual portal step)
      5. Remove-ApplicationAccessPolicy — Delete the legacy policy

.EXAMPLE
    .\Audit-ExoAppAccessPolicies.ps1

    Connects to Microsoft Graph and Exchange Online (both via device code auth), audits
    all Application Access Policies, and saves an HTML report to the desktop.

.EXAMPLE
    # Pre-connect to Graph and Exchange before running:
    Connect-MgGraph -Scopes "Application.Read.All", "Directory.Read.All"
    Connect-ExchangeOnline
    .\Audit-ExoAppAccessPolicies.ps1

    Uses existing sessions if already connected.

.INPUTS
    None. This script does not accept pipeline input.

.OUTPUTS
    An HTML report saved to the desktop:
    AppAccessPolicyMigration_<TenantName>_<Date>.html

.NOTES
    Author:  Mike Crowley
    https://mikecrowley.us

    Permissions required:
      - Microsoft Graph: Application.Read.All, Directory.Read.All
      - Exchange Online: Organization Management or Security Reader

    The permission-to-RBAC-role mapping covers the most common Graph and Exchange Online
    application permissions. Unmapped permissions will appear in the report but will not
    generate role assignment commands — review these manually.

.LINK
    https://learn.microsoft.com/en-us/exchange/permissions/application-rbac
#>

Disconnect-MgGraph -ErrorAction SilentlyContinue
Connect-MgGraph -DeviceCode -NoWelcome -ContextScope Process
Connect-ExchangeOnline -Device -ShowBanner:$false

# Get tenant name for filename
$Org = Invoke-MgGraphRequest -Uri "v1.0/organization"
$TenantName = ($Org.value[0].displayName -replace '[^\w\-]', '')

# Permission to RBAC role mapping
$PermissionToRbacRole = @{
    "Mail.Read"                 = "Application Mail.Read"
    "Mail.ReadBasic"            = "Application Mail.ReadBasic"
    "Mail.ReadWrite"            = "Application Mail.ReadWrite"
    "Mail.Send"                 = "Application Mail.Send"
    "MailboxSettings.Read"      = "Application MailboxSettings.Read"
    "MailboxSettings.ReadWrite" = "Application MailboxSettings.ReadWrite"
    "Calendars.Read"            = "Application Calendars.Read"
    "Calendars.ReadWrite"       = "Application Calendars.ReadWrite"
    "Contacts.Read"             = "Application Contacts.Read"
    "Contacts.ReadWrite"        = "Application Contacts.ReadWrite"
    "full_access_as_app"        = "Application EWS.AccessAsApp"
    "IMAP.AccessAsApp"          = "Application EWS.AccessAsApp"
}

# Cache service principal app roles for permission resolution
$ServicePrincipals = @{}
@(
    "00000003-0000-0000-c000-000000000000"  # Microsoft Graph
    "00000002-0000-0ff1-ce00-000000000000"  # Exchange Online
) | ForEach-Object {
    $ServicePrincipals[$_] = Invoke-MgGraphRequest -Uri "v1.0/servicePrincipals(appId='$_')"
}

# Cache existing Exchange Service Principals
$ExchangeServicePrincipals = @{}
Get-ServicePrincipal -ErrorAction SilentlyContinue | ForEach-Object {
    $ExchangeServicePrincipals[$_.AppId] = $_
}

function Resolve-PermissionName {
    param([string]$ResourceAppId, [string]$PermissionId)
    $sp = $ServicePrincipals[$ResourceAppId]
    if (-not $sp) { return $PermissionId }
    $match = $sp.appRoles | Where-Object { $_.id -eq $PermissionId }
    if ($match) { return $match.value } else { return $PermissionId }
}

function Find-GroupByScopeName {
    param([string]$ScopeName)
    if (-not $ScopeName) { return $null }
    $escaped = $ScopeName.Replace("'", "''")
    foreach ($prop in @('mailNickname', 'displayName', 'mail')) {
        try {
            $result = Invoke-MgGraphRequest -Uri "v1.0/groups?`$filter=$prop eq '$escaped'&`$select=id,displayName,mail,mailNickname"
            if ($result.value.Count -gt 0) { return $result.value[0] }
        }
        catch { }
    }
    return $null
}

function Find-RecipientByScopeName {
    param([string]$ScopeName)
    if (-not $ScopeName) { return $null }
    try {
        return Get-Recipient -Identity $ScopeName -ErrorAction Stop | Select-Object DisplayName, PrimarySmtpAddress, RecipientType, RecipientTypeDetails
    }
    catch { }
    if ($ScopeName -match '^(.+),\s*(\w+)\s*\(') {
        try {
            return Get-Recipient -Identity $Matches[2] -ErrorAction Stop | Select-Object DisplayName, PrimarySmtpAddress, RecipientType, RecipientTypeDetails
        }
        catch { }
    }
    return $null
}

function Get-GroupDistinguishedName {
    param([string]$GroupEmail)
    if (-not $GroupEmail) { return $null }
    try {
        $group = Get-Group -Identity $GroupEmail -ErrorAction Stop
        return $group.DistinguishedName
    }
    catch {
        return $null
    }
}

$Policies = Get-ApplicationAccessPolicy

$Report = foreach ($Policy in $Policies) {
    $AppId = $Policy.AppId
    $Issues = @()
    $MigrationStatus = "Ready"
    $MigrationBlockers = @()

    # App registration
    $AppFound = $true
    try {
        $App = Invoke-MgGraphRequest -Uri "v1.0/applications(appId='$AppId')"
    }
    catch {
        $App = @{ displayName = $AppId; id = $null }
        $AppFound = $false
        $Issues += "Orphaned"
        $MigrationStatus = "Delete Only"
        $MigrationBlockers += "App deleted - remove policy only"
    }

    # Check if Exchange Service Principal already exists
    $ExoSpExists = $ExchangeServicePrincipals.ContainsKey($AppId)

    # Get Entra Service Principal details (need ObjectId for New-ServicePrincipal)
    $EntraSpObjectId = $null
    if ($AppFound) {
        try {
            $EntraSp = Invoke-MgGraphRequest -Uri "v1.0/servicePrincipals(appId='$AppId')?`$select=id,appId,displayName"
            $EntraSpObjectId = $EntraSp.id
        }
        catch { }
    }

    # Granted permissions from service principal
    $GrantedPermissions = @()
    $RbacRoles = @()
    try {
        $AppSp = Invoke-MgGraphRequest -Uri "v1.0/servicePrincipals(appId='$AppId')"
        if ($AppSp.id) {
            $assignments = Invoke-MgGraphRequest -Uri "v1.0/servicePrincipals/$($AppSp.id)/appRoleAssignments"
            foreach ($assignment in $assignments.value) {
                try {
                    $resourceSp = Invoke-MgGraphRequest -Uri "v1.0/servicePrincipals/$($assignment.resourceId)?`$select=appId"
                    $permName = Resolve-PermissionName -ResourceAppId $resourceSp.appId -PermissionId $assignment.appRoleId
                    $resourceName = switch ($resourceSp.appId) {
                        "00000003-0000-0000-c000-000000000000" { "Graph" }
                        "00000002-0000-0ff1-ce00-000000000000" { "EXO" }
                        default { "Other" }
                    }
                    $GrantedPermissions += "$resourceName`:$permName"

                    if ($PermissionToRbacRole.ContainsKey($permName)) {
                        $RbacRoles += $PermissionToRbacRole[$permName]
                    }
                }
                catch { }
            }
        }
    }
    catch { }

    $RbacRoles = $RbacRoles | Select-Object -Unique

    if ($GrantedPermissions.Count -eq 0 -and $AppFound) {
        $Issues += "No Permissions"
        $MigrationStatus = "Review"
        $MigrationBlockers += "No mail/calendar permissions - verify if policy needed"
    }

    # Resolve target
    $GroupInfo = Find-GroupByScopeName -ScopeName $Policy.ScopeName
    $GroupDN = $null
    if ($GroupInfo) {
        $TargetType = "Group"
        $TargetName = $GroupInfo.displayName
        $TargetEmail = $GroupInfo.mail
        $GroupDN = Get-GroupDistinguishedName -GroupEmail $TargetEmail
    }
    else {
        $Recipient = Find-RecipientByScopeName -ScopeName $Policy.ScopeName
        if ($Recipient) {
            $TargetType = $Recipient.RecipientTypeDetails
            $TargetName = $Recipient.DisplayName
            $TargetEmail = $Recipient.PrimarySmtpAddress
            if ($TargetType -match 'Mailbox') {
                $Issues += "Single Mailbox"
                $MigrationStatus = "Blocked"
                $MigrationBlockers += "Create mail-enabled security group first"
            }
        }
        else {
            $TargetType = "Not Found"
            $TargetName = $Policy.ScopeName
            $TargetEmail = $null
            $Issues += "Target Missing"
            $MigrationStatus = "Blocked"
            $MigrationBlockers += "Target group/mailbox not found"
        }
    }

    # Build migration commands
    $MigrationCommands = @()
    if ($MigrationStatus -eq "Ready" -and $AppFound -and $GroupDN) {
        $safeName = ($App.displayName -replace '[^\w\-]', '_')
        $scopeName = "Scope_$safeName"

        if (-not $ExoSpExists) {
            $MigrationCommands += "# Step 1: Create Exchange Service Principal"
            $MigrationCommands += "New-ServicePrincipal -AppId '$AppId' -ObjectId '$EntraSpObjectId' -DisplayName '$($App.displayName)'"
            $MigrationCommands += ""
        }

        $MigrationCommands += "# Step 2: Create Management Scope"
        $MigrationCommands += "New-ManagementScope -Name '$scopeName' -RecipientRestrictionFilter `"MemberOfGroup -eq '$GroupDN'`""
        $MigrationCommands += ""

        if ($RbacRoles.Count -gt 0) {
            $MigrationCommands += "# Step 3: Create RBAC Role Assignments"
            foreach ($role in $RbacRoles) {
                $MigrationCommands += "New-ManagementRoleAssignment -App '$AppId' -Role '$role' -CustomResourceScope '$scopeName'"
            }
            $MigrationCommands += ""
        }

        $MigrationCommands += "# Step 4: Remove permissions from Entra ID (Azure Portal)"
        $MigrationCommands += "# Navigate to: Entra ID > App registrations > $($App.displayName) > API permissions"
        $MigrationCommands += "# Remove: $($GrantedPermissions -join ', ')"
        $MigrationCommands += ""

        $MigrationCommands += "# Step 5: Remove legacy Application Access Policy"
        $MigrationCommands += "Remove-ApplicationAccessPolicy -Identity '$($Policy.Identity)'"
    }
    elseif ($MigrationStatus -eq "Delete Only") {
        $MigrationCommands += "# Orphaned policy - safe to remove"
        $MigrationCommands += "Remove-ApplicationAccessPolicy -Identity '$($Policy.Identity)'"
    }

    [PSCustomObject]@{
        App_DisplayName    = $App.displayName
        App_Id             = $App.id
        App_ClientId       = $AppId
        Entra_SpObjectId   = $EntraSpObjectId
        Exo_SpExists       = $ExoSpExists
        Policy_Identity    = $Policy.Identity
        Policy_AccessRight = $Policy.AccessRight
        Target_Type        = $TargetType
        Target_Name        = $TargetName
        Target_Email       = $TargetEmail
        Target_DN          = $GroupDN
        Permissions        = if ($GrantedPermissions) { $GrantedPermissions -join "; " } else { "" }
        RbacRoles          = if ($RbacRoles) { $RbacRoles -join "; " } else { "" }
        Issues             = $Issues -join ", "
        MigrationStatus    = $MigrationStatus
        MigrationBlockers  = $MigrationBlockers -join "; "
        MigrationCommands  = $MigrationCommands -join "`n"
    }
}

# Summary stats
$TotalPolicies = $Report.Count
$OrphanedApps = ($Report | Where-Object { $_.Issues -match "Orphaned" }).Count
$NoPermissions = ($Report | Where-Object { $_.Issues -match "No Permissions" }).Count
$SingleMailbox = ($Report | Where-Object { $_.Issues -match "Single Mailbox" }).Count
$ReadyToMigrate = ($Report | Where-Object { $_.MigrationStatus -eq "Ready" }).Count
$Blocked = ($Report | Where-Object { $_.MigrationStatus -eq "Blocked" }).Count

# Build table rows
$HtmlRows = foreach ($item in $Report | Sort-Object MigrationStatus, App_DisplayName) {
    $rowClass = switch ($item.MigrationStatus) {
        "Blocked" { "status-blocked" }
        "Review" { "status-review" }
        "Delete Only" { "status-delete" }
        default { "" }
    }

    $statusBadge = switch ($item.MigrationStatus) {
        "Ready" { '<span class="status-badge ready">Ready</span>' }
        "Blocked" { '<span class="status-badge blocked">Blocked</span>' }
        "Review" { '<span class="status-badge review">Review</span>' }
        "Delete Only" { '<span class="status-badge delete">Delete Only</span>' }
        default { '' }
    }

    $issuesBadges = ""
    if ($item.Issues -match "Orphaned") { $issuesBadges += '<span class="badge orphaned">Orphaned App</span>' }
    if ($item.Issues -match "No Permissions") { $issuesBadges += '<span class="badge no-perms">No Permissions</span>' }
    if ($item.Issues -match "Single Mailbox") { $issuesBadges += '<span class="badge single-mbx">Single Mailbox</span>' }
    if ($item.Issues -match "Target Missing") { $issuesBadges += '<span class="badge target-missing">Target Missing</span>' }

    $exoSpStatus = if ($item.Exo_SpExists) { '<span class="sp-exists">Exists</span>' } else { '<span class="sp-missing">Create needed</span>' }

    $permsList = if ($item.Permissions) {
        ($item.Permissions -split "; " | ForEach-Object { "<span class='perm'>$_</span>" }) -join ""
    }
    else { "<span class='none'>None granted</span>" }

    $rbacList = if ($item.RbacRoles) {
        ($item.RbacRoles -split "; " | ForEach-Object { "<span class='rbac-role'>$_</span>" }) -join ""
    }
    else { "<span class='none'>N/A</span>" }

    $commandsHtml = if ($item.MigrationCommands) {
        $escaped = [System.Web.HttpUtility]::HtmlEncode($item.MigrationCommands)
        "<pre class='commands'>$escaped</pre>"
    }
    else { "" }

    $blockersHtml = if ($item.MigrationBlockers) {
        "<div class='blockers'>Warning: $($item.MigrationBlockers)</div>"
    }
    else { "" }

    @"
    <tr class="$rowClass">
        <td>
            <strong>$($item.App_DisplayName)</strong>$issuesBadges<br>
            <span class="app-id">$($item.App_ClientId)</span>
        </td>
        <td>$statusBadge$blockersHtml</td>
        <td>
            <span class="target-type $($item.Target_Type.ToLower() -replace '\s','')">$($item.Target_Type)</span><br>
            $($item.Target_Name)<br>
            <span class="email">$($item.Target_Email)</span>
        </td>
        <td>$exoSpStatus</td>
        <td class="perms">$permsList</td>
        <td class="perms">$rbacList</td>
    </tr>
    <tr class="commands-row $rowClass">
        <td colspan="6">$commandsHtml</td>
    </tr>
"@
}

$Html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Application Access Policy Migration Report - $TenantName</title>
    <style>
        * { box-sizing: border-box; }
        body {
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
            background: linear-gradient(145deg, #f0f4f8 0%, #d9e2ec 100%);
            background-attachment: fixed;
            color: #3d4852;
            line-height: 1.5;
            margin: 0;
            padding: 2rem;
            min-height: 100vh;
        }
        .container { max-width: 1600px; margin: 0 auto; }
        h1 { font-weight: 600; margin-bottom: 0.5rem; color: #2d3748; }
        .subtitle { color: #718096; margin-bottom: 2rem; }

        .summary {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
            gap: 1rem;
            margin-bottom: 2rem;
        }
        .stat {
            background: rgba(255,255,255,0.8);
            border-radius: 8px;
            padding: 1rem;
            border: 1px solid #cbd5e0;
        }
        .stat-value { font-size: 1.75rem; font-weight: 600; color: #4a5568; }
        .stat-label { color: #718096; font-size: 0.8rem; }
        .stat.warning .stat-value { color: #b7791f; }
        .stat.danger .stat-value { color: #9b2c2c; }
        .stat.success .stat-value { color: #276749; }
        .stat.info .stat-value { color: #2b6cb0; }

        .migration-note {
            background: rgba(255,255,255,0.8);
            border-left: 4px solid #4299e1;
            padding: 1rem;
            margin-bottom: 2rem;
            border-radius: 0 8px 8px 0;
        }
        .migration-note h3 { margin: 0 0 0.5rem 0; color: #2b6cb0; }
        .migration-note p { margin: 0; color: #4a5568; }

        table {
            width: 100%;
            background: rgba(255,255,255,0.85);
            border-radius: 8px;
            border-collapse: collapse;
            border: 1px solid #cbd5e0;
        }
        th {
            text-align: left;
            padding: 0.75rem;
            background: rgba(247,250,252,0.9);
            border-bottom: 2px solid #cbd5e0;
            font-weight: 600;
            font-size: 0.75rem;
            text-transform: uppercase;
            letter-spacing: 0.025em;
            color: #4a5568;
        }
        td {
            padding: 0.75rem;
            border-bottom: 1px solid #e2e8f0;
            vertical-align: top;
            font-size: 0.875rem;
        }
        tr:last-child td { border-bottom: none; }
        tr.status-blocked { background: rgba(254,215,215,0.3); }
        tr.status-review { background: rgba(250,240,137,0.3); }
        tr.status-delete { background: rgba(198,198,198,0.2); }

        .commands-row td { padding: 0 0.75rem 0.75rem 0.75rem; }
        .commands-row.status-blocked td { background: rgba(254,215,215,0.15); }
        .commands-row.status-review td { background: rgba(250,240,137,0.15); }
        .commands-row.status-delete td { background: rgba(198,198,198,0.1); }

        .app-id { font-family: "SF Mono", Monaco, monospace; font-size: 0.7rem; color: #a0aec0; }

        .status-badge {
            display: inline-block;
            font-size: 0.7rem;
            font-weight: 600;
            padding: 0.25rem 0.5rem;
            border-radius: 4px;
            text-transform: uppercase;
        }
        .status-badge.ready { background: #c6f6d5; color: #276749; }
        .status-badge.blocked { background: #feb2b2; color: #742a2a; }
        .status-badge.review { background: #faf089; color: #744210; }
        .status-badge.delete { background: #e2e8f0; color: #4a5568; }

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
        .badge.orphaned { background: #feb2b2; color: #742a2a; }
        .badge.no-perms { background: #faf089; color: #744210; }
        .badge.single-mbx { background: #90cdf4; color: #2a4365; }
        .badge.target-missing { background: #feb2b2; color: #742a2a; }

        .target-type {
            display: inline-block;
            font-size: 0.7rem;
            font-weight: 500;
            padding: 0.2rem 0.4rem;
            border-radius: 4px;
            background: #e2e8f0;
            color: #4a5568;
        }
        .target-type.group { background: #9ae6b4; color: #22543d; }
        .target-type.usermailbox { background: #90cdf4; color: #2a4365; }
        .target-type.notfound { background: #feb2b2; color: #742a2a; }

        .email { color: #a0aec0; font-size: 0.75rem; }

        .sp-exists { color: #276749; font-size: 0.75rem; }
        .sp-missing { color: #975a16; font-size: 0.75rem; }

        .perm, .rbac-role {
            display: inline-block;
            background: #e2e8f0;
            padding: 0.1rem 0.35rem;
            border-radius: 3px;
            margin: 0.1rem;
            font-family: "SF Mono", Monaco, monospace;
            font-size: 0.7rem;
            color: #4a5568;
        }
        .rbac-role { background: #c6f6d5; color: #276749; }
        .none { color: #a0aec0; font-style: italic; font-size: 0.75rem; }

        .blockers {
            margin-top: 0.5rem;
            font-size: 0.75rem;
            color: #9b2c2c;
        }

        .commands {
            background: #2d3748;
            color: #e2e8f0;
            padding: 1rem;
            border-radius: 6px;
            font-family: "SF Mono", Monaco, monospace;
            font-size: 0.75rem;
            white-space: pre-wrap;
            overflow-x: auto;
            margin: 0;
        }

        .footer {
            margin-top: 2rem;
            padding-top: 1rem;
            border-top: 1px solid #cbd5e0;
            color: #718096;
            font-size: 0.8rem;
        }
        .footer code {
            background: #e2e8f0;
            padding: 0.1rem 0.3rem;
            border-radius: 3px;
            font-size: 0.75rem;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>Application Access Policy &rarr; RBAC Migration Report</h1>
        <p class="subtitle"><strong>$TenantName</strong> &mdash; $(Get-Date -Format "MMMM d, yyyy 'at' h:mm tt")</p>

        <div class="migration-note">
            <h3>Migration Required</h3>
            <p>Application Access Policies are deprecated and will be replaced by RBAC for Applications.
            This report identifies current policies and generates the PowerShell commands needed to migrate each one.</p>
        </div>

        <div class="summary">
            <div class="stat">
                <div class="stat-value">$TotalPolicies</div>
                <div class="stat-label">Total Policies</div>
            </div>
            <div class="stat success">
                <div class="stat-value">$ReadyToMigrate</div>
                <div class="stat-label">Ready to Migrate</div>
            </div>
            <div class="stat danger">
                <div class="stat-value">$Blocked</div>
                <div class="stat-label">Blocked</div>
            </div>
            <div class="stat warning">
                <div class="stat-value">$NoPermissions</div>
                <div class="stat-label">No Permissions</div>
            </div>
            <div class="stat info">
                <div class="stat-value">$SingleMailbox</div>
                <div class="stat-label">Single Mailbox</div>
            </div>
            <div class="stat">
                <div class="stat-value">$OrphanedApps</div>
                <div class="stat-label">Orphaned</div>
            </div>
        </div>

        <table>
            <thead>
                <tr>
                    <th>Application</th>
                    <th>Migration Status</th>
                    <th>Target Scope</th>
                    <th>EXO Service Principal</th>
                    <th>Current Permissions</th>
                    <th>RBAC Roles</th>
                </tr>
            </thead>
            <tbody>
                $($HtmlRows -join "`n")
            </tbody>
        </table>

        <div class="footer">
            <strong>Migration Steps:</strong><br>
            1. <code>New-ServicePrincipal</code> - Create Exchange pointer to Entra service principal<br>
            2. <code>New-ManagementScope</code> - Define mailbox scope using <code>MemberOfGroup</code> filter<br>
            3. <code>New-ManagementRoleAssignment</code> - Assign RBAC roles with scope<br>
            4. Remove Graph/EXO permissions from Entra ID (Azure Portal)<br>
            5. <code>Remove-ApplicationAccessPolicy</code> - Delete legacy policy<br><br>
            <strong>Blockers:</strong> Single Mailbox targets require creating a mail-enabled security group first.
            Add the mailbox as a member, then use that group for the management scope.
        </div>
    </div>
</body>
</html>
"@

$Desktop = [Environment]::GetFolderPath("Desktop")
$ExportPath = "$Desktop\AppAccessPolicyMigration_$TenantName`_$(Get-Date -Format 'yyyyMMdd').html"
$Html | Out-File -FilePath $ExportPath -Encoding UTF8
Write-Host "Report saved to: $ExportPath" -ForegroundColor Green

Start-Process $ExportPath
