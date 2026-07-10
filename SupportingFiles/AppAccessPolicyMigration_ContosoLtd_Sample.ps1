# Application Access Policy -> RBAC for Applications migration commands
# Generated 2026-07-10 14:05:38 for tenant ContosoLtd by Audit-ExoAppAccessPolicies.ps1
# Prereqs: Connect-ExchangeOnline (Organization Management / Exchange Administrator);
#          Connect-MgGraph -Scopes 'AppRoleAssignment.ReadWrite.All' (for the cutover revocations).
# Steps 1-4 are additive. Cutover blocks self-verify with Test-ServicePrincipalAuthorization
# and skip themselves if RBAC is not live; Remove-ApplicationAccessPolicy prompts.
# Blocks for name-matched targets are fully commented out - verify the target first.

# ==== Contoso Room Booking (ffffffff-0000-0000-0000-000000000001) -> Facilities Rooms [RestrictAccess] ====
# Steps 1-4 are additive and do not change existing access.

# Step 1: Exchange service principal pointer (idempotent)
if (-not (Get-ServicePrincipal -Identity 'ffffffff-0000-0000-0000-000000000001' -ErrorAction SilentlyContinue)) {
    New-ServicePrincipal -AppId 'ffffffff-0000-0000-0000-000000000001' -ObjectId 'abababab-0000-0000-0000-000000000001' -DisplayName 'Contoso Room Booking'
}

# Step 2: Management scope (idempotent; warns if the name is taken by a different filter)
# NOTE: MemberOfGroup covers DIRECT members only - nested group members are out of scope.
$scope = $null
$scopeConflict = $false
$scope = Get-ManagementScope -Identity 'AppRBAC_facilities-rooms' -ErrorAction SilentlyContinue
if (-not $scope) {
    New-ManagementScope -Name 'AppRBAC_facilities-rooms' -RecipientRestrictionFilter "MemberOfGroup -eq 'CN=Facilities Rooms,OU=contoso.onmicrosoft.com,OU=Microsoft Exchange Hosted Organizations,DC=EURPR01A001,DC=prod,DC=outlook,DC=com'"
} elseif ($scope.RecipientFilter -notlike '*CN=Facilities Rooms,OU=contoso.onmicrosoft.com,OU=Microsoft Exchange Hosted Organizations,DC=EURPR01A001,DC=prod,DC=outlook,DC=com*') {
    $scopeConflict = $true
    Write-Warning "Scope 'AppRBAC_facilities-rooms' already exists with a DIFFERENT filter: $($scope.RecipientFilter) - role assignments skipped; resolve the conflict first."
}

# Step 3: RBAC role assignments (idempotent - skips roles already assigned; skipped
# entirely on a scope-name conflict)
if (-not $scopeConflict) {
    $liveRoles = @(Test-ServicePrincipalAuthorization -Identity 'ffffffff-0000-0000-0000-000000000001' -ErrorAction SilentlyContinue | ForEach-Object { $_.RoleName })
    if ($liveRoles -notcontains 'Application Calendars.Read') { New-ManagementRoleAssignment -App 'ffffffff-0000-0000-0000-000000000001' -Role 'Application Calendars.Read' -CustomResourceScope 'AppRBAC_facilities-rooms' }
    if ($liveRoles -notcontains 'Application Calendars.ReadWrite') { New-ManagementRoleAssignment -App 'ffffffff-0000-0000-0000-000000000001' -Role 'Application Calendars.ReadWrite' -CustomResourceScope 'AppRBAC_facilities-rooms' }
}

# Step 4: VERIFY - expect InScope = True for an in-scope mailbox
Test-ServicePrincipalAuthorization -Identity 'ffffffff-0000-0000-0000-000000000001' -Resource 'member1@contoso.com' | Format-Table
# Also confirm the application itself still works before continuing.

# ---- Step 5-6: CUTOVER. Verifies EVERY migrated role is live and InScope, revokes the
# tenant-wide grants (each one checked), then removes the legacy policy (Remove-* prompts
# for confirmation). Entra + RBAC grants are a union - scoping only takes effect after
# the tenant-wide grant is revoked. The policy is only removed if ALL revocations succeed.
$expectedRoles = @('Application Calendars.Read', 'Application Calendars.ReadWrite')
$cutoverTestMailbox = 'member1@contoso.com'   # in-scope member found by the audit; substitute if needed
$auth = if ($cutoverTestMailbox -notlike '<*') { @(Test-ServicePrincipalAuthorization -Identity 'ffffffff-0000-0000-0000-000000000001' -Resource $cutoverTestMailbox -ErrorAction SilentlyContinue) } else { @() }
$missingRoles = @($expectedRoles | Where-Object { $role = $_; -not ($auth | Where-Object { $_.RoleName -eq $role -and $_.InScope }) })
if ($missingRoles.Count -gt 0) {
    Write-Warning "Contoso Room Booking: roles not verified InScope for $cutoverTestMailbox (missing: $($missingRoles -join ', ')) - cutover skipped. Run steps 1-4; if they succeeded, try a different in-scope mailbox."
} elseif (@((Get-MgContext).Scopes) -notcontains 'AppRoleAssignment.ReadWrite.All') {
    Write-Warning 'Graph session lacks AppRoleAssignment.ReadWrite.All - run: Connect-MgGraph -Scopes AppRoleAssignment.ReadWrite.All'
} else {
    $revokeFailed = $false
    # KEEP (NOT replaced by RBAC - do not revoke): Graph:User.Read.All
    try { Invoke-MgGraphRequest -Method DELETE -Uri 'v1.0/servicePrincipals/abababab-0000-0000-0000-000000000001/appRoleAssignments/assign-000001-1' -ErrorAction Stop; Write-Host 'Revoked Graph:Calendars.Read' } catch { $revokeFailed = $true; Write-Warning "Revocation FAILED (Graph:Calendars.Read): $_" }
    try { Invoke-MgGraphRequest -Method DELETE -Uri 'v1.0/servicePrincipals/abababab-0000-0000-0000-000000000001/appRoleAssignments/assign-000001-2' -ErrorAction Stop; Write-Host 'Revoked Graph:Calendars.ReadWrite' } catch { $revokeFailed = $true; Write-Warning "Revocation FAILED (Graph:Calendars.ReadWrite): $_" }
    if ($revokeFailed) {
        Write-Warning 'One or more revocations FAILED - the legacy policy was NOT removed. Fix the errors above and rerun this cutover block.'
    } else {
        Write-Host 'Tenant-wide grants revoked. Exchange caches app permissions 30 min - 2 h; re-test the app.'
        Write-Host 'Then confirm removal of the legacy policy:'
        Remove-ApplicationAccessPolicy -Identity '11111111-1111-1111-1111-111111111111\ffffffff-0000-0000-0000-000000000001:S-1-5-21-1004336348-1177238915-682003330-8175;cccccccc-0000-0000-0000-000000000001'
    }
}
# Hygiene afterwards: App registrations > API permissions - delete the revoked rows
# ('not granted' leftovers). Removing entries there does NOT revoke access by itself.

# ==== HR Notification Bot (ffffffff-0000-0000-0000-000000000004) -> HR Alerts [RestrictAccess] ====
# Steps 1-4 are additive and do not change existing access.

# Step 1: Exchange service principal pointer (idempotent)
if (-not (Get-ServicePrincipal -Identity 'ffffffff-0000-0000-0000-000000000004' -ErrorAction SilentlyContinue)) {
    New-ServicePrincipal -AppId 'ffffffff-0000-0000-0000-000000000004' -ObjectId 'abababab-0000-0000-0000-000000000004' -DisplayName 'HR Notification Bot'
}

# Step 2: Management scope (idempotent; warns if the name is taken by a different filter)
# Single-mailbox scope. To cover more mailboxes later, create a mail-enabled security
# group instead and use: "MemberOfGroup -eq '<group DN>'"
$scope = $null
$scopeConflict = $false
$scope = Get-ManagementScope -Identity 'AppRBAC_hr-alerts' -ErrorAction SilentlyContinue
if (-not $scope) {
    New-ManagementScope -Name 'AppRBAC_hr-alerts' -RecipientRestrictionFilter "ExternalDirectoryObjectId -eq 'eeeeeeee-0000-0000-0000-000000000001'"
} elseif ($scope.RecipientFilter -notlike '*eeeeeeee-0000-0000-0000-000000000001*') {
    $scopeConflict = $true
    Write-Warning "Scope 'AppRBAC_hr-alerts' already exists with a DIFFERENT filter: $($scope.RecipientFilter) - role assignments skipped; resolve the conflict first."
}

# Step 3: RBAC role assignments (idempotent - skips roles already assigned; skipped
# entirely on a scope-name conflict)
if (-not $scopeConflict) {
    $liveRoles = @(Test-ServicePrincipalAuthorization -Identity 'ffffffff-0000-0000-0000-000000000004' -ErrorAction SilentlyContinue | ForEach-Object { $_.RoleName })
    if ($liveRoles -notcontains 'Application Mail.Send') { New-ManagementRoleAssignment -App 'ffffffff-0000-0000-0000-000000000004' -Role 'Application Mail.Send' -CustomResourceScope 'AppRBAC_hr-alerts' }
}

# Step 4: VERIFY - expect InScope = True for an in-scope mailbox
Test-ServicePrincipalAuthorization -Identity 'ffffffff-0000-0000-0000-000000000004' -Resource 'hr-alerts@contoso.com' | Format-Table
# Also confirm the application itself still works before continuing.

# ---- Step 5-6: CUTOVER. Verifies EVERY migrated role is live and InScope, revokes the
# tenant-wide grants (each one checked), then removes the legacy policy (Remove-* prompts
# for confirmation). Entra + RBAC grants are a union - scoping only takes effect after
# the tenant-wide grant is revoked. The policy is only removed if ALL revocations succeed.
$expectedRoles = @('Application Mail.Send')
$cutoverTestMailbox = 'hr-alerts@contoso.com'   # in-scope member found by the audit; substitute if needed
$auth = if ($cutoverTestMailbox -notlike '<*') { @(Test-ServicePrincipalAuthorization -Identity 'ffffffff-0000-0000-0000-000000000004' -Resource $cutoverTestMailbox -ErrorAction SilentlyContinue) } else { @() }
$missingRoles = @($expectedRoles | Where-Object { $role = $_; -not ($auth | Where-Object { $_.RoleName -eq $role -and $_.InScope }) })
if ($missingRoles.Count -gt 0) {
    Write-Warning "HR Notification Bot: roles not verified InScope for $cutoverTestMailbox (missing: $($missingRoles -join ', ')) - cutover skipped. Run steps 1-4; if they succeeded, try a different in-scope mailbox."
} elseif (@((Get-MgContext).Scopes) -notcontains 'AppRoleAssignment.ReadWrite.All') {
    Write-Warning 'Graph session lacks AppRoleAssignment.ReadWrite.All - run: Connect-MgGraph -Scopes AppRoleAssignment.ReadWrite.All'
} else {
    $revokeFailed = $false
    try { Invoke-MgGraphRequest -Method DELETE -Uri 'v1.0/servicePrincipals/abababab-0000-0000-0000-000000000004/appRoleAssignments/assign-000004-1' -ErrorAction Stop; Write-Host 'Revoked Graph:Mail.Send' } catch { $revokeFailed = $true; Write-Warning "Revocation FAILED (Graph:Mail.Send): $_" }
    if ($revokeFailed) {
        Write-Warning 'One or more revocations FAILED - the legacy policy was NOT removed. Fix the errors above and rerun this cutover block.'
    } else {
        Write-Host 'Tenant-wide grants revoked. Exchange caches app permissions 30 min - 2 h; re-test the app.'
        Write-Host 'Then confirm removal of the legacy policy:'
        Remove-ApplicationAccessPolicy -Identity '11111111-1111-1111-1111-111111111111\ffffffff-0000-0000-0000-000000000004:S-1-5-21-1004336348-1177238915-682003330-4061;eeeeeeee-0000-0000-0000-000000000001'
    }
}
# Hygiene afterwards: App registrations > API permissions - delete the revoked rows
# ('not granted' leftovers). Removing entries there does NOT revoke access by itself.

# ==== Invoice Mailer (ffffffff-0000-0000-0000-000000000002) -> AP Team [RestrictAccess] ====
# Steps 1-4 are additive and do not change existing access.

# Step 1: Exchange service principal pointer (idempotent)
# Already exists in Exchange Online - nothing to do.

# Step 2: Management scope (idempotent; warns if the name is taken by a different filter)
# NOTE: MemberOfGroup covers DIRECT members only - nested group members are out of scope.
$scope = $null
$scopeConflict = $false
$scope = Get-ManagementScope -Identity 'AppRBAC_ap-team' -ErrorAction SilentlyContinue
if (-not $scope) {
    New-ManagementScope -Name 'AppRBAC_ap-team' -RecipientRestrictionFilter "MemberOfGroup -eq 'CN=AP Team,OU=contoso.onmicrosoft.com,OU=Microsoft Exchange Hosted Organizations,DC=EURPR01A001,DC=prod,DC=outlook,DC=com'"
} elseif ($scope.RecipientFilter -notlike '*CN=AP Team,OU=contoso.onmicrosoft.com,OU=Microsoft Exchange Hosted Organizations,DC=EURPR01A001,DC=prod,DC=outlook,DC=com*') {
    $scopeConflict = $true
    Write-Warning "Scope 'AppRBAC_ap-team' already exists with a DIFFERENT filter: $($scope.RecipientFilter) - role assignments skipped; resolve the conflict first."
}

# Step 3: RBAC role assignments (idempotent - skips roles already assigned; skipped
# entirely on a scope-name conflict)
if (-not $scopeConflict) {
    $liveRoles = @(Test-ServicePrincipalAuthorization -Identity 'ffffffff-0000-0000-0000-000000000002' -ErrorAction SilentlyContinue | ForEach-Object { $_.RoleName })
    if ($liveRoles -notcontains 'Application Mail.Send') { New-ManagementRoleAssignment -App 'ffffffff-0000-0000-0000-000000000002' -Role 'Application Mail.Send' -CustomResourceScope 'AppRBAC_ap-team' }
}

# Step 4: VERIFY - expect InScope = True for an in-scope mailbox
Test-ServicePrincipalAuthorization -Identity 'ffffffff-0000-0000-0000-000000000002' -Resource 'member2@contoso.com' | Format-Table
# Also confirm the application itself still works before continuing.

# ---- Step 5-6: CUTOVER. Verifies EVERY migrated role is live and InScope, revokes the
# tenant-wide grants (each one checked), then removes the legacy policy (Remove-* prompts
# for confirmation). Entra + RBAC grants are a union - scoping only takes effect after
# the tenant-wide grant is revoked. The policy is only removed if ALL revocations succeed.
$expectedRoles = @('Application Mail.Send')
$cutoverTestMailbox = 'member2@contoso.com'   # in-scope member found by the audit; substitute if needed
$auth = if ($cutoverTestMailbox -notlike '<*') { @(Test-ServicePrincipalAuthorization -Identity 'ffffffff-0000-0000-0000-000000000002' -Resource $cutoverTestMailbox -ErrorAction SilentlyContinue) } else { @() }
$missingRoles = @($expectedRoles | Where-Object { $role = $_; -not ($auth | Where-Object { $_.RoleName -eq $role -and $_.InScope }) })
if ($missingRoles.Count -gt 0) {
    Write-Warning "Invoice Mailer: roles not verified InScope for $cutoverTestMailbox (missing: $($missingRoles -join ', ')) - cutover skipped. Run steps 1-4; if they succeeded, try a different in-scope mailbox."
} elseif (@((Get-MgContext).Scopes) -notcontains 'AppRoleAssignment.ReadWrite.All') {
    Write-Warning 'Graph session lacks AppRoleAssignment.ReadWrite.All - run: Connect-MgGraph -Scopes AppRoleAssignment.ReadWrite.All'
} else {
    $revokeFailed = $false
    try { Invoke-MgGraphRequest -Method DELETE -Uri 'v1.0/servicePrincipals/abababab-0000-0000-0000-000000000002/appRoleAssignments/assign-000002-1' -ErrorAction Stop; Write-Host 'Revoked Graph:Mail.Send' } catch { $revokeFailed = $true; Write-Warning "Revocation FAILED (Graph:Mail.Send): $_" }
    if ($revokeFailed) {
        Write-Warning 'One or more revocations FAILED - the legacy policy was NOT removed. Fix the errors above and rerun this cutover block.'
    } else {
        Write-Host 'Tenant-wide grants revoked. Exchange caches app permissions 30 min - 2 h; re-test the app.'
        Write-Host 'Then confirm removal of the legacy policy:'
        Remove-ApplicationAccessPolicy -Identity '11111111-1111-1111-1111-111111111111\ffffffff-0000-0000-0000-000000000002:S-1-5-21-1004336348-1177238915-682003330-3061;cccccccc-0000-0000-0000-000000000002'
    }
}
# Hygiene afterwards: App registrations > API permissions - delete the revoked rows
# ('not granted' leftovers). Removing entries there does NOT revoke access by itself.

# ==== Legacy EWS Archiver (ffffffff-0000-0000-0000-000000000003) -> Records Management [RestrictAccess] ====
# Steps 1-4 are additive and do not change existing access.

# Step 1: Exchange service principal pointer (idempotent)
if (-not (Get-ServicePrincipal -Identity 'ffffffff-0000-0000-0000-000000000003' -ErrorAction SilentlyContinue)) {
    New-ServicePrincipal -AppId 'ffffffff-0000-0000-0000-000000000003' -ObjectId 'abababab-0000-0000-0000-000000000003' -DisplayName 'Legacy EWS Archiver'
}

# Step 2: Management scope (idempotent; warns if the name is taken by a different filter)
# NOTE: MemberOfGroup covers DIRECT members only - nested group members are out of scope.
$scope = $null
$scopeConflict = $false
$scope = Get-ManagementScope -Identity 'AppRBAC_records-mgmt' -ErrorAction SilentlyContinue
if (-not $scope) {
    New-ManagementScope -Name 'AppRBAC_records-mgmt' -RecipientRestrictionFilter "MemberOfGroup -eq 'CN=Records Management,OU=contoso.onmicrosoft.com,OU=Microsoft Exchange Hosted Organizations,DC=EURPR01A001,DC=prod,DC=outlook,DC=com'"
} elseif ($scope.RecipientFilter -notlike '*CN=Records Management,OU=contoso.onmicrosoft.com,OU=Microsoft Exchange Hosted Organizations,DC=EURPR01A001,DC=prod,DC=outlook,DC=com*') {
    $scopeConflict = $true
    Write-Warning "Scope 'AppRBAC_records-mgmt' already exists with a DIFFERENT filter: $($scope.RecipientFilter) - role assignments skipped; resolve the conflict first."
}

# Step 3: RBAC role assignments (idempotent - skips roles already assigned; skipped
# entirely on a scope-name conflict)
if (-not $scopeConflict) {
    $liveRoles = @(Test-ServicePrincipalAuthorization -Identity 'ffffffff-0000-0000-0000-000000000003' -ErrorAction SilentlyContinue | ForEach-Object { $_.RoleName })
    if ($liveRoles -notcontains 'Application EWS.AccessAsApp') { New-ManagementRoleAssignment -App 'ffffffff-0000-0000-0000-000000000003' -Role 'Application EWS.AccessAsApp' -CustomResourceScope 'AppRBAC_records-mgmt' }
    if ($liveRoles -notcontains 'Application Mail.Read') { New-ManagementRoleAssignment -App 'ffffffff-0000-0000-0000-000000000003' -Role 'Application Mail.Read' -CustomResourceScope 'AppRBAC_records-mgmt' }
}
# WARNING: EWS is blocked for non-Microsoft apps starting Oct 1, 2026 (EWSAllowedAppIDs allowlist) and removed after Apr 2027. Plan a Microsoft Graph migration for this app.

# Step 4: VERIFY - expect InScope = True for an in-scope mailbox
Test-ServicePrincipalAuthorization -Identity 'ffffffff-0000-0000-0000-000000000003' -Resource 'member3@contoso.com' | Format-Table
# Also confirm the application itself still works before continuing.

# ---- Step 5-6: CUTOVER. Verifies EVERY migrated role is live and InScope, revokes the
# tenant-wide grants (each one checked), then removes the legacy policy (Remove-* prompts
# for confirmation). Entra + RBAC grants are a union - scoping only takes effect after
# the tenant-wide grant is revoked. The policy is only removed if ALL revocations succeed.
$expectedRoles = @('Application EWS.AccessAsApp', 'Application Mail.Read')
$cutoverTestMailbox = 'member3@contoso.com'   # in-scope member found by the audit; substitute if needed
$auth = if ($cutoverTestMailbox -notlike '<*') { @(Test-ServicePrincipalAuthorization -Identity 'ffffffff-0000-0000-0000-000000000003' -Resource $cutoverTestMailbox -ErrorAction SilentlyContinue) } else { @() }
$missingRoles = @($expectedRoles | Where-Object { $role = $_; -not ($auth | Where-Object { $_.RoleName -eq $role -and $_.InScope }) })
if ($missingRoles.Count -gt 0) {
    Write-Warning "Legacy EWS Archiver: roles not verified InScope for $cutoverTestMailbox (missing: $($missingRoles -join ', ')) - cutover skipped. Run steps 1-4; if they succeeded, try a different in-scope mailbox."
} elseif (@((Get-MgContext).Scopes) -notcontains 'AppRoleAssignment.ReadWrite.All') {
    Write-Warning 'Graph session lacks AppRoleAssignment.ReadWrite.All - run: Connect-MgGraph -Scopes AppRoleAssignment.ReadWrite.All'
} else {
    $revokeFailed = $false
    # KEEP (NOT replaced by RBAC - do not revoke): EXO:IMAP.AccessAsApp
    try { Invoke-MgGraphRequest -Method DELETE -Uri 'v1.0/servicePrincipals/abababab-0000-0000-0000-000000000003/appRoleAssignments/assign-000003-1' -ErrorAction Stop; Write-Host 'Revoked EXO:full_access_as_app' } catch { $revokeFailed = $true; Write-Warning "Revocation FAILED (EXO:full_access_as_app): $_" }
    try { Invoke-MgGraphRequest -Method DELETE -Uri 'v1.0/servicePrincipals/abababab-0000-0000-0000-000000000003/appRoleAssignments/assign-000003-3' -ErrorAction Stop; Write-Host 'Revoked Graph:Mail.Read' } catch { $revokeFailed = $true; Write-Warning "Revocation FAILED (Graph:Mail.Read): $_" }
    if ($revokeFailed) {
        Write-Warning 'One or more revocations FAILED - the legacy policy was NOT removed. Fix the errors above and rerun this cutover block.'
    } else {
        Write-Host 'Tenant-wide grants revoked. Exchange caches app permissions 30 min - 2 h; re-test the app.'
        Write-Host 'Then confirm removal of the legacy policy:'
        Remove-ApplicationAccessPolicy -Identity '11111111-1111-1111-1111-111111111111\ffffffff-0000-0000-0000-000000000003:S-1-5-21-1004336348-1177238915-682003330-1688;cccccccc-0000-0000-0000-000000000003'
    }
}
# Hygiene afterwards: App registrations > API permissions - delete the revoked rows
# ('not granted' leftovers). Removing entries there does NOT revoke access by itself.

# ==== Ticketing Mail Connector (ffffffff-0000-0000-0000-000000000008) -> Helpdesk Intake [RestrictAccess] ====
# Steps 1-4 are additive and do not change existing access.

# Step 1: Exchange service principal pointer (idempotent)
# Already exists in Exchange Online - nothing to do.

# Step 2: Management scope (idempotent; warns if the name is taken by a different filter)
# NOTE: MemberOfGroup covers DIRECT members only - nested group members are out of scope.
$scope = $null
$scopeConflict = $false
$scope = Get-ManagementScope -Identity 'AppRBAC_helpdesk-intake' -ErrorAction SilentlyContinue
if (-not $scope) {
    New-ManagementScope -Name 'AppRBAC_helpdesk-intake' -RecipientRestrictionFilter "MemberOfGroup -eq 'CN=Helpdesk Intake,OU=contoso.onmicrosoft.com,OU=Microsoft Exchange Hosted Organizations,DC=EURPR01A001,DC=prod,DC=outlook,DC=com'"
} elseif ($scope.RecipientFilter -notlike '*CN=Helpdesk Intake,OU=contoso.onmicrosoft.com,OU=Microsoft Exchange Hosted Organizations,DC=EURPR01A001,DC=prod,DC=outlook,DC=com*') {
    $scopeConflict = $true
    Write-Warning "Scope 'AppRBAC_helpdesk-intake' already exists with a DIFFERENT filter: $($scope.RecipientFilter) - role assignments skipped; resolve the conflict first."
}

# Step 3: RBAC role assignments (idempotent - skips roles already assigned; skipped
# entirely on a scope-name conflict)
if (-not $scopeConflict) {
    $liveRoles = @(Test-ServicePrincipalAuthorization -Identity 'ffffffff-0000-0000-0000-000000000008' -ErrorAction SilentlyContinue | ForEach-Object { $_.RoleName })
    if ($liveRoles -notcontains 'Application Mail.Read') { New-ManagementRoleAssignment -App 'ffffffff-0000-0000-0000-000000000008' -Role 'Application Mail.Read' -CustomResourceScope 'AppRBAC_helpdesk-intake' }
    if ($liveRoles -notcontains 'Application Mail.Send') { New-ManagementRoleAssignment -App 'ffffffff-0000-0000-0000-000000000008' -Role 'Application Mail.Send' -CustomResourceScope 'AppRBAC_helpdesk-intake' }
}

# Step 4: VERIFY - expect InScope = True for an in-scope mailbox
Test-ServicePrincipalAuthorization -Identity 'ffffffff-0000-0000-0000-000000000008' -Resource 'member7@contoso.com' | Format-Table
# Also confirm the application itself still works before continuing.

# ---- Step 5-6: CUTOVER. Verifies EVERY migrated role is live and InScope, revokes the
# tenant-wide grants (each one checked), then removes the legacy policy (Remove-* prompts
# for confirmation). Entra + RBAC grants are a union - scoping only takes effect after
# the tenant-wide grant is revoked. The policy is only removed if ALL revocations succeed.
$expectedRoles = @('Application Mail.Read', 'Application Mail.Send')
$cutoverTestMailbox = 'member7@contoso.com'   # in-scope member found by the audit; substitute if needed
$auth = if ($cutoverTestMailbox -notlike '<*') { @(Test-ServicePrincipalAuthorization -Identity 'ffffffff-0000-0000-0000-000000000008' -Resource $cutoverTestMailbox -ErrorAction SilentlyContinue) } else { @() }
$missingRoles = @($expectedRoles | Where-Object { $role = $_; -not ($auth | Where-Object { $_.RoleName -eq $role -and $_.InScope }) })
if ($missingRoles.Count -gt 0) {
    Write-Warning "Ticketing Mail Connector: roles not verified InScope for $cutoverTestMailbox (missing: $($missingRoles -join ', ')) - cutover skipped. Run steps 1-4; if they succeeded, try a different in-scope mailbox."
} elseif (@((Get-MgContext).Scopes) -notcontains 'AppRoleAssignment.ReadWrite.All') {
    Write-Warning 'Graph session lacks AppRoleAssignment.ReadWrite.All - run: Connect-MgGraph -Scopes AppRoleAssignment.ReadWrite.All'
} else {
    $revokeFailed = $false
    try { Invoke-MgGraphRequest -Method DELETE -Uri 'v1.0/servicePrincipals/abababab-0000-0000-0000-000000000008/appRoleAssignments/assign-000008-1' -ErrorAction Stop; Write-Host 'Revoked Graph:Mail.Read' } catch { $revokeFailed = $true; Write-Warning "Revocation FAILED (Graph:Mail.Read): $_" }
    try { Invoke-MgGraphRequest -Method DELETE -Uri 'v1.0/servicePrincipals/abababab-0000-0000-0000-000000000008/appRoleAssignments/assign-000008-2' -ErrorAction Stop; Write-Host 'Revoked Graph:Mail.Send' } catch { $revokeFailed = $true; Write-Warning "Revocation FAILED (Graph:Mail.Send): $_" }
    if ($revokeFailed) {
        Write-Warning 'One or more revocations FAILED - the legacy policy was NOT removed. Fix the errors above and rerun this cutover block.'
    } else {
        Write-Host 'Tenant-wide grants revoked. Exchange caches app permissions 30 min - 2 h; re-test the app.'
        Write-Host 'Then confirm removal of the legacy policy:'
        Remove-ApplicationAccessPolicy -Identity '11111111-1111-1111-1111-111111111111\ffffffff-0000-0000-0000-000000000008:S-1-5-21-1004336348-1177238915-682003330-5287;cccccccc-0000-0000-0000-000000000007'
    }
}
# Hygiene afterwards: App registrations > API permissions - delete the revoked rows
# ('not granted' leftovers). Removing entries there does NOT revoke access by itself.

# No Exchange application permissions -> this policy has no effect today.
# KEEP (not Exchange-related): Graph:DeviceManagementManagedDevices.Read.All
# Removing the policy changes no behavior (confirmation prompt follows). If the app is
# granted Exchange permissions later, scope it with RBAC at that time.
Remove-ApplicationAccessPolicy -Identity '11111111-1111-1111-1111-111111111111\ffffffff-0000-0000-0000-000000000010:S-1-5-21-1004336348-1177238915-682003330-4478;cccccccc-0000-0000-0000-000000000009'

# App registration exists but there is no enterprise application (service principal),
# so the app cannot get app-only tokens and the policy is dormant.
# If the app should work: New-MgServicePrincipal -AppId 'ffffffff-0000-0000-0000-000000000017'   # needs Application.ReadWrite.All
# then rerun this audit. If the app is retired, remove the policy and the registration.

# ==== Bulk Invite Tool (ffffffff-0000-0000-0000-000000000016) -> Invite Scope [RestrictAccess] ====
# Steps 1-4 are additive and do not change existing access.
# NOTE: this app cannot currently obtain tokens (deactivated / sign-in disabled).
# Steps 1-3 can be staged and the Step 4 test cmdlet still evaluates, but the app
# itself cannot be end-to-end tested until re-enabled. If it is being RETIRED, skip
# migration: revoke its grants and remove the policy instead.

# Step 1: Exchange service principal pointer (idempotent)
if (-not (Get-ServicePrincipal -Identity 'ffffffff-0000-0000-0000-000000000016' -ErrorAction SilentlyContinue)) {
    New-ServicePrincipal -AppId 'ffffffff-0000-0000-0000-000000000016' -ObjectId 'abababab-0000-0000-0000-000000000016' -DisplayName 'Bulk Invite Tool'
}

# Step 2: Management scope (idempotent; warns if the name is taken by a different filter)
# NOTE: MemberOfGroup covers DIRECT members only - nested group members are out of scope.
$scope = $null
$scopeConflict = $false
$scope = Get-ManagementScope -Identity 'AppRBAC_invite-scope' -ErrorAction SilentlyContinue
if (-not $scope) {
    New-ManagementScope -Name 'AppRBAC_invite-scope' -RecipientRestrictionFilter "MemberOfGroup -eq 'CN=Invite Scope,OU=contoso.onmicrosoft.com,OU=Microsoft Exchange Hosted Organizations,DC=EURPR01A001,DC=prod,DC=outlook,DC=com'"
} elseif ($scope.RecipientFilter -notlike '*CN=Invite Scope,OU=contoso.onmicrosoft.com,OU=Microsoft Exchange Hosted Organizations,DC=EURPR01A001,DC=prod,DC=outlook,DC=com*') {
    $scopeConflict = $true
    Write-Warning "Scope 'AppRBAC_invite-scope' already exists with a DIFFERENT filter: $($scope.RecipientFilter) - role assignments skipped; resolve the conflict first."
}

# Step 3: RBAC role assignments (idempotent - skips roles already assigned; skipped
# entirely on a scope-name conflict)
if (-not $scopeConflict) {
    $liveRoles = @(Test-ServicePrincipalAuthorization -Identity 'ffffffff-0000-0000-0000-000000000016' -ErrorAction SilentlyContinue | ForEach-Object { $_.RoleName })
    if ($liveRoles -notcontains 'Application Mail.Send') { New-ManagementRoleAssignment -App 'ffffffff-0000-0000-0000-000000000016' -Role 'Application Mail.Send' -CustomResourceScope 'AppRBAC_invite-scope' }
}

# Step 4: VERIFY - expect InScope = True for an in-scope mailbox
Test-ServicePrincipalAuthorization -Identity 'ffffffff-0000-0000-0000-000000000016' -Resource 'member15@contoso.com' | Format-Table
# Also confirm the application itself still works before continuing.

# ---- Step 5-6: CUTOVER. Verifies EVERY migrated role is live and InScope, revokes the
# tenant-wide grants (each one checked), then removes the legacy policy (Remove-* prompts
# for confirmation). Entra + RBAC grants are a union - scoping only takes effect after
# the tenant-wide grant is revoked. The policy is only removed if ALL revocations succeed.
$expectedRoles = @('Application Mail.Send')
$cutoverTestMailbox = 'member15@contoso.com'   # in-scope member found by the audit; substitute if needed
$auth = if ($cutoverTestMailbox -notlike '<*') { @(Test-ServicePrincipalAuthorization -Identity 'ffffffff-0000-0000-0000-000000000016' -Resource $cutoverTestMailbox -ErrorAction SilentlyContinue) } else { @() }
$missingRoles = @($expectedRoles | Where-Object { $role = $_; -not ($auth | Where-Object { $_.RoleName -eq $role -and $_.InScope }) })
if ($missingRoles.Count -gt 0) {
    Write-Warning "Bulk Invite Tool: roles not verified InScope for $cutoverTestMailbox (missing: $($missingRoles -join ', ')) - cutover skipped. Run steps 1-4; if they succeeded, try a different in-scope mailbox."
} elseif (@((Get-MgContext).Scopes) -notcontains 'AppRoleAssignment.ReadWrite.All') {
    Write-Warning 'Graph session lacks AppRoleAssignment.ReadWrite.All - run: Connect-MgGraph -Scopes AppRoleAssignment.ReadWrite.All'
} else {
    $revokeFailed = $false
    try { Invoke-MgGraphRequest -Method DELETE -Uri 'v1.0/servicePrincipals/abababab-0000-0000-0000-000000000016/appRoleAssignments/assign-000016-1' -ErrorAction Stop; Write-Host 'Revoked Graph:Mail.Send' } catch { $revokeFailed = $true; Write-Warning "Revocation FAILED (Graph:Mail.Send): $_" }
    if ($revokeFailed) {
        Write-Warning 'One or more revocations FAILED - the legacy policy was NOT removed. Fix the errors above and rerun this cutover block.'
    } else {
        Write-Host 'Tenant-wide grants revoked. Exchange caches app permissions 30 min - 2 h; re-test the app.'
        Write-Host 'Then confirm removal of the legacy policy:'
        Remove-ApplicationAccessPolicy -Identity '11111111-1111-1111-1111-111111111111\ffffffff-0000-0000-0000-000000000016:S-1-5-21-1004336348-1177238915-682003330-4608;cccccccc-0000-0000-0000-000000000015'
    }
}
# Hygiene afterwards: App registrations > API permissions - delete the revoked rows
# ('not granted' leftovers). Removing entries there does NOT revoke access by itself.

# ==== Contractor Portal (ffffffff-0000-0000-0000-000000000015) -> Contractor Mailboxes [RestrictAccess] ====
# Steps 1-4 are additive and do not change existing access.
# NOTE: this app cannot currently obtain tokens (deactivated / sign-in disabled).
# Steps 1-3 can be staged and the Step 4 test cmdlet still evaluates, but the app
# itself cannot be end-to-end tested until re-enabled. If it is being RETIRED, skip
# migration: revoke its grants and remove the policy instead.

# Step 1: Exchange service principal pointer (idempotent)
if (-not (Get-ServicePrincipal -Identity 'ffffffff-0000-0000-0000-000000000015' -ErrorAction SilentlyContinue)) {
    New-ServicePrincipal -AppId 'ffffffff-0000-0000-0000-000000000015' -ObjectId 'abababab-0000-0000-0000-000000000015' -DisplayName 'Contractor Portal'
}

# Step 2: Management scope (idempotent; warns if the name is taken by a different filter)
# NOTE: MemberOfGroup covers DIRECT members only - nested group members are out of scope.
$scope = $null
$scopeConflict = $false
$scope = Get-ManagementScope -Identity 'AppRBAC_contractor-mbx' -ErrorAction SilentlyContinue
if (-not $scope) {
    New-ManagementScope -Name 'AppRBAC_contractor-mbx' -RecipientRestrictionFilter "MemberOfGroup -eq 'CN=Contractor Mailboxes,OU=contoso.onmicrosoft.com,OU=Microsoft Exchange Hosted Organizations,DC=EURPR01A001,DC=prod,DC=outlook,DC=com'"
} elseif ($scope.RecipientFilter -notlike '*CN=Contractor Mailboxes,OU=contoso.onmicrosoft.com,OU=Microsoft Exchange Hosted Organizations,DC=EURPR01A001,DC=prod,DC=outlook,DC=com*') {
    $scopeConflict = $true
    Write-Warning "Scope 'AppRBAC_contractor-mbx' already exists with a DIFFERENT filter: $($scope.RecipientFilter) - role assignments skipped; resolve the conflict first."
}

# Step 3: RBAC role assignments (idempotent - skips roles already assigned; skipped
# entirely on a scope-name conflict)
if (-not $scopeConflict) {
    $liveRoles = @(Test-ServicePrincipalAuthorization -Identity 'ffffffff-0000-0000-0000-000000000015' -ErrorAction SilentlyContinue | ForEach-Object { $_.RoleName })
    if ($liveRoles -notcontains 'Application Mail.Read') { New-ManagementRoleAssignment -App 'ffffffff-0000-0000-0000-000000000015' -Role 'Application Mail.Read' -CustomResourceScope 'AppRBAC_contractor-mbx' }
}

# Step 4: VERIFY - expect InScope = True for an in-scope mailbox
Test-ServicePrincipalAuthorization -Identity 'ffffffff-0000-0000-0000-000000000015' -Resource 'member14@contoso.com' | Format-Table
# Also confirm the application itself still works before continuing.

# ---- Step 5-6: CUTOVER. Verifies EVERY migrated role is live and InScope, revokes the
# tenant-wide grants (each one checked), then removes the legacy policy (Remove-* prompts
# for confirmation). Entra + RBAC grants are a union - scoping only takes effect after
# the tenant-wide grant is revoked. The policy is only removed if ALL revocations succeed.
$expectedRoles = @('Application Mail.Read')
$cutoverTestMailbox = 'member14@contoso.com'   # in-scope member found by the audit; substitute if needed
$auth = if ($cutoverTestMailbox -notlike '<*') { @(Test-ServicePrincipalAuthorization -Identity 'ffffffff-0000-0000-0000-000000000015' -Resource $cutoverTestMailbox -ErrorAction SilentlyContinue) } else { @() }
$missingRoles = @($expectedRoles | Where-Object { $role = $_; -not ($auth | Where-Object { $_.RoleName -eq $role -and $_.InScope }) })
if ($missingRoles.Count -gt 0) {
    Write-Warning "Contractor Portal: roles not verified InScope for $cutoverTestMailbox (missing: $($missingRoles -join ', ')) - cutover skipped. Run steps 1-4; if they succeeded, try a different in-scope mailbox."
} elseif (@((Get-MgContext).Scopes) -notcontains 'AppRoleAssignment.ReadWrite.All') {
    Write-Warning 'Graph session lacks AppRoleAssignment.ReadWrite.All - run: Connect-MgGraph -Scopes AppRoleAssignment.ReadWrite.All'
} else {
    $revokeFailed = $false
    try { Invoke-MgGraphRequest -Method DELETE -Uri 'v1.0/servicePrincipals/abababab-0000-0000-0000-000000000015/appRoleAssignments/assign-000015-1' -ErrorAction Stop; Write-Host 'Revoked Graph:Mail.Read' } catch { $revokeFailed = $true; Write-Warning "Revocation FAILED (Graph:Mail.Read): $_" }
    if ($revokeFailed) {
        Write-Warning 'One or more revocations FAILED - the legacy policy was NOT removed. Fix the errors above and rerun this cutover block.'
    } else {
        Write-Host 'Tenant-wide grants revoked. Exchange caches app permissions 30 min - 2 h; re-test the app.'
        Write-Host 'Then confirm removal of the legacy policy:'
        Remove-ApplicationAccessPolicy -Identity '11111111-1111-1111-1111-111111111111\ffffffff-0000-0000-0000-000000000015:S-1-5-21-1004336348-1177238915-682003330-1062;cccccccc-0000-0000-0000-000000000014'
    }
}
# Hygiene afterwards: App registrations > API permissions - delete the revoked rows
# ('not granted' leftovers). Removing entries there does NOT revoke access by itself.

# ==== CRM Mailbox Sync (ffffffff-0000-0000-0000-000000000005) -> Sales Department [RestrictAccess] ====
# Steps 1-4 are additive and do not change existing access.

# Step 1: Exchange service principal pointer (idempotent)
if (-not (Get-ServicePrincipal -Identity 'ffffffff-0000-0000-0000-000000000005' -ErrorAction SilentlyContinue)) {
    New-ServicePrincipal -AppId 'ffffffff-0000-0000-0000-000000000005' -ObjectId 'abababab-0000-0000-0000-000000000005' -DisplayName 'CRM Mailbox Sync'
}

# Step 2: Management scope (idempotent; warns if the name is taken by a different filter)
# NOTE: MemberOfGroup covers DIRECT members only - nested group members are out of scope.
$scope = $null
$scopeConflict = $false
$scope = Get-ManagementScope -Identity 'AppRBAC_sales-dept' -ErrorAction SilentlyContinue
if (-not $scope) {
    New-ManagementScope -Name 'AppRBAC_sales-dept' -RecipientRestrictionFilter "MemberOfGroup -eq 'CN=Sales Department,OU=contoso.onmicrosoft.com,OU=Microsoft Exchange Hosted Organizations,DC=EURPR01A001,DC=prod,DC=outlook,DC=com'"
} elseif ($scope.RecipientFilter -notlike '*CN=Sales Department,OU=contoso.onmicrosoft.com,OU=Microsoft Exchange Hosted Organizations,DC=EURPR01A001,DC=prod,DC=outlook,DC=com*') {
    $scopeConflict = $true
    Write-Warning "Scope 'AppRBAC_sales-dept' already exists with a DIFFERENT filter: $($scope.RecipientFilter) - role assignments skipped; resolve the conflict first."
}

# Step 3: RBAC role assignments (idempotent - skips roles already assigned; skipped
# entirely on a scope-name conflict)
if (-not $scopeConflict) {
    $liveRoles = @(Test-ServicePrincipalAuthorization -Identity 'ffffffff-0000-0000-0000-000000000005' -ErrorAction SilentlyContinue | ForEach-Object { $_.RoleName })
    if ($liveRoles -notcontains 'Application Mail.ReadWrite') { New-ManagementRoleAssignment -App 'ffffffff-0000-0000-0000-000000000005' -Role 'Application Mail.ReadWrite' -CustomResourceScope 'AppRBAC_sales-dept' }
}

# Step 4: VERIFY - expect InScope = True for an in-scope mailbox
Test-ServicePrincipalAuthorization -Identity 'ffffffff-0000-0000-0000-000000000005' -Resource 'member4@contoso.com' | Format-Table
# Also confirm the application itself still works before continuing.

# ---- Step 5-6: CUTOVER. Verifies EVERY migrated role is live and InScope, revokes the
# tenant-wide grants (each one checked), then removes the legacy policy (Remove-* prompts
# for confirmation). Entra + RBAC grants are a union - scoping only takes effect after
# the tenant-wide grant is revoked. The policy is only removed if ALL revocations succeed.
$expectedRoles = @('Application Mail.ReadWrite')
$cutoverTestMailbox = 'member4@contoso.com'   # in-scope member found by the audit; substitute if needed
$auth = if ($cutoverTestMailbox -notlike '<*') { @(Test-ServicePrincipalAuthorization -Identity 'ffffffff-0000-0000-0000-000000000005' -Resource $cutoverTestMailbox -ErrorAction SilentlyContinue) } else { @() }
$missingRoles = @($expectedRoles | Where-Object { $role = $_; -not ($auth | Where-Object { $_.RoleName -eq $role -and $_.InScope }) })
$nestedGroupsHandled = $false   # set to $true after adding nested-group members DIRECTLY to the group
if ($missingRoles.Count -gt 0) {
    Write-Warning "CRM Mailbox Sync: roles not verified InScope for $cutoverTestMailbox (missing: $($missingRoles -join ', ')) - cutover skipped. Run steps 1-4; if they succeeded, try a different in-scope mailbox."
} elseif (@((Get-MgContext).Scopes) -notcontains 'AppRoleAssignment.ReadWrite.All') {
    Write-Warning 'Graph session lacks AppRoleAssignment.ReadWrite.All - run: Connect-MgGraph -Scopes AppRoleAssignment.ReadWrite.All'
} elseif (-not $nestedGroupsHandled) {
    Write-Warning "CRM Mailbox Sync: the scope covers DIRECT members only - flatten nested groups, then set `$nestedGroupsHandled = `$true and rerun."
} else {
    $revokeFailed = $false
    try { Invoke-MgGraphRequest -Method DELETE -Uri 'v1.0/servicePrincipals/abababab-0000-0000-0000-000000000005/appRoleAssignments/assign-000005-1' -ErrorAction Stop; Write-Host 'Revoked Graph:Mail.ReadWrite' } catch { $revokeFailed = $true; Write-Warning "Revocation FAILED (Graph:Mail.ReadWrite): $_" }
    if ($revokeFailed) {
        Write-Warning 'One or more revocations FAILED - the legacy policy was NOT removed. Fix the errors above and rerun this cutover block.'
    } else {
        Write-Host 'Tenant-wide grants revoked. Exchange caches app permissions 30 min - 2 h; re-test the app.'
        Write-Host 'Then confirm removal of the legacy policy:'
        Remove-ApplicationAccessPolicy -Identity '11111111-1111-1111-1111-111111111111\ffffffff-0000-0000-0000-000000000005:S-1-5-21-1004336348-1177238915-682003330-1765;cccccccc-0000-0000-0000-000000000004'
    }
}
# Hygiene afterwards: App registrations > API permissions - delete the revoked rows
# ('not granted' leftovers). Removing entries there does NOT revoke access by itself.

# ==== Marketing Blaster 2019 (ffffffff-0000-0000-0000-000000000014) -> Marketing Lists [RestrictAccess] ====
# Steps 1-4 are additive and do not change existing access.
# NOTE: this app cannot currently obtain tokens (deactivated / sign-in disabled).
# Steps 1-3 can be staged and the Step 4 test cmdlet still evaluates, but the app
# itself cannot be end-to-end tested until re-enabled. If it is being RETIRED, skip
# migration: revoke its grants and remove the policy instead.

# Step 1: Exchange service principal pointer (idempotent)
if (-not (Get-ServicePrincipal -Identity 'ffffffff-0000-0000-0000-000000000014' -ErrorAction SilentlyContinue)) {
    New-ServicePrincipal -AppId 'ffffffff-0000-0000-0000-000000000014' -ObjectId 'abababab-0000-0000-0000-000000000014' -DisplayName 'Marketing Blaster 2019'
}

# Step 2: Management scope (idempotent; warns if the name is taken by a different filter)
# NOTE: MemberOfGroup covers DIRECT members only - nested group members are out of scope.
$scope = $null
$scopeConflict = $false
$scope = Get-ManagementScope -Identity 'AppRBAC_marketing-lists' -ErrorAction SilentlyContinue
if (-not $scope) {
    New-ManagementScope -Name 'AppRBAC_marketing-lists' -RecipientRestrictionFilter "MemberOfGroup -eq 'CN=Marketing Lists,OU=contoso.onmicrosoft.com,OU=Microsoft Exchange Hosted Organizations,DC=EURPR01A001,DC=prod,DC=outlook,DC=com'"
} elseif ($scope.RecipientFilter -notlike '*CN=Marketing Lists,OU=contoso.onmicrosoft.com,OU=Microsoft Exchange Hosted Organizations,DC=EURPR01A001,DC=prod,DC=outlook,DC=com*') {
    $scopeConflict = $true
    Write-Warning "Scope 'AppRBAC_marketing-lists' already exists with a DIFFERENT filter: $($scope.RecipientFilter) - role assignments skipped; resolve the conflict first."
}

# Step 3: RBAC role assignments (idempotent - skips roles already assigned; skipped
# entirely on a scope-name conflict)
if (-not $scopeConflict) {
    $liveRoles = @(Test-ServicePrincipalAuthorization -Identity 'ffffffff-0000-0000-0000-000000000014' -ErrorAction SilentlyContinue | ForEach-Object { $_.RoleName })
    if ($liveRoles -notcontains 'Application Mail.Send') { New-ManagementRoleAssignment -App 'ffffffff-0000-0000-0000-000000000014' -Role 'Application Mail.Send' -CustomResourceScope 'AppRBAC_marketing-lists' }
}

# Step 4: VERIFY - expect InScope = True for an in-scope mailbox
Test-ServicePrincipalAuthorization -Identity 'ffffffff-0000-0000-0000-000000000014' -Resource 'member13@contoso.com' | Format-Table
# Also confirm the application itself still works before continuing.

# ---- Step 5-6: CUTOVER. Verifies EVERY migrated role is live and InScope, revokes the
# tenant-wide grants (each one checked), then removes the legacy policy (Remove-* prompts
# for confirmation). Entra + RBAC grants are a union - scoping only takes effect after
# the tenant-wide grant is revoked. The policy is only removed if ALL revocations succeed.
$expectedRoles = @('Application Mail.Send')
$cutoverTestMailbox = 'member13@contoso.com'   # in-scope member found by the audit; substitute if needed
$auth = if ($cutoverTestMailbox -notlike '<*') { @(Test-ServicePrincipalAuthorization -Identity 'ffffffff-0000-0000-0000-000000000014' -Resource $cutoverTestMailbox -ErrorAction SilentlyContinue) } else { @() }
$missingRoles = @($expectedRoles | Where-Object { $role = $_; -not ($auth | Where-Object { $_.RoleName -eq $role -and $_.InScope }) })
if ($missingRoles.Count -gt 0) {
    Write-Warning "Marketing Blaster 2019: roles not verified InScope for $cutoverTestMailbox (missing: $($missingRoles -join ', ')) - cutover skipped. Run steps 1-4; if they succeeded, try a different in-scope mailbox."
} elseif (@((Get-MgContext).Scopes) -notcontains 'AppRoleAssignment.ReadWrite.All') {
    Write-Warning 'Graph session lacks AppRoleAssignment.ReadWrite.All - run: Connect-MgGraph -Scopes AppRoleAssignment.ReadWrite.All'
} else {
    $revokeFailed = $false
    try { Invoke-MgGraphRequest -Method DELETE -Uri 'v1.0/servicePrincipals/abababab-0000-0000-0000-000000000014/appRoleAssignments/assign-000014-1' -ErrorAction Stop; Write-Host 'Revoked Graph:Mail.Send' } catch { $revokeFailed = $true; Write-Warning "Revocation FAILED (Graph:Mail.Send): $_" }
    if ($revokeFailed) {
        Write-Warning 'One or more revocations FAILED - the legacy policy was NOT removed. Fix the errors above and rerun this cutover block.'
    } else {
        Write-Host 'Tenant-wide grants revoked. Exchange caches app permissions 30 min - 2 h; re-test the app.'
        Write-Host 'Then confirm removal of the legacy policy:'
        Remove-ApplicationAccessPolicy -Identity '11111111-1111-1111-1111-111111111111\ffffffff-0000-0000-0000-000000000014:S-1-5-21-1004336348-1177238915-682003330-6522;cccccccc-0000-0000-0000-000000000013'
    }
}
# Hygiene afterwards: App registrations > API permissions - delete the revoked rows
# ('not granted' leftovers). Removing entries there does NOT revoke access by itself.

# FINISH MIGRATION: RBAC is live (Application Calendars.ReadWrite) and the matching tenant-wide
# Entra grants are gone, so this legacy policy constrains nothing. Verify, then remove it:
Test-ServicePrincipalAuthorization -Identity 'ffffffff-0000-0000-0000-000000000007' -Resource 'member6@contoso.com' | Format-Table   # expect InScope = True
# Confirm the application still works, then remove (confirmation prompt follows):
Remove-ApplicationAccessPolicy -Identity '11111111-1111-1111-1111-111111111111\ffffffff-0000-0000-0000-000000000007:S-1-5-21-1004336348-1177238915-682003330-2391;cccccccc-0000-0000-0000-000000000006'

# !! TARGET MATCHED BY NAME ONLY - VERIFY BEFORE RUNNING !!
# The policy identity carried no object id, so 'Payroll Users' was found by name match.
# Confirm it is the intended target, then uncomment this block.
# ==== Payroll Ingest (ffffffff-0000-0000-0000-000000000006) -> Payroll Users [RestrictAccess] ====
# Steps 1-4 are additive and do not change existing access.

# Step 1: Exchange service principal pointer (idempotent)
# if (-not (Get-ServicePrincipal -Identity 'ffffffff-0000-0000-0000-000000000006' -ErrorAction SilentlyContinue)) {
#     New-ServicePrincipal -AppId 'ffffffff-0000-0000-0000-000000000006' -ObjectId 'abababab-0000-0000-0000-000000000006' -DisplayName 'Payroll Ingest'
# }

# Step 2: Management scope (idempotent; warns if the name is taken by a different filter)
# NOTE: MemberOfGroup covers DIRECT members only - nested group members are out of scope.
# $scope = $null
# $scopeConflict = $false
# $scope = Get-ManagementScope -Identity 'AppRBAC_payroll-users' -ErrorAction SilentlyContinue
# if (-not $scope) {
#     New-ManagementScope -Name 'AppRBAC_payroll-users' -RecipientRestrictionFilter "MemberOfGroup -eq 'CN=Payroll Users,OU=contoso.onmicrosoft.com,OU=Microsoft Exchange Hosted Organizations,DC=EURPR01A001,DC=prod,DC=outlook,DC=com'"
# } elseif ($scope.RecipientFilter -notlike '*CN=Payroll Users,OU=contoso.onmicrosoft.com,OU=Microsoft Exchange Hosted Organizations,DC=EURPR01A001,DC=prod,DC=outlook,DC=com*') {
#     $scopeConflict = $true
#     Write-Warning "Scope 'AppRBAC_payroll-users' already exists with a DIFFERENT filter: $($scope.RecipientFilter) - role assignments skipped; resolve the conflict first."
# }

# Step 3: RBAC role assignments (idempotent - skips roles already assigned; skipped
# entirely on a scope-name conflict)
# if (-not $scopeConflict) {
#     $liveRoles = @(Test-ServicePrincipalAuthorization -Identity 'ffffffff-0000-0000-0000-000000000006' -ErrorAction SilentlyContinue | ForEach-Object { $_.RoleName })
#     if ($liveRoles -notcontains 'Application Mail.Read') { New-ManagementRoleAssignment -App 'ffffffff-0000-0000-0000-000000000006' -Role 'Application Mail.Read' -CustomResourceScope 'AppRBAC_payroll-users' }
# }

# Step 4: VERIFY - expect InScope = True for an in-scope mailbox
# Test-ServicePrincipalAuthorization -Identity 'ffffffff-0000-0000-0000-000000000006' -Resource 'member5@contoso.com' | Format-Table
# Also confirm the application itself still works before continuing.

# ---- Step 5-6: CUTOVER. Verifies EVERY migrated role is live and InScope, revokes the
# tenant-wide grants (each one checked), then removes the legacy policy (Remove-* prompts
# for confirmation). Entra + RBAC grants are a union - scoping only takes effect after
# the tenant-wide grant is revoked. The policy is only removed if ALL revocations succeed.
# $expectedRoles = @('Application Mail.Read')
# $cutoverTestMailbox = 'member5@contoso.com'   # in-scope member found by the audit; substitute if needed
# $auth = if ($cutoverTestMailbox -notlike '<*') { @(Test-ServicePrincipalAuthorization -Identity 'ffffffff-0000-0000-0000-000000000006' -Resource $cutoverTestMailbox -ErrorAction SilentlyContinue) } else { @() }
# $missingRoles = @($expectedRoles | Where-Object { $role = $_; -not ($auth | Where-Object { $_.RoleName -eq $role -and $_.InScope }) })
# if ($missingRoles.Count -gt 0) {
#     Write-Warning "Payroll Ingest: roles not verified InScope for $cutoverTestMailbox (missing: $($missingRoles -join ', ')) - cutover skipped. Run steps 1-4; if they succeeded, try a different in-scope mailbox."
# } elseif (@((Get-MgContext).Scopes) -notcontains 'AppRoleAssignment.ReadWrite.All') {
#     Write-Warning 'Graph session lacks AppRoleAssignment.ReadWrite.All - run: Connect-MgGraph -Scopes AppRoleAssignment.ReadWrite.All'
# } else {
#     $revokeFailed = $false
#     try { Invoke-MgGraphRequest -Method DELETE -Uri 'v1.0/servicePrincipals/abababab-0000-0000-0000-000000000006/appRoleAssignments/assign-000006-1' -ErrorAction Stop; Write-Host 'Revoked Graph:Mail.Read' } catch { $revokeFailed = $true; Write-Warning "Revocation FAILED (Graph:Mail.Read): $_" }
#     if ($revokeFailed) {
#         Write-Warning 'One or more revocations FAILED - the legacy policy was NOT removed. Fix the errors above and rerun this cutover block.'
#     } else {
#         Write-Host 'Tenant-wide grants revoked. Exchange caches app permissions 30 min - 2 h; re-test the app.'
#         Write-Host 'Then confirm removal of the legacy policy:'
#         Remove-ApplicationAccessPolicy -Identity '11111111-1111-1111-1111-111111111111\ffffffff-0000-0000-0000-000000000006:S-1-5-21-1004336348-1177238915-682003330-4223'
#     }
# }
# Hygiene afterwards: App registrations > API permissions - delete the revoked rows
# ('not granted' leftovers). Removing entries there does NOT revoke access by itself.

# Only IMAP/POP app permissions. Policies do not constrain IMAP/POP, so this policy has
# no effect; mailbox reach is whatever FullAccess grants exist for the service principal.
# Review the Exchange service principal and its mailbox permissions:
Get-ServicePrincipal -Identity 'ffffffff-0000-0000-0000-000000000011' | Format-List DisplayName, AppId, ObjectId
# Per suspect mailbox: Get-MailboxPermission -Identity '<mailbox>' | Where-Object { $_.User -like '*POS Mail Poller*' }
# The policy itself can be removed after review (confirmation prompt follows):
Remove-ApplicationAccessPolicy -Identity '11111111-1111-1111-1111-111111111111\ffffffff-0000-0000-0000-000000000011:S-1-5-21-1004336348-1177238915-682003330-5776;cccccccc-0000-0000-0000-000000000010'

# Exchange permissions are DELEGATED only (Mail.Send, Calendars.Read).
# Policies constrain app-only access, so this policy does nothing today. Removing it
# changes no behavior (confirmation prompt follows):
Remove-ApplicationAccessPolicy -Identity '11111111-1111-1111-1111-111111111111\ffffffff-0000-0000-0000-000000000009:S-1-5-21-1004336348-1177238915-682003330-2530;cccccccc-0000-0000-0000-000000000008'

# DenyAccess policy - do NOT migrate with restrict-style commands (a scope on this
# group would GRANT access to exactly the mailboxes currently denied).
# Review what is denied and who is in the target (links in this row), then either keep
# this policy or redesign scoping so the allow-side groups exclude these recipients.
Get-ApplicationAccessPolicy -Identity '11111111-1111-1111-1111-111111111111\ffffffff-0000-0000-0000-000000000013:S-1-5-21-1004336348-1177238915-682003330-2549;cccccccc-0000-0000-0000-000000000012' | Format-List

# This policy constrains permissions that have no RBAC role: Graph:Calendars.ReadBasic
# KEEP THE POLICY - removing it would widen those permissions to every mailbox.
# Move the app to a supported permission (e.g. Calendars.Read) and rerun this audit.
# No removal command generated on purpose.

# Target group exists in Entra but is not an Exchange recipient. MemberOfGroup scopes
# require an Exchange-recognized group (M365 group, mail-enabled security group, or DL).
# Inspect what Exchange sees for it:
Get-Recipient -Identity 'dept-alerts@contoso.com' -ErrorAction SilentlyContinue | Format-List RecipientTypeDetails, DistinguishedName, ExternalDirectoryObjectId
# Fix: create a mail-enabled security group with the same members and rerun this audit,
# or scope by recipient attributes instead (New-ManagementScope -RecipientRestrictionFilter).

# Multiple groups share this policy's scope name - identify the right one, then rerun:
Invoke-MgGraphRequest -Uri "v1.0/groups?`$filter=displayName eq 'kiosk'&`$select=id,displayName,mail" | Select-Object -ExpandProperty value

# The policy's target no longer resolves. A RestrictAccess policy with an empty/deleted
# target denies the app on ALL mailboxes - the app is either broken or unused.
# 1) Check sign-in logs / app owners to see if the app still needs Exchange access.
# 2) If yes: build an RBAC scope for the correct mailboxes FIRST (see other rows for the pattern).
# 3) Removing this policy WITHOUT an RBAC scope restores TENANT-WIDE access, so it is
#    gated behind an explicit opt-in:
$iUnderstandThisRestoresTenantWideAccess = $false
if ($iUnderstandThisRestoresTenantWideAccess) {
    Remove-ApplicationAccessPolicy -Identity '11111111-1111-1111-1111-111111111111\ffffffff-0000-0000-0000-000000000019:S-1-5-21-1004336348-1177238915-682003330-4151;cccccccc-0000-0000-0000-000000000777'
}

# The policy's target no longer resolves. A RestrictAccess policy with an empty/deleted
# target denies the app on ALL mailboxes - the app is either broken or unused.
# 1) Check sign-in logs / app owners to see if the app still needs Exchange access.
# 2) If yes: build an RBAC scope for the correct mailboxes FIRST (see other rows for the pattern).
# 3) Removing this policy WITHOUT an RBAC scope restores TENANT-WIDE access, so it is
#    gated behind an explicit opt-in:
$iUnderstandThisRestoresTenantWideAccess = $false
if ($iUnderstandThisRestoresTenantWideAccess) {
    Remove-ApplicationAccessPolicy -Identity '11111111-1111-1111-1111-111111111111\ffffffff-0000-0000-0000-000000000022:S-1-5-21-1004336348-1177238915-682003330-7793;eeeeeeee-0000-0000-0000-000000000002'
}

# A Graph lookup failed during this run (see the note in this row) - the state of this
# policy is UNKNOWN. Rerun the audit; if it persists, check Graph consent and throttling.

# App and service principal are gone from Entra ID - the policy is inert.
# Cross-check your app inventory, then remove (confirmation prompt follows):
Remove-ApplicationAccessPolicy -Identity '11111111-1111-1111-1111-111111111111\ffffffff-0000-0000-0000-000000000018:S-1-5-21-1004336348-1177238915-682003330-5842;cccccccc-0000-0000-0000-000000000002'

