#Requires -Modules ExchangeOnlineManagement, Microsoft.Graph.Authentication

<#
.SYNOPSIS
    Inventories Exchange Application Access Policies, generates RBAC migration commands,
    and audits the tenant for unconstrained Exchange application access.
.DESCRIPTION
    Maps Application Access Policies to their Entra apps, target groups/mailboxes, and
    granted permissions. Also scans Graph + Exchange Online service principals for apps
    holding Exchange application permissions with NO mailbox scoping (Microsoft's model
    is insecure-by-default: portal admin consent grants org-wide mailbox access unless a
    policy or RBAC scope is added separately).

    Outputs an HTML report and a companion .ps1 with the migration commands.

    Safety model of the generated commands:
      - Steps 1-4 (service principal pointer, management scope, role assignments, test)
        are additive: they change nothing that exists.
      - Cutover (revoke Entra grants + remove policy) is LIVE code guarded by its own
        verification: it runs Test-ServicePrincipalAuthorization first and aborts unless
        RBAC is confirmed InScope, and removes the (now-inert) policy only after every
        grant revocation succeeded. Remove-ApplicationAccessPolicy is called with
        -Confirm:$false because it does NOT reliably prompt in the modern REST-based EXO
        module (the docs claim Remove-* cmdlets pause, but these don't) - the verification
        is the real gate, not a phantom prompt.
      - Actions that would RESTORE tenant-wide access (removing a policy whose target was
        deleted) are gated behind an explicit opt-in variable.
      - Blocks whose policy target could only be matched by NAME (not object id) are
        emitted fully commented and must be verified by a human first.
.PARAMETER UseDeviceCode
    Sign in to Microsoft Graph and Exchange Online with device-code flow instead of the
    default interactive browser prompt (maps to Connect-MgGraph -UseDeviceCode and
    Connect-ExchangeOnline -Device). Useful for headless/remote sessions - but note that
    some tenants restrict device-code sign-in via Conditional Access.

.EXAMPLE
    .\Audit-ExoAppAccessPolicies.ps1

    Signs in to Microsoft Graph and Exchange Online interactively (browser prompt),
    audits all Application Access Policies and Exchange app permissions, and saves an
    HTML report plus a companion migration .ps1 to the desktop.

.EXAMPLE
    .\Audit-ExoAppAccessPolicies.ps1 -UseDeviceCode

    Same audit, but both sign-ins use device-code flow (for remote/headless sessions).

.INPUTS
    None. This script does not accept pipeline input.

.OUTPUTS
    Saved to an 'AppAccessPolicyMigration' folder on the desktop:
      AppAccessPolicyMigration_<TenantName>_<yyyyMMdd_HHmmss>.html  (report)
      AppAccessPolicyMigration_<TenantName>_<yyyyMMdd_HHmmss>.ps1   (migration commands)

.NOTES
    Author:  Mike Crowley
    https://mikecrowley.us

    Permissions required:
      - Microsoft Graph: Application.Read.All, Directory.Read.All
        (the generated cutover blocks additionally need AppRoleAssignment.ReadWrite.All)
      - Exchange Online: Organization Management (View-Only recipients minimum for the audit)

    Notable caveats surfaced by the report:
      - MemberOfGroup scopes cover DIRECT group members only (policies honored nesting)
      - DenyAccess policies have no RBAC equivalent and must not be migrated blindly
      - IMAP/POP app permissions have no RBAC roles (scoped via Add-MailboxPermission)
      - Delegated-only apps never needed a policy (policies constrain app-only access)
      - EWS is blocked for non-Microsoft apps Oct 1, 2026 and removed after Apr 2027

.LINK
    https://learn.microsoft.com/en-us/exchange/permissions-exo/application-rbac

.LINK
    https://learn.microsoft.com/en-us/entra/identity/enterprise-apps/deactivate-app-registration

.LINK
    https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/new-applicationaccesspolicy

.LINK
    https://learn.microsoft.com/en-us/powershell/exchange/recipientfilter-properties
#>

[CmdletBinding()]
param(
    # Device-code sign-in for headless/remote sessions. Interactive browser auth is the
    # default because a growing number of tenants restrict device-code flow.
    [switch]$UseDeviceCode
)

Disconnect-MgGraph -ErrorAction SilentlyContinue
$graphConnect = @{ NoWelcome = $true; ContextScope = 'Process'; Scopes = @('Application.Read.All', 'Directory.Read.All') }
$exoConnect = @{ ShowBanner = $false }
if ($UseDeviceCode) {
    $graphConnect['UseDeviceCode'] = $true
    $exoConnect['Device'] = $true
}

# ExchangeOnlineManagement 3.7+ signs in through the Windows broker (WAM) by default,
# and MSAL's broker crashes with a NullReferenceException in RuntimeBroker inside
# windowless hosts (VS Code integrated console, ISE). Prefer browser auth in those
# hosts, and if the broker still crashes, retry once with -DisableWAM (module 3.7.2+).
$ExoHasDisableWam = (Get-Command Connect-ExchangeOnline).Parameters.ContainsKey('DisableWAM')
if (-not $UseDeviceCode -and $ExoHasDisableWam -and $Host.Name -ne 'ConsoleHost') {
    $exoConnect['DisableWAM'] = $true
}

Connect-MgGraph @graphConnect
try {
    Connect-ExchangeOnline @exoConnect
}
catch {
    $wamCrash = "$_" -match 'RuntimeBroker|Object reference not set'
    if ($wamCrash -and $ExoHasDisableWam -and -not $exoConnect.ContainsKey('DisableWAM')) {
        Write-Warning 'Windows broker (WAM) sign-in failed in this host - retrying with browser auth (-DisableWAM).'
        $exoConnect['DisableWAM'] = $true
        Connect-ExchangeOnline @exoConnect
    }
    elseif ($wamCrash) {
        throw ("Connect-ExchangeOnline failed initializing the Windows broker (WAM), which is known to crash in " +
               "windowless hosts like the VS Code integrated console. Fixes: run from a regular PowerShell console, " +
               "update ExchangeOnlineManagement to 3.7.2 or later (adds -DisableWAM), or rerun with -UseDeviceCode. " +
               "Original error: $_")
    }
    else { throw }
}

# Fail fast if Graph is not usable. A silent Graph failure later must never be
# misclassified as "app not found" (which would generate policy-deletion guidance).
try {
    $Org = Invoke-MgGraphRequest -Uri "v1.0/organization" -ErrorAction Stop
}
catch {
    throw "Microsoft Graph query failed. Verify consent for Application.Read.All and Directory.Read.All. Error: $_"
}
$TenantName = ($Org.value[0].displayName -replace '[^\w\-]', '')
$TenantId = "$($Org.value[0].id)"

#region Helpers

function Invoke-GraphSafe {
    # Distinguishes "object not found" (404) from every other failure (auth, throttling,
    # network). Only a true 404 may drive orphan/not-found classifications.
    param([string]$Uri)
    try {
        return @{ Ok = $true; Data = (Invoke-MgGraphRequest -Uri $Uri -ErrorAction Stop) }
    }
    catch {
        $msg = $_.Exception.Message
        $status = $null
        try { $status = [int]$_.Exception.StatusCode } catch { }
        if (-not $status) { try { $status = [int]$_.Exception.Response.StatusCode } catch { } }
        $notFound = ($status -eq 404) -or
                    ($msg -match 'Request_ResourceNotFound|ResourceNotFound|\bNotFound\b|\bNot Found\b|Status:\s*404|\b404\b')
        return @{ Ok = $false; NotFound = [bool]$notFound; Error = $msg }
    }
}

function EscSq { param([string]$s) if ($null -eq $s) { '' } else { $s.Replace("'", "''") } }
function EscDq { param([string]$s) if ($null -eq $s) { '' } else { $s.Replace('`', '``').Replace('$', '`$').Replace('"', '`"') } }
function HtmlEnc { param([string]$s) [System.Net.WebUtility]::HtmlEncode("$s") }

function Get-ExoGroupDn {
    # Resolve an Entra group to its Exchange DistinguishedName for a MemberOfGroup filter.
    # Strategies are independent try/catches so one unsupported filter or cmdlet cannot
    # take the others down (Get-Group does not accept every filterable property, and
    # M365 groups are only visible to Get-UnifiedGroup / Get-Recipient).
    param([string]$GroupId, [string]$GroupEmail)

    try {
        $r = Get-Recipient -Filter "ExternalDirectoryObjectId -eq '$GroupId'" -ErrorAction Stop | Select-Object -First 1
        if ($r.DistinguishedName) { return "$($r.DistinguishedName)" }
    }
    catch { }
    if ($GroupEmail) {
        try {
            $r = Get-Recipient -Identity $GroupEmail -ErrorAction Stop | Select-Object -First 1
            # Guard against a stale email pointing at a different object
            if ($r.DistinguishedName -and (-not $r.ExternalDirectoryObjectId -or "$($r.ExternalDirectoryObjectId)" -eq $GroupId)) {
                return "$($r.DistinguishedName)"
            }
        }
        catch { }
        try {
            $g = Get-Group -Identity $GroupEmail -ErrorAction Stop | Select-Object -First 1
            if ($g.DistinguishedName -and (-not $g.ExternalDirectoryObjectId -or "$($g.ExternalDirectoryObjectId)" -eq $GroupId)) {
                return "$($g.DistinguishedName)"
            }
        }
        catch { }
        try {
            $u = Get-UnifiedGroup -Identity $GroupEmail -ErrorAction Stop | Select-Object -First 1
            if ($u.DistinguishedName -and (-not $u.ExternalDirectoryObjectId -or "$($u.ExternalDirectoryObjectId)" -eq $GroupId)) {
                return "$($u.DistinguishedName)"
            }
        }
        catch { }
    }
    return $null
}

function Find-GroupByScopeName {
    # Last-resort fallback, used only when the policy Identity carries no object GUID.
    # Requires an UNAMBIGUOUS match; name matches are always flagged for human review.
    param([string]$ScopeName)
    if (-not $ScopeName) { return @{ Match = $null; Ambiguous = $false } }
    $escaped = $ScopeName.Replace("'", "''")
    foreach ($prop in @('mailNickname', 'displayName', 'mail')) {
        $r = Invoke-GraphSafe "v1.0/groups?`$filter=$prop eq '$escaped'&`$select=id,displayName,mail,mailNickname"
        if ($r.Ok -and $r.Data.value.Count -eq 1) { return @{ Match = $r.Data.value[0]; Ambiguous = $false } }
        if ($r.Ok -and $r.Data.value.Count -gt 1) { return @{ Match = $null; Ambiguous = $true } }
    }
    return @{ Match = $null; Ambiguous = $false }
}

function Find-RecipientByScopeName {
    param([string]$ScopeName)
    if (-not $ScopeName) { return $null }
    try {
        return Get-Recipient -Identity $ScopeName -ErrorAction Stop |
            Select-Object DisplayName, PrimarySmtpAddress, RecipientType, RecipientTypeDetails, ExternalDirectoryObjectId
    }
    catch { }
    return $null
}

#endregion Helpers

#region Permission -> RBAC role maps
# Per https://learn.microsoft.com/en-us/exchange/permissions-exo/application-rbac
# (Supported Application Roles table). RBAC for Applications covers Microsoft Graph
# and EWS only. Note: some Graph permission names carry a .All suffix that the role
# name drops (e.g. MailboxFolder.Read.All -> Application MailboxFolder.Read).

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
    'MailboxItem.Export.All'       = 'Application MailboxItem.Export'
    'MailboxItem.ImportExport.All' = 'Application MailboxItem.ImportExport'
    'MailboxConfigItem.Read'       = 'Application MailboxConfigItem.Read'
    'MailboxConfigItem.ReadWrite'  = 'Application MailboxConfigItem.ReadWrite'
    'MailTips.ReadBasic.All'       = 'Application MailTips.ReadBasic.All'
}
$ExoRoleMap = @{
    'full_access_as_app' = 'Application EWS.AccessAsApp'
    'SMTP.SendAsApp'     = 'Application SMTP.SendAsApp'
}
# EXO app roles with NO RBAC equivalent (RBAC supports Graph + EWS only). IMAP/POP
# app-only access is authorized per-mailbox via Add-MailboxPermission on the Exchange
# service principal - it is not replaced by RBAC and must not be revoked blindly.
$UnmappableExoRoles = @('IMAP.AccessAsApp', 'POP.AccessAsApp')
# Graph Exchange-data app permissions that App Access Policies constrain but that have
# no RBAC role (removing the policy would leave them unscoped).
$UnmappableGraphExchange = @('Calendars.ReadBasic')

$EwsRetirementWarning = '# WARNING: EWS is blocked for non-Microsoft apps starting Oct 1, 2026 (EWSAllowedAppIDs allowlist) and removed after Apr 2027. Plan a Microsoft Graph migration for this app.'

#endregion Permission -> RBAC role maps

# Cache well-known resource service principals for permission-name resolution.
# These MUST load - without them, granted permissions cannot be identified and every
# app would be misreported as having no Exchange permissions.
$ServicePrincipals = @{}
foreach ($id in @($GraphAppId, $ExoAppId)) {
    $r = Invoke-GraphSafe "v1.0/servicePrincipals(appId='$id')"
    if ($r.Ok) { $ServicePrincipals[$id] = $r.Data }
    else { throw "Could not load well-known service principal $id (needed to resolve permission names). Error: $($r.Error)" }
}
$ResourceAppIdCache = @{}   # resource SP objectId -> appId (successes only; failures are never cached)

function Resolve-PermissionName {
    param([string]$ResourceAppId, [string]$PermissionId)
    $sp = $ServicePrincipals[$ResourceAppId]
    if (-not $sp) { return $PermissionId }
    $match = $sp.appRoles | Where-Object { $_.id -eq $PermissionId }
    if ($match) { return $match.value } else { return $PermissionId }
}

# Cache existing Exchange Service Principals. Track failure explicitly - an empty cache
# from a failed call would silently break RBAC detection and the EXO SP column.
$ExchangeServicePrincipals = @{}
$ExoSpCacheOk = $true
try {
    Get-ServicePrincipal -ErrorAction Stop | ForEach-Object {
        $ExchangeServicePrincipals[$_.AppId] = $_
    }
}
catch { $ExoSpCacheOk = $false }

$Policies = Get-ApplicationAccessPolicy

$ScopeRegistry = @{}    # target object GUID -> management scope name (reuse across apps sharing a target)
$UsedScopeNames = @{}   # scope name -> target GUID (uniqueness)

$Report = foreach ($Policy in $Policies) {
    $AppId = $Policy.AppId
    $AccessRight = "$($Policy.AccessRight)"
    $IsDeny = $AccessRight -match 'Deny'
    $Issues = @()
    $MigrationStatus = 'Ready'
    $MigrationBlockers = @()
    $PolicyIdSafe = EscSq "$($Policy.Identity)"

    # --- Application identity: the SERVICE PRINCIPAL is authoritative. Multi-tenant /
    # third-party apps have no application object in this tenant, only a service principal.
    $spRes = Invoke-GraphSafe "v1.0/servicePrincipals(appId='$AppId')?`$select=id,appId,displayName,accountEnabled"
    $appRes = Invoke-GraphSafe "v1.0/applications(appId='$AppId')?`$select=id,displayName"

    $SpLive = $spRes.Ok
    $EntraSpObjectId = if ($SpLive) { $spRes.Data.id } else { $null }
    $SpAccountEnabled = if ($SpLive) { $spRes.Data.accountEnabled } else { $null }

    # Deactivation state (https://learn.microsoft.com/en-us/entra/identity/enterprise-apps/deactivate-app-registration):
    # isDisabled lives on the APPLICATION object (global token block; beta endpoint),
    # accountEnabled on the SERVICE PRINCIPAL (tenant-scoped sign-in block). Supplementary
    # read - failure here must never affect classification.
    $AppDeactivated = $null
    $DisabledByMicrosoft = $null
    if ($appRes.Ok) {
        $disRes = Invoke-GraphSafe "beta/applications(appId='$AppId')?`$select=isDisabled,disabledByMicrosoftStatus"
        if ($disRes.Ok) {
            $AppDeactivated = [bool]$disRes.Data.isDisabled
            if ($disRes.Data.disabledByMicrosoftStatus) { $DisabledByMicrosoft = "$($disRes.Data.disabledByMicrosoftStatus)" }
        }
    }
    $AppDisplayName = if ($SpLive) { "$($spRes.Data.displayName)".Trim() }
                      elseif ($appRes.Ok) { "$($appRes.Data.displayName)".Trim() }
                      else { $AppId }
    $AppObjectId = if ($appRes.Ok) { $appRes.Data.id } else { $null }

    if (-not $spRes.Ok -and -not $spRes.NotFound) {
        $Issues += 'Lookup Error'
        $MigrationStatus = 'Error'
        $MigrationBlockers += "Graph service principal lookup failed - rerun or investigate: $($spRes.Error)"
    }
    elseif (-not $SpLive -and $appRes.NotFound) {
        $Issues += 'Orphaned'
        $MigrationStatus = 'Delete Only'
        $MigrationBlockers += 'App and service principal deleted in Entra ID - policy is inert; remove policy only'
    }
    elseif (-not $SpLive -and -not $appRes.Ok) {
        $Issues += 'Lookup Error'
        $MigrationStatus = 'Error'
        $MigrationBlockers += "Service principal not found and application lookup failed - rerun before acting: $($appRes.Error)"
    }
    elseif (-not $SpLive) {
        $Issues += 'No Service Principal'
        $MigrationStatus = 'Review'
        $MigrationBlockers += 'App registration exists but has no service principal in this tenant - the app cannot get app-only tokens, so the policy is dormant'
    }

    $ExoSpExists = $ExchangeServicePrincipals.ContainsKey($AppId)

    # Actual (live) RBAC application role assignments in Exchange for this app. Detects
    # half-finished migrations: grants already revoked but the legacy policy left behind.
    $LiveRbacRoles = @()
    if ($ExoSpExists) {
        try {
            $LiveRbacRoles = @(Test-ServicePrincipalAuthorization -Identity $AppId -ErrorAction Stop |
                ForEach-Object { "$($_.RoleName)" } | Where-Object { $_ } | Select-Object -Unique)
        }
        catch { }
    }

    # --- Granted APPLICATION permissions (appRoleAssignments on the service principal) ---
    $MigratedPerms = @()      # display strings for permissions replaced by RBAC roles
    $KeepPerms = @()          # display strings for permissions NOT replaced by RBAC - must be kept
    $RbacRoles = @()
    $RevocationCmds = @()     # per-permission grant revocation commands (run inside the verified cutover gate)
    $HasImapPop = $false
    $HasEws = $false
    $HasUnmappableExchange = $false

    if ($SpLive) {
        $uri = "v1.0/servicePrincipals/$EntraSpObjectId/appRoleAssignments"
        while ($uri) {
            $aRes = Invoke-GraphSafe $uri
            if (-not $aRes.Ok) {
                $Issues += 'Permission Lookup Error'
                if ($MigrationStatus -ne 'Error') { $MigrationStatus = 'Error' }
                $MigrationBlockers += "Could not enumerate app permissions: $($aRes.Error)"
                break
            }
            foreach ($assignment in $aRes.Data.value) {
                if (-not $ResourceAppIdCache.ContainsKey($assignment.resourceId)) {
                    $rRes = Invoke-GraphSafe "v1.0/servicePrincipals/$($assignment.resourceId)?`$select=appId"
                    if ($rRes.Ok) { $ResourceAppIdCache[$assignment.resourceId] = $rRes.Data.appId }
                    else {
                        $Issues += 'Permission Lookup Error'
                        if ($MigrationStatus -ne 'Error') { $MigrationStatus = 'Error' }
                        $MigrationBlockers += "Could not resolve a permission's resource ($($assignment.resourceId)): $($rRes.Error)"
                        continue
                    }
                }
                $resourceAppId = $ResourceAppIdCache[$assignment.resourceId]
                $permName = Resolve-PermissionName -ResourceAppId $resourceAppId -PermissionId $assignment.appRoleId
                $resourceName = switch ($resourceAppId) {
                    $GraphAppId { 'Graph' }
                    $ExoAppId { 'EXO' }
                    default { 'Other' }
                }
                $display = "$resourceName`:$permName"

                $role = $null
                if ($resourceAppId -eq $GraphAppId -and $GraphRoleMap.ContainsKey($permName)) { $role = $GraphRoleMap[$permName] }
                elseif ($resourceAppId -eq $ExoAppId -and $ExoRoleMap.ContainsKey($permName)) { $role = $ExoRoleMap[$permName] }

                if ($role) {
                    $MigratedPerms += $display
                    $RbacRoles += $role
                    if ($role -eq 'Application EWS.AccessAsApp') { $HasEws = $true }
                    # A 404 on the DELETE means the grant is already gone - the desired end
                    # state - so reruns of the cutover stay green (idempotent).
                    $RevocationCmds += "    try { Invoke-MgGraphRequest -Method DELETE -Uri 'v1.0/servicePrincipals/$EntraSpObjectId/appRoleAssignments/$($assignment.id)' -ErrorAction Stop; Write-Host 'Revoked $display' } catch { if (`"`$_`" -match 'Request_ResourceNotFound|\b404\b') { Write-Host 'Already revoked: $display' } else { `$revokeFailed = `$true; Write-Warning `"Revocation FAILED ($display): `$_`" } }"
                }
                else {
                    $KeepPerms += $display
                    if ($resourceAppId -eq $ExoAppId -and $UnmappableExoRoles -contains $permName) { $HasImapPop = $true }
                    elseif ($resourceAppId -eq $ExoAppId) {
                        # Unmapped EXO role (e.g. legacy Outlook REST Mail.*): assume the policy
                        # constrains it, so removal guidance must NOT be generated (safe direction).
                        $HasUnmappableExchange = $true
                    }
                    if ($resourceAppId -eq $GraphAppId -and $UnmappableGraphExchange -contains $permName) { $HasUnmappableExchange = $true }
                }
            }
            $uri = $aRes.Data.'@odata.nextLink'
        }
    }
    $RbacRoles = @($RbacRoles | Select-Object -Unique)

    # --- Granted DELEGATED permissions (oauth2PermissionGrants). Policies and RBAC only
    # constrain app-only access; a policy on a delegated-only app has no effect.
    $DelegatedExchangeScopes = @()
    $DelegatedLookupOk = $false
    if ($SpLive) {
        $uri = "v1.0/servicePrincipals/$EntraSpObjectId/oauth2PermissionGrants"
        $grants = @()
        $grantsOk = $true
        while ($uri) {
            $gRes = Invoke-GraphSafe $uri
            if (-not $gRes.Ok) { $grantsOk = $false; break }
            $grants += $gRes.Data.value
            $uri = $gRes.Data.'@odata.nextLink'
        }
        if ($grantsOk) {
            $DelegatedLookupOk = $true
            $exchangeResourceIds = @($ServicePrincipals.Values | ForEach-Object { $_.id })
            foreach ($grant in $grants) {
                if ($exchangeResourceIds -contains $grant.resourceId -and $grant.scope) {
                    foreach ($s in ($grant.scope -split '\s+' | Where-Object { $_ })) {
                        if ($s -match '^(Mail|Calendars|Contacts|MailboxSettings|MailboxFolder|MailboxItem|MailTips|EWS|IMAP|POP|SMTP)([._-]|$)' -or $s -eq 'full_access_as_user') {
                            $DelegatedExchangeScopes += $s
                        }
                    }
                }
            }
            $DelegatedExchangeScopes = @($DelegatedExchangeScopes | Select-Object -Unique)
        }
    }

    # --- Exchange relevance (runs before the deny check so deny rows keep this context) ---
    if ($SpLive -and $RbacRoles.Count -eq 0 -and $MigrationStatus -eq 'Ready') {
        # Unmappable-but-policy-constrained permissions MUST win over every branch that
        # emits policy-removal guidance - removing the policy would unscope them.
        if ($HasUnmappableExchange) {
            $Issues += 'Unmappable Permission'
            $MigrationStatus = 'Review'
            $MigrationBlockers += 'App holds Exchange permissions this policy constrains but that have no RBAC role. Keep the policy, or move the app to a supported permission, before removing anything.'
            if ($HasImapPop) {
                $MigrationBlockers += 'App also holds IMAP/POP permissions (never policy-constrained; scoped via Add-MailboxPermission on the Exchange service principal).'
            }
        }
        elseif ($HasImapPop) {
            $Issues += 'IMAP/POP Only'
            $MigrationStatus = 'Review'
            $MigrationBlockers += 'Only IMAP/POP app permissions: policies do not constrain IMAP/POP (they cover Graph/REST/EWS) and RBAC has no equivalent. Mailbox reach is controlled via Add-MailboxPermission on the Exchange service principal.'
        }
        elseif ($LiveRbacRoles.Count -gt 0) {
            # Grants revoked + RBAC assignments live = an interrupted cutover. The policy
            # constrains nothing now (it only ever constrained Entra grants).
            $Issues += 'Finish Migration'
            $MigrationStatus = 'Review'
            $MigrationBlockers += "Live RBAC role assignments exist ($($LiveRbacRoles -join ', ')) and no matching tenant-wide Entra grants remain. Migration is nearly complete - remove the legacy policy to finish."
        }
        elseif ($DelegatedLookupOk -and $DelegatedExchangeScopes.Count -gt 0) {
            $Issues += 'Delegated Only'
            $MigrationStatus = 'Review'
            $MigrationBlockers += "App's Exchange permissions are DELEGATED only ($($DelegatedExchangeScopes -join ', ')). Policies constrain app-only access, so this policy has no effect - likely created under a misunderstanding. Delegated access is already limited to what each signed-in user can reach."
        }
        else {
            $Issues += 'No Exchange Permissions'
            $MigrationStatus = 'Review'
            $note = 'App has no Exchange (Graph Outlook/EWS) application permissions - this policy has no effect today.'
            if (-not $DelegatedLookupOk) { $note += ' (Delegated permissions could not be checked in this run.)' }
            $MigrationBlockers += $note
        }
    }
    if ($HasImapPop -and $RbacRoles.Count -gt 0) {
        $Issues += 'IMAP/POP'
        $MigrationBlockers += 'App also holds IMAP/POP permissions, which RBAC cannot replace. Keep them; IMAP/POP mailbox access is controlled via Add-MailboxPermission on the Exchange service principal.'
    }
    if ($RbacRoles.Count -gt 0 -and $LiveRbacRoles.Count -gt 0 -and $MigrationStatus -in @('Ready', 'Review')) {
        # Tenant-wide grants remain AND RBAC assignments exist - likely a partial cutover.
        $Issues += 'RBAC Live'
        $MigrationBlockers += "Live RBAC assignments already exist ($($LiveRbacRoles -join ', ')) while tenant-wide grants remain. If a previous cutover was interrupted, rerun this block's cutover to finish revoking and remove the policy."
    }

    # --- Deny check AFTER relevance so those notes are preserved; deny dominates below ---
    if ($IsDeny -and $MigrationStatus -in @('Ready', 'Review')) {
        $Issues += 'Deny Policy'
        $MigrationStatus = 'Review'
        $MigrationBlockers = @('DenyAccess policy: RBAC has no deny equivalent. Do NOT create a scope on this group (that would invert the policy). Keep the policy for now or redesign scoping.') + $MigrationBlockers
    }

    # --- Deactivation / sign-in state. A blocked app gets no new tokens, so the policy has
    # no live effect and this row is a retire-vs-migrate decision, not a plain migration.
    $SignInBlocked = $false
    if ($MigrationStatus -in @('Ready', 'Review')) {
        if ($DisabledByMicrosoft) {
            $SignInBlocked = $true
            $Issues += 'Disabled by Microsoft'
            $MigrationStatus = 'Review'
            $MigrationBlockers += "App is disabled BY MICROSOFT (disabledByMicrosoftStatus = $DisabledByMicrosoft) - investigate for fraud/compromise before migrating anything."
        }
        if ($AppDeactivated -eq $true) {
            $SignInBlocked = $true
            $Issues += 'Deactivated'
            $MigrationStatus = 'Review'
            $MigrationBlockers += 'App registration is DEACTIVATED (isDisabled = true): no new tokens are issued, so this policy has no live effect. Retiring? Remove the policy and revoke remaining grants. Keeping? Reactivate (App registrations > Deactivated applications) before end-to-end testing.'
        }
        elseif ($SpLive -and $SpAccountEnabled -eq $false) {
            $SignInBlocked = $true
            $Issues += 'Sign-in Disabled'
            $MigrationStatus = 'Review'
            $MigrationBlockers += 'Enterprise application sign-in is DISABLED in this tenant (accountEnabled = false): no new tokens, so this policy has no live effect. Retiring? Remove the policy and revoke remaining grants. Keeping? Re-enable sign-in before end-to-end testing.'
        }
    }

    # --- Resolve the policy target. The trailing GUID of the policy Identity is the
    # target's Entra object id - exact, unlike display-name matching.
    $ScopeGuid = $null
    if ("$($Policy.Identity)" -match ';([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})\s*$') {
        $ScopeGuid = $Matches[1]
    }

    $TargetType = 'Not Found'
    $TargetName = $Policy.ScopeName
    $TargetEmail = $null
    $GroupInfo = $null
    $UserInfo = $null
    $GroupDN = $null
    $TargetEdoid = $null
    $TestMailbox = $null
    $TestMailboxChecks = 0
    $HasNestedGroups = $false
    $ResolvedByName = $false

    if ($ScopeGuid) {
        $g = Invoke-GraphSafe "v1.0/groups/$ScopeGuid`?`$select=id,displayName,mail,mailNickname"
        if ($g.Ok) { $GroupInfo = $g.Data }
        elseif ($g.NotFound) {
            $u = Invoke-GraphSafe "v1.0/users/$ScopeGuid`?`$select=id,displayName,mail,userPrincipalName"
            if ($u.Ok) { $UserInfo = $u.Data }
            elseif ($u.NotFound) {
                # Both 404: the target object was deleted. Never fall back to name matching
                # here - any name match would necessarily be a DIFFERENT object.
                if ($MigrationStatus -in @('Ready', 'Review')) {
                    $Issues += 'Target Missing'
                    $MigrationStatus = 'Blocked'
                    $MigrationBlockers += 'Policy target object was DELETED in Entra ID. For a RestrictAccess policy that means the app is currently denied on ALL mailboxes.'
                }
            }
            else {
                $Issues += 'Lookup Error'
                if ($MigrationStatus -ne 'Error') { $MigrationStatus = 'Error' }
                $MigrationBlockers += "Target user lookup failed - rerun before acting: $($u.Error)"
            }
        }
        else {
            $Issues += 'Lookup Error'
            if ($MigrationStatus -ne 'Error') { $MigrationStatus = 'Error' }
            $MigrationBlockers += "Target group lookup failed - rerun before acting: $($g.Error)"
        }
    }
    else {
        $nameResult = Find-GroupByScopeName -ScopeName $Policy.ScopeName
        if ($nameResult.Ambiguous -and $MigrationStatus -in @('Ready', 'Review')) {
            $Issues += 'Ambiguous Name'
            $MigrationStatus = 'Blocked'
            $MigrationBlockers += "Multiple groups match the policy scope name '$($Policy.ScopeName)' - cannot determine the intended target"
        }
        elseif ($nameResult.Match) {
            $GroupInfo = $nameResult.Match
            $ResolvedByName = $true
        }
    }

    if ($GroupInfo) {
        $TargetType = 'Group'
        $TargetName = $GroupInfo.displayName
        $TargetEmail = $GroupInfo.mail
        $GroupDN = Get-ExoGroupDn -GroupId $GroupInfo.id -GroupEmail $TargetEmail
        if (-not $GroupDN -and $MigrationStatus -in @('Ready', 'Review')) {
            $Issues += 'Group Not In EXO'
            $MigrationStatus = 'Blocked'
            $MigrationBlockers += 'Group is not an Exchange recipient - MemberOfGroup scopes only work with Exchange-recognized groups (M365 group, mail-enabled security group, or DL)'
        }
        # Nested groups: MemberOfGroup scopes match DIRECT members only, while App Access
        # Policies honored nested membership (per PolicyScopeGroupId documentation).
        if ($GroupDN) {
            $membersOk = $true
            $memberUri = "v1.0/groups/$($GroupInfo.id)/members?`$select=id&`$top=999"
            while ($memberUri) {
                $m = Invoke-GraphSafe $memberUri
                if (-not $m.Ok) { $membersOk = $false; break }
                if (-not $HasNestedGroups) {
                    $HasNestedGroups = [bool]($m.Data.value | Where-Object { $_.'@odata.type' -eq '#microsoft.graph.group' })
                }
                if (-not $TestMailbox) {
                    foreach ($memberUser in ($m.Data.value | Where-Object { $_.'@odata.type' -eq '#microsoft.graph.user' })) {
                        if ($TestMailboxChecks -ge 5) { break }
                        $TestMailboxChecks++
                        $fu = Invoke-GraphSafe "v1.0/users/$($memberUser.id)?`$select=mail,userPrincipalName"
                        if (-not $fu.Ok) { continue }
                        $candidate = if ($fu.Data.mail) { $fu.Data.mail } else { $fu.Data.userPrincipalName }
                        if (-not $candidate) { continue }
                        # Must be a real Exchange recipient, or the cutover InScope test can never pass
                        try {
                            if (Get-Recipient -Identity $candidate -ErrorAction Stop) { $TestMailbox = "$candidate"; break }
                        }
                        catch { }
                    }
                }
                if ($HasNestedGroups -and $TestMailbox) { break }
                $memberUri = $m.Data.'@odata.nextLink'
            }
            if ($HasNestedGroups -and $MigrationStatus -eq 'Ready') {
                $Issues += 'Nested Groups'
                $MigrationStatus = 'Review'
                $MigrationBlockers += 'Group contains nested groups: MemberOfGroup scopes cover DIRECT members only. Flatten membership (add nested members directly) before cutover.'
            }
            if (-not $membersOk -and $MigrationStatus -eq 'Ready') {
                $Issues += 'Members Unverified'
                $MigrationStatus = 'Review'
                $MigrationBlockers += 'Could not enumerate group members to check for nested groups - verify manually before cutover'
            }
            if ($ResolvedByName -and $MigrationStatus -in @('Ready', 'Review')) {
                $Issues += 'Resolved By Name'
                $MigrationStatus = 'Review'
                $MigrationBlockers += 'Target was matched by NAME only (policy identity carried no object id). Verify this is the intended group; the generated commands are commented out until then.'
            }
        }
    }
    elseif ($UserInfo) {
        # Require an actual Exchange recipient behind the Entra user, otherwise the
        # generated scope filter would match nothing.
        $Recipient = $null
        try {
            $Recipient = Get-Recipient -Filter "ExternalDirectoryObjectId -eq '$($UserInfo.id)'" -ErrorAction Stop | Select-Object -First 1
        }
        catch { }
        if (-not $Recipient -and $UserInfo.userPrincipalName) {
            try { $Recipient = Get-Recipient -Identity $UserInfo.userPrincipalName -ErrorAction Stop | Select-Object -First 1 } catch { }
        }
        if ($Recipient) {
            $TargetType = "$($Recipient.RecipientTypeDetails)"
            $TargetName = $UserInfo.displayName
            $TargetEmail = if ($Recipient.PrimarySmtpAddress) { "$($Recipient.PrimarySmtpAddress)" } else { $UserInfo.mail }
            $TargetEdoid = $UserInfo.id
            $TestMailbox = $TargetEmail
            $Issues += 'Single Mailbox'
        }
        elseif ($MigrationStatus -in @('Ready', 'Review')) {
            $TargetType = 'User'
            $TargetName = $UserInfo.displayName
            $TargetEmail = $UserInfo.mail
            $Issues += 'No EXO Recipient'
            $MigrationStatus = 'Blocked'
            $MigrationBlockers += 'Entra user exists but has no Exchange recipient - a RestrictAccess policy pointing at it denies the app on ALL mailboxes'
        }
    }
    elseif ($MigrationStatus -in @('Ready', 'Review')) {
        $Recipient = Find-RecipientByScopeName -ScopeName $Policy.ScopeName
        if ($Recipient -and $Recipient.ExternalDirectoryObjectId) {
            $TargetType = "$($Recipient.RecipientTypeDetails)"
            $TargetName = $Recipient.DisplayName
            $TargetEmail = "$($Recipient.PrimarySmtpAddress)"
            $TargetEdoid = "$($Recipient.ExternalDirectoryObjectId)"
            $TestMailbox = $TargetEmail
            $Issues += 'Single Mailbox'
            if (-not $ScopeGuid) {
                $ResolvedByName = $true
                $Issues += 'Resolved By Name'
                $MigrationStatus = 'Review'
                $MigrationBlockers += 'Target was matched by NAME only. Verify this is the intended recipient; the generated commands are commented out until then.'
            }
        }
        elseif (-not ($Issues -contains 'Target Missing') -and -not ($Issues -contains 'Ambiguous Name')) {
            $Issues += 'Target Missing'
            $MigrationStatus = 'Blocked'
            $MigrationBlockers += 'Target group/mailbox not found'
        }
    }

    # Entra portal deep links
    $TargetGuidForLink = if ($GroupInfo) { $GroupInfo.id } elseif ($UserInfo) { $UserInfo.id } else { $TargetEdoid }
    $Links = [ordered]@{}
    if ($AppObjectId) { $Links['App registration'] = "https://entra.microsoft.com/#view/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/~/Overview/appId/$AppId" }
    if ($EntraSpObjectId) { $Links['Enterprise app'] = "https://entra.microsoft.com/#view/Microsoft_AAD_IAM/ManagedAppMenuBlade/~/Overview/objectId/$EntraSpObjectId/appId/$AppId" }
    $TargetLink = if ($GroupInfo) { "https://entra.microsoft.com/#view/Microsoft_AAD_IAM/GroupDetailsMenuBlade/~/Overview/groupId/$($GroupInfo.id)" }
                  elseif ($TargetGuidForLink) { "https://entra.microsoft.com/#view/Microsoft_AAD_IAM/UserDetailsMenuBlade/~/Profile/userId/$TargetGuidForLink" }
                  else { $null }

    # --- Build migration commands / call to action. EVERY row gets one. ---
    $MigrationCommands = @()
    $CanGenerate = (-not $IsDeny) -and $SpLive -and $EntraSpObjectId -and $RbacRoles.Count -gt 0 -and
                   ($GroupDN -or $TargetEdoid) -and $MigrationStatus -in @('Ready', 'Review')

    if ($CanGenerate) {
        $targetKey = if ($GroupInfo) { $GroupInfo.id } else { $TargetEdoid }
        if ($ScopeRegistry.ContainsKey($targetKey)) {
            $scopeName = $ScopeRegistry[$targetKey]
        }
        else {
            $targetSafe = if ($GroupInfo -and $GroupInfo.mailNickname) { $GroupInfo.mailNickname }
                          elseif ($GroupInfo) { $GroupInfo.displayName }
                          elseif ($TargetEmail) { ($TargetEmail -split '@')[0] }
                          else { $TargetName }
            $targetSafe = ("$targetSafe".Trim() -replace '[^\w\-]', '_')
            $scopeName = "AppRBAC_$targetSafe"
            $n = 2
            while ($UsedScopeNames.ContainsKey($scopeName) -and $UsedScopeNames[$scopeName] -ne $targetKey) {
                $scopeName = "AppRBAC_${targetSafe}_$n"; $n++
            }
            $UsedScopeNames[$scopeName] = $targetKey
            $ScopeRegistry[$targetKey] = $scopeName
        }

        if ($GroupDN) {
            $filter = "MemberOfGroup -eq '$(EscDq (EscSq $GroupDN))'"
            $filterKey = $GroupDN
        }
        else {
            $filter = "ExternalDirectoryObjectId -eq '$targetKey'"
            $filterKey = $targetKey
        }

        $MigrationCommands += "# ==== $AppDisplayName ($AppId) -> $TargetName [$AccessRight] ===="
        $MigrationCommands += "# Steps 1-4 are additive and do not change existing access."
        if ($SignInBlocked) {
            $MigrationCommands += "# NOTE: this app cannot currently obtain tokens (deactivated / sign-in disabled)."
            $MigrationCommands += "# Steps 1-3 can be staged and the Step 4 test cmdlet still evaluates, but the app"
            $MigrationCommands += "# itself cannot be end-to-end tested until re-enabled. If it is being RETIRED, skip"
            $MigrationCommands += "# migration: revoke its grants and remove the policy instead."
        }
        $MigrationCommands += ""
        $MigrationCommands += "# Step 1: Exchange service principal pointer (idempotent)"
        if ($ExoSpExists) {
            $MigrationCommands += "# Already exists in Exchange Online - nothing to do."
        }
        else {
            $MigrationCommands += "if (-not (Get-ServicePrincipal -Identity '$AppId' -ErrorAction SilentlyContinue)) {"
            $MigrationCommands += "    New-ServicePrincipal -AppId '$AppId' -ObjectId '$EntraSpObjectId' -DisplayName '$(EscSq $AppDisplayName)'"
            $MigrationCommands += "}"
        }
        $MigrationCommands += ""
        $MigrationCommands += "# Step 2: Management scope (idempotent; warns if the name is taken by a different filter)"
        if ($GroupDN) {
            $MigrationCommands += "# NOTE: MemberOfGroup covers DIRECT members only - nested group members are out of scope."
        }
        else {
            $MigrationCommands += "# Single-mailbox scope. To cover more mailboxes later, create a mail-enabled security"
            $MigrationCommands += "# group instead and use: `"MemberOfGroup -eq '<group DN>'`""
        }
        $MigrationCommands += "`$scope = `$null"
        $MigrationCommands += "`$scopeConflict = `$false"
        $MigrationCommands += "`$scope = Get-ManagementScope -Identity '$scopeName' -ErrorAction SilentlyContinue"
        $MigrationCommands += "if (-not `$scope) {"
        $MigrationCommands += "    New-ManagementScope -Name '$scopeName' -RecipientRestrictionFilter `"$filter`""
        $MigrationCommands += "} elseif (`$scope.RecipientFilter -notlike '*$(EscSq $filterKey)*') {"
        $MigrationCommands += "    `$scopeConflict = `$true"
        $MigrationCommands += "    Write-Warning `"Scope '$scopeName' already exists with a DIFFERENT filter: `$(`$scope.RecipientFilter) - role assignments skipped; resolve the conflict first.`""
        $MigrationCommands += "}"
        $MigrationCommands += ""
        $MigrationCommands += "# Step 3: RBAC role assignments (idempotent - skips roles already assigned; skipped"
        $MigrationCommands += "# entirely on a scope-name conflict)"
        $MigrationCommands += "if (-not `$scopeConflict) {"
        $MigrationCommands += "    `$liveRoles = @(Test-ServicePrincipalAuthorization -Identity '$AppId' -ErrorAction SilentlyContinue | ForEach-Object { `$_.RoleName })"
        foreach ($role in $RbacRoles) {
            $MigrationCommands += "    if (`$liveRoles -notcontains '$role') { New-ManagementRoleAssignment -App '$AppId' -Role '$role' -CustomResourceScope '$scopeName' }"
        }
        $MigrationCommands += "}"
        if ($HasEws) { $MigrationCommands += $EwsRetirementWarning }
        $MigrationCommands += ""
        $MigrationCommands += "# Step 4: VERIFY - expect InScope = True for an in-scope mailbox"
        if ($TestMailbox) {
            $MigrationCommands += "Test-ServicePrincipalAuthorization -Identity '$AppId' -Resource '$(EscSq $TestMailbox)' | Format-Table"
        }
        else {
            $MigrationCommands += "# No member mailbox found automatically - substitute one:"
            $MigrationCommands += "# Test-ServicePrincipalAuthorization -Identity '$AppId' -Resource '<member mailbox>' | Format-Table"
        }
        $MigrationCommands += "# Also confirm the application itself still works before continuing."
        $MigrationCommands += ""
        $NeedsFlattenGate = ($Issues -contains 'Nested Groups') -or ($Issues -contains 'Members Unverified')
        $rolesList = ($RbacRoles | ForEach-Object { "'$(EscSq $_)'" }) -join ', '
        $MigrationCommands += "# ---- Step 5-6: CUTOVER. Verifies EVERY migrated role is live and InScope, revokes the"
        $MigrationCommands += "# tenant-wide grants (each one checked), then removes the now-inert legacy policy."
        $MigrationCommands += "# Entra + RBAC grants are a union - scoping only takes effect after the tenant-wide grant"
        $MigrationCommands += "# is revoked. Nothing here runs unverified: the policy is removed ONLY after RBAC is"
        $MigrationCommands += "# confirmed InScope AND every revocation succeeded. Remove is explicit (-Confirm:`$false)"
        $MigrationCommands += "# because Remove-ApplicationAccessPolicy does NOT reliably prompt in the modern EXO module."
        $MigrationCommands += "`$expectedRoles = @($rolesList)"
        if ($TestMailbox) {
            $MigrationCommands += "`$cutoverTestMailbox = '$(EscSq $TestMailbox)'   # in-scope member found by the audit; substitute if needed"
        }
        else {
            $MigrationCommands += "`$cutoverTestMailbox = '<member mailbox>'   # REQUIRED: set to a mailbox inside the new scope"
        }
        $MigrationCommands += "`$auth = if (`$cutoverTestMailbox -notlike '<*') { @(Test-ServicePrincipalAuthorization -Identity '$AppId' -Resource `$cutoverTestMailbox -ErrorAction SilentlyContinue) } else { @() }"
        $MigrationCommands += "`$missingRoles = @(`$expectedRoles | Where-Object { `$role = `$_; -not (`$auth | Where-Object { `$_.RoleName -eq `$role -and `$_.InScope }) })"
        if ($NeedsFlattenGate) {
            $MigrationCommands += "`$nestedGroupsHandled = `$false   # set to `$true after adding nested-group members DIRECTLY to the group"
        }
        $MigrationCommands += "if (`$missingRoles.Count -gt 0) {"
        $MigrationCommands += "    Write-Warning `"$(EscDq $AppDisplayName): roles not verified InScope for `$cutoverTestMailbox (missing: `$(`$missingRoles -join ', ')) - cutover skipped. Run steps 1-4; if they succeeded, try a different in-scope mailbox.`""
        $MigrationCommands += "} elseif (@((Get-MgContext).Scopes) -notcontains 'AppRoleAssignment.ReadWrite.All') {"
        $MigrationCommands += "    Write-Warning 'Graph session lacks AppRoleAssignment.ReadWrite.All - run: Connect-MgGraph -Scopes AppRoleAssignment.ReadWrite.All'"
        if ($NeedsFlattenGate) {
            $MigrationCommands += "} elseif (-not `$nestedGroupsHandled) {"
            $MigrationCommands += "    Write-Warning `"$(EscDq $AppDisplayName): the scope covers DIRECT members only - flatten nested groups, then set ```$nestedGroupsHandled = ```$true and rerun.`""
        }
        $MigrationCommands += "} else {"
        $MigrationCommands += "    `$revokeFailed = `$false"
        if ($KeepPerms.Count -gt 0) {
            $MigrationCommands += "    # KEEP (NOT replaced by RBAC - do not revoke): $($KeepPerms -join ', ')"
        }
        $MigrationCommands += $RevocationCmds
        $MigrationCommands += "    if (`$revokeFailed) {"
        $MigrationCommands += "        Write-Warning 'One or more revocations FAILED - the legacy policy was NOT removed. Fix the errors above and rerun this cutover block.'"
        $MigrationCommands += "    } elseif (@(Get-ApplicationAccessPolicy | Where-Object { `$_.Identity -eq '$PolicyIdSafe' }).Count -eq 0) {"
        $MigrationCommands += "        Write-Host '$(EscSq $AppDisplayName): legacy policy already removed - migration complete.'"
        $MigrationCommands += "    } else {"
        $MigrationCommands += "        Write-Host 'Tenant-wide grants revoked. Exchange caches app permissions 30 min - 2 h; re-test the app.'"
        $MigrationCommands += "        # The policy is now inert (it only constrained the Entra grants, which are gone)."
        $MigrationCommands += "        Remove-ApplicationAccessPolicy -Identity '$PolicyIdSafe' -Confirm:`$false"
        $MigrationCommands += "        if (@(Get-ApplicationAccessPolicy | Where-Object { `$_.Identity -eq '$PolicyIdSafe' }).Count -eq 0) {"
        $MigrationCommands += "            Write-Host '$(EscSq $AppDisplayName): legacy policy removed - migration complete.'"
        $MigrationCommands += "        } else {"
        $MigrationCommands += "            Write-Warning '$(EscSq $AppDisplayName): legacy policy still present (removal failed) - rerun this cutover block to finish.'"
        $MigrationCommands += "        }"
        $MigrationCommands += "    }"
        $MigrationCommands += "}"
        $MigrationCommands += "# Hygiene afterwards: App registrations > API permissions - delete the revoked rows"
        $MigrationCommands += "# ('not granted' leftovers). Removing entries there does NOT revoke access by itself."
        if (-not $AppObjectId) {
            $MigrationCommands += "# (App is registered in another tenant - only the grant revocation applies here.)"
        }

        if ($ResolvedByName) {
            $MigrationCommands = @(
                "# !! TARGET MATCHED BY NAME ONLY - VERIFY BEFORE RUNNING !!"
                "# The policy identity carried no object id, so '$TargetName' was found by name match."
                "# Confirm it is the intended target, then uncomment this block."
            ) + ($MigrationCommands | ForEach-Object { if ($_ -and $_ -notmatch '^\s*#') { "# $_" } else { $_ } })
        }
    }
    elseif ($MigrationStatus -eq 'Delete Only') {
        $MigrationCommands += "# App and service principal are gone from Entra ID - the policy is inert."
        $MigrationCommands += "# Cross-check your app inventory, then remove it:"
        $MigrationCommands += "Remove-ApplicationAccessPolicy -Identity '$PolicyIdSafe' -Confirm:`$false"
    }
    elseif ($IsDeny) {
        $MigrationCommands += "# DenyAccess policy - do NOT migrate with restrict-style commands (a scope on this"
        $MigrationCommands += "# group would GRANT access to exactly the mailboxes currently denied)."
        $MigrationCommands += "# Review what is denied and who is in the target (links in this row), then either keep"
        $MigrationCommands += "# this policy or redesign scoping so the allow-side groups exclude these recipients."
        $MigrationCommands += "Get-ApplicationAccessPolicy -Identity '$PolicyIdSafe' | Format-List"
    }
    elseif ($Issues -contains 'Finish Migration') {
        $MigrationCommands += "# FINISH MIGRATION: RBAC is live ($($LiveRbacRoles -join ', ')) and the matching tenant-wide"
        $MigrationCommands += "# Entra grants are gone, so this legacy policy constrains nothing. Verify, then remove it:"
        if ($TestMailbox) {
            $MigrationCommands += "Test-ServicePrincipalAuthorization -Identity '$AppId' -Resource '$(EscSq $TestMailbox)' | Format-Table   # expect InScope = True"
        }
        else {
            $MigrationCommands += "Test-ServicePrincipalAuthorization -Identity '$AppId' | Format-Table   # spot-check InScope with an in-scope mailbox via -Resource"
        }
        $MigrationCommands += "# Confirm the application still works, then remove the now-inert policy:"
        $MigrationCommands += "Remove-ApplicationAccessPolicy -Identity '$PolicyIdSafe' -Confirm:`$false"
        $MigrationCommands += "if (@(Get-ApplicationAccessPolicy | Where-Object { `$_.Identity -eq '$PolicyIdSafe' }).Count -eq 0) {"
        $MigrationCommands += "    Write-Host '$(EscSq $AppDisplayName): legacy policy removed - migration complete.'"
        $MigrationCommands += "} else {"
        $MigrationCommands += "    Write-Warning '$(EscSq $AppDisplayName): legacy policy still present (removal failed).'"
        $MigrationCommands += "}"
    }
    elseif ($Issues -contains 'Delegated Only') {
        $MigrationCommands += "# Exchange permissions are DELEGATED only ($($DelegatedExchangeScopes -join ', '))."
        $MigrationCommands += "# Policies constrain app-only access, so this policy does nothing today. Removing it"
        $MigrationCommands += "# changes no behavior:"
        $MigrationCommands += "Remove-ApplicationAccessPolicy -Identity '$PolicyIdSafe' -Confirm:`$false"
    }
    elseif ($Issues -contains 'No Exchange Permissions') {
        $MigrationCommands += "# No Exchange application permissions -> this policy has no effect today."
        if ($KeepPerms.Count -gt 0) { $MigrationCommands += "# KEEP (not Exchange-related): $($KeepPerms -join ', ')" }
        $MigrationCommands += "# Removing the policy changes no behavior. If the app is granted Exchange permissions"
        $MigrationCommands += "# later, scope it with RBAC at that time."
        $MigrationCommands += "Remove-ApplicationAccessPolicy -Identity '$PolicyIdSafe' -Confirm:`$false"
    }
    elseif ($Issues -contains 'Unmappable Permission') {
        $MigrationCommands += "# This policy constrains permissions that have no RBAC role: $($KeepPerms -join ', ')"
        $MigrationCommands += "# KEEP THE POLICY - removing it would widen those permissions to every mailbox."
        $MigrationCommands += "# Move the app to a supported permission (e.g. Calendars.Read) and rerun this audit."
        $MigrationCommands += "# No removal command generated on purpose."
    }
    elseif ($Issues -contains 'IMAP/POP Only') {
        $MigrationCommands += "# Only IMAP/POP app permissions. Policies do not constrain IMAP/POP, so this policy has"
        $MigrationCommands += "# no effect; mailbox reach is whatever FullAccess grants exist for the service principal."
        $MigrationCommands += "# Review the Exchange service principal and its mailbox permissions:"
        $MigrationCommands += "Get-ServicePrincipal -Identity '$AppId' | Format-List DisplayName, AppId, ObjectId"
        $MigrationCommands += "# Per suspect mailbox: Get-MailboxPermission -Identity '<mailbox>' | Where-Object { `$_.User -like '*$(EscSq $AppDisplayName)*' }"
        $MigrationCommands += "# The policy itself can be removed after review:"
        $MigrationCommands += "Remove-ApplicationAccessPolicy -Identity '$PolicyIdSafe' -Confirm:`$false"
    }
    elseif ($Issues -contains 'No Service Principal') {
        $MigrationCommands += "# App registration exists but there is no enterprise application (service principal),"
        $MigrationCommands += "# so the app cannot get app-only tokens and the policy is dormant."
        $MigrationCommands += "# If the app should work: New-MgServicePrincipal -AppId '$AppId'   # needs Application.ReadWrite.All"
        $MigrationCommands += "# then rerun this audit. If the app is retired, remove the policy and the registration."
    }
    elseif ($Issues -contains 'Group Not In EXO') {
        $MigrationCommands += "# Target group exists in Entra but is not an Exchange recipient. MemberOfGroup scopes"
        $MigrationCommands += "# require an Exchange-recognized group (M365 group, mail-enabled security group, or DL)."
        $MigrationCommands += "# Inspect what Exchange sees for it:"
        $idOrMail = if ($TargetEmail) { EscSq $TargetEmail } elseif ($GroupInfo) { $GroupInfo.id } else { EscSq "$TargetName" }
        $MigrationCommands += "Get-Recipient -Identity '$idOrMail' -ErrorAction SilentlyContinue | Format-List RecipientTypeDetails, DistinguishedName, ExternalDirectoryObjectId"
        $MigrationCommands += "# Fix: create a mail-enabled security group with the same members and rerun this audit,"
        $MigrationCommands += "# or scope by recipient attributes instead (New-ManagementScope -RecipientRestrictionFilter)."
    }
    elseif ($Issues -contains 'Target Missing' -or $Issues -contains 'No EXO Recipient') {
        $MigrationCommands += "# The policy's target no longer resolves. A RestrictAccess policy with an empty/deleted"
        $MigrationCommands += "# target denies the app on ALL mailboxes - the app is either broken or unused."
        $MigrationCommands += "# 1) Check sign-in logs / app owners to see if the app still needs Exchange access."
        $MigrationCommands += "# 2) If yes: build an RBAC scope for the correct mailboxes FIRST (see other rows for the pattern)."
        $MigrationCommands += "# 3) Removing this policy WITHOUT an RBAC scope restores TENANT-WIDE access, so it is"
        $MigrationCommands += "#    gated behind an explicit opt-in:"
        $MigrationCommands += "`$iUnderstandThisRestoresTenantWideAccess = `$false"
        $MigrationCommands += "if (`$iUnderstandThisRestoresTenantWideAccess) {"
        $MigrationCommands += "    Remove-ApplicationAccessPolicy -Identity '$PolicyIdSafe' -Confirm:`$false"
        $MigrationCommands += "}"
    }
    elseif ($Issues -contains 'Ambiguous Name') {
        $MigrationCommands += "# Multiple groups share this policy's scope name - identify the right one, then rerun:"
        $MigrationCommands += "Invoke-MgGraphRequest -Uri `"v1.0/groups?```$filter=displayName eq '$(EscSq "$($Policy.ScopeName)")'&```$select=id,displayName,mail`" | Select-Object -ExpandProperty value"
    }
    elseif ($MigrationStatus -eq 'Error') {
        $MigrationCommands += "# A Graph lookup failed during this run (see the note in this row) - the state of this"
        $MigrationCommands += "# policy is UNKNOWN. Rerun the audit; if it persists, check Graph consent and throttling."
    }

    # Safety net: a row must never present as 'Ready' without runnable commands.
    if ($MigrationStatus -eq 'Ready' -and -not ($MigrationCommands | Where-Object { $_ })) {
        $MigrationStatus = 'Review'
        $MigrationBlockers += 'No migration commands could be generated - investigate this row manually'
    }

    [PSCustomObject]@{
        App_DisplayName    = $AppDisplayName
        App_Id             = $AppObjectId
        App_ClientId       = $AppId
        Entra_SpObjectId   = $EntraSpObjectId
        Exo_SpExists       = $ExoSpExists
        Policy_Identity    = "$($Policy.Identity)"
        Policy_AccessRight = $AccessRight
        Target_Type        = "$TargetType"
        Target_Name        = $TargetName
        Target_Email       = $TargetEmail
        Target_DN          = $GroupDN
        Target_Link        = $TargetLink
        App_Links          = $Links
        MigratedPerms      = $MigratedPerms -join '; '
        KeepPerms          = $KeepPerms -join '; '
        DelegatedPerms     = $DelegatedExchangeScopes -join '; '
        RbacRoles          = $RbacRoles -join '; '
        LiveRbacRoles      = $LiveRbacRoles -join '; '
        Issues             = $Issues -join ', '
        MigrationStatus    = $MigrationStatus
        MigrationBlockers  = $MigrationBlockers -join ' | '
        MigrationCommands  = $MigrationCommands -join "`n"
    }
}

#region Security scan: Exchange app permissions with NO mailbox scoping
# Microsoft's model is insecure-by-default: admin consent in the portal grants org-wide
# mailbox access; scoping requires a separate policy or RBAC assignment that the GUI
# never mentions. Enumerate every service principal holding Exchange application
# permissions and check whether anything constrains it.

$SecurityScanOk = $true
$SecurityRows = @()
$MsFirstPartySkipped = 0
$MsTenantIds = @('f8cdef31-a31e-4b4a-93e4-5f571e91255a', '72f988bf-86f1-41af-91ab-2d7cd011db47')

# Apps with a RestrictAccess policy (the legacy constraint)
$RestrictPolicyAppIds = @($Policies | Where-Object { "$($_.AccessRight)" -match 'Restrict' } | ForEach-Object { $_.AppId } | Select-Object -Unique)

# Apps with an EXO RBAC application role assignment. Match on unambiguous keys only
# (ObjectId/Identity/AppId) - DisplayName collisions would mislabel apps as RBAC-scoped.
$RbacAppIds = @{}
if (-not $ExoSpCacheOk) { $SecurityScanOk = $false }
try {
    $spAssignments = Get-ManagementRoleAssignment -ErrorAction Stop | Where-Object { "$($_.RoleAssigneeType)" -eq 'ServicePrincipal' }
    $assigneeNames = @{}
    foreach ($a in $spAssignments) { $assigneeNames["$($a.RoleAssignee)"] = $true }
    foreach ($sp in $ExchangeServicePrincipals.Values) {
        foreach ($k in @($sp.Identity, $sp.ObjectId, $sp.AppId)) {
            if ($k -and $assigneeNames.ContainsKey("$k")) { $RbacAppIds[$sp.AppId] = $true }
        }
    }
}
catch { $SecurityScanOk = $false }

function Get-PermRiskTier {
    param([string]$Perm)
    switch -Regex ($Perm) {
        '^(Mail\.Read|Mail\.ReadWrite|Mail\.Send|full_access_as_app|IMAP\.AccessAsApp|POP\.AccessAsApp|SMTP\.SendAsApp|MailboxItem\.|MailboxFolder\.ReadWrite|Exchange\.ManageAsApp|EWS\.AccessAsApp)' { 'high'; break }
        '^(Calendars\.|Contacts\.|MailboxSettings\.|MailboxFolder\.Read|Mail\.ReadBasic)' { 'medium'; break }
        default { 'low' }
    }
}

$GrantHolders = @{}   # principal objectId -> @{ Perms = [list]; }
foreach ($resAppId in @($GraphAppId, $ExoAppId)) {
    $resSp = $ServicePrincipals[$resAppId]
    $roleNameById = @{}
    foreach ($ar in $resSp.appRoles) { $roleNameById[$ar.id] = $ar.value }
    # Mailbox-data roles only. Admin/API EXO roles (Exchange.ManageAsApp, Mailbox.Migration,
    # ReportingWebService...) are out of scope for mailbox scoping and would get inapplicable
    # remediation advice; legacy Outlook REST roles (Mail.*/Calendars.*/...) ARE included.
    $relevant = if ($resAppId -eq $GraphAppId) { @($GraphRoleMap.Keys) + $UnmappableGraphExchange }
                else { @('full_access_as_app', 'IMAP.AccessAsApp', 'POP.AccessAsApp', 'SMTP.SendAsApp') }
    $prefix = if ($resAppId -eq $GraphAppId) { 'Graph' } else { 'EXO' }
    $uri = "v1.0/servicePrincipals/$($resSp.id)/appRoleAssignedTo?`$top=999"
    while ($uri) {
        $res = Invoke-GraphSafe $uri
        if (-not $res.Ok) { $SecurityScanOk = $false; break }
        foreach ($a in $res.Data.value) {
            if ("$($a.principalType)" -ne 'ServicePrincipal') { continue }
            $permName = $roleNameById[$a.appRoleId]
            if (-not $permName) { continue }
            $isRelevant = $relevant -contains $permName
            if (-not $isRelevant -and $resAppId -eq $ExoAppId -and $permName -match '^(Mail|Calendars|Contacts|MailboxSettings)\.') { $isRelevant = $true }
            if (-not $isRelevant) { continue }
            if (-not $GrantHolders.ContainsKey($a.principalId)) {
                $GrantHolders[$a.principalId] = @{ DisplayName = "$($a.principalDisplayName)"; Perms = @() }
            }
            $GrantHolders[$a.principalId].Perms += "$prefix`:$permName"
        }
        $uri = $res.Data.'@odata.nextLink'
    }
}

$DanglingGrants = 0
foreach ($principalId in @($GrantHolders.Keys)) {
    $info = Invoke-GraphSafe "v1.0/servicePrincipals/$principalId`?`$select=appId,displayName,appOwnerOrganizationId,accountEnabled"
    if ($info.NotFound) {
        # Grant pointing at a deleted principal: definitive answer, not a scan failure.
        $DanglingGrants++
        continue
    }
    if (-not $info.Ok) {
        # Keep the app visible rather than silently dropping it from the security table.
        $SecurityScanOk = $false
        $SecurityRows += [PSCustomObject]@{
            AppId       = ''
            SpObjectId  = $principalId
            DisplayName = $GrantHolders[$principalId].DisplayName
            Enabled     = $null
            Deactivated = $false
            Perms       = @($GrantHolders[$principalId].Perms | Select-Object -Unique)
            Constraint  = 'Lookup failed'
            RiskTier    = 'high'
            Notes       = "Graph lookup for this service principal failed - constraint status UNKNOWN; rerun the audit. $($info.Error)"
        }
        continue
    }
    if ($MsTenantIds -contains "$($info.Data.appOwnerOrganizationId)") { $MsFirstPartySkipped++; continue }

    $appId = $info.Data.appId

    # App-level deactivation (isDisabled) is only readable for apps registered in THIS tenant.
    $rowDeactivated = $false
    if ("$($info.Data.appOwnerOrganizationId)" -eq $TenantId) {
        $disRes = Invoke-GraphSafe "beta/applications(appId='$appId')?`$select=isDisabled"
        if ($disRes.Ok -and $disRes.Data.isDisabled) { $rowDeactivated = $true }
    }
    $perms = @($GrantHolders[$principalId].Perms | Select-Object -Unique)
    $graphEwsPerms = @($perms | Where-Object { $_ -match '^Graph:' -or $_ -eq 'EXO:full_access_as_app' })
    $protoPerms = @($perms | Where-Object { $_ -match '^EXO:(IMAP|POP|SMTP)' })
    $hasRestrict = $RestrictPolicyAppIds -contains $appId
    $hasRbac = $RbacAppIds.ContainsKey($appId)

    $notes = @()
    if ($graphEwsPerms.Count -eq 0 -and $protoPerms.Count -gt 0) {
        $constraint = 'Mailbox-permission model'
        $notes += 'IMAP/POP/SMTP reach = FullAccess grants to the service principal (verify with Get-MailboxPermission); policies do not apply'
    }
    elseif ($hasRestrict) {
        $constraint = 'Legacy policy'
        $notes += 'Constrained by a RestrictAccess Application Access Policy (see table above) - migrate to RBAC'
        if ($protoPerms.Count -gt 0) { $notes += "IMAP/POP/SMTP permissions are NOT covered by the policy: $($protoPerms -join ', ')" }
    }
    elseif ($hasRbac) {
        $constraint = 'Unconstrained (RBAC exists)'
        $notes += 'Appears to have an Exchange RBAC assignment, BUT the tenant-wide Entra grant is still in place - grants are a union, so the RBAC scope is not limiting anything. Verify the assignment (Get-ManagementRoleAssignment / Test-ServicePrincipalAuthorization) covers these permissions, then revoke the Entra grant.'
    }
    else {
        $constraint = 'Unconstrained'
        $notes += 'Tenant-wide mailbox access with no Application Access Policy and no RBAC scope'
    }

    $tier = 'low'
    foreach ($p in $perms) {
        $t = Get-PermRiskTier ($p -replace '^(Graph|EXO):', '')
        if ($t -eq 'high') { $tier = 'high'; break }
        if ($t -eq 'medium') { $tier = 'medium' }
    }

    if ($rowDeactivated -or $info.Data.accountEnabled -eq $false) {
        $notes += 'App cannot obtain new tokens right now (deactivated / sign-in disabled) - lower urgency, but the standing grants become live again the moment it is re-enabled.'
    }

    $SecurityRows += [PSCustomObject]@{
        AppId       = $appId
        SpObjectId  = $principalId
        DisplayName = if ($info.Data.displayName) { "$($info.Data.displayName)".Trim() } else { $GrantHolders[$principalId].DisplayName }
        Enabled     = $info.Data.accountEnabled
        Deactivated = $rowDeactivated
        Perms       = $perms
        Constraint  = $constraint
        RiskTier    = $tier
        Notes       = $notes -join ' | '
    }
}

$UnconstrainedCount = @($SecurityRows | Where-Object { $_.Constraint -match '^Unconstrained' }).Count

#endregion Security scan

#region Output

# Summary stats
$TotalPolicies = @($Report).Count
$ReadyToMigrate = @($Report | Where-Object { $_.MigrationStatus -eq 'Ready' }).Count
$NeedsReview = @($Report | Where-Object { $_.MigrationStatus -eq 'Review' }).Count
$Blocked = @($Report | Where-Object { $_.MigrationStatus -in @('Blocked', 'Error') }).Count
$DenyPolicies = @($Report | Where-Object { $_.Policy_AccessRight -match 'Deny' }).Count
$OrphanedApps = @($Report | Where-Object { $_.Issues -match 'Orphaned' }).Count

$SortedReport = $Report | Sort-Object @{Expression = { switch ($_.MigrationStatus) {
    'Ready' { 0 } 'Review' { 1 } 'Blocked' { 2 } 'Error' { 3 } 'Delete Only' { 4 } default { 5 } } } }, App_DisplayName

function New-LinkHtml {
    param($Item)
    $parts = @()
    foreach ($k in $Item.App_Links.Keys) {
        $parts += "<a class='plink' href='$($Item.App_Links[$k])' target='_blank' rel='noopener'>$(HtmlEnc $k) &#8599;</a>"
    }
    $parts -join ' '
}

# Build policy table rows
$HtmlRows = foreach ($item in $SortedReport) {
    $rowClass = switch ($item.MigrationStatus) {
        'Blocked' { 'status-blocked' }
        'Error' { 'status-error' }
        'Review' { 'status-review' }
        'Delete Only' { 'status-delete' }
        default { '' }
    }

    $statusBadge = switch ($item.MigrationStatus) {
        'Ready' { '<span class="status-badge ready">&#10003; Ready</span>' }
        'Blocked' { '<span class="status-badge blocked">&#10005; Blocked</span>' }
        'Error' { '<span class="status-badge error">&#8252; Error</span>' }
        'Review' { '<span class="status-badge review">&#9888; Review</span>' }
        'Delete Only' { '<span class="status-badge delete">&#8709; Delete Only</span>' }
        default { '' }
    }

    $arBadge = if ($item.Policy_AccessRight -match 'Deny') {
        '<span class="ar-badge deny">&#10005; DenyAccess</span>'
    } else {
        '<span class="ar-badge restrict">RestrictAccess</span>'
    }

    $issuesBadges = ''
    foreach ($iss in ($item.Issues -split ', ' | Where-Object { $_ })) {
        $cls = switch -Regex ($iss) {
            'Orphaned|Target Missing|Deny|Error|Ambiguous|No EXO Recipient|Disabled by Microsoft' { 'crit' }
            'Delegated' { 'purp' }
            'Nested|Resolved By Name|Members Unverified' { 'warn' }
            'Single Mailbox|Finish Migration|RBAC Live' { 'info' }
            'Deactivated|Sign-in Disabled' { 'neut' }
            default { 'warn' }
        }
        $issuesBadges += "<span class='badge $cls'>$(HtmlEnc $iss)</span>"
    }

    $exoSpStatus = if (-not $ExoSpCacheOk) { '<span class="sp-missing">Unknown (lookup failed)</span>' }
                   elseif ($item.Exo_SpExists) { '<span class="sp-exists">&#10003; Exists</span>' }
                   else { '<span class="sp-missing">Create needed</span>' }

    $permsList = ''
    if ($item.MigratedPerms) {
        $permsList += ($item.MigratedPerms -split '; ' | ForEach-Object { "<span class='perm migrate' title='Replaced by an RBAC role'>$(HtmlEnc $_)</span>" }) -join ''
    }
    if ($item.KeepPerms) {
        $permsList += ($item.KeepPerms -split '; ' | ForEach-Object { "<span class='perm keep' title='Not replaced by RBAC - keep in Entra ID'>$(HtmlEnc $_) &#183; keep</span>" }) -join ''
    }
    if ($item.DelegatedPerms) {
        $permsList += ($item.DelegatedPerms -split '; ' | ForEach-Object { "<span class='perm delegated' title='Delegated permission - not affected by policies or RBAC for Applications'>$(HtmlEnc $_) &#183; delegated</span>" }) -join ''
    }
    if (-not $permsList) { $permsList = "<span class='none'>None granted</span>" }

    $rbacList = ''
    if ($item.LiveRbacRoles) {
        $rbacList += ($item.LiveRbacRoles -split '; ' | ForEach-Object { "<span class='rbac-role live' title='Assignment is LIVE in Exchange Online'>$(HtmlEnc $_) &#183; live</span>" }) -join ''
    }
    if ($item.RbacRoles) {
        $liveSet = @($item.LiveRbacRoles -split '; ')
        $rbacList += (@($item.RbacRoles -split '; ') | Where-Object { $liveSet -notcontains $_ } |
            ForEach-Object { "<span class='rbac-role' title='Proposed - created by this row''s commands'>$(HtmlEnc $_)</span>" }) -join ''
    }
    if (-not $rbacList) { $rbacList = "<span class='none'>N/A</span>" }

    $commandsHtml = if ($item.MigrationCommands) {
        $lineCount = @($item.MigrationCommands -split "`n").Count
        "<details class='cmd'><summary>PowerShell <span class='cmd-meta'>$lineCount lines</span><button type='button' class='copy-btn'>Copy</button></summary><pre class='commands'>$(HtmlEnc $item.MigrationCommands)</pre></details>"
    }
    else { '' }

    $blockersHtml = if ($item.MigrationBlockers) {
        "<div class='blockers'>&#9888; $(HtmlEnc $item.MigrationBlockers)</div>"
    }
    else { '' }

    $targetTypeClass = ("$($item.Target_Type)".ToLower() -replace '\s', '')
    $targetNameHtml = if ($item.Target_Link) {
        "<a class='tlink' href='$($item.Target_Link)' target='_blank' rel='noopener'>$(HtmlEnc $item.Target_Name) &#8599;</a>"
    } else { HtmlEnc $item.Target_Name }

    $commandsRow = if ($commandsHtml) {
        "`n    <tr class=`"commands-row $rowClass`">`n        <td colspan=`"6`">$commandsHtml</td>`n    </tr>"
    } else { '' }

    @"
    <tr class="$rowClass">
        <td>
            <strong>$(HtmlEnc $item.App_DisplayName)</strong>$issuesBadges<br>
            <span class="app-id">$(HtmlEnc $item.App_ClientId)</span><br>
            $(New-LinkHtml $item)
        </td>
        <td>$statusBadge $arBadge$blockersHtml</td>
        <td>
            <span class="target-type $targetTypeClass">$(HtmlEnc $item.Target_Type)</span><br>
            $targetNameHtml<br>
            <span class="email">$(HtmlEnc $item.Target_Email)</span>
        </td>
        <td>$exoSpStatus</td>
        <td class="perms">$permsList</td>
        <td class="perms">$rbacList</td>
    </tr>$commandsRow
"@
}

# Build security table rows (unconstrained first, then by risk)
$SecuritySorted = $SecurityRows | Sort-Object @{Expression = { switch -Regex ($_.Constraint) {
    '^Unconstrained' { 0 } 'Mailbox-permission' { 1 } default { 2 } } } },
    @{Expression = { switch ($_.RiskTier) { 'high' { 0 } 'medium' { 1 } default { 2 } } } }, DisplayName

$SecurityRowsHtml = foreach ($row in $SecuritySorted) {
    $constraintBadge = switch -Regex ($row.Constraint) {
        '^Unconstrained' { "<span class='status-badge unconstrained'>&#9888; $(HtmlEnc $row.Constraint)</span>" }
        'Mailbox-permission' { "<span class='status-badge review'>&#9888; Verify grants</span>" }
        'Lookup failed' { "<span class='status-badge error'>&#8252; Lookup failed</span>" }
        default { "<span class='status-badge ready'>&#10003; $(HtmlEnc $row.Constraint)</span>" }
    }
    $tierBadge = switch ($row.RiskTier) {
        'high' { "<span class='badge crit'>High</span>" }
        'medium' { "<span class='badge warn'>Medium</span>" }
        default { "<span class='badge neut'>Low</span>" }
    }
    $permChips = ($row.Perms | ForEach-Object { "<span class='perm keep'>$(HtmlEnc $_)</span>" }) -join ''
    $links = "<a class='plink' href='https://entra.microsoft.com/#view/Microsoft_AAD_IAM/ManagedAppMenuBlade/~/Overview/objectId/$($row.SpObjectId)/appId/$($row.AppId)' target='_blank' rel='noopener'>Enterprise app &#8599;</a>"
    $disabledNote = ''
    if ($row.Deactivated) { $disabledNote += " <span class='badge neut'>deactivated</span>" }
    if ($row.Enabled -eq $false) { $disabledNote += " <span class='badge neut'>sign-in disabled</span>" }
    $rowCls = if ($row.Constraint -match '^Unconstrained') { 'status-blocked' } else { '' }
    @"
    <tr class="$rowCls">
        <td><strong>$(HtmlEnc $row.DisplayName)</strong>$disabledNote<br><span class="app-id">$(HtmlEnc $row.AppId)</span><br>$links</td>
        <td>$constraintBadge<div class='blockers-neutral'>$(HtmlEnc $row.Notes)</div></td>
        <td>$tierBadge</td>
        <td class="perms">$permChips</td>
    </tr>
"@
}

$SecurityScanNote = if (-not $SecurityScanOk) {
    "<p class='scan-warn'>&#9888; Parts of the security scan failed (Graph or Exchange query error) - this list may be incomplete. Rerun the audit to confirm.</p>"
} else { '' }

$Html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Application Access Policy Migration Report - $TenantName</title>
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
        /* An app row and its commands row are ONE group: no divider inside the group,
           a full divider only after it - otherwise the pill reads as ambiguous. */
        tbody tr:has(+ tr.commands-row) td { border-bottom: none; }
        tr.commands-row td { padding-top: 0; }
        tr.status-blocked { background: rgba(178, 39, 39, 0.05); }
        tr.status-error { background: rgba(236, 131, 90, 0.07); }
        tr.status-review { background: rgba(250, 178, 25, 0.07); }
        tr.status-delete { background: rgba(137, 135, 129, 0.07); }

        .commands-row td { padding: 0 0.75rem 0.75rem 0.75rem; }

        .app-id { font-family: ui-monospace, "Cascadia Mono", Consolas, monospace; font-size: 0.7rem; color: var(--ink-3); }
        a.plink, a.tlink { color: var(--link); font-size: 0.75rem; text-decoration: none; }
        a.tlink { font-size: 0.875rem; }
        a.plink:hover, a.tlink:hover { text-decoration: underline; }

        .status-badge {
            display: inline-block;
            font-size: 0.7rem;
            font-weight: 600;
            padding: 0.25rem 0.5rem;
            border-radius: 4px;
            text-transform: uppercase;
        }
        .status-badge.ready { background: var(--good-bg); color: var(--good-fg); }
        .status-badge.blocked { background: var(--crit-bg); color: var(--crit-fg); }
        .status-badge.error { background: var(--err-bg); color: var(--err-fg); }
        .status-badge.review { background: var(--warn-bg); color: var(--warn-fg); }
        .status-badge.delete { background: var(--neut-bg); color: var(--neut-fg); }
        .status-badge.unconstrained { background: var(--crit-fg); color: #ffffff; }

        .ar-badge {
            display: inline-block;
            font-size: 0.65rem;
            font-weight: 600;
            padding: 0.15rem 0.4rem;
            border-radius: 4px;
        }
        .ar-badge.restrict { background: var(--neut-bg); color: var(--neut-fg); }
        .ar-badge.deny { background: var(--crit-solid); color: #ffffff; }

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

        .target-type {
            display: inline-block;
            font-size: 0.7rem;
            font-weight: 500;
            padding: 0.2rem 0.4rem;
            border-radius: 4px;
            background: var(--neut-bg);
            color: var(--neut-fg);
        }
        .target-type.group { background: var(--good-bg); color: var(--good-fg); }
        .target-type.usermailbox, .target-type.mailuser { background: var(--info-bg); color: var(--info-fg); }
        .target-type.notfound, .target-type.user { background: var(--crit-bg); color: var(--crit-fg); }

        .email { color: var(--ink-3); font-size: 0.75rem; }
        .sp-exists { color: var(--good-fg); font-size: 0.75rem; }
        .sp-missing { color: var(--warn-fg); font-size: 0.75rem; }

        .perm, .rbac-role {
            display: inline-block;
            padding: 0.1rem 0.35rem;
            border-radius: 3px;
            margin: 0.1rem;
            font-family: ui-monospace, "Cascadia Mono", Consolas, monospace;
            font-size: 0.7rem;
        }
        .perm.migrate { background: var(--info-bg); color: var(--info-fg); }
        .perm.keep { background: var(--neut-bg); color: var(--neut-fg); border: 1px dashed var(--ink-3); }
        .perm.delegated { background: var(--purp-bg); color: var(--purp-fg); }
        .rbac-role { background: var(--good-bg); color: var(--good-fg); }
        .rbac-role.live { background: var(--good-fg); color: #ffffff; }
        .none { color: var(--ink-3); font-style: italic; font-size: 0.75rem; }

        .blockers { margin-top: 0.5rem; font-size: 0.75rem; color: var(--crit-fg); }
        .blockers-neutral { margin-top: 0.5rem; font-size: 0.75rem; color: var(--ink-2); max-width: 42rem; }
        .scan-warn { color: var(--err-fg); font-size: 0.85rem; }

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

        /* Collapsed-by-default command disclosure */
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
        .cmd-tools { margin: 0 0 0.6rem 0; font-size: 0.75rem; color: var(--ink-3); }
        .cmd-tools a { color: var(--link); cursor: pointer; text-decoration: none; }
        .cmd-tools a:hover { text-decoration: underline; }

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
    </style>
</head>
<body>
    <div class="container">
        <h1>Application Access Policy &rarr; RBAC Migration Report</h1>
        <p class="subtitle"><strong>$TenantName</strong> &mdash; $(Get-Date -Format "MMMM d, yyyy 'at' h:mm:ss tt")</p>

        <div class="migration-note">
            <h3>Migration Required</h3>
            <p>Application Access Policies are replaced by RBAC for Applications (Microsoft: don't create
            new policies; no firm retirement date announced yet). This report identifies current policies,
            generates migration commands, and audits the tenant for unconstrained Exchange app access.</p>
            <p><strong>Safety:</strong> Steps 1&ndash;4 are additive. Cutover blocks verify RBAC with
            <code>Test-ServicePrincipalAuthorization</code> and confirm every grant revocation succeeded
            before the (now-inert) policy is removed &mdash; nothing runs unverified. Actions that would restore
            tenant-wide access are gated behind an explicit opt-in variable. Name-matched targets are fully
            commented out until verified. (<code>Remove-ApplicationAccessPolicy</code> is called with
            <code>-Confirm:$false</code> because it does not reliably prompt in the modern EXO module; the
            verification above is the real gate.)</p>
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
            <div class="stat warning">
                <div class="stat-value">$NeedsReview</div>
                <div class="stat-label">Needs Review</div>
            </div>
            <div class="stat danger">
                <div class="stat-value">$Blocked</div>
                <div class="stat-label">Blocked / Error</div>
            </div>
            <div class="stat danger">
                <div class="stat-value">$DenyPolicies</div>
                <div class="stat-label">Deny Policies</div>
            </div>
            <div class="stat">
                <div class="stat-value">$OrphanedApps</div>
                <div class="stat-label">Orphaned</div>
            </div>
            <div class="stat danger">
                <div class="stat-value">$UnconstrainedCount</div>
                <div class="stat-label">Unconstrained Apps</div>
            </div>
        </div>

        <p class="cmd-tools">Commands are collapsed &mdash;
            <a onclick="document.querySelectorAll('details.cmd').forEach(function(d){d.open=true})">expand all</a> &middot;
            <a onclick="document.querySelectorAll('details.cmd').forEach(function(d){d.open=false})">collapse all</a></p>
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

        <h2>Exchange App Permissions Without Mailbox Scoping</h2>
        <p class="section-sub">Every non-Microsoft service principal holding Exchange <em>application</em> permissions.
        Admin consent alone grants <strong>org-wide</strong> mailbox access &mdash; scoping requires a policy or an RBAC
        assignment that the portal never prompts for (insecure by default). Rows marked Unconstrained can reach every
        mailbox in the tenant. $(if ($MsFirstPartySkipped -gt 0) { "$MsFirstPartySkipped Microsoft first-party service principals were excluded." }) $(if ($DanglingGrants -gt 0) { "$DanglingGrants grant(s) pointing at deleted service principals were skipped." })</p>
        $SecurityScanNote
        <table>
            <thead>
                <tr>
                    <th>Application</th>
                    <th>Constraint Status</th>
                    <th>Risk</th>
                    <th>Exchange App Permissions</th>
                </tr>
            </thead>
            <tbody>
                $($SecurityRowsHtml -join "`n")
            </tbody>
        </table>
        <p class="section-sub" style="margin-top:0.75rem">To constrain an unconstrained app: create a management scope for
        the mailboxes it actually needs, add the matching <code>Application *</code> RBAC role assignment (pattern in any
        Ready row above), verify with <code>Test-ServicePrincipalAuthorization</code>, then revoke the tenant-wide Entra
        grant. Order matters: RBAC first, revoke second &mdash; the app never loses access it should have.</p>

        <div class="footer">
            <strong>Migration Steps (per app):</strong><br>
            1. <code>New-ServicePrincipal</code> &mdash; Exchange pointer to the Entra service principal<br>
            2. <code>New-ManagementScope</code> &mdash; mailbox scope (<code>MemberOfGroup</code> DN filter, or
               <code>ExternalDirectoryObjectId</code> for a single mailbox)<br>
            3. <code>New-ManagementRoleAssignment</code> &mdash; assign RBAC application roles with the scope<br>
            4. <code>Test-ServicePrincipalAuthorization</code> &mdash; verify InScope = True, and test the app itself<br>
            5. Revoke <em>only the migrated</em> permission <em>grants</em> in Entra ID (per-permission via Graph; the
               portal's &quot;Revoke admin consent&quot; button revokes everything at once). Removing rows from the app
               registration's API-permissions list afterwards is hygiene only &mdash; it does not revoke access.
               Permissions marked <em>keep</em> are not replaced by RBAC and must stay.<br>
            6. <code>Remove-ApplicationAccessPolicy</code> &mdash; delete the legacy policy last<br><br>
            <strong>Caveats:</strong>
            RBAC scoping only takes effect once the tenant-wide Entra grant is revoked (grants are a union).
            Exchange caches app permissions for 30&nbsp;minutes&ndash;2&nbsp;hours; allow for that between steps 5 and 6.
            <code>MemberOfGroup</code> scopes match <em>direct</em> group members only &mdash; flatten nested groups first.
            DenyAccess policies have no RBAC equivalent and must not be migrated with these steps.
            Policies and RBAC apply to app-only (application) permissions &mdash; apps with only <em>delegated</em>
            Exchange permissions never needed a policy.
            IMAP/POP app permissions have no RBAC roles; that access is controlled via <code>Add-MailboxPermission</code>.
            EWS (<code>Application EWS.AccessAsApp</code>) is blocked for non-Microsoft apps starting Oct&nbsp;1,&nbsp;2026 and
            removed after Apr&nbsp;2027 &mdash; plan Graph migrations for EWS apps.
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

# Reports collect in a dedicated Desktop folder rather than littering the Desktop itself.
$OutDir = Join-Path ([Environment]::GetFolderPath('Desktop')) 'AppAccessPolicyMigration'
if (-not (Test-Path $OutDir)) { $null = New-Item -ItemType Directory -Path $OutDir }
$RunStamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$ExportPath = Join-Path $OutDir "AppAccessPolicyMigration_$TenantName`_$RunStamp.html"
$Html | Out-File -FilePath $ExportPath -Encoding UTF8
Write-Host "Report saved to: $ExportPath" -ForegroundColor Green

# Companion script: all generated command blocks. Steps 1-4 run as-is (additive); each
# cutover verifies RBAC and confirms revocations succeeded before removing the inert policy.
$CommandBlocks = $SortedReport | Where-Object { $_.MigrationCommands } | ForEach-Object { $_.MigrationCommands + "`n" }
if ($CommandBlocks) {
    $ScriptPath = Join-Path $OutDir "AppAccessPolicyMigration_$TenantName`_$RunStamp.ps1"
    @(
        "# Application Access Policy -> RBAC for Applications migration commands"
        "# Generated $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') for tenant $TenantName by Audit-ExoAppAccessPolicies.ps1"
        "# Prereqs: Connect-ExchangeOnline (Organization Management / Exchange Administrator);"
        "#          Connect-MgGraph -Scopes 'AppRoleAssignment.ReadWrite.All' (for the cutover revocations)."
        "# Steps 1-4 are additive. Cutover blocks self-verify with Test-ServicePrincipalAuthorization,"
        "# skip themselves if RBAC is not live, and remove the inert policy only after revocations succeed."
        "# Blocks for name-matched targets are fully commented out - verify the target first."
        ""
    ) + $CommandBlocks | Out-File -FilePath $ScriptPath -Encoding UTF8
    Write-Host "Migration commands saved to: $ScriptPath" -ForegroundColor Green
}

Start-Process $ExportPath

#endregion Output
