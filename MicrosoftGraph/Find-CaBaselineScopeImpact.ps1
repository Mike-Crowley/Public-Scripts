#Requires -Version 5.1
#Requires -Modules Microsoft.Graph.Authentication

# ============================================================================
# Find-CaBaselineScopeImpact.ps1
#
# Identifies applications likely to break under Microsoft's Conditional Access
# "baseline scopes" enforcement change (rollout began June 15, 2026):
#   https://learn.microsoft.com/en-us/entra/identity/conditional-access/concept-enforcement-resource-exclusions
#
# The change: when an ALL-resources CA policy has one or more resource
# exclusions, sign-ins that request ONLY baseline scopes (openid, profile,
# email, offline_access, User.Read, User.Read.All, User.ReadBasic.All,
# People.Read, People.Read.All, GroupMember.Read.All, Member.Read.Hidden)
# used to bypass the policy entirely. After enforcement they are evaluated
# as directory access (audience: Windows Azure Active Directory) and hit the
# policy's controls (MFA, compliant device, block).
#
# Who breaks, per the doc:
#   1. PUBLIC clients requesting only baseline scopes (e.g. VS Code requests
#      openid+profile; Azure CLI requests only User.Read) - affected whether
#      or not they are excluded anywhere.
#   2. CONFIDENTIAL clients that are EXCLUDED from an All-resources policy
#      and consented only to baseline DIRECTORY scopes (OIDC-only
#      confidential clients are explicitly unaffected).
#
# Detection strategy (100k-user friendly - NO bulk sign-in log pull):
#   - CA policies:            one small query
#   - oauth2PermissionGrants: one flat paged pull (the delegated-consent
#                             footprint is a static proxy for "what scopes
#                             does this client request"); -SkipPerUserConsents
#                             restricts to admin (AllPrincipals) grants if the
#                             tenant has heavy per-user consent volume
#   - service principals:     resolved only for the candidates, via
#                             getByIds / appId-filter batches
#   - sign-in logs:           OFF by default; -IncludeSignInSample runs one
#                             bounded query per flagged app (top 100, last
#                             -SignInSampleDays days, default 7) via Graph
#                             $batch, 20 queries per round trip
#
# Coverage note: consent grants cannot prove what an app requests at runtime.
# For authoritative detection Microsoft's method is the Customize-behavior
# placeholder app + conditionalAccessAudiences sign-in filter (see doc). This
# script is the fast, read-only first pass.
#
# Tenant enforcement state: the "Baseline scope settings (Preview)" blade
# (aka.ms/BaselineScopesSettingsUX, or aka.ms/BaselineScopesSettingsUX-gov;
# guidance at aka.ms/BaselineScopesSettings) is hidden unless reached via the
# direct link, which appends feature.isbaselinescopesenabled=true. A HAR
# capture of the blade (2026-08-03) shows it reads the setting via
#   GET /beta/identity/conditionalAccess/settings
#   (x-ms-command-name: PolicyManagement - GetBaselineScopesSettings)
# When no ADMIN selection has been saved, the API synthesizes the object at
# read time: modifiedDateTime equals the query time and exclusions and
# advancedSettings are null. That null state is ambiguous: either Microsoft's
# rollout has not reached the tenant yet (GCC tenants often land late in the
# waves even when they are not GCC High), or the rollout completed and
# enforcement is now the silent default - automatic enforcement records no
# selection, so the API looks identical either way. The blade itself shows if
# and when Microsoft enabled enforcement for the tenant.
# When a selection HAS been saved, the choice lives at
# advancedSettings.baselineScopes with createdDateTime = when the option was
# selected and resourceAppId = the placeholder app for Customize behavior, or
# the all-zeros GUID for Disable enforcement (field-verified: a tenant where
# Microsoft's rollout enabled enforcement on 2026-07-27 and admins then
# disabled it on 2026-07-31 shows zeros). The endpoint is marked
# PrivatePreview with Deprecation/Sunset headers, so this script probes it
# best-effort and degrades gracefully if it moves.
#
# READ-ONLY: only GET requests plus the getByIds read POST. Nothing in the
# tenant is created, changed, or deleted. Scopes requested are *.Read.All.
#
# Output: timestamped HTML report + CSV in Desktop\CaBaselineScopeImpact\
# ============================================================================

[CmdletBinding()]
param(
    # Pin the sign-in to a specific tenant (recommended when you hold guest access elsewhere).
    [string]$TenantId,

    # Device-code auth for hosts where interactive/WAM auth fails (VS Code integrated console, ISE, ssh).
    [switch]$UseDeviceCode,

    # Per-app bounded sign-in queries for flagged apps (adds AuditLog.Read.All to the consent).
    # Default window is 7 days: sign-in log query cost scales with the window,
    # and 7 days is enough to separate live apps from dormant ones.
    [switch]$IncludeSignInSample,
    [ValidateRange(1, 30)][int]$SignInSampleDays = 7,
    [ValidateRange(1, 200)][int]$MaxAppsToSample = 40,

    # Only admin-consented (AllPrincipals) grants; skips per-user consents for speed in huge tenants.
    [switch]$SkipPerUserConsents,

    # Graph national cloud (Global covers commercial and GCC; USGov = GCC High, USGovDoD = DoD).
    [ValidateSet('Global', 'USGov', 'USGovDoD', 'China')]
    [string]$Environment = 'Global',

    # Defaults to Desktop\CaBaselineScopeImpact
    [string]$OutputFolder
)

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
$OidcScopes = @('openid', 'profile', 'email', 'offline_access')
$DirectoryBaselineScopes = @('User.Read', 'User.Read.All', 'User.ReadBasic.All',
    'People.Read', 'People.Read.All', 'GroupMember.Read.All', 'Member.Read.Hidden')
$BaselineLower = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
foreach ($s in ($OidcScopes + $DirectoryBaselineScopes)) { [void]$BaselineLower.Add($s) }
$DirectoryLower = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
foreach ($s in $DirectoryBaselineScopes) { [void]$DirectoryLower.Add($s) }

$MsGraphAppId  = '00000003-0000-0000-c000-000000000000'
$AadGraphAppId = '00000002-0000-0000-c000-000000000000'   # Windows Azure Active Directory
$graphBase = 'https://graph.microsoft.com'                # replaced after connect for national clouds
$MicrosoftTenantIds = @('f8cdef31-a31e-4b4a-93e4-5f571e91255a', '72f988bf-86f1-41af-91ab-2d7cd011db47')

# Well-known first-party PUBLIC clients (ISV/Microsoft apps have no local app
# object, so client type is otherwise unknowable statically).
$KnownPublicClients = @{
    '04b07795-8ddb-461a-bbee-02f9e1bf7b46' = 'Microsoft Azure CLI'
    '1950a258-227b-4e31-a9cf-717495945fc9' = 'Microsoft Azure PowerShell'
    'aebc6443-996d-45c2-90f0-388ff96faa56' = 'Visual Studio Code'
    '872cd9fa-d31f-45e0-9eab-6e460a02d1f1' = 'Visual Studio'
    '14d82eec-204b-4c2f-b7e8-296a70dab67e' = 'Microsoft Graph PowerShell'
    '1b730954-1685-4b74-9bfd-dac224a7b894' = 'Azure AD PowerShell (legacy)'
    'd3590ed6-52b3-4102-aeff-aad2292ab01c' = 'Microsoft Office'
    '27922004-5251-4030-b22d-91ecd9a37ea4' = 'Outlook Mobile'
    'f8d98a96-0999-43f5-8af3-69971c7bb423' = 'Apple Internet Accounts (iOS Accounts)'
    'ed42a417-9cae-4d2d-a834-84ec9a60bb53' = 'GroupMe'
}

function Test-IsGuid { param([string]$Value) $g = [guid]::Empty; [guid]::TryParse($Value, [ref]$g) }
function EscHtml { param($s) [System.Net.WebUtility]::HtmlEncode("$s") }

function Get-GraphPages {
    param([string]$Uri, [string]$Label)
    # ::new(), not New-Object: PS 7.6.4's binder throws 'Argument types do not
    # match' when @() enumerates a PSObject-wrapped generic List downstream.
    $items = [System.Collections.Generic.List[object]]::new()
    $next = $Uri
    $page = 0
    while ($next) {
        $resp = Invoke-MgGraphRequest -Method GET -Uri $next -OutputType PSObject
        if ($resp.value) { foreach ($v in $resp.value) { $items.Add($v) } }
        $next = $resp.'@odata.nextLink'
        $page++
        if ($page % 10 -eq 0) { Write-Host "    $Label... $($items.Count) records so far" }
    }
    return , $items.ToArray()
}

# Resolve service principals by OBJECT id in batches of 1000 (read-only POST).
function Resolve-SpByObjectId {
    param([string[]]$Ids)
    $map = @{}
    for ($i = 0; $i -lt $Ids.Count; $i += 1000) {
        $batch = $Ids[$i..([Math]::Min($i + 999, $Ids.Count - 1))]
        $body = @{ ids = @($batch); types = @('servicePrincipal') } | ConvertTo-Json -Depth 3
        $resp = Invoke-MgGraphRequest -Method POST -Uri "$graphBase/v1.0/directoryObjects/getByIds" -Body $body -ContentType 'application/json' -OutputType PSObject
        foreach ($o in @($resp.value)) { $map[$o.id] = $o }
    }
    return $map
}

# Resolve service principals by APP id in filter batches (in-operator, small batches).
function Resolve-SpByAppId {
    param([string[]]$AppIds)
    $map = @{}
    $guids = @($AppIds | Where-Object { Test-IsGuid $_ } | Sort-Object -Unique)
    for ($i = 0; $i -lt $guids.Count; $i += 10) {
        $batch = $guids[$i..([Math]::Min($i + 9, $guids.Count - 1))]
        $inList = ($batch | ForEach-Object { "'$_'" }) -join ','
        $uri = "$graphBase/v1.0/servicePrincipals?`$filter=appId in ($inList)&`$select=id,appId,displayName,appOwnerOrganizationId,accountEnabled&`$top=999"
        $resp = Invoke-MgGraphRequest -Method GET -Uri $uri -OutputType PSObject
        foreach ($o in @($resp.value)) { $map[$o.appId] = $o }
    }
    return $map
}

# ---------------------------------------------------------------------------
# Connect (read-only scopes; Graph SDK falls back to device code on WAM-hostile hosts)
# ---------------------------------------------------------------------------
$scopes = @('Policy.Read.All', 'Directory.Read.All')
if ($IncludeSignInSample) { $scopes += 'AuditLog.Read.All' }

$connect = @{ Scopes = $scopes; NoWelcome = $true; Environment = $Environment }
if ($TenantId) { $connect.TenantId = $TenantId }
if ($UseDeviceCode) { $connect.UseDeviceCode = $true }

Write-Host 'Signing in to Microsoft Graph (read-only)...' -ForegroundColor Cyan
Connect-MgGraph @connect

$ctx = Get-MgContext
if (-not $ctx -or -not $ctx.TenantId) { throw 'Graph sign-in did not produce a usable context.' }
$tenant = $ctx.TenantId
try {
    $envInfo = Get-MgEnvironment -Name $ctx.Environment -ErrorAction Stop
    if ($envInfo.GraphEndpoint) { $graphBase = "$($envInfo.GraphEndpoint)".TrimEnd('/') }
} catch { }
$org = Invoke-MgGraphRequest -Method GET -Uri "$graphBase/v1.0/organization?`$select=displayName,id" -OutputType PSObject
$orgName = "$(@($org.value)[0].displayName)"
Write-Host "Connected: $orgName ($tenant) via $graphBase" -ForegroundColor Green

# ---------------------------------------------------------------------------
# Step 0: Tenant enforcement state (best-effort; private-preview endpoint the
# Baseline scope settings blade reads; see header for provenance)
# ---------------------------------------------------------------------------
$baselineSettings = $null
$baselineStateKind = 'Unknown'
$baselinePlaceholderAppId = ''
$baselineState = 'Could not read the Baseline scopes setting (private-preview endpoint unavailable to this account, tenant, or cloud). Check the blade manually.'
try {
    $baselineSettings = Invoke-MgGraphRequest -Method GET -Uri "$graphBase/beta/identity/conditionalAccess/settings" -OutputType PSObject
    $bs = $baselineSettings.advancedSettings.baselineScopes
    if ($null -eq $baselineSettings.exclusions -and $null -eq $baselineSettings.advancedSettings) {
        $baselineStateKind = 'DefaultRollout'
        $baselineState = 'No admin selection saved in Baseline scope settings. That means either the rollout has not reached this tenant yet (GCC tenants often land late in the waves), or it already completed and enforcement is now the silent default; the API looks identical in both cases because automatic enforcement records no selection. Open the blade, which shows if and when Microsoft enabled enforcement for this tenant. The modifiedDateTime this API returns is synthesized at read time; it is not a selection date.'
    }
    elseif ($bs) {
        $baselineStateKind = 'SelectionSaved'
        $rid = "$($bs.resourceAppId)"
        if ($rid -eq $AadGraphAppId) {
            $baselineState = "Enforcement was explicitly enabled on $($bs.createdDateTime) (baseline scopes target Windows Azure Active Directory)."
        }
        elseif ($rid -and $rid -ne '00000000-0000-0000-0000-000000000000') {
            $baselinePlaceholderAppId = $rid
            $baselineState = "Customize behavior is active (selected $($bs.createdDateTime)): baseline scopes are evaluated against placeholder app $rid. Legacy behavior is retained ONLY for policies that exclude that app; verify each triggering policy's exclusions include it where intended."
        }
        else {
            $baselineState = "A selection was saved on $($bs.createdDateTime) with resourceAppId set to the all-zeros GUID. Field evidence maps this shape to Disable enforcement (legacy behavior retained for all policies), which Microsoft does not recommend. Confirm in the blade, which also shows if and when Microsoft's rollout enabled enforcement for this tenant."
        }
    }
    else {
        $baselineStateKind = 'SelectionSaved'
        $baselineState = "Settings object contains saved values in an unrecognized shape (modified $($baselineSettings.modifiedDateTime)); raw JSON is in the report."
    }
    Write-Host "Baseline scopes setting: $baselineStateKind" -ForegroundColor Yellow
} catch {
    Write-Warning "Baseline scopes settings probe failed: $($_.Exception.Message)"
}

$now = Get-Date
$stamp = $now.ToString('yyyyMMdd_HHmmss')
if (-not $OutputFolder) { $OutputFolder = Join-Path ([Environment]::GetFolderPath('Desktop')) 'CaBaselineScopeImpact' }
if (-not (Test-Path $OutputFolder)) { $null = New-Item -ItemType Directory -Path $OutputFolder -Force }
$orgSafe = ($orgName -replace '[^A-Za-z0-9]+', '')
if (-not $orgSafe) { $orgSafe = 'tenant' }
$htmlPath = Join-Path $OutputFolder "CaBaselineScopeImpact_${orgSafe}_$stamp.html"
$csvPath  = Join-Path $OutputFolder "CaBaselineScopeImpact_${orgSafe}_$stamp.csv"

# ---------------------------------------------------------------------------
# Step 1: Conditional Access policies - find the triggering shape
#         (targets All resources AND has resource exclusions)
# ---------------------------------------------------------------------------
Write-Host 'Step 1: Reading Conditional Access policies...' -ForegroundColor Cyan
$policies = Get-GraphPages -Uri "$graphBase/v1.0/identity/conditionalAccess/policies" -Label 'CA policies'

function Get-PolicyControls {
    param($p)
    $parts = @()
    $sevRank = 0
    $g = $p.grantControls
    if ($g) {
        foreach ($c in @($g.builtInControls)) {
            switch ($c) {
                'block'                { $parts += 'Block access';            if ($sevRank -lt 4) { $sevRank = 4 } }
                'compliantDevice'      { $parts += 'Require compliant device'; if ($sevRank -lt 3) { $sevRank = 3 } }
                'domainJoinedDevice'   { $parts += 'Require hybrid-joined device'; if ($sevRank -lt 3) { $sevRank = 3 } }
                'compliantApplication' { $parts += 'Require app protection policy'; if ($sevRank -lt 3) { $sevRank = 3 } }
                'approvedApplication'  { $parts += 'Require approved client app'; if ($sevRank -lt 3) { $sevRank = 3 } }
                'mfa'                  { $parts += 'Require MFA';             if ($sevRank -lt 2) { $sevRank = 2 } }
                'passwordChange'       { $parts += 'Require password change'; if ($sevRank -lt 2) { $sevRank = 2 } }
                default                { if ($c) { $parts += "$c" } }
            }
        }
        if ($g.authenticationStrength -and $g.authenticationStrength.displayName) {
            $parts += "Auth strength: $($g.authenticationStrength.displayName)"
            if ($sevRank -lt 2) { $sevRank = 2 }
        }
    }
    if (-not $parts -and $p.sessionControls) { $parts = @('Session controls only'); $sevRank = 1 }
    if (-not $parts) { $parts = @('(no grant controls)') }
    [pscustomobject]@{ Text = ($parts -join ' + '); Rank = $sevRank }
}

function Get-PolicyUserSummary {
    param($p)
    $u = $p.conditions.users
    if (-not $u) { return '(no user condition)' }
    # @($null).Count is 1, so drop null/empty values before counting
    $inc = @()
    if (@($u.includeUsers) -contains 'All') { $inc += 'All users' }
    else {
        $n = @($u.includeUsers  | Where-Object { $_ }).Count; if ($n) { $inc += "$n users" }
        $n = @($u.includeGroups | Where-Object { $_ }).Count; if ($n) { $inc += "$n groups" }
        $n = @($u.includeRoles  | Where-Object { $_ }).Count; if ($n) { $inc += "$n roles" }
    }
    if (-not $inc) { $inc = @('(none)') }
    $exc = @()
    $n = @($u.excludeUsers  | Where-Object { $_ }).Count; if ($n) { $exc += "$n users" }
    $n = @($u.excludeGroups | Where-Object { $_ }).Count; if ($n) { $exc += "$n groups" }
    $n = @($u.excludeRoles  | Where-Object { $_ }).Count; if ($n) { $exc += "$n roles" }
    $s = ($inc -join ', ')
    if ($exc) { $s += " (excl. $($exc -join ', '))" }
    return $s
}

$triggering = @()
foreach ($p in $policies) {
    $apps = $p.conditions.applications
    if (-not $apps) { continue }
    if ($p.state -notin @('enabled', 'enabledForReportingButNotEnforced')) { continue }
    $inc = @($apps.includeApplications)
    $exc = @($apps.excludeApplications) | Where-Object { $_ -and $_ -ne 'None' }
    $reason = ''
    if (($inc -contains 'All') -and $exc.Count -gt 0) { $reason = 'All resources + resource exclusions' }
    elseif ($inc -contains $AadGraphAppId) { $reason = 'Explicitly targets Windows Azure Active Directory' }
    if (-not $reason) { continue }
    $ctl = Get-PolicyControls $p
    $triggering += [pscustomobject]@{
        Id           = $p.id
        Name         = $p.displayName
        State        = $p.state
        Enforced     = ($p.state -eq 'enabled')
        Reason       = $reason
        Exclusions   = $exc
        Controls     = $ctl.Text
        SevRank      = $ctl.Rank
        Users        = Get-PolicyUserSummary $p
    }
}

# The behavior CHANGE fires only for All-resources-with-exclusions policies;
# WAAD-targeting policies are listed because their controls also hit
# baseline-scope sign-ins once those are evaluated as directory access.
$changeTriggers = @($triggering | Where-Object { $_.Reason -like 'All resources*' })
$enabledTriggering = @($triggering | Where-Object Enforced)
$maxSev = 0
foreach ($t in $enabledTriggering) { if ($t.SevRank -gt $maxSev) { $maxSev = $t.SevRank } }
$sevLabel = switch ($maxSev) { 4 { 'Critical' } 3 { 'High' } 2 { 'Medium' } 1 { 'Low' } default { 'None' } }

Write-Host ("  {0} CA policies total; {1} target All resources WITH exclusions; {2} explicitly target Windows Azure AD" -f `
    @($policies).Count, $changeTriggers.Count, (@($triggering).Count - $changeTriggers.Count))

# Union of excluded resource appIds across change-triggering policies, with display names.
$excludedAppIds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$excludedNonGuid = @{}
foreach ($t in $changeTriggers) {
    foreach ($e in $t.Exclusions) {
        if (Test-IsGuid $e) { [void]$excludedAppIds.Add($e) } else { $excludedNonGuid[$e] = $true }
    }
}
$excludedSpMap = @{}
if ($excludedAppIds.Count) {
    Write-Host '  Resolving excluded-resource display names...'
    $excludedSpMap = Resolve-SpByAppId -AppIds @($excludedAppIds)
}
if ($baselinePlaceholderAppId) {
    try {
        $phMap = Resolve-SpByAppId -AppIds @($baselinePlaceholderAppId)
        if ($phMap.ContainsKey($baselinePlaceholderAppId)) {
            $baselineState = $baselineState -replace [regex]::Escape("placeholder app $baselinePlaceholderAppId"), "placeholder app '$($phMap[$baselinePlaceholderAppId].displayName)' ($baselinePlaceholderAppId)"
        }
    } catch { }
}
function Get-ExclusionLabel {
    param([string]$Value)
    if (-not (Test-IsGuid $Value)) { return "$Value (app group)" }
    if ($excludedSpMap.ContainsKey($Value)) { return $excludedSpMap[$Value].displayName }
    return "$Value (no service principal found)"
}

# ---------------------------------------------------------------------------
# Step 2: Delegated consent footprints (the only tenant-wide pull; flat and paged)
# ---------------------------------------------------------------------------
$rows = @()
$clientsScanned = 0
$baselineOnlyCount = 0
$oidcOnlyConfidential = 0

if (-not $changeTriggers) {
    Write-Host 'No All-resources policies with resource exclusions found. This tenant is NOT affected by the enforcement change; skipping app analysis.' -ForegroundColor Green
}
else {
    Write-Host 'Step 2: Pulling delegated permission grants (consent footprints)...' -ForegroundColor Cyan
    $grantUri = "$graphBase/v1.0/oauth2PermissionGrants?`$select=clientId,consentType,resourceId,scope&`$top=999"
    if ($SkipPerUserConsents) {
        $grantUri = "$graphBase/v1.0/oauth2PermissionGrants?`$filter=consentType eq 'AllPrincipals'&`$select=clientId,consentType,resourceId,scope&`$top=999"
    }
    $grants = Get-GraphPages -Uri $grantUri -Label 'grants'
    Write-Host "  $(@($grants).Count) grant records"

    # Aggregate per client: union of scopes, set of resource SPs
    $byClient = @{}
    foreach ($g in $grants) {
        if (-not $g.clientId) { continue }
        if (-not $byClient.ContainsKey($g.clientId)) {
            $byClient[$g.clientId] = @{
                Scopes    = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
                Resources = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
            }
        }
        foreach ($s in ("$($g.scope)" -split '\s+')) { if ($s) { [void]$byClient[$g.clientId].Scopes.Add($s) } }
        if ($g.resourceId) { [void]$byClient[$g.clientId].Resources.Add($g.resourceId) }
    }
    $clientsScanned = $byClient.Count
    Write-Host "  $clientsScanned distinct client apps with delegated consent"

    # Which resource SPs are Graph / AAD Graph (baseline scopes only exist there)?
    Write-Host '  Resolving resource service principals...'
    $allResourceIds = @($byClient.Values | ForEach-Object { $_.Resources } | ForEach-Object { $_ } | Sort-Object -Unique)
    $resourceMap = Resolve-SpByObjectId -Ids $allResourceIds

    # Classify: candidates = clients whose entire consent footprint is baseline
    $candidates = @{}
    foreach ($clientId in $byClient.Keys) {
        $c = $byClient[$clientId]
        $beyond = $false
        foreach ($rid in $c.Resources) {
            $rApp = ''
            if ($resourceMap.ContainsKey($rid)) { $rApp = "$($resourceMap[$rid].appId)" }
            if ($rApp -ne $MsGraphAppId -and $rApp -ne $AadGraphAppId) { $beyond = $true; break }
        }
        if (-not $beyond) {
            foreach ($s in $c.Scopes) { if (-not $BaselineLower.Contains($s)) { $beyond = $true; break } }
        }
        if ($beyond) { continue }
        if ($c.Scopes.Count -eq 0) { continue }
        $hasDirectory = $false
        foreach ($s in $c.Scopes) { if ($DirectoryLower.Contains($s)) { $hasDirectory = $true; break } }
        $candidates[$clientId] = [pscustomobject]@{
            Scopes       = (@($c.Scopes) | Sort-Object) -join ' '
            HasDirectory = $hasDirectory
        }
    }
    $baselineOnlyCount = $candidates.Count
    Write-Host "  $baselineOnlyCount clients consented ONLY to baseline scopes" -ForegroundColor Yellow

    # ---------------------------------------------------------------------------
    # Step 3: Resolve candidate clients and classify impact
    # ---------------------------------------------------------------------------
    Write-Host 'Step 3: Classifying candidate clients...' -ForegroundColor Cyan
    $clientSpMap = @{}
    if ($candidates.Count) { $clientSpMap = Resolve-SpByObjectId -Ids @($candidates.Keys) }

    # Batch-resolve app registration objects for tenant-owned candidates
    # (client-type detection): one 'appId in' query per 15 apps, not one each.
    $appObjMap = @{}
    $tenantOwnedAppIds = @()
    foreach ($clientId in $candidates.Keys) {
        $sp = $clientSpMap[$clientId]
        if ($sp -and "$($sp.appOwnerOrganizationId)" -eq $tenant -and $sp.appId) { $tenantOwnedAppIds += "$($sp.appId)" }
    }
    $uniqueOwned = @($tenantOwnedAppIds | Sort-Object -Unique)
    for ($i = 0; $i -lt $uniqueOwned.Count; $i += 15) {
        $batch = $uniqueOwned[$i..([Math]::Min($i + 14, $uniqueOwned.Count - 1))]
        $inList = ($batch | ForEach-Object { "'$_'" }) -join ','
        try {
            $appResp = Invoke-MgGraphRequest -Method GET -Uri "$graphBase/v1.0/applications?`$filter=appId in ($inList)&`$select=appId,publicClient,spa,web,isFallbackPublicClient&`$top=999" -OutputType PSObject
            foreach ($a in @($appResp.value)) { $appObjMap["$($a.appId)"] = $a }
        } catch { }
        if ($i -gt 0 -and ($i / 15) % 10 -eq 0) { Write-Host "    app registrations... $([Math]::Min($i + 15, $uniqueOwned.Count))/$($uniqueOwned.Count)" }
    }

    foreach ($clientId in $candidates.Keys) {
        $cand = $candidates[$clientId]
        $sp = $null
        if ($clientSpMap.ContainsKey($clientId)) { $sp = $clientSpMap[$clientId] }
        $appId = ''
        $name = "(service principal $clientId not found)"
        $enabled = $true
        $ownerOrg = ''
        if ($sp) { $appId = "$($sp.appId)"; $name = "$($sp.displayName)"; $enabled = [bool]$sp.accountEnabled; $ownerOrg = "$($sp.appOwnerOrganizationId)" }

        $ownership = 'ISV'
        if ($ownerOrg -eq $tenant) { $ownership = 'Tenant-owned' }
        elseif ($MicrosoftTenantIds -contains $ownerOrg) { $ownership = 'Microsoft' }

        # Client type: app object for tenant-owned; known-list for foreign; else Unknown
        $clientType = 'Unknown'
        if ($ownership -eq 'Tenant-owned' -and $appId) {
            $app = $appObjMap[$appId]
            if ($app) {
                $isPublic = $false
                if ($app.isFallbackPublicClient) { $isPublic = $true }
                if (@($app.publicClient.redirectUris | Where-Object { $_ }).Count -gt 0) { $isPublic = $true }
                if (@($app.spa.redirectUris | Where-Object { $_ }).Count -gt 0) { $isPublic = $true }
                if ($isPublic) { $clientType = 'Public' } else { $clientType = 'Confidential' }
            }
        }
        elseif ($KnownPublicClients.ContainsKey($appId)) { $clientType = 'Public' }

        $isExcluded = $appId -and $excludedAppIds.Contains($appId)

        # Verdict per the doc's affected-scenarios table. Unverifiable ISV
        # clients that are NOT excluded get a lower 'Possible' tier: most are
        # multi-tenant web SSO integrations (confidential clients), which are
        # unaffected unless excluded - flagging them all as Affected would bury
        # the real breakage list in large tenants.
        $verdict = ''
        $action = ''
        if ($clientType -eq 'Public' -or ($clientType -eq 'Unknown' -and $isExcluded)) {
            $verdict = 'Affected'
            if ($clientType -eq 'Unknown') { $verdict = 'Affected (verify client type)' }
            switch ($ownership) {
                'Microsoft' { $action = "Interactive Microsoft client: users will start receiving the policy's challenge. Breaks where the policy blocks access or requires a managed device and devices are unmanaged, or where this client runs headless/automated. Review who uses it and from where." }
                'Tenant-owned' { $action = 'Confirm the app can handle Conditional Access challenges (MFA / device claims). If it cannot, update it, or retain legacy behavior for the specific policy via the Customize behavior placeholder app.' }
                default { $action = 'Verify with the vendor whether this client handles Conditional Access challenges. If it cannot, retain legacy behavior for the specific policy via the Customize behavior placeholder app until it is fixed.' }
            }
        }
        elseif ($clientType -eq 'Unknown') {
            $verdict = 'Possible'
            $action = 'Affected only if this is actually a public (native or desktop) client. Multi-tenant web SSO integrations are confidential clients and are unaffected unless excluded from an All-resources policy. If users reach this app through a desktop or mobile client that requests only these scopes, treat it as Affected; sign-in activity below helps triage dormant entries.'
        }
        elseif ($clientType -eq 'Confidential' -and $cand.HasDirectory -and $isExcluded) {
            $verdict = 'Affected'
            if ($ownership -eq 'Tenant-owned') {
                $action = 'Excluded confidential client consented only to baseline directory scopes: ask the developers to request OIDC scopes (openid, profile) instead of User.Read-style scopes. OIDC-only confidential clients are explicitly unaffected. Otherwise use Customize behavior.'
            } else {
                $action = 'Excluded confidential client consented only to baseline directory scopes: engage the ISV about moving to OIDC scopes. If they cannot update in time, retain legacy behavior for the policy via Customize behavior.'
            }
        }
        elseif ($clientType -eq 'Confidential' -and $cand.HasDirectory) {
            $verdict = 'Monitor'
            $action = 'Not excluded from any triggering policy, so its Microsoft Graph sign-ins are already CA-enforced today; no change expected. Listed because its footprint is baseline-directory-only: if it is ever added to an exclusion, it lands in the affected set.'
        }
        else {
            # Confidential + OIDC-only: explicitly unaffected
            $oidcOnlyConfidential++
            continue
        }

        $rows += [pscustomobject]@{
            App            = $name
            AppId          = $appId
            SpObjectId     = $clientId
            ClientType     = $clientType
            Ownership      = $ownership
            SignInEnabled  = $enabled
            ConsentedScopes = $cand.Scopes
            ExcludedFromPolicy = $isExcluded
            Verdict        = $verdict
            Severity       = $(if ($verdict -like 'Affected*') { $sevLabel } elseif ($verdict -eq 'Possible') { 'Low' } else { 'Info' })
            Action         = $action
            RecentSignIns  = ''
            RecentUsers    = ''
            LastSignIn     = ''
            CaFailures     = ''
        }
    }
}

# Default report order: severity first, then app name (headers re-sort client-side)
$SevRankMap = @{ 'Critical' = 5; 'High' = 4; 'Medium' = 3; 'Low' = 2; 'Info' = 1 }
$affectedRows = @($rows | Where-Object { $_.Verdict -like 'Affected*' } | Sort-Object @{ e = { $SevRankMap[$_.Severity] }; Descending = $true }, App)
$monitorRows  = @($rows | Where-Object { $_.Verdict -eq 'Monitor' } | Sort-Object App)
$possibleRows = @($rows | Where-Object { $_.Verdict -eq 'Possible' } | Sort-Object App)

# ---------------------------------------------------------------------------
# Step 4 (optional): bounded per-app sign-in samples for flagged apps
# ---------------------------------------------------------------------------
if ($IncludeSignInSample -and ($affectedRows.Count + $possibleRows.Count) -gt 0) {
    Write-Host "Step 4: Sampling sign-ins for up to $MaxAppsToSample flagged apps (last $SignInSampleDays days, interactive only, batched 20 per request)..." -ForegroundColor Cyan
    # InvariantCulture: ':' is a culture placeholder in .NET format strings, and
    # locales like fi-FI render it as '.', which Graph rejects in the filter
    $since = $now.AddDays(-$SignInSampleDays).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ', [System.Globalization.CultureInfo]::InvariantCulture)
    # Affected apps first; Possible (unverified ISV) apps fill remaining slots.
    # Sign-in-disabled SPs cannot produce sign-ins: record zeros, skip the query.
    foreach ($r in @($affectedRows) + @($possibleRows)) {
        if (-not $r.SignInEnabled) { $r.RecentSignIns = '0'; $r.RecentUsers = '0'; $r.CaFailures = '0' }
    }
    $targets = @(@($affectedRows) + @($possibleRows) | Where-Object { $_.AppId -and $_.SignInEnabled } | Select-Object -First $MaxAppsToSample)
    for ($i = 0; $i -lt $targets.Count; $i += 20) {
        $chunk = @($targets[$i..([Math]::Min($i + 19, $targets.Count - 1))])
        $reqs = @()
        for ($j = 0; $j -lt $chunk.Count; $j++) {
            $filter = "appId eq '$($chunk[$j].AppId)' and createdDateTime ge $since"
            $reqs += @{ id = "$j"; method = 'GET'; url = "/auditLogs/signIns?`$filter=$([uri]::EscapeDataString($filter))&`$top=100&`$select=createdDateTime,userPrincipalName,conditionalAccessStatus" }
        }
        $batchResp = $null
        try {
            $body = @{ requests = $reqs } | ConvertTo-Json -Depth 5
            $batchResp = Invoke-MgGraphRequest -Method POST -Uri "$graphBase/v1.0/`$batch" -Body $body -ContentType 'application/json' -OutputType PSObject
        } catch { Write-Warning "  Sign-in batch request failed: $($_.Exception.Message)" }
        for ($j = 0; $j -lt $chunk.Count; $j++) {
            $r = $chunk[$j]
            $events = $null
            $item = $null
            if ($batchResp) { $item = @($batchResp.responses) | Where-Object { "$($_.id)" -eq "$j" } | Select-Object -First 1 }
            if ($item -and $item.status -eq 200) { $events = @($item.body.value) }
            else {
                # Per-app fallback: some tenants only serve sign-ins on beta
                try {
                    $filter = "appId eq '$($r.AppId)' and createdDateTime ge $since"
                    $single = Invoke-MgGraphRequest -Method GET -Uri "$graphBase/beta/auditLogs/signIns?`$filter=$([uri]::EscapeDataString($filter))&`$top=100&`$select=createdDateTime,userPrincipalName,conditionalAccessStatus" -OutputType PSObject
                    $events = @($single.value)
                } catch { }
            }
            if ($null -eq $events) { $r.RecentSignIns = 'query failed'; continue }
            $r.RecentSignIns = "$($events.Count)"
            if ($events.Count -eq 100) { $r.RecentSignIns = '100+' }
            if ($events.Count) {
                $r.RecentUsers = "$(@($events.userPrincipalName | Sort-Object -Unique).Count)"
                $r.LastSignIn = "$(@($events | Sort-Object createdDateTime -Descending)[0].createdDateTime)"
                $r.CaFailures = "$(@($events | Where-Object { $_.conditionalAccessStatus -eq 'failure' }).Count)"
            } else {
                $r.RecentUsers = '0'; $r.CaFailures = '0'
            }
        }
        Write-Host ("  sampled {0}/{1} apps" -f ([Math]::Min($i + 20, $targets.Count)), $targets.Count)
    }
}

# ---------------------------------------------------------------------------
# Step 5: Outputs
# ---------------------------------------------------------------------------
Write-Host 'Step 5: Writing report...' -ForegroundColor Cyan

if ($rows) { $rows | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8 }

$sevColors = @{
    'Critical' = @{ fg = '#ffffff'; bg = '#b3261e' }
    'High'     = @{ fg = '#ffffff'; bg = '#b45309' }
    'Medium'   = @{ fg = '#402c00'; bg = '#fcd34d' }
    'Low'      = @{ fg = '#1e3a5f'; bg = '#dbeafe' }
    'None'     = @{ fg = '#14532d'; bg = '#dcfce7' }
    'Info'     = @{ fg = '#1e3a5f'; bg = '#dbeafe' }
}

$policyRowsHtml = ''
foreach ($t in ($triggering | Sort-Object @{e = 'Enforced'; Descending = $true }, @{e = 'SevRank'; Descending = $true })) {
    $stateBadge = if ($t.Enforced) { '<span class="badge enforced">Enabled</span>' } else { '<span class="badge reportonly">Report-only</span>' }
    $exclList = (@($t.Exclusions | ForEach-Object { EscHtml (Get-ExclusionLabel $_) }) -join '<br>')
    $policyRowsHtml += "<tr><td>$(EscHtml $t.Name)</td><td>$stateBadge</td><td>$(EscHtml $t.Reason)</td><td>$(EscHtml $t.Controls)</td><td>$(EscHtml $t.Users)</td><td>$exclList</td></tr>`n"
}
if (-not $policyRowsHtml) {
    $policyRowsHtml = '<tr><td colspan="6" class="empty">No Conditional Access policies target All resources with resource exclusions (or Windows Azure AD directly). This tenant is not affected by the enforcement change.</td></tr>'
}

$sampleCols = ''
if ($IncludeSignInSample) { $sampleCols = "<th>Sign-ins (last $SignInSampleDays d)</th><th>Users</th><th>Last seen</th><th>CA failures</th>" }

function New-AppRowHtml {
    param($r)
    $sc = $sevColors[$r.Severity]
    if (-not $sc) { $sc = $sevColors['Info'] }
    $sevBadge = "<span class='badge' style='background:$($sc.bg);color:$($sc.fg)'>$(EscHtml $r.Severity)</span>"
    $links = ''
    if ($r.AppId -and $r.SpObjectId) {
        $links = " <a href='https://entra.microsoft.com/#view/Microsoft_AAD_IAM/ManagedAppMenuBlade/~/Overview/objectId/$($r.SpObjectId)/appId/$($r.AppId)' target='_blank'>Enterprise app</a>"
        if ($r.Ownership -eq 'Tenant-owned') {
            $links += " &middot; <a href='https://entra.microsoft.com/#view/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/~/Overview/appId/$($r.AppId)' target='_blank'>App registration</a>"
        }
    }
    $flags = @()
    if ($r.ExcludedFromPolicy) { $flags += 'Excluded from policy' }
    if (-not $r.SignInEnabled) { $flags += 'Sign-in disabled' }
    $flagsHtml = ''
    if ($flags) { $flagsHtml = "<div class='flags'>$(EscHtml ($flags -join ' | '))</div>" }
    $sampleCells = ''
    if ($IncludeSignInSample) {
        # data-v carries the numeric value for client-side sorting ('100+' -> 100, blank/failed -> -1)
        $nums = foreach ($v in $r.RecentSignIns, $r.RecentUsers, $r.CaFailures) {
            $d = "$v" -replace '[^0-9]', ''
            if ($d -ne '') { [int]$d } else { -1 }
        }
        $sampleCells = "<td data-v='$($nums[0])'>$(EscHtml $r.RecentSignIns)</td><td data-v='$($nums[1])'>$(EscHtml $r.RecentUsers)</td><td>$(EscHtml $r.LastSignIn)</td><td data-v='$($nums[2])'>$(EscHtml $r.CaFailures)</td>"
    }
    $rank = $SevRankMap[$r.Severity]
    if ($null -eq $rank) { $rank = 0 }
    return "<tr><td><strong>$(EscHtml $r.App)</strong><div class='sub'>$(EscHtml $r.AppId)$links</div>$flagsHtml</td>" +
        "<td>$(EscHtml $r.ClientType)</td><td>$(EscHtml $r.Ownership)</td>" +
        "<td class='scopes'>$(EscHtml $r.ConsentedScopes)</td>" +
        "<td data-v='$rank'>$(EscHtml $r.Verdict)<br>$sevBadge</td>$sampleCells<td class='action'>$(EscHtml $r.Action)</td></tr>`n"
}

$colCount = 6
if ($IncludeSignInSample) { $colCount = 10 }
$appRowsHtml = ''
foreach ($r in $affectedRows) { $appRowsHtml += New-AppRowHtml $r }
$monitorRowsHtml = ''
foreach ($r in $monitorRows) { $monitorRowsHtml += New-AppRowHtml $r }
$possibleRowsHtml = ''
foreach ($r in $possibleRows) { $possibleRowsHtml += New-AppRowHtml $r }
if (-not $appRowsHtml) {
    $emptyMsg = 'No affected client apps were found.'
    if (-not $changeTriggers) { $emptyMsg = 'App analysis skipped: no triggering policies exist.' }
    $appRowsHtml = "<tr><td colspan='$colCount' class='empty'>$emptyMsg</td></tr>"
}
if (-not $monitorRowsHtml)  { $monitorRowsHtml = "<tr><td colspan='$colCount' class='empty'>None found.</td></tr>" }
if (-not $possibleRowsHtml) { $possibleRowsHtml = "<tr><td colspan='$colCount' class='empty'>None found.</td></tr>" }
$appTableHead = "<thead><tr><th>Application</th><th>Client type</th><th>Ownership</th><th>Consented scopes</th><th>Verdict</th>$sampleCols<th>Recommended action</th></tr></thead>"

$stateBannerHtml = ''
switch ($baselineStateKind) {
    'DefaultRollout' {
        $stateBannerHtml = "<div class='banner'><strong>Tenant state: no selection saved.</strong> $(EscHtml $baselineState)</div>"
    }
    'SelectionSaved' {
        $rawJson = ''
        try { $rawJson = ($baselineSettings | Select-Object id, modifiedDateTime, advancedSettings, exclusions | ConvertTo-Json -Depth 8) } catch { }
        $stateBannerHtml = "<div class='banner'><strong>Tenant state: a selection is saved.</strong> $(EscHtml $baselineState)<pre style='margin:8px 0 0;font-size:12px;overflow:auto'>$(EscHtml $rawJson)</pre></div>"
    }
    default {
        $stateBannerHtml = "<div class='banner'><strong>Tenant state: unknown.</strong> $(EscHtml $baselineState)</div>"
    }
}

$scanNote = 'All delegated consents (admin and per-user) were analyzed.'
if ($SkipPerUserConsents) { $scanNote = 'Only admin (AllPrincipals) consents were analyzed (-SkipPerUserConsents); per-user consented apps are not covered in this run.' }
$sampleNote = 'Sign-in sampling was not run (-IncludeSignInSample to enable); activity columns omitted.'
if ($IncludeSignInSample) { $sampleNote = "Sign-in samples cover interactive sign-ins only, last $SignInSampleDays days, capped at 100 events and $MaxAppsToSample apps." }

$html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>CA Baseline Scopes Enforcement Impact - $(EscHtml $orgName)</title>
<style>
  :root { --ink:#1a1f2b; --sub:#5b6474; --line:#e3e7ee; --card:#ffffff; --bg:#f4f6f9; --accent:#1f4e79; }
  * { box-sizing:border-box; }
  body { margin:0; font-family:'Segoe UI',system-ui,sans-serif; background:var(--bg); color:var(--ink); font-size:14px; }
  header { background:linear-gradient(120deg,#16324f,#1f4e79); color:#fff; padding:28px 36px; }
  header h1 { margin:0 0 6px; font-size:22px; font-weight:600; }
  header .meta { color:#c8d6e5; font-size:13px; }
  main { max-width:1500px; margin:0 auto; padding:24px 36px 60px; }
  .banner { background:#fff7ed; border:1px solid #fdba74; border-left:5px solid #ea580c; border-radius:8px; padding:14px 18px; margin:0 0 20px; }
  .banner strong { color:#9a3412; }
  .tiles { display:flex; gap:14px; flex-wrap:wrap; margin:0 0 24px; }
  .tile { background:var(--card); border:1px solid var(--line); border-radius:10px; padding:14px 20px; min-width:170px; }
  .tile .num { font-size:26px; font-weight:700; }
  .tile .lbl { color:var(--sub); font-size:12px; text-transform:uppercase; letter-spacing:.04em; }
  h2 { font-size:17px; margin:30px 0 10px; color:var(--accent); }
  table { width:100%; border-collapse:collapse; background:var(--card); border:1px solid var(--line); border-radius:10px; overflow:hidden; }
  th { text-align:left; background:#eef2f7; padding:9px 12px; font-size:12px; text-transform:uppercase; letter-spacing:.03em; color:var(--sub); border-bottom:1px solid var(--line); }
  td { padding:10px 12px; border-bottom:1px solid var(--line); vertical-align:top; }
  tr:last-child td { border-bottom:none; }
  .sub { color:var(--sub); font-size:12px; margin-top:2px; }
  .scopes { font-family:Consolas,monospace; font-size:12px; max-width:260px; word-break:break-word; }
  .action { max-width:420px; }
  .flags { margin-top:4px; font-size:12px; color:#9a3412; }
  .empty { text-align:center; color:var(--sub); padding:22px; }
  .badge { display:inline-block; padding:2px 9px; border-radius:999px; font-size:12px; font-weight:600; }
  .badge.enforced { background:#dcfce7; color:#14532d; }
  .badge.reportonly { background:#dbeafe; color:#1e3a5f; }
  .note { color:var(--sub); font-size:13px; margin:8px 0 0; }
  a { color:var(--accent); }
  table.sortable thead th { cursor:pointer; user-select:none; white-space:nowrap; }
  table.sortable thead th:hover { background:#e2e9f2; }
  th.s-asc::after { content:' \25B2'; font-size:9px; }
  th.s-desc::after { content:' \25BC'; font-size:9px; }
  .tfilter { margin:0 0 8px; padding:7px 12px; border:1px solid var(--line); border-radius:8px; width:300px; font:inherit; font-size:13px; background:var(--card); }
  summary { cursor:pointer; color:var(--sub); }
  section.guide { background:var(--card); border:1px solid var(--line); border-radius:10px; padding:6px 20px 16px; margin-top:26px; }
  section.guide li { margin:6px 0; }
</style>
</head>
<body>
<header>
  <h1>Conditional Access Baseline Scopes Enforcement - Impact Report</h1>
  <div class="meta">$(EscHtml $orgName) &middot; Tenant $(EscHtml $tenant) &middot; Generated $($now.ToString('yyyy-MM-dd HH:mm')) &middot; Read-only analysis</div>
</header>
<main>
<div class="banner">
  <strong>Rollout is live.</strong> Microsoft began enforcing baseline-scope evaluation for All-resources policies with exclusions on <strong>June 15, 2026</strong>, rolling out over several weeks. This tenant may already be enforced.
  Manage the behavior in <a href="https://aka.ms/BaselineScopesSettingsUX" target="_blank">Baseline scope settings</a> (US Gov: <a href="https://aka.ms/BaselineScopesSettingsUX-gov" target="_blank">gov link</a>): the blade is hidden unless reached via these direct links, which append the feature.isbaselinescopesenabled flag.
  Guidance: <a href="https://aka.ms/BaselineScopesSettings" target="_blank">aka.ms/BaselineScopesSettings</a> and the
  <a href="https://learn.microsoft.com/en-us/entra/identity/conditional-access/concept-enforcement-resource-exclusions" target="_blank">enforcement concept doc</a>.
</div>
$stateBannerHtml

<div class="tiles">
  <div class="tile"><div class="num">$($changeTriggers.Count)</div><div class="lbl">Triggering policies</div></div>
  <div class="tile"><div class="num">$($affectedRows.Count)</div><div class="lbl">Apps flagged Affected</div></div>
  <div class="tile"><div class="num">$($monitorRows.Count)</div><div class="lbl">Apps to Monitor</div></div>
  <div class="tile"><div class="num">$($possibleRows.Count)</div><div class="lbl">Unverified ISV clients</div></div>
  <div class="tile"><div class="num">$baselineOnlyCount</div><div class="lbl">Baseline-only clients</div></div>
  <div class="tile"><div class="num">$clientsScanned</div><div class="lbl">Clients scanned</div></div>
  <div class="tile"><div class="num">$(EscHtml $sevLabel)</div><div class="lbl">Max enforced control</div></div>
</div>

<h2>Policies that will evaluate baseline-scope sign-ins</h2>
<table>
  <tr><th>Policy</th><th>State</th><th>Why listed</th><th>Grant controls</th><th>User scope</th><th>Resource exclusions</th></tr>
  $policyRowsHtml
</table>
<p class="note">After enforcement, sign-ins that request only baseline scopes are evaluated against these policies with Windows Azure Active Directory as the audience. Report-only policies do not enforce yet but show what would apply.</p>

<h2>Client applications at risk ($($affectedRows.Count))</h2>
<input class="tfilter" data-target="tblAffected" type="search" placeholder="Filter apps...">
<table class="sortable" id="tblAffected">
  $appTableHead
  <tbody>
  $appRowsHtml
  </tbody>
</table>
<p class="note">$(EscHtml $scanNote) $(EscHtml $sampleNote) Confidential clients consented only to OIDC scopes were skipped as explicitly unaffected ($oidcOnlyConfidential found). Click a column header to sort; sorted by severity, then name, by default.</p>

<h2>Monitored apps ($($monitorRows.Count))</h2>
<details>
<summary>Confidential clients with baseline-directory-only footprints that are NOT excluded from any triggering policy. Their Microsoft Graph sign-ins are already CA-enforced today, so no change is expected; they matter only if someone later adds them to a policy exclusion. Expand to review.</summary>
<input class="tfilter" data-target="tblMonitor" type="search" placeholder="Filter apps..." style="margin-top:10px">
<table class="sortable" id="tblMonitor" style="margin-top:6px">
  $appTableHead
  <tbody>
  $monitorRowsHtml
  </tbody>
</table>
</details>

<h2>Unverified ISV clients ($($possibleRows.Count))</h2>
<details>
<summary>These baseline-only ISV apps are not excluded from any triggering policy and their client type cannot be determined from this tenant. Most are web SSO integrations (confidential, unaffected); any that are native or desktop clients are affected. Expand to review; sign-in activity helps separate live apps from dormant entries.</summary>
<input class="tfilter" data-target="tblPossible" type="search" placeholder="Filter apps..." style="margin-top:10px">
<table class="sortable" id="tblPossible" style="margin-top:6px">
  $appTableHead
  <tbody>
  $possibleRowsHtml
  </tbody>
</table>
</details>

<section class="guide">
<h2>How to read this and what to do</h2>
<ul>
  <li><strong>Baseline scopes:</strong> openid, profile, email, offline_access, User.Read, User.Read.All, User.ReadBasic.All, People.Read, People.Read.All, GroupMember.Read.All, Member.Read.Hidden.</li>
  <li><strong>Affected public clients</strong> break (or start prompting) regardless of policy exclusions. Interactive apps on compliant devices mostly just see a new MFA prompt; headless or kiosk usage, unmanaged devices under a compliant-device control, and anything under a block control will fail.</li>
  <li><strong>Affected confidential clients</strong> (excluded + baseline directory scopes only) have a clean fix: switch the app to OIDC scopes (openid, profile). OIDC-only confidential clients are explicitly unaffected.</li>
  <li><strong>To keep an exemption on purpose:</strong> use <em>Customize behavior</em> in the <a href="https://aka.ms/BaselineScopesSettingsUX" target="_blank">Baseline scopes settings</a>: register a placeholder single-tenant app, exclude it from the specific policy, and select it as the baseline-scopes target resource. That retains legacy behavior for that policy only. Disabling enforcement tenant-wide is not recommended.</li>
  <li><strong>Authoritative detection:</strong> this report infers "requests only baseline scopes" from consent footprints, which is static evidence. Apps that were never consented in this tenant, or whose runtime requests differ from their consents, are not visible here. Microsoft's authoritative method: configure Customize behavior with a placeholder app, then filter sign-in logs on <code>conditionalAccessAudiences</code> for that app id over several days.</li>
  <li><strong>Not in scope of this change:</strong> app-only (client credentials) tokens, apps requesting any scope beyond baseline (already enforced), and All-resources policies without exclusions.</li>
</ul>
</section>
</main>
<script>
(function () {
  function cellVal(row, idx) {
    var cell = row.cells[idx];
    if (!cell) { return ''; }
    if (cell.hasAttribute('data-v')) { return parseFloat(cell.getAttribute('data-v')); }
    var t = cell.textContent.trim();
    var n = parseFloat(t.replace(/[^0-9.\-]/g, ''));
    if (t !== '' && !isNaN(n) && /^[0-9]/.test(t)) { return n; }
    return t.toLowerCase();
  }
  document.querySelectorAll('table.sortable thead th').forEach(function (th) {
    th.addEventListener('click', function () {
      var table = th.closest('table');
      var tbody = table.tBodies[0];
      var idx = Array.prototype.indexOf.call(th.parentNode.children, th);
      var dir = th.dataset.dir === 'desc' ? 'asc' : 'desc';
      table.querySelectorAll('thead th').forEach(function (h) { delete h.dataset.dir; h.classList.remove('s-asc', 's-desc'); });
      th.dataset.dir = dir;
      th.classList.add(dir === 'asc' ? 's-asc' : 's-desc');
      var rows = Array.prototype.slice.call(tbody.rows);
      rows.sort(function (a, b) {
        var va = cellVal(a, idx), vb = cellVal(b, idx);
        if (typeof va === 'number' && typeof vb === 'number') { return dir === 'asc' ? va - vb : vb - va; }
        va = String(va); vb = String(vb);
        return dir === 'asc' ? va.localeCompare(vb) : vb.localeCompare(va);
      });
      rows.forEach(function (r) { tbody.appendChild(r); });
    });
  });
  document.querySelectorAll('input.tfilter').forEach(function (inp) {
    inp.addEventListener('input', function () {
      var table = document.getElementById(inp.dataset.target);
      if (!table) { return; }
      var q = inp.value.toLowerCase();
      Array.prototype.forEach.call(table.tBodies[0].rows, function (r) {
        r.style.display = r.textContent.toLowerCase().indexOf(q) === -1 ? 'none' : '';
      });
    });
  });
})();
</script>
</body>
</html>
"@

Set-Content -Path $htmlPath -Value $html -Encoding UTF8

Write-Host ''
Write-Host '=========================== SUMMARY ===========================' -ForegroundColor Cyan
Write-Host "Baseline scopes tenant state: $baselineStateKind"
Write-Host ("Triggering policies (All resources + exclusions): {0} ({1} enforced); WAAD-targeting policies listed: {2}" -f $changeTriggers.Count, @($changeTriggers | Where-Object Enforced).Count, (@($triggering).Count - $changeTriggers.Count))
Write-Host ("Clients scanned: {0}; baseline-only footprints: {1}" -f $clientsScanned, $baselineOnlyCount)
Write-Host ("Apps flagged AFFECTED: {0}" -f $affectedRows.Count) -ForegroundColor $(if ($affectedRows.Count) { 'Yellow' } else { 'Green' })
Write-Host ("Apps to monitor: {0}; unverified ISV clients (possible): {1}" -f $monitorRows.Count, $possibleRows.Count)
Write-Host ''
Write-Host "Report: $htmlPath"
if ($rows) { Write-Host "CSV:    $csvPath" }
