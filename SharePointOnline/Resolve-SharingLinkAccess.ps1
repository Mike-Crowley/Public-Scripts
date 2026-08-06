#Requires -Version 5.1
#Requires -Modules Microsoft.Graph.Authentication

<#
.SYNOPSIS
    Resolves a SharePoint/OneDrive sharing link to the item it points at and reports
    every principal who can reach it, without accepting the share.

.DESCRIPTION
    Answers the question an admin gets during a data-incident review: "someone sent me
    this link, who has access to it?"

    Clicking a sharing link starts the redemption flow and can add the investigator to
    the item's ACL, contaminating the very thing being investigated. This script resolves
    the link through the Microsoft Graph /shares endpoint instead, which only redeems when
    an explicit Prefer header is sent, and by default none is sent at all.

    The trade-off is real and worth stating: without redeeming, a caller who has no
    existing access to the item gets a 403 rather than an answer. The routes that resolve
    that without touching the ACL are an app-only certificate sign-in, which is treated as
    an owner, or the item's plain path URL, which needs no token. -RedeemIfNecessary is
    available as a last resort and is recorded in the report when used.

    From the resolved item it reports:
      - Direct grants (grantedToV2) and sharing-link grants (grantedToIdentitiesV2).
        Reading only the singular property is the most common bug in hand-rolled versions
        of this: the invitees on a "specific people" link live in the PLURAL collection,
        so they silently render blank - exactly the population an investigation needs.
      - Decoded SharePoint claim login names, so 'c:0o.c|federateddirectoryclaimprovider|<guid>'
        reads as "Microsoft 365 group" rather than an opaque string.
      - Expanded Entra group membership, so a "Can edit" row that is really a 400-person
        group is reported as 400 people (-ExpandGroups).
      - Anonymous ("Anyone"), organization-wide, and guest access called out by risk.
      - An inherited-vs-unique inference by walking ancestors, because SharePoint document
        libraries do not return the inheritedFrom property that Graph documents elsewhere.
      - Optional unified audit log pivot: who created the link, who was added to it, and
        who actually redeemed or used it (-IncludeAuditLog).

    READ-ONLY: GETs only, plus the Search-UnifiedAuditLog read cmdlet when -IncludeAuditLog
    is used. Nothing is created, changed, deleted, or redeemed.

.PARAMETER Url
    One or more sharing URLs (the https://tenant.sharepoint.com/:f:/s/... form) or plain
    item URLs (the https://tenant.sharepoint.com/sites/team/Shared%20Documents/... form).
    Accepts pipeline input. Sharing URLs resolve through /shares; plain URLs fall back to
    a site/drive/path walk if /shares declines them.

.PARAMETER TenantId
    Usually unnecessary. The tenant is derived from the link host by default:
    contoso.sharepoint.com maps to contoso.onmicrosoft.com, and Connect-MgGraph -TenantId
    accepts any verified domain of a tenant, so that value is passed straight through and
    Graph validates it. Pinning matters most for consultants holding guest access elsewhere,
    where an unpinned sign-in can authenticate against the wrong directory and return a bare
    404 that reads like a deleted link.

    Supply this only to override the derivation, or when it fails because the tenant uses a
    vanity SharePoint domain. An explicit value always wins and is not second-guessed: a
    tenant's vanity domain and its onmicrosoft domain are different strings for the same
    tenant, and nothing available locally tells that apart from a real mismatch. Accepts a
    GUID or any verified domain, though -IncludeAuditLog additionally needs a domain
    because Connect-ExchangeOnline -Organization rejects GUIDs.

.PARAMETER ClientId
    Entra application (client) id for app-only certificate sign-in. Use with
    -CertificateThumbprint and -TenantId. Preferred for incident work: an application
    identity is treated as an owner by the permissions endpoint, so it sees the COMPLETE
    ACL, and it never appears in the ACL itself. See the -Scopes notes below.

.PARAMETER CertificateThumbprint
    Thumbprint of the certificate registered on the app in -ClientId. The certificate must
    be in the current user's or machine's certificate store.
    https://learn.microsoft.com/en-us/entra/identity-platform/howto-create-self-signed-certificate

.PARAMETER UseDeviceCode
    Device-code sign-in for hosts where interactive/WAM auth fails (VS Code integrated
    console, ISE, ssh sessions).

.PARAMETER Scopes
    Override the delegated scopes requested. The default set is read-only. Microsoft
    documents /shares as requiring Files.ReadWrite (delegated) or Files.ReadWrite.All /
    Sites.ReadWrite.All (application), which reads as a documentation artifact of the
    redemption capability rather than a real write requirement: read-only scopes work in
    practice. If /shares returns 403, rerun with -Scopes to add the documented write scope,
    or pass the translated plain item URL instead, which uses a different path entirely.

.PARAMETER Environment
    Graph national cloud. Also derived from the link host by default and rarely needs
    supplying: .sharepoint.com is Global (which covers commercial and GCC), .sharepoint.us
    is GCC High, .sharepoint-mil.us is DoD, and .sharepoint.cn is 21Vianet China. An
    explicit value always wins. Note that -IncludeAuditLog works in every cloud because it
    uses Search-UnifiedAuditLog; the Graph audit search API alternative is Global-only.

.PARAMETER MaxAncestorDepth
    How many parent levels to walk when inferring inherited permissions. Default 10, which
    reaches the library root of all but the most deeply nested folder structures.

.PARAMETER ExpandGroups
    Expand Entra ID (security and Microsoft 365) groups found in the ACL to their
    transitive members. SharePoint groups cannot be expanded through Graph; those rows
    carry a direct link to the site's group membership page instead.

.PARAMETER MaxGroupMembers
    Cap on members listed per expanded group. Default 200. Truncation is reported.

.PARAMETER IncludeChildren
    Also check the folder's children for permissions that differ from the folder itself,
    which surfaces broken inheritance underneath the shared folder.

.PARAMETER MaxChildren
    Cap on children inspected when -IncludeChildren is used. Default 50.

.PARAMETER SkipInheritanceCheck
    Skip the ancestor walk. The walk costs one permissions call per ancestor level and is
    what distinguishes "this folder was shared" from "this folder inherited a share made
    higher up", which is usually the actual finding.

.PARAMETER IncludeAuditLog
    Pivot into the unified audit log for the resolved item: link creation, invitees,
    redemptions, and revocations. Requires the ExchangeOnlineManagement module and the
    View-Only Audit Logs role. This is the only way to answer "who actually opened it",
    which permissions alone cannot answer.

.PARAMETER AuditDays
    Audit log lookback in days. Default 180. Retention is 180 days on E3 and 365 on E5.

.PARAMETER IncludeFileAccess
    Add FileAccessed / FileDownloaded / FilePreviewed events to the audit pivot. Noisy on
    busy items, but it is the record of who actually read the content.

.PARAMETER RedeemIfNecessary
    Send Prefer: redeemSharingLinkIfNecessary when resolving the link. Off by default.

    Without it, a bare GET on /shares grants nothing, which is the point. The cost is that a
    caller with no existing access to the item gets a 403, and that is exactly the position
    of an investigator who deliberately did not accept the share. Microsoft documents this
    header as being for the case where "the intention is simply to peek at the link's
    metadata", with access "only guaranteed to be granted for the duration of this request".

    Read that hedge carefully: it promises the access is temporary, not that no permission
    record is written. In a data-incident review, where the ACL is the evidence, prefer the
    routes that cannot alter it at all: an app-only certificate sign-in (-ClientId and
    -CertificateThumbprint) is treated as an owner and sees the complete ACL without ever
    appearing in it, and a plain item path URL resolves without a token. Use this switch
    when neither is available, and note that the report records that it was used.

    The durable form, Prefer: redeemSharingLink, is equivalent to clicking the link in a
    browser and is never sent by this script under any switch.

.PARAMETER GridView
    Browse the results in Out-GridView, which gives sorting and filtering interactively and
    is usually the fastest way to triage a large or batched result. Falls back to a wrapped
    console table where Out-GridView is unavailable (any non-Windows host). When
    -IncludeAuditLog also ran, the sharing timeline opens in a second grid.

.PARAMETER OutputFolder
    Defaults to Desktop\SharingLinkAccess.

.PARAMETER NoOpen
    Do not open the HTML report when the run finishes.

.EXAMPLE
    .\Resolve-SharingLinkAccess.ps1 -Url 'https://contoso.sharepoint.com/:f:/s/team/Eq0ibm9uc1pJvLmeWUEZm_4B...'

    Resolves the link interactively, reports every principal with access, and writes an
    HTML report plus CSV. The link is never redeemed, so the operator does not appear in
    the results of their own investigation.

.EXAMPLE
    .\Resolve-SharingLinkAccess.ps1 -Url $link -ClientId <app-guid> -CertificateThumbprint <thumbprint> -ExpandGroups -IncludeAuditLog

    App-only run for a formal investigation: complete ACL (application identities are
    treated as owners), groups flattened to named people, and the sharing timeline from
    the audit log.

.EXAMPLE
    Get-Content .\links.txt | .\Resolve-SharingLinkAccess.ps1 -ExpandGroups

    Batch mode. One report covering every link.

.INPUTS
    System.String. Sharing URLs or item URLs, by property name or by value.

.OUTPUTS
    [pscustomobject] rows (one per principal per permission) to the pipeline, plus
    Desktop\SharingLinkAccess\SharingLinkAccess_<item>_<yyyyMMdd_HHmmss>.{html,csv}

.NOTES
    Author:  Mike Crowley
    https://mikecrowley.us

    Two Graph behaviours dominate the accuracy of any tool in this space:

    1. Partial ACLs for non-owners. Per the driveItem permissions reference: "For the
       owner of the item, all sharing permissions will be returned... For a non-owner
       caller, only the sharing permissions that apply to the caller are returned."
       A delegated run as a non-owner therefore returns a SHORT list with no error and no
       warning. This script detects that case and banners it. App-only, site collection
       admin, or the item owner are the only ways to see everything.

    2. shareId and the link webUrl are "only returned for callers that are able to create
       the sharing permission", so a low-privilege delegated run also loses the correlation
       key the audit pivot uses.

    Known platform limits this script works around or reports rather than hides:
      - OneDrive for Business and SharePoint document libraries do not return inheritedFrom,
        so inheritance here is INFERRED by comparing ancestor ACLs and is labelled as such.
      - hasPassword is documented as OneDrive Personal only, so it is not trustworthy for
        SharePoint links.
      - Graph exposes SharePoint groups as an id plus login name and cannot enumerate their
        members; only Entra groups can be expanded.
      - "Specific people" access is physically stored in hidden SharePoint groups named
        SharingLinks.{UniqueId}.{SharingLinkKind}.{ShareId}. Those are not visible in the
        UI, which is why the browser "Manage access" pane and this report can disagree.

    Alternatives considered, and why this exists:
      - The SharePoint admin "Shared with external users" site usage report covers a whole
        site and requires site admin rights on it, which an investigator often should not
        take on the site under review.
      - Data Access Governance reports rank sites by sharing volume and require SharePoint
        Premium / Advanced Management licensing. They do not resolve a single link.
      - Get-PnPFolderSharingLink and Get-PnPFileSharingLink wrap the same Graph endpoint but
        start from a known server-relative path, not from a sharing link, and filter to link
        permissions only.
      - The legacy SharePoint REST GetSharingLinkData / GetFileByGuestUrl route is
        unsupported for user-supplied links; Microsoft's guidance is the Graph shares API.

    The tenant is derived from the link host by string mapping only. No tenant lookup is
    performed here: Graph validates the domain at sign-in. If you want the tenant GUID,
    region scope, or federation state for a domain, that is Get-EntraCredentialInfo.ps1 in
    this repo, which already queries the unauthenticated OpenID configuration endpoint.

.LINK
    https://learn.microsoft.com/en-us/graph/api/shares-get

.LINK
    https://learn.microsoft.com/en-us/graph/api/driveitem-list-permissions

.LINK
    https://learn.microsoft.com/en-us/graph/api/resources/permission

.LINK
    https://learn.microsoft.com/en-us/purview/audit-log-sharing
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName, Position = 0)]
    [Alias('SharingUrl', 'SharingLink', 'WebUrl')]
    [string[]]$Url,

    [string]$TenantId,

    [string]$ClientId,
    [string]$CertificateThumbprint,

    [switch]$UseDeviceCode,

    [string[]]$Scopes = @('Sites.Read.All', 'Files.Read.All', 'Group.Read.All', 'User.Read.All'),

    [ValidateSet('Global', 'USGov', 'USGovDoD', 'China')]
    [string]$Environment = 'Global',

    [switch]$ExpandGroups,
    [ValidateRange(1, 5000)][int]$MaxGroupMembers = 200,

    [switch]$IncludeChildren,
    [ValidateRange(1, 500)][int]$MaxChildren = 50,

    [switch]$SkipInheritanceCheck,
    [ValidateRange(1, 30)][int]$MaxAncestorDepth = 10,

    [switch]$IncludeAuditLog,
    [ValidateRange(1, 365)][int]$AuditDays = 180,
    [switch]$IncludeFileAccess,

    # Opt in to Prefer: redeemSharingLinkIfNecessary. Off by default; see .PARAMETER.
    [switch]$RedeemIfNecessary,

    [switch]$GridView,

    [string]$OutputFolder,
    [switch]$NoOpen
)

begin {
    $ErrorActionPreference = 'Stop'
    $CollectedUrls = [System.Collections.Generic.List[string]]::new()

    # -----------------------------------------------------------------------
    # Constants
    # -----------------------------------------------------------------------
    $graphBase = 'https://graph.microsoft.com'   # replaced after connect for national clouds

    # Sharing-link path markers SharePoint uses between the host and the token:
    # :f: folder, :w: Word, :x: Excel, :p: PowerPoint, :b: PDF/other binary, :o: OneNote,
    # :v: video, :i: image, :t: text, :u: generic, :l: list, :g: page. Matching the shape
    # rather than the letter keeps new markers working.
    $SharingUrlPattern = '/:[A-Za-z]:/[A-Za-z]/'

    # SharePoint sharing operations, per the Purview sharing-audit schema.
    $SharingOperations = @(
        'SharingSet', 'SharingRevoked', 'SharingInvitationCreated', 'SharingInvitationAccepted',
        'SharingInvitationBlocked', 'SharingInvitationRevoked',
        'SecureLinkCreated', 'SecureLinkUsed', 'SecureLinkUpdated', 'SecureLinkDeleted',
        'AddedToSecureLink', 'RemovedFromSecureLink',
        'AnonymousLinkCreated', 'AnonymousLinkUsed', 'AnonymousLinkUpdated', 'AnonymousLinkRemoved',
        'CompanyLinkCreated', 'CompanyLinkUsed', 'CompanyLinkRemoved',
        'AddedToGroup', 'RemovedFromGroup',
        'SharingInheritanceBroken', 'SharingInheritanceReset'
    )
    $FileOperations = @('FileAccessed', 'FileDownloaded', 'FilePreviewed', 'FileAccessedExtended')

    function EscHtml { param($s) [System.Net.WebUtility]::HtmlEncode("$s") }

    function Get-TenantHintFromUrl {
        <#
            A sharing URL already names its tenant: contoso.sharepoint.com belongs to
            contoso.onmicrosoft.com, because the SharePoint hostname prefix is minted from
            the tenant's initial onmicrosoft prefix and stays in lockstep with it.

            Connect-MgGraph -TenantId accepts any verified domain of a tenant, so that
            derived domain is handed straight to it. No tenant lookup of our own: Graph
            validates it, and a wrong value produces a real sign-in error rather than a
            silent fallback.

            Pinning matters most for consultants holding guest access in several tenants,
            where an unpinned sign-in can authenticate against the wrong directory and
            return a bare 404 that reads like a deleted link.

            The host suffix also identifies the cloud, which saves supplying -Environment.
            Vanity SharePoint domains break the prefix rule and are reported as such.
        #>
        param([string]$Value)

        $out = [pscustomobject]@{ HostName = ''; Domain = ''; Environment = ''; Note = '' }
        $uri = $null
        try { $uri = [uri]$Value } catch { $out.Note = 'Not a parsable URL.'; return $out }
        $h = "$($uri.Host)".ToLowerInvariant()
        $out.HostName = $h

        # Longest suffix first: .sharepoint-mil.us must win over .sharepoint.us.
        $map = @(
            @{ Suffix = '.sharepoint-mil.us'; Domain = '.onmicrosoft.us'; Env = 'USGovDoD' }
            @{ Suffix = '.sharepoint.us'; Domain = '.onmicrosoft.us'; Env = 'USGov' }
            @{ Suffix = '.sharepoint.cn'; Domain = '.partner.onmschina.cn'; Env = 'China' }
            @{ Suffix = '.sharepoint.com'; Domain = '.onmicrosoft.com'; Env = 'Global' }
        )
        $m = $map | Where-Object { $h.EndsWith($_.Suffix) } | Select-Object -First 1
        if (-not $m) {
            $out.Note = "Host '$h' is not a recognized SharePoint hostname; a vanity SharePoint domain does this."
            return $out
        }

        # contoso-my (OneDrive) and contoso-admin (admin center) share the tenant prefix.
        $prefix = $h.Substring(0, $h.Length - $m.Suffix.Length) -replace '-(my|admin)$', ''
        if (-not $prefix) { $out.Note = 'Could not extract a tenant prefix from the hostname.'; return $out }

        $out.Domain = "$prefix$($m.Domain)"
        $out.Environment = $m.Env
        return $out
    }

    function ConvertTo-GraphShareToken {
        # https://learn.microsoft.com/en-us/graph/api/shares-get#encoding-sharing-urls
        # base64, strip '=' padding, '/' -> '_', '+' -> '-', prefix 'u!'.
        param([string]$Value)
        $b64 = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($Value))
        return 'u!' + $b64.TrimEnd('=').Replace('/', '_').Replace('+', '-')
    }

    function Get-GraphPages {
        # Emits each record individually. Deliberately NOT "return , $array": that idiom
        # hands the whole array back as ONE object, so a caller wrapping the result in @()
        # gets a single nested element instead of N records. Member access on that nested
        # array then silently MERGES every record (two ACEs become one with roles
        # 'write, owner' and both ids joined), which is worse than an error because it
        # looks like a plausible answer.
        param([string]$Uri)
        $items = [System.Collections.Generic.List[object]]::new()
        $next = $Uri
        while ($next) {
            $resp = Invoke-MgGraphRequest -Method GET -Uri $next -OutputType PSObject
            if ($resp.value) { foreach ($v in $resp.value) { $items.Add($v) } }
            $next = $resp.'@odata.nextLink'
        }
        return $items.ToArray()
    }

    function Invoke-GraphSafe {
        # Separates "not there / not allowed" from transport failures so callers can fall
        # back on a 403/404 without swallowing throttling or auth breakage.
        param([string]$Uri, [hashtable]$Headers)
        try {
            $splat = @{ Method = 'GET'; Uri = $Uri; OutputType = 'PSObject'; ErrorAction = 'Stop' }
            if ($Headers -and $Headers.Count) { $splat.Headers = $Headers }
            return @{ Ok = $true; Data = (Invoke-MgGraphRequest @splat) }
        }
        catch {
            $msg = "$($_.Exception.Message)"
            $status = 0
            try { $status = [int]$_.Exception.StatusCode } catch { }
            if (-not $status) { try { $status = [int]$_.Exception.Response.StatusCode } catch { } }
            if (-not $status) {
                if ($msg -match '\b403\b|[Ff]orbidden|accessDenied') { $status = 403 }
                elseif ($msg -match '\b404\b|[Nn]ot ?[Ff]ound|itemNotFound') { $status = 404 }
            }
            return @{ Ok = $false; Status = $status; Error = $msg }
        }
    }

    function ConvertFrom-SharePointLoginName {
        <#
            SharePoint stores principals as claims strings. Decoding them is what turns an
            unreadable ACL row into an answer. Encodings covered here are the ones that
            actually appear in SharePoint Online ACLs; anything else falls through to
            'Unknown' with the raw value preserved rather than being guessed at.
        #>
        param([string]$LoginName)

        $result = [pscustomobject]@{
            Type       = 'Unknown'
            Identifier = "$LoginName"
            DirectoryId = ''
            IsBroad    = $false
            Note       = ''
        }
        if (-not $LoginName) { return $result }

        $parts = "$LoginName" -split '\|'
        $prefix = $parts[0]
        $last = $parts[-1]

        switch -Regex ("$LoginName") {
            '^i:0#\.f\|membership\|' {
                $result.Type = 'Entra user'
                $result.Identifier = $last
                if ($last -match '#EXT#') { $result.Type = 'Guest user' }
                return $result
            }
            '^i:0#\.w\|' {
                $result.Type = 'On-premises AD user'
                $result.Identifier = $last
                return $result
            }
            '^i:0[e5]\.t\|' {
                $result.Type = 'Federated user'
                $result.Identifier = $last
                return $result
            }
            '^c:0\(\.s\|true$' {
                $result.Type = 'Everyone'
                $result.Identifier = 'Everyone (includes external users)'
                $result.IsBroad = $true
                $result.Note = 'Every authenticated user plus previously invited external users.'
                return $result
            }
            '^c:0!\.s\|windows$' {
                $result.Type = 'All authenticated users'
                $result.Identifier = 'NT AUTHORITY\authenticated users'
                $result.IsBroad = $true
                return $result
            }
            '^c:0-\.f\|rolemanager\|spo-grid-all-users' {
                $result.Type = 'Everyone except external users'
                $result.Identifier = 'Everyone except external users'
                $result.IsBroad = $true
                $result.Note = 'Every member of the tenant directory.'
                return $result
            }
            '^c:0o\.c\|federateddirectoryclaimprovider\|' {
                $result.Type = 'Microsoft 365 group'
                $id = $last
                if ($id -match '_o$') {
                    $result.Type = 'Microsoft 365 group (owners)'
                    $id = $id -replace '_o$', ''
                }
                $result.DirectoryId = $id
                $result.Identifier = $id
                return $result
            }
            '^c:0t\.c\|tenant\|' {
                $result.Type = 'Entra security group'
                $result.DirectoryId = $last
                $result.Identifier = $last
                return $result
            }
            '^c:0u\.c\|tenant\|' {
                $result.Type = 'Entra directory object'
                $result.DirectoryId = $last
                $result.Identifier = $last
                return $result
            }
            '^c:0h\.f\|rolemanager\|' {
                $result.Type = 'Directory role'
                $result.Identifier = $last
                $result.IsBroad = $true
                return $result
            }
            '^i:0#\.f\|rolemanager\|' {
                $result.Type = 'Role'
                $result.Identifier = $last
                return $result
            }
            '^SharingLinks\.' {
                # SharingLinks.{UniqueId}.{SharingLinkKind}.{ShareId}
                $result.Type = 'Sharing link group'
                $bits = "$LoginName" -split '\.'
                if ($bits.Count -ge 4) { $result.Note = "Link kind: $($bits[2])" }
                $result.Identifier = "$LoginName"
                return $result
            }
            '^c:0\(\.s\|' {
                $result.Type = 'Claim'
                $result.Identifier = $last
                return $result
            }
        }

        # Unrecognised claim: keep the last segment, which is the useful half of every
        # encoding seen so far, but do not assert a type.
        if ($parts.Count -gt 1) { $result.Identifier = $last; $result.Note = "Unrecognised claim prefix '$prefix'." }
        return $result
    }

    function Get-AccessLabel {
        param([string[]]$Roles)
        $r = @($Roles | Where-Object { $_ })
        if (-not $r) { return '(none)' }
        $labels = foreach ($x in $r) {
            switch -Regex ("$x") {
                '^owner$'          { 'Full control' }
                '^write$'          { 'Can edit' }
                '^read$'           { 'Can view' }
                '^review$'         { 'Can review' }
                '^restricted'      { 'Restricted view' }
                '^sp\.full ?control' { 'Full control' }
                default            { "$x" }
            }
        }
        return (($labels | Select-Object -Unique) -join ', ')
    }

    function Get-LinkDescription {
        param($Link)
        if (-not $Link) { return '' }
        $scope = "$($Link.scope)"
        $type = "$($Link.type)"
        $scopeText = switch ($scope) {
            'anonymous'      { 'Anyone with the link' }
            'organization'   { 'Anyone in the organization' }
            'users'          { 'Specific people' }
            'existingAccess' { 'People with existing access' }
            ''               { 'Sharing link' }
            default          { $scope }
        }
        $typeText = switch ($type) {
            'view'           { 'view' }
            'edit'           { 'edit' }
            'review'         { 'review' }
            'embed'          { 'embed' }
            'blocksDownload' { 'view, no download' }
            'createOnly'     { 'upload only' }
            ''               { '' }
            default          { $type }
        }
        if ($typeText) { return "$scopeText ($typeText)" }
        return $scopeText
    }
}

process {
    foreach ($u in $Url) { if ("$u".Trim()) { $CollectedUrls.Add("$u".Trim()) } }
}

end {
    $targets = @($CollectedUrls | Select-Object -Unique)
    if (-not $targets) { throw 'No URLs supplied.' }

    # -----------------------------------------------------------------------
    # Tenant pinning, derived from the link host unless overridden
    # -----------------------------------------------------------------------
    $tenantHints = @()
    foreach ($t in $targets) { $tenantHints += Get-TenantHintFromUrl -Value $t }
    $distinctTenants = @($tenantHints | Where-Object { $_.Domain } | Select-Object -ExpandProperty Domain -Unique)
    if ($distinctTenants.Count -gt 1) {
        # A warning rather than a throw: a tenant that renamed its SharePoint domain serves
        # both the old and new hostnames, so two hostnames are not proof of two tenants and
        # blocking here would reject a legitimate old-link plus new-link comparison. If they
        # really are different tenants, the sign-in serves only one and the others fail
        # individually with their own message, which is more useful than refusing to start.
        Write-Warning ("These URLs carry $($distinctTenants.Count) different SharePoint hostnames ($($distinctTenants -join ', ')). " +
            "Signing in once, against $($distinctTenants[0]). That is fine when a tenant has renamed its domain, since both " +
            'hostnames belong to it. If these really are separate tenants, run the script once per tenant.')
    }
    # Prefer a hint that produced a domain, so a vanity-domain URL still yields an object to
    # warn from instead of failing silently.
    $derived = @($tenantHints | Where-Object { $_.Domain })[0]
    if (-not $derived) { $derived = @($tenantHints)[0] }

    # Connect-ExchangeOnline -Organization requires a domain and rejects a GUID, so it is
    # tracked separately from -TenantId, which accepts either.
    $ExoOrganization = "$($derived.Domain)"
    if ($PSBoundParameters.ContainsKey('TenantId') -and $TenantId) {
        # An explicit value wins and is never second-guessed. A tenant's vanity domain and
        # its onmicrosoft domain are different strings for the same tenant, and nothing
        # available here distinguishes that from a genuine mismatch, so warning on a string
        # difference would fire on correct usage.
        if ($TenantId -notmatch '^[0-9a-f]{8}-[0-9a-f]{4}-') { $ExoOrganization = $TenantId }
    }
    elseif ($derived.Domain) {
        # Connect-MgGraph -TenantId takes any verified domain of the tenant, so the derived
        # onmicrosoft domain goes straight in and Graph validates it.
        $TenantId = $derived.Domain
        Write-Host "Pinning sign-in to $($derived.Domain), derived from the link host $($derived.HostName)." -ForegroundColor Green
    }
    else {
        Write-Warning ("Could not derive a tenant from the link host. $($derived.Note) " +
            'Signing in without pinning a tenant; pass -TenantId if the sign-in lands in the wrong directory.')
    }

    # The link host also identifies the cloud, so -Environment only needs supplying when
    # the caller is overriding that inference.
    if (-not $PSBoundParameters.ContainsKey('Environment') -and $derived.Environment -and $Environment -ne $derived.Environment) {
        Write-Host "Graph environment set to $($derived.Environment) from the link host." -ForegroundColor Green
        $Environment = $derived.Environment
    }

    # -----------------------------------------------------------------------
    # Connect
    # -----------------------------------------------------------------------
    $connect = @{ NoWelcome = $true; ContextScope = 'Process'; Environment = $Environment }
    $AppOnly = $false
    if ($ClientId -or $CertificateThumbprint) {
        if (-not ($ClientId -and $CertificateThumbprint)) {
            throw 'App-only sign-in needs -ClientId and -CertificateThumbprint together.'
        }
        if (-not $TenantId) {
            throw ('App-only sign-in needs a tenant, and one could not be derived from the link host. ' +
                "$(if ($derived) { $derived.Note }) Pass -TenantId explicitly.")
        }
        $connect.ClientId = $ClientId
        $connect.CertificateThumbprint = $CertificateThumbprint
        $connect.TenantId = $TenantId
        $AppOnly = $true
        Write-Host 'Signing in to Microsoft Graph (app-only certificate)...' -ForegroundColor Cyan
    }
    else {
        $connect.Scopes = $Scopes
        if ($TenantId) { $connect.TenantId = $TenantId }
        if ($UseDeviceCode) { $connect.UseDeviceCode = $true }
        Write-Host 'Signing in to Microsoft Graph (delegated, read-only scopes)...' -ForegroundColor Cyan
    }
    Connect-MgGraph @connect

    $ctx = Get-MgContext
    if (-not $ctx -or -not $ctx.TenantId) { throw 'Graph sign-in did not produce a usable context.' }
    try {
        $envInfo = Get-MgEnvironment -Name $ctx.Environment -ErrorAction Stop
        if ($envInfo.GraphEndpoint) { $graphBase = "$($envInfo.GraphEndpoint)".TrimEnd('/') }
    }
    catch { }

    $orgName = ''
    try {
        $org = Invoke-MgGraphRequest -Method GET -Uri "$graphBase/v1.0/organization?`$select=displayName,id" -OutputType PSObject
        $orgName = "$(@($org.value)[0].displayName)"
    }
    catch { }
    if (-not $orgName) { $orgName = "$($ctx.TenantId)" }
    $signedInAs = if ($AppOnly) { "app $($ctx.ClientId)" } else { "$($ctx.Account)" }
    Write-Host "Connected: $orgName as $signedInAs" -ForegroundColor Green

    if (-not $AppOnly) {
        Write-Warning ('Delegated sign-in. Graph returns the COMPLETE permission list only to an owner or co-owner of ' +
            'the item; every other caller silently receives just the permissions that apply to them. If this account ' +
            'is not an owner or site collection admin of the target, the results below are incomplete with no error. ' +
            'Use -ClientId/-CertificateThumbprint for an authoritative answer.')
    }

    # -----------------------------------------------------------------------
    # Item resolution
    # -----------------------------------------------------------------------
    function Resolve-ItemFromShareToken {
        # A bare GET does NOT redeem: redemption requires an explicit Prefer header, and by
        # default none is sent. That is the whole reason to resolve links here rather than
        # in a browser.
        #
        # The cost of sending nothing is that a caller with no existing access to the item
        # gets a 403, which is precisely the investigator's position when they have
        # deliberately not accepted the share. -RedeemIfNecessary opts in to
        # redeemSharingLinkIfNecessary, which Microsoft documents as the option for peeking
        # at a link's metadata, with access "only guaranteed to be granted for the duration
        # of this request". Note the hedge in that wording: it promises the access is
        # temporary, not that no permission record is written. Left off by default for that
        # reason. redeemSharingLink, the durable form equivalent to clicking the link, is
        # never sent by this script under any switch.
        param([string]$Value, [switch]$Redeem)
        $token = ConvertTo-GraphShareToken -Value $Value
        $headers = $null
        if ($Redeem) { $headers = @{ Prefer = 'redeemSharingLinkIfNecessary' } }
        return Invoke-GraphSafe -Uri "$graphBase/v1.0/shares/$token/driveItem?`$select=id,name,webUrl,size,folder,file,createdBy,createdDateTime,lastModifiedBy,lastModifiedDateTime,parentReference,sharepointIds" -Headers $headers
    }

    function Resolve-ItemFromPathUrl {
        <#
            Fallback for a plain item URL: find the site by trying progressively shorter
            server-relative paths (handles /sites/x, /teams/x, /personal/x, subsites, and
            the root site), then match the drive whose webUrl prefixes the target, then
            address the item by path. Uses only read scopes, so it also works when /shares
            is refused.
        #>
        param([string]$Value)

        $uri = $null
        try { $uri = [uri]$Value } catch { return @{ Ok = $false; Status = 0; Error = "Not a parsable URL: $Value" } }
        $host_ = $uri.Host
        $segments = @(($uri.AbsolutePath.Trim('/') -split '/') | Where-Object { $_ })
        if (-not $segments) { return @{ Ok = $false; Status = 0; Error = 'URL has no path.' } }

        $site = $null
        for ($take = [Math]::Min(4, $segments.Count); $take -ge 0; $take--) {
            $candidate = ''
            if ($take -gt 0) { $candidate = '/' + (($segments[0..($take - 1)] | ForEach-Object { [uri]::UnescapeDataString($_) }) -join '/') }
            $siteUri = if ($candidate) { "$graphBase/v1.0/sites/$($host_):$($candidate):?`$select=id,webUrl,displayName" }
                       else { "$graphBase/v1.0/sites/$($host_)?`$select=id,webUrl,displayName" }
            $r = Invoke-GraphSafe $siteUri
            if ($r.Ok -and $r.Data.id) { $site = $r.Data; break }
        }
        if (-not $site) { return @{ Ok = $false; Status = 404; Error = "Could not resolve a site from $Value" } }

        $drives = Invoke-GraphSafe "$graphBase/v1.0/sites/$($site.id)/drives?`$select=id,name,webUrl"
        if (-not $drives.Ok) { return @{ Ok = $false; Status = $drives.Status; Error = "Site resolved but drives could not be listed: $($drives.Error)" } }

        $targetDecoded = [uri]::UnescapeDataString("$Value").TrimEnd('/')
        $best = $null
        foreach ($d in @($drives.Data.value)) {
            $dw = [uri]::UnescapeDataString("$($d.webUrl)").TrimEnd('/')
            if ($targetDecoded.StartsWith($dw, [StringComparison]::OrdinalIgnoreCase)) {
                if (-not $best -or $dw.Length -gt ([uri]::UnescapeDataString("$($best.webUrl)").TrimEnd('/')).Length) { $best = $d }
            }
        }
        if (-not $best) { return @{ Ok = $false; Status = 404; Error = "No document library in $($site.webUrl) matches $Value" } }

        $driveWeb = [uri]::UnescapeDataString("$($best.webUrl)").TrimEnd('/')
        $rel = $targetDecoded.Substring($driveWeb.Length).Trim('/')
        $selectClause = "`$select=id,name,webUrl,size,folder,file,createdBy,createdDateTime,lastModifiedBy,lastModifiedDateTime,parentReference,sharepointIds"
        $itemUri = if ($rel) {
            $encRel = (($rel -split '/') | ForEach-Object { [uri]::EscapeDataString($_) }) -join '/'
            "$graphBase/v1.0/drives/$($best.id)/root:/$($encRel):?$selectClause"
        }
        else { "$graphBase/v1.0/drives/$($best.id)/root?$selectClause" }

        return Invoke-GraphSafe $itemUri
    }

    function Get-ItemPermissions {
        # Returns a FLAT enumeration of permission objects. See Get-GraphPages for why the
        # comma idiom is avoided on both sides of this call.
        param([string]$DriveId, [string]$ItemId)
        try { return @(Get-GraphPages -Uri "$graphBase/v1.0/drives/$DriveId/items/$ItemId/permissions") }
        catch { Write-Warning "  Could not read permissions for item $ItemId : $($_.Exception.Message)"; return @() }
    }

    # Entra group expansion. SharePoint groups are deliberately not attempted: Graph exposes
    # them as an id plus a login name with no membership relationship.
    $GroupMemberCache = @{}
    function Get-GroupMembersCached {
        param([string]$GroupId)
        if (-not $GroupId) { return $null }
        if ($GroupMemberCache.ContainsKey($GroupId)) { return $GroupMemberCache[$GroupId] }
        $out = @{ Ok = $false; Members = @(); Truncated = $false; DisplayName = ''; Error = '' }
        $meta = Invoke-GraphSafe "$graphBase/v1.0/groups/$($GroupId)?`$select=id,displayName,mail,groupTypes,securityEnabled"
        if ($meta.Ok) { $out.DisplayName = "$($meta.Data.displayName)" }
        $r = Invoke-GraphSafe "$graphBase/v1.0/groups/$GroupId/transitiveMembers/microsoft.graph.user?`$select=id,displayName,userPrincipalName,mail,userType&`$top=999"
        if ($r.Ok) {
            $members = @($r.Data.value)
            $next = $r.Data.'@odata.nextLink'
            while ($next -and $members.Count -lt $MaxGroupMembers) {
                $more = Invoke-GraphSafe $next
                if (-not $more.Ok) { break }
                $members += @($more.Data.value)
                $next = $more.Data.'@odata.nextLink'
            }
            $out.Ok = $true
            if ($members.Count -gt $MaxGroupMembers) {
                $out.Truncated = $true
                $members = @($members | Select-Object -First $MaxGroupMembers)
            }
            elseif ($next) { $out.Truncated = $true }
            $out.Members = $members
        }
        else { $out.Error = $r.Error }
        $GroupMemberCache[$GroupId] = $out
        return $out
    }

    # -----------------------------------------------------------------------
    # Flatten one permission (ACE) into one row per principal
    # -----------------------------------------------------------------------
    function ConvertTo-AccessRows {
        param($Ace, $Item, [string]$ItemLabel, [string]$SiteWebUrl)

        $rows = [System.Collections.Generic.List[object]]::new()
        $roles = @($Ace.roles)
        $access = Get-AccessLabel -Roles $roles
        $link = $Ace.link
        $isLink = [bool]$link
        $linkDesc = Get-LinkDescription -Link $link
        $linkScope = "$($link.scope)"
        # Graph hands this back as a [DateTime] here, not the ISO string the docs show, so a
        # string match on the documented '0001-01-01' sentinel never fires: en-US renders
        # DateTime.MinValue as '01/01/0001 00:00:00'. Normalise to a real DateTime, drop the
        # sentinel, and emit ISO 8601 UTC, because '09/05/2026' is ambiguous in an artifact
        # that may be read outside the US.
        $expires = ''
        if ($Ace.expirationDateTime) {
            $expDt = [datetime]::MinValue
            if ($Ace.expirationDateTime -is [datetime]) { $expDt = $Ace.expirationDateTime }
            else { [void][datetime]::TryParse("$($Ace.expirationDateTime)", [ref]$expDt) }
            if ($expDt -gt [datetime]::MinValue) { $expires = $expDt.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ') }
        }
        $shareId = "$($Ace.shareId)"
        $invitationEmail = "$($Ace.invitation.email)"

        # Direct grants populate grantedToV2 (singular). Sharing links populate
        # grantedToIdentitiesV2 (PLURAL). Reading only the singular property is why
        # "specific people" invitees commonly render blank.
        $identities = @()
        if ($Ace.grantedToIdentitiesV2) { $identities += @($Ace.grantedToIdentitiesV2) }
        elseif ($Ace.grantedToIdentities) { $identities += @($Ace.grantedToIdentities) }
        if ($Ace.grantedToV2) { $identities += @($Ace.grantedToV2) }
        elseif ($Ace.grantedTo -and -not $Ace.grantedToV2) { $identities += @($Ace.grantedTo) }

        function New-Row {
            param(
                [string]$Principal, [string]$Identifier, [string]$PrincipalType,
                [string]$ViaGroup = '', [string]$Notes = '', [bool]$Broad = $false,
                [string]$DirectoryId = ''
            )
            $isGuest = ($PrincipalType -like '*Guest*') -or ($Identifier -match '#EXT#') -or
                       ($PrincipalType -eq 'Anonymous')
            # A group row stands for an unknown number of people. Ranking it Info would
            # bury the single row that actually represents the largest population.
            $isGroup = ($PrincipalType -match 'group') -and ($PrincipalType -notmatch 'Sharing link group')
            $risk = 'Info'
            if ($linkScope -eq 'anonymous') { $risk = 'Critical' }
            elseif ($Broad -and ($roles -contains 'write' -or $roles -contains 'owner')) { $risk = 'Critical' }
            elseif ($Broad) { $risk = 'High' }
            elseif ($linkScope -eq 'organization' -and ($roles -contains 'write' -or $roles -contains 'owner')) { $risk = 'High' }
            elseif ($linkScope -eq 'organization') { $risk = 'Medium' }
            elseif ($isGuest -and ($roles -contains 'write' -or $roles -contains 'owner')) { $risk = 'High' }
            elseif ($isGuest) { $risk = 'Medium' }
            elseif ($isGroup -and ($roles -contains 'write' -or $roles -contains 'owner')) { $risk = 'Medium' }
            elseif ($isGroup -or ($roles -contains 'owner')) { $risk = 'Low' }

            if ($isGroup -and -not $ViaGroup) {
                $Notes = if ($ExpandGroups) { "$Notes" }
                         else { ("$Notes Group membership was not expanded, so the number of people this row represents is unknown. Rerun with -ExpandGroups.").Trim() }
            }

            [pscustomobject]@{
                Item          = $ItemLabel
                ItemUrl       = "$($Item.webUrl)"
                Grant         = $(if ($isLink) { 'Sharing link' } else { 'Direct grant' })
                LinkAudience  = $linkDesc
                LinkUrl       = "$($link.webUrl)"
                Principal     = $Principal
                Identifier    = $Identifier
                PrincipalType = $PrincipalType
                ViaGroup      = $ViaGroup
                Access        = $access
                Roles         = ($roles -join ', ')
                External      = $isGuest
                Expires       = $expires
                Risk          = $risk
                PermissionId  = "$($Ace.id)"
                ShareId       = $shareId
                DirectoryId   = $DirectoryId
                SiteUrl       = $SiteWebUrl
                Notes         = $Notes
            }
        }

        # An anonymous link grants access with no identity attached at all.
        if ($linkScope -eq 'anonymous') {
            $rows.Add((New-Row -Principal 'Anyone with the link' -Identifier '(no sign-in required)' `
                        -PrincipalType 'Anonymous' -Broad $true `
                        -Notes 'Anonymous "Anyone" link. Anyone holding the URL can open this, inside or outside the tenant, with no authentication and no audit identity.'))
        }
        elseif ($linkScope -eq 'organization' -and -not $identities) {
            $rows.Add((New-Row -Principal 'Everyone in the organization' -Identifier "$orgName" `
                        -PrincipalType 'Organization' -Broad $true `
                        -Notes 'Organization-wide link. Every signed-in tenant member who holds the URL can open this.'))
        }

        foreach ($idset in $identities) {
            $handled = $false

            foreach ($kind in @('user', 'group', 'application', 'device')) {
                $ident = $idset.$kind
                if (-not $ident) { continue }
                $handled = $true
                $name = "$($ident.displayName)"
                $id = "$($ident.id)"
                $type = switch ($kind) {
                    'user'        { 'Entra user' }
                    'group'       { 'Entra group' }
                    'application' { 'Application' }
                    'device'      { 'Device' }
                }
                $ident2 = $id
                # SharePoint mirrors the same principal into siteUser/siteGroup with the
                # claim login name, which carries the UPN and the real principal class.
                $site = $idset.siteUser
                if (-not $site) { $site = $idset.siteGroup }
                if (-not $site) { $site = $idset.sharePointGroup }
                $decoded = $null
                if ($site -and $site.loginName) {
                    $decoded = ConvertFrom-SharePointLoginName -LoginName "$($site.loginName)"
                    if ($decoded.Identifier -and $decoded.Identifier -ne "$($site.loginName)") { $ident2 = $decoded.Identifier }
                    if ($decoded.Type -ne 'Unknown') { $type = $decoded.Type }
                }
                if (-not $name -and $site) { $name = "$($site.displayName)" }

                $broad = [bool]($decoded -and $decoded.IsBroad)
                $note = ''
                if ($decoded) { $note = $decoded.Note }
                $dirId = ''
                if ($decoded -and $decoded.DirectoryId) { $dirId = $decoded.DirectoryId }
                elseif ($kind -eq 'group') { $dirId = $id }

                $rows.Add((New-Row -Principal $name -Identifier $ident2 -PrincipalType $type `
                            -Notes $note -Broad $broad -DirectoryId $dirId))

                # Flatten Entra groups to named people. This is the difference between
                # "6 rows with access" and "6 rows covering 431 people".
                if ($ExpandGroups -and $dirId -and $type -match 'group') {
                    $g = Get-GroupMembersCached -GroupId $dirId
                    if ($g -and $g.Ok) {
                        foreach ($m in $g.Members) {
                            $mType = if ("$($m.userType)" -eq 'Guest') { 'Guest user' } else { 'Entra user' }
                            $rows.Add((New-Row -Principal "$($m.displayName)" -Identifier "$($m.userPrincipalName)" `
                                        -PrincipalType $mType -ViaGroup $name `
                                        -Notes 'Transitive member of the group above.'))
                        }
                        if ($g.Truncated) {
                            $rows.Add((New-Row -Principal "(membership truncated)" -Identifier "$dirId" `
                                        -PrincipalType 'Notice' -ViaGroup $name `
                                        -Notes "More than $MaxGroupMembers members. Raise -MaxGroupMembers to list them all."))
                        }
                    }
                    elseif ($g) {
                        $rows.Add((New-Row -Principal "(membership unavailable)" -Identifier "$dirId" `
                                    -PrincipalType 'Notice' -ViaGroup $name `
                                    -Notes "Group members could not be read: $($g.Error)"))
                    }
                }
                break
            }

            if ($handled) { continue }

            # No user/group facet: SharePoint-only principal (site user or site group).
            $site = $idset.siteUser
            if (-not $site) { $site = $idset.siteGroup }
            if (-not $site) { $site = $idset.sharePointGroup }
            if (-not $site) { continue }

            $decoded = ConvertFrom-SharePointLoginName -LoginName "$($site.loginName)"
            $type = $decoded.Type
            $note = $decoded.Note
            if ($idset.siteGroup -or $idset.sharePointGroup) {
                if ($type -eq 'Unknown') { $type = 'SharePoint group' }
                if ($SiteWebUrl -and $site.id -match '^\d+$') {
                    $note = ("$note Graph cannot enumerate SharePoint group members; view them at " +
                             "$SiteWebUrl/_layouts/15/people.aspx?MembershipGroupId=$($site.id)").Trim()
                }
                else {
                    $note = "$note Graph cannot enumerate SharePoint group members.".Trim()
                }
            }

            $rows.Add((New-Row -Principal "$($site.displayName)" -Identifier $decoded.Identifier `
                        -PrincipalType $type -Notes $note -Broad $decoded.IsBroad -DirectoryId $decoded.DirectoryId))

            if ($ExpandGroups -and $decoded.DirectoryId -and $type -match 'group') {
                $g = Get-GroupMembersCached -GroupId $decoded.DirectoryId
                if ($g -and $g.Ok) {
                    foreach ($m in $g.Members) {
                        $mType = if ("$($m.userType)" -eq 'Guest') { 'Guest user' } else { 'Entra user' }
                        $rows.Add((New-Row -Principal "$($m.displayName)" -Identifier "$($m.userPrincipalName)" `
                                    -PrincipalType $mType -ViaGroup "$($site.displayName)" `
                                    -Notes 'Transitive member of the group above.'))
                    }
                    if ($g.Truncated) {
                        $rows.Add((New-Row -Principal '(membership truncated)' -Identifier "$($decoded.DirectoryId)" `
                                    -PrincipalType 'Notice' -ViaGroup "$($site.displayName)" `
                                    -Notes "More than $MaxGroupMembers members. Raise -MaxGroupMembers to list them all."))
                    }
                }
            }
        }

        # An invitation that has not been redeemed carries no principal yet, only an email.
        if (-not $rows.Count -and $invitationEmail) {
            $rows.Add((New-Row -Principal $invitationEmail -Identifier $invitationEmail `
                        -PrincipalType 'Unredeemed invitation' `
                        -Notes 'Invitation sent but not yet redeemed; no account is attached to the grant yet.'))
        }
        if (-not $rows.Count) {
            # An existingAccess link legitimately carries no identities: it grants nothing new
            # and exists only so a URL can be handed to people who already have access. Saying
            # "identity withheld" there would invent a privilege problem that does not exist.
            $noIdNote = if ($linkScope -eq 'existingAccess') {
                'An "existing access" link grants no new permissions, so it carries no identities by design. Who can open it is determined entirely by the other rows on this item.'
            }
            else {
                'Graph returned this permission with no identity facet. The most common cause is a low-privilege caller: identity and secret fields are withheld from callers who could not have created the permission.'
            }
            $rows.Add((New-Row -Principal '(no identity returned)' -Identifier '' -PrincipalType 'Unknown' -Notes $noIdNote))
        }
        # Flat emission, same reasoning as Get-GraphPages. Today's callers all use
        # foreach, which enumerates and would survive the comma idiom, but a future
        # caller reaching for @() would silently get one merged row instead of N.
        return $rows.ToArray()
    }

    # -----------------------------------------------------------------------
    # Per-URL processing
    # -----------------------------------------------------------------------
    $now = Get-Date
    $stamp = $now.ToString('yyyyMMdd_HHmmss')
    if (-not $OutputFolder) { $OutputFolder = Join-Path ([Environment]::GetFolderPath('Desktop')) 'SharingLinkAccess' }
    if (-not (Test-Path $OutputFolder)) { $null = New-Item -ItemType Directory -Path $OutputFolder -Force }

    $allRows = [System.Collections.Generic.List[object]]::new()
    $reports = [System.Collections.Generic.List[object]]::new()

    # A tenant that has renamed its SharePoint domain keeps the old hostname alive as a
    # redirect, and old links keep circulating for years. Following that redirect from
    # /shares can loop, which surfaces as "Too many redirects performed" rather than as
    # anything mentioning a rename. Ask the tenant what its current host actually is so a
    # legacy-host link can be retried against it instead of guessed at.
    $TenantSpoHost = ''
    try {
        $rootSite = Invoke-MgGraphRequest -Method GET -Uri "$graphBase/v1.0/sites/root?`$select=webUrl" -OutputType PSObject
        if ($rootSite.webUrl) { $TenantSpoHost = ([uri]"$($rootSite.webUrl)").Host }
    }
    catch { Write-Verbose "Could not read the tenant root site: $($_.Exception.Message)" }

    function Get-HostRewrittenUrl {
        # Plain string swap of the authority. UriBuilder re-encodes the path, which mangles
        # the ':f:' segment and the base64url token that identify a sharing link.
        param([string]$Value, [string]$NewHost)
        if (-not $NewHost) { return '' }
        $current = ''
        try { $current = ([uri]$Value).Host } catch { return '' }
        if (-not $current -or $current -ieq $NewHost) { return '' }
        return ($Value -replace ('(?i)^(https?://)' + [regex]::Escape($current)), ('${1}' + $NewHost))
    }

    foreach ($target in $targets) {
        Write-Host ''
        Write-Host "Resolving: $target" -ForegroundColor Cyan

        # What the caller actually handed over, kept separate from the working URL. A
        # hostname rewrite must not silently rewrite the record of what was submitted:
        # in an evidence artifact those are two different facts.
        $submittedUrl = $target
        $rewrittenUrl = ''

        $linkHost = ''
        try { $linkHost = ([uri]$target).Host } catch { }
        if ($TenantSpoHost -and $linkHost -and $linkHost -ine $TenantSpoHost) {
            Write-Host ("  Link host is $linkHost but this tenant now serves SharePoint at $TenantSpoHost (renamed domain, or a link from elsewhere).") -ForegroundColor Yellow
        }

        $looksLikeSharingLink = $target -match $SharingUrlPattern
        $res = Resolve-ItemFromShareToken -Value $target -Redeem:$RedeemIfNecessary
        $resolvedVia = 'shares API'
        if ($RedeemIfNecessary) { $resolvedVia = 'shares API with redeemSharingLinkIfNecessary' }

        # Retry on the tenant's current host. The sharing token identifies the item, not the
        # hostname, so it stays valid once the legacy-domain redirect is taken out of play.
        $retryError = ''
        if (-not $res.Ok) {
            $alt = Get-HostRewrittenUrl -Value $target -NewHost $TenantSpoHost
            if ($alt) {
                Write-Host "  Retrying on the tenant's current SharePoint host ($TenantSpoHost)..." -ForegroundColor Yellow
                $altRes = Resolve-ItemFromShareToken -Value $alt -Redeem:$RedeemIfNecessary
                if ($altRes.Ok) {
                    $res = $altRes
                    $target = $alt
                    $rewrittenUrl = $alt
                    $resolvedVia = "$resolvedVia, after rewriting a legacy SharePoint hostname"
                }
                else { $retryError = "$($altRes.Error)" }
            }
        }

        if (-not $res.Ok) {
            $firstError = $res.Error
            $firstStatus = $res.Status
            if (-not $looksLikeSharingLink) {
                Write-Host '  /shares declined; falling back to site/drive/path resolution...' -ForegroundColor Yellow
                $res = Resolve-ItemFromPathUrl -Value $target
                $resolvedVia = 'site/drive path walk'
                if (-not $res.Ok) {
                    $alt = Get-HostRewrittenUrl -Value $target -NewHost $TenantSpoHost
                    if ($alt) {
                        $res = Resolve-ItemFromPathUrl -Value $alt
                        if ($res.Ok) { $resolvedVia = 'site/drive path walk, after rewriting a legacy SharePoint hostname' }
                    }
                }
            }
            if (-not $res.Ok) {
                # Additive, not exclusive. The first attempt and the rewritten-host retry can
                # fail for different reasons, and the second reason is often the more useful
                # one: a redirect loop followed by a 403 means the rewrite got past the
                # redirect and then hit a permission wall. Reporting only the first failure
                # buries that.
                $hints = @()
                if ($firstError -match 'Too many redirects') {
                    $h = ' A redirect loop means the link host keeps handing the request onward, which is what a renamed SharePoint domain does to links minted before the rename.'
                    if ($TenantSpoHost) {
                        # Do not claim the item is gone. A sharing token is minted against the
                        # original host, so rewriting the host can invalidate the token itself;
                        # a failed retry says nothing either way about whether the item exists.
                        $h += " Retrying on $TenantSpoHost also failed, which does NOT establish that the item is gone: the sharing token is issued against the original hostname, so moving it to another host can invalidate the token rather than the item."
                        $h += " To settle whether the content still exists, pass the item's plain path URL on $TenantSpoHost instead of the sharing link. That route never uses a token."
                    }
                    else { $h += ' The tenant root site could not be read to find the current hostname, so no retry was attempted.' }
                    $hints += $h
                }
                if ($firstStatus -eq 403 -or $retryError -match 'Forbidden|accessDenied|\b403\b') {
                    $hints += (' A 403 was returned, meaning Graph reached the item and refused, which is different from the item not existing.' +
                        ' Two causes, in order of likelihood:' +
                        ' (1) The signed-in account has no access to the item. That is the normal state for an investigator who deliberately did not accept the share, and a bare /shares GET grants nothing.' +
                        " Resolve it by signing in app-only with -ClientId and -CertificateThumbprint, which is treated as an owner and sees the whole ACL, or by passing the item's plain path URL if you know it, or with -RedeemIfNecessary to send the documented metadata-peek header." +
                        ' (2) Scopes. Microsoft documents /shares as requiring Files.ReadWrite (delegated) or Files.ReadWrite.All / Sites.ReadWrite.All (application); this script requests read-only scopes by default, so -Scopes can add one.')
                }
                if ($firstStatus -eq 404 -and -not $hints.Count) {
                    $hints += ' A 404 here usually means the link was deleted, expired, or was issued by a different tenant than the one signed in to.'
                }
                $hint = ($hints -join '')
                Write-Warning "Could not resolve '$target'. shares API: $firstError$hint"
                if ($retryError) { Write-Warning "  Retry on ${TenantSpoHost}: $retryError" }
                if ($resolvedVia -like 'site/drive path walk*') { Write-Warning "  Path fallback also failed: $($res.Error)" }
                continue
            }
        }

        $item = $res.Data
        $driveId = "$($item.parentReference.driveId)"
        $itemId = "$($item.id)"
        $isFolder = [bool]$item.folder
        $itemLabel = "$($item.name)"
        if (-not $itemLabel) { $itemLabel = '(root)' }
        $siteWebUrl = ''
        if ($item.sharepointIds -and $item.sharepointIds.siteUrl) { $siteWebUrl = "$($item.sharepointIds.siteUrl)".TrimEnd('/') }

        Write-Host "  $($(if ($isFolder) { 'Folder' } else { 'File' })): $itemLabel" -ForegroundColor Green
        Write-Host "  Actual location: $($item.webUrl)"
        # Never print "not redeemed" when a Prefer header was in fact sent. This line is the
        # script's central safety claim, so it has to track what actually happened.
        if ($RedeemIfNecessary) {
            Write-Host "  Resolved via $resolvedVia" -ForegroundColor Yellow
            Write-Host "  Prefer: redeemSharingLinkIfNecessary WAS sent. $signedInAs may now appear in this item's access list." -ForegroundColor Yellow
        }
        else {
            Write-Host "  Resolved via $resolvedVia (link NOT redeemed; no Prefer header sent)"
        }
        if ($rewrittenUrl) { Write-Host "  Note: resolved against the rewritten host, not the URL as submitted." -ForegroundColor Yellow }

        if (-not $driveId -or -not $itemId) {
            Write-Warning '  Resolved item is missing driveId/id; cannot read permissions.'
            continue
        }

        $perms = @(Get-ItemPermissions -DriveId $driveId -ItemId $itemId)
        Write-Host "  $($perms.Count) permission entries"

        $rows = [System.Collections.Generic.List[object]]::new()
        foreach ($ace in $perms) {
            foreach ($r in (ConvertTo-AccessRows -Ace $ace -Item $item -ItemLabel $itemLabel -SiteWebUrl $siteWebUrl)) { $rows.Add($r) }
        }

        # -------------------------------------------------------------------
        # Inheritance inference. SharePoint document libraries do not return
        # inheritedFrom, so the only available evidence is whether the same
        # (principal + roles) signature also exists on an ancestor. Reported as an
        # inference, never as fact.
        # -------------------------------------------------------------------
        $ancestors = [System.Collections.Generic.List[object]]::new()
        $inheritedSignatures = @{}
        if (-not $SkipInheritanceCheck) {
            Write-Host '  Walking ancestors to infer inherited vs unique permissions...'
            $cursorId = "$($item.parentReference.id)"
            $depth = 0
            while ($cursorId -and $depth -lt $MaxAncestorDepth) {
                $depth++
                $a = Invoke-GraphSafe "$graphBase/v1.0/drives/$driveId/items/$($cursorId)?`$select=id,name,webUrl,parentReference,root"
                if (-not $a.Ok) { break }
                $aPerms = @(Get-ItemPermissions -DriveId $driveId -ItemId "$($a.Data.id)")
                $aRows = [System.Collections.Generic.List[object]]::new()
                foreach ($ace in $aPerms) {
                    foreach ($r in (ConvertTo-AccessRows -Ace $ace -Item $a.Data -ItemLabel "$($a.Data.name)" -SiteWebUrl $siteWebUrl)) { $aRows.Add($r) }
                }
                foreach ($r in $aRows) {
                    if ($r.ViaGroup) { continue }
                    $sig = "$($r.Identifier)|$($r.Roles)|$($r.Grant)"
                    if (-not $inheritedSignatures.ContainsKey($sig)) { $inheritedSignatures[$sig] = "$($a.Data.name)" }
                }
                $ancestors.Add([pscustomobject]@{
                        Name  = "$($a.Data.name)"
                        Url   = "$($a.Data.webUrl)"
                        Count = $aRows.Count
                    })
                if ($a.Data.root) { break }
                $cursorId = "$($a.Data.parentReference.id)"
            }
            foreach ($r in $rows) {
                if ($r.ViaGroup) { continue }
                $sig = "$($r.Identifier)|$($r.Roles)|$($r.Grant)"
                if ($inheritedSignatures.ContainsKey($sig)) {
                    $r.Notes = ("Also present on ancestor '$($inheritedSignatures[$sig])', so this access most likely predates the link and was inherited. $($r.Notes)").Trim()
                }
            }
        }

        # -------------------------------------------------------------------
        # Children with permissions that differ from the parent
        # -------------------------------------------------------------------
        $childFindings = [System.Collections.Generic.List[object]]::new()
        if ($IncludeChildren -and $isFolder) {
            Write-Host "  Checking up to $MaxChildren children for permissions that differ..."
            $kids = @()
            $kr = Invoke-GraphSafe "$graphBase/v1.0/drives/$driveId/items/$itemId/children?`$select=id,name,webUrl,folder,file&`$top=$MaxChildren"
            if ($kr.Ok) { $kids = @($kr.Data.value | Select-Object -First $MaxChildren) }
            $parentSigs = @{}
            foreach ($r in $rows) { if (-not $r.ViaGroup) { $parentSigs["$($r.Identifier)|$($r.Roles)|$($r.Grant)"] = $true } }
            foreach ($kid in $kids) {
                $kp = @(Get-ItemPermissions -DriveId $driveId -ItemId "$($kid.id)")
                foreach ($ace in $kp) {
                    foreach ($r in (ConvertTo-AccessRows -Ace $ace -Item $kid -ItemLabel "$($kid.name)" -SiteWebUrl $siteWebUrl)) {
                        if ($r.ViaGroup) { continue }
                        $sig = "$($r.Identifier)|$($r.Roles)|$($r.Grant)"
                        if (-not $parentSigs.ContainsKey($sig)) {
                            $r.Notes = ("Present on this child but not on the shared folder: inheritance is broken here. $($r.Notes)").Trim()
                            $childFindings.Add($r)
                        }
                    }
                }
            }
            Write-Host "    $($childFindings.Count) child-level permissions not present on the folder"
        }

        foreach ($r in $rows) { $allRows.Add($r) }
        foreach ($r in $childFindings) { $allRows.Add($r) }

        $reports.Add([pscustomobject]@{
                Input        = $submittedUrl
                RewrittenUrl = $rewrittenUrl
                Item        = $item
                ItemLabel   = $itemLabel
                IsFolder    = $isFolder
                ResolvedVia = $resolvedVia
                SiteUrl     = $siteWebUrl
                Rows        = @($rows)
                Children    = @($childFindings)
                Ancestors   = @($ancestors)
                Permissions = $perms
                Audit       = @()
                AuditNote   = ''
            })
    }

    if (-not $reports.Count) { throw 'No links resolved. Nothing to report.' }

    # -----------------------------------------------------------------------
    # Audit log pivot: permissions say who COULD open it, the audit log says who DID.
    # -----------------------------------------------------------------------
    if ($IncludeAuditLog) {
        Write-Host ''
        Write-Host "Audit log pivot (last $AuditDays days)..." -ForegroundColor Cyan
        $exoOk = $false
        if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
            Write-Warning '  ExchangeOnlineManagement is not installed. Install-Module ExchangeOnlineManagement, then rerun with -IncludeAuditLog.'
        }
        else {
            try {
                Import-Module ExchangeOnlineManagement -ErrorAction Stop
                $connected = $false
                try { $connected = [bool](Get-ConnectionInformation -ErrorAction Stop) } catch { }
                if (-not $connected) {
                    $exoConnect = @{ ShowBanner = $false }
                    if ($AppOnly) {
                        # Same certificate identity as Graph; the app additionally needs
                        # Exchange.ManageAsApp plus the View-Only Audit Logs role.
                        # -Organization requires a DOMAIN, never a tenant GUID, which is why
                        # the onmicrosoft domain derived from the link host is carried
                        # separately from the GUID used for Graph.
                        $exoConnect.AppId = $ClientId
                        $exoConnect.CertificateThumbprint = $CertificateThumbprint
                        $exoConnect.Organization = $ExoOrganization
                        if (-not $ExoOrganization) {
                            throw ('The audit pivot needs a tenant domain for Connect-ExchangeOnline -Organization, ' +
                                'and none could be derived from the link host. Pass -TenantId contoso.onmicrosoft.com.')
                        }
                    }
                    elseif ($UseDeviceCode) { $exoConnect.Device = $true }
                    Write-Host '  Signing in to Exchange Online (audit log search)...'
                    Connect-ExchangeOnline @exoConnect
                }
                $exoOk = $true
            }
            catch { Write-Warning "  Exchange Online sign-in failed: $($_.Exception.Message)" }
        }

        if ($exoOk) {
            $ops = $SharingOperations
            if ($IncludeFileAccess) { $ops = $ops + $FileOperations }
            $start = $now.AddDays(-$AuditDays)

            foreach ($rep in $reports) {
                # The audit log stores ObjectId as the item URL. Graph returns it percent
                # encoded, Purview usually stores it decoded, so both forms are supplied.
                $encoded = "$($rep.Item.webUrl)"
                $decoded = [uri]::UnescapeDataString($encoded)
                $objectIds = @($encoded, $decoded) | Select-Object -Unique

                $records = [System.Collections.Generic.List[object]]::new()
                try {
                    # ReturnLargeSet returns unsorted data with duplicates and must be paged
                    # until an empty page, then deduplicated by Identity.
                    $session = "SharingLinkAccess_$([guid]::NewGuid().ToString('N').Substring(0,12))"
                    do {
                        $page = @(Search-UnifiedAuditLog -StartDate $start -EndDate $now `
                                -ObjectIds $objectIds -Operations $ops `
                                -SessionId $session -SessionCommand ReturnLargeSet -ResultSize 5000 -ErrorAction Stop)
                        foreach ($p in $page) { $records.Add($p) }
                    } while ($page.Count -gt 0 -and $records.Count -lt 50000)
                }
                catch {
                    $rep.AuditNote = "Audit search failed: $($_.Exception.Message)"
                    Write-Warning "  $($rep.AuditNote)"
                }

                $clean = @($records | Sort-Object Identity -Unique)
                $events = foreach ($rec in $clean) {
                    $data = $null
                    try { $data = $rec.AuditData | ConvertFrom-Json } catch { }
                    [pscustomobject]@{
                        When        = $rec.CreationDate
                        Operation   = "$($rec.Operations)"
                        ActingUser  = "$($rec.UserIds)"
                        TargetUser  = "$($data.TargetUserOrGroupName)"
                        TargetType  = "$($data.TargetUserOrGroupType)"
                        ObjectId    = "$($data.ObjectId)"
                        SharingId   = "$($data.UniqueSharingId)"
                        LinkType    = "$($data.EventData)"
                        ClientIp    = "$($data.ClientIP)"
                    }
                }
                $rep.Audit = @($events | Sort-Object { $_.When -as [datetime] })
                if (-not $rep.AuditNote) {
                    $rep.AuditNote = "$($rep.Audit.Count) sharing events in the last $AuditDays days."
                    if (-not $rep.Audit.Count) {
                        $rep.AuditNote += ' An empty result is not proof nothing happened: unified audit retention is 180 days on E3 and 365 on E5, and events older than the retention window are gone.'
                    }
                }
                Write-Host "  $($rep.ItemLabel): $($rep.Audit.Count) audit events"
            }
        }
    }

    # -----------------------------------------------------------------------
    # Outputs
    # -----------------------------------------------------------------------
    $primary = $reports[0]
    $safeName = ("$($primary.ItemLabel)" -replace '[^A-Za-z0-9]+', '')
    if (-not $safeName) { $safeName = 'item' }
    if ($reports.Count -gt 1) { $safeName = "$safeName-plus$($reports.Count - 1)" }
    $htmlPath = Join-Path $OutputFolder "SharingLinkAccess_${safeName}_$stamp.html"
    $csvPath = Join-Path $OutputFolder "SharingLinkAccess_${safeName}_$stamp.csv"

    $RiskRank = @{ 'Critical' = 5; 'High' = 4; 'Medium' = 3; 'Low' = 2; 'Info' = 1 }
    $ordered = @($allRows | Sort-Object @{ e = { $RiskRank["$($_.Risk)"] }; Descending = $true }, Item, Principal)
    if ($ordered) { $ordered | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8 }

    $riskColors = @{
        'Critical' = @{ fg = '#ffffff'; bg = '#b3261e' }
        'High'     = @{ fg = '#ffffff'; bg = '#b45309' }
        'Medium'   = @{ fg = '#402c00'; bg = '#fcd34d' }
        'Low'      = @{ fg = '#1e3a5f'; bg = '#dbeafe' }
        'Info'     = @{ fg = '#1e3a5f'; bg = '#dbeafe' }
    }

    function New-AccessRowHtml {
        param($r)
        $rc = $riskColors["$($r.Risk)"]
        if (-not $rc) { $rc = $riskColors['Info'] }
        $badge = "<span class='badge' style='background:$($rc.bg);color:$($rc.fg)'>$(EscHtml $r.Risk)</span>"
        $who = "<strong>$(EscHtml $r.Principal)</strong>"
        if ($r.Identifier -and $r.Identifier -ne $r.Principal) { $who += "<div class='sub'>$(EscHtml $r.Identifier)</div>" }
        if ($r.ViaGroup) { $who += "<div class='via'>via $(EscHtml $r.ViaGroup)</div>" }
        $ext = ''
        if ($r.External) { $ext = "<div class='flags'>External / guest</div>" }
        $grant = EscHtml $r.Grant
        if ($r.LinkAudience) { $grant += "<div class='sub'>$(EscHtml $r.LinkAudience)</div>" }
        $exp = ''
        if ($r.Expires) { $exp = "<div class='sub'>Expires $(EscHtml $r.Expires)</div>" }
        $rank = $RiskRank["$($r.Risk)"]
        if (-not $rank) { $rank = 0 }
        return ("<tr><td>$who$ext</td><td>$(EscHtml $r.PrincipalType)</td><td>$grant</td>" +
            "<td>$(EscHtml $r.Access)$exp</td><td data-v='$rank'>$badge</td>" +
            "<td class='note-cell'>$(EscHtml $r.Notes)</td></tr>`n")
    }

    $sectionsHtml = ''
    foreach ($rep in $reports) {
        $rows = @($rep.Rows | Sort-Object @{ e = { $RiskRank["$($_.Risk)"] }; Descending = $true }, Principal)
        $people = @($rows | Where-Object { $_.PrincipalType -match 'user' } | Select-Object -ExpandProperty Identifier -Unique)
        $guests = @($rows | Where-Object { $_.External } | Select-Object -ExpandProperty Identifier -Unique)
        $links = @($rep.Permissions | Where-Object { $_.link })
        $anon = @($rows | Where-Object { $_.PrincipalType -eq 'Anonymous' })
        $broad = @($rows | Where-Object { $_.PrincipalType -in @('Anonymous', 'Organization', 'Everyone', 'Everyone except external users', 'All authenticated users') })

        $rowsHtml = ''
        foreach ($r in $rows) { $rowsHtml += New-AccessRowHtml $r }
        if (-not $rowsHtml) { $rowsHtml = "<tr><td colspan='6' class='empty'>No permissions returned.</td></tr>" }

        $childHtml = ''
        foreach ($r in @($rep.Children)) { $childHtml += New-AccessRowHtml $r }

        $ancestorHtml = ''
        foreach ($a in @($rep.Ancestors)) {
            $ancestorHtml += "<li>$(EscHtml $a.Name) <span class='sub'>($($a.Count) permission rows)</span></li>`n"
        }
        if (-not $ancestorHtml) { $ancestorHtml = '<li class="sub">Ancestor walk was skipped or returned nothing.</li>' }

        $auditHtml = ''
        foreach ($e in @($rep.Audit)) {
            $auditHtml += ("<tr><td>$(EscHtml $e.When)</td><td>$(EscHtml $e.Operation)</td><td>$(EscHtml $e.ActingUser)</td>" +
                "<td>$(EscHtml $e.TargetUser)</td><td>$(EscHtml $e.TargetType)</td><td class='mono'>$(EscHtml $e.SharingId)</td></tr>`n")
        }
        if (-not $auditHtml) {
            $msg = 'Audit pivot not run. Add -IncludeAuditLog to see who created the link, who was invited, and who redeemed it.'
            if ($IncludeAuditLog) { $msg = "$($rep.AuditNote)" }
            $auditHtml = "<tr><td colspan='6' class='empty'>$(EscHtml $msg)</td></tr>"
        }

        # Dashboard chrome earns its place at report scale, not at lookup scale. A search
        # box and sortable headers over nine rows are noise, so they only appear once the
        # table is long enough to need them.
        $idx = $reports.IndexOf($rep)
        $needsChrome = $rows.Count -ge 15
        $tblClass = if ($needsChrome) { 'sortable' } else { '' }
        $filterHtml = if ($needsChrome) { "<input class=`"tfilter`" data-target=`"tbl$idx`" type=`"search`" placeholder=`"Filter principals...`">" } else { '' }

        $anonBanner = ''
        if ($anon.Count) {
            $anonBanner = "<div class='banner crit'><strong>Anonymous link present.</strong> This item is reachable by anyone holding the URL, with no sign-in and no attributable identity in the audit log. Treat the URL itself as the credential.</div>"
        }

        $sectionsHtml += @"
<section class="card">
  <h2>$(EscHtml $rep.ItemLabel)</h2>
  <div class="kv"><span>Submitted</span><code>$(EscHtml $rep.Input)</code></div>
  $(if ($rep.RewrittenUrl) { "<div class='kv'><span>Resolved as</span><code>$(EscHtml $rep.RewrittenUrl)</code></div>" })
  <div class="kv"><span>Actual location</span><a href="$(EscHtml $rep.Item.webUrl)" target="_blank">$(EscHtml ([uri]::UnescapeDataString("$($rep.Item.webUrl)")))</a></div>
  <div class="kv"><span>Type</span>$(if ($rep.IsFolder) { 'Folder' } else { 'File' }) &middot; resolved via $(EscHtml $rep.ResolvedVia) &middot; link not redeemed</div>
  <div class="kv"><span>Last modified</span>$(EscHtml $rep.Item.lastModifiedDateTime) by $(EscHtml $rep.Item.lastModifiedBy.user.displayName)</div>
  $anonBanner
  <div class="tiles">
    <div class="tile"><div class="num">$($rows.Count)</div><div class="lbl">Access rows</div></div>
    <div class="tile"><div class="num">$($people.Count)</div><div class="lbl">Distinct people</div></div>
    <div class="tile"><div class="num">$($guests.Count)</div><div class="lbl">External / guest</div></div>
    <div class="tile"><div class="num">$($links.Count)</div><div class="lbl">Sharing links</div></div>
    <div class="tile"><div class="num">$($broad.Count)</div><div class="lbl">Broad-audience grants</div></div>
    <div class="tile"><div class="num">$(@($rep.Children).Count)</div><div class="lbl">Child exceptions</div></div>
  </div>

  <h3>Who can reach this</h3>
  $filterHtml
  <table class="$tblClass" id="tbl$idx">
    <thead><tr><th>Principal</th><th>Type</th><th>Granted by</th><th>Access</th><th>Risk</th><th>Notes</th></tr></thead>
    <tbody>
    $rowsHtml
    </tbody>
  </table>

  <h3>Inheritance</h3>
  <p class="note">SharePoint document libraries do not return the <code>inheritedFrom</code> property, so inheritance below is <em>inferred</em> by comparing this item's grants against its ancestors' grants. Ancestors checked, nearest first:</p>
  <ul class="ancestors">$ancestorHtml</ul>

  <h3>Child items with permissions the folder does not have ($(@($rep.Children).Count))</h3>
  $(if ($childHtml) {
    "<table class='sortable'><thead><tr><th>Principal</th><th>Type</th><th>Granted by</th><th>Access</th><th>Risk</th><th>Notes</th></tr></thead><tbody>$childHtml</tbody></table>"
  } elseif ($IncludeChildren) {
    "<p class='note'>No child item carries a permission that the folder does not already have.</p>"
  } else {
    "<p class='note'>Not checked. Add <code>-IncludeChildren</code> to look for broken inheritance beneath this folder.</p>"
  })

  <h3>Sharing timeline</h3>
  <table class="sortable">
    <thead><tr><th>When</th><th>Operation</th><th>Acting user</th><th>Target user</th><th>Target type</th><th>Sharing id</th></tr></thead>
    <tbody>
    $auditHtml
    </tbody>
  </table>
</section>
"@
    }

    $ownerBanner = ''
    if (-not $AppOnly) {
        $ownerBanner = @"
<div class="banner warn">
  <strong>Delegated sign-in: this list may be incomplete.</strong>
  Microsoft returns the full permission collection only to an owner or co-owner of the item. Every other caller receives
  <em>only the permissions that apply to them</em>, with no error and no warning, and the <code>shareId</code> and link
  <code>webUrl</code> fields are withheld entirely. If <code>$(EscHtml $signedInAs)</code> is not an owner or site collection
  admin of this item, treat these results as a floor, not a complete answer. Rerun with
  <code>-ClientId</code> and <code>-CertificateThumbprint</code> for an authoritative list.
</div>
"@
    }

    $totalGuests = @($allRows | Where-Object { $_.External } | Select-Object -ExpandProperty Identifier -Unique).Count
    $totalCritical = @($allRows | Where-Object { $_.Risk -eq 'Critical' }).Count

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>Sharing Link Access - $(EscHtml $primary.ItemLabel)</title>
<style>
  :root { --ink:#1a1f2b; --sub:#5b6474; --line:#e3e7ee; --card:#ffffff; --bg:#f4f6f9; --accent:#1f4e79; }
  * { box-sizing:border-box; }
  body { margin:0; font-family:'Segoe UI',system-ui,sans-serif; background:var(--bg); color:var(--ink); font-size:14px; }
  header { background:linear-gradient(120deg,#16324f,#1f4e79); color:#fff; padding:28px 36px; }
  header h1 { margin:0 0 6px; font-size:22px; font-weight:600; }
  header .meta { color:#c8d6e5; font-size:13px; }
  main { max-width:1500px; margin:0 auto; padding:24px 36px 60px; }
  .banner { border-radius:8px; padding:14px 18px; margin:0 0 20px; }
  .banner.warn { background:#fff7ed; border:1px solid #fdba74; border-left:5px solid #ea580c; }
  .banner.warn strong { color:#9a3412; }
  .banner.crit { background:#fef2f2; border:1px solid #fca5a5; border-left:5px solid #b3261e; }
  .banner.crit strong { color:#7f1d1d; }
  .banner.info { background:#eff6ff; border:1px solid #93c5fd; border-left:5px solid #1f4e79; }
  .tiles { display:flex; gap:14px; flex-wrap:wrap; margin:16px 0 24px; }
  .tile { background:var(--bg); border:1px solid var(--line); border-radius:10px; padding:12px 18px; min-width:140px; }
  .tile .num { font-size:24px; font-weight:700; }
  .tile .lbl { color:var(--sub); font-size:12px; text-transform:uppercase; letter-spacing:.04em; }
  section.card { background:var(--card); border:1px solid var(--line); border-radius:10px; padding:20px 24px 26px; margin:0 0 26px; }
  h2 { font-size:18px; margin:0 0 12px; color:var(--accent); }
  h3 { font-size:15px; margin:26px 0 8px; color:var(--accent); }
  .kv { display:flex; gap:12px; margin:4px 0; font-size:13px; align-items:baseline; }
  .kv span { color:var(--sub); min-width:120px; flex:none; }
  .kv code { background:var(--bg); padding:2px 6px; border-radius:4px; word-break:break-all; font-size:12px; }
  table { width:100%; border-collapse:collapse; background:var(--card); border:1px solid var(--line); border-radius:10px; overflow:hidden; }
  th { text-align:left; background:#eef2f7; padding:9px 12px; font-size:12px; text-transform:uppercase; letter-spacing:.03em; color:var(--sub); border-bottom:1px solid var(--line); }
  td { padding:10px 12px; border-bottom:1px solid var(--line); vertical-align:top; }
  tr:last-child td { border-bottom:none; }
  .sub { color:var(--sub); font-size:12px; margin-top:2px; word-break:break-all; }
  .via { color:#1f4e79; font-size:12px; margin-top:2px; }
  .flags { margin-top:4px; font-size:12px; color:#9a3412; }
  .note-cell { max-width:420px; color:var(--sub); font-size:12px; }
  .mono { font-family:Consolas,monospace; font-size:11px; word-break:break-all; }
  .empty { text-align:center; color:var(--sub); padding:22px; }
  .badge { display:inline-block; padding:2px 9px; border-radius:999px; font-size:12px; font-weight:600; }
  .note { color:var(--sub); font-size:13px; margin:8px 0; }
  .ancestors { color:var(--ink); font-size:13px; margin:6px 0 0; }
  a { color:var(--accent); }
  table.sortable thead th { cursor:pointer; user-select:none; white-space:nowrap; }
  table.sortable thead th:hover { background:#e2e9f2; }
  th.s-asc::after { content:' \25B2'; font-size:9px; }
  th.s-desc::after { content:' \25BC'; font-size:9px; }
  .tfilter { margin:0 0 8px; padding:7px 12px; border:1px solid var(--line); border-radius:8px; width:300px; font:inherit; font-size:13px; background:var(--card); }
  section.guide { background:var(--card); border:1px solid var(--line); border-radius:10px; padding:6px 24px 18px; }
  section.guide li { margin:7px 0; }
</style>
</head>
<body>
<header>
  <h1>Sharing Link Access Report</h1>
  <div class="meta">$(EscHtml $orgName) &middot; $($reports.Count) link(s) resolved &middot; Generated $($now.ToString('yyyy-MM-dd HH:mm')) &middot; Signed in as $(EscHtml $signedInAs) &middot; Read-only</div>
</header>
<main>
$ownerBanner
$(if ($RedeemIfNecessary) { @"
<div class="banner warn">
  <strong>Run with -RedeemIfNecessary.</strong> Links were resolved with the
  <code>Prefer: redeemSharingLinkIfNecessary</code> header, which Microsoft documents as granting access
  "only guaranteed to be granted for the duration of this request". That wording promises the access is temporary, not that
  no permission record was written, so <strong>$(EscHtml $signedInAs) may now appear in this item's own access list</strong>.
  Treat any row naming that account as potentially an artifact of this report rather than pre-existing access. The durable
  <code>Prefer: redeemSharingLink</code> header was not sent.
</div>
"@ } else { @"
<div class="banner info">
  <strong>The link was not redeemed.</strong> Every link here was resolved through the Microsoft Graph <code>/shares</code>
  endpoint with no <code>Prefer</code> header at all, so no account was added to any ACL by running this report.
  Opening one of these links in a browser instead would have accepted the sharing gesture and placed the operator into the
  very access list under review.
</div>
"@ })
<div class="banner info">
  $(if ($totalCritical) { "<strong>$totalCritical critical-risk grant(s)</strong> and <strong>$totalGuests external identit(ies)</strong> were found across the resolved items." } else { "No critical-risk grants were found. $totalGuests external identit(ies) across the resolved items." })
</div>

$sectionsHtml

<section class="guide">
<h2>How to read this</h2>
<ul>
  <li><strong>Granted by</strong> separates a <em>direct grant</em> (someone was added to the item or inherited it from a parent) from a <em>sharing link</em> (a URL was minted). Only sharing links can be revoked by deleting the link; direct grants survive it.</li>
  <li><strong>"Specific people" links</strong> store their invitees in Graph's <code>grantedToIdentitiesV2</code> collection, not <code>grantedToV2</code>. Underneath, SharePoint keeps them in hidden groups named <code>SharingLinks.{UniqueId}.{SharingLinkKind}.{ShareId}</code> that never appear in the site's group list, which is why this report and the browser "Manage access" pane can legitimately disagree.</li>
  <li><strong>Inheritance is inferred, not read.</strong> OneDrive for Business and SharePoint document libraries do not return <code>inheritedFrom</code>. A row marked as also present on an ancestor almost certainly predates the link.</li>
  <li><strong>Groups.</strong> Entra security and Microsoft 365 groups are expanded to transitive members with <code>-ExpandGroups</code>. SharePoint groups cannot be expanded through Graph at all; those rows carry a link to the site's group membership page.</li>
  <li><strong>Broad-audience grants</strong> (Anyone, Everyone, Everyone except external users, organization-wide links) matter more than any individual row: they are the ones that turn a folder into a tenant-wide or internet-wide resource.</li>
  <li><strong>Permissions answer "who could", not "who did."</strong> Only the audit log answers the second question, which is usually the one that matters in an incident. Run with <code>-IncludeAuditLog</code>, and <code>-IncludeFileAccess</code> for actual reads.</li>
  <li><strong>Absence of audit events is not absence of activity.</strong> Unified audit retention is 180 days on E3 and 365 days on E5. Anything older is gone.</li>
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

    # The console has to answer the question on its own. The usual case is one link and a
    # single-digit number of principals, and having to open a browser to read nine rows
    # would make the report the tool instead of the artifact.
    $ConsoleRowCap = 40
    Write-Host ''
    Write-Host '======================= WHO CAN REACH THIS =======================' -ForegroundColor Cyan
    foreach ($rep in $reports) {
        $kind = if ($rep.IsFolder) { 'folder' } else { 'file' }
        Write-Host ''
        Write-Host ("{0}  ({1})" -f $rep.ItemLabel, $kind) -ForegroundColor White
        Write-Host ("  {0}" -f [uri]::UnescapeDataString("$($rep.Item.webUrl)")) -ForegroundColor Gray

        $shown = @($rep.Rows | Sort-Object @{ e = { $RiskRank["$($_.Risk)"] }; Descending = $true }, Principal)
        $props = @('Risk', 'Access', 'Principal', 'PrincipalType', 'Grant')
        if (@($shown | Where-Object { $_.ViaGroup }).Count) { $props += 'ViaGroup' }
        $overflow = 0
        if ($shown.Count -gt $ConsoleRowCap) {
            $overflow = $shown.Count - $ConsoleRowCap
            $shown = @($shown | Select-Object -First $ConsoleRowCap)
        }
        if ($shown.Count) {
            $table = ($shown | Select-Object $props | Format-Table -AutoSize | Out-String -Width 240).Trim("`r", "`n")
            foreach ($line in ($table -split "`r?`n")) { Write-Host "  $line" }
        }
        else { Write-Host '  (no permissions returned)' -ForegroundColor Yellow }
        if ($overflow) { Write-Host "  ... and $overflow more row(s); see the HTML report." -ForegroundColor Gray }

        $extCount = @($rep.Rows | Where-Object { $_.External } | Select-Object -ExpandProperty Identifier -Unique).Count
        $linkCount = @($rep.Permissions | Where-Object { $_.link }).Count
        Write-Host ("  {0} access row(s), {1} external identit(ies), {2} sharing link(s)" -f @($rep.Rows).Count, $extCount, $linkCount)
        if (@($rep.Children).Count) { Write-Host "  $(@($rep.Children).Count) child item(s) carry permissions the folder does not." -ForegroundColor Yellow }
        if ($rep.AuditNote) { Write-Host "  $($rep.AuditNote)" -ForegroundColor Gray }
    }

    Write-Host ''
    if ($totalCritical) { Write-Host "Critical-risk grants: $totalCritical" -ForegroundColor Red }
    if (-not $AppOnly) { Write-Host 'Delegated run: complete only if this account owns the item. See the report banner.' -ForegroundColor Yellow }
    Write-Host "Report: $htmlPath"
    if ($ordered) { Write-Host "CSV:    $csvPath" }

    # Out-GridView sorts and filters for free, which is the fastest way to work a large or
    # batched result. It is Windows-only, so the guard-and-fall-back idiom matches
    # Get-TeamsChatMessages rather than assuming the cmdlet exists.
    if ($GridView) {
        $gridProps = 'Risk', 'Access', 'Principal', 'Identifier', 'PrincipalType', 'Grant',
                     'LinkAudience', 'ViaGroup', 'External', 'Expires', 'Item', 'Notes'
        $gridTitle = "Sharing link access - $($primary.ItemLabel) ($($ordered.Count) rows)"
        if ($reports.Count -gt 1) { $gridTitle = "Sharing link access - $($reports.Count) links, $($ordered.Count) rows" }
        if (Get-Command Out-GridView -ErrorAction SilentlyContinue) {
            $ordered | Select-Object $gridProps | Out-GridView -Title $gridTitle
        }
        else {
            Write-Host 'Out-GridView is unavailable on this host; showing a console table instead.' -ForegroundColor Yellow
            $ordered | Select-Object $gridProps | Format-Table -Wrap
        }

        # The timeline answers "who did", which the permission rows cannot. It gets its own
        # grid rather than being wedged into the same one: different shape, different question.
        $auditRows = @($reports | ForEach-Object { $_.Audit })
        if ($auditRows.Count) {
            if (Get-Command Out-GridView -ErrorAction SilentlyContinue) {
                $auditRows | Out-GridView -Title "Sharing timeline - $($auditRows.Count) events, last $AuditDays days"
            }
            else { $auditRows | Format-Table -Wrap }
        }
    }

    if (-not $NoOpen) { Start-Process $htmlPath }

    $ordered
}
