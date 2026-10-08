<#
.SYNOPSIS
    Reports on users with SMS-based MFA factors in Okta.

.DESCRIPTION
    Get-OktaSmsFactors queries the Okta Factors API to identify users enrolled in SMS-based
    multi-factor authentication. This is useful for:

        - SMS deprecation planning (Okta and security frameworks recommend phasing out SMS MFA)
        - MFA migration audits (identifying users to move to push or FIDO2, or carrying a verified
          phone number into another identity provider)
        - Compliance reporting on authentication methods

    Okta has no org-wide "list every factor" endpoint, so the script reads one factors list per
    user, and that endpoint is rate limited (300 or 600 calls a minute on most orgs). Runtime is
    therefore set by how many users you hand it, and the script offers three ways to keep that
    number small:

        -InputFile:     A CSV of pre-filtered users. Either a column of Okta user ids, or a column
                        of logins (the export of Okta's "MFA Enrollment by User" report, filtered
                        on the Phone authenticator type, is the natural source), which the script
                        resolves to ids.

        -GroupId:       Members of one or more Okta groups instead of every active user.

        -ProviderType:  Only accounts whose credential provider is in the list. An account that
                        federation created just in time (provider FEDERATION) under an enrollment
                        policy that forbids the phone authenticator can never hold an SMS factor,
                        so "-ProviderType OKTA, IMPORT, SOCIAL" skips the bulk of a federated
                        tenant without missing an enrollment.

    Without -InputFile or -GroupId the script enumerates every ACTIVE user through the Users API
    with pagination, as before.

    Every call reuses one HTTPS session (connection reuse roughly halves the time per call), and
    rate limiting is handled automatically: the script reads Okta's x-rate-limit-remaining and
    x-rate-limit-reset response headers, pauses when the window is nearly spent and retries a
    429 after the window resets.

    Authentication:
        Provide an Okta API token (SSWS token) via the -ApiToken parameter. Generate one in
        the Okta admin console under Security > API > Tokens. The token needs permissions to
        read users, groups and their factors (typically an Org Admin or Read-Only Admin role).
        Keep the token out of your command history by reading it into a variable first.

.PARAMETER OktaDomain
    Your Okta organization domain (e.g., "mycompany.okta.com"). Do not include the protocol
    prefix (https://) -- the script adds it automatically.

.PARAMETER ApiToken
    An Okta API token (SSWS token) with permissions to read users, groups and factors. Generate
    one in the Okta admin console under Security > API > Tokens.

.PARAMETER InputFile
    Optional path to a CSV file of pre-filtered users. Two shapes are accepted:

        - An "id" column of Okta user ids (optional columns "login", "email", "fullName",
          "status" and "provider" are used when present).
        - No "id" column but a login column named "login", "user login", "username" or "user"
          (case does not matter), as in the CSV export of the Okta "MFA Enrollment by User"
          report. Each login is resolved to its Okta user with one Users API call.

.PARAMETER GroupId
    One or more Okta group ids (00g...). The script lists the members of each group instead of
    enumerating every user in the org. Members of several groups are de-duplicated.

.PARAMETER ProviderType
    Only process accounts whose credential provider type is in this list: OKTA (Okta-managed
    password), ACTIVE_DIRECTORY, LDAP, FEDERATION (created by SAML or OIDC just-in-time
    provisioning), SOCIAL, IMPORT. Applies to users enumerated from the API and to CSV rows that
    carry a "provider" column or were resolved from a login; CSV rows with an id and no provider
    column are processed regardless, with one warning.

.PARAMETER IncludeInactive
    Include users whose status is not ACTIVE (SUSPENDED, DEPROVISIONED, STAGED, PROVISIONED,
    PASSWORD_EXPIRED, LOCKED_OUT, RECOVERY). By default only ACTIVE users are processed, in every
    input mode.

.PARAMETER UserLimit
    Maximum number of users to process. Defaults to 0 (no limit). Useful for testing or
    sampling a subset of users before running against the full population.

.EXAMPLE
    $token = (Get-Content -Path "$HOME\okta-token.txt" -Raw).Trim()
    .\Get-OktaSmsFactors.ps1 -OktaDomain "mycompany.okta.com" -ApiToken $token

    Queries all active users in the Okta tenant and reports on those with SMS factors.

.EXAMPLE
    .\Get-OktaSmsFactors.ps1 -OktaDomain "mycompany.okta.com" -ApiToken $token -GroupId "00gabc123" -ProviderType OKTA, IMPORT, SOCIAL

    Scans the members of one group, skipping accounts that federation created just in time. On a
    mostly federated tenant this turns a run of hours into one of minutes, with the same result.

.EXAMPLE
    .\Get-OktaSmsFactors.ps1 -OktaDomain "mycompany.okta.com" -ApiToken $token -InputFile ".\MFA_Enrollment_by_User.csv"

    Feeds the export of the "MFA Enrollment by User" report (Reports > Multifactor Authentication,
    filtered on authenticator type Phone). The report lists logins, which the script resolves to
    user ids before reading each user's factors.

.EXAMPLE
    .\Get-OktaSmsFactors.ps1 -OktaDomain "mycompany.okta.com" -ApiToken $token -UserLimit 50

    Processes only the first 50 users. Useful for testing.

.EXAMPLE
    . .\Get-OktaSmsFactors.ps1
    $results = Get-OktaSmsFactors -OktaDomain "mycompany.okta.com" -ApiToken $token -GroupId "00gabc123"
    $results | Where-Object { $_.SmsPhoneNumber } | Export-Csv ".\sms_users.csv" -NoTypeInformation

    Dot-source the script, capture results, and export users with SMS factors to CSV. The
    SmsFactorIds column carries the factor ids for a later cleanup once the numbers have moved.

.EXAMPLE
    $results = .\Get-OktaSmsFactors.ps1 -OktaDomain "mycompany.okta.com" -ApiToken $token
    $results | Where-Object { $_.MultipleSms } | Format-Table Login, FullName, SmsPhoneNumber

    Find users with multiple SMS factors enrolled.

.NOTES
    Author: Mike Crowley
    https://mikecrowley.us

    Okta API Reference:
        List Users          - https://developer.okta.com/docs/api/openapi/okta-management/management/tag/User/#tag/User/operation/listUsers
        List Group Members  - https://developer.okta.com/docs/api/openapi/okta-management/management/tag/Group/#tag/Group/operation/listGroupUsers
        List Factors        - https://developer.okta.com/docs/api/openapi/okta-management/management/tag/UserFactor/#tag/UserFactor/operation/listFactors

    Rate Limits:
        The Okta /api/v1/users/{id}/factors endpoint is subject to a per-minute rate limit that
        varies by org (300 a minute on some orgs, 600 or more on others). The script runs flat
        out until fewer than 5 requests remain in the window, sleeps until the window resets,
        and retries a 429 the same way. Requests in parallel do not raise the cap; they only
        spend it faster. Reducing the user list (see -GroupId, -ProviderType, -InputFile) is
        what shortens a run.

    Required Okta Role:
        Read-Only Admin, Org Admin, or a custom role with user, group and factor read permissions.

    CSV Format Examples (for -InputFile):
        id,login,fullName,email,status,provider
        00u1abc2def3ghi4j5k6,user1@mikecrowley.us,Jane Doe,jane@mikecrowley.us,ACTIVE,OKTA
        00u7lmn8opq9rst0u1v2,user2@mikecrowley.us,John Smith,john@mikecrowley.us,ACTIVE,FEDERATION

        login
        user1@mikecrowley.us
        user2@mikecrowley.us

    Output object properties:
        UserId         - Okta user ID
        Login          - Okta login (typically email)
        FullName       - User's display name
        Email          - User's email address
        Status         - Okta user status (ACTIVE, SUSPENDED, etc.)
        ProviderType   - Credential provider type (OKTA, FEDERATION, SOCIAL, IMPORT, ...) when known
        SmsPhoneNumber - Phone number(s) from SMS factor enrollment, E.164, "; " separated
        SmsFactorIds   - Factor id(s) of the SMS enrollment(s), "; " separated
        MultipleSms    - $true if the user has more than one SMS factor

.LINK
    https://developer.okta.com/docs/api/openapi/okta-management/management/tag/UserFactor/

.LINK
    https://help.okta.com/oie/en-us/content/topics/reports/mfa-enrollment-user-report.htm

.LINK
    https://github.com/Mike-Crowley/Public-Scripts
#>

[CmdletBinding()]
param(
    # OktaDomain and ApiToken are required when the script is run directly (checked at the bottom of the file);
    # they are not marked Mandatory so that dot-sourcing the script to load the function does not prompt for them
    [ValidateScript({
        if ($_ -match '^https?://') { throw "Provide the domain only (e.g., 'mycompany.okta.com'), not a full URL." }
        if ($_ -notmatch '\.okta\.com$|\.oktapreview\.com$|\.okta-emea\.com$') {
            Write-Warning "Domain '$_' does not end with .okta.com -- verify this is correct."
        }
        $true
    })]
    [string]$OktaDomain,

    [ValidateNotNullOrEmpty()]
    [string]$ApiToken,

    [ValidateScript({
        if (Test-Path $_ -PathType Leaf) { $true }
        else { throw "File not found: $_" }
    })]
    [string]$InputFile,

    [ValidatePattern('^00g[A-Za-z0-9]+$')]
    [string[]]$GroupId,

    [ValidateSet('OKTA', 'ACTIVE_DIRECTORY', 'LDAP', 'FEDERATION', 'SOCIAL', 'IMPORT')]
    [string[]]$ProviderType,

    [switch]$IncludeInactive,

    [ValidateRange(0, [int]::MaxValue)]
    [int]$UserLimit = 0
)

#region Main Function

function Get-OktaSmsFactors {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$OktaDomain,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$ApiToken,

        [string]$InputFile,

        [string[]]$GroupId,

        [ValidateSet('OKTA', 'ACTIVE_DIRECTORY', 'LDAP', 'FEDERATION', 'SOCIAL', 'IMPORT')]
        [string[]]$ProviderType,

        [switch]$IncludeInactive,

        [int]$UserLimit = 0
    )

    $script:OktaBaseUri = "https://$OktaDomain/api/v1"
    $script:OktaHeaders = @{
        "Authorization" = "SSWS $ApiToken"
        "Accept"        = "application/json"
        "Content-Type"  = "application/json"
    }
    $script:OktaSession = $null      # one HTTPS session for every call: connection reuse roughly halves the round trip
    $script:OktaCallCount = 0
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    # Collect users to process
    $users = [System.Collections.Generic.List[PSCustomObject]]::new()
    $providerUnknown = $false

    if ($InputFile) {
        Write-Host "Loading users from: $InputFile" -ForegroundColor Cyan
        $csvData = @(Import-Csv -Path $InputFile)
        if ($csvData.Count -eq 0) { throw "The CSV file has no rows." }

        $csvColumns = @($csvData[0].PSObject.Properties.Name)
        $idColumn = Get-CsvColumnName -Columns $csvColumns -Candidates 'id', 'userid', 'user id', 'okta id'
        $loginColumn = Get-CsvColumnName -Columns $csvColumns -Candidates 'login', 'user login', 'username', 'user name', 'user'
        $nameColumn = Get-CsvColumnName -Columns $csvColumns -Candidates 'fullName', 'full name', 'name', 'display name'
        $emailColumn = Get-CsvColumnName -Columns $csvColumns -Candidates 'email', 'email address'
        $statusColumn = Get-CsvColumnName -Columns $csvColumns -Candidates 'status', 'user status'
        $providerColumn = Get-CsvColumnName -Columns $csvColumns -Candidates 'provider', 'providerType', 'provider type'

        if (-not $idColumn -and -not $loginColumn) {
            throw "The CSV needs an 'id' column of Okta user ids or a login column ('login', 'user login', 'username' or 'user'). Found columns: $($csvColumns -join ', ')"
        }

        if ($idColumn) {
            foreach ($row in $csvData) {
                if ([string]::IsNullOrWhiteSpace($row.$idColumn)) { continue }
                $users.Add([pscustomobject]@{
                    id       = $row.$idColumn.Trim()
                    login    = if ($loginColumn) { $row.$loginColumn } else { $null }
                    fullName = if ($nameColumn) { $row.$nameColumn } else { $null }
                    email    = if ($emailColumn) { $row.$emailColumn } else { $null }
                    status   = if ($statusColumn) { $row.$statusColumn } else { $null }
                    provider = if ($providerColumn) { $row.$providerColumn } else { $null }
                })
            }
            if (-not $providerColumn) { $providerUnknown = $true }
            Write-Host "Loaded $($users.Count) user id(s) from CSV." -ForegroundColor Cyan
        }
        else {
            # A login-only file, such as the "MFA Enrollment by User" report export: resolve each login to its user
            $logins = @($csvData | ForEach-Object { $_.$loginColumn } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { $_.Trim() } | Select-Object -Unique)
            Write-Host "CSV carries $($logins.Count) login(s) and no id column; resolving each through the Users API..." -ForegroundColor Cyan
            $resolved = 0
            foreach ($login in $logins) {
                $resolved++
                Write-Progress -Activity "Resolving logins" -Status "$resolved of $($logins.Count) - $login" -PercentComplete ([math]::Round(($resolved / $logins.Count) * 100))
                $u = Invoke-OktaRequest -Uri "$script:OktaBaseUri/users/$([uri]::EscapeDataString($login))" -AllowNotFound
                if (-not $u) { Write-Warning "Login not found in Okta: $login"; continue }
                $users.Add((ConvertTo-UserRecord -User $u))
            }
            Write-Progress -Activity "Resolving logins" -Completed
            Write-Host "Resolved $($users.Count) of $($logins.Count) login(s)." -ForegroundColor Cyan
        }
    }
    elseif ($GroupId) {
        $seen = @{}
        foreach ($gid in $GroupId) {
            $group = Invoke-OktaRequest -Uri "$script:OktaBaseUri/groups/$gid"
            Write-Host "Listing members of group $gid ($($group.profile.name))..." -ForegroundColor Cyan
            $members = @(Invoke-OktaPagedRequest -Uri "$script:OktaBaseUri/groups/$gid/users?limit=200")
            $added = 0
            foreach ($u in $members) {
                if ($seen.ContainsKey($u.id)) { continue }
                $seen[$u.id] = $true
                $users.Add((ConvertTo-UserRecord -User $u))
                $added++
            }
            Write-Host "  $($members.Count) member(s), $added new to the list." -ForegroundColor Gray
        }
    }
    else {
        Write-Host "Querying users from Okta..." -ForegroundColor Cyan
        $usersUrl = "$script:OktaBaseUri/users?limit=200"
        if (-not $IncludeInactive) { $usersUrl += '&filter=' + [uri]::EscapeDataString('status eq "ACTIVE"') }
        $all = @(Invoke-OktaPagedRequest -Uri $usersUrl -StopAfter $UserLimit)
        foreach ($u in $all) { $users.Add((ConvertTo-UserRecord -User $u)) }
        Write-Host "Retrieved $($users.Count) user(s) from Okta." -ForegroundColor Cyan
    }

    # Status filter (ACTIVE only unless -IncludeInactive); rows that carry no status are kept
    if (-not $IncludeInactive) {
        $before = $users.Count
        $users = [System.Collections.Generic.List[PSCustomObject]]::new([PSCustomObject[]]@($users | Where-Object { -not $_.status -or $_.status -eq 'ACTIVE' }))
        if ($users.Count -lt $before) { Write-Host "Skipping $($before - $users.Count) user(s) whose status is not ACTIVE (use -IncludeInactive to keep them)." -ForegroundColor Gray }
    }

    # Provider filter: an account that federation created just in time cannot hold an SMS factor when its enrollment policy forbids the phone authenticator
    if ($ProviderType) {
        if ($providerUnknown) {
            Write-Warning "-ProviderType given, but the CSV has no 'provider' column; rows are processed regardless."
        }
        else {
            $before = $users.Count
            $users = [System.Collections.Generic.List[PSCustomObject]]::new([PSCustomObject[]]@($users | Where-Object { $_.provider -and ($ProviderType -contains $_.provider) }))
            Write-Host "Provider filter ($($ProviderType -join ', ')): $($users.Count) of $before user(s) kept." -ForegroundColor Cyan
        }
    }

    # Apply UserLimit
    if ($UserLimit -gt 0 -and $users.Count -gt $UserLimit) {
        $users = [System.Collections.Generic.List[PSCustomObject]]::new([PSCustomObject[]]@($users | Select-Object -First $UserLimit))
        Write-Host "Limited to $UserLimit user(s) per -UserLimit parameter." -ForegroundColor Yellow
    }

    if ($users.Count -eq 0) {
        Write-Host "No users to process." -ForegroundColor Yellow
        return
    }

    # Query factors for each user
    $results = [System.Collections.Generic.List[PSCustomObject]]::new()
    $counter = 0

    foreach ($user in $users) {
        $counter++
        Write-Progress -Activity "Querying Okta SMS Factors" `
            -Status "$counter of $($users.Count) - $($user.login)" `
            -PercentComplete ([math]::Round(($counter / $users.Count) * 100))

        try {
            $factors = @(Invoke-OktaRequest -Uri "$script:OktaBaseUri/users/$($user.id)/factors")
            $smsFactors = @($factors | Where-Object { $_.factorType -eq 'sms' })

            $results.Add([pscustomobject]@{
                UserId         = $user.id
                Login          = $user.login
                FullName       = $user.fullName
                Email          = $user.email
                Status         = $user.status
                ProviderType   = $user.provider
                SmsPhoneNumber = if ($smsFactors.Count) { @($smsFactors | ForEach-Object { $_.profile.phoneNumber }) -join '; ' } else { $null }
                SmsFactorIds   = if ($smsFactors.Count) { @($smsFactors | ForEach-Object { $_.id }) -join '; ' } else { $null }
                MultipleSms    = $smsFactors.Count -gt 1
            })
        }
        catch {
            Write-Warning "[$($user.id)] $($user.login): $($_.Exception.Message)"
            $results.Add([pscustomobject]@{
                UserId         = $user.id
                Login          = $user.login
                FullName       = $user.fullName
                Email          = $user.email
                Status         = $user.status
                ProviderType   = $user.provider
                SmsPhoneNumber = $null
                SmsFactorIds   = $null
                MultipleSms    = $false
            })
        }
    }

    Write-Progress -Activity "Querying Okta SMS Factors" -Completed
    $stopwatch.Stop()

    # Summary
    $smsUsers = @($results | Where-Object { $_.SmsPhoneNumber }).Count
    Write-Host ("`nComplete. {0} of {1} user(s) have SMS factors enrolled. {2} API call(s) in {3:n1} minutes." -f $smsUsers, $results.Count, $script:OktaCallCount, $stopwatch.Elapsed.TotalMinutes) -ForegroundColor Cyan

    # Pipeline output
    $results
}

#endregion

#region Helper Functions

function ConvertTo-UserRecord {
    # The fields the factor loop needs, from an Okta user object
    param([Parameter(Mandatory)] $User)
    [pscustomobject]@{
        id       = $User.id
        login    = $User.profile.login
        fullName = ("$($User.profile.firstName) $($User.profile.lastName)").Trim()
        email    = $User.profile.email
        status   = $User.status
        provider = $User.credentials.provider.type
    }
}

function Get-CsvColumnName {
    # The first CSV column whose name matches one of the candidates, ignoring case and spaces; $null when none does
    param(
        [Parameter(Mandatory)] [string[]]$Columns,
        [Parameter(Mandatory)] [string[]]$Candidates
    )
    foreach ($candidate in $Candidates) {
        $wanted = ($candidate -replace '\s', '').ToLowerInvariant()
        foreach ($column in $Columns) {
            if ((($column -replace '\s', '').ToLowerInvariant()) -eq $wanted) { return $column }
        }
    }
    return $null
}

function Invoke-OktaRequest {
    <#
        One GET against the Okta API on the shared session: reads the rate-limit headers, sleeps out a
        nearly spent window, retries a 429 after the window resets, and returns the parsed body.
        Returns $null for a 404 when -AllowNotFound is set; any other error status throws.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$Uri,
        [switch]$AllowNotFound,
        [ref]$ResponseHeaders
    )

    for ($attempt = 1; $attempt -le 5; $attempt++) {
        $params = @{
            Uri                     = $Uri
            Headers                 = $script:OktaHeaders
            ResponseHeadersVariable = 'headers'
            StatusCodeVariable      = 'status'
            SkipHttpErrorCheck      = $true
            ErrorAction             = 'Stop'
        }
        if ($script:OktaSession) { $params.WebSession = $script:OktaSession } else { $params.SessionVariable = 'newSession' }

        $body = Invoke-RestMethod @params
        $script:OktaCallCount++
        if (-not $script:OktaSession -and $newSession) { $script:OktaSession = $newSession }

        if ($status -eq 429) {
            Wait-OktaRateLimitReset -ResponseHeaders $headers -Reason "429 received"
            continue
        }

        Invoke-OktaRateLimitCheck -ResponseHeaders $headers

        if ($status -eq 404 -and $AllowNotFound) { return $null }
        if ($status -lt 200 -or $status -ge 300) {
            $summary = if ($body.errorSummary) { $body.errorSummary } else { "HTTP $status" }
            throw "Okta answered $status for $Uri : $summary"
        }

        if ($ResponseHeaders) { $ResponseHeaders.Value = $headers }
        return $body
    }
    throw "Okta kept answering 429 for $Uri"
}

function Invoke-OktaPagedRequest {
    # Follows the Link rel="next" header until the list ends, or until -StopAfter items have been collected (0 = no limit)
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$Uri,
        [int]$StopAfter = 0
    )
    $items = [System.Collections.Generic.List[object]]::new()
    $next = $Uri
    while ($next) {
        $headers = $null
        $page = @(Invoke-OktaRequest -Uri $next -ResponseHeaders ([ref]$headers))
        foreach ($item in $page) { $items.Add($item) }
        Write-Host "  Retrieved so far: $($items.Count)" -ForegroundColor Gray
        $next = $null
        if ($headers -and $headers['Link']) {
            $linkHeader = $headers['Link'] -join ', '
            if ($linkHeader -match '<([^>]+)>;\s*rel="next"') { $next = $Matches[1] }
        }
        if ($page.Count -eq 0) { $next = $null }
        if ($StopAfter -gt 0 -and $items.Count -ge $StopAfter) { break }
    }
    return $items.ToArray()
}

function Invoke-OktaRateLimitCheck {
    # Prints the remaining budget only when it is getting low, and sleeps out the window when 5 or fewer calls remain
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $ResponseHeaders
    )

    $remaining = $ResponseHeaders['x-rate-limit-remaining']
    $reset = $ResponseHeaders['x-rate-limit-reset']

    if ($remaining -and $reset) {
        $rateLimitRemaining = [int]$remaining[0]
        if ($rateLimitRemaining -lt 50) {
            Write-Host "  Rate Limit: $rateLimitRemaining remaining" -ForegroundColor $(if ($rateLimitRemaining -lt 10) { "Red" } else { "Yellow" })
        }
        else {
            Write-Verbose "Rate Limit: $rateLimitRemaining remaining"
        }

        if ($rateLimitRemaining -le 5) {
            Wait-OktaRateLimitReset -ResponseHeaders $ResponseHeaders -Reason "Rate limit approaching"
        }
    }
}

function Wait-OktaRateLimitReset {
    # Sleeps until the window named by x-rate-limit-reset has passed (a 60-second fallback when the header is missing)
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] $ResponseHeaders,
        [string]$Reason = "Rate limit"
    )
    $reset = $ResponseHeaders['x-rate-limit-reset']
    $waitTime = 60
    if ($reset) {
        $resetTime = [DateTimeOffset]::FromUnixTimeSeconds([int64]$reset[0]).LocalDateTime
        $waitTime = ($resetTime - (Get-Date)).TotalSeconds
    }
    if ($waitTime -gt 0) {
        Write-Warning "$Reason. Waiting $([math]::Ceiling($waitTime)) seconds..."
        Start-Sleep -Seconds ([math]::Ceiling($waitTime) + 1)
    }
}

#endregion

# Direct invocation support
if ($MyInvocation.InvocationName -ne '.') {
    if (-not $OktaDomain -or -not $ApiToken) {
        throw "OktaDomain and ApiToken are required. Run the script with both, or dot-source it (. .\Get-OktaSmsFactors.ps1) and call Get-OktaSmsFactors."
    }
    $scriptParams = @{}
    foreach ($key in $PSBoundParameters.Keys) {
        $scriptParams[$key] = $PSBoundParameters[$key]
    }
    Get-OktaSmsFactors @scriptParams
}
