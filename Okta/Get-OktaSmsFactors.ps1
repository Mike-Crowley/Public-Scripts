<#
.SYNOPSIS
    Reports on users with SMS-based MFA factors in Okta.

.DESCRIPTION
    Get-OktaSmsFactors queries the Okta Factors API to identify users enrolled in SMS-based
    multi-factor authentication. This is useful for:

        - SMS deprecation planning (Okta and security frameworks recommend phasing out SMS MFA)
        - MFA migration audits (identifying users to move to push or FIDO2)
        - Compliance reporting on authentication methods

    The script supports two modes of operation:

        With -InputFile:    Reads a CSV of pre-filtered Okta user IDs and queries their factors.
                            Useful when you have a large tenant and only want to check a subset
                            of users (e.g., those already known to have phone-type authenticators
                            from an Okta admin portal export).

        Without -InputFile: Queries all active users from the Okta Users API with pagination,
                            then checks each user's enrolled factors. For very large tenants,
                            consider using -InputFile or -UserLimit to control scope.

    Rate limiting is handled automatically. The script reads Okta's x-rate-limit-remaining and
    x-rate-limit-reset response headers and pauses when approaching the limit.

    Authentication:
        Provide an Okta API token (SSWS token) via the -ApiToken parameter. Generate one in
        the Okta admin console under Security > API > Tokens. The token needs permissions to
        read users and their factors (typically an Org Admin or Read-Only Admin role).

.PARAMETER OktaDomain
    Your Okta organization domain (e.g., "mycompany.okta.com"). Do not include the protocol
    prefix (https://) -- the script adds it automatically.

.PARAMETER ApiToken
    An Okta API token (SSWS token) with permissions to read users and factors. Generate one
    in the Okta admin console under Security > API > Tokens.

.PARAMETER InputFile
    Optional path to a CSV file containing pre-filtered Okta users. The CSV must contain at
    least an "id" column with Okta user IDs. Optional columns "login", "email", "fullName",
    and "status" will be used if present, otherwise the script queries each user's profile.

    This is useful for large tenants where enumerating all users via API is impractical. You
    can export users from the Okta admin portal and filter them before passing to this script.

.PARAMETER UserLimit
    Maximum number of users to process. Defaults to 0 (no limit). Useful for testing or
    sampling a subset of users before running against the full population.

.EXAMPLE
    .\Get-OktaSmsFactors.ps1 -OktaDomain "mycompany.okta.com" -ApiToken "00abc123..."

    Queries all active users in the Okta tenant and reports on those with SMS factors.

.EXAMPLE
    .\Get-OktaSmsFactors.ps1 -OktaDomain "mycompany.okta.com" -ApiToken "00abc123..." -UserLimit 50

    Processes only the first 50 users. Useful for testing.

.EXAMPLE
    .\Get-OktaSmsFactors.ps1 -OktaDomain "mycompany.okta.com" -ApiToken "00abc123..." -InputFile ".\phone_users.csv"

    Reads user IDs from a CSV file and queries their SMS factors.

.EXAMPLE
    . .\Get-OktaSmsFactors.ps1
    $results = Get-OktaSmsFactors -OktaDomain "mycompany.okta.com" -ApiToken "00abc123..."
    $results | Where-Object { $_.SmsPhoneNumber } | Export-Csv ".\sms_users.csv" -NoTypeInformation

    Dot-source the script, capture results, and export users with SMS factors to CSV.

.EXAMPLE
    $results = .\Get-OktaSmsFactors.ps1 -OktaDomain "mycompany.okta.com" -ApiToken "00abc123..."
    $results | Where-Object { $_.MultipleSms } | Format-Table Login, FullName, SmsPhoneNumber

    Find users with multiple SMS factors enrolled.

.NOTES
    Author: Mike Crowley

    Okta API Reference:
        List Users   - https://developer.okta.com/docs/api/openapi/okta-management/management/tag/User/#tag/User/operation/listUsers
        List Factors - https://developer.okta.com/docs/api/openapi/okta-management/management/tag/UserFactor/#tag/UserFactor/operation/listFactors

    Rate Limits:
        The Okta /api/v1/users/{id}/factors endpoint is subject to rate limiting.
        This script monitors the x-rate-limit-remaining header and automatically
        pauses when fewer than 5 requests remain, resuming after the rate limit
        window resets.

    Required Okta Role:
        Read-Only Admin, Org Admin, or a custom role with user and factor read permissions.

    CSV Format Example (for -InputFile):
        id,login,fullName,email,status
        00u1abc2def3ghi4j5k6,user1@mikecrowley.us,Jane Doe,jane@mikecrowley.us,ACTIVE
        00u7lmn8opq9rst0u1v2,user2@mikecrowley.us,John Smith,john@mikecrowley.us,ACTIVE

    Output object properties:
        UserId         - Okta user ID
        Login          - Okta login (typically email)
        FullName       - User's display name
        Email          - User's email address
        Status         - Okta user status (ACTIVE, SUSPENDED, etc.)
        SmsPhoneNumber - Phone number(s) from SMS factor enrollment
        MultipleSms    - $true if the user has more than one SMS factor

.LINK
    https://developer.okta.com/docs/api/openapi/okta-management/management/tag/UserFactor/

.LINK
    https://github.com/Mike-Crowley/Public-Scripts
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidateScript({
        if ($_ -match '^https?://') { throw "Provide the domain only (e.g., 'mycompany.okta.com'), not a full URL." }
        if ($_ -notmatch '\.okta\.com$|\.oktapreview\.com$|\.okta-emea\.com$') {
            Write-Warning "Domain '$_' does not end with .okta.com -- verify this is correct."
        }
        $true
    })]
    [string]$OktaDomain,

    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$ApiToken,

    [ValidateScript({
        if (Test-Path $_ -PathType Leaf) { $true }
        else { throw "File not found: $_" }
    })]
    [string]$InputFile,

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

        [int]$UserLimit = 0
    )

    $baseUri = "https://$OktaDomain/api/v1"
    $headers = @{
        "Authorization" = "SSWS $ApiToken"
        "Accept"        = "application/json"
        "Content-Type"  = "application/json"
    }

    # Collect users to process
    if ($InputFile) {
        Write-Host "Loading users from: $InputFile" -ForegroundColor Cyan
        $csvData = Import-Csv -Path $InputFile

        $csvColumns = $csvData[0].PSObject.Properties.Name
        if ('id' -notin $csvColumns) {
            throw "Required CSV column 'id' not found. Found columns: $($csvColumns -join ', ')"
        }

        $users = foreach ($row in $csvData) {
            [pscustomobject]@{
                id       = $row.id
                login    = if ($row.PSObject.Properties['login']) { $row.login } else { $null }
                fullName = if ($row.PSObject.Properties['fullName']) { $row.fullName } else { $null }
                email    = if ($row.PSObject.Properties['email']) { $row.email } else { $null }
                status   = if ($row.PSObject.Properties['status']) { $row.status } else { $null }
            }
        }

        Write-Host "Loaded $($users.Count) user(s) from CSV." -ForegroundColor Cyan
    }
    else {
        Write-Host "Querying users from Okta..." -ForegroundColor Cyan
        $users = [System.Collections.Generic.List[PSCustomObject]]::new()
        $usersUrl = "$baseUri/users?limit=200&filter=status eq `"ACTIVE`""

        while ($usersUrl) {
            try {
                $response = Invoke-RestMethod -Uri $usersUrl -Headers $headers -ResponseHeadersVariable responseHeaders -ErrorAction Stop
            }
            catch {
                throw "Failed to query Okta users: $($_.Exception.Message)"
            }

            foreach ($u in $response) {
                $users.Add([pscustomobject]@{
                    id       = $u.id
                    login    = $u.profile.login
                    fullName = "$($u.profile.firstName) $($u.profile.lastName)"
                    email    = $u.profile.email
                    status   = $u.status
                })
            }

            Write-Host "  Users retrieved so far: $($users.Count)" -ForegroundColor Gray

            # Handle rate limiting on user enumeration
            Invoke-OktaRateLimitCheck -ResponseHeaders $responseHeaders

            # Pagination via Link header
            $usersUrl = $null
            if ($responseHeaders['Link']) {
                $linkHeader = $responseHeaders['Link'] -join ', '
                if ($linkHeader -match '<([^>]+)>;\s*rel="next"') {
                    $usersUrl = $Matches[1]
                }
            }

            if ($UserLimit -gt 0 -and $users.Count -ge $UserLimit) {
                $users = [System.Collections.Generic.List[PSCustomObject]]::new(
                    [PSCustomObject[]]($users | Select-Object -First $UserLimit)
                )
                break
            }
        }

        Write-Host "Retrieved $($users.Count) user(s) from Okta." -ForegroundColor Cyan
    }

    # Apply UserLimit to CSV input as well
    if ($UserLimit -gt 0 -and $users.Count -gt $UserLimit) {
        $users = $users | Select-Object -First $UserLimit
        Write-Host "Limited to $UserLimit user(s) per -UserLimit parameter." -ForegroundColor Yellow
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
            $factorsUrl = "$baseUri/users/$($user.id)/factors"
            $factorsResponse = Invoke-RestMethod -Uri $factorsUrl -Headers $headers -ResponseHeadersVariable factorsHeaders -ErrorAction Stop
            $smsFactors = $factorsResponse | Where-Object factorType -eq "sms"

            Invoke-OktaRateLimitCheck -ResponseHeaders $factorsHeaders

            $result = [pscustomobject]@{
                UserId         = $user.id
                Login          = $user.login
                FullName       = $user.fullName
                Email          = $user.email
                Status         = $user.status
                SmsPhoneNumber = if ($smsFactors) { ($smsFactors.profile.phoneNumber) -join '; ' } else { $null }
                MultipleSms    = ($smsFactors | Measure-Object).Count -gt 1
            }

            $results.Add($result)
        }
        catch {
            Write-Warning "[$($user.id)] $($user.login): $($_.Exception.Message)"
            $results.Add([pscustomobject]@{
                UserId         = $user.id
                Login          = $user.login
                FullName       = $user.fullName
                Email          = $user.email
                Status         = $user.status
                SmsPhoneNumber = $null
                MultipleSms    = $false
            })
        }
    }

    Write-Progress -Activity "Querying Okta SMS Factors" -Completed

    # Summary
    $smsUsers = ($results | Where-Object { $_.SmsPhoneNumber }).Count
    Write-Host "`nComplete. $smsUsers of $($results.Count) user(s) have SMS factors enrolled." -ForegroundColor Cyan

    # Pipeline output
    $results
}

#endregion

#region Helper Functions

function Invoke-OktaRateLimitCheck {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $ResponseHeaders
    )

    $remaining = $ResponseHeaders['x-rate-limit-remaining']
    $reset = $ResponseHeaders['x-rate-limit-reset']

    if ($remaining -and $reset) {
        $rateLimitRemaining = [int]$remaining[0]
        Write-Host "  Rate Limit: $rateLimitRemaining remaining" -ForegroundColor $(
            if ($rateLimitRemaining -lt 10) { "Red" }
            elseif ($rateLimitRemaining -lt 50) { "Yellow" }
            else { "Green" }
        )

        if ($rateLimitRemaining -le 5) {
            $resetTime = [DateTimeOffset]::FromUnixTimeSeconds([int64]$reset[0]).LocalDateTime
            $waitTime = ($resetTime - (Get-Date)).TotalSeconds
            if ($waitTime -gt 0) {
                Write-Warning "Rate limit approaching. Waiting $([math]::Ceiling($waitTime)) seconds..."
                Start-Sleep -Seconds ([math]::Ceiling($waitTime) + 1)
            }
        }
    }
}

#endregion

# Direct invocation support
if ($MyInvocation.InvocationName -ne '.') {
    $scriptParams = @{}
    foreach ($key in $PSBoundParameters.Keys) {
        $scriptParams[$key] = $PSBoundParameters[$key]
    }
    if ($scriptParams.Count -gt 0) {
        Get-OktaSmsFactors @scriptParams
    }
}
