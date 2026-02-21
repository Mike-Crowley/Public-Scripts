<#
.SYNOPSIS
    Imports hardware OATH tokens into Microsoft Entra ID and activates them for assigned users.

.DESCRIPTION
    Import-EntraHardwareOathToken reads a CSV file containing hardware OATH token data
    (serial numbers, hex-encoded secret keys, user principal names) and performs the following
    for each row:

      1. Looks up the user in Entra ID by UPN
      2. Converts the hex-encoded secret key to Base32 (RFC 4648)
      3. Creates the hardware OATH device in the directory and assigns it to the user
      4. Generates a TOTP verification code (RFC 6238) from the secret key
      5. Activates the token for the user

    Activation is normally an interactive step where a user reads the code off their physical
    token. This script automates it by computing the TOTP value from the seed, enabling
    seamless bulk migrations between MFA providers without end-user involvement.

    Secret Key Format:
        The Graph API requires secret keys in Base32 encoding, but most hardware token
        vendors (Fortinet, FEITIAN, TOKEN2, Deepnet, etc.) provide them as hexadecimal
        strings. This script handles the hex-to-Base32 conversion automatically. Provide
        the raw hex values from your vendor in the SecretKey CSV column.

        Note: Microsoft's documentation shows example secret keys that appear to contain
        characters outside the Base32 alphabet (0, 1, 8, 9, lowercase). These examples are
        misleading. The API strictly enforces Base32 (A-Z, 2-7 only) and will reject keys
        containing invalid characters.

    This script uses the Microsoft Graph beta API endpoints:
        POST /directory/authenticationMethodDevices/hardwareOathDevices
        POST /users/{id}/authentication/hardwareOathMethods/assignAndActivateBySerialNumber

    Connect to Microsoft Graph prior to running this script:
        Connect-MgGraph -Scopes "Policy.ReadWrite.AuthenticationMethod", "UserAuthenticationMethod.ReadWrite.All", "User.Read.All"

    The calling user also needs the Authentication Administrator or Privileged Authentication
    Administrator Entra role.

    Note: These Graph API endpoints are currently in beta and subject to change.

.PARAMETER InputFile
    Path to a CSV file containing token data.

    Required columns:
        SerialNumber          - The token's serial number (e.g., "FTK-000001")
        SecretKey             - The token's shared secret as a hexadecimal string
        UserPrincipalName     - The user's UPN in Entra ID (e.g., "user@contoso.com")

    Optional columns (override parameter defaults per-row):
        Manufacturer          - Token manufacturer name
        Model                 - Token model name
        TimeIntervalInSeconds - TOTP refresh interval (30 or 60)
        HashFunction          - "hmacsha1" or "hmacsha256"
        DisplayName           - Custom display name for the token in Entra

.PARAMETER Manufacturer
    Default manufacturer name applied when the CSV row lacks a Manufacturer value.
    Defaults to "Hardware Token".

.PARAMETER Model
    Default model name applied when the CSV row lacks a Model value.
    Defaults to "TOTP".

.PARAMETER TimeIntervalInSeconds
    TOTP refresh interval in seconds. Must be 30 or 60. Defaults to 60.
    Most hardware tokens (FortiToken, FEITIAN) use 60; most authenticator apps use 30.
    Can be overridden per-row if the CSV contains a TimeIntervalInSeconds column.

.PARAMETER HashFunction
    Hash algorithm for TOTP generation. Must be "hmacsha1" or "hmacsha256".
    Defaults to "hmacsha1". Can be overridden per-row via CSV column.

.PARAMETER Digits
    Number of digits in the TOTP code. Must be 6 or 8. Defaults to 6.

.EXAMPLE
    .\Import-EntraHardwareOathToken.ps1 -InputFile "C:\tokens\token_inventory.csv"

    Imports all tokens from the CSV using default settings (60-second interval, hmacsha1, 6 digits).

.EXAMPLE
    .\Import-EntraHardwareOathToken.ps1 -InputFile "C:\tokens\tokens.csv" -Manufacturer "FEITIAN" -Model "OTP c200" -TimeIntervalInSeconds 60

    Imports FEITIAN hardware tokens with a 60-second TOTP interval.

.EXAMPLE
    . .\Import-EntraHardwareOathToken.ps1
    Import-EntraHardwareOathToken -InputFile "C:\tokens\tokens.csv" -WhatIf

    Dot-source then call with -WhatIf to preview what would be created without making changes.

.EXAMPLE
    $results = .\Import-EntraHardwareOathToken.ps1 -InputFile "C:\tokens\tokens.csv"
    $results | Where-Object { $_.Error } | Export-Csv "C:\tokens\failures.csv" -NoTypeInformation

    Capture results and export failures for review.

.NOTES
    Author: Mike Crowley
    Requires: Microsoft.Graph.Authentication module
    Dependencies: ..\Utilities\Get-TOTP.ps1 (Get-TOTP and Convert-HexToBase32 functions)
    API: Microsoft Graph Beta (subject to change)

    Required Entra Role:
        Authentication Administrator  -or-  Privileged Authentication Administrator

    Required Graph Permissions:
        Policy.ReadWrite.AuthenticationMethod    - Create hardware OATH devices
        UserAuthenticationMethod.ReadWrite.All    - Assign and activate tokens for users
        User.Read.All                            - Look up users by UPN

    CSV Format Example:
        SerialNumber,SecretKey,UserPrincipalName,Manufacturer,Model
        FTK-000001,31323334353637383930,user1@contoso.com,FEITIAN,OTP c200
        FTK-000002,48656C6C6F21,user2@contoso.com,,

    Output object properties:
        SerialNumber      - Token serial number
        UserPrincipalName - Target user
        Created           - $true if the device was created and assigned in Entra
        Activated         - $true if the token was activated with a TOTP code
        Error             - Error message, or $null on success

.LINK
    https://learn.microsoft.com/en-us/graph/api/authenticationmethoddevice-post-hardwareoathdevices

.LINK
    https://learn.microsoft.com/en-us/entra/identity/authentication/how-to-mfa-manage-oath-tokens

.LINK
    https://github.com/Mike-Crowley/Public-Scripts
#>

#Requires -Modules Microsoft.Graph.Authentication

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [ValidateScript({
        if (Test-Path $_ -PathType Leaf) { $true }
        else { throw "File not found: $_" }
    })]
    [string]$InputFile,

    [string]$Manufacturer = "Hardware Token",

    [string]$Model = "TOTP",

    [ValidateSet(30, 60)]
    [int]$TimeIntervalInSeconds = 60,

    [ValidateSet("hmacsha1", "hmacsha256")]
    [string]$HashFunction = "hmacsha1",

    [ValidateSet(6, 8)]
    [int]$Digits = 6
)

# Load TOTP utility functions (Get-TOTP, Convert-HexToBase32)
. "$PSScriptRoot\..\Utilities\Get-TOTP.ps1"

#region Main Function

function Import-EntraHardwareOathToken {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory)]
        [ValidateScript({
            if (Test-Path $_ -PathType Leaf) { $true }
            else { throw "File not found: $_" }
        })]
        [string]$InputFile,

        [string]$Manufacturer = "Hardware Token",

        [string]$Model = "TOTP",

        [ValidateSet(30, 60)]
        [int]$TimeIntervalInSeconds = 60,

        [ValidateSet("hmacsha1", "hmacsha256")]
        [string]$HashFunction = "hmacsha1",

        [ValidateSet(6, 8)]
        [int]$Digits = 6
    )

    # Pre-flight: verify Graph context
    $context = Get-MgContext
    if ($null -eq $context) {
        throw "No Microsoft Graph context found. Run Connect-MgGraph first."
    }

    $requiredScopes = @('Policy.ReadWrite.AuthenticationMethod', 'UserAuthenticationMethod.ReadWrite.All')
    if ($context.Scopes) {
        foreach ($scope in $requiredScopes) {
            if ($scope -notin $context.Scopes) {
                Write-Warning "Permission scope '$scope' may not be consented. You may encounter errors."
            }
        }
    }

    # Import CSV and validate required columns
    $tokenData = Import-Csv -Path $InputFile
    if ($tokenData.Count -eq 0) {
        throw "CSV file is empty: $InputFile"
    }

    $csvColumns = $tokenData[0].PSObject.Properties.Name
    foreach ($col in @('SerialNumber', 'SecretKey', 'UserPrincipalName')) {
        if ($col -notin $csvColumns) {
            throw "Required CSV column '$col' not found. Found columns: $($csvColumns -join ', ')"
        }
    }

    Write-Host "Loaded $($tokenData.Count) token(s) from: $InputFile" -ForegroundColor Cyan

    # Process each token
    $results = [System.Collections.Generic.List[PSCustomObject]]::new()
    $errorUsers = [System.Collections.Generic.List[PSCustomObject]]::new()
    $counter = 0

    foreach ($token in $tokenData) {
        $counter++
        Write-Progress -Activity "Importing Hardware OATH Tokens" `
            -Status "$counter of $($tokenData.Count) - $($token.UserPrincipalName)" `
            -PercentComplete ([math]::Round(($counter / $tokenData.Count) * 100))

        # Per-row overrides (CSV values take precedence over parameters)
        $rowManufacturer = if ($token.PSObject.Properties['Manufacturer'] -and $token.Manufacturer) { $token.Manufacturer } else { $Manufacturer }
        $rowModel = if ($token.PSObject.Properties['Model'] -and $token.Model) { $token.Model } else { $Model }
        $rowInterval = if ($token.PSObject.Properties['TimeIntervalInSeconds'] -and $token.TimeIntervalInSeconds) { [int]$token.TimeIntervalInSeconds } else { $TimeIntervalInSeconds }
        $rowHashFunction = if ($token.PSObject.Properties['HashFunction'] -and $token.HashFunction) { $token.HashFunction } else { $HashFunction }
        $rowDisplayName = if ($token.PSObject.Properties['DisplayName'] -and $token.DisplayName) { $token.DisplayName } else { $null }

        $result = [pscustomobject]@{
            SerialNumber      = $token.SerialNumber
            UserPrincipalName = $token.UserPrincipalName
            Created           = $false
            Activated         = $false
            Error             = $null
        }

        try {
            # Look up user
            Write-Host "Processing $($token.UserPrincipalName)..." -ForegroundColor Cyan
            $userRequest = Invoke-MgGraphRequest -Uri "v1.0/users/$($token.UserPrincipalName)" -OutputType PSObject
            $userId = $userRequest.id

            if ($null -eq $userId) {
                throw "User not found: $($token.UserPrincipalName)"
            }

            if ($PSCmdlet.ShouldProcess("$($token.SerialNumber) -> $($token.UserPrincipalName)", "Create, assign, and activate hardware OATH token")) {

                # Step 1: Create device and assign to user
                $serialSuffix = $token.SerialNumber
                if ($serialSuffix.Length -gt 7) {
                    $serialSuffix = $serialSuffix.Substring($serialSuffix.Length - 7)
                }

                $createBody = @{
                    displayName           = if ($rowDisplayName) { $rowDisplayName } else { "$rowManufacturer $serialSuffix" }
                    serialNumber          = $token.SerialNumber
                    manufacturer          = $rowManufacturer
                    model                 = $rowModel
                    secretKey             = $token.SecretKey | Convert-HexToBase32
                    timeIntervalInSeconds = $rowInterval
                    hashFunction          = $rowHashFunction
                    assignTo              = @{ id = $userId }
                }

                try {
                    Invoke-MgGraphRequest -Uri "beta/directory/authenticationMethodDevices/hardwareOathDevices" -Body $createBody -Method POST | Out-Null
                    $result.Created = $true
                }
                catch {
                    $errorDetail = try { (($_.Exception.Message -split "`n")[-1] | ConvertFrom-Json).error.message } catch { $_.Exception.Message }
                    throw "Create/assign failed: $errorDetail"
                }

                # Step 2: Generate TOTP and activate
                $activateBody = @{
                    verificationCode = Get-TOTP -SecretHex $token.SecretKey -TimeStep $rowInterval -Digits $Digits -HashFunction $rowHashFunction
                    serialNumber     = $token.SerialNumber
                    displayName      = "$($userRequest.displayName)'s $rowModel"
                }

                try {
                    Invoke-MgGraphRequest -Uri "beta/users/$userId/authentication/hardwareOathMethods/assignAndActivateBySerialNumber" -Body $activateBody -Method POST | Out-Null
                    $result.Activated = $true
                }
                catch {
                    $errorDetail = try { (($_.Exception.Message -split "`n")[-1] | ConvertFrom-Json).error.message } catch { $_.Exception.Message }
                    throw "Activation failed: $errorDetail"
                }
            }
        }
        catch {
            $result.Error = $_.Exception.Message
            $errorUsers.Add($result)
            Write-Warning "[$($token.SerialNumber)] $($token.UserPrincipalName): $($_.Exception.Message)"
        }

        $results.Add($result)
    }

    Write-Progress -Activity "Importing Hardware OATH Tokens" -Completed

    # Summary
    $succeeded = ($results | Where-Object { $_.Activated }).Count
    $failed = $errorUsers.Count

    Write-Host "`nImport complete." -ForegroundColor Cyan
    if ($failed -gt 0) {
        Write-Warning "$failed of $($results.Count) token(s) failed:"
        $errorUsers | ForEach-Object { Write-Warning "  $($_.UserPrincipalName)" }
    }
    else {
        Write-Host "All $($results.Count) token(s) succeeded." -ForegroundColor Green
    }

    # Pipeline output
    $results
}

#endregion

# Direct invocation support
if ($MyInvocation.InvocationName -ne '.') {
    $scriptParams = @{}
    foreach ($key in $PSBoundParameters.Keys) {
        $scriptParams[$key] = $PSBoundParameters[$key]
    }
    if ($scriptParams.Count -gt 0) {
        Import-EntraHardwareOathToken @scriptParams
    }
}
