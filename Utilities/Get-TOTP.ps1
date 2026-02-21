<#
.SYNOPSIS
    Generates Time-based One-Time Passwords (TOTP) and converts hex strings to Base32.

.DESCRIPTION
    This script provides two standalone functions with no external dependencies:

    Get-TOTP
        Implements RFC 6238 (TOTP) to generate time-based one-time passwords from a
        hex-encoded shared secret. Uses .NET's System.Security.Cryptography.HMACSHA1
        (or HMACSHA256) for the HMAC computation and RFC 4226 dynamic truncation to
        produce the final code. Useful for verifying hardware token output or automating
        token activation against APIs that require a verification code.

    Convert-HexToBase32
        Converts hexadecimal strings to Base32 encoding per RFC 4648 (alphabet: A-Z, 2-7).
        Many services (Microsoft Entra ID, Google Authenticator, etc.) require shared
        secrets in Base32 format, but most hardware token vendors provide them as hex.
        This function bridges that gap.

    Common use cases:
        - Verifying that a hardware token is producing the expected codes
        - Automating OATH token activation in Microsoft Entra ID (see
          Import-EntraHardwareOathToken.ps1 in the MicrosoftGraph folder)
        - Converting vendor-supplied hex secrets to Base32 for manual registration
          in authenticator apps or identity provider portals

.PARAMETER SecretHex
    The shared secret as a hexadecimal string (characters 0-9, A-F). This is typically
    provided by the hardware token vendor or generated during token provisioning.
    Example: "31323334353637383930" (the ASCII hex encoding of "1234567890").

.PARAMETER TimeStep
    The time interval in seconds for TOTP generation. Defaults to 30.
        30 - Standard for most software authenticator apps (Google Authenticator, Authy)
        60 - Used by many hardware tokens (FortiToken 200B, FEITIAN OTP c200)

.PARAMETER Digits
    Number of digits in the generated OTP. Must be 6 or 8. Defaults to 6.
    Most implementations use 6 digits.

.PARAMETER HashFunction
    Hash algorithm for HMAC computation. Must be "hmacsha1" or "hmacsha256".
    Defaults to "hmacsha1" (most widely supported across hardware tokens and services).

.EXAMPLE
    . .\Get-TOTP.ps1
    Get-TOTP -SecretHex "31323334353637383930"

    Generates a 6-digit TOTP code using HMAC-SHA1 with a 30-second time step.

.EXAMPLE
    . .\Get-TOTP.ps1
    Get-TOTP -SecretHex "48656C6C6F21" -TimeStep 60 -Digits 6

    Generates a 6-digit TOTP with a 60-second interval, typical for hardware tokens.

.EXAMPLE
    . .\Get-TOTP.ps1
    Convert-HexToBase32 "48656C6C6F21"

    Converts a hex string to Base32 encoding. Returns: JBSWY3DPBI======

.EXAMPLE
    . .\Get-TOTP.ps1
    "48656C6C6F21" | Convert-HexToBase32

    Pipeline input is supported for Convert-HexToBase32.

.EXAMPLE
    . .\Get-TOTP.ps1
    $hex = "31323334353637383930"
    Write-Host "Base32: $(Convert-HexToBase32 $hex)"
    Write-Host "Current TOTP: $(Get-TOTP -SecretHex $hex -TimeStep 30)"

    Convert a secret to Base32 and generate its current TOTP code in one session.

.NOTES
    Author: Mike Crowley

    No external modules or dependencies required. Uses only built-in .NET classes:
        System.Security.Cryptography.HMACSHA1
        System.Security.Cryptography.HMACSHA256

    RFC References:
        RFC 6238 - TOTP: Time-Based One-Time Password Algorithm
        RFC 4226 - HOTP: An HMAC-Based One-Time Password Algorithm
        RFC 4648 - The Base16, Base32, and Base64 Data Encodings

    Note: The system clock must be reasonably accurate for TOTP codes to match.
    Codes are valid for one TimeStep window. If activation against an API fails,
    ensure the system time is synced (e.g., w32tm /resync).

.LINK
    https://github.com/Mike-Crowley/Public-Scripts
#>

function Convert-HexToBase32 {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline, Position = 0)]
        [string]$HexString
    )
    process {
        if ($HexString -notmatch '^[0-9A-Fa-f]+$') {
            Write-Error "Input is not a valid hexadecimal string."
            return
        }
        if ($HexString.Length % 2 -ne 0) {
            $HexString = "0" + $HexString
        }

        # Convert hex to byte array
        $bytes = [byte[]]::new($HexString.Length / 2)
        for ($i = 0; $i -lt $HexString.Length; $i += 2) {
            $bytes[$i / 2] = [Convert]::ToByte($HexString.Substring($i, 2), 16)
        }

        # RFC 4648 Base32 alphabet
        $alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ234567"

        # Convert bytes to binary string
        $binaryString = -join ($bytes | ForEach-Object { [Convert]::ToString($_, 2).PadLeft(8, '0') })

        # Process 5 bits at a time
        $base32 = ""
        for ($i = 0; $i -lt $binaryString.Length; $i += 5) {
            if ($i + 5 -le $binaryString.Length) {
                $chunk = $binaryString.Substring($i, 5)
            }
            else {
                $chunk = $binaryString.Substring($i).PadRight(5, '0')
            }
            $base32 += $alphabet[[Convert]::ToInt32($chunk, 2)]
        }

        # Add padding to nearest multiple of 8
        $padding = (8 - ($base32.Length % 8)) % 8
        if ($padding -gt 0) {
            $base32 += "=" * $padding
        }

        return $base32
    }
}

function Get-TOTP {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$SecretHex,

        [int]$TimeStep = 30,

        [ValidateSet(6, 8)]
        [int]$Digits = 6,

        [ValidateSet("hmacsha1", "hmacsha256")]
        [string]$HashFunction = "hmacsha1"
    )

    # Convert hex secret to byte array
    $keyBytes = [byte[]]::new($SecretHex.Length / 2)
    for ($i = 0; $i -lt $SecretHex.Length; $i += 2) {
        $keyBytes[$i / 2] = [Convert]::ToByte($SecretHex.Substring($i, 2), 16)
    }

    # Create HMAC with the appropriate algorithm
    if ($HashFunction -eq "hmacsha256") {
        $hmac = [System.Security.Cryptography.HMACSHA256]::new($keyBytes)
    }
    else {
        $hmac = [System.Security.Cryptography.HMACSHA1]::new($keyBytes)
    }

    # Calculate time counter (RFC 6238)
    $counter = [math]::Floor(([DateTimeOffset]::UtcNow.ToUnixTimeSeconds()) / $TimeStep)

    # Convert counter to 8-byte big-endian array
    $counterBytes = [BitConverter]::GetBytes([int64]$counter)
    if ([BitConverter]::IsLittleEndian) {
        [Array]::Reverse($counterBytes)
    }

    # Compute HMAC hash
    $hash = $hmac.ComputeHash($counterBytes)
    $hmac.Dispose()

    # Dynamic truncation (RFC 4226 Section 5.4)
    $offset = $hash[$hash.Length - 1] -band 0xF
    $code = (($hash[$offset] -band 0x7F) -shl 24) -bor
            (($hash[$offset + 1] -band 0xFF) -shl 16) -bor
            (($hash[$offset + 2] -band 0xFF) -shl 8) -bor
             ($hash[$offset + 3] -band 0xFF)

    return ($code % [math]::Pow(10, $Digits)).ToString().PadLeft($Digits, '0')
}
