<#
.SYNOPSIS
    Finds enabled Active Directory accounts that lack AES Kerberos keys and therefore still
    rely on RC4, so you can remediate them before RC4 is disabled. Read-only; no key material
    is written to disk.

.DESCRIPTION
    As Microsoft disables RC4 by default (the msDS-SupportedEncryptionTypes hardening wave and
    the DefaultDomainSupportedEncTypes change), accounts that have never had AES keys generated
    will start failing Kerberos once RC4 is turned off. The usual signals -- Kerberos error
    events, audit logs, or pwdLastSet heuristics -- only catch accounts that have authenticated
    recently. This script instead enumerates EVERY account directly from the directory so the
    impact list is complete, independent of activity or denial events.

    It runs in two phases:

        Phase A (dump)      Streams Kerberos key-type *presence* for every account to a
                            newline-delimited JSON file (.jsonl). Memory stays constant
                            regardless of directory size, and NO secrets are written -- only
                            which encryption types each account holds a key for.

        Phase B (report)    Filters the dump to enabled accounts with a manually managed
                            password (excludes machine and managed service accounts -- see
                            NOTES) that have neither an AES128 nor an AES256 key, decorates each
                            with LastLogonDate, OU, SPN count, msDS-SupportedEncryptionTypes,
                            and Description from AD, then writes the impact CSV with the most
                            recently active accounts first (the highest remediation priority).

    Re-run the report against the existing dump without pulling from AD again with -SkipDump.

    Remediating a flagged account is typically a password reset -- which generates AES keys on a
    modern domain functional level -- and/or setting msDS-SupportedEncryptionTypes to include AES.

.PARAMETER Server
    FQDN of the writable domain controller to replicate from. Defaults to a discovered DC for
    the current domain.

.PARAMETER OutputFolder
    Folder for the key-type dump and the impact CSV. Defaults to an "RC4-Impact" folder on the
    current user's Desktop, resolved via [Environment]::GetFolderPath("Desktop") so it follows
    OneDrive Known Folder redirection. The dump file is reused by -SkipDump.

.PARAMETER SkipDump
    Reuse the existing dump file and run Phase B only (seconds). Throws if no dump exists yet.

.EXAMPLE
    Install-Module DSInternals -Scope CurrentUser
    .\Find-Rc4Impact.ps1

    Full run from an account with DCSync rights (see NOTES): dump key-type metadata from a
    discovered DC, then write the enriched RC4 impact CSV to <Desktop>\RC4-Impact.

.EXAMPLE
    .\Find-Rc4Impact.ps1 -Server dc01.corp.example.com -OutputFolder C:\Audit\RC4

    Replicate from a specific DC and write the dump and impact CSV to a custom folder.

.EXAMPLE
    .\Find-Rc4Impact.ps1 -SkipDump -Verbose

    Re-run the report from the existing dump (no AD replication) -- useful for re-querying the
    AD enrichment without a second multi-minute replication pass.

.NOTES
    Author: Mike Crowley
    https://mikecrowley.us

    Requires: DSInternals module (Install-Module DSInternals) and the ActiveDirectory module (RSAT).

    Connectivity: the ActiveDirectory module calls (Get-ADDomain / Get-ADUser) reach the DC over
    AD Web Services (TCP 9389), which the script preflights before dumping. Get-ADReplAccount
    replicates over RPC (TCP 135 plus a dynamic high port) -- open those too on any firewall.

    Privilege: Phase A uses Get-ADReplAccount, which performs a directory replication (DCSync).
    The account running it needs the "Replicating Directory Changes" and "Replicating Directory
    Changes All" rights on the domain -- effectively Domain Admin equivalent. Run it from a
    trusted host. The dump file holds account names and key-type presence only; it contains no
    password hashes or key material.

    Scope: accounts whose sAMAccountName ends in "$" -- computer accounts and (group) Managed
    Service Accounts -- are treated as machine accounts and excluded from the impact list. They
    rotate their own passwords automatically and obtain AES keys without admin action, so they
    are not the focus of this audit. Phase A still dumps them.

    DSInternals version note: the credential path used here is
    SupplementalCredentials.KerberosNew.Credentials, with a .KeyType per entry. Confirm it on
    your version before a large run:

        $p = Get-ADReplAccount -SamAccountName krbtgt -Domain CORP -Server dc01.corp.example.com
        $p.SupplementalCredentials.KerberosNew.Credentials | Format-Table KeyType

.LINK
    https://github.com/Mike-Crowley/Public-Scripts

.LINK
    https://github.com/MichaelGrafnetter/DSInternals/blob/master/Documentation/PowerShell/Get-ADReplAccount.md

.LINK
    https://learn.microsoft.com/en-us/windows-server/security/kerberos/detect-remediate-rc4-kerberos

.LINK
    https://www.microsoft.com/en-us/windows-server/blog/2025/12/03/beyond-rc4-for-windows-authentication/

.LINK
    https://support.microsoft.com/en-us/topic/kb5021131-how-to-manage-the-kerberos-protocol-changes-related-to-cve-2022-37966-fd837ac3-cdec-4e76-a6ec-86e67501407d
#>

#Requires -Modules DSInternals, ActiveDirectory

[CmdletBinding()]
param(
    [string]$Server,

    [ValidateNotNullOrEmpty()]
    [string]$OutputFolder = (Join-Path ([Environment]::GetFolderPath("Desktop")) 'RC4-Impact'),

    [switch]$SkipDump
)

$ErrorActionPreference = 'Stop'

# Resolve a DC to replicate from if one wasn't supplied.
if (-not $Server) {
    $Server = (Get-ADDomainController -Discover -ErrorAction Stop).HostName | Select-Object -First 1
    Write-Verbose "No -Server supplied; using discovered DC: $Server"
}

# Preflight: confirm AD Web Services (TCP 9389) is reachable before any ActiveDirectory module
# calls. Get-ADReplAccount replicates over RPC separately, but Get-ADDomain/Get-ADUser need ADWS.
Write-Host "Checking AD Web Services connectivity to $Server (TCP 9389)..." -ForegroundColor Cyan
if (-not (Test-NetConnection -ComputerName $Server -Port 9389 -InformationLevel Quiet -WarningAction SilentlyContinue)) {
    throw "Cannot reach AD Web Services on $Server (TCP 9389). Confirm the DC is online, the ADWS service is running, and the port is open through any firewall."
}

if (-not (Test-Path $OutputFolder)) {
    New-Item -Path $OutputFolder -ItemType Directory -Force | Out-Null
}

$dumpPath = Join-Path $OutputFolder 'aes-keydump.jsonl'
$tmpPath  = "$dumpPath.tmp"
$stamp    = Get-Date -Format 'yyyyMMdd-HHmm'
$outCsv   = Join-Path $OutputFolder "rc4-impact-$stamp.csv"

if ($SkipDump -and -not (Test-Path $dumpPath)) {
    throw "No dump found at $dumpPath -- run once without -SkipDump first."
}

# ---- Phase A: stream key-type presence for every account to NDJSON (no secrets) ----
if (-not $SkipDump) {
    $namingContext = (Get-ADDomain -Server $Server).DistinguishedName
    Write-Host "Phase A: dumping key metadata from $Server ($namingContext)..." -ForegroundColor Cyan
    # Explicit UTF-8 without BOM so ConvertFrom-Json can parse the first record on every PS edition.
    $writer = [System.IO.StreamWriter]::new($tmpPath, $false, [System.Text.UTF8Encoding]::new($false))
    $count  = 0
    try {
        $dumpParams = @{
            All           = $true
            Server        = $Server
            NamingContext = $namingContext
            ErrorAction   = 'Stop'
        }
        Get-ADReplAccount @dumpParams | ForEach-Object {
            $keyTypes = @()
            if ($_.SupplementalCredentials -and $_.SupplementalCredentials.KerberosNew) {
                $keyTypes = @($_.SupplementalCredentials.KerberosNew.Credentials | ForEach-Object { "$($_.KeyType)" })
            }
            $record = [pscustomobject]@{
                Sam        = $_.SamAccountName
                Enabled    = [bool]$_.Enabled
                IsComputer = $_.SamAccountName.EndsWith('$')
                Dn         = $_.DistinguishedName
                PwdLastSet = if ($_.PasswordLastSet) { $_.PasswordLastSet.ToString('o') } else { $null }
                HasAES256  = $keyTypes -contains 'AES256_CTS_HMAC_SHA1_96'
                HasAES128  = $keyTypes -contains 'AES128_CTS_HMAC_SHA1_96'
                # The RC4 key IS the NT hash; per MS-SAMR the KerberosNew package only holds
                # DES/AES key types, so HasRC4 is effectively always False here even though
                # every account with a password "has RC4." Kept for schema completeness only --
                # do not use it to reason about RC4 presence. AES absence is the real signal.
                HasRC4     = $keyTypes -contains 'RC4_HMAC_NT'
            }
            $writer.WriteLine(($record | ConvertTo-Json -Compress))
            $count++
            if ($count % 5000 -eq 0) { Write-Progress -Activity 'Dumping accounts' -Status "$count processed" }
        }
    }
    finally {
        $writer.Dispose()
        Write-Progress -Activity 'Dumping accounts' -Completed
    }
    Move-Item -Path $tmpPath -Destination $dumpPath -Force
    Write-Host "  wrote $count records to $dumpPath" -ForegroundColor Green
}

# ---- Phase B: filter the dump to enabled, non-machine, AES-less accounts ----
Write-Host 'Phase B: scanning dump for AES-less accounts...' -ForegroundColor Cyan
$impact = [System.Collections.Generic.List[object]]::new()
$total  = 0
foreach ($line in [System.IO.File]::ReadLines($dumpPath)) {
    if (-not $line) { continue }
    $total++
    $a = $line | ConvertFrom-Json
    if (-not $a.Enabled)               { continue }
    if ($a.IsComputer)                 { continue }
    if ($a.HasAES256 -or $a.HasAES128) { continue }
    $impact.Add($a)
}
Write-Host "Scanned $total accounts; $($impact.Count) enabled accounts have NO AES keys." -ForegroundColor Yellow

if ($impact.Count -eq 0) {
    Write-Host 'Nothing to report -- every enabled account already has an AES key.' -ForegroundColor Green
    return
}

# Enrich each impacted account from AD to prioritize remediation. Get-ADUser is scoped to the
# (usually small) impact list, not the whole directory, so this stays cheap.
Write-Host '  enriching impact list from AD...' -ForegroundColor Cyan
Write-Verbose "Enriching via $Server"
$enrichProps = 'LastLogonDate', 'servicePrincipalName', 'msDS-SupportedEncryptionTypes', 'Description'
$report = foreach ($r in $impact) {
    $ad = Get-ADUser -Identity $r.Dn -Properties $enrichProps -Server $Server -ErrorAction SilentlyContinue
    [pscustomobject]@{
        Account                       = $r.Sam
        LastLogon                     = $ad.LastLogonDate
        PwdLastSet                    = $r.PwdLastSet
        SPNs                          = if ($ad) { @($ad.servicePrincipalName).Count } else { $null }
        msDS_SupportedEncryptionTypes = $ad.'msDS-SupportedEncryptionTypes'
        OU                            = if ($ad) { ($ad.DistinguishedName -split '(?<!\\),', 2)[1] } else { $null }
        Description                   = $ad.Description
    }
}
$report | Sort-Object LastLogon -Descending | Export-Csv -Path $outCsv -NoTypeInformation -Encoding UTF8
Write-Host "Impact report: $outCsv" -ForegroundColor Green
