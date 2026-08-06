#Requires -Version 5.1
<#
    Rebuilds SupportingFiles/PassKey_AAGUID_Catalog.json from the FIDO Alliance
    Metadata Service (MDS3), the authoritative registry of certified authenticators.
    The BLOB is a JWS; this verifies both the signing chain and the RS256 signature
    before any data is trusted.
#>
param(
    [Parameter(Mandatory)][string]$OutFile,
    [string]$BlobUri = 'https://mds3.fidoalliance.org/',
    # MDS rate-limits repeat downloads. A cached copy is re-verified from scratch, so
    # reusing one is not a shortcut around the signature check.
    [string]$CachedBlob
)
$ErrorActionPreference = 'Stop'

function FromB64Url([string]$s) {
    $s = $s.Replace('-', '+').Replace('_', '/')
    switch ($s.Length % 4) { 1 { $s = $s.Substring(0, $s.Length - 1) } 2 { $s += '==' } 3 { $s += '=' } }
    return [Convert]::FromBase64String($s)
}

if ($CachedBlob -and (Test-Path $CachedBlob)) {
    Write-Host "Using cached BLOB $CachedBlob (signature is still verified below)" -ForegroundColor Cyan
    $raw = Get-Content $CachedBlob -Raw
}
else {
    Write-Host "Downloading $BlobUri ..." -ForegroundColor Cyan
    $raw = (Invoke-WebRequest -Uri $BlobUri -TimeoutSec 120).Content
    if ($raw -is [byte[]]) { $raw = [Text.Encoding]::UTF8.GetString($raw) }
}
$parts = $raw.Trim() -split '\.'
if ($parts.Count -ne 3) { throw "Expected a 3-part JWS, got $($parts.Count) parts." }

$hdr = [Text.Encoding]::UTF8.GetString((FromB64Url $parts[0])) | ConvertFrom-Json
$signer = [Security.Cryptography.X509Certificates.X509Certificate2]::new([Convert]::FromBase64String($hdr.x5c[0]))

# 1. Chain must validate to a trusted root.
$chain = [Security.Cryptography.X509Certificates.X509Chain]::new()
$chain.ChainPolicy.RevocationMode = 'Online'
foreach ($c in $hdr.x5c) {
    $null = $chain.ChainPolicy.ExtraStore.Add([Security.Cryptography.X509Certificates.X509Certificate2]::new([Convert]::FromBase64String($c)))
}
if (-not $chain.Build($signer)) {
    throw "MDS signing certificate chain did not validate: $(($chain.ChainStatus | ForEach-Object { $_.Status }) -join ', ')"
}

# 2. Signature must verify over header.payload.
$rsa = [Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPublicKey($signer)
$ok = $rsa.VerifyData([Text.Encoding]::ASCII.GetBytes("$($parts[0]).$($parts[1])"), (FromB64Url $parts[2]),
    [Security.Cryptography.HashAlgorithmName]::SHA256, [Security.Cryptography.RSASignaturePadding]::Pkcs1)
if (-not $ok) { throw 'MDS BLOB signature did not verify. Refusing to use the data.' }
Write-Host "  Signature verified. Signed by: $($signer.Subject)" -ForegroundColor Green

$mds = [Text.Encoding]::UTF8.GetString((FromB64Url $parts[1])) | ConvertFrom-Json
Write-Host "  BLOB no.$($mds.no), nextUpdate $($mds.nextUpdate), $($mds.entries.Count) entries" -ForegroundColor Green

$CertifiedStatuses = @('FIDO_CERTIFIED', 'FIDO_CERTIFIED_L1', 'FIDO_CERTIFIED_L1plus',
    'FIDO_CERTIFIED_L2', 'FIDO_CERTIFIED_L2plus', 'FIDO_CERTIFIED_L3', 'FIDO_CERTIFIED_L3plus')

$out = foreach ($e in ($mds.entries | Where-Object { $_.aaguid })) {
    $ms = $e.metadataStatement
    $gi = $ms.authenticatorGetInfo
    $reports = @($e.statusReports)

    # Highest certification level attained, and the date it took effect.
    $certReports = @($reports | Where-Object { $CertifiedStatuses -contains $_.status })
    $best = $certReports | Sort-Object @{ e = { $CertifiedStatuses.IndexOf($_.status) }; Descending = $true } | Select-Object -First 1
    $certNumber = ($certReports | Where-Object { $_.certificateNumber } | Select-Object -First 1).certificateNumber

    $versions = @($gi.versions)
    $ctap = $null
    if ($versions -contains 'FIDO_2_1') { $ctap = '2.1' }
    elseif ($versions -contains 'FIDO_2_1_PRE') { $ctap = '2.1-pre' }
    elseif ($versions -contains 'FIDO_2_0') { $ctap = '2.0' }
    elseif ($versions -contains 'U2F_V2') { $ctap = 'u2f' }

    $kp = @($ms.keyProtection)

    [ordered]@{
        aaguid        = "$($e.aaguid)"
        name          = $ms.description
        certification = [ordered]@{
            isFidoCertified   = ($certReports.Count -gt 0)
            isRevoked         = ($reports.status -contains 'REVOKED')
            level             = if ($best) { $best.status } else { ($reports | Select-Object -First 1).status }
            certificationDate = if ($best) { $best.effectiveDate } else { $null }
            certificateNumber = $certNumber
            statusHistory     = @($reports | ForEach-Object {
                    [ordered]@{ status = $_.status; effectiveDate = $_.effectiveDate; certificateNumber = $_.certificateNumber }
                })
        }
        security      = [ordered]@{
            # Derived from the vendor's own model name in MDS. MDS has no FIPS field;
            # vendors mark FIPS models in the description (e.g. "YubiKey 5 FIPS Series").
            isFips            = [bool]($ms.description -match '\bFIPS\b')
            # Enterprise ATTESTATION capability, from authenticatorGetInfo.options.ep.
            # Not the same thing as a vendor's "Enterprise Edition" product name.
            isEnterprise      = [bool]($gi.options.ep)
            hardwareProtected = [bool](($kp -contains 'hardware') -or ($kp -contains 'secure_element') -or ($kp -contains 'tee'))
            keyProtection     = $kp
            matcherProtection = @($ms.matcherProtection)
            cryptoStrength    = $ms.cryptoStrength
        }
        connectivity  = [ordered]@{
            transports     = @($gi.transports)
            attachmentHint = @($ms.attachmentHint)
        }
        protocol      = [ordered]@{
            family     = $ms.protocolFamily
            versions   = $versions
            ctapLevel  = $ctap
            extensions = @($gi.extensions)
        }
        userVerification     = @($ms.userVerificationDetails | ForEach-Object { @($_ | ForEach-Object { $_.userVerificationMethod }) } | Where-Object { $_ } | Sort-Object -Unique)
        attestationTypes     = @($ms.attestationTypes)
        authenticatorVersion = $ms.authenticatorVersion
        timeOfLastStatusChange = $e.timeOfLastStatusChange
    }
}

$catalog = [ordered]@{
    source        = [ordered]@{
        name              = 'FIDO Alliance Metadata Service (MDS3)'
        url               = $BlobUri
        blobSerialNumber  = $mds.no
        blobNextUpdate    = $mds.nextUpdate
        retrievedUtc      = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
        signatureVerified = $true
        signedBy          = $signer.Subject
        legalHeader       = $mds.legalHeader
        derivedFields     = @(
            'security.isFips is derived from the vendor model name in metadataStatement.description; MDS has no FIPS field.',
            'security.isEnterprise reflects enterprise ATTESTATION support (authenticatorGetInfo.options.ep), not a product edition name.',
            'protocol.ctapLevel is derived from authenticatorGetInfo.versions.',
            'certification.level is the highest FIDO_CERTIFIED_* status present in statusReports.'
        )
        notes             = @(
            'Only entries carrying an aaguid are included. MDS also lists U2F and UAF authenticators keyed by attestationCertificateKeyIdentifiers or AAID; those have no AAGUID and are out of scope.',
            'Rebuild with SupportingFiles/Build-AaguidCatalog.ps1. The BLOB is re-signed periodically, so refresh after blobNextUpdate.',
            'Microsoft Entra ID ingests MDS on an unpublished monthly cadence, so a newly certified AAGUID can be absent from Entra attestation enforcement for some time after it appears here.'
        )
    }
    authenticatorCount = @($out).Count
    authenticators = @($out | Sort-Object { $_.name }, { $_.aaguid })
}

$json = $catalog | ConvertTo-Json -Depth 12
[IO.File]::WriteAllText($OutFile, $json, [Text.UTF8Encoding]::new($false))
Write-Host "Wrote $(@($out).Count) authenticators to $OutFile ($([Math]::Round((Get-Item $OutFile).Length/1KB)) KB)" -ForegroundColor Green
