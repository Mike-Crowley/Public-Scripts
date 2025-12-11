<#

.SYNOPSIS
    This tool displays the signing and encrypting certificates published in ADFS or Entra ID federation metadata as well as the HTTPS ("SSL") certificate used in the connection itself.

    This tool does not authenticate to the server or investigate each ADFS farm node directly. For this, use the ADFS Cert Diag tool

    Version: Dec 2025

.DESCRIPTION
    This tool displays the signing and encrypting certificates published in ADFS or Entra ID federation metadata as well as the HTTPS ("SSL") certificate used in the connection itself.

    Supports two modes:
    - ADFS mode: Query an ADFS farm by FQDN (uses -FarmFqdn parameter)
    - Entra ID mode: Query Entra ID federation metadata by URL (uses -MetadataUrl parameter)

    DaysToExpiry values are color-coded: Green (90+ days), Yellow (30-89 days), Red (<30 days or expired).

    Sample Output with -Display $true (default):

        EntityID: http://adfs.contoso.com/adfs/services/trust

        SSL (HTTPS) Certificate:

            SSL_Subject:       CN=ADFS.CONTOSO, O=CONTOSO CORP, OID.1.3.6.1.4.1.311.60.2.1.3=US
            SSL_NotAfter:      1/14/2024 6:59:59 PM
            SSL_Thumbprint:    21321F3C2E225480F112A7BC2B3347B58B439842
            SSL_Issuer:        CN=CONTOSO CORP
            SSL_DaysToExpiry:  25

        Encryption Certificate:

            Encryption_Subject:     CN=ADFS Encryption - adfs.contoso.com
            Encryption_NotAfter:    7/7/2023 7:05:31 PM
            Encryption_Thumbprint:  0507D8E023B8715FE3F5F4A6421F47A36C6DD3AD
            Encryption_Issuer:      CN=ADFS Encryption - adfs.contoso.com
            Encryption_DaysToExpiry:  129

        Token Signing Certificate:

            FirstSigning_Subject:     CN=ADFS Signing - adfs.contoso.com
            FirstSigning_NotAfter:    7/7/2023 7:05:32 PM
            FirstSigning_Thumbprint:  0507D8E023B8715FE3F5F4A6421F47A36C6DD3AD
            FirstSigning_Issuer:      CN=ADFS Signing - adfs.contoso.com
            FirstSigning_DaysToExpiry:  129

        Second Token Signing Certificate:

            !! No Second Token Signing Certificate Found !!

    Sample Output with -Display $false (for use with loops, pipeline, etc):

            EntityID                      : http://adfs.contoso.com/adfs/services/trust
            SSL_Subject                   : CN=ADFS.CONTOSO, O=CONTOSO CORP, OID.1.3.6.1.4.1.311.60.2.1.3=US
            SSL_NotAfter                  : 11/14/2023 6:59:59 PM
            SSL_Thumbprint                : 21321F3C2E225480F112A7BC2B3347B58B439842
            SSL_Issuer                    : CN=CONTOSO CORP
            SSL_DaysToExpiry              : 256
            FirstSigning_Subject          : CN=ADFS Signing - adfs.contoso.com
            FirstSigning_NotAfter         : 7/5/2023 7:05:32 PM
            FirstSigning_Thumbprint       : 0507D8E023B8715FE3F5F4A6421F47A36C6DD3AD
            FirstSigning_Issuer           : CN=ADFS Signing - adfs.contoso.com
            FirstSigning_DaysToExpiry     : 124
            SecondSigning_Subject         :
            SecondSigning_NotAfter        :
            SecondSigning_Thumbprint      :
            SecondSigning_Issuer          :
            SecondSigning_DaysToExpiry    :
            Encryption_Subject            : CN=ADFS Encryption - adfs.contoso.com
            Encryption_NotAfter           : 7/5/2023 7:05:31 PM
            Encryption_Thumbprint         : 0507D8E023B8715FE3F5F4A6421F47A36C6DD3AD
            Encryption_Issuer             : CN=ADFS Encryption - adfs.contoso.com
            Encryption_DaysToExpiry       : 124

    NOTE: This tool by Microsoft may be handy as well: https://adfshelp.microsoft.com/MetadataExplorer/GetFederationMetadata

    Author:
    Mike Crowley
    https://mikecrowley.us

.EXAMPLE
    Request-FederationCerts -FarmFqdn adfs.contoso.com

.EXAMPLE
    Request-FederationCerts -FarmFqdn adfs.contoso.com -Display $false

.EXAMPLE
    Request-FederationCerts -MetadataUrl "https://login.microsoftonline.com/contoso.onmicrosoft.com/federationmetadata/2007-06/federationmetadata.xml"

.EXAMPLE
    Request-FederationCerts -MetadataUrl "https://login.microsoftonline.com/50afff0a-571a-4822-915b-d6823ca9fe63/federationmetadata/2007-06/federationmetadata.xml?appid=df3d5533-18a3-400a-9027-ae7da4836fc7"

.EXAMPLE
    Request-FederationCerts -FarmFqdn adfs.contoso.com -ExportCsv "C:\reports\certs.csv" -ExportJson "C:\reports\certs.json"

.LINK
    https://github.com/mike-crowley_blkln
    https://github.com/Mike-Crowley

#>

function Request-FederationCerts {
    [CmdletBinding(DefaultParameterSetName = 'ADFS')]
    param (
        [Parameter(ParameterSetName = 'ADFS', ValueFromPipeline = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$FarmFqdn,

        [Parameter(ParameterSetName = 'Entra', ValueFromPipeline = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$MetadataUrl,

        [bool]$Display = $true,

        [string]$ExportCsv,

        [string]$ExportJson
    )

    $IsPSCore = $PSVersionTable.PSEdition -eq 'Core'

    # Validate that at least one parameter is provided
    if (-not $FarmFqdn -and -not $MetadataUrl) {
        Write-Warning "Please specify either -FarmFqdn or -MetadataUrl"
        return
    }

    if ($FarmFqdn) {
        # ADFS mode - test connection first
        if (-not (Test-NetConnection -ComputerName $FarmFqdn -Port 443 -InformationLevel Quiet -Verbose)) {
            Write-Warning "Cannot connect to: $FarmFqdn"
            return
        }
        $url = "https://$FarmFqdn/FederationMetadata/2007-06/FederationMetadata.xml"
        $hostHeader = $FarmFqdn
    }
    else {
        # Entra mode - use the provided URL directly
        $url = $MetadataUrl
        $hostHeader = ([uri]$MetadataUrl).Host
    }

    # Ensure TLS 1.2
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    # For Windows PowerShell, ignore SSL warnings for self-signed certs
    if (-not $IsPSCore) {
        [Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }
    }

    # Make HTTPS connection and get content
    try {
        if ($IsPSCore) {
            # PowerShell 7+ - use Invoke-WebRequest with -SkipCertificateCheck for self-signed certs
            $webResponse = Invoke-WebRequest -Uri $url -UseBasicParsing -SkipCertificateCheck
            $content = $webResponse.Content

            # Get SSL cert from the connection via TcpClient
            $tcpClient = [System.Net.Sockets.TcpClient]::new($hostHeader, 443)
            $sslStream = [System.Net.Security.SslStream]::new($tcpClient.GetStream(), $false, { $true })
            $sslStream.AuthenticateAsClient($hostHeader)
            $SSLCert_x509 = [Security.Cryptography.X509Certificates.X509Certificate2]::new($sslStream.RemoteCertificate)
            $sslStream.Close()
            $tcpClient.Close()
        }
        elseif ($MetadataUrl) {
            # Windows PowerShell + Entra mode - use Invoke-WebRequest
            $webResponse = Invoke-WebRequest -Uri $url -UseBasicParsing
            $content = $webResponse.Content

            # Get SSL cert from the connection
            $tcpClient = [System.Net.Sockets.TcpClient]::new($hostHeader, 443)
            $sslStream = [System.Net.Security.SslStream]::new($tcpClient.GetStream(), $false, { $true })
            $sslStream.AuthenticateAsClient($hostHeader)
            $SSLCert_x509 = [Security.Cryptography.X509Certificates.X509Certificate2]::new($sslStream.RemoteCertificate)
            $sslStream.Close()
            $tcpClient.Close()
        }
        else {
            # Windows PowerShell + ADFS mode - use HttpWebRequest to handle custom Host header
            $request = [Net.HttpWebRequest]::Create($url)
            $request.Host = $hostHeader
            $request.AllowAutoRedirect = $false
            $response = $request.GetResponse()

            $HttpsCertBytes = $request.ServicePoint.Certificate.GetRawCertData()
            $contentStream = $response.GetResponseStream()
            $reader = [IO.StreamReader]::new($contentStream)
            $content = $reader.ReadToEnd()
            $reader.Close()
            $contentStream.Close()
            $response.Close()

            # Extract HTTPS cert (ADFS calls this the "SSL" cert)
            $CertInBase64 = [convert]::ToBase64String($HttpsCertBytes)
            $SSLCert_x509 = [Security.Cryptography.X509Certificates.X509Certificate2]([System.Convert]::FromBase64String($CertInBase64))
        }
    }
    catch {
        Write-Warning "Failed to retrieve metadata from: $url"
        Write-Warning $_.Exception.Message
        return
    }

    # Parse FederationMetadata for certs
    # Remove BOM if present (can cause XML parsing issues in PowerShell 7)
    $content = $content -replace '^\xEF\xBB\xBF', '' -replace '^\uFEFF', ''
    # ADFS uses SPSSODescriptor, Entra ID uses IDPSSODescriptor
    $xmlContent = [xml]$content

    # Extract EntityID from the metadata
    $EntityID = $xmlContent.EntityDescriptor.entityID

    $KeyDescriptors = $xmlContent.EntityDescriptor.SPSSODescriptor.KeyDescriptor
    if (-not $KeyDescriptors) {
        $KeyDescriptors = $xmlContent.EntityDescriptor.IDPSSODescriptor.KeyDescriptor
    }

    $FirstSigningCert_base64 = ([array]($KeyDescriptors | Where-Object use -eq 'signing').KeyInfo)[0].X509Data.X509Certificate
    $FirstSigningCert_x509 = if ($FirstSigningCert_base64) { [Security.Cryptography.X509Certificates.X509Certificate2][System.Convert]::FromBase64String($FirstSigningCert_base64) } else { $null }

    $SecondSigningCert_base64 = ([array]($KeyDescriptors | Where-Object use -eq 'signing').KeyInfo)[1].X509Data.X509Certificate
    $SecondSigningCert_x509 = if ($SecondSigningCert_base64) { [Security.Cryptography.X509Certificates.X509Certificate2][System.Convert]::FromBase64String($SecondSigningCert_base64) } else { $null }

    $EncryptionCert_base64 = ($KeyDescriptors | Where-Object use -eq 'encryption').KeyInfo.X509Data.X509Certificate
    $EncryptionCert_x509 = if ($EncryptionCert_base64) { [Security.Cryptography.X509Certificates.X509Certificate2][System.Convert]::FromBase64String($EncryptionCert_base64) } else { $null }

    $Now = Get-Date

    $CertReportObject = [pscustomobject]@{
        EntityID                   = $EntityID

        SSL_Subject                = $SSLCert_x509.Subject
        SSL_NotAfter               = $SSLCert_x509.NotAfter
        SSL_Thumbprint             = $SSLCert_x509.Thumbprint
        SSL_Issuer                 = $SSLCert_x509.Issuer
        SSL_DaysToExpiry           = ($SSLCert_x509.NotAfter - $Now).Days

        FirstSigning_Subject       = $FirstSigningCert_x509.Subject
        FirstSigning_NotAfter      = $FirstSigningCert_x509.NotAfter
        FirstSigning_Thumbprint    = $FirstSigningCert_x509.Thumbprint
        FirstSigning_Issuer        = $FirstSigningCert_x509.Issuer
        FirstSigning_DaysToExpiry  = if ($FirstSigningCert_x509) { ($FirstSigningCert_x509.NotAfter - $Now).Days } else { $null }

        SecondSigning_Subject      = $SecondSigningCert_x509.Subject
        SecondSigning_NotAfter     = $SecondSigningCert_x509.NotAfter
        SecondSigning_Thumbprint   = $SecondSigningCert_x509.Thumbprint
        SecondSigning_Issuer       = $SecondSigningCert_x509.Issuer
        SecondSigning_DaysToExpiry = if ($SecondSigningCert_x509) { ($SecondSigningCert_x509.NotAfter - $Now).Days } else { $null }

        Encryption_Subject         = $EncryptionCert_x509.Subject
        Encryption_NotAfter        = $EncryptionCert_x509.NotAfter
        Encryption_Thumbprint      = $EncryptionCert_x509.Thumbprint
        Encryption_Issuer          = $EncryptionCert_x509.Issuer
        Encryption_DaysToExpiry    = if ($EncryptionCert_x509) { ($EncryptionCert_x509.NotAfter - $Now).Days } else { $null }
    }

    # Helper function for expiry color coding
    function Get-ExpiryColor {
        param([int]$Days)
        if ($null -eq $Days) { return 'Gray' }
        if ($Days -lt 0) { return 'Red' }
        if ($Days -lt 30) { return 'Red' }
        if ($Days -lt 90) { return 'Yellow' }
        return 'Green'
    }

    if ($Display -eq $true) {

        Write-Host "`n    EntityID: " -ForegroundColor Cyan -NoNewline
        Write-Host $CertReportObject.EntityID

        Write-Host "    `nSSL (HTTPS) Certificate:`n" -ForegroundColor Green
        Write-Host "    SSL_Subject:     " $CertReportObject.SSL_Subject
        Write-Host "    SSL_NotAfter:    " $CertReportObject.SSL_NotAfter
        Write-Host "    SSL_Thumbprint:  " $CertReportObject.SSL_Thumbprint
        Write-Host "    SSL_Issuer:      " $CertReportObject.SSL_Issuer
        Write-Host "    SSL_DaysToExpiry: " -NoNewline
        Write-Host $CertReportObject.SSL_DaysToExpiry -ForegroundColor (Get-ExpiryColor $CertReportObject.SSL_DaysToExpiry)

        if ($null -ne $CertReportObject.Encryption_Subject) {
            Write-Host "    `nEncryption Certificate:`n" -ForegroundColor DarkMagenta
            Write-Host "    Encryption_Subject:     " $CertReportObject.Encryption_Subject
            Write-Host "    Encryption_NotAfter:    " $CertReportObject.Encryption_NotAfter
            Write-Host "    Encryption_Thumbprint:  " $CertReportObject.Encryption_Thumbprint
            Write-Host "    Encryption_Issuer:      " $CertReportObject.Encryption_Issuer
            Write-Host "    Encryption_DaysToExpiry: " -NoNewline
            Write-Host $CertReportObject.Encryption_DaysToExpiry -ForegroundColor (Get-ExpiryColor $CertReportObject.Encryption_DaysToExpiry)
        }
        else {
            Write-Host "    `nEncryption Certificate:`n" -ForegroundColor DarkMagenta
            Write-Host "    !! No Encryption Certificate Found (typical for Entra ID metadata) !!`n"
        }

        if ($null -eq $CertReportObject.SecondSigning_Subject) {
            Write-Host "    `nToken Signing Certificate:`n" -ForegroundColor Yellow
        }
        else {
            Write-Host "    `nFirst Token Signing Certificate:`n" -ForegroundColor Yellow
        }
        Write-Host "    FirstSigning_Subject:     " $CertReportObject.FirstSigning_Subject
        Write-Host "    FirstSigning_NotAfter:    " $CertReportObject.FirstSigning_NotAfter
        Write-Host "    FirstSigning_Thumbprint:  " $CertReportObject.FirstSigning_Thumbprint
        Write-Host "    FirstSigning_Issuer:      " $CertReportObject.FirstSigning_Issuer
        Write-Host "    FirstSigning_DaysToExpiry: " -NoNewline
        Write-Host $CertReportObject.FirstSigning_DaysToExpiry -ForegroundColor (Get-ExpiryColor $CertReportObject.FirstSigning_DaysToExpiry)

        Write-Host "`nSecond Token Signing Certificate:`n" -ForegroundColor DarkYellow

        if ($null -ne $CertReportObject.SecondSigning_Subject) {
            Write-Host "    SecondSigning_Subject:     " $CertReportObject.SecondSigning_Subject
            Write-Host "    SecondSigning_NotAfter:    " $CertReportObject.SecondSigning_NotAfter
            Write-Host "    SecondSigning_Thumbprint:  " $CertReportObject.SecondSigning_Thumbprint
            Write-Host "    SecondSigning_Issuer:      " $CertReportObject.SecondSigning_Issuer
            Write-Host "    SecondSigning_DaysToExpiry: " -NoNewline
            Write-Host $CertReportObject.SecondSigning_DaysToExpiry -ForegroundColor (Get-ExpiryColor $CertReportObject.SecondSigning_DaysToExpiry)
            Write-Host "`n    NOTE: A federation metadata document can have multiple signing keys." -ForegroundColor Gray
            Write-Host "      When a federation metadata document includes more than one certificate, a service that is validating the tokens should support all certificates in the document." -ForegroundColor Gray
            Write-Host "      The 'First' certificate in the metadata may not be the 'Primary' certificate in the ADFS configuration"
            Write-Host "      https://learn.microsoft.com/en-us/entra/identity-platform/federation-metadata#token-signing-certificates"
        }
        else { Write-Host "    !! No Second Token Signing Certificate Found !!`n" }

        Write-Host "`n"
    }

    # Export to CSV if requested
    if ($ExportCsv) {
        try {
            $CertReportObject | Export-Csv -Path $ExportCsv -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
            Write-Host "Exported to CSV: $ExportCsv" -ForegroundColor Green
        }
        catch {
            Write-Warning "Failed to export CSV: $_"
        }
    }

    # Export to JSON if requested
    if ($ExportJson) {
        try {
            $CertReportObject | ConvertTo-Json -Depth 3 | Out-File -FilePath $ExportJson -Encoding UTF8 -ErrorAction Stop
            Write-Host "Exported to JSON: $ExportJson" -ForegroundColor Green
        }
        catch {
            Write-Warning "Failed to export JSON: $_"
        }
    }

    # Return object if not displaying (for pipeline use)
    if ($Display -eq $false) {
        return $CertReportObject
    }
}

# Examples
# ADFS mode:
# Request-FederationCerts -FarmFqdn adfs.contoso.com -Display $true

# Entra ID mode:
# Request-FederationCerts -MetadataUrl "https://login.microsoftonline.com/contoso.onmicrosoft.com/federationmetadata/2007-06/federationmetadata.xml"