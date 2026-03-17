<span style="font-size:large;">Author: Mike Crowley — Co-Founder, <a href="https://www.baselinetechnologies.com/leadership/">Baseline Technologies</a></span>

<p align="left">
    <a href="https://mikecrowley.us">
        <img alt="Mike's Blog"
            src="https://img.shields.io/badge/Mike's-Blog-darkgreen?link=https%3A%2F%2Fmikecrowley.us">
    </a>
    <a href="https://www.baselinetechnologies.com/leadership/">
        <img alt="Baseline Technologies"
            src="https://img.shields.io/badge/Baseline-Technologies-darkorange?link=https%3A%2F%2Fwww.baselinetechnologies.com">
    </a>
    <a href="https://github.com/Mike-Crowley/Public-Scripts">
        <img alt="Microsoft MVP"
            src="https://img.shields.io/badge/Microsoft_MVP-2010--2018-blue">
    </a>
    <a href="https://mikecrowley.files.wordpress.com/2020/06/8f158f9484a5cee37192077e0979564af679d0bb.asc">
        <img alt="Public PGP Key"
            src="https://img.shields.io/badge/PGP%2FGPG-Key-darkred?link=https%3A%2F%2Fmikecrowley.files.wordpress.com%2F2020%2F06%2F8f158f9484a5cee37192077e0979564af679d0bb.asc">
    </a>
    <a href="http://www.linkedin.com/in/mikecrowley">
        <img alt="linkedin.com/in/MikeCrowley"
            src="https://img.shields.io/badge/LinkedIn-mikecrowley-0077B5.svg?logo=LinkedIn">
    </a>
</p>

<p align="center">
    <br>
    <a href="https://www.buymeacoffee.com/mikecrowley" target="_blank">
        <img alt="Buy Me A Coffee" src="https://cdn.buymeacoffee.com/buttons/default-orange.png" height="41"
            width="174">
    </a>
</p>

# Public-Scripts Repository

<p align="right">
<img alt="GitHub License" src="https://img.shields.io/github/license/Mike-Crowley/Public-Scripts">
<img alt="GitHub top language" src="https://img.shields.io/github/languages/top/Mike-Crowley/Public-Scripts">
<img alt="GitHub commit activity" src="https://img.shields.io/github/commit-activity/t/Mike-Crowley/Public-Scripts">
<img alt="GitHub code size in bytes" src="https://img.shields.io/github/languages/code-size/Mike-Crowley/Public-Scripts">
</p>

PowerShell tools for identity security, Microsoft 365, and Windows infrastructure. Focused on Entra ID, OSINT, and practical automation for enterprise environments.

Many of these scripts were originally shared on the [TechNet Gallery](https://learn.microsoft.com/en-us/teamblog/technet-gallery-retirement) and [Microsoft forums](https://social.msdn.microsoft.com/Profile/mike%20crowley), where Mike was among the top contributors and earned [Microsoft MVP](https://mvp.microsoft.com/en-US/faq) status eight consecutive years (2010–2018).

## OSINT

+ ~~[Get-EntraCredentialType.ps1](./OSINT/Get-EntraCredentialType.ps1)~~ — *Deprecated. Use Get-EntraCredentialInfo instead.*

+ [Get-EntraCredentialInfo.ps1](./OSINT/Get-EntraCredentialInfo.ps1)

  + Query Entra for the CredentialType and openid-configuration of a user for a combined output.

+ [Request-FederationCerts.ps1](./OSINT/Request-FederationCerts.ps1)

  + Remotely query ADFS or Entra ID federation metadata to see information about the certificates being used. Supports both `-FarmFqdn` for ADFS servers and `-MetadataUrl` for Entra ID federation metadata endpoints.

+ [Get-AutoDetect.ps1](./OSINT/AutoDetect-AutoDiscover-v2)

  + Query AutoDiscover v2 / and the AutoDetect service (two files).

+ [Get-ExODomains.ps1](./OSINT/Get-ExODomains.ps1)

  + Query the domains in a tenant from the Exchange AutoDiscover service.

+ [Get-ShodanIpLookup.ps1](./OSINT/Get-ShodanIpLookup.ps1)

  + Query the Shodan database for information about an IP address, using the free InternetDB service or the paid API.

## Microsoft Graph

+ [Import-EntraHardwareOathToken.ps1](./MicrosoftGraph/Import-EntraHardwareOathToken.ps1)

  + Bulk-import hardware OATH tokens into Microsoft Entra ID via the Graph beta API. Automates the full create, assign, and activate workflow by computing TOTP verification codes from the seed, enabling seamless MFA provider migrations without end-user involvement. Depends on [Get-TOTP.ps1](./Utilities/Get-TOTP.ps1).

+ [Graph_SignInActivity_Report.ps1](./MicrosoftGraph/Graph_SignInActivity_Report.ps1)

  + Report on user SignInActivity and license detail via Invoke-RestMethod from Microsoft Graph.

+ [MailUser-MgUser-Activity-Report.ps1](./MicrosoftGraph/MailUser-MgUser-Activity-Report.ps1)

  + Export login information for mail users via Microsoft Graph.

+ [MgUserMail.ps1](./MicrosoftGraph/MgUserMail.ps1)

  + Send email via Microsoft Graph.

+ [Get-TeamsChatMessages.ps1](./MicrosoftGraph/Get-TeamsChatMessages.ps1)

  + Retrieve and export Microsoft Teams chat messages via Microsoft Graph. Supports delegated auth (your own chats) and app auth (any user's chats). Includes date filtering (relative, specific range, or sliding window), pagination, Out-GridView chat picker, and JSON/CSV/clipboard export. Also provides `-RegisterApp` to create the required Entra ID app registration.

## Okta

+ [Get-OktaSmsFactors.ps1](./Okta/Get-OktaSmsFactors.ps1)

  + Query the Okta Factors API to report on users with SMS-based MFA factors. Useful for SMS deprecation planning and MFA migration audits. Supports both direct Okta user enumeration and pre-filtered CSV input for large tenants. Handles API rate limiting automatically.

## Exchange

+ [Audit-ExoAppAccessPolicies.ps1](./Exchange/Audit-ExoAppAccessPolicies.ps1)

  + Audit all Exchange Online Application Access Policies and generate an HTML migration report with ready-to-run PowerShell commands for converting each policy to RBAC for Applications. Identifies orphaned apps, missing targets, permission gaps, and single-mailbox blockers.

+ [Get-AlternateMailboxes.ps1](./Exchange/Get-AlternateMailboxes.ps1)

  + Query Exchange Online AutoDiscover to enumerate mailbox delegates with modern auth.

+ [Get-AlternateMailboxes_BasicAuth.ps1](./Exchange/Get-AlternateMailboxes_BasicAuth.ps1)

  + Query Exchange Online AutoDiscover to enumerate mailbox delegates. The old basic auth version.

+ [RecipientReportv5.ps1](./Exchange/RecipientReportv5.ps1)

  + Dump all recipients and their email addresses (proxy addresses) to CSV.

+ [LowerCaseUPNs.ps1](./Exchange/LowerCaseUPNs.ps1)

  + Change Exchange user's email addresses to lowercase.

## Azure

+ [Find-AzFileShareDuplicates.ps1](./Azure/Find-AzFileShareDuplicates.ps1)

  + Find duplicate files in an Azure File Share by comparing MD5 hashes. Generates HTML, CSV, and JSON reports with metrics on potential wasted space. Requires Az.Accounts, Az.Storage, Az.Resources modules.

## SharePoint Online

+ [Find-DriveItemDuplicates.ps1](./SharePointOnline/Find-DriveItemDuplicates.ps1)

  + Identify duplicate files across OneDrive and SharePoint Online document libraries via Microsoft Graph. Uses dual-confidence detection: high confidence (quickXorHash match) and low confidence (filename match for files without hashes). Includes an interactive `-SitePicker` mode with Out-GridView for browsing SharePoint sites and OneDrive users, optional `-IncludeStorageMetrics` via the Graph Reports API, and generates an HTML dashboard with CSV/JSON exports in a timestamped folder on the desktop.

+ [Get-SPOStorageInsights.ps1](./SharePointOnline/Get-SPOStorageInsights.ps1)

  + Analyze SharePoint Online and OneDrive storage for duplicate files, version history bloat, and preservation hold libraries via Microsoft Graph. Supports multi-site scanning with interactive site picker, dual-confidence duplicate detection, and generates an HTML dashboard with CSV/JSON exports.

+ [Restore-FromRecycleBin.ps1](./SharePointOnline/Restore-FromRecycleBin.ps1)

  + Restore files from SPO recycle bin in bulk with logging and progress tracking.

## Active Directory

+ [Convert-GuidFormat.ps1](./AD_DS/Convert-GuidFormat.ps1)

  + Convert between standard GUID, ImmutableId (Base64), and hex byte representations. Auto-detects the input format and returns all three encodings — useful when cross-referencing objects across Azure AD Connect, Entra ID, and Active Directory.

+ [Resolve-AdAceGuid.ps1](./AD_DS/Resolve-AdAceGuid.ps1)

  + Resolve GUIDs on Active Directory ACEs to their friendly attribute, class, or extended right names. Builds a lookup hashtable from the schema and Extended-Rights container with just two LDAP queries, then resolves individual GUIDs locally — no per-ACE round-trips. Useful for auditing OU delegations and understanding who can write which attributes.

+ [Update-UseNotifyReplication.ps1](./AD_DS/Update-UseNotifyReplication.ps1)

  + Evaluate and optionally enable the Use_Notify option on AD Site Links and Replication Connections to reduce replication latency. Generates an HTML dashboard with health score, site link details (cost, schedule, sites), replication connections, and actionable recommendations. Use `-GetRegistrySettings` to query DC notification timers and AvoidPdcOnWan settings via WinRM.

+ [Get-ADSiteByIp.ps1](./AD_DS/Get-ADSiteByIp.ps1)

  + Determine which Active Directory site an IP address belongs to by querying AD subnets via LDAP. Requires the IP-Calc script from the PowerShell Gallery.

+ [Test-AdPassword.ps1](./AD_DS/Test-AdPassword.ps1)

  + Validate the credentials of a PSCredential object against the local machine or Active Directory domain.

## Windows

+ [RDPConnectionParser.ps1](./Windows/RDPConnectionParser.ps1)

  + Extract interactive (local and remote desktop) login information and save to CSV.

## Utilities

+ [Get-TOTP.ps1](./Utilities/Get-TOTP.ps1)

  + Generate Time-based One-Time Passwords (RFC 6238) and convert hex strings to Base32 (RFC 4648). Standalone functions with no external dependencies — useful for verifying hardware token output, automating OATH token activation, or converting vendor-supplied hex secrets for authenticator apps.

+ [Compare-ObjectsInVSCode.ps1](./Utilities/Compare-ObjectsInVSCode.ps1)

  + Compare two PowerShell Objects in Visual Studio Code.

+ [Convert-CaesarCipher.ps1](./Utilities/Convert-CaesarCipher.ps1)

  + Encode or decode case-sensitive English strings using the Caesar Cipher.

+ [Find-SensitiveInfo.ps1](./Utilities/Find-SensitiveInfo.ps1)

  + Check strings against Microsoft Purview DLP keyword and regex patterns to identify potentially sensitive data (IP addresses, account numbers, personal identifiers). Uses the same classification rules as Microsoft 365 DLP, exported from Exchange Online.

+ [Read-Fonts.ps1](./Utilities/Read-Fonts.ps1)

  + Extract font names from DOCX, XLSX, and PPTX files by parsing the embedded Office Open XML. Reports whether each font is installed on the local system. Requires PowerShell 7.

# Gists

A few smaller things live here: https://gist.github.com/Mike-Crowley

+ [Get-Superscript.ps1](https://gist.github.com/Mike-Crowley/b2a63bfe6bd533452bca3125037594a1)

  + Replace a given letter with the superscript letter.

+ [Get-WordscapesResults.ps1](https://gist.github.com/Mike-Crowley/09a03b770ab94af01147d4c7f9a10460)

  + Generate words for the wordscapes game so I can answer faster than my mom.

+ [Verify-SmbSigning.ps1](https://gist.github.com/Mike-Crowley/4aa9d0913ef0518e79034e5cdc56daf4)

  + Makes an SMB connection to a remote server, captures the traffic with Wireshark (tshark), and then parses the capture to report on the use of SMB signing.

---

<span style="font-size:large;">Visit https://mikecrowley.us/tag/powershell for additional functions and scripts.</span>
