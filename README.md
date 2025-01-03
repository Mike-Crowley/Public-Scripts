<span style="font-size:large;">Author: Mike Crowley</span>

<p align="left">
    <a href="https://mikecrowley.us">
        <img alt="Mike's Blog"
            src="https://img.shields.io/badge/Mike's-Blog-darkgreen?link=https%3A%2F%2Fmikecrowley.us">
    </a>
    <a href="https://www.baselinetechnologies.com">
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

Microsoft [retired the TechNet Gallery](https://learn.microsoft.com/en-us/teamblog/technet-gallery-retirement), so I've re-uploaded a few scripts that were formally posted here: https://social.msdn.microsoft.com/Profile/mike%20crowley

## Miscellaneous

+ [Get-AlternateMailboxes.ps1](./Get-AlternateMailboxes.ps1)

  + Query Exchange Online AutoDiscover to enumerate mailbox delegates with modern auth.

+ [Get-AlternateMailboxes_BasicAuth.ps1](./Get-AlternateMailboxes_BasicAuth.ps1)

  + Query Exchange Online AutoDiscover to enumerate mailbox delegates. The old basic auth version.

+ [Graph_SignInActivity_Report.ps1](./Graph_SignInActivity_Report.ps1)

  + Report on user SignInActivity and license detail via Invoke-RestMethod from Microsoft Graph.

+ [LowerCaseUPNs.ps1](./LowerCaseUPNs.ps1)

  + Change Exchange user's email addresses to lowercase.

+ [MailUser-MgUser-Activity-Report.ps1](./MailUser-MgUser-Activity-Report.ps1)

  + Get Export login information for mail users via Microsoft Graph.

+ [MgUserMail.ps1](./MgUserMail.ps1)

  + Send email via Microsoft Graph.

+ [RDPConnectionParser.ps1](./RDPConnectionParser.ps1)

  + Extract interactive (local and remote desktop) login information and save to CSV.

+ [RecipientReportv5.ps1](./RecipientReportv5.ps1)

  + Dump all recipients and their email addresses (proxy addresses) to CSV.

+ [Restore-FromRecycleBin.ps1](./Restore-FromRecycleBin.ps1)

  + Restore files from SPO in bulk.
  
+ [Find-DriveItemDuplicates.ps1](./Find-DriveItemDuplicates.ps1)

  + A utility to find duplicate files on OneDrive for Business or SharePoint Online.

+ [Compare-ObjectsInVSCode.ps1](./Compare-ObjectsInVSCode.ps1)

  + Compare two PowerShell Objects in Visual Studio Code.

## OSINT

+ [Get-EntraCredentialType.ps1](./OSINT/Get-EntraCredentialType.ps1)

  + Query Entra for the CredentialType of a user.

+ [Get-EntraCredentialInfo.ps1](./OSINT/Get-EntraCredentialInfo.ps1)

  + Query Entra for the CredentialType and openid-configuration of a user for a combined output.

+ [Request-AdfsCerts.ps1](./OSINT/Request-AdfsCerts.ps1)

  + Remotley query ADFS to see information about the certificates it is using.

+ [Get-AutoDetect.ps1](./OSINT/AutoDetect-AutoDiscover-v2)

  + Query AutoDiscover v2 / and the AutoDetect service (two files).

+ [Get-ExODomains.ps1](./OSINT/Get-ExODomains.ps1)

  + Query the domains in a tenant from the Exchange AutoDiscover service.

# Gists

<p align="right">  
<img alt="GitHub License" src="https://img.shields.io/github/license/Mike-Crowley/Public-Scripts">
<img alt="GitHub top language" src="https://img.shields.io/github/languages/top/Mike-Crowley/Public-Scripts">
</p>

There are also a few things over here: https://gist.github.com/Mike-Crowley

+ [Get-ADSiteByIp.ps1](https://gist.github.com/Mike-Crowley/3ad9472a2ab365c723f2272da197eabf)

  + Enter an IP address and this will lookup the AD site to which it belongs.

+ [Test-AdPassword.ps1](https://gist.github.com/Mike-Crowley/0cfaf1a8733b530e8f00acb59dec771f)

  + Determine if an AD user's password is valid.

+ [Get-Superscript.ps1](https://gist.github.com/Mike-Crowley/b2a63bfe6bd533452bca3125037594a1)

  + Replace a given letter with the superscript letter.

+ [Get-ShodanIpLookup](https://gist.github.com/Mike-Crowley/ff3c432ad921799b736b45dff828acca)

  + Query the Shodan database for an IP address with or without an API key

+ [Get-WordscapesResults.ps1](https://gist.github.com/Mike-Crowley/09a03b770ab94af01147d4c7f9a10460)

  + Generate words for the wordscapes game so I can answer faster than my mom.

+ [Verify-SmbSigning.ps1](https://gist.github.com/Mike-Crowley/4aa9d0913ef0518e79034e5cdc56daf4)

  + Makes an SMB connection to a remote server, captures the traffic with Wireshark (tshark), and then parses the capture to report on the use of SMB signing.

#

<span style="font-size:large;">Be sure to read the comments in the scripts themselves for more detail!</span>

<span style="font-size:large;">Visit https://mikecrowley.us/tag/powershell for additional functions and scripts.</span>
