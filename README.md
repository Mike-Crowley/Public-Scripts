# Public-Scripts Repo

Microsoft [retired the TechNet Gallery](https://docs.microsoft.com/en-us/teamblog/technet-gallery-retirement), so I've re-uploaded a few scripts that were formally posted here: https://social.msdn.microsoft.com/Profile/mike%20crowley


+ [Get-AlternateMailboxes.ps1](https://github.com/Mike-Crowley/Public-Scripts/blob/main/Get-AlternateMailboxes.ps1)

  - Query Exchange Online AutoDiscover to enumerate mailbox delegates with modern auth.

+ [Get-AlternateMailboxes_BasicAuth.ps1](https://github.com/Mike-Crowley/Public-Scripts/blob/main/Get-AlternateMailboxes_BasicAuth.ps1)

  - Query Exchange Online AutoDiscover to enumerate mailbox delegates. The old basic auth version.

+ [Graph_SignInActivity_Report.ps1](https://github.com/Mike-Crowley/Public-Scripts/blob/main/Graph_SignInActivity_Report.ps1)

  - Report on user SignInActivity and license detail via Invoke-RestMethod from Microsoft Graph.

+ [LowerCaseUPNs.ps1](https://github.com/Mike-Crowley/Public-Scripts/blob/main/LowerCaseUPNs.ps1)

  - Change Exchange user's email addresses to lowercase.

+ [MailUser-MgUser-Activity-Report.ps1](https://github.com/Mike-Crowley/Public-Scripts/blob/main/MailUser-MgUser-Activity-Report.ps1)

  - Get Export login information for mail users via Microsoft Graph.

+ [MgUserMail.ps1](https://github.com/Mike-Crowley/Public-Scripts/blob/main/MgUserMail.ps1)

  - Send email via Microsoft Graph.

+ [RDPConnectionParser.ps1](https://github.com/Mike-Crowley/Public-Scripts/blob/main/RDPConnectionParser.ps1)

  - Extract interactive (local and remote desktop) login information and save to CSV.

+ [RecipientReportv5.ps1](https://github.com/Mike-Crowley/Public-Scripts/blob/main/RecipientReportv5.ps1)

  - Dump all recipients and their email addresses (proxy addresses) to CSV.

+ [Request-AdfsCerts.ps1](https://github.com/Mike-Crowley/Public-Scripts/blob/main/Request-AdfsCerts.ps1)

  - Remotley query ADFS to see information about the certificates it is using.

+ [Restore-FromRecycleBin.ps1](https://github.com/Mike-Crowley/Public-Scripts/blob/main/Restore-FromRecycleBin.ps1)

  - Restore files from SPO in bulk.


# Gists

There are also a few things over here: https://gist.github.com/Mike-Crowley


+ [Get-ADSiteByIp.ps1](https://gist.github.com/Mike-Crowley/3ad9472a2ab365c723f2272da197eabf)

  - Enter an IP address and this will lookup the AD site to which it belongs.

+ [Get-AutoDetect.ps1](https://gist.github.com/Mike-Crowley/521680c3f84105378d2eb2358bd539cf)

  - Query AutoDiscover v2 / and the AutoDetect service (two files).

+ [Get-ExODomains.ps1](https://gist.github.com/Mike-Crowley/5da3f3fd69519f06866d580ebbd5b5b7)

  - Query the domains in a tenant from the Exchange AutoDiscover service.

+ [Test-AdPassword.ps1](https://gist.github.com/Mike-Crowley/0cfaf1a8733b530e8f00acb59dec771f)

  - Determine if an AD user's password is valid.

+ [Get-Superscript.ps1](https://gist.github.com/Mike-Crowley/b2a63bfe6bd533452bca3125037594a1)

  -  Replace a given letter with the superscript letter.

+ [Get-ShodanIpLookup](https://gist.github.com/Mike-Crowley/ff3c432ad921799b736b45dff828acca)

  -  Query the Shodan database for an IP address with or without an API key

+ [Get-WordscapesResults.ps1](https://gist.github.com/Mike-Crowley/09a03b770ab94af01147d4c7f9a10460)

  - Generate words for the wordscapes game so I can answer faster than my mom.


#### Be sure to read the comments in the scripts themselves for more detail!
