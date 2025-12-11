# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Repository Overview

This is a collection of standalone PowerShell scripts for Microsoft 365 administration, OSINT, and Windows management. Scripts are authored by Mike Crowley (Microsoft MVP 2010-2018). There is no build system, test framework, or package management - each script is self-contained.

## Repository Structure

- **Root directory**: Main utility scripts for Exchange Online, Microsoft Graph, and general administration
- **OSINT/**: Open-source intelligence scripts for querying Entra ID, ADFS, and Exchange autodiscover endpoints
- **AD_DS/**: Active Directory Domain Services scripts (e.g., replication notification configuration)
- **SupportingFiles/**: Reference data files (CSV/JSON) used by some scripts

## Common Patterns

### Script Structure
Scripts follow standard PowerShell conventions:
- Comment-based help with `.SYNOPSIS`, `.DESCRIPTION`, `.EXAMPLE`, `.LINK`
- Functions use `Verb-Noun` naming (e.g., `Get-EntraCredentialType`, `Find-DriveItemDuplicates`)
- Parameters with validation attributes (`[ValidateScript()]`, `[ValidateSet()]`, `[ValidateRange()]`)
- Many scripts require pre-connection to Microsoft Graph (`Connect-MgGraph`)

### Microsoft Graph Integration
Several scripts use the Microsoft Graph PowerShell SDK:
- Require `Microsoft.Graph.Authentication` module
- Use `Invoke-MgGraphRequest` for REST API calls
- Require appropriate permission scopes (e.g., `Files.Read`, `Sites.Read`)

### Output Patterns
Scripts typically output:
- `[pscustomobject]` for structured data
- CSV/JSON exports to Desktop for reports
- Pipeline-friendly output for further processing

## Running Scripts

Scripts are run directly in PowerShell. Example:
```powershell
# Dot-source to load function into session
. .\Compare-ObjectsInVSCode.ps1

# Then call the function
Compare-ObjectsInVSCode $Object1 $Object2 -Depth 2
```

For Graph-dependent scripts:
```powershell
Connect-MgGraph -Scopes Files.Read
Find-DriveItemDuplicates -RootPath "Desktop" -OutputStyle Report
```
