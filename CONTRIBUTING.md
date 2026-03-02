# Contributing

Thanks for your interest in this project! This is a personal collection of PowerShell scripts for Microsoft 365 administration, OSINT, and Windows infrastructure. Contributions are welcome, but please keep the following in mind.

## Guidelines

- **Bug fixes and improvements** to existing scripts are appreciated. Please open an issue first to discuss the change.
- **New scripts** should be standalone, follow PowerShell `Verb-Noun` naming, and include full comment-based help (`.SYNOPSIS`, `.DESCRIPTION`, `.EXAMPLE`, `.NOTES`, `.LINK`).
- **Pull requests** should target a single script or a closely related set of changes. Keep PRs focused and easy to review.
- **Code style** should match existing scripts: `[CmdletBinding()]`, proper parameter validation, and `[pscustomobject]` output where appropriate.

## Before Submitting

1. Test your script in a real environment (or describe your test setup in the PR).
2. Ensure comment-based help is complete and accurate.
3. Do not include secrets, API keys, or credentials in your submission.

## Expectations

This is a side project, so response times may vary. Not every PR will be merged, but all feedback is valued.
