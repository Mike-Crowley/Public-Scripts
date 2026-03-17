function Find-SensitiveInfo {
    <#
    .SYNOPSIS
        Identifies sensitive information in content using Microsoft DLP rule patterns.

    .DESCRIPTION
        Checks input strings against Microsoft's predefined DLP keywords and regex patterns
        to identify potentially sensitive data such as IP addresses, account numbers, driver
        license numbers, and personal identifiers.

        Pattern sources are extracted from the Microsoft Rule Package, which contains the
        same classification rules used by Microsoft Purview DLP. The function matches input
        against both keyword lists and regex patterns, returning which rules matched.

        Note: Microsoft also uses internal functions (e.g., Func_credit_card) for some
        sensitive information types that involve complex validation beyond simple regex.
        These function-based detections are not replicated here.

    .PARAMETER ContentToCheck
        One or more strings to check against DLP patterns. Accepts pipeline input.

    .PARAMETER RulePackagePath
        Path to the ClassificationRuleCollectionXml file exported from Exchange Online.
        If not specified, defaults to ClassificationRuleCollectionXml.xml in the same
        directory as this script.

        To generate this file, run the following in an Exchange Online PowerShell session:

            Connect-IPPSSession
            $xml = (Get-DlpSensitiveInformationTypeRulePackage |
                Where-Object RuleCollectionName -eq 'Microsoft Rule Package'
            ).ClassificationRuleCollectionXml
            [xml]$xml | Export-Clixml .\ClassificationRuleCollectionXml.xml

    .EXAMPLE
        Find-SensitiveInfo "1.2.3.4"

        Content  KeywordMatches RegexMatches
        -------  -------------- ------------
        1.2.3.4                 Regex_ipv4_address

    .EXAMPLE
        Find-SensitiveInfo "Date of Birth"

        Content       KeywordMatches             RegexMatches
        -------       --------------             ------------
        Date of Birth Keyword_us_drivers_license

    .EXAMPLE
        "5.6.7.8", "Driv Lic", "hello" | Find-SensitiveInfo

        Checks multiple strings from the pipeline against all DLP patterns.

    .INPUTS
        System.String

    .OUTPUTS
        PSCustomObject with properties: Content, KeywordMatches, RegexMatches

    .NOTES
        Author:  Mike Crowley
        https://mikecrowley.us

    .LINK
        https://learn.microsoft.com/en-us/microsoft-365/compliance/sit-functions
    #>

    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory, Position = 0, ValueFromPipeline)]
        [ValidateNotNullOrEmpty()]
        [string[]]$ContentToCheck,

        [Parameter()]
        [string]$RulePackagePath
    )

    begin {
        if (-not $RulePackagePath) {
            $RulePackagePath = Join-Path $PSScriptRoot 'ClassificationRuleCollectionXml.xml'
        }

        if (-not (Test-Path $RulePackagePath)) {
            throw "Rule package not found at '$RulePackagePath'. See Get-Help Find-SensitiveInfo -Parameter RulePackagePath for export instructions."
        }

        [xml]$ClassificationRuleCollectionXml = Import-Clixml $RulePackagePath
        $Keywords = $ClassificationRuleCollectionXml.RulePackage.Rules.Keyword
        $RegexPatterns = $ClassificationRuleCollectionXml.RulePackage.Rules.Regex
    }

    process {
        foreach ($Element in $ContentToCheck) {
            $KeywordMatches = $Keywords | Where-Object { $_.group.term -contains $Element }
            $RegexMatches = $RegexPatterns | Where-Object { $Element -match $_.'#text' }

            [PSCustomObject]@{
                Content        = $Element
                KeywordMatches = $KeywordMatches.id -join '; '
                RegexMatches   = $RegexMatches.id -join '; '
            }
        }
    }
}
