function Read-Fonts {
    <#
    .SYNOPSIS
        Extracts font information from DOCX, XLSX, and PPTX files.

    .DESCRIPTION
        Reads font names used within Microsoft Office Open XML files by extracting and
        parsing the embedded XML. For each font found, indicates whether it is installed
        on the local system.

        Supported file types:
          - DOCX — reads word/fontTable.xml
          - XLSX — reads xl/styles.xml
          - PPTX — parses typeface attributes from slide XML files

    .PARAMETER Filepath
        Path to a .docx, .xlsx, or .pptx file. Accepts pipeline input.

    .EXAMPLE
        Read-Fonts 'C:\Reports\quarterly.docx'

        Name              SystemFont
        ----              ----------
        Calibri           True
        Cambria           True
        Wingdings         True

    .EXAMPLE
        Read-Fonts 'C:\Data\budget.xlsx'

        Extracts fonts from an Excel workbook's style definitions.

    .EXAMPLE
        Get-ChildItem *.pptx | ForEach-Object { Read-Fonts $_.FullName }

        Extracts fonts from all PowerPoint files in the current directory.

    .INPUTS
        System.String

    .OUTPUTS
        PSCustomObject with properties: Name, SystemFont

    .NOTES
        Author:  Mike Crowley
        https://mikecrowley.us

        Requires PowerShell 7 or later due to Expand-Archive behavior.

        The SystemFont check uses locally installed fonts only (approximately 256 on a
        default Windows installation). Fonts bundled with Microsoft Office are not
        included in this check and will show SystemFont = False even if available.

    .LINK
        https://learn.microsoft.com/en-us/typography/fonts/windows_11_font_list
    #>

    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory, Position = 0, ValueFromPipeline)]
        [ValidateScript({
                if ($_ -match '\.(docx|xlsx|pptx)$') { $true }
                else { throw "Filepath must be a .docx, .xlsx, or .pptx file." }
            })]
        [string]$Filepath
    )

    process {
        if ((Get-Command Expand-Archive).Version -gt [version]'1.0.1.0') {
            if (Test-Path $Filepath) {

                $TempPath = "$env:TEMP\Read-Fonts_$(Get-Random)"
                Expand-Archive $Filepath -DestinationPath $TempPath -Force

                if ($Filepath -like '*.docx') {
                    [xml]$FileXML = Get-Content ($TempPath + "\word\fontTable.xml")
                    $FontList = $FileXML.fonts.font.name
                }
                elseif ($Filepath -like '*.xlsx') {
                    [xml]$FileXML = Get-Content ($TempPath + "\xl\styles.xml")
                    $FontList = $FileXML.styleSheet.fonts.font.name.val
                }
                elseif ($Filepath -like '*.pptx') {
                    $SlideXMLs = Get-ChildItem ($TempPath + "\ppt\slides") -Filter *.xml
                    $FontList = $SlideXMLs | ForEach-Object {
                        $FileContent = Get-Content $_.FullName
                        $pattern = 'typeface="([^"]+)"'
                        $regexMatches = [regex]::Matches($FileContent, $pattern)
                        foreach ($m in $regexMatches) {
                            $m.Groups[1].Value
                        }
                    } | Sort-Object -Unique
                }

                Remove-Item $TempPath -Recurse -Force

                $SystemFonts = ([System.Drawing.Text.InstalledFontCollection]::new()).Families

                $FontList | ForEach-Object {
                    [PSCustomObject]@{
                        Name       = $_
                        SystemFont = $_ -in $SystemFonts.Name
                    }
                } | Sort-Object Name
            }
            else { Write-Error "Cannot find file: $Filepath" }
        }
        else { Write-Error "Requires PowerShell 7 or later (Expand-Archive version too old)." }
    }
}
