<#
.SYNOPSIS
    Compares two PowerShell objects side-by-side in Visual Studio Code.

.DESCRIPTION
    Converts two objects to JSON and opens them in VS Code's diff viewer for
    visual comparison. Useful for comparing configuration objects, API responses,
    or any PowerShell objects.

.PARAMETER Object1
    The first object to compare.

.PARAMETER Object2
    The second object to compare.

.PARAMETER Depth
    The JSON serialization depth (1-10). Defaults to 1. Increase for deeply nested objects.

.EXAMPLE
    . .\Compare-ObjectsInVSCode.ps1
    $Process1 = Get-Process mspaint
    $Process2 = Get-Process excel
    Compare-ObjectsInVSCode $Process1 $Process2 -Depth 2

.NOTES
    Author: Mike Crowley
    https://mikecrowley.us

    Requires: Visual Studio Code (code.exe) in PATH

.LINK
    https://github.com/Mike-Crowley/Public-Scripts
#>

function Compare-ObjectsInVSCode {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [PSObject]$Object1,

        [Parameter(Mandatory = $true)]
        [PSObject]$Object2,

        [Parameter(Mandatory = $false)]
        [ValidateRange(1, 10)]
        [int]$Depth = 1
    )

    if (-not (Get-Command code -ErrorAction SilentlyContinue)) {
        Write-Error "Visual Studio Code couldn't be found."
        return
    }

    $tempDir = $env:TMP
    $file1Path = Join-Path -Path $tempDir -ChildPath "object1.json"
    $file2Path = Join-Path -Path $tempDir -ChildPath "object2.json"

    $json1 = $Object1 | ConvertTo-Json -Depth $Depth
    $json2 = $Object2 | ConvertTo-Json -Depth $Depth

    try {
        $json1 | Out-File -FilePath $file1Path -ErrorAction Stop
        $json2 | Out-File -FilePath $file2Path -ErrorAction Stop
    }
    catch {
        Write-Error "Failed to write temporary files: $_"
        return
    }

    # Open files in VS Code for comparison
    code -d $file1Path $file2Path
}
