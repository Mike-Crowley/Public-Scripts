
function Compare-ObjectsInVSCode {
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

    $json1 | Out-File -FilePath $file1Path
    $json2 | Out-File -FilePath $file2Path

    # Open files in VS Code for comparison
    code -d $file1Path $file2Path
}

<# Example
    $Process1 = Get-Process mspaint
    $Process2 = Get-Process excel
    Compare-ObjectsInVSCode $Process1 $Process2 -Depth 2
#>
