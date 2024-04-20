<#
.DESCRIPTION
    This script finds duplicate files via Microsoft Graph by comparing their hashes.
    -It requires the manual modification of the $upn variable, which defines who's OneDrive will be searched.
    -It also requires a modification to the Desktop/DupesDirectory text, if another directory path is to be searched.
    -If using the test files, wait for them to sync before attempting to detect them.

    To do: Convert to a function

.LINK

    https://mikecrowley.us/2024/04/20/onedrive-and-sharepoint-online-file-deduplication-report-microsoft-graph-api
#>


<#
    Create some duplicate files for testing

    $Desktop = [Environment]::GetFolderPath("Desktop")
    $TestDir = mkdir $Desktop\DupesDirectory -Force
    $TestLogFile = (Invoke-WebRequest "https://gist.githubusercontent.com/Mike-Crowley/d4275d6abd78ad8d19a6f1bcf9671ec4/raw/66fe537cfe8e58b1a5eb1c1336c4fdf6a9f05145/log.log.log").content
    1..25 | ForEach-Object { $TestLogFile | Out-File "$TestDir\$(Get-Random).log" }

    Create more, if you'd like:

    1..25 | ForEach-Object { "Hello World 1" | Out-File "$TestDir\$(Get-Random).log" }
    1..25 | ForEach-Object { "Hello World 2" | Out-File "$TestDir\$(Get-Random).log" }

    # Create some non-duplicate files
    1..25  | ForEach-Object { Get-Random | Out-File "$TestDir\$(Get-Random).log" }

    !! Wait for the files to sync via KFM !!
#>

# Which User's OneDrive we're searching:
$Upn = "mike@mikecrowley.fake"

# Which folder path in the OneDrive we're searching
$FolderPath = "Desktop/DupesDirectory"

##################################################

# Connect to Graph
Connect-MgGraph -NoWelcome

$Drive = Invoke-MgGraphRequest -Uri "beta/users/$upn/drive"

# Find the files
$uri = "beta/drives/$($Drive.id)/root:/$($FolderPath):/children"
$AllChildren = [Collections.Generic.List[Object]]::new()
do {
    $PageResults = Invoke-MgGraphRequest -Uri $uri
    if ($PageResults.value) {
        $AllChildren.AddRange($PageResults.value)
    }
    else {
        $AllChildren.Add($PageResults)
    }
    $uri = $PageResults.'@odata.nextlink'
} until (-not $uri)
$Files = $AllChildren | Where-Object { $null -ne $_.file } # remove non-files such as folders

# Calculate the properties
$FilesCustom = foreach ($DriveItem in $Files) {
    [pscustomobject] @{
        name                 = $DriveItem.name
        lastModifiedDateTime = $DriveItem.lastModifiedDateTime
        quickXorHash         = $DriveItem.file.hashes.quickXorHash
        size                 = $DriveItem.size
        webUrl               = $DriveItem.webUrl
    }
}

# Create the groups of dupes
$GroupsOfDupes = $FilesCustom | Where-Object { $null -ne $_.quickXorHash } | Group-Object quickXorHash | Where-Object count -ge 2

$Report = foreach ($Group in $GroupsOfDupes) {
    [pscustomobject] @{
        QuickXorHash  = $Group.Name
        FileSizeKB    = $Group.Group.size[0] / 1KB
        NumberOfFiles = $Group.Count
        FileNames     = ($Group.Group.name | Sort-Object -Unique) -join ';'
        WebLocalPaths = ($Group.Group.webUrl | ForEach-Object { ([uri]$_).LocalPath }) -join ";"
    }
}

# Create report
Write-Host "Found $(($GroupsOfDupes | Measure-Object).count) group(s) of duplicate files. See desktop reports for details." -ForegroundColor Cyan
$Desktop = [Environment]::GetFolderPath("Desktop")
$Report | Export-Csv $Desktop\DupeReport_csv.csv -NoTypeInformation
$Report | ConvertTo-Json | Out-File $Desktop\DupeReport_json.json # Possibly easier to read
