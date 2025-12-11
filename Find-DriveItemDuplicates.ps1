<#
.DESCRIPTION
    Find-DriveItemDuplicates examines the hash values of OneDrive files to determine if there are duplicates.
    Duplicates are reported to the desktop, the pipeline, or both.

    Connect to Graph prior to running Find-DriveItemDuplicates.
    e.g.
        Connect-MgGraph -Scopes Files.Read

    If you are checking files outside of your own OneDrive, you'll want to register an app with the appropriate permissions.
    Reference:
        https://learn.microsoft.com/en-us/powershell/microsoftgraph/app-only?view=graph-powershell-1.0

.EXAMPLE

    Find-DriveItemDuplicates

    # Runs with parameter defaults:
        # UPN:         Current user
        # RootPath:    Root of drive
        # NoRecursion: False (It will recurse through folders)
        # OutputStyle: ReportAndPassThru
        # ResultSize:  32767
        # Silent:      False (It will write comments to the screen)

.EXAMPLE

    Find-DriveItemDuplicates -RootPath "Desktop" -OutputStyle Report -ResultSize 500

    # Searches the current user's desktop folder at the root of the drive and its subfolders.
    # Checks up to 500 files and generates the desktop report only (good for testing).

.EXAMPLE

    Find-DriveItemDuplicates -Upn user1@example.com -RootPath "Desktop/DupesDirectory" -OutputStyle Report -NoRecursion

    # Searches the user1's Desktop/DupesDirectory folder at the root of the drive.
    # Does not check subfolders. Creates the reports but does not return the items to the pipeline.

.NOTES

    Create some duplicate files for testing if needed.

        $Desktop = [Environment]::GetFolderPath("Desktop")
        $TestDir = mkdir $Desktop\DupesDirectory -Force
        $TestLogFile = (Invoke-WebRequest "https://gist.githubusercontent.com/Mike-Crowley/d4275d6abd78ad8d19a6f1bcf9671ec4/raw/66fe537cfe8e58b1a5eb1c1336c4fdf6a9f05145/log.log.log").content
        1..25 | ForEach-Object { $TestLogFile | Out-File "$TestDir\$(Get-Random).log" }

    Create more, if you'd like.

        1..25 | ForEach-Object { "Hello World 1" | Out-File "$TestDir\$(Get-Random).log" }
        1..25 | ForEach-Object { "Hello World 2" | Out-File "$TestDir\$(Get-Random).log" }

    # Create some non-duplicate files.

        1..25  | ForEach-Object { Get-Random | Out-File "$TestDir\$(Get-Random).log" }

    !! Wait for the files to sync via OneDrive's sync client (If using Known Folder Move - KFM) !!

.LINK

    https://mikecrowley.us/2024/04/20/onedrive-and-sharepoint-online-file-deduplication-report-microsoft-graph-api
#>

function Find-DriveItemDuplicates {
    param(
        [ValidateScript(
            {
                if ($_ -match "^\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*$") { $true } else { throw "Invalid UPN." }
            }
        )]
        [string]$Upn = (whoami /upn),
        [string]$RootPath = "/",
        [switch]$NoRecursion = $false,
        [ValidateSet(
            "Report", "PassThru", "ReportAndPassThru"
        )]
        [string]$OutputStyle = "ReportAndPassThru",
        [int32]$ResultSize = [int16]::MaxValue,
        [switch]$Silent = $false
    )
    $StartTime = Get-Date

    # Some pre-run checks
    if ($null -eq (Get-Command Invoke-MgGraphRequest) ) { Throw "Invoke-MgGraphRequest cmdlet not found. Install the Microsoft.Graph.Authentication PowerShell module. `nhttps://learn.microsoft.com/en-us/graph/sdks/sdk-installation#install-the-microsoft-graph-powershell-sdk" }
    if ($null -eq (Get-MgContext)) { Throw "No Graph context found. Please call Connect-MgGraph." }
    if (
        (((Get-MgContext).Scopes | Out-String ) -notlike '*Files.Read*') -and
        (((Get-MgContext).Scopes | Out-String ) -notlike '*Sites.Read*')
    ) { Write-Warning "Permission scope may be missing. `nhttps://learn.microsoft.com/en-us/graph/api/driveitem-list-children?view=graph-rest-beta&tabs=http#permissions" }

    #Find the user's drive
    $Drive = Invoke-MgGraphRequest -Uri "beta/users/$upn/drive"
    if ($Drive.webUrl -notlike 'https*') { Throw "Drive not found." }

    if (-not $Silent) { Write-Host "`nFound Drive: " -ForegroundColor Cyan -NoNewline }
    if (-not $Silent) { Write-Host $Drive.webUrl -ForegroundColor DarkCyan }
    if (-not $Silent) { Write-Host "" }

    $baseUri = "beta/drives/$($Drive.id)/root"

    # Create a new list to hold all file objects
    $FileList = [Collections.Generic.List[Object]]::new()

    # To do:
    # Normalize if/else - move shared components above if/else
    # Implement better graph error handling

    if ($NoRecursion) {
        $uri = "beta/drives/$($Drive.id)/root:/$($RootPath):/children"

        do {
            if (-not $Silent) { Write-Host "Searching: " -ForegroundColor Cyan -NoNewline }
            if (-not $Silent) { Write-Host $uri  -ForegroundColor DarkCyan }
            $PageResults = Invoke-MgGraphRequest -Uri $uri
            if ($PageResults.value) {
                $FileList.AddRange($PageResults.value)
            }
            else {
                $FileList.Add($PageResults)
            }
            $uri = $PageResults.'@odata.nextlink'
        } until (-not $uri)
        $FileList = $FileList | Where-Object { $null -ne $_.file } # remove non-files such as folders

        $FileList = foreach ($DriveItem in $FileList) {
            [pscustomobject] @{
                name                 = $DriveItem.name
                lastModifiedDateTime = $DriveItem.lastModifiedDateTime
                quickXorHash         = $DriveItem.file.hashes.quickXorHash
                size                 = $DriveItem.size
                webUrl               = $DriveItem.webUrl
            }
        }
    }
    else {
        # Internal function to recursively get files
        function Get-FolderItemsRecursively {
            param(
                [string]$Path,
                [Collections.Generic.List[Object]]$AllFiles
            )

            if ($path -eq "/") {
                $uri = "$baseUri/children"
            }
            else {
                $uri = "$baseUri`:/$path`:/children"
            }

            if ($AllFiles.Count -lt $ResultSize) {
                do {
                    if (-not $Silent) { Write-Host "Searching: " -ForegroundColor Cyan -NoNewline }
                    if (-not $Silent) { Write-Host $path  -ForegroundColor DarkCyan }

                    if ($AllFiles.count -gt 1) {
                        if (-not $Silent) { Write-Host "Files Evaluated: " -ForegroundColor Cyan -NoNewline }
                        if (-not $Silent) { Write-Host  $AllFiles.count -ForegroundColor DarkCyan }
                    }

                    $Response = Invoke-MgGraphRequest -Uri $Uri

                    foreach ($Item in $Response.value) {
                        if ($null -ne $Item.folder) {
                            # Recursive call if the item is a folder
                            $FolderPath = "$Path/$($Item.name)"
                            Get-FolderItemsRecursively -Path $FolderPath -AllFiles $AllFiles
                        }
                        elseif ($null -ne $Item.file) {
                            # Add to the list if the item is a file
                            $AllFiles.Add(
                                [pscustomobject]@{
                                    Name                 = $Item.name
                                    LastModifiedDateTime = $Item.lastModifiedDateTime
                                    QuickXorHash         = $Item.file.hashes.quickXorHash
                                    Size                 = $Item.size
                                    WebUrl               = $Item.webUrl
                                }
                            )
                        }
                    }
                    # Pagination handling
                    $Uri = $Response.'@odata.nextLink'
                } until ((-not $Uri) -or ($AllFiles.Count -ge $ResultSize))
            }
        }
        # Start recursion from the root Path
        Get-FolderItemsRecursively -Path $RootPath -AllFiles $FileList
    }

    # Create the groups of dupes. The last graph page may have returned more than the $ResultSize, so limit output as needed.
    $GroupsOfDupes = $FileList[0..($ResultSize - 1)] | Where-Object { $null -ne $_.quickXorHash } | Group-Object quickXorHash | Where-Object count -ge 2

    # Select final columns for output
    $Output = foreach ($Group in $GroupsOfDupes) {
        $FileGroupSize = ($Group.Group.size | Measure-Object -Sum).Sum
        [pscustomobject] @{
            QuickXorHash    = $Group.Name
            NumberOfFiles   = $Group.Count
            FileSizeKB      = $Group.Group.size[0] / 1KB
            FileGroupSizeKB = $FileGroupSize / 1KB
            PossibleWasteKB = ( $FileGroupSize - $Group.Group.size[0] ) / 1KB
            FileNames       = ($Group.Group.name | Sort-Object -Unique) -join ';'
            WebLocalPaths   = ($Group.Group.webUrl | ForEach-Object { ([uri]$_).LocalPath }) -join ";"
        }
    }

    # Create pipeline output if requested
    if (($OutputStyle -eq "PassThru") -or ($OutputStyle -eq "ReportAndPassThru")) {
        $Output
    }

    # Create reports if requested
    if (($OutputStyle -eq "Report") -or ($OutputStyle -eq "ReportAndPassThru")) {
        $FileDate = Get-Date -Format 'ddMMMyyyy_HHmm.s'
        $Desktop = [Environment]::GetFolderPath("Desktop")
        $CsvOutputPath = "$Desktop\$UPN-DupeReport-$FileDate.csv"
        $JsonOutputPath = $CsvOutputPath -replace ".csv", ".json"

        $Output | Export-Csv $CsvOutputPath -NoTypeInformation -Encoding UTF8
        $Output | ConvertTo-Json | Out-File $JsonOutputPath -Encoding UTF8
    }

    # Report status to console
    $EndTime = Get-Date
    if (-not $Silent) { Write-Host "`nJob Duration: " -ForegroundColor Cyan }
    if (-not $Silent) { Write-Host " Job Start: " -ForegroundColor Cyan -NoNewline }
    if (-not $Silent) { Write-Host $StartTime -ForegroundColor DarkCyan }
    if (-not $Silent) { Write-Host " Job End: " -ForegroundColor Cyan -NoNewline }
    if (-not $Silent) { Write-Host $EndTime -ForegroundColor DarkCyan }
    if (-not $Silent) { Write-Host " Total Seconds: " -ForegroundColor Cyan -NoNewline }
    if (-not $Silent) { Write-Host $([math]::Ceiling(($EndTime - $StartTime).TotalSeconds)) -ForegroundColor DarkCyan }

    if (-not $Silent) { Write-Host "`nJob Parameters: " -ForegroundColor Cyan }
    if (-not $Silent) { Write-Host " ResultSize: " -ForegroundColor Cyan -NoNewline }
    if (-not $Silent) { Write-Host $ResultSize -ForegroundColor DarkCyan }

    if (-not $Silent) { Write-Host " Recurse: " -ForegroundColor Cyan -NoNewline }
    if (-not $Silent) { Write-Host (-not $NoRecursion) -ForegroundColor DarkCyan }

    if (-not $Silent) { Write-Host " OutputStyle: " -ForegroundColor Cyan -NoNewline }
    if (-not $Silent) { Write-Host $OutputStyle -ForegroundColor DarkCyan }

    if (-not $Silent) { Write-Host "`nTotal Files Evaluated: " -ForegroundColor Cyan -NoNewline }
    if (-not $Silent) { Write-Host "$(($FileList[0..($ResultSize - 1)] ).count)" -ForegroundColor DarkCyan }

    if (-not $Silent) { Write-Host "`nFound "  -ForegroundColor Cyan -NoNewline }
    if (-not $Silent) { Write-Host "$($($GroupsOfDupes | Measure-Object).count)" -ForegroundColor DarkCyan -NoNewline }
    if (-not $Silent) { Write-Host " group(s) of duplicate files." -ForegroundColor Cyan  -NoNewline }

    if ($($GroupsOfDupes | Measure-Object).count -ge 1) {
        if (-not $Silent) { Write-Host " See desktop reports for details." -ForegroundColor Cyan }
    }

    if (-not $Silent) { Write-Host "`nPotential Waste in MB: "  -ForegroundColor Cyan -NoNewline }
    if (-not $Silent) { Write-Host "$( ($Output.PossibleWasteKB | Measure-Object -Sum).sum / 1MB)" -ForegroundColor DarkCyan -NoNewline }

}