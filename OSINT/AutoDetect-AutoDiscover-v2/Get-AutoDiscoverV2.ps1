Function Get-AutoDiscoverV2 {
    param (
        [parameter(Mandatory = $true)][string]
        $Upn
    )

    try {
        $Response = Invoke-WebRequest -Uri "https://outlook.office365.com/autodiscover/autodiscover.json/v1.0/$($Upn)?Protocol=activesync" -UseBasicParsing -ErrorAction Stop  # or change Protocol to ews
    }
    catch {
        Write-Warning "Failed to retrieve autodiscover information for $Upn : $($_.Exception.Message)"
        return
    }

    $Response.Content | ConvertFrom-Json
}

Get-AutoDiscoverV2 -Upn user1@mikecrowley.us # must be a valid smtp address