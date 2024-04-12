Function Get-AutoDiscoverV2 {
    param (
        [parameter(Mandatory = $true)][string]
        $Upn
    )    
    $Response = Invoke-WebRequest "https://outlook.office365.com/autodiscover/autodiscover.json/v1.0/$($upn)?Protocol=activesync" #or change to ews
    $Response.Content | ConvertFrom-Json
}

Get-AutoDiscoverV2 -Upn user1@mikecrowley.us # must be a valid smtp address