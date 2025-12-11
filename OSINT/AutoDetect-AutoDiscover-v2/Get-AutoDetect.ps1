Function Get-AutoDetect {
    param (
        [parameter(Mandatory = $true)][string]
        $Upn
    )

    try {
        $Response = Invoke-WebRequest -Uri "https://prod-autodetect.outlookmobile.com/autodetect/detect" -Headers @{"X-Email" = $Upn } -UseBasicParsing -ErrorAction Stop
    }
    catch {
        Write-Warning "Failed to retrieve autodetect information for $Upn : $($_.Exception.Message)"
        return
    }

    $Response.Content | ConvertFrom-Json | Select-Object -ExpandProperty protocols
}

Get-AutoDetect -Upn user1@mikecrowley.us  # must be a valid smtp address