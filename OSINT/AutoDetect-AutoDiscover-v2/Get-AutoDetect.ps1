Function Get-AutoDetect {
    param (
        [parameter(Mandatory = $true)][string]
        $Upn
    )
    $Response = Invoke-WebRequest  "https://prod-autodetect.outlookmobile.com/autodetect/detect" -Headers @{"X-Email" = $upn }
    $Response.Content | ConvertFrom-Json | select -ExpandProperty protocols
}

Get-AutoDetect -Upn user1@mikecrowley.us  # must be a valid smtp address