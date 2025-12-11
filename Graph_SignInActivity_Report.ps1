#requires -modules importexcel, msal.ps

<#
    In this example, I query Microsoft Graph with Invoke-RestMethod, using a custom app registration,
    looking for users that have a UPN that ends with subdomain.mikecrowley.us
    
    It will ultimatly build an xlsx file with the following headers:    
        displayName, userPrincipalName, lastSignInDateTime, lastNonInteractiveSignInDateTime, skuPartNumbers, licenseAssignmentStates, userId

    Needs graph permissions:
    AuditLog.Read.All
    Organization.Read.All
#>

#auth
$tokenParams = @{
    clientID    = <guid>
    tenantID    = <guid>
    RedirectUri = <uri>
}
$myToken = Get-MsalToken @tokenParams 

#build request headers
$uri = @'
https://graph.microsoft.com/beta/users?$count=true&$filter=endswith(userPrincipalName,'subdomain.mikecrowley.us')&$select=userPrincipalName,displayName,signInActivity,licenseAssignmentStates,assignedLicenses
'@

$requestParams = @{
    Headers = @{
        Authorization    = "Bearer $($myToken.AccessToken)"
        ConsistencyLevel = "eventual"
    } 
    Method  = "Get"    
}

#collect users displayName and signInActivity
#ref: https://learn.microsoft.com/en-us/azure/active-directory/reports-monitoring/howto-manage-inactive-user-accounts#how-to-detect-inactive-user-accounts

$queryResults = @()

do {
    if ((get-date).AddMinutes(5) -lt $myToken.ExpiresOn.LocalDateTime) {       
        $pageResults = Invoke-RestMethod -Method $requestParams.Method -Headers $requestParams.Headers -Uri $uri
        if ($null -eq $PageResults.value) {
            $QueryResults += $PageResults.value
        }
        else { $QueryResults += $PageResults }
        $uri = $PageResults.'@odata.nextlink'
    }
    else {
        Write-Output "Please wait - renewing token..."
        $myToken = Get-MsalToken @tokenParams
    }   

    Write-Output ("Users downloaded: " + $queryResults.Count)
}
until ($null -eq $uri)



# Static list of AAD SKUs
# $SubscribedSkus = Import-Csv .\SupportingFiles\Graph_SignInActivity_Report_AAD_SKUs.csv
# Most are also available here: https://learn.microsoft.com/en-us/azure/active-directory/enterprise-users/licensing-service-plan-reference
# And as a CSV here: https://download.microsoft.com/download/e/3/e/e3e9faf2-f28b-490a-9ada-c6089a1fc5b0/Product%20names%20and%20service%20plan%20identifiers%20for%20licensing.csv

#Get the list from the tenant
$SubscribedSkus = (Invoke-RestMethod -Method $requestParams.Method -Headers $requestParams.Headers -Uri "https://graph.microsoft.com/v1.0/subscribedSkus").value

$Skus = [ordered]@{}
foreach ($SKU in $SubscribedSkus) {
    $Skus.Add($SKU.SkuId, $SKU.SkuPartNumber)
}

#Format Output
$FinalOutput = foreach ($User in $queryResults.value) {
    [PSCustomObject]@{
        displayName                      = $user.displayName
        userPrincipalName                = $user.userPrincipalName
        lastSignInDateTime               = $user.signInActivity.lastSignInDateTime
        lastNonInteractiveSignInDateTime = $user.signInActivity.lastNonInteractiveSignInDateTime
        skuPartNumbers                   = If ($user.licenseAssignmentStates -ne "") { ( $user.assignedlicenses.skuid | ForEach-Object { $SKUs.get_item($_) }) -join ';' }        
        licenseAssignmentStates          = $user.licenseAssignmentStates.state -join ';'
        userId                           = $user.id        
    }
}  
#Write to file
$ReportDate = Get-Date -Format 'ddMMMyyyy_HHmm'
$DesktopPath = ([Environment]::GetFolderPath("Desktop") + '\Graph_Reporting\Graph_Reporting_' + $ReportDate + '\')
mkdir $DesktopPath -Force

$Common_ExportExcelParams = @{
    BoldTopRow   = $true
    AutoSize     = $true
    AutoFilter   = $true
    FreezeTopRow = $true
}

$FinalOutput | Sort-Object lastSignInDateTime -Descending | Export-Excel @Common_ExportExcelParams -Path ($DesktopPath + $ReportDate + " _signInActivity.xlsx")