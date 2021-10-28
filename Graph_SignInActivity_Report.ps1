#v 27Aug2021 v4

<#
    In this example, I'm using a custom app registration
    I looking for users that have a UPN that ends with subdomain.mikecrowley.us
    
    It will ultimatly build an xlsx file with the following headers:
    lastSignInDateTime lastNonInteractiveSignInDateTime skuName @odata.id userPrincipalName displayName id licenseAssignmentStates assignedLicenses signInActivity 
    
    #to do
    #fix - license objects are flattened in excel file

#>


#auth
$tokenParams = @{
    clientID     = '<app id>' 
    tenantID     = '<your tenant>' 
    RedirectUri = '<redirect uri>'
  
}
$myToken = Get-MsalToken @tokenParams


#build request headers

$uri = @'
https://graph.microsoft.com/beta/users?$count=true&$filter=endswith(userPrincipalName,'subdomain.mikecrowley.us')&$select=userPrincipalName,displayName,signInActivity,licenseAssignmentStates,assignedLicenses
'@

$requestParams = @{
    Headers  = @{
        Authorization = "Bearer $($myToken.AccessToken)"
        ConsistencyLevel = "eventual"
        } 
    Method   = "Get"    
}

#collect users displayName and signInActivity
#ref: https://docs.microsoft.com/en-us/azure/active-directory/reports-monitoring/howto-manage-inactive-user-accounts#how-to-detect-inactive-user-accounts

$queryResults = @()

do {
if ((get-date).AddMinutes(5) -lt $myToken.ExpiresOn.LocalDateTime) {       
    $pageResults = Invoke-RestMethod -Method $requestParams.Method -Headers $requestParams.Headers -Uri $uri
    if ($PageResults.value -ne $null) {
        $QueryResults += $PageResults.value
        }
        else {$QueryResults += $PageResults}
        $uri = $PageResults.'@odata.nextlink'
    }
    else {
        Write-Output "Please wait - renewing token..."
        $myToken = Get-MsalToken @tokenParams
    }   

    Write-Output ("Users downloaded: " + $queryResults.Count)
}
until ($uri -eq $null)


#Sku lookup table for later output - not 100%
$SKUs = @{                                  
   'e823ca47-49c4-46b3-b38d-ca11d5abe3d2' = 'M365_G3_GOV'        
   '4ae99959-6b0f-43b0-b1ce-68146001bdba' = 'VISIOCLIENT_GOV'
   '65758a5f-2e16-43b3-a8cb-296cd8f69e09' = 'D365_ENTERPRISE_CUSTOMER_SERVICE_GOV'
   'f0da4861-bc71-4962-942b-9b6a79d15f41' = 'D365_FIELD_SERVICE_ATTACH_GOV'
   '229fa362-9d30-4dbc-8110-21b77a7f9b26' = 'D365_SALES_PRO_GOV'
   '1b399f66-be2a-479c-a79d-84a43a46f79e' = 'DYN365_CS_CHAT_GOV'
   'eca22b68-b31f-4e9c-a20c-4d40287bc5dd' = 'POWERAPPS_P1_GOV'
   'a460366a-ade7-4791-b581-9fbff1bdaa85' = 'MCOEV_GOV'
   'cb9bc974-a47b-4123-998d-a383390168cc' = 'CRM_ONLINE_PORTAL_GCC'
   'f2230877-72be-4fec-b1ba-7156d6f75bd6' = 'PROJECTPREMIUM_GOV'
   'f0612879-44ea-47fb-baf0-3d76d9235576' = 'POWERBI_PRO_GOV'
   '6470687e-a428-4b7a-bef2-8a291ad947c9' = 'WINDOWS_STORE'
   'c793db86-5237-494e-9b11-dcd4877c2c8c' = 'EMS_GOV'
   '8e4c6baa-f2ff-4884-9c38-93785d0d7ba1' = 'POWERAPPS_PER_USER_GCC'
   '923f58ab-fca1-46a1-92f9-89fda21238a8' = 'MCOPSTN_1_GOV'
   '8900a2c0-edba-4079-bdf3-b276e293b6a8' = 'ENTERPRISEPREMIUM_GOV'
   '2d3091c7-0712-488b-b3d8-6b97bde6a1f5' = 'MCOMEETADV_GOV'
   '1d2756cb-2147-4b05-b4d5-f013c022dcb9' = 'CRMTESTINSTANCE_GCC'
   '3f4babde-90ec-47c6-995d-d223749065d1' = 'STANDARDPACK_GOV'
   'df845ce7-05f9-4894-b5f2-11bbfbcfd2b6' = 'ADALLOM_STANDALONE'
   'efccb6f7-5641-4e0e-bd10-b4976e1bf68e' = 'EMS'
   '98defdf7-f6c1-44f5-a1f6-943b6764e7a5' = 'ATA'
   '093e8d14-a334-43d9-93e3-30589a8b47d0' = 'RMSBASIC'
   '2bd3cb20-1bb6-446b-b4d0-089af3a05c52' = 'CRMINSTANCE_GCC'
   '7be8dc28-4da4-4e6d-b9b9-c60f2806df8a' = 'EXCHANGEENTERPRISE_GOV'
   '535a3a29-c5f0-42fe-8215-d3b9e1f38c4a' = 'ENTERPRISEPACK_GOV'
   '074c6829-b3a0-430a-ba3d-aca365e57065' = 'PROJECTPROFESSIONAL_GOV'
   '8a180c2b-f4cf-4d44-897c-3d32acc4a60b' = 'EMSPREMIUM_GOV'
   'd0d1ca43-b81a-4f51-81e5-a5b1ad7bb005' = 'ATP_ENTERPRISE_GOV'
   'e4fa3838-3d01-42df-aa28-5e0a4c68604b' = 'Office 365 Education E3 for Faculty'
   '94763226-9b3c-4e75-a931-5c89701abe66' = 'Office 365 Education for Faculty'
   '78e66a63-337a-4a9a-8959-41c6654dfb56' = 'Office 365 Education for Faculty'
   'a4585165-0533-458a-97e3-c400570268c4' = 'Office 365 Education E5 for Faculty'
   '9a320620-ca3d-4705-a79d-27c135c96e05' = 'Office 365 Education E5 without PSTN Conferencing for Faculty'
   'a19037fc-48b4-4d57-b079-ce44b7832473' = 'Office 365 Education E1 for Faculty'
   'f5a9147f-b4f8-4924-a9f0-8fadaac4982f' = 'Office 365 Education E3 for Faculty'
   '16732e85-c0e3-438e-a82f-71f39cbe2acb' = 'Office 365 Education E4 for Faculty'
   '4b590615-0888-425a-a965-b3bf7789848d' = 'Microsoft 365 Education A3 for Faculty'
   'e97c048c-37a4-45fb-ab50-922fbf07a370' = 'Microsoft 365 Education A5 for Faculty'
   'e578b273-6db4-4691-bba0-8d691f4da603' = 'Microsoft 365 A5 without Audio Conferencing for Faculty'
   '43e691ad-1491-4e8c-8dc9-da6b8262c03b' = 'Office 365 Education for Home school for Faculty'
   'af4e28de-6b52-4fd3-a5f4-6bf708a304d3' = 'Office 365 A1 for Faculty (for Device)'
   '8fc2205d-4e51-4401-97f0-5c89ef1aafbb' = 'Office 365 Education E3 for Students'
   '314c4481-f395-4525-be8b-2ec4bb1e9d91' = 'Office 365 Education for Students'
   'ee656612-49fa-43e5-b67e-cb1fdf7699df' = 'Office 365 Education E5 for Students'
   '1164451b-e2e5-4c9e-8fa6-e5122d90dbdc' = 'Office 365 Education E5 without PSTN Conferencing for Students'
   'd37ba356-38c5-4c82-90da-3d714f72a382' = 'Office 365 Education E1 for Students'
   '05e8cabf-68b5-480f-a930-2143d472d959' = 'Office 365 Education E4 for Students'
   '7cfd9a2b-e110-4c39-bf20-c6a3f36a3121' = 'Microsoft 365 Education A3 for Students'
   '18250162-5d87-4436-a834-d795c15c80f3' = 'Microsoft 365 Education A3 for Students use benefits'
   '46c119d4-0379-4a9d-85e4-97c66d3f909e' = 'Microsoft 365 Education A5 for Students'
   '31d57bc7-3a05-4867-ab53-97a17835a411' = 'Microsoft 365 A5 Student use benefits'
   'a25c01ce-bab1-47e9-a6d0-ebe939b99ff9' = 'Microsoft 365 A5 without Audio Conferencing for Students'
   '81441ae1-0b31-4185-a6c0-32b6b84d419f' = 'Microsoft 365 A5 without Audio Conferencing for Students use benefit'
   '98b6e773-24d4-4c0d-a968-6e787a1f8204' = 'Office 365 A3 for Students'
   '476aad1e-7a7f-473c-9d20-35665a5cbd4f' = 'Office 365 A3 Student use benefit'
   'f6e603f1-1a6d-4d32-a730-34b809cb9731' = 'Office 365 A5 Student use benefit'
   'bc86c9cd-3058-43ba-9972-141678675ac1' = 'Office 365 A5 without Audio Conferencing for Students use benefit'
   'afbb89a7-db5f-45fb-8af0-1bc5c5015709' = 'Office 365 Education forHomeschool for Students'
   '160d609e-ab08-4fce-bc1c-ea13321942ac' = 'Office 365 A1 for Students (forDevice)'
   'e82ae690-a2d5-4d76-8d30-7c6e01e6022e' = 'Office 365 A1 Plus for Students'  
}


#Format Output
$SelectFilter = @(
    @{n="lastSignInDateTime";                e={[datetime]($_.signInActivity.lastSignInDateTime)}}
    #@{n="lastSignInRequestId";               e={$_.signInActivity.lastSignInRequestId}}
    @{n="lastNonInteractiveSignInDateTime";  e={[datetime]($_.signInActivity.lastNonInteractiveSignInDateTime)}}
    #@{n="lastNonInteractiveSignInRequestId"; e={$_.signInActivity.lastNonInteractiveSignInRequestId}}
    @{n='skuName';                           e={( $_.assignedlicenses.skuid | foreach {$SKUs.get_item($_) }) -join ';'}} 
    '*'        
)
$finalOutput = $queryResults | select $SelectFilter | sort lastSignInDateTime -Descending


#Write to file
$ReportDate = Get-Date -format ddMMMyyyy_HHmm
$DesktopPath = ([Environment]::GetFolderPath("Desktop") + '\Graph_Reporting\Graph_Reporting_'+ $ReportDate + '\')
md $DesktopPath -Force

$Common_ExportExcelParams = @{
    #PassThru     = $true
    BoldTopRow   = $true
    AutoSize     = $true
    AutoFilter   = $true
    FreezeTopRow = $true
}

# $finalOutput | export-clixml c:\tmp\dcpsfile.xml -Force
$finalOutput | Export-Excel @Common_ExportExcelParams -Path ($DesktopPath + $ReportDate + " DCPS_signInActivity.xlsx")
