#requires -modules importexcel, msal.ps

#v 27Oct2021

<#
    In this example, I'm using a custom app registration, looking for
    users that have a UPN that ends with subdomain.mikecrowley.us
    
    It will ultimatly build an xlsx file with the following headers:
    lastSignInDateTime lastNonInteractiveSignInDateTime skuName @odata.id userPrincipalName displayName id licenseAssignmentStates assignedLicenses signInActivity 

    Needs graph permissions:
    AuditLogs.Read.All
    Organization.Read.All

    
    #to do
    #Promote licenseAssignmentStates	assignedLicenses columns, if
    desired (there is already a skuName column)

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
    '113feb6c-3fe4-4440-bddc-54d774bf0318' = 'EXCHANGE_S_FOUNDATION'
    '2e2ddb96-6af9-4b1d-a3f0-d6ecfd22edb2' = 'ADALLOM_S_STANDALONE'
    '14ab5db5-e6c4-4b20-b4bc-13e36fd2227f' = 'ATA'
    '61d18b02-6889-479f-8f36-56e6e0fe5792' = 'ADALLOM_FOR_AATP'
    '493ff600-6a2b-4db6-ad37-a7d4eb214516' = 'ATP_ENTERPRISE_GOV'
    'eac6b45b-aa89-429f-a37b-c8ce00e8367e' = 'CRM_ONLINE_PORTAL_GCC'
    '483cc331-f4df-4a3b-b8ca-fe1a247569f6' = 'CRMINSTANCE_GCC'
    '6d99eb83-7b5f-4947-8e99-cc12f1adb399' = 'CRMTESTINSTANCE_GCC'
    '3089c02b-e533-4b73-96a5-01fa648c3c3c' = 'POWERAPPS_DYN_APPS_GOV'
    '8f9f0f3b-ca90-406c-a842-95579171f8ec' = 'SHAREPOINTWAC_GOV'
    'fdcb7064-f45c-46fa-b056-7e0e9fdf4bf3' = 'PROJECT_ESSENTIALS_GOV'
    '153f85dd-d912-4762-af6c-d6e0fb4f6692' = 'SHAREPOINTENTERPRISE_GOV'
    '2c6af4f1-e7b6-4d59-bbc8-eaa884f42d69' = 'FLOW_DYN_APPS_GOV'
    'bb681a9b-58f5-42ee-9926-674325be8aaa' = 'Forms_Pro_Service_GCC'
    'dc6643d9-1e72-4dce-9f64-1d6eac1f1c5a' = 'DYN365_ENTERPRISE_CUSTOMER_SERVICE_GOV'
    '922ba911-5694-4e99-a794-73aed9bfeec8' = 'EXCHANGE_S_FOUNDATION_GOV'
    '74594e2b-fdad-447e-b7f9-e7eb4b9da4f7' = 'D365_FIELD_SERVICE_ATTACH_GOV'
    'e62ffe5b-7612-441f-a72d-c11cf456d33a' = 'FLOW_SALES_PRO_GOV '
    '12cf31f8-754f-4efe-87a8-167c19e30831' = 'POWERAPPS_SALES_PRO_GOV'
    'dd89efa0-5a55-4892-ba30-82e3f8008339' = 'DYN365_SALES_PRO_GOV'
    'b9f7ce72-67ff-4695-a9d9-5ff620232024' = 'DYN365_CS_CHAT_FPA_GOV'
    'ffb878a5-3184-472b-800b-65eadc63d764' = 'DYN365_CS_CHAT_GOV'
    '41781fb2-bc02-4b7c-bd55-b576c07bb09d' = 'AAD_PREMIUM'
    'bea4c11e-220a-4e6d-8eb8-8ea15d019f90' = 'RMS_S_ENTERPRISE'
    'c1ec4a95-1f05-45b3-a911-aa3fa01094f5' = 'INTUNE_A'
    '6c57d4b6-3b23-47a5-9bc9-69f17b4947b3' = 'RMS_S_PREMIUM'
    '932ad362-64a8-4783-9106-97849a1a30b9' = 'ADALLOM_S_DISCOVERY'
    '8a256a2b-b617-496d-b51b-e76466e88db0' = 'MFA_PREMIUM'
    '5689bec4-755d-4753-8b61-40975025187c' = 'RMS_S_PREMIUM2'
    'eec0eb4f-6444-4f95-aba0-50c24d67f998' = 'AAD_PREMIUM_P2'
    '0a20c815-5e81-4727-9bdc-2b5a117850c3' = 'POWERAPPS_O365_P2_GOV'
    '2c1ada27-dbaa-46f9-bda6-ecb94445f758' = 'STREAM_O365_E3_GOV'
    '24af5f65-d0f3-467b-9f78-ea798c4aeffc' = 'FORMS_GOV_E3'
    '5b4ef465-7ea1-459a-9f91-033317755a51' = 'PROJECTWORKMANAGEMENT_GOV'
    '304767db-7d23-49e8-a945-4a7eb65f9f28' = 'TEAMS_GOV'
    '882e1d05-acd1-4ccb-8708-6ee03664b117' = 'INTUNE_O365'
    '6a76346d-5d6e-4051-9fe3-ed3f312b5597' = 'RMS_S_ENTERPRISE_GOV'
    'c537f360-6a00-4ace-a7f5-9128d0ac1e4b' = 'FLOW_O365_P2_GOV'
    'de9234ff-6483-44d9-b15e-dca72fdd27af' = 'OFFICESUBSCRIPTION_GOV'
    '8c3069c0-ccdb-44be-ab77-986203a67df2' = 'EXCHANGE_S_ENTERPRISE_GOV'
    '6e5b7995-bd4f-4cbd-9d19-0e32010c72f0' = 'MYANALYTICS_P2_GOV'
    '199a5c09-e0ca-4e37-8f7c-b05d533e1ea2' = 'MICROSOFTBOOKINGS'
    '5136a095-5cf0-4aff-bec3-e84448b38ea5' = 'MIP_S_CLP1'
    'a31ef4a2-f787-435e-8335-e47eb0cafc94' = 'MCOSTANDARD_GOV'
    '89b5d3b1-3855-49fe-b46c-87c66dbc1526' = 'LOCKBOX_ENTERPRISE_GOV'
    'f544b08d-1645-4287-82de-8d91f37c02a1' = 'MCOMEETADV_GOV'
    'db23fce2-a974-42ef-9002-d78dd42a0f22' = 'MCOEV_GOV'
    'd1cbfb67-18a8-4792-b643-630b7f19aad1' = 'EQUIVIO_ANALYTICS_GOV'
    '208120d1-9adb-4daf-8c22-816bd5d237e7' = 'EXCHANGE_ANALYTICS_GOV'
    '944e9726-f011-4353-b654-5f7d2663db76' = 'BI_AZURE_P_2_GOV'
    '843da3a8-d2cc-4e7a-9e90-dc46019f964c' = 'FORMS_GOV_E5'
    '2f442157-a11c-46b9-ae5b-6e39ff4e5849' = 'M365_ADVANCED_AUDITING'
    '617b097b-4b93-4ede-83de-5f075bb5fb2f' = 'PREMIUM_ENCRYPTION'
    'efb0351d-3b08-4503-993d-383af8de41e3' = 'MIP_S_CLP2'
    '900018f1-0cdb-4ecb-94d4-90281760fdc6' = 'THREAT_INTELLIGENCE_GOV'
    '0eacfc38-458a-40d3-9eab-9671258f1a3e' = 'POWERAPPS_O365_P3_GOV'
    '8055d84a-c172-42eb-b997-6c2ae4628246' = 'FLOW_O365_P3_GOV'
    '92c2089d-9a53-49fe-b1a6-9e6bdf959547' = 'STREAM_O365_E5_GOV'
    '3c8a8792-7866-409b-bb61-1b20ace0368b' = 'MCOPSTN1_GOV'
    'ce361df2-f2a5-4713-953f-4050ba09aad8' = 'DYN365_CDS_P1_GOV'
    '774da41c-a8b3-47c1-8322-b9c1ab68be9f' = 'FLOW_P1_GOV'
    '5ce719f1-169f-4021-8a64-7d24dcaec15f' = 'POWERAPPS_P1_GOV'
    'be6e5cba-3661-424c-b79a-6d95fa1d849a' = 'POWERAPPS_PER_APP_GCC'
    '8e2c2c3d-07f6-4da7-86a9-e78cc8c2c8b9' = 'Flow_Per_APP_GCC'
    'd7f9c9bc-0a28-4da4-b5f1-731acb27a3e4' = 'CDS_PER_APP_GCC'
    '8f55b472-f8bf-40a9-be30-e29919d4ddfe' = 'POWERAPPS_PER_USER_GCC'
    '37396c73-2203-48e6-8be1-d882dae53275' = 'DYN365_CDS_P2_GOV'
    '8e3eb3bd-bc99-4221-81b8-8b8bc882e128' = 'Flow_PowerApps_PerUser_GCC'
    '45c6831b-ad74-4c7f-bd03-7c2b3fa39067' = 'PROJECT_CLIENT_SUBSCRIPTION_GOV'
    'e57afa78-1f19-4542-ba13-b32cd4d8f472' = 'SHAREPOINT_PROJECT_GOV'
    '31cf2cfc-6b0d-4adc-a336-88b724ed8122' = 'RMS_S_BASIC'
    '4ccb60ee-9523-48fd-8f63-4b090f1ad77a' = 'OFFICEMOBILE_SUBSCRIPTION_GOV'
    'c42aa49a-f357-45d5-9972-bc29df885fee' = 'POWERAPPS_O365_P1_GOV'
    'ad6c8870-6356-474c-901c-64d7da8cea48' = 'FLOW_O365_P1_GOV'
    '15267263-5986-449d-ac5c-124f3b49b2d6' = 'STREAM_O365_E1_GOV'
    'f4cba850-4f34-4fd2-a341-0fddfdce1e8f' = 'FORMS_GOV_E1'
    'f9c43823-deb4-46a8-aa65-8b551f0c4f8a' = 'SHAREPOINTSTANDARD_GOV'
    'e9b4930a-925f-45e2-ac2a-3f7788ca6fdd' = 'EXCHANGE_S_STANDARD_GOV'
    'f85945f4-7a55-4009-bc39-6a5f14a8eac1' = 'VISIO_CLIENT_SUBSCRIPTION_GOV'
    '98709c2e-96b5-4244-95f5-a0ebe139fb8a' = 'ONEDRIVE_BASIC_GOV'
    '8a9ecb07-cfc0-48ab-866c-f83c4d911576' = 'VISIOONLINE_GOV'
    'a420f25f-a7b3-4ff5-a9d0-5d58f73b537d' = 'WINDOWS_STORE'
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
    @{n="lastNonInteractiveSignInDateTime";  e={[datetime]($_.signInActivity.lastNonInteractiveSignInDateTime)}}    
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

$finalOutput | Export-Excel @Common_ExportExcelParams -Path ($DesktopPath + $ReportDate + " _signInActivity.xlsx")
