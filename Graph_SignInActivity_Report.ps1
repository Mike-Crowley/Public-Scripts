#requires -modules importexcel, msal.ps

<#
    In this example, I'm using a custom app registration, looking for
    users that have a UPN that ends with subdomain.mikecrowley.us
    
    It will ultimatly build an xlsx file with the following headers:
    lastSignInDateTime lastNonInteractiveSignInDateTime skuName @odata.id userPrincipalName displayName id licenseAssignmentStates assignedLicenses signInActivity 

    Needs graph permissions:
    AuditLog.Read.All
    Organization.Read.All
    
    #to do
    #Promote licenseAssignmentStatesassignedLicenses columns, if
    desired (there is already a skuName column)

#>


#auth
$tokenParams = @{
    clientID    = '<app id>' 
    tenantID    = '<your tenant>' 
    RedirectUri = '<redirect uri>'
  
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
#ref: https://docs.microsoft.com/en-us/azure/active-directory/reports-monitoring/howto-manage-inactive-user-accounts#how-to-detect-inactive-user-accounts

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

# Ridiculously long, static list for public reference more than the actual function of the script at this point 😊
$Skus = @{
    '1d0f309f-fdf9-4b2a-9ae7-9c48b91f1426' = 'AAD_BASIC_EDU'
    '3a3976ce-de18-4a87-a78e-5e9245e252df' = 'AAD_EDU'
    '41781fb2-bc02-4b7c-bd55-b576c07bb09d' = 'AAD_PREMIUM'
    'eec0eb4f-6444-4f95-aba0-50c24d67f998' = 'AAD_PREMIUM_P2'
    '61d18b02-6889-479f-8f36-56e6e0fe5792' = 'ADALLOM_FOR_AATP'
    '932ad362-64a8-4783-9106-97849a1a30b9' = 'ADALLOM_S_DISCOVERY'
    '8c098270-9dd4-4350-9b30-ba4703f3b36b' = 'ADALLOM_S_O365'
    '2e2ddb96-6af9-4b1d-a3f0-d6ecfd22edb2' = 'ADALLOM_S_STANDALONE'
    '14ab5db5-e6c4-4b20-b4bc-13e36fd2227f' = 'ATA'
    'f20fedf3-f3c3-43c3-8267-2bfdd51c0939' = 'ATP_ENTERPRISE'
    '493ff600-6a2b-4db6-ad37-a7d4eb214516' = 'ATP_ENTERPRISE_GOV'
    '944e9726-f011-4353-b654-5f7d2663db76' = 'BI_AZURE_P_2_GOV'
    '2049e525-b859-401b-b2a0-e0a31c4b1fe4' = 'BI_AZURE_P0'
    '70d33638-9c74-4d01-bfd3-562de28bd4ba' = 'BI_AZURE_P2'
    '0bf3c642-7bb5-4ccc-884e-59d09df0266c' = 'BI_AZURE_P3'
    '32d15238-9a8c-46da-af3f-21fc5351d365' = 'BI_AZURE_P3_GOV'
    '5e62787c-c316-451f-b873-1d05acd4d12c' = 'BPOS_S_TODO_1'
    'c87f142c-d1e9-4363-8630-aaea9c4d9ae5' = 'BPOS_S_TODO_2'
    '3fb82609-8c27-4f7b-bd51-30634711ee67' = 'BPOS_S_TODO_3'
    '80f0ae31-0dfb-425c-b3fc-36f40170eb35' = 'CAREERCOACH_EDU'
    '3da2fd4c-1bee-4b61-a17f-94c31e5cab93' = 'CDS_ATTENDED_RPA'
    '4802707d-47e1-45dc-82c5-b6981f0fb38c' = 'CDS_ATTENDED_RPA_GCC'
    'c84e52ae-1906-4947-ac4d-6fb3e5bf7c2e' = 'CDS_Flow_Business_Process'
    '54b61386-c818-4634-8400-61c9e8f6acd3' = 'CDS_Flow_Business_Process_GCC'
    'bed136c6-b799-4462-824d-fc045d3a9d25' = 'CDS_O365_P1'
    '95b76021-6a53-4741-ab8b-1d1f3d66a95a' = 'CDS_O365_P2'
    'a70bbf38-cdda-470d-adb8-5804b8770f41' = 'CDS_O365_P2_GCC'
    'afa73018-811e-46e9-988f-f75d2b1b8430' = 'CDS_O365_P3'
    'bce5e5ca-c2fd-4d53-8ee2-58dfffed4c10' = 'CDS_O365_P3_GCC'
    'd7f9c9bc-0a28-4da4-b5f1-731acb27a3e4' = 'CDS_PER_APP_GCC'
    '32ad3a4e-2272-43b4-88d0-80d284258208' = 'CDS_POWERAPPS_PORTALS_LOGIN'
    '0f7b9a29-7990-44ff-9d05-a76be778f410' = 'CDS_POWERAPPS_PORTALS_LOGIN_GCC'
    '72c30473-7845-460a-9feb-b58f216e8694' = 'CDS_POWERAPPS_PORTALS_PAGEVIEW'
    '352257a9-db78-4217-a29d-8b8d4705b014' = 'CDS_POWERAPPS_PORTALS_PAGEVIEW_GCC'
    '5141c408-df3d-456a-9878-a65119b0a750' = 'CDS_UNATTENDED_RPA_GCC'
    'e4d0b25d-e440-4ee9-aac4-1d5a5db9f3ef' = 'CDS_Virtual_Agent_Base_Gov'
    '95df1203-fee7-4726-b7e1-8037a8e899eb' = 'CDS_Virtual_Agent_Usl_GCC'
    'bcc0702e-ba97-48d9-ae04-fa8689c53bba' = 'CDS_Virtual_Agent_Usl_Gov'
    '5d7a2e9a-4ee5-4f1c-bc9f-abc481bf39d8' = 'CDSAICAPACITY_PERAPP'
    '91f50f7b-2204-4803-acac-5cf5668b8b39' = 'CDSAICAPACITY_PERUSER'
    '74d93933-6f22-436e-9441-66d205435abb' = 'CDSAICAPACITY_PERUSER_NEW'
    '41fcdd7d-4733-4863-9cf4-c65b83ce2df4' = 'COMMUNICATIONS_COMPLIANCE'
    '6dc145d6-95dd-4191-b9c3-185575ee6f6b' = 'COMMUNICATIONS_DLP'
    'd9fa6af4-e046-4c89-9226-729a0786685d' = 'Content_Explorer'
    '2b815d45-56e4-4e3a-b65c-66cb9175b560' = 'ContentExplorer_Standard'
    '3efff3fe-528a-4fc5-b1ba-845802cc764f' = 'CPC_2'
    'eac6b45b-aa89-429f-a37b-c8ce00e8367e' = 'CRM_ONLINE_PORTAL_GCC'
    '483cc331-f4df-4a3b-b8ca-fe1a247569f6' = 'CRMINSTANCE_GCC'
    '62edd427-6067-4274-93c4-29afdeb30707' = 'CRMSTORAGE_GCC'
    '6d99eb83-7b5f-4947-8e99-cc12f1adb399' = 'CRMTESTINSTANCE_GCC'
    '6db1f1db-2b46-403f-be40-e39395f08dbb' = 'CUSTOMER_KEY'
    '1412cdc1-d593-4ad1-9050-40c30ad0b023' = 'D365_CSI_EMBED_CE'
    '69f07c66-bee4-4222-b051-195095efee5b' = 'D365_ProjectOperations'
    '18fa3aba-b085-4105-87d7-55617b8585e6' = 'D365_ProjectOperationsCDS'
    '46129a58-a698-46f0-aa5b-17f6586297d9' = 'DATA_INVESTIGATIONS'
    '59231cdf-b40d-4534-a93e-14d0cd31d27e' = 'DATAVERSE_FOR_POWERAUTOMATE_DESKTOP'
    '6f0e9100-ff66-41ce-96fc-3d8b7ad26887' = 'DATAVERSE_POWERAPPS_PER_APP_NEW'
    '8c7d2df8-86f0-4902-b2ed-a0458298f3b3' = 'Deskless'
    '7aae746a-3463-4737-b295-3c1a16c31438' = 'DV_PowerPages_Authenticated_User'
    '40b010bb-0b69-4654-ac5e-ba161433f4b4' = 'DYN365_CDS_O365_P1'
    '4ff01e01-1ba7-4d71-8cf8-ce96c3bbcf14' = 'DYN365_CDS_O365_P2'
    '06162da2-ebf9-4954-99a0-00fee96f95cc' = 'DYN365_CDS_O365_P2_GCC'
    '28b0fa46-c39a-4188-89e2-58e979a6b014' = 'DYN365_CDS_O365_P3'
    'a7d3fb37-b6df-4085-b509-50810d991a39' = 'DYN365_CDS_O365_P3_GCC'
    'ce361df2-f2a5-4713-953f-4050ba09aad8' = 'DYN365_CDS_P1_GOV'
    '6ea4c1ef-c259-46df-bce2-943342cd3cb2' = 'DYN365_CDS_P2'
    '37396c73-2203-48e6-8be1-d882dae53275' = 'DYN365_CDS_P2_GOV'
    '50554c47-71d9-49fd-bc54-42a2765c555c' = 'DYN365_CDS_PROJECT'
    '83837d9c-c21a-46a0-873e-d834c94015d6' = 'DYN365_CDS_PROJECT_GCC'
    '17ab22cd-a0b3-4536-910a-cb6eb12696c0' = 'DYN365_CDS_VIRAL'
    'b9f7ce72-67ff-4695-a9d9-5ff620232024' = 'DYN365_CS_CHAT_FPA_GOV'
    'ffb878a5-3184-472b-800b-65eadc63d764' = 'DYN365_CS_CHAT_GOV'
    'e304c3c3-f86c-4200-b174-1ade48805b22' = 'DYN365_CS_MESSAGING_GOV'
    '9d37aa61-3cc3-457c-8b54-e6f3853aa6b6' = 'DYN365_CS_MESSAGING_TPS_GOV'
    '79bb0a8d-e686-4e16-ac59-2b3fd0014a61' = 'DYN365_ENTERPRISE_CASE_MANAGEMENT_GOV'
    'dc6643d9-1e72-4dce-9f64-1d6eac1f1c5a' = 'DYN365_ENTERPRISE_CUSTOMER_SERVICE_GOV'
    '8c66ef8a-177f-4c0d-853c-d4f219331d09' = 'DYN365_ENTERPRISE_FIELD_SERVICE'
    'd56f3deb-50d8-465a-bedb-f079817ccac1' = 'DYN365_ENTERPRISE_P1'
    'dd89efa0-5a55-4892-ba30-82e3f8008339' = 'DYN365_SALES_PRO_GOV'
    'a9b86446-fa4e-498f-a92a-41b447e03337' = 'EducationAnalyticsP1'
    '4de31727-a228-4ec3-a5bf-8e45b5ca48cc' = 'EQUIVIO_ANALYTICS'
    'd1cbfb67-18a8-4792-b643-630b7f19aad1' = 'EQUIVIO_ANALYTICS_GOV'
    '531ee2f8-b1cb-453b-9c21-d2180d014ca5' = 'EXCEL_PREMIUM'
    '34c0d7a0-a70f-4668-9238-47f9fc208882' = 'EXCHANGE_ANALYTICS'
    '208120d1-9adb-4daf-8c22-816bd5d237e7' = 'EXCHANGE_ANALYTICS_GOV'
    'efb87545-963c-4e0d-99df-69c6916d9eb0' = 'EXCHANGE_S_ENTERPRISE'
    '8c3069c0-ccdb-44be-ab77-986203a67df2' = 'EXCHANGE_S_ENTERPRISE_GOV'
    '113feb6c-3fe4-4440-bddc-54d774bf0318' = 'EXCHANGE_S_FOUNDATION'
    '922ba911-5694-4e99-a794-73aed9bfeec8' = 'EXCHANGE_S_FOUNDATION_GOV'
    '9aaf7827-d63c-4b61-89c3-182f06f82e5c' = 'EXCHANGE_S_STANDARD'
    '7e017b61-a6e0-4bdc-861a-932846591f6e' = 'FLOW_BUSINESS_PROCESS'
    'cb83e771-a077-4a73-9201-d955585b29fa' = 'FLOW_BUSINESS_PROCESS_GCC'
    '7e6d7d78-73de-46ba-83b1-6d25117334ba' = 'FLOW_DYN_APPS'
    '2c6af4f1-e7b6-4d59-bbc8-eaa884f42d69' = 'FLOW_DYN_APPS_GOV'
    'b650d915-9886-424b-a08d-633cede56f57' = 'FLOW_DYN_P2'
    'fa200448-008c-4acb-abd4-ea106ed2199d' = 'FLOW_FOR_PROJECT'
    '16687e20-06f9-4577-9cc0-34a2704260fc' = 'FLOW_FOR_PROJECT_GOV'
    '0f9b09cb-62d1-4ff4-9129-43f4996f83f4' = 'FLOW_O365_P1'
    '76846ad7-7776-4c40-a281-a386362dd1b9' = 'FLOW_O365_P2'
    'c537f360-6a00-4ace-a7f5-9128d0ac1e4b' = 'FLOW_O365_P2_GOV'
    '07699545-9485-468e-95b6-2fca3738be01' = 'FLOW_O365_P3'
    '8055d84a-c172-42eb-b997-6c2ae4628246' = 'FLOW_O365_P3_GOV'
    '774da41c-a8b3-47c1-8322-b9c1ab68be9f' = 'FLOW_P1_GOV'
    '50e68c76-46c6-4674-81f9-75456511b170' = 'FLOW_P2_VIRAL'
    'd20bfa21-e9ae-43fc-93c2-20783f0840c3' = 'FLOW_P2_VIRAL_REAL'
    'c539fa36-a64e-479a-82e1-e40ff2aa83ee' = 'Flow_Per_APP'
    '8e2c2c3d-07f6-4da7-86a9-e78cc8c2c8b9' = 'Flow_Per_APP_GCC'
    '769b8bee-2779-4c5a-9456-6f4f8629fd41' = 'FLOW_PER_USER_GCC'
    'dc789ed8-0170-4b65-a415-eb77d5bb350a' = 'Flow_PowerApps_PerUser'
    '8e3eb3bd-bc99-4221-81b8-8b8bc882e128' = 'Flow_PowerApps_PerUser_GCC'
    'e62ffe5b-7612-441f-a72d-c11cf456d33a' = 'FLOW_SALES_PRO_GOV'
    'f9f6db16-ace6-4838-b11c-892ee75e810a' = 'FLOW_Virtual_Agent_Base_Gov'
    '0b939472-1861-45f1-ab6d-208f359c05cd' = 'Flow_Virtual_Agent_Usl_Gov'
    '24af5f65-d0f3-467b-9f78-ea798c4aeffc' = 'FORMS_GOV_E3'
    '843da3a8-d2cc-4e7a-9e90-dc46019f964c' = 'FORMS_GOV_E5'
    '159f4cd6-e380-449f-a816-af1a9ef76344' = 'FORMS_PLAN_E1'
    '2789c901-c14e-48ab-a76a-be334d9d793a' = 'FORMS_PLAN_E3'
    'e212cbc7-0961-4c40-9825-01117710dcb1' = 'FORMS_PLAN_E5'
    '97f29a83-1a20-44ff-bf48-5e4ad11f3e51' = 'Forms_Pro_CE'
    '9c439259-63b0-46cc-a258-72be4313a42d' = 'Forms_Pro_FS'
    'bb681a9b-58f5-42ee-9926-674325be8aaa' = 'Forms_Pro_Service_GCC'
    'a6520331-d7d4-4276-95f5-15c0933bc757' = 'GRAPH_CONNECTORS_SEARCH_INDEX'
    'e26c2fcc-ab91-4a61-b35c-03cdc8dddf66' = 'INFO_GOVERNANCE'
    'c4801e8a-cb58-4c35-aca6-f2dcc106f287' = 'INFORMATION_BARRIERS'
    'd587c7a3-bda9-4f99-8776-9bcf59c84f75' = 'INSIDER_RISK'
    '9d0c4ee5-e4a1-4625-ab39-d82b619b1a34' = 'INSIDER_RISK_MANAGEMENT'
    'c1ec4a95-1f05-45b3-a911-aa3fa01094f5' = 'INTUNE_A'
    'd216f254-796f-4dab-bbfa-710686e646b9' = 'INTUNE_A_GOV'
    '1689aade-3d6a-4bfc-b017-46d2672df5ad' = 'Intune_Defender'
    'da24caf9-af8e-485c-b7c8-e73336da2693' = 'INTUNE_EDU'
    '882e1d05-acd1-4ccb-8708-6ee03664b117' = 'INTUNE_O365'
    '54fc630f-5a40-48ee-8965-af0503c1386e' = 'KAIZALA_O365_P2'
    'aebd3021-9f8f-4bf8-bbe3-0ed2f4f047a1' = 'KAIZALA_O365_P3'
    '0898bdbb-73b0-471a-81e5-20f1fe4dd66e' = 'KAIZALA_STANDALONE'
    '9f431833-0334-42de-a7dc-70aa40db46db' = 'LOCKBOX_ENTERPRISE'
    '89b5d3b1-3855-49fe-b46c-87c66dbc1526' = 'LOCKBOX_ENTERPRISE_GOV'
    '2f442157-a11c-46b9-ae5b-6e39ff4e5849' = 'M365_ADVANCED_AUDITING'
    'f6de4823-28fa-440b-b886-4783fa86ddba' = 'M365_AUDIT_PLATFORM'
    '6f23d6a9-adbf-481c-8538-b4c095654487' = 'M365_LIGHTHOUSE_CUSTOMER_PLAN1'
    'd55411c9-cfff-40a9-87c7-240f14df7da5' = 'M365_LIGHTHOUSE_PARTNER_PLAN1'
    '711413d0-b36e-4cd4-93db-0a50a4ab7ea3' = 'MCO_VIRTUAL_APPT'
    '4828c8ec-dc2e-4779-b502-87ac9ce28ab7' = 'MCOEV'
    'db23fce2-a974-42ef-9002-d78dd42a0f22' = 'MCOEV_GOV'
    'f47330e9-c134-43b3-9993-e7f004506889' = 'MCOEV_VIRTUALUSER'
    '3e26ee1f-8a5f-4d52-aee2-b81ce45c8f40' = 'MCOMEETADV'
    'f544b08d-1645-4287-82de-8d91f37c02a1' = 'MCOMEETADV_GOV'
    '4ed3ff63-69d7-4fb7-b984-5aec7f605ca8' = 'MCOPSTN1'
    '3c8a8792-7866-409b-bb61-1b20ace0368b' = 'MCOPSTN1_GOV'
    '5a10155d-f5c1-411a-a8ec-e99aae125390' = 'MCOPSTN2'
    '54a152dc-90de-4996-93d2-bc47e670fc06' = 'MCOPSTN5'
    '505e180f-f7e0-4b65-91d4-00d670bbd18c' = 'MCOPSTNC'
    '0feaeb32-d00e-4d66-bd5a-43b5b83db82c' = 'MCOSTANDARD'
    'a31ef4a2-f787-435e-8335-e47eb0cafc94' = 'MCOSTANDARD_GOV'
    '292cc034-7b7c-4950-aaf5-943befd3f1d4' = 'MDE_LITE'
    '8a256a2b-b617-496d-b51b-e76466e88db0' = 'MFA_PREMIUM'
    'acffdce6-c30f-4dc2-81c0-372e33c515ec' = 'Microsoft Stream'
    'a413a9ff-720c-4822-98ef-2f37c2a21f4c' = 'MICROSOFT_COMMUNICATION_COMPLIANCE'
    '85704d55-2e73-47ee-93b4-4b8ea14db92b' = 'MICROSOFT_ECDN'
    '94065c59-bc8e-4e8b-89e5-5138d471eaff' = 'MICROSOFT_SEARCH'
    '199a5c09-e0ca-4e37-8f7c-b05d533e1ea2' = 'MICROSOFTBOOKINGS'
    '64bfac92-2b17-4482-b5e5-a0304429de3e' = 'MICROSOFTENDPOINTDLP'
    '4c246bbc-f513-4311-beff-eba54c353256' = 'MINECRAFT_EDUCATION_EDITION'
    '5136a095-5cf0-4aff-bec3-e84448b38ea5' = 'MIP_S_CLP1'
    'efb0351d-3b08-4503-993d-383af8de41e3' = 'MIP_S_CLP2'
    'cd31b152-6326-4d1b-ae1b-997b625182e6' = 'MIP_S_Exchange'
    'd2d51368-76c9-4317-ada2-a12c004c432f' = 'ML_CLASSIFICATION'
    'bf28f719-7844-4079-9c78-c1307898e192' = 'MTP'
    '33c4f319-9bdd-48d6-9c4d-410b750a4a5a' = 'MYANALYTICS_P2'
    '6e5b7995-bd4f-4cbd-9d19-0e32010c72f0' = 'MYANALYTICS_P2_GOV'
    '03acaee3-9492-4f40-aed4-bcb6b32981b6' = 'NBENTERPRISE'
    '7dbc2d88-20e2-4eb6-b065-4510b38d6eb2' = 'NONPROFIT_PORTAL'
    'db4d623d-b514-490b-b7ef-8885eee514de' = 'Nucleus'
    '9b5de886-f035-4ff2-b3d8-c9127bea3620' = 'OFFICE_FORMS_PLAN_2'
    '96c1e14a-ef43-418d-b115-9636cdaa8eed' = 'OFFICE_FORMS_PLAN_3'
    'c63d4d19-e8cb-460e-b37c-4d6c34603745' = 'OFFICEMOBILE_SUBSCRIPTION'
    '43de0ff5-c92c-492b-9116-175376d08c38' = 'OFFICESUBSCRIPTION'
    'de9234ff-6483-44d9-b15e-dca72fdd27af' = 'OFFICESUBSCRIPTION_GOV'
    'da792a53-cbc0-4184-a10d-e544dd34b3c1' = 'ONEDRIVE_BASIC'
    '98709c2e-96b5-4244-95f5-a0ebe139fb8a' = 'ONEDRIVE_BASIC_GOV'
    'b1188c4c-1b36-4018-b48b-ee07604f6feb' = 'PAM_ENTERPRISE'
    '9da49a6d-707a-48a1-b44a-53dcde5267f8' = 'PBI_PREMIUM_P1_ADDON'
    '30df3dbd-5bf6-4d74-9417-cccc096595e4' = 'PBI_PREMIUM_P1_ADDON_GCC'
    '375cd0ad-c407-49fd-866a-0bff4f8a9a4d' = 'POWER_AUTOMATE_ATTENDED_RPA'
    'fb613c67-1a58-4645-a8df-21e95a37d433' = 'POWER_AUTOMATE_ATTENDED_RPA_GCC'
    '45e63e9f-6dd9-41fd-bd41-93bfa008c537' = 'POWER_AUTOMATE_UNATTENDED_RPA_GCC'
    '60bf28f9-2b70-4522-96f7-335f5e06c941' = 'Power_Pages_Internal_User'
    '0bdd5466-65c3-470a-9fa6-f679b48286b0' = 'Power_Virtual_Agent_Usl_GCC'
    '9023fe69-f9e0-4c1e-bfde-654954469162' = 'POWER_VIRTUAL_AGENTS_D365_CS_CHAT_GOV'
    'e501d49b-1176-4816-aece-2563c0d995db' = 'POWER_VIRTUAL_AGENTS_D365_CS_MESSAGING_GOV'
    '0683001c-0492-4d59-9515-d9a6426b5813' = 'POWER_VIRTUAL_AGENTS_O365_P1'
    '041fe683-03e4-45b6-b1af-c0cdc516daee' = 'POWER_VIRTUAL_AGENTS_O365_P2'
    'ded3d325-1bdc-453e-8432-5bac26d7a014' = 'POWER_VIRTUAL_AGENTS_O365_P3'
    '874fc546-6efe-4d22-90b8-5c4e7aa59f4b' = 'POWERAPPS_DYN_APPS'
    '3089c02b-e533-4b73-96a5-01fa648c3c3c' = 'POWERAPPS_DYN_APPS_GOV'
    '0b03f40b-c404-40c3-8651-2aceb74365fa' = 'POWERAPPS_DYN_P2'
    '92f7a6f3-b89b-4bbd-8c30-809e6da5ad1c' = 'POWERAPPS_O365_P1'
    'c68f8d98-5534-41c8-bf36-22fa496fa792' = 'POWERAPPS_O365_P2'
    '0a20c815-5e81-4727-9bdc-2b5a117850c3' = 'POWERAPPS_O365_P2_GOV'
    '9c0dab89-a30c-4117-86e7-97bda240acd2' = 'POWERAPPS_O365_P3'
    '0eacfc38-458a-40d3-9eab-9671258f1a3e' = 'POWERAPPS_O365_P3_GOV'
    '5ce719f1-169f-4021-8a64-7d24dcaec15f' = 'POWERAPPS_P1_GOV'
    'd5368ca3-357e-4acb-9c21-8495fb025d1f' = 'POWERAPPS_P2_VIRAL'
    'be6e5cba-3661-424c-b79a-6d95fa1d849a' = 'POWERAPPS_PER_APP_GCC'
    '14f8dac2-0784-4daa-9cb2-6d670b088d64' = 'POWERAPPS_PER_APP_NEW'
    'ea2cf03b-ac60-46ae-9c1d-eeaeb63cec86' = 'POWERAPPS_PER_USER'
    '8f55b472-f8bf-40a9-be30-e29919d4ddfe' = 'POWERAPPS_PER_USER_GCC'
    '084747ad-b095-4a57-b41f-061d84d69f6f' = 'POWERAPPS_PORTALS_LOGIN'
    'bea6aef1-f52d-4cce-ae09-bed96c4b1811' = 'POWERAPPS_PORTALS_LOGIN_GCC'
    '1c5a559a-ec06-4f76-be5b-6a315418495f' = 'POWERAPPS_PORTALS_PAGEVIEW'
    '483d5646-7724-46ac-ad71-c78b7f099d8d' = 'POWERAPPS_PORTALS_PAGEVIEW_GCC'
    '12cf31f8-754f-4efe-87a8-167c19e30831' = 'POWERAPPS_SALES_PRO_GOV'
    '2d589a15-b171-4e61-9b5f-31d15eeb2872' = 'POWERAUTOMATE_DESKTOP_FOR_WIN'
    'cdf787bd-1546-48d2-9e93-b21f9ea7067a' = 'PowerPages_Authenticated_User_GCC'
    '617b097b-4b93-4ede-83de-5f075bb5fb2f' = 'PREMIUM_ENCRYPTION'
    'fafd7243-e5c1-4a3a-9e40-495efcb1d3c3' = 'PROJECT_CLIENT_SUBSCRIPTION'
    '45c6831b-ad74-4c7f-bd03-7c2b3fa39067' = 'PROJECT_CLIENT_SUBSCRIPTION_GOV'
    '1259157c-8581-4875-bca7-2ffb18c51bda' = 'PROJECT_ESSENTIALS'
    'fdcb7064-f45c-46fa-b056-7e0e9fdf4bf3' = 'PROJECT_ESSENTIALS_GOV'
    '0a05d977-a21a-45b2-91ce-61c240dbafa2' = 'PROJECT_FOR_PROJECT_OPERATIONS'
    'a55dfd10-0864-46d9-a3cd-da5991a3e0e2' = 'PROJECT_O365_P1'
    '31b4e2fc-4cd6-4e7d-9c1b-41407303bd66' = 'PROJECT_O365_P2'
    'e7d09ae4-099a-4c34-a2a2-3e166e95c44a' = 'PROJECT_O365_P2_GOV'
    'b21a6b06-1988-436e-a07b-51ec6d9f52ad' = 'PROJECT_O365_P3'
    '9b7c50ec-cd50-44f2-bf48-d72de6f90717' = 'PROJECT_O365_P3_GOV'
    '818523f5-016b-4355-9be8-ed6944946ea7' = 'PROJECT_PROFESSIONAL'
    '22572403-045f-432b-a660-af949c0a77b5' = 'PROJECT_PROFESSIONAL_FACULTY'
    '49c7bc16-7004-4df6-8cd5-4ec48b7e9ea0' = 'PROJECT_PROFESSIONAL_FOR_GOV'
    'b737dad2-2f6c-4c65-90e3-ca563267e8b9' = 'PROJECTWORKMANAGEMENT'
    '5b4ef465-7ea1-459a-9f91-033317755a51' = 'PROJECTWORKMANAGEMENT_GOV'
    '65cc641f-cccd-4643-97e0-a17e3045e541' = 'RECORDS_MANAGEMENT'
    'a4c6cf29-1168-4076-ba5c-e8fe0e62b17e' = 'REMOTE_HELP'
    '31cf2cfc-6b0d-4adc-a336-88b724ed8122' = 'RMS_S_BASIC'
    'bea4c11e-220a-4e6d-8eb8-8ea15d019f90' = 'RMS_S_ENTERPRISE'
    '6a76346d-5d6e-4051-9fe3-ed3f312b5597' = 'RMS_S_ENTERPRISE_GOV'
    '6c57d4b6-3b23-47a5-9bc9-69f17b4947b3' = 'RMS_S_PREMIUM'
    '1b66aedf-8ca1-4f73-af76-ec76c6180f98' = 'RMS_S_PREMIUM_GOV'
    '5689bec4-755d-4753-8b61-40975025187c' = 'RMS_S_PREMIUM2'
    'bf6f5520-59e3-4f82-974b-7dbbc4fd27c7' = 'SAFEDOCS'
    'c33802dd-1b50-4b9a-8bb9-f13d2cdeadac' = 'SCHOOL_DATA_SYNC_P1'
    '500b6a2a-7a50-4f40-b5f9-160e5b8c2f48' = 'SCHOOL_DATA_SYNC_P2'
    'fe71d6c3-a2ea-4499-9778-da042bf08063' = 'SHAREPOINT_PROJECT'
    '664a2fed-6c7a-468e-af35-d61740f0ec90' = 'SHAREPOINT_PROJECT_EDU'
    'e57afa78-1f19-4542-ba13-b32cd4d8f472' = 'SHAREPOINT_PROJECT_GOV'
    '5dbe027f-2339-4123-9542-606e4d348a72' = 'SHAREPOINTENTERPRISE'
    '63038b2c-28d0-45f6-bc36-33062963b498' = 'SHAREPOINTENTERPRISE_EDU'
    '153f85dd-d912-4762-af6c-d6e0fb4f6692' = 'SHAREPOINTENTERPRISE_GOV'
    'c7699d2e-19aa-44de-8edf-1736da088ca1' = 'SHAREPOINTSTANDARD'
    '0a4983bb-d3e5-4a09-95d8-b2d0127b3df5' = 'SHAREPOINTSTANDARD_EDU'
    'be5a7ed5-c598-4fcd-a061-5e6724c68a58' = 'SHAREPOINTSTORAGE'
    'e95bec33-7c88-4a70-8e19-b10bd9d0c014' = 'SHAREPOINTWAC'
    'e03c7e47-402c-463c-ab25-949079bedb21' = 'SHAREPOINTWAC_EDU'
    '8f9f0f3b-ca90-406c-a842-95579171f8ec' = 'SHAREPOINTWAC_GOV'
    '743dd19e-1ce3-4c62-a3ad-49ba8f63a2f6' = 'STREAM_O365_E1'
    '9e700747-8b1d-45e5-ab8d-ef187ceec156' = 'STREAM_O365_E3'
    '2c1ada27-dbaa-46f9-bda6-ecb94445f758' = 'STREAM_O365_E3_GOV'
    '6c6042f5-6f01-4d67-b8c1-eb99d36eed3e' = 'STREAM_O365_E5'
    '92c2089d-9a53-49fe-b1a6-9e6bdf959547' = 'STREAM_O365_E5_GOV'
    'a23b959c-7ce8-4e57-9140-b90eb88a9e97' = 'SWAY'
    '304767db-7d23-49e8-a945-4a7eb65f9f28' = 'TEAMS_GOV'
    '92c6b761-01de-457a-9dd9-793a975238f7' = 'Teams_Room_Standard'
    '57ff2da0-773e-42df-b2af-ffb7a2317929' = 'TEAMS1'
    'cc8c0802-a325-43df-8cba-995d0c6cb373' = 'TEAMSPRO_CUST'
    '0504111f-feb8-4a3c-992a-70280f9a2869' = 'TEAMSPRO_MGMT'
    'f8b44f54-18bb-46a3-9658-44ab58712968' = 'TEAMSPRO_PROTECTION'
    '9104f592-f2a7-4f77-904c-ca5a5715883f' = 'TEAMSPRO_VIRTUALAPPT'
    '78b58230-ec7e-4309-913c-93a45cc4735b' = 'TEAMSPRO_WEBINAR'
    '8e0c0a52-6a6c-4d40-8370-dd62790dcd70' = 'THREAT_INTELLIGENCE'
    '900018f1-0cdb-4ecb-94d4-90281760fdc6' = 'THREAT_INTELLIGENCE_GOV'
    '795f6fe0-cc4d-4773-b050-5dde4dc704c9' = 'UNIVERSAL_PRINT_01'
    'b67adbaf-a096-42c9-967e-5a84edbe0086' = 'UNIVERSAL_PRINT_NO_SEEDING'
    'e425b9f6-1543-45a0-8efb-f8fdaf18cba1' = 'Virtual_Agent_Base_GCC'
    '00b6f978-853b-4041-9de0-a233d18669aa' = 'Virtual_Agent_Usl_Gov'
    'e7c91390-7625-45be-94e0-e16907e03118' = 'Virtualization Rights for Windows 10 (E3/E5+VDA)'
    '663a804f-1c30-4ff0-9915-9db84f0d1cea' = 'VISIO_CLIENT_SUBSCRIPTION'
    'f85945f4-7a55-4009-bc39-6a5f14a8eac1' = 'VISIO_CLIENT_SUBSCRIPTION_GOV'
    '2bdbaf8f-738f-4ac7-9234-3c3ee2ce7d0f' = 'VISIOONLINE'
    '8a9ecb07-cfc0-48ab-866c-f83c4d911576' = 'VISIOONLINE_GOV'
    'b76fb638-6ba6-402a-b9f9-83d28acb3d86' = 'VIVA_LEARNING_SEEDED'
    'a82fbf69-b4d7-49f4-83a6-915b2cf354f4' = 'VIVAENGAGE_CORE'
    'b8afc642-032e-4de5-8c0a-507a7bba7e5d' = 'WHITEBOARD_PLAN1'
    '94a54592-cd8b-425e-87c6-97868b000b91' = 'WHITEBOARD_PLAN2'
    '4a51bca5-1eff-43f5-878c-177680f191af' = 'WHITEBOARD_PLAN3'
    '21b439ba-a0ca-424f-a6cc-52f954a5b111' = 'WIN10_PRO_ENT_SUB'
    '871d91ec-ec1a-452b-a83f-bd76c7d770ef' = 'WINDEFATP'
    '9a6eeb79-0b4b-4bf0-9808-39d99a2cd5a3' = 'Windows_Autopatch'
    'a420f25f-a7b3-4ff5-a9d0-5d58f73b537d' = 'WINDOWS_STORE'
    '7bf960f6-2cd9-443a-8046-5dbff9558365' = 'WINDOWSUPDATEFORBUSINESS_DEPLOYMENTSERVICE'
    '2078e8df-cff6-4290-98cb-5408261a760a' = 'YAMMER_EDU'
    '7547a3fe-08ee-4ccb-b430-5077c5041653' = 'YAMMER_ENTERPRISE'
}

<# learn the above via:
    Connect-AzureAD
    $Skus = [ordered]@{}
    foreach ($plan in (Get-AzureADSubscribedSku).ServicePlans | sort ServicePlanName -Unique) {
        $Skus.Add($plan.ServicePlanId, $plan.ServicePlanName)
    }
    $Skus | ft -AutoSize
#>

#Format Output
$SelectFilter = @(
    @{n = "lastSignInDateTime"; e = { [datetime]($_.signInActivity.lastSignInDateTime) } }    
    @{n = "lastNonInteractiveSignInDateTime"; e = { [datetime]($_.signInActivity.lastNonInteractiveSignInDateTime) } }    
    @{n = 'skuName'; e = { ( $_.assignedlicenses.skuid | ForEach-Object { $SKUs.get_item($_) }) -join ';' } } 
    '*'        
)
$finalOutput = $queryResults | Select-Object $SelectFilter | Sort-Object lastSignInDateTime -Descending


#Write to file
$ReportDate = Get-Date -format ddMMMyyyy_HHmm
$DesktopPath = ([Environment]::GetFolderPath("Desktop") + '\Graph_Reporting\Graph_Reporting_' + $ReportDate + '\')
mkdir $DesktopPath -Force

$Common_ExportExcelParams = @{
    #PassThru     = $true
    BoldTopRow   = $true
    AutoSize     = $true
    AutoFilter   = $true
    FreezeTopRow = $true
}

$finalOutput | Export-Excel @Common_ExportExcelParams -Path ($DesktopPath + $ReportDate + " _signInActivity.xlsx")
