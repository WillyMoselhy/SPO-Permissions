#region: Create Azure Resources
# Parameters
$Location = 'SouthAfricaNorth'
$RGName = 'RGP-ECLOSIA-SPOREP-PROD-01'
$FunctionAppName = 'func-ecl-SPOPerm-01'
$StorageAccountName = 'saeclfuncspoperm0122'
$KeyVaultName = 'funcecpSPOPerm01kv22516'
$PnPApplicationName = "func-ecl-SPOPerm-01-PnPApp"

# Validate inputs
if ($KeyVaultName.Length -gt 24) { Throw "Keyvault name too long" }
if ($StorageAccountName.Length -gt 24) { Throw "Keyvault name too long" }

# Login in to Azure using the right subscription
Connect-AzAccount -UseDeviceAuthentication
$subscription = Get-AzSubscription | Out-GridView -OutputMode Single -Title 'Select Target Subscription'
$azContext = Set-AzContext -SubscriptionObject $subscription

# Get tenant information
$tenantId = $azContext.Tenant.Id
$azTenant = Get-AzTenant -TenantId $tenantId 
$rootDomain = $azTenant.Domains | Where-Object { $_ -match "^(?<RootDomain>[a-z0-9]*)\.onmicrosoft\.com$" }
if ($rootDomain.count -eq 1) {
    $rootDomain = $matches.RootDomain
}
else {
    throw "More than one root doamin found!"
}
$defaultDomain = $azTenant.DefaultDomain
$sharePointDomain = "$rootDomain.sharepoint.com"

# We will use this to assign keyvault permissions to the account running the deployment.
$azAccount = Get-AzADUser -UserPrincipalName $azContext.Account

#region: Setup PnP PowerShell
Import-Module PnP.PowerShell

$pnpSerivcePrinicapl = Get-AzADServicePrincipal -DisplayName $PnPApplicationName

if ($null -eq $pnpSerivcePrinicapl) {
    Write-Host "Registering PnP Application"
    $pnpSerivcePrinicapl = Register-PnPAzureADApp -ApplicationName $PnPApplicationName -Tenant $defaultDomain -Interactive -ErrorAction Stop
    $certBase64 = $pnpSerivcePrinicapl.Base64Encoded
    $pnpClientID = $pnpSerivcePrinicapl.'AzureAppId/ClientId'
}
else {
    Write-Host "PnP App is already registered: $PnPApplicationName"
    $certBase64 = [system.Convert]::ToBase64String(([System.IO.File]::ReadAllBytes('.\func-ecl-SPOPerm-01-PnPApp.pfx')))
    $pnpClientID = $pnpSerivcePrinicapl.AppId
}


# Create Resource Group
$resourceGroup = New-AzResourceGroup -Name $RGName -Location $Location -Force

# Deploy function app resources
$deploymentParams = @{
    Name                  = "SPOPermissions-FunctionApp-{0}utc" -f (Get-Date -AsUTC -Format yyyy-MM-dd_HH-mm-ss)
    ResourceGroupName     = $RGName
    TemplateFile          = '.\Build\Bicep\FunctionApp.bicep'
    FunctionAppName       = $FunctionAppName
    StorageAccountName    = $StorageAccountName
    KeyVaultName          = $KeyVaultName
    AZTenantDefaultDomain = $defaultDomain
    SharePointDomain      = $sharePointDomain
    
    # For SharePoint PnP Module
    PnPApplicationName    = $PnPApplicationName
    PnPClientID           = $pnpClientID

    AccountId             = $azAccount.Id
    Verbose               = $true
}
$bicepDeployment = New-AzResourceGroupDeployment @deploymentParams
# Get the function app MSI and publish Profile


$msiID = $bicepDeployment.Outputs.msiID.Value

# Update permissions for MSI(s) to access key vault

$appId = (Get-AzADServicePrincipal -ObjectId $msiID).AppId
$keyVaultRoleName = 'Key Vault Secrets User' 
if (-Not (Get-AzRoleAssignment -Scope $bicepDeployment.Outputs.keyvaultId.Value -RoleDefinitionName $keyVaultRoleName -ObjectId $msiID)) {
    New-AzRoleAssignment -ApplicationId $appId -RoleDefinitionName $keyVaultRoleName -Scope $bicepDeployment.Outputs.keyvaultId.Value
}
else {
    "Permission is already applied"
}

# Update permissions for the storage account output blob container
$storageRoleName = "Storage Blob Data Contributor"
if (-Not (Get-AzRoleAssignment -Scope $bicepDeployment.Outputs.outputContainerId.Value -RoleDefinitionName $storageRoleName -ObjectId $msiID)) {
    Write-Host "Assigning role $storageRoleName to Function App MSI $msiID"
    New-AzRoleAssignment -Scope $bicepDeployment.Outputs.outputContainerId.Value -RoleDefinitionName $storageRoleName -ApplicationId $appId 
}
else {
    "Permission is already applied"
}


#endregion

#region: Github publishing
$publishProfile = Get-AzWebAppPublishingProfile -ResourceGroupName $RGName -Name "$FunctionAppName"

# We are not currently automatically adding thh publish Profile to GitHub Secret Actions, so asking the user to do it.
$publishProfile | Set-Clipboard
Read-Host "Function App Publish Profile is in clipboard, please paste it as a new GitHub Secrets"
# Now we need to edit the yml file to publish the function app

#endregion

#region: Import PnP Certificate to Keyvault

Import-AzKeyVaultCertificate -VaultName $KeyVaultName -Name $PnPApplicationName -CertificateString $certbase64


# TEST - DELETE LATER
Connect-PnPOnline -Interactive -Url https://M365x21720695.sharepoint.com
Get-PnPTenant


$LOCAL_TenantId = "ed559cd0-4ff1-413f-9d46-9dc213a5158f"
$LOCAL_ClientId = "2e1fee6b-7fe5-48ac-b51a-da35e149f1c5"
$LOCAL_ClientSecret = "6Es8Q~Q66_0aeT_ka6ps~pBBkDtaOuq38jjBbafO"

Connect-PnPOnline -Url "https://$($LoginInfo.TenantName).Sharepoint.com" -ClientId $LoginInfo.AppID -Tenant "$($LoginInfo.TenantName).OnMicrosoft.com" -CertificatePath $LoginInfo.CertificatePath -ErrorAction Stop
Connect-PnPOnline -Url 'https://M365x21720695.sharepoint.com' -Tenant 'M365x21720695.onmicrosoft.com' -ClientId "2e1fee6b-7fe5-48ac-b51a-da35e149f1c5" -CertificatePath 'C:\GitDevOps\SPO-Permissions\FunctionApp\Cert\PnP Rocks2.pfx'
Connect-PnPOnline -Url 'https://M365x21720695.sharepoint.com' -Tenant 'M365x21720695.onmicrosoft.com' -ClientId "2e1fee6b-7fe5-48ac-b51a-da35e149f1c5" -CertificatePath 'C:\GitDevOps\SPO-Permissions\FunctionApp\Cert\PnP Rocks2.pfx'



Connect-PnPOnline -Url 'https://M365x21720695.sharepoint.com' -Tenant 'M365x21720695.onmicrosoft.com' -ClientId "8df3c7d6-2adf-4e42-a549-2f8f665c80e5" -CertificateBase64Encoded $certBase64
Get-PnPTenant

$Body = @{
    'tenant'        = $LOCAL_TenantId
    'client_id'     = $LOCAL_ClientId
    'scope'         = 'https://m365x21720695.sharepoint.com/.default'
    'client_secret' = $LOCAL_ClientSecret
    'grant_type'    = 'client_credentials'
}

# Assemble a hashtable for splatting parameters, for readability
# The tenant id is used in the uri of the request as well as the body
$Params = @{

    'Uri'         = "https://login.microsoftonline.com/$LOCAL_TenantId/oauth2/v2.0/token"
    'Method'      = 'Post'
    'Body'        = $Body
    'ContentType' = 'application/x-www-form-urlencoded'
}

$mgToken = (Invoke-RestMethod @Params).access_token
Disconnect-PnPOnline
Connect-PnPOnline -Url 'https://M365x21720695.sharepoint.com' -AccessToken $mgToken

#region: Assign Graph API permission to use PnP PowerShell with MSI
# Reference: https://pnp.github.io/powershell/articles/azurefunctions.html#assigning-microsoft-graph-permissions-to-the-managed-identity
# The reference uses AzureAD module, we are using GraphAPI here
Import-Module Microsoft.Graph.Applications
Connect-MgGraph -Scopes Application.ReadWrite.All, Directory.Read.All, Directory.ReadWrite.All, AppRoleAssignment.ReadWrite.All

$GraphAppId = "00000003-0000-0000-c000-000000000000"
$graphSP = Get-MgServicePrincipal -Search "AppId:$GraphAppId" -ConsistencyLevel eventual
$msiSP = Get-MgServicePrincipal -ServicePrincipalId '082d0922-03a4-4e55-b0ee-089e5dd3a6d0' $msiIDdev # This is obtained while deploying the function RG, if not get it from Azure Portal. Note: there is one for dev and one for prod
$msGraphPermissions = @(
    'Directory.Read.All' #Used to read user and group permissions
)
$msGraphAppRoles = $graphSP.AppRoles | Where-Object { $_.Value -in $msGraphPermissions }
Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $msiSP.Id # This is check what permissions are currently assigned
$msGraphAppRoles | ForEach-Object {
    $params = @{
        PrincipalId = $msiSP.Id
        ResourceId  = $graphSP.Id
        AppRoleId   = $_.Id
    }
    New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $msiSP.Id -BodyParameter $params
}
#endregion
