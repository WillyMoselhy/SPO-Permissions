#region: Create Azure Resources
# Parameters
$Location = 'EastUS'
$RGName = 'rg-SPOPermissions-01'
$FunctionAppName = 'func-SPOPermission-01'
$StorageAccountName = 'safuncspopermissions01'


# Login in to Azure using the right subscription
Connect-AzAccount -UseDeviceAuthentication
$subscription = Get-AzSubscription | Out-GridView -OutputMode Single -Title 'Select Target Subscription'
$azContext = Set-AzContext -SubscriptionObject $subscription

# We will use this to assign keyvault permissions to the account running the deployment.
$azAccount = Get-AzADUser -UserPrincipalName $azContext.Account

# Create Resource Group
$resourceGroup = New-AzResourceGroup -Name $RGName -Location $Location -Force

# Deploy function app resources
$deploymentParams = @{
    Name               = "SPOPermissions-FunctionApp-{0}utc" -f (Get-Date -AsUTC -Format yyyy-MM-dd_HH-mm-ss)
    ResourceGroupName  = $RGName
    TemplateFile       = '.\Build\Bicep\FunctionApp.bicep'
    FunctionAppName    = $FunctionAppName
    StorageAccountName = $StorageAccountName
    AccountId          = $azAccount.Id
    Verbose            = $true
}
$bicepDeployment = New-AzResourceGroupDeployment @deploymentParams
# Get the function app MSI and publish Profile

#TODD: Convert this to array at output from bicep
$msiIDs = @($bicepDeployment.Outputs.msiIDprod.Value,$bicepDeployment.Outputs.msiIDdev.Value )

# Update permissions for MSI(s) to access key vault
$msiIDs | ForEach-Object {
    $appId = (Get-AzADServicePrincipal -ObjectId $_).AppId
    if(-Not (Get-AzRoleAssignment -Scope $bicepDeployment.Outputs.keyvault.Value -RoleDefinitionName 'Key Vault Secrets User' -ObjectId $appId)){
        New-AzRoleAssignment -ApplicationId $appId -RoleDefinitionName 'Key Vault Secrets User' -Scope $bicepDeployment.Outputs.keyvault.Value
    }
    else{
        "Permission is already applied"
    }
}

#endregion

#region: Github publishing
$publishProfile = Get-AzWebAppPublishingProfile -ResourceGroupName $RGName -Name "$FunctionAppName/slots/dev"

# We are not currently automatically adding thh publish Profile to GitHub Secret Actions, so asking the user to do it.
$publishProfile | Set-Clipboard
Read-Host "Function App Publish Profile is in clipboard, please paste it as a new GitHub Secrets"
# Now we need to edit the yml file to publish the function app

#endregion

#region: Setup PnP PowerShell
Import-Module PnP.PowerShell
Register-PnPManagementShellAccess # This creates an enterprise application, delete it to undo: 31359c7f-bd7e-475c-86db-fdb8c937548e (hard coded Application ID)
$result = Register-PnPAzureADApp -ApplicationName 'PnP Rocks' -Tenant 'M365x21720695.onmicrosoft.com' -Interactive

Add-AzKeyVaultCertificate -VaultName 'func-SPOPermission-01-kv' -Name 'PnPRocks'

$certBase64 = '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'


$certificateName = 'PnPPowerShell'
$keyVaulrReaderRole =
$keyVault = Get-AzResource -ResourceId $bicepDeployment.Outputs.keyvault.Value
$cert = Import-AzKeyVaultCertificate -VaultName $keyVault.Name -Name $certificateName -FilePath 'C:\GitDevOps\SPO-Permissions\FunctionApp\Cert\PnP Rocks2.pfx'
$secret = Get-AzKeyVaultSecret -VaultName $keyVault.Name -Name $certificateName
$secret.SecretValue | ConvertFrom-SecureString -AsPlainText

Get-AzKeyVaultCertificate -VaultName $keyVault.Name -Name $certificateName

# TEST - DELETE LATER
Connect-PnPOnline -Interactive -Url https://M365x21720695.sharepoint.com
Get-PnPTenant

$result = Register-PnPAzureADApp -ApplicationName 'PnP Rocks' -Tenant 'M365x21720695.onmicrosoft.com' -Interactive -OutPath 'c:\temp\pnprocks.pfx'

$certBase64 = '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'


$Cert = New-Object security.cryptography.x509certificates.x509certificate2 -ArgumentList 'C:\GitDevOps\SPO-Permissions\FunctionApp\Cert\PnP Rocks2.pfx'

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
