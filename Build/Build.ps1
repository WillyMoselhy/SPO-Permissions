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

$publishProfile = Get-AzWebAppPublishingProfile -ResourceGroupName $RGName -Name "$FunctionAppName/slots/dev"

# We are not currently automatically adding thh publish Profile to GitHub Secret Actions, so asking the user to do it.
$publishProfile | Set-Clipboard
Read-Host "Function App Publish Profile is in clipboard, please paste it as a new GitHub Secrets"
# Now we need to edit the yml file to publish the function app

#region: Setup PnP PowerShell
Import-Module PnP.PowerShell
Register-PnPManagementShellAccess # This creates an enterprise application, delete it to undo: 31359c7f-bd7e-475c-86db-fdb8c937548e (hard coded Application ID)


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
$result = Register-PnPAzureADApp -ApplicationName 'PnP Rocks' -Tenant 'M365x21720695.onmicrosoft.com' -Interactive -OutPath c:\temp\

$Cert = New-Object security.cryptography.x509certificates.x509certificate2 -ArgumentList 'C:\GitDevOps\SPO-Permissions\FunctionApp\Cert\PnP Rocks2.pfx'

$LOCAL_TenantId = "ed559cd0-4ff1-413f-9d46-9dc213a5158f"
$LOCAL_ClientId = "2e1fee6b-7fe5-48ac-b51a-da35e149f1c5"
$LOCAL_ClientSecret = "6Es8Q~Q66_0aeT_ka6ps~pBBkDtaOuq38jjBbafO"

Connect-PnPOnline -Url "https://$($LoginInfo.TenantName).Sharepoint.com" -ClientId $LoginInfo.AppID -Tenant "$($LoginInfo.TenantName).OnMicrosoft.com" -CertificatePath $LoginInfo.CertificatePath -ErrorAction Stop
Connect-PnPOnline -Url 'https://M365x21720695.sharepoint.com' -Tenant 'M365x21720695.onmicrosoft.com' -ClientId "2e1fee6b-7fe5-48ac-b51a-da35e149f1c5" -CertificatePath 'C:\GitDevOps\SPO-Permissions\FunctionApp\Cert\PnP Rocks2.pfx'
Connect-PnPOnline -Url 'https://M365x21720695.sharepoint.com' -Tenant 'M365x21720695.onmicrosoft.com' -ClientId "2e1fee6b-7fe5-48ac-b51a-da35e149f1c5" -CertificatePath 'C:\GitDevOps\SPO-Permissions\FunctionApp\Cert\PnP Rocks2.pfx'

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
$msiSP = Get-MgServicePrincipal -ServicePrincipalId $msiIDdev # This is obtained while deploying the function RG, if not get it from Azure Portal. Note: there is one for dev and one for prod
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
