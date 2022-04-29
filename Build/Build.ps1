# Parameters
$Location = 'EastUS'
$RGName = 'rg-SPOPermissions-01'
$FunctionAppName = 'func-SPOPermission-01'
$StorageAccountName = 'safuncspopermissions01'


# Login in to Azure using the right subscription
Connect-AzAccount -UseDeviceAuthentication
$subscription = Get-AzSubscription | Out-GridView -OutputMode Single -Title 'Select Target Subscription'
Set-AzContext -SubscriptionObject $subscription
# Create Resource Group
$resourceGroup = New-AzResourceGroup -Name $RGName -Location $Location

# Deploy function app resources
$deploymentParams = @{
    Name               = "SPOPermissions-FunctionApp-{0}utc" -f (Get-Date -AsUTC -Format yyyy-MM-dd_HH-mm-ss)
    ResourceGroupName  = $RGName
    TemplateFile       = '.\Build\Bicep\FunctionApp.bicep'
    FunctionAppName    = $FunctionAppName
    StorageAccountName = $StorageAccountName
    Verbose            = $true
}
$bicepDeployment = New-AzResourceGroupDeployment @deploymentParams
# Get the function app MSI and publish Profile

$msiID = $bicepDeployment.Outputs.msiIDprod.Value
$msiIDdev = $bicepDeployment.Outputs.msiIDdev.Value
$publishProfile = Get-AzWebAppPublishingProfile -ResourceGroupName $RGName -Name "$FunctionAppName/slots/dev"

# We are not currently automatically adding teh publish Profile to GitHub Secret Actions, so asking the user to do it.
$publishProfile | scb
Read-Host "Function App Publish Profile is in clipboard, please paste it as a new GitHub Secrets"
# Now we need to edit the yml file to publish the function app

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
$msGraphAppRoles = $graphSP.AppRoles | where { $_.Value -in $msGraphPermissions }
Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $msiSP.Id # This is check what permissions are currently assigned
$msGraphAppRoles | foreach {
    $params = @{
        PrincipalId = $msiSP.Id
        ResourceId  = $graphSP.Id
        AppRoleId   = $_.Id
    }
    New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $msiSP.Id -BodyParameter $params
}
#endregion
