# Parameters
$Location = 'EastUS'
$RGName = 'rg-SPOPermissions-01'
$FunctionAppName = 'func-SPOPermission-01'
$StorageAccountName  = 'safuncspopermissions01'


# Login in to Azure using the right subscription
Connect-AzAccount -UseDeviceAuthentication
$subscription = Get-AzSubscription | Out-GridView -OutputMode Single -Title 'Select Target Subscription'
Set-AzContext -SubscriptionObject $subscription
# Create Resource Group
$resourceGroup = New-AzResourceGroup -Name $RGName -Location $Location

# Deploy function app resources
$deploymentParams = @{
    Name = "SPOPermissions-FunctionApp-{0}utc" -f (Get-Date -AsUTC -Format yyyy-MM-dd_HH-mm-ss)
    ResourceGroupName = $RGName
    TemplateFile = '.\Build\Bicep\FunctionApp.bicep'
    FunctionAppName = $FunctionAppName
    StorageAccountName = $StorageAccountName
    Verbose = $true
}
$bicepDeployment = New-AzResourceGroupDeployment @deploymentParams
# Get the function app MSI and publish Profile

$msiID = $bicepDeployment.Outputs.msiID.Value
$publishProfile = Get-AzWebAppPublishingProfile -ResourceGroupName $RGName -Name "$FunctionAppName/slots/dev"

# We are not currently automatically adding teh publish Profile to GitHub Secret Actions, so asking the user to do it.
$publishProfile | scb
Read-Host "Function App Publish Profile is in clipboard, please paste it as a new GitHub Secrets"
# Now we need to edit the yml file to publish the function app

$publishProfile