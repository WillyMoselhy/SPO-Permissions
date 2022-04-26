# Parameters
$Location = 'EastUS'
$RGName = 'rg-SPOPermissions-01'
$functionAppGithubSource =


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
    Verbose = $true
}
$bicepDeployment = New-AzResourceGroupDeployment @deploymentParams
$msiID = $bicepDeployment.Outputs.msiID.Value


$publishProfile = Get-AzWebAppPublishingProfile -ResourceGroupName $RGName -Name 'func-SPOPermission-01/slots/dev'
$publishProfile | scb