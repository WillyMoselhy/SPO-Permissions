/*
The function app is made up of:
1- Storage Account
2- App Service Plan
3- Log Analytics Workspace
4- Application Insights (which requires a log analytics workspace)
5- Function App
7 - Keyvault to store PnP App certificate
*/

// Parameters
param StorageAccountName string
param FunctionAppName string
param KeyVaultName string
param Location string = resourceGroup().location
param AccountId string
param AZTenantDefaultDomain string
param SharePointDomain string
param PnPClientID string
param PnPApplicationName string

// Variables
var keyVaultAdministratorRoleId = '/subscriptions/${subscription().subscriptionId}/providers/Microsoft.Authorization/roleDefinitions/00482a5a-887f-4fb3-b363-3b7fe8e74483'

var functionAppSettings = [
  {
    name: 'FUNCTIONS_EXTENSION_VERSION'
    value: '~4'
  }
  {
    name: 'FUNCTIONS_WORKER_RUNTIME'
    value: 'powershell'
  }
  {
    name: 'AzureWebJobsStorage'
    value: 'DefaultEndpointsProtocol=https;AccountName=${storageAccount.name};EndpointSuffix=${environment().suffixes.storage};AccountKey=${storageAccount.listKeys().keys[0].value}'
  }
  {
    name: 'WEBSITE_CONTENTAZUREFILECONNECTIONSTRING'
    value: 'DefaultEndpointsProtocol=https;AccountName=${storageAccount.name};EndpointSuffix=${environment().suffixes.storage};AccountKey=${storageAccount.listKeys().keys[0].value}'
  }
  {
    name: 'WEBSITE_CONTENTSHARE'
    value: toLower(FunctionAppName)
  }
  {
    name: 'APPINSIGHTS_INSTRUMENTATIONKEY'
    value: appInsights.properties.InstrumentationKey
  }
  {
    name: '_PnPPowerShell_KeyVaultId'
    value: keyVault.id
  }
  {
    name: '_AZTenantDefaultDomain'
    value: AZTenantDefaultDomain
  }
  {
    name: '_SharePointDomain'
    value: SharePointDomain
  }
  {
    name: '_PnPClientID'
    value: PnPClientID
  }
  {
    name: '_PnPApplicationName'
    value: PnPApplicationName
  }
]

resource storageAccount 'Microsoft.Storage/storageAccounts@2021-08-01' = {
  name: StorageAccountName
  location: Location
  kind: 'StorageV2'
  sku: {
    name: 'Standard_LRS'
  }
  properties: {
    supportsHttpsTrafficOnly: true
    minimumTlsVersion: 'TLS1_2'
  }
}

resource appServicePlan 'Microsoft.Web/serverfarms@2021-03-01' = {
  name: '${FunctionAppName}-asp'
  location: Location
  sku: {
    name: 'Y1'
    tier: 'Dynamic'
  }
}

resource logAnalytics 'Microsoft.OperationalInsights/workspaces@2021-06-01' = {
  name: '${FunctionAppName}-law'
  location: Location
  properties: {
    sku: {
      name: 'PerGB2018'
    }
    retentionInDays: 30
    features: {
      enableLogAccessUsingOnlyResourcePermissions: true
    }
    publicNetworkAccessForIngestion: 'Enabled'
    publicNetworkAccessForQuery: 'Enabled'
  }
}

resource appInsights 'Microsoft.Insights/components@2020-02-02' = {
  name: FunctionAppName
  location: Location
  kind: 'web'
  properties: {
    Application_Type: 'web'
    publicNetworkAccessForIngestion: 'Enabled'
    publicNetworkAccessForQuery: 'Enabled'
    WorkspaceResourceId: logAnalytics.id
  }
}

// Create a function app with managed system identity (MSI)
resource functionApp 'Microsoft.Web/sites@2021-03-01' = {
  name: FunctionAppName
  location: Location
  kind: 'functionApp'
  identity: {
    type: 'SystemAssigned'
  }
  properties: {
    httpsOnly: true
    serverFarmId: appServicePlan.id
    siteConfig: {
      use32BitWorkerProcess: false
      powerShellVersion: '7.2'
      netFrameworkVersion: 'v6.0'
      appSettings: functionAppSettings
    }
  }
}

resource keyVault 'Microsoft.KeyVault/vaults@2021-10-01' = {
  name: KeyVaultName
  location: Location
  properties: {
    tenantId: subscription().tenantId
    enableRbacAuthorization: true
    sku: {
      family: 'A'
      name: 'standard'
    }
  }
}

resource keyVaultAdminRoleAssignment 'Microsoft.Authorization/roleAssignments@2020-10-01-preview' = {
  scope: keyVault
  name: guid(keyVault.id, AccountId, keyVaultAdministratorRoleId)
  properties: {
    roleDefinitionId: keyVaultAdministratorRoleId
    principalId: AccountId
    principalType: 'User'
  }
}

output msiID string = functionApp.identity.principalId
output keyvault string = keyVault.id
