// Parameters
param StorageAccountName string
param FunctionAppName string
param Location string = resourceGroup().location

/*
The function app is made up of four resources:
1- Storage Account
2- App Service Plan
3- Log Analytics Workspace
4- Application Insights (which requires a log analytics workspace)
5- Function App
6- Function App Dev slot
*/

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

resource logAnalytics 'Microsoft.OperationalInsights/workspaces@2021-06-01' ={
  name: '${FunctionAppName}-law'
  location: Location
  properties:{
    sku:{
      name: 'PerGB2018'
    }
    retentionInDays: 30
    features:{
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
  properties:{
    httpsOnly: true
    serverFarmId: appServicePlan.id
    siteConfig:{
      use32BitWorkerProcess: false
      powerShellVersion: '7.2'
      netFrameworkVersion: 'v6.0'
      appSettings:[
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
      ]
    }
  }
}

resource devSlot 'Microsoft.Web/sites/slots@2021-03-01' ={
  name: '${functionApp.name}/dev'
  location: Location
  kind: 'functionApp'
  properties:{
    enabled: true
  }
}

output msiID string = functionApp.identity.principalId