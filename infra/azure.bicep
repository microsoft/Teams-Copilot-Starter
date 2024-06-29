@maxLength(24)
@minLength(4)
@description('Used to generate names for all resources in this file')
param botResourceBaseName string

@description('Required when create Azure Bot service')
param botAadAppClientId string

@secure()
@description('Required by Bot Framework package in your bot project')
param botAadAppClientSecret string
param botAppType string
param teamsFxEnv string
param appVersion string

param storageSKU string
param botWebAppSKU string

@maxLength(42)
param botDisplayName string

param teamsAppId string

param botServerfarmsName string = '${botResourceBaseName}plan'
param botWebAppName string = '${botResourceBaseName}web'

param location string = resourceGroup().location
param aadAppClientId string
param aadAppTenantId string
param aadAppOauthAuthorityHost string
@secure()
param aadAppClientSecret string
param openAIEndpoint string
@secure()
param openAIKey string
param openAIModel string
param openAIEmbeddingModel string
param openAIApiVersion string
param defaultPromptName string
param storageContainerName string
param maxTurns string
param botChatHistoryStorageName string = '${botResourceBaseName}sta'
param maxFileSize string
param maxPages string
param webDataSource string
param documentDataSource string
param indexFolderPath string
param azureAISearch string

var oauthAuthority = uri(aadAppOauthAuthorityHost, aadAppTenantId)
var teamsMobileOrDesktopAppClientId = '1fec8e78-bce4-4aaf-ab1b-5451cc387264'
var teamsWebAppClientId = '5e3ce6c0-2b1f-4285-8d4b-75ee78787346'
var officeWebAppClientId1 = '4345a7b9-9a63-4910-a426-35363201d503'
var officeWebAppClientId2 = '4765445b-32c6-49b0-83e6-1d93765276ca'
var outlookDesktopAppClientId = 'd3590ed6-52b3-4102-aeff-aad2292ab01c'
var outlookWebAppClientId = '00000002-0000-0ff1-ce00-000000000000'
var authorizedClientApplicationIds = '${teamsMobileOrDesktopAppClientId};${teamsWebAppClientId};${officeWebAppClientId1};${officeWebAppClientId2};${outlookDesktopAppClientId};${outlookWebAppClientId}'

///////////////////////////////////////////////////////
// BOT
///////////////////////////////////////////////////////

//Create AppInsights for the Bot
resource botAppInsights 'Microsoft.Insights/components@2020-02-02' = {
  location: location
  name: '${botResourceBaseName}ai'
  kind: 'web'
  properties: {
    Application_Type: 'web'
  }
}

// Azure Storage that hosts Tab static web site and Bot chat history
resource botStorageAccount 'Microsoft.Storage/storageAccounts@2022-09-01' = {
  kind: 'StorageV2'
  location: location
  name: botChatHistoryStorageName
  properties: {
    supportsHttpsTrafficOnly: true
  }
  sku: {
    name: storageSKU
  }
}

// Compute resources for the Bot Web App
resource botServerfarm 'Microsoft.Web/serverfarms@2021-02-01' = {
  kind: 'app'
  location: location
  name: botServerfarmsName
  sku: {
    name: botWebAppSKU
  }
}

// Web App that hosts Bot
resource botWebApp 'Microsoft.Web/sites@2021-02-01' = {
  kind: 'app'
  location: location
  name: botWebAppName
  properties: {
    serverFarmId: botServerfarm.id
    httpsOnly: true
    siteConfig: {
      alwaysOn: true
      use32BitWorkerProcess: false
      appSettings: [
        {
          name: 'WEBSITE_RUN_FROM_PACKAGE'
          value: '1' // Run Azure APP Service from a package file
        }
        {
          name: 'WEBSITE_NODE_DEFAULT_VERSION'
          value: '~18' // Set NodeJS version to 18.x for your site
        }
        {
          name: 'RUNNING_ON_AZURE'
          value: '1'
        }
        {
          name: 'BOT_ID'
          value: botAadAppClientId
        }
        {
          name: 'BOT_PASSWORD'
          value: botAadAppClientSecret
        }
      ]
      ftpsState: 'FtpsOnly'
    }
  }
}

// Configure the Bot Web App settings
resource botWebAppSettings 'Microsoft.Web/sites/config@2021-02-01' = {
  parent: botWebApp
  name: 'appsettings'
  properties: {
    WEBSITE_NODE_DEFAULT_VERSION: '~18'
    WEBSITE_RUN_FROM_PACKAGE: '1'
    TEAMSFX_ENV: teamsFxEnv
    APP_VERSION: appVersion
    TEAMS_APP_ID: teamsAppId
    BOT_ID: botAadAppClientId
    BOT_PASSWORD: botAadAppClientSecret
    BOT_DOMAIN: botWebApp.properties.defaultHostName
    BOT_APP_TYPE: botAppType
    AAD_APP_CLIENT_ID: aadAppClientId
    AAD_APP_CLIENT_SECRET: aadAppClientSecret
    AAD_APP_TENANT_ID: aadAppTenantId
    AAD_APP_OAUTH_AUTHORITY_HOST: aadAppOauthAuthorityHost
    OPENAI_KEY: openAIKey
    OPENAI_ENDPOINT: openAIEndpoint
    OPENAI_MODEL: openAIModel
    OPENAI_EMBEDDING_MODEL: openAIEmbeddingModel
    STORAGE_ACCOUNT_NAME: botStorageAccount.name
    STORAGE_ACCOUNT_KEY: botStorageAccount.listKeys().keys[0].value
    OPENAI_API_VERSION: openAIApiVersion
    VECTRA_INDEX_PATH: indexFolderPath
    DEFAULT_PROMPT_NAME: defaultPromptName
    STORAGE_CONTAINER_NAME: storageContainerName
    WEBDATA_SOURCE_NAME: webDataSource
    DOCUMENTDATA_SOURCE_NAME: documentDataSource
    AZURE_SEARCH_SOURCE_NAME: azureAISearch
    APPLICATIONINSIGHTS_INSTRUMENTATION_KEY: botAppInsights.properties.InstrumentationKey
    MAX_TURNS: maxTurns
    MAX_FILE_SIZE: maxFileSize
    MAX_PAGES: maxPages
    RUNNING_ON_AZURE: '1'
  }
}

// Register Bot web service as a bot with the Bot Framework
module azureBotRegistration './botRegistration/azurebot.bicep' = {
  name: 'Azure-Bot-registration'
  params: {
    resourceBaseName: botResourceBaseName
    botAadAppClientId: botAadAppClientId
    botAppDomain: botWebApp.properties.defaultHostName
    botDisplayName: botDisplayName
  }
}

// Create a blob service for the Bot chat history
resource botBlobService 'Microsoft.Storage/storageAccounts/blobServices@2022-09-01' = {
  name: 'default'
  parent: botStorageAccount
}

// Create a blob container to store Bot chat history
resource botStorageContainer 'Microsoft.Storage/storageAccounts/blobServices/containers@2022-09-01' = {
  parent: botBlobService
  name: 'conversations'
  properties: {
    publicAccess: 'None'
  }
}

// The output will be persisted in .env.{envName}. Visit https://aka.ms/teamsfx-actions/arm-deploy for more details.
output BOT_AZURE_APP_SERVICE_RESOURCE_ID string = botWebApp.id
output BOT_DOMAIN string = botWebApp.properties.defaultHostName
output STORAGE_ACCOUNT_NAME string = botStorageAccount.name
output STORAGE_ACCOUNT_KEY string = botStorageAccount.listKeys().keys[0].value
