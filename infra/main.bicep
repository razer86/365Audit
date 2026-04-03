// ── 365Audit Azure Function Infrastructure ──────────────────────────────────
// Deploys: Function App (B1 plan) + Storage Account + Key Vault + Managed Identity
//
// Usage:
//   az group create -n rg-365audit -l australiaeast
//   az deployment group create -g rg-365audit -f infra/main.bicep \
//     -p functionAppName=func-365audit huduBaseUrl=https://your-hudu.com
//
// After deployment:
//   1. Store Hudu API key in Key Vault:
//      az keyvault secret set --vault-name <keyVaultName-output> \
//        --name 365Audit-HuduApiKey --value '<your-hudu-api-key>'
//   2. Upload UnattendedCustomers.psd1 to the Function App:
//      az functionapp deployment source config-zip ...  (handled by GitHub Actions)
//      Then upload the customer list via Kudu SCM or az webapp deploy.

@description('Name of the Function App (must be globally unique).')
param functionAppName string

@description('Hudu instance base URL.')
param huduBaseUrl string

@description('Azure region for all resources. Defaults to the resource group location.')
param location string = resourceGroup().location

@description('Max concurrent audit jobs. Passed as AUDIT_THROTTLE_LIMIT app setting.')
@minValue(1)
@maxValue(10)
param throttleLimit int = 3

@description('Skip publishing reports to Hudu. Set to true for dry-run testing.')
param skipPublish bool = true

@description('Hudu asset layout ID for audit credential assets.')
param huduAssetLayoutId int = 67

@description('Hudu asset layout ID for monthly report assets.')
param huduReportLayoutId int = 68

@description('Display name prefix for monthly report assets in Hudu.')
param huduReportAssetName string = 'M365 - Monthly Audit Report'

@description('Comma-separated MSP email domains for technical contact checking.')
param mspDomains string = ''

// ── Naming conventions ──────────────────────────────────────────────────────
var uniqueSuffix = uniqueString(resourceGroup().id)
var storageAccountName = 'st365audit${uniqueSuffix}'
var keyVaultName = 'kv-365a-${substring(uniqueSuffix, 0, 8)}'
var appServicePlanName = 'asp-365audit'

// ── Storage Account (required by Azure Functions runtime) ───────────────────
resource storageAccount 'Microsoft.Storage/storageAccounts@2023-05-01' = {
  name: storageAccountName
  location: location
  kind: 'StorageV2'
  sku: {
    name: 'Standard_LRS'
  }
  properties: {
    supportsHttpsTrafficOnly: true
    minimumTlsVersion: 'TLS1_2'
  }
}

// ── App Service Plan (B1 — no timeout limits, ~$20 AUD/month) ──────────────
// To save costs, you can stop this plan between monthly runs via:
//   az appservice plan update -n asp-365audit -g rg-365audit --sku FREE
//   (then scale back to B1 before the next run)
resource appServicePlan 'Microsoft.Web/serverfarms@2023-12-01' = {
  name: appServicePlanName
  location: location
  kind: 'linux'
  sku: {
    name: 'B1'
    tier: 'Basic'
  }
  properties: {
    reserved: true // Required for Linux
  }
}

// ── Function App ────────────────────────────────────────────────────────────
resource functionApp 'Microsoft.Web/sites@2023-12-01' = {
  name: functionAppName
  location: location
  kind: 'functionapp,linux'
  identity: {
    type: 'SystemAssigned'
  }
  properties: {
    serverFarmId: appServicePlan.id
    httpsOnly: true
    siteConfig: {
      linuxFxVersion: 'PowerShell|7.4'
      ftpsState: 'Disabled'
      minTlsVersion: '1.2'
      appSettings: [
        {
          name: 'AzureWebJobsStorage'
          value: 'DefaultEndpointsProtocol=https;AccountName=${storageAccount.name};EndpointSuffix=${environment().suffixes.storage};AccountKey=${storageAccount.listKeys().keys[0].value}'
        }
        {
          name: 'FUNCTIONS_EXTENSION_VERSION'
          value: '~4'
        }
        {
          name: 'FUNCTIONS_WORKER_RUNTIME'
          value: 'powershell'
        }
        {
          name: 'FUNCTIONS_WORKER_RUNTIME_VERSION'
          value: '7.4'
        }
        {
          name: 'KEY_VAULT_NAME'
          value: keyVault.name
        }
        {
          name: 'HUDU_BASE_URL'
          value: huduBaseUrl
        }
        {
          name: 'AUDIT_THROTTLE_LIMIT'
          value: string(throttleLimit)
        }
        {
          name: 'SKIP_PUBLISH'
          value: skipPublish ? 'true' : 'false'
        }
        {
          name: 'HUDU_ASSET_LAYOUT_ID'
          value: string(huduAssetLayoutId)
        }
        {
          name: 'HUDU_REPORT_LAYOUT_ID'
          value: string(huduReportLayoutId)
        }
        {
          name: 'HUDU_REPORT_ASSET_NAME'
          value: huduReportAssetName
        }
        {
          name: 'MSP_DOMAINS'
          value: mspDomains
        }
        {
          name: 'WEBSITE_RUN_FROM_PACKAGE'
          value: '1'
        }
      ]
    }
  }
}

// ── Key Vault ───────────────────────────────────────────────────────────────
resource keyVault 'Microsoft.KeyVault/vaults@2023-07-01' = {
  name: keyVaultName
  location: location
  properties: {
    tenantId: subscription().tenantId
    sku: {
      family: 'A'
      name: 'standard'
    }
    enableRbacAuthorization: true
    enableSoftDelete: true
    softDeleteRetentionInDays: 7
  }
}

// ── Key Vault role assignment (Function App → Key Vault Secrets User) ───────
// Allows the Function App's Managed Identity to read secrets.
var keyVaultSecretsUserRoleId = '4633458b-17de-408a-b874-0445c86b69e6'

resource kvRoleAssignment 'Microsoft.Authorization/roleAssignments@2022-04-01' = {
  name: guid(keyVault.id, functionApp.id, keyVaultSecretsUserRoleId)
  scope: keyVault
  properties: {
    principalId: functionApp.identity.principalId
    roleDefinitionId: subscriptionResourceId('Microsoft.Authorization/roleDefinitions', keyVaultSecretsUserRoleId)
    principalType: 'ServicePrincipal'
  }
}

// ── Outputs ─────────────────────────────────────────────────────────────────
@description('Function App name (for GitHub Actions deployment target).')
output functionAppName string = functionApp.name

@description('Function App default hostname.')
output functionAppUrl string = 'https://${functionApp.properties.defaultHostName}'

@description('Key Vault name (store secrets here).')
output keyVaultName string = keyVault.name

@description('Function App Managed Identity principal ID.')
output managedIdentityPrincipalId string = functionApp.identity.principalId

@description('Storage Account name.')
output storageAccountName string = storageAccount.name
