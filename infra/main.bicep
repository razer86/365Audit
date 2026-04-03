// ── 365Audit Container Apps Job Infrastructure ──────────────────────────────
// Deploys: Container Registry + Container Apps Job + Managed Identity + Log Analytics
// Reuses existing Key Vault for secret storage.
//
// Usage:
//   az deployment group create -g rg-365audit -f infra/main.bicep \
//     -p keyVaultName=kv-365a-t36i7raz huduBaseUrl=https://your-hudu.com \
//       mspDomains=domain1.com,domain2.com
//
// After deployment:
//   1. Ensure Hudu API key exists in Key Vault:
//      az keyvault secret set --vault-name <keyVaultName> \
//        --name 365Audit-HuduApiKey --value '<your-hudu-api-key>'
//   2. Grant AcrPush to GitHub Actions service principal:
//      az role assignment create --assignee <github-app-id> \
//        --role AcrPush --scope <acr-resource-id>
//   3. Push code to trigger GitHub Actions → builds image → updates job

// ── Parameters ─────────────────────────────────────────────────────────────

@description('Name of the existing Key Vault (already deployed, contains Hudu API key).')
param keyVaultName string

@description('Hudu instance base URL.')
param huduBaseUrl string

@description('Azure region for all resources. Defaults to the resource group location.')
param location string = resourceGroup().location

@description('Max concurrent audit jobs per container. Passed as AUDIT_THROTTLE_LIMIT.')
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

@description('Container image tag. Updated by GitHub Actions on each deployment.')
param imageTag string = 'latest'

// ── Naming conventions ──────────────────────────────────────────────────────
var uniqueSuffix = uniqueString(resourceGroup().id)
var acrName = 'acr365a${substring(uniqueSuffix, 0, 8)}'

// ── Existing Key Vault ─────────────────────────────────────────────────────
resource keyVault 'Microsoft.KeyVault/vaults@2023-07-01' existing = {
  name: keyVaultName
}

// ── User-Assigned Managed Identity ─────────────────────────────────────────
// User-assigned (not system-assigned) because Container Apps Jobs are ephemeral.
// A user-assigned identity persists independently of job executions.
resource managedIdentity 'Microsoft.ManagedIdentity/userAssignedIdentities@2023-01-31' = {
  name: 'id-365audit'
  location: location
}

// ── Azure Container Registry (Basic SKU, ~$7 AUD/month) ───────────────────
resource acr 'Microsoft.ContainerRegistry/registries@2023-07-01' = {
  name: acrName
  location: location
  sku: {
    name: 'Basic'
  }
  properties: {
    adminUserEnabled: false    // use managed identity for image pull
  }
}

// ── AcrPull role assignment (Managed Identity → ACR) ───────────────────────
var acrPullRoleId = '7f951dda-4ed3-4680-a7ca-43fe172d538d'

resource acrPullRole 'Microsoft.Authorization/roleAssignments@2022-04-01' = {
  name: guid(acr.id, managedIdentity.id, acrPullRoleId)
  scope: acr
  properties: {
    principalId: managedIdentity.properties.principalId
    roleDefinitionId: subscriptionResourceId('Microsoft.Authorization/roleDefinitions', acrPullRoleId)
    principalType: 'ServicePrincipal'
  }
}

// ── Key Vault Secrets User role assignment (Managed Identity → Key Vault) ──
var kvSecretsUserRoleId = '4633458b-17de-408a-b874-0445c86b69e6'

resource kvRoleAssignment 'Microsoft.Authorization/roleAssignments@2022-04-01' = {
  name: guid(keyVault.id, managedIdentity.id, kvSecretsUserRoleId)
  scope: keyVault
  properties: {
    principalId: managedIdentity.properties.principalId
    roleDefinitionId: subscriptionResourceId('Microsoft.Authorization/roleDefinitions', kvSecretsUserRoleId)
    principalType: 'ServicePrincipal'
  }
}

// ── Log Analytics Workspace ────────────────────────────────────────────────
resource logAnalytics 'Microsoft.OperationalInsights/workspaces@2023-09-01' = {
  name: 'log-365audit'
  location: location
  properties: {
    sku: {
      name: 'PerGB2018'
    }
    retentionInDays: 30
  }
}

// ── Container Apps Environment ─────────────────────────────────────────────
resource containerEnv 'Microsoft.App/managedEnvironments@2024-03-01' = {
  name: 'cae-365audit'
  location: location
  properties: {
    appLogsConfiguration: {
      destination: 'log-analytics'
      logAnalyticsConfiguration: {
        customerId: logAnalytics.properties.customerId
        sharedKey: logAnalytics.listKeys().primarySharedKey
      }
    }
  }
}

// ── Container Apps Job (scheduled trigger — 2am on 1st of each month) ─────
resource auditJob 'Microsoft.App/jobs@2024-03-01' = {
  name: 'job-365audit'
  location: location
  identity: {
    type: 'UserAssigned'
    userAssignedIdentities: {
      '${managedIdentity.id}': {}
    }
  }
  properties: {
    environmentId: containerEnv.id
    configuration: {
      triggerType: 'Schedule'
      scheduleTriggerConfig: {
        cronExpression: '0 2 1 * *'
        parallelism: 1
        replicaCompletionCount: 1
      }
      replicaTimeout: 14400    // 4 hours in seconds
      replicaRetryLimit: 0     // no auto-retry — errors are handled internally
      registries: [
        {
          server: acr.properties.loginServer
          identity: managedIdentity.id
        }
      ]
    }
    template: {
      containers: [
        {
          name: 'audit'
          image: '${acr.properties.loginServer}/365audit:${imageTag}'
          resources: {
            cpu: json('2.0')
            memory: '4Gi'
          }
          env: [
            { name: 'AZURE_CLIENT_ID';         value: managedIdentity.properties.clientId }
            { name: 'KEY_VAULT_NAME';           value: keyVault.name }
            { name: 'HUDU_BASE_URL';            value: huduBaseUrl }
            { name: 'AUDIT_THROTTLE_LIMIT';     value: string(throttleLimit) }
            { name: 'SKIP_PUBLISH';             value: skipPublish ? 'true' : 'false' }
            { name: 'HUDU_ASSET_LAYOUT_ID';     value: string(huduAssetLayoutId) }
            { name: 'HUDU_REPORT_LAYOUT_ID';    value: string(huduReportLayoutId) }
            { name: 'HUDU_REPORT_ASSET_NAME';   value: huduReportAssetName }
            { name: 'MSP_DOMAINS';              value: mspDomains }
          ]
        }
      ]
    }
  }
}

// ── Outputs ─────────────────────────────────────────────────────────────────

@description('Container Registry login server (for docker push).')
output acrLoginServer string = acr.properties.loginServer

@description('Container Registry name.')
output acrName string = acr.name

@description('Container Apps Job name.')
output jobName string = auditJob.name

@description('Key Vault name.')
output keyVaultName string = keyVault.name

@description('Managed Identity client ID (used for Connect-AzAccount -Identity).')
output managedIdentityClientId string = managedIdentity.properties.clientId

@description('Managed Identity principal ID (for troubleshooting RBAC).')
output managedIdentityPrincipalId string = managedIdentity.properties.principalId
