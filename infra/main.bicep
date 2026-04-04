// ── 365Audit Azure Automation Infrastructure ──────────────────────────────
// Deploys an Automation Account with a system-assigned managed identity,
// a Key Vault for the Hudu API key, and a Log Analytics workspace.
//
// Usage:
//   az deployment group create \
//     --resource-group rg-365audit \
//     --template-file infra/main.bicep \
//     --parameters huduBaseUrl='https://your-hudu.huducloud.com'

param location string = resourceGroup().location

@description('Hudu instance base URL (no trailing slash)')
param huduBaseUrl string

@description('Comma-separated MSP email domains')
param mspDomains string = ''

@description('Hudu asset layout ID for credential assets')
param huduAssetLayoutId int = 67

@description('Hudu asset layout ID for monthly report assets')
param huduReportLayoutId int = 68

@description('Display name prefix for monthly report assets')
param huduReportAssetName string = 'M365 - Monthly Audit Report'

@description('Maximum concurrent customer audits')
@minValue(1)
@maxValue(10)
param throttleLimit int = 3

@description('Skip Hudu report publishing (dry-run mode)')
param skipPublish bool = true

@description('Name of existing Key Vault (must contain 365Audit-HuduApiKey secret)')
param keyVaultName string

// ── Automation Account ──────────────────────────────────────────────────────
resource automationAccount 'Microsoft.Automation/automationAccounts@2023-11-01' = {
  name: 'aa-365audit'
  location: location
  identity: {
    type: 'SystemAssigned'
  }
  properties: {
    sku: {
      name: 'Free'
    }
    publicNetworkAccess: true
  }
}

// ── Log Analytics Workspace ─────────────────────────────────────────────────
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

// ── Link Automation Account to Log Analytics ────────────────────────────────
resource automationDiagnostics 'Microsoft.Insights/diagnosticSettings@2021-05-01-preview' = {
  name: 'aa-365audit-diag'
  scope: automationAccount
  properties: {
    workspaceId: logAnalytics.id
    logs: [
      {
        categoryGroup: 'allLogs'
        enabled: true
      }
    ]
  }
}

// ── Key Vault access for Automation Account managed identity ────────────────
resource keyVault 'Microsoft.KeyVault/vaults@2023-07-01' existing = {
  name: keyVaultName
}

resource kvRoleAssignment 'Microsoft.Authorization/roleAssignments@2022-04-01' = {
  name: guid(keyVault.id, automationAccount.id, '4633458b-17de-408a-b874-0445c86b69e6')
  scope: keyVault
  properties: {
    roleDefinitionId: subscriptionResourceId('Microsoft.Authorization/roleDefinitions', '4633458b-17de-408a-b874-0445c86b69e6') // Key Vault Secrets User
    principalId: automationAccount.identity.principalId
    principalType: 'ServicePrincipal'
  }
}

// ── Schedule: 2:00 AM UTC on 1st of each month ─────────────────────────────
resource monthlySchedule 'Microsoft.Automation/automationAccounts/schedules@2023-11-01' = {
  parent: automationAccount
  name: 'monthly-audit'
  properties: {
    frequency: 'Month'
    interval: 1
    startTime: '2026-05-01T02:00:00+00:00'
    timeZone: 'UTC'
    advancedSchedule: {
      monthDays: [1]
    }
  }
}

// ── Outputs ─────────────────────────────────────────────────────────────────
output automationAccountName string = automationAccount.name
output automationAccountPrincipalId string = automationAccount.identity.principalId
output keyVaultName string = keyVault.name
output logAnalyticsWorkspaceId string = logAnalytics.id
