# Azure Automation Deployment Guide

Deploys 365Audit as a scheduled PowerShell Runbook in Azure Automation.

---

## Architecture

```
GitHub Actions (on push to main)
  └─ Import Runbook + scripts to Automation Account

Azure Automation (scheduled: 2am on 1st of month)
  └─ Connect-AzAccount -Identity (managed identity)
  └─ Fetch Hudu API key from Key Vault
  └─ Sync customer list from Hudu
  └─ For each customer:
      └─ Fetch credentials from Hudu
      └─ Run audit modules (1-7)
      └─ Generate summary report
      └─ Publish to Hudu
```

---

## Prerequisites

| Tool | Install | Check |
|------|---------|-------|
| Azure CLI | `winget install Microsoft.AzureCLI` | `az --version` |
| GitHub CLI | `winget install GitHub.cli` | `gh --version` |

---

## Step 1 — Login and Resource Group

```bash
az login --tenant <tenantId>
az account set --subscription "<subscriptionId>"

# Create resource group (or reuse existing)
az group create --name rg-365audit --location australiaeast
```

## Step 2 — Ensure Key Vault Exists

```bash
# Only if Key Vault doesn't exist yet:
az keyvault create --name <keyVaultName> --resource-group rg-365audit \
  --location australiaeast --enable-rbac-authorization

az keyvault secret set --vault-name <keyVaultName> \
  --name 365Audit-HuduApiKey --value '<your-hudu-api-key>'
```

## Step 3 — Deploy Infrastructure

```bash
az deployment group create \
  --resource-group rg-365audit \
  --template-file infra/main.bicep \
  --parameters \
    keyVaultName='<keyVaultName>' \
    huduBaseUrl='https://your-hudu.huducloud.com'
```

This creates:
- **Azure Automation Account** (Free tier — 500 min/month)
- **System-assigned Managed Identity** with Key Vault Secrets User role
- **Log Analytics Workspace** (30-day retention)
- **Monthly Schedule** (2:00 AM UTC on 1st of each month)

Save the outputs:

```bash
az deployment group show --resource-group rg-365audit --name main \
  --query properties.outputs -o table
```

## Step 4 — Grant Graph Permissions to Managed Identity

The Automation Account's managed identity needs permission to read the
Key Vault secret. The Bicep template handles this automatically.

No Graph permissions are needed on the Automation Account identity —
each customer tenant's audit uses the per-tenant app registration
credentials stored in Hudu.

## Step 5 — Configure Automation Account Variables

In the Azure portal (Automation Account → Shared Resources → Variables),
or via CLI:

```bash
aa="aa-365audit"
rg="rg-365audit"

az automation variable create --automation-account-name $aa -g $rg \
  --name HUDU_BASE_URL --value '"https://your-hudu.huducloud.com"'

az automation variable create --automation-account-name $aa -g $rg \
  --name KEY_VAULT_NAME --value '"<keyVaultName>"'

az automation variable create --automation-account-name $aa -g $rg \
  --name SKIP_PUBLISH --value '"true"'

# Optional:
az automation variable create --automation-account-name $aa -g $rg \
  --name MSP_DOMAINS --value '"domain1.com,domain2.com"'

az automation variable create --automation-account-name $aa -g $rg \
  --name TEST_CUSTOMERS --value '"<one-customer-slug>"'
```

## Step 6 — Wire Up GitHub Actions

### 6a — Create GitHub OIDC App Registration (if not already done)

```bash
az ad app create --display-name github-365audit-deploy
# Note the appId from output

az ad sp create --id <appId>
```

### 6b — Add Federated Credential

```json
{
  "name": "github-main",
  "issuer": "https://token.actions.githubusercontent.com",
  "subject": "repo:<owner>/365Audit:ref:refs/heads/main",
  "audiences": ["api://AzureADTokenExchange"]
}
```

```bash
az ad app federated-credential create --id <appId> --parameters <path-to-json>
```

### 6c — Grant Contributor on Resource Group

```bash
spObjectId=$(az ad sp show --id <appId> --query id -o tsv)
az role assignment create \
  --assignee-object-id $spObjectId \
  --assignee-principal-type ServicePrincipal \
  --role Contributor \
  --scope /subscriptions/<subId>/resourceGroups/rg-365audit
```

### 6d — GitHub Repository Configuration

**Secrets** (Settings → Secrets → Actions):

| Secret | Value |
|--------|-------|
| `AZURE_CLIENT_ID` | GitHub OIDC app client ID |
| `AZURE_TENANT_ID` | Azure AD tenant ID |
| `AZURE_SUBSCRIPTION_ID` | Subscription ID |

**Variables** (Settings → Variables → Actions):

| Variable | Value |
|----------|-------|
| `AUTOMATION_ACCOUNT` | `aa-365audit` |
| `RESOURCE_GROUP` | `rg-365audit` |

## Step 7 — First Deploy

Push to main to trigger GitHub Actions:

```bash
git push origin main
```

Watch the workflow: **Actions → Deploy to Azure Automation**

## Step 8 — Test

```bash
# Trigger a manual run
az automation runbook start \
  --automation-account-name aa-365audit \
  --resource-group rg-365audit \
  --name Invoke-AuditRunbook
```

Check execution in the portal: **Automation Account → Jobs**

### What to look for

- Managed identity auth succeeds
- Key Vault secret retrieval works
- Customer list sync from Hudu completes
- Audit runs and generates report files
- No Hudu publish (SKIP_PUBLISH is on)

## Step 9 — Go Live

Update the variable to enable publishing:

```bash
az automation variable update \
  --automation-account-name aa-365audit \
  --resource-group rg-365audit \
  --name SKIP_PUBLISH --value '"false"'

az automation variable delete \
  --automation-account-name aa-365audit \
  --resource-group rg-365audit \
  --name TEST_CUSTOMERS
```

The Runbook will run automatically at **2:00 AM UTC on the 1st of each month**.

---

## Cost

| Resource | Monthly Cost (AUD) |
|----------|-------------------|
| Automation Account (Free tier) | $0 (500 min/month included) |
| Log Analytics (minimal) | ~$1-2 |
| Key Vault (existing) | ~$0.05 |
| **Total** | **~$1-2** |

---

## Manual Trigger

```bash
az automation runbook start \
  --automation-account-name aa-365audit \
  --resource-group rg-365audit \
  --name Invoke-AuditRunbook
```

## Change Schedule

Update in the Azure portal (Automation Account → Schedules) or redeploy
the Bicep with updated schedule parameters.
