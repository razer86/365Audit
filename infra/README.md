# Azure Container Apps Job Deployment Guide

Deploys 365Audit as a scheduled Container Apps Job. Run this process once
per environment (test tenant, then production).

---

## Prerequisites

| Tool | Install | Check |
|------|---------|-------|
| Azure CLI | `winget install Microsoft.AzureCLI` or [Cloud Shell](https://shell.azure.com) | `az --version` |
| GitHub CLI | `winget install GitHub.cli` | `gh --version` |
| Docker | [Docker Desktop](https://docs.docker.com/desktop/) (optional, for local testing) | `docker --version` |

---

## Architecture

```
GitHub Actions (on push)
  └─ Build Docker image → push to ACR
  └─ Update Container Apps Job with new image

Container Apps Job (cron: 2am on 1st of month)
  └─ Pull image from ACR (Managed Identity)
  └─ Run container-entrypoint.ps1
      └─ Connect-AzAccount -Identity
      └─ Fetch Hudu API key from Key Vault
      └─ Sync customer list from Hudu
      └─ Run audit for each customer
      └─ Generate reports
```

---

## Step 1 — Login and Resource Group

```bash
az login --tenant <tenantId>
az account set --subscription "<subscriptionId>"

# Create resource group (or reuse existing)
az group create --name rg-365audit --location australiaeast
```

## Step 2 — Ensure Key Vault Exists

If migrating from Azure Functions, the Key Vault already exists. If deploying
fresh, create it first and store the Hudu API key:

```bash
# Only if Key Vault doesn't exist yet:
az keyvault create --name <keyVaultName> --resource-group rg-365audit \
  --location australiaeast --enable-rbac-authorization

az keyvault secret set --vault-name <keyVaultName> \
  --name 365Audit-HuduApiKey --value '<your-hudu-api-key>'
```

Grant yourself admin access if needed:

```bash
az role assignment create --assignee <your-email> \
  --role "Key Vault Administrator" \
  --scope /subscriptions/<subId>/resourceGroups/rg-365audit/providers/Microsoft.KeyVault/vaults/<keyVaultName>
```

## Step 3 — Deploy Infrastructure

```bash
az deployment group create \
  --resource-group rg-365audit \
  --template-file infra/main.bicep \
  --parameters \
    keyVaultName='<keyVaultName>' \
    huduBaseUrl='https://neconnect.huducloud.com' \
    mspDomains='ntit.com.au,nqbe.com.au,capconnect.com.au,widebayit.com.au,neconnect.com.au'
```

This creates:
- **Azure Container Registry** (Basic, ~$7 AUD/month)
- **User-Assigned Managed Identity** with AcrPull + Key Vault Secrets User roles
- **Log Analytics Workspace** (30-day retention)
- **Container Apps Environment**
- **Container Apps Job** (scheduled: 2am on 1st of month, 2 CPU / 4 GB, 4hr timeout)

Save the outputs:

```bash
az deployment group show --resource-group rg-365audit --name main \
  --query properties.outputs -o table
```

| Output | Used for |
|--------|----------|
| `acrLoginServer` | Docker push target |
| `acrName` | GitHub Actions variable |
| `jobName` | Manual trigger commands |
| `managedIdentityClientId` | Troubleshooting auth |

## Step 4 — Wire Up GitHub Actions

### 4a — Grant AcrPush to GitHub OIDC Service Principal

```bash
az role assignment create \
  --assignee <github-app-id> \
  --role AcrPush \
  --scope /subscriptions/<subId>/resourceGroups/rg-365audit/providers/Microsoft.ContainerRegistry/registries/<acrName>
```

### 4b — Add Federated Credential (if new branch)

Write the JSON to a temp file (PowerShell mangles inline JSON):

```json
{
  "name": "github-container-apps",
  "issuer": "https://token.actions.githubusercontent.com",
  "subject": "repo:razer86/365Audit:ref:refs/heads/feature/container-apps",
  "audiences": ["api://AzureADTokenExchange"]
}
```

```bash
az ad app federated-credential create --id <github-app-id> --parameters <path-to-json>
```

### 4c — GitHub Repository Configuration

**Secrets** (Settings → Secrets and variables → Actions → Secrets):

| Secret | Value |
|--------|-------|
| `AZURE_CLIENT_ID` | GitHub OIDC app registration client ID |
| `AZURE_TENANT_ID` | Azure AD tenant ID |
| `AZURE_SUBSCRIPTION_ID` | Subscription ID |

**Variables** (Settings → Secrets and variables → Actions → Variables):

| Variable | Value |
|----------|-------|
| `ACR_NAME` | Container Registry name from Step 3 output |
| `JOB_NAME` | `job-365audit` (default) |
| `RESOURCE_GROUP` | `rg-365audit` (default) |

## Step 5 — First Deploy

Push to the branch to trigger GitHub Actions:

```bash
git push origin feature/container-apps
```

Watch the workflow: **Actions → Build & Deploy to Container Apps**

The first Docker build will take 10-15 minutes (downloading PowerShell modules).
Subsequent builds are faster due to layer caching.

## Step 6 — Test

Set test configuration:

```bash
az containerapp job update \
  --name job-365audit --resource-group rg-365audit \
  --set-env-vars TEST_CUSTOMERS=<one-customer-slug> SKIP_PUBLISH=true
```

Trigger a manual run:

```bash
az containerapp job start \
  --name job-365audit --resource-group rg-365audit
```

Check execution status:

```bash
az containerapp job execution list \
  --name job-365audit --resource-group rg-365audit -o table
```

View logs:

```bash
az containerapp logs show \
  --name job-365audit --resource-group rg-365audit --type console
```

### What to look for

- Managed Identity auth succeeds
- Key Vault secret retrieval works
- Customer list sync from Hudu completes
- Graph SDK connects without assembly errors
- Audit runs and generates report files
- No Hudu publish (SKIP_PUBLISH is on)

## Step 7 — Go Live

Remove test restrictions and enable publishing:

```bash
az containerapp job update \
  --name job-365audit --resource-group rg-365audit \
  --set-env-vars SKIP_PUBLISH=false \
  --remove-env-vars TEST_CUSTOMERS
```

The job will run automatically at **2:00 AM UTC on the 1st of each month**.

---

## Repeating for Production Tenant

1. `az login --tenant <productionTenantId>`
2. Repeat Steps 1-6 with production values
3. Use separate GitHub Environments with their own secrets/variables

---

## Cost

| Resource | Monthly Cost (AUD) |
|----------|-------------------|
| Container Registry (Basic) | ~$7 |
| Container Apps Job (per execution) | ~$0.50 per 2hr run |
| Log Analytics | ~$1-2 (minimal ingestion) |
| **Total** | **~$8-10** |

Compare: Azure Functions B1 plan was ~$20/month always-on.

---

## Manual Trigger

```bash
az containerapp job start --name job-365audit --resource-group rg-365audit
```

## Change Schedule

Update the cron expression in `infra/main.bicep` and redeploy, or:

```bash
az containerapp job update \
  --name job-365audit --resource-group rg-365audit \
  --cron-expression "0 2 1 * *"
```

Format: `{min} {hour} {day} {month} {dow}` (5-field standard cron, UTC).

---

## Cleanup Old Azure Functions Resources

After confirming Container Apps Job works:

```bash
az functionapp delete --name azfunc-m365audit-neconnect --resource-group rg-365audit
az appservice plan delete --name asp-365audit --resource-group rg-365audit --yes
az storage account delete --name st365auditt36i7razytcpo --resource-group rg-365audit --yes
```

## Tear Down (Test Environment)

```bash
az group delete --name rg-365audit --yes --no-wait
az ad app delete --id <github-app-id>
```
