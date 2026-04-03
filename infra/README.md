# Azure Functions Deployment Guide

Step-by-step guide to deploy 365Audit as an Azure Function. This process
will need to be run twice — once in the **private/test tenant** and again
in the **company tenant** when ready for production.

---

## Prerequisites

| Tool | Install | Check |
|------|---------|-------|
| Azure CLI | `winget install Microsoft.AzureCLI` or [Azure Cloud Shell](https://shell.azure.com) | `az --version` |
| GitHub CLI | `winget install GitHub.cli` | `gh --version` |
| Azure subscription | With credits available | `az account show` |

---

## Overview

```
┌─────────────────────────────────────────────────────────────┐
│  1. Deploy infra (Bicep)    → Resource Group, Function App, │
│                                Key Vault, Storage, Identity │
│  2. Store secrets           → Hudu API key in Key Vault     │
│  3. First deploy (manual)   → Zip-deploy the function app   │
│  4. Test with SkipPublish   → Validate audits run correctly │
│  5. Wire GitHub Actions     → Auto-deploy on push to main   │
│  6. Go live                 → Set SkipPublish=false          │
└─────────────────────────────────────────────────────────────┘
```

---

## Step 1 — Login and Create Resource Group

```bash
# Login (opens browser)
az login

# Verify correct subscription
az account show --query "{name:name, id:id}" -o table

# If you need to switch subscriptions:
# az account set --subscription "<subscription-id>"

# Create resource group (australiaeast — closest region)
az group create --name rg-365audit --location australiaeast
```

## Step 2 — Deploy Infrastructure

The Bicep template creates:
- **App Service Plan** (B1 Linux, ~$20 AUD/month)
- **Function App** (PowerShell 7.4, system-assigned Managed Identity)
- **Storage Account** (required by Functions runtime)
- **Key Vault** (RBAC-enabled, Function App granted Secrets User role)

```bash
az deployment group create \
  --resource-group rg-365audit \
  --template-file infra/main.bicep \
  --parameters \
    functionAppName='func-365audit-<suffix>' \
    huduBaseUrl='https://neconnect.huducloud.com'
```

> **Naming:** `functionAppName` must be globally unique. Use a short suffix
> like your initials or `dev`/`prod` (e.g. `func-365audit-rs`, `func-365audit-prod`).

> **SkipPublish:** Defaults to `true` in the Bicep template, so the first
> deployment will not push reports to Hudu. You'll flip this later.

Save the outputs — you'll need them in later steps:

```bash
# Show deployment outputs
az deployment group show \
  --resource-group rg-365audit \
  --name main \
  --query properties.outputs -o table
```

Key outputs:
| Output | Used for |
|--------|----------|
| `functionAppName` | All subsequent az commands, GitHub secret |
| `keyVaultName` | Storing the Hudu API key |
| `managedIdentityPrincipalId` | Troubleshooting RBAC if needed |

## Step 3 — Store Hudu API Key in Key Vault

```bash
az keyvault secret set \
  --vault-name <keyVaultName> \
  --name '365Audit-HuduApiKey' \
  --value '<your-hudu-api-key>'
```

Verify it was stored:

```bash
az keyvault secret show \
  --vault-name <keyVaultName> \
  --name '365Audit-HuduApiKey' \
  --query "name" -o tsv
```

## Step 4 — First Deploy (Manual Zip)

Before wiring up GitHub Actions, do a manual deployment to verify
everything works.

```bash
# From the repo root — package the function app
mkdir -p deploy

# Azure Function scaffolding
cp AzureFunction/host.json deploy/
cp AzureFunction/profile.ps1 deploy/
cp AzureFunction/requirements.psd1 deploy/
cp -r AzureFunction/AuditBatchTimer deploy/

# Toolkit scripts (alongside host.json — run.ps1 expects this layout)
cp Invoke-AzAuditBatch.ps1 deploy/
cp Start-365Audit.ps1 deploy/
cp Start-UnattendedAudit.ps1 deploy/
cp Setup-365AuditApp.ps1 deploy/
cp Generate-AuditSummary.ps1 deploy/
cp Invoke-*.ps1 deploy/
cp -r Common deploy/
cp -r Helpers deploy/
mkdir -p deploy/Resources

# Create zip and deploy
cd deploy && zip -r ../365audit-func.zip . && cd ..

az functionapp deployment source config-zip \
  --resource-group rg-365audit \
  --name <functionAppName> \
  --src 365audit-func.zip

# Clean up
rm -rf deploy 365audit-func.zip
```

> **Windows without zip:** Use PowerShell instead:
> ```powershell
> Compress-Archive -Path deploy\* -DestinationPath 365audit-func.zip -Force
> ```

## Step 5 — Upload Customer List (or let Sync handle it)

The customer list (`UnattendedCustomers.psd1`) is not included in the
deploy package because it contains customer data.

**Option A — Let Sync-UnattendedCustomers create it automatically:**
The batch script runs `Sync-UnattendedCustomers.ps1` at startup, which
queries Hudu and generates the file. This is the default behaviour —
no manual upload needed if Hudu credentials are working.

**Option B — Upload manually (if sync is disabled or for testing):**

```bash
az functionapp deploy \
  --resource-group rg-365audit \
  --name <functionAppName> \
  --src-path UnattendedCustomers.psd1 \
  --target-path UnattendedCustomers.psd1 \
  --type static
```

Or via the Azure Portal: Function App → Advanced Tools (Kudu) →
Debug console → navigate to `/home/site/wwwroot/` → drag and drop.

## Step 6 — Test the Function

Trigger the function manually. Since `SKIP_PUBLISH=true`, it will run
audits and generate reports but won't push anything to Hudu.

**Via Azure Portal:**
Function App → Functions → AuditBatchTimer → Code + Test → Test/Run

**Via CLI:**

```bash
az functionapp function invoke \
  --resource-group rg-365audit \
  --name <functionAppName> \
  --function-name AuditBatchTimer
```

**Check logs:**

```bash
# Live log stream
az functionapp log tail \
  --resource-group rg-365audit \
  --name <functionAppName>
```

Or in the Portal: Function App → Log stream

### What to look for

- Managed Identity auth succeeds (profile.ps1)
- Key Vault secret retrieval works
- Customer list sync from Hudu completes
- At least one customer audit runs to completion
- Reports are generated in `$env:TEMP` (visible in logs)
- No Hudu publish attempts (SkipPublish is on)

---

## Step 7 — Wire Up GitHub Actions (CI/CD)

This enables automatic deployment when code is pushed to `main`.
Skip this step if you only want manual deployments for now.

### 7a — Create an Azure AD App Registration for OIDC

```bash
# Create app registration
az ad app create --display-name 'github-365audit-deploy' --query appId -o tsv
# Save this appId ↑

# Create service principal
az ad sp create --id <appId>

# Add federated credential for GitHub Actions
az ad app federated-credential create --id <appId> --parameters '{
  "name": "github-main",
  "issuer": "https://token.actions.githubusercontent.com",
  "subject": "repo:razer86/365Audit:ref:refs/heads/main",
  "audiences": ["api://AzureADTokenExchange"]
}'

# Grant Contributor role on the resource group
az role assignment create \
  --assignee <appId> \
  --role Contributor \
  --scope "/subscriptions/<subscriptionId>/resourceGroups/rg-365audit"
```

### 7b — Add GitHub Repository Secrets

Go to **Settings → Secrets and variables → Actions** and add:

| Secret | Value |
|--------|-------|
| `AZURE_CLIENT_ID` | App registration appId from 7a |
| `AZURE_TENANT_ID` | Azure AD tenant ID (`az account show --query tenantId -o tsv`) |
| `AZURE_SUBSCRIPTION_ID` | Subscription ID (`az account show --query id -o tsv`) |
| `AZURE_FUNCTION_APP_NAME` | Function App name from Step 2 output |

### 7c — Test the Workflow

Push a commit to `main` (or merge the feature branch) and check:
GitHub → Actions → "Deploy to Azure Functions" → verify it completes.

---

## Step 8 — Go Live

When you're satisfied the audits are running correctly:

```bash
# Disable SkipPublish so reports are pushed to Hudu
az functionapp config appsettings set \
  --resource-group rg-365audit \
  --name <functionAppName> \
  --settings SKIP_PUBLISH=false
```

Or in the Portal: Function App → Configuration → Application settings →
edit `SKIP_PUBLISH` → set to `false` → Save.

The timer trigger fires at **2:00 AM UTC on the 1st of each month**.
To change the schedule, edit `AzureFunction/AuditBatchTimer/function.json`:

```json
{ "schedule": "0 0 2 1 * *" }
```

Format: `{sec} {min} {hour} {day} {month} {day-of-week}` (NCRONTAB).

---

## Repeating for Company Tenant

When deploying to the production/company tenant:

1. `az login` with company credentials (or switch tenant: `az login --tenant <companyTenantId>`)
2. Repeat Steps 1–6 with production values:
   - Resource group: `rg-365audit` (or your naming convention)
   - Function app name: `func-365audit-prod` (or similar)
   - Hudu base URL: production Hudu instance
   - Hudu API key: production key
3. For GitHub Actions (Step 7): add a second federated credential and use
   GitHub Environments (`production`) with separate secrets, or maintain
   a separate workflow file per tenant
4. Set `SKIP_PUBLISH=false` when ready (Step 8)

---

## Cost Management

The B1 App Service Plan costs ~$20 AUD/month. Since audits run monthly,
you can scale down between runs to save costs:

```bash
# After the audit completes — scale to Free (stops billing)
az appservice plan update -n asp-365audit -g rg-365audit --sku FREE

# Before the next audit — scale back to B1
az appservice plan update -n asp-365audit -g rg-365audit --sku B1
```

> **Note:** On the Free tier the Function App will still exist but won't
> execute timer triggers. The next run will fire when you scale back to B1.

---

## Troubleshooting

| Symptom | Check |
|---------|-------|
| Function doesn't trigger | Is the plan on B1? (Free tier can't run timers) |
| Key Vault access denied | Managed Identity role assignment: `az role assignment list --scope <kvResourceId>` |
| Graph SDK assembly errors | Check `host.json` timeout — long-running audits may hit the 4hr limit |
| Customer sync fails | Verify Hudu API key is correct: test from Cloud Shell with `curl` |
| Zip deploy fails | Ensure `WEBSITE_RUN_FROM_PACKAGE=1` is set in app settings |
| Logs not appearing | Enable Application Insights (not included in Bicep — add if needed) |

---

## Tearing Down (Test Environment)

```bash
# Remove everything in one go
az group delete --name rg-365audit --yes --no-wait

# Clean up the app registration (if created for GitHub Actions)
az ad app delete --id <appId>
```
