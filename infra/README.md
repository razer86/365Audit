# Azure Deployment Guide

## Prerequisites

- Azure CLI installed (`az --version`)
- GitHub CLI installed (`gh --version`)
- An Azure subscription with credits

## Step 1: Deploy Infrastructure

```bash
# Login to Azure
az login

# Create resource group (Australia East — closest to you)
az group create --name rg-365audit --location australiaeast

# Deploy infrastructure
az deployment group create \
  --resource-group rg-365audit \
  --template-file infra/main.bicep \
  --parameters functionAppName=func-365audit-<yoursuffix> \
               huduBaseUrl=https://neconnect.huducloud.com

# Note the outputs — you'll need functionAppName and keyVaultName
```

## Step 2: Store Hudu API Key in Key Vault

```bash
az keyvault secret set \
  --vault-name <keyVaultName-from-output> \
  --name 365Audit-HuduApiKey \
  --value '<your-hudu-api-key>'
```

## Step 3: Set Up GitHub OIDC for Deployments

This lets GitHub Actions deploy to Azure without storing long-lived secrets.

```bash
# Create an app registration for GitHub
az ad app create --display-name "github-365audit-deploy"
# Note the appId from the output

# Create a service principal
az ad sp create --id <appId>

# Add federated credential (replace <your-github-org/repo>)
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
  --scope /subscriptions/<subscriptionId>/resourceGroups/rg-365audit
```

## Step 4: Add GitHub Secrets

Go to your repo → Settings → Secrets and variables → Actions, and add:

| Secret | Value |
|--------|-------|
| `AZURE_CLIENT_ID` | App registration client ID from Step 3 |
| `AZURE_TENANT_ID` | Your Azure AD tenant ID |
| `AZURE_SUBSCRIPTION_ID` | Your Azure subscription ID |
| `AZURE_FUNCTION_APP_NAME` | Function App name from Step 1 output |

## Step 5: Upload Customer List

The `UnattendedCustomers.psd1` file contains customer data and is not
deployed via GitHub Actions. Upload it separately:

```bash
# Upload via Azure CLI (Kudu ZIP deploy of a single file)
az functionapp deploy \
  --resource-group rg-365audit \
  --name <functionAppName> \
  --src-path UnattendedCustomers.psd1 \
  --target-path UnattendedCustomers.psd1 \
  --type static
```

Or upload via the Azure Portal: Function App → Advanced Tools (Kudu) →
Debug console → navigate to `/home/site/wwwroot/` → drag and drop the file.

## Step 6: Test

Trigger the function manually from the Azure Portal:

Function App → Functions → AuditBatchTimer → Code + Test → Test/Run

Or via CLI:

```bash
az functionapp function invoke \
  --resource-group rg-365audit \
  --name <functionAppName> \
  --function-name AuditBatchTimer
```

## Cost Management

The B1 App Service Plan costs ~$20 AUD/month. Since audits run monthly,
you can scale down between runs:

```bash
# After the audit completes — scale to Free (stops billing)
az appservice plan update -n asp-365audit -g rg-365audit --sku FREE

# Before the next audit — scale back to B1
az appservice plan update -n asp-365audit -g rg-365audit --sku B1
```

This can be automated with a second timer-triggered function or Azure
Automation runbook if desired.
