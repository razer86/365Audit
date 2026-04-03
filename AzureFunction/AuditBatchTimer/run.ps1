# Azure Function entry point — thin wrapper around Invoke-AzAuditBatch.ps1.
# Timer trigger fires at 2:00 AM on the 1st of each month (see function.json).
# Configure KEY_VAULT_NAME and HUDU_BASE_URL in Application Settings.

param($Timer)

Write-Host "AuditBatchTimer triggered at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss UTC' -AsUTC)"

# In the deployed package, toolkit scripts sit alongside host.json (one level up).
# Locally the repo structure has an extra AzureFunction/ layer — use TOOLKIT_ROOT
# env var to override when testing run.ps1 outside of a deployed Function App.
$toolkitRoot = if ($env:TOOLKIT_ROOT) { $env:TOOLKIT_ROOT }
               else { Split-Path -Parent $PSScriptRoot }

$batchParams = @{
    KeyVaultName        = $env:KEY_VAULT_NAME
    HuduBaseUrl         = $env:HUDU_BASE_URL
    OutputRoot          = $env:TEMP
    CleanupLocalReports = $true    # Always clean up in Azure — temp storage
    ErrorAction         = 'Stop'
}

if ($env:AUDIT_THROTTLE_LIMIT)                                    { $batchParams['ThrottleLimit']      = [int]$env:AUDIT_THROTTLE_LIMIT }
if ($env:HUDU_ASSET_LAYOUT_ID)                                    { $batchParams['HuduAssetLayoutId']  = [int]$env:HUDU_ASSET_LAYOUT_ID }
if ($env:HUDU_REPORT_LAYOUT_ID)                                   { $batchParams['HuduReportLayoutId'] = [int]$env:HUDU_REPORT_LAYOUT_ID }
if ($env:HUDU_REPORT_ASSET_NAME)                                  { $batchParams['HuduReportAssetName']= $env:HUDU_REPORT_ASSET_NAME }
if ($env:MSP_DOMAINS)                                             { $batchParams['MspDomains']         = $env:MSP_DOMAINS -split ',' }
if ($env:SKIP_PUBLISH -eq 'true')                                 { $batchParams['SkipPublish']        = $true }
if ($env:TEST_CUSTOMERS)                                          { $batchParams['Customers']          = $env:TEST_CUSTOMERS -split ',' }

& "$toolkitRoot/Invoke-AzAuditBatch.ps1" @batchParams
