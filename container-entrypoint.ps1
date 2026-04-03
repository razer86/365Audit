<#
.SYNOPSIS
    Container Apps Job entry point for 365Audit.

.DESCRIPTION
    Authenticates with Azure Managed Identity, then runs Invoke-AzAuditBatch.ps1
    with configuration from environment variables.

    Replaces the Azure Functions AuditBatchTimer/run.ps1 entry point.
    In a standalone pwsh container:
      - Start-Job works (no hosted PowerShell restriction)
      - Graph SDK + Az module assembly conflicts do not occur
      - Filesystem is fully writable
      - $env:TEMP is set via Dockerfile ENV

.NOTES
    Author      : Raymond Slater
    Version     : 1.0.0
#>

$ErrorActionPreference = 'Stop'

Write-Host "365Audit container started at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss UTC' -AsUTC)"

# ── Authenticate with User-Assigned Managed Identity ─────────────────────────
# AZURE_CLIENT_ID is the clientId of the user-assigned managed identity,
# injected by the Container Apps Job environment configuration (Bicep).
$clientId = $env:AZURE_CLIENT_ID
if (-not $clientId) {
    Write-Error "AZURE_CLIENT_ID environment variable is not set. Cannot authenticate with Managed Identity."
}

Write-Host "Authenticating with Managed Identity (ClientId: $clientId)..."
Connect-AzAccount -Identity -AccountId $clientId | Out-Null
Write-Host "Authenticated."

# ── Build parameters for Invoke-AzAuditBatch.ps1 ────────────────────────────
$batchParams = @{
    KeyVaultName        = $env:KEY_VAULT_NAME
    HuduBaseUrl         = $env:HUDU_BASE_URL
    OutputRoot          = '/tmp/365audit'
    CleanupLocalReports = $true
    ErrorAction         = 'Stop'
}

if ($env:AUDIT_THROTTLE_LIMIT)   { $batchParams['ThrottleLimit']       = [int]$env:AUDIT_THROTTLE_LIMIT }
if ($env:HUDU_ASSET_LAYOUT_ID)   { $batchParams['HuduAssetLayoutId']   = [int]$env:HUDU_ASSET_LAYOUT_ID }
if ($env:HUDU_REPORT_LAYOUT_ID)  { $batchParams['HuduReportLayoutId']  = [int]$env:HUDU_REPORT_LAYOUT_ID }
if ($env:HUDU_REPORT_ASSET_NAME) { $batchParams['HuduReportAssetName'] = $env:HUDU_REPORT_ASSET_NAME }
if ($env:MSP_DOMAINS)            { $batchParams['MspDomains']          = $env:MSP_DOMAINS -split ',' }
if ($env:SKIP_PUBLISH -eq 'true'){ $batchParams['SkipPublish']         = $true }
if ($env:TEST_CUSTOMERS)         { $batchParams['Customers']           = $env:TEST_CUSTOMERS -split ',' }

& "$PSScriptRoot/Invoke-AzAuditBatch.ps1" @batchParams

Write-Host "365Audit container finished at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss UTC' -AsUTC)"
