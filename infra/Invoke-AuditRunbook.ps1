<#
.SYNOPSIS
    Azure Automation Runbook for 365Audit monthly batch execution.

.DESCRIPTION
    Runs as a PowerShell 7.2 Runbook in an Azure Automation Account.
    Authenticates via system-assigned managed identity, retrieves the
    Hudu API key from Key Vault, syncs the customer list from Hudu,
    then runs the full audit pipeline for each customer:
      1. Cert check / renewal (non-interactive)
      2. Run audit modules (Entra, Exchange, SharePoint, Mail Security, Intune, Teams, Maester)
      3. Generate HTML summary report
      4. Publish report to Hudu (unless SkipPublish is set)

    All scripts are deployed alongside this Runbook by GitHub Actions.

.NOTES
    Author      : Raymond Slater
    Version     : 1.0.0

    This Runbook expects the following Automation Account variables:
      HUDU_BASE_URL          — Hudu instance URL
      KEY_VAULT_NAME         — Azure Key Vault containing the Hudu API key
      MSP_DOMAINS            — Comma-separated MSP email domains (optional)
      HUDU_ASSET_LAYOUT_ID   — Credential asset layout ID (default: 67)
      HUDU_REPORT_LAYOUT_ID  — Report asset layout ID (default: 68)
      HUDU_REPORT_ASSET_NAME — Report asset name prefix (optional)
      SKIP_PUBLISH           — 'true' to skip Hudu publishing (default: 'true')
      THROTTLE_LIMIT         — Max concurrent audits (default: 3)
      TEST_CUSTOMERS         — Comma-separated customer slugs for testing (optional)
#>

#Requires -Version 7.2

$ErrorActionPreference = 'Stop'

Write-Output "365Audit Runbook started at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss UTC' -AsUTC)"

# ── Read Automation Account variables ───────────────────────────────────────
$_huduBaseUrl       = Get-AutomationVariable -Name 'HUDU_BASE_URL'
$_keyVaultName      = Get-AutomationVariable -Name 'KEY_VAULT_NAME'
$_mspDomains        = Get-AutomationVariable -Name 'MSP_DOMAINS'        -ErrorAction SilentlyContinue
$_assetLayoutId     = Get-AutomationVariable -Name 'HUDU_ASSET_LAYOUT_ID'   -ErrorAction SilentlyContinue
$_reportLayoutId    = Get-AutomationVariable -Name 'HUDU_REPORT_LAYOUT_ID'  -ErrorAction SilentlyContinue
$_reportAssetName   = Get-AutomationVariable -Name 'HUDU_REPORT_ASSET_NAME' -ErrorAction SilentlyContinue
$_skipPublish       = Get-AutomationVariable -Name 'SKIP_PUBLISH'       -ErrorAction SilentlyContinue
$_throttleLimit     = Get-AutomationVariable -Name 'THROTTLE_LIMIT'     -ErrorAction SilentlyContinue
$_testCustomers     = Get-AutomationVariable -Name 'TEST_CUSTOMERS'     -ErrorAction SilentlyContinue

# ── Defaults ────────────────────────────────────────────────────────────────
if (-not $_assetLayoutId)   { $_assetLayoutId   = 67 }
if (-not $_reportLayoutId)  { $_reportLayoutId  = 68 }
if (-not $_reportAssetName) { $_reportAssetName = 'M365 - Monthly Audit Report' }
if (-not $_skipPublish)     { $_skipPublish     = 'true' }
if (-not $_throttleLimit)   { $_throttleLimit   = 3 }

# ── Authenticate with Managed Identity ──────────────────────────────────────
Write-Output "Authenticating with Managed Identity..."
Connect-AzAccount -Identity | Out-Null
Write-Output "Authenticated."

# ── Preload Graph SDK assemblies (before Az modules clash) ──────────────────
$_scriptRoot = $PSScriptRoot
. (Join-Path $_scriptRoot 'Common' 'Audit-Common.ps1')
Initialize-GraphSdk

# ── Retrieve Hudu API key from Key Vault ────────────────────────────────────
Write-Output "Fetching Hudu API key from Key Vault '$_keyVaultName'..."
Import-Module Az.KeyVault -ErrorAction Stop
$_huduApiKey = Get-AzKeyVaultSecret -VaultName $_keyVaultName -Name '365Audit-HuduApiKey' -AsPlainText -ErrorAction Stop
if (-not $_huduApiKey) {
    throw "Key Vault secret '365Audit-HuduApiKey' not found in vault '$_keyVaultName'."
}
Write-Output "Hudu API key retrieved."

# ── Sync customer list from Hudu ────────────────────────────────────────────
$_tempDir          = $env:TEMP ?? '/tmp'
$_customerListPath = Join-Path $_tempDir 'UnattendedCustomers.psd1'
$_syncScript       = Join-Path $_scriptRoot 'Helpers' 'Sync-UnattendedCustomers.ps1'

if (Test-Path $_syncScript) {
    Write-Output "Syncing customer list from Hudu..."
    & $_syncScript `
        -HuduBaseUrl       $_huduBaseUrl `
        -HuduApiKey        $_huduApiKey `
        -HuduAssetLayoutId $_assetLayoutId `
        -OutputFilePath    $_customerListPath `
        -ErrorAction Stop
}

# ── Load customer list ──────────────────────────────────────────────────────
if (-not (Test-Path $_customerListPath)) {
    throw "Customer list not found at $_customerListPath. Sync may have failed."
}

$_customerData = Import-PowerShellDataFile -Path $_customerListPath
$_customerList = @($_customerData.Customers)

if ($_customerList.Count -eq 0) {
    throw "No customers in synced customer list."
}

# Filter to test customers if specified
if ($_testCustomers) {
    $_testSlugs    = $_testCustomers -split ','
    $_customerList = @($_customerList | Where-Object { $_.HuduCompanySlug -in $_testSlugs })
    if ($_customerList.Count -eq 0) {
        throw "None of the test customers matched the synced list."
    }
}

# ── Output configuration ────────────────────────────────────────────────────
$_outputRoot = Join-Path $_tempDir '365audit'
New-Item -ItemType Directory -Path $_outputRoot -Force | Out-Null

# ── Run audit for each customer ─────────────────────────────────────────────
$_auditScript   = Join-Path $_scriptRoot 'Start-365Audit.ps1'
$_publishScript = Join-Path $_scriptRoot 'Helpers' 'Publish-HuduAuditReport.ps1'
$_totalCustomers = $_customerList.Count
$_currentIndex   = 0
$_results        = [System.Collections.Generic.List[PSCustomObject]]::new()

Write-Output "`n$('=' * 72)"
Write-Output "  365Audit Runbook — $_totalCustomers customer(s)"
Write-Output "$('=' * 72)"

foreach ($_entry in $_customerList) {
    $_currentIndex++
    $_slug = $_entry.HuduCompanySlug
    $_mods = @($_entry.Modules ?? @('A'))
    $_label = "[$_currentIndex/$_totalCustomers] $_slug"

    Write-Output "`n$('=' * 72)"
    Write-Output "  $_label  (modules: $($_mods -join ','))"
    Write-Output "$('=' * 72)"

    $_tenantStart = Get-Date
    $_result = [PSCustomObject]@{
        Customer    = $_slug
        Modules     = $_mods -join ','
        AuditStatus = 'Pending'
        Error       = $null
        Elapsed     = $null
    }

    try {
        # Run audit
        $_auditParams = @{
            HuduCompanyId = $_slug
            HuduBaseUrl   = $_huduBaseUrl
            HuduApiKey    = $_huduApiKey
            Modules       = $_mods
            OutputRoot    = $_outputRoot
            ErrorAction   = 'Stop'
        }
        if ($_mspDomains) { $_auditParams['MspDomains'] = $_mspDomains -split ',' }

        & $_auditScript @_auditParams

        $_result.AuditStatus = 'Completed'
        $_elapsed = (Get-Date) - $_tenantStart
        $_result.Elapsed = [math]::Round($_elapsed.TotalMinutes, 1)
        Write-Output "  $_label — DONE ($($_result.Elapsed)m)"

        # Publish to Hudu
        if ($_skipPublish -ne 'true' -and (Test-Path $_publishScript)) {
            $_lastOutputFile = Join-Path $env:TEMP '365Audit_LastOutput.txt'
            $_customerOutputPath = if (Test-Path $_lastOutputFile) {
                (Get-Content $_lastOutputFile -Raw -ErrorAction SilentlyContinue).Trim()
            } else { $null }

            if ($_customerOutputPath -and (Test-Path $_customerOutputPath)) {
                Write-Output "  Publishing report to Hudu..."
                try {
                    & $_publishScript `
                        -OutputPath       $_customerOutputPath `
                        -CompanySlug      $_slug `
                        -HuduBaseUrl      $_huduBaseUrl `
                        -HuduApiKey       $_huduApiKey `
                        -ReportLayoutId   $_reportLayoutId `
                        -ReportAssetName  $_reportAssetName
                }
                catch {
                    Write-Warning "  Hudu publish failed for ${_slug}: $($_.Exception.Message)"
                }
            }
        }
    }
    catch {
        $_result.AuditStatus = 'Failed'
        $_result.Error = $_.Exception.Message
        $_elapsed = (Get-Date) - $_tenantStart
        $_result.Elapsed = [math]::Round($_elapsed.TotalMinutes, 1)
        Write-Output "  $_label — FAILED: $($_.Exception.Message)"
    }

    $_results.Add($_result)
}

# ── Summary ─────────────────────────────────────────────────────────────────
Write-Output "`n$('=' * 72)"
Write-Output "  Batch complete — $($_results.Count) customer(s)"
Write-Output "$('=' * 72)"

$_results | Format-Table -AutoSize | Out-String -Width 120

$_completedCount = @($_results | Where-Object { $_.AuditStatus -eq 'Completed' }).Count
$_failedCount    = @($_results | Where-Object { $_.AuditStatus -eq 'Failed'    }).Count

Write-Output "365Audit Runbook finished at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss UTC' -AsUTC)"
Write-Output "  Completed: $_completedCount  |  Failed: $_failedCount"

if ($_failedCount -gt 0) {
    Write-Error "$_failedCount customer(s) failed — review output above."
}
