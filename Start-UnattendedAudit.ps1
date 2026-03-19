<#
.SYNOPSIS
    Runs the 365Audit toolkit across multiple customers in sequence.

.DESCRIPTION
    Reads the customer list from UnattendedCustomers.psd1 (one entry per customer),
    then for each customer:
      1. Fetches credentials from the Hudu '365Audit' asset automatically.
      2. Checks the audit certificate expiry — if within 30 days, calls
         Setup-365AuditApp.ps1 to renew the cert and push new credentials
         back to Hudu (no browser login required).
      3. Runs the audit modules defined for that customer via Start-365Audit.ps1.
      4. Generates the HTML summary report (not opened automatically).

    SETUP:
      1. Copy UnattendedCustomers.psd1.example to UnattendedCustomers.psd1
      2. Edit UnattendedCustomers.psd1 — add a HuduCompanySlug + Modules entry per customer
         (UnattendedCustomers.psd1 is excluded from git to keep customer data private)
      3. Add HuduApiKey to config.psd1 (or pass -HuduApiKey at runtime)
      4. Run: .\Start-UnattendedAudit.ps1

    NOTE: Non-interactive cert renewal requires that Setup-365AuditApp.ps1
    has been run interactively at least once per customer tenant. This grants
    Application.ReadWrite.OwnedBy and registers the service principal as an
    owner of the app registration.

.PARAMETER Customers
    Optional filter — one or more HuduCompanySlugs to process. When omitted,
    all entries in UnattendedCustomers.psd1 are processed.

.PARAMETER Modules
    Global module override — applies to every customer this run.
    When omitted, each customer uses the Modules array from the PSD1 file.
    Valid values: 1, 2, 3, 4, 9.

.PARAMETER HuduBaseUrl
    Hudu instance base URL. Falls back to HUDU_BASE_URL env var, then
    'https://neconnect.huducloud.com'.

.PARAMETER HuduApiKey
    Hudu API key. Falls back to HUDU_API_KEY env var.

.PARAMETER SkipCertCheck
    Skip the certificate expiry check and renewal step.

.EXAMPLE
    .\Start-UnattendedAudit.ps1
    Processes all customers in UnattendedCustomers.psd1 using HuduApiKey from config.psd1.

.EXAMPLE
    .\Start-UnattendedAudit.ps1 -Customers 'contoso','fabrikam' -Modules 1,2
    Runs only Entra and Exchange audits for two specific customers.

.NOTES
    Author      : Raymond Slater
    Version     : 2.3.0

.LINK
    https://github.com/razer86/365Audit
#>

#Requires -Version 7.2

param (
    [string[]]$Customers,

    [ValidateSet(1, 2, 3, 4, 9)]
    [int[]]$Modules,

    [string]$HuduBaseUrl,
    [string]$HuduApiKey,

    [switch]$SkipCertCheck
)

$ScriptVersion = "2.3.0"
Write-Verbose "Start-UnattendedAudit.ps1 loaded (v$ScriptVersion)"

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Load config.psd1 from the script root — fallback for HuduApiKey / HuduBaseUrl.
# Explicit command-line parameters always take precedence over config file values.
$_configPath = Join-Path $PSScriptRoot 'config.psd1'
if (Test-Path $_configPath) {
    try {
        $_config = Import-PowerShellDataFile -Path $_configPath
        if (-not $HuduApiKey  -and $_config.HuduApiKey)  { $HuduApiKey  = $_config.HuduApiKey }
        if (-not $HuduBaseUrl -and $_config.HuduBaseUrl) { $HuduBaseUrl = $_config.HuduBaseUrl }
    }
    catch { Write-Warning "Could not load config.psd1: $_" }
}
if (-not $HuduBaseUrl) { $HuduBaseUrl = 'https://neconnect.huducloud.com' }

# ── Load customer list from PSD1 ───────────────────────────────────────────────
$customerListPath = Join-Path $PSScriptRoot 'UnattendedCustomers.psd1'
if (-not (Test-Path $customerListPath)) {
    Write-Error ("Customer list not found: $customerListPath`n" +
        "Copy UnattendedCustomers.psd1.example to UnattendedCustomers.psd1 " +
        "and add your customers.") -ErrorAction Stop
}

try {
    $_customerData = Import-PowerShellDataFile -Path $customerListPath
}
catch {
    Write-Error "Could not load UnattendedCustomers.psd1: $_" -ErrorAction Stop
}
$customerList = @($_customerData.Customers)

if ($customerList.Count -eq 0) {
    Write-Error "No customers defined in UnattendedCustomers.psd1." -ErrorAction Stop
}

# Optional: filter to specific slugs passed on the command line
if ($Customers) {
    $customerList = @($customerList | Where-Object { $_.HuduCompanySlug -in $Customers })
    if ($customerList.Count -eq 0) {
        Write-Error "None of the specified customers matched entries in UnattendedCustomers.psd1." -ErrorAction Stop
    }
}
# ─────────────────────────────────────────────────────────────────────────────

if (-not $HuduApiKey) {
    Write-Error "HUDU_API_KEY is not set. Set the environment variable or pass -HuduApiKey." -ErrorAction Stop
}

$scriptRoot   = $PSScriptRoot
$setupScript  = Join-Path $scriptRoot 'Setup-365AuditApp.ps1'
$auditScript  = Join-Path $scriptRoot 'Start-365Audit.ps1'
$results      = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($scriptPath in @($setupScript, $auditScript)) {
    if (-not (Test-Path $scriptPath)) {
        Write-Error "Required script not found: $scriptPath" -ErrorAction Stop
    }
}

$totalCustomers = $customerList.Count
$currentIndex   = 0
$failed         = @()

foreach ($entry in $customerList) {
        $currentIndex++
        $customerId   = $entry.HuduCompanySlug
        $customerMods = if ($Modules) { $Modules } else { [int[]]($entry.Modules ?? @(1, 2, 3, 4)) }
        $customerLabel = "[$currentIndex/$totalCustomers] $customerId"

        Write-Host "`n$('=' * 72)" -ForegroundColor Cyan
        Write-Host "  $customerLabel  (modules: $($customerMods -join ','))" -ForegroundColor Cyan
        Write-Host "$('=' * 72)" -ForegroundColor Cyan

        $result = [PSCustomObject]@{
            Customer    = $customerId
            Modules     = $customerMods -join ','
            CertRenewed = $false
            AuditStatus = 'Pending'
            Error       = $null
        }

        try {
            # --- Step 1: Cert check and renewal ---
            if (-not $SkipCertCheck) {
                Write-Host "  Checking certificate expiry..." -ForegroundColor DarkCyan
                & $setupScript `
                    -HuduCompanyId $customerId `
                    -HuduBaseUrl   $HuduBaseUrl `
                    -HuduApiKey    $HuduApiKey `
                    -ErrorAction Stop
            }

            # --- Step 2: Run audit ---
            Write-Host "  Starting audit (modules: $($customerMods -join ','))..." -ForegroundColor DarkCyan
            & $auditScript `
                -HuduCompanyId $customerId `
                -HuduBaseUrl   $HuduBaseUrl `
                -HuduApiKey    $HuduApiKey `
                -Modules       $customerMods `
                -ErrorAction Stop

            $result.AuditStatus = 'Completed'
            Write-Host "  $customerLabel — DONE" -ForegroundColor Green
        }
        catch {
            $result.AuditStatus = 'Failed'
            $result.Error       = $_.Exception.Message
            Write-Host "  $customerLabel — FAILED: $($_.Exception.Message)" -ForegroundColor Red
        }

        $results.Add($result)
}

# ── Final summary ─────────────────────────────────────────────────────────────
Write-Host "`n$('=' * 72)" -ForegroundColor Cyan
Write-Host "  Bulk run complete — $($results.Count) customer(s)" -ForegroundColor Cyan
Write-Host "$('=' * 72)" -ForegroundColor Cyan

$results | Format-Table -AutoSize

$failed = @($results | Where-Object { $_.AuditStatus -eq 'Failed' })
if ($failed.Count -gt 0) {
    Write-Host "  $($failed.Count) customer(s) failed — review errors above." -ForegroundColor Red
    exit 1
}
