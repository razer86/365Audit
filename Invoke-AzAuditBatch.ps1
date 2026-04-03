<#
.SYNOPSIS
    Runs 365Audit across multiple customers concurrently.

.DESCRIPTION
    Reads the customer list from UnattendedCustomers.psd1 and processes each
    customer audit as a separate PowerShell background job for process isolation.
    Designed to run both locally (for testing) and as an Azure Function.

    For each customer:
      1. Checks certificate expiry and renews if needed (via Setup-365AuditApp.ps1).
      2. Runs the audit modules (via Start-365Audit.ps1).
      3. Generates the HTML summary report.
      4. Publishes the report to Hudu (sequential, to avoid API rate limits).

    Credential resolution order:
      1. -HuduApiKey parameter (explicit, for local testing).
      2. Azure Key Vault secret via -KeyVaultName (Azure mode, Managed Identity).
      3. config.psd1 file in the script root.
      4. HUDU_API_KEY environment variable.

    SETUP:
      1. Copy UnattendedCustomers.psd1.example to UnattendedCustomers.psd1
      2. Add customer entries with HuduCompanySlug and Modules
      3. Ensure credentials are available (Key Vault or config.psd1)
      4. Run: .\Invoke-AzAuditBatch.ps1

.PARAMETER Customers
    Optional filter -- one or more HuduCompanySlugs to process. When omitted,
    all entries in UnattendedCustomers.psd1 are processed.

.PARAMETER Modules
    Global module override -- applies to every customer this run.
    When omitted, each customer uses the Modules array from the PSD1 file.

.PARAMETER ThrottleLimit
    Maximum number of concurrent customer audits. Each job runs in a
    separate pwsh process (required for Graph SDK assembly isolation).
    Default: 3. Range: 1-10.

.PARAMETER Sequential
    Disable concurrency and process customers one at a time in the current
    process. Useful for local debugging where job output is harder to inspect.

.PARAMETER HuduBaseUrl
    Hudu instance base URL. Falls back to config.psd1, then env var, then
    'https://neconnect.huducloud.com'.

.PARAMETER HuduApiKey
    Hudu API key. Falls back to Key Vault (if -KeyVaultName set), config.psd1,
    or HUDU_API_KEY env var.

.PARAMETER KeyVaultName
    Azure Key Vault name. When set, the Hudu API key is fetched from Key Vault
    using Managed Identity. Requires Az.KeyVault module and an authenticated
    Azure context (Managed Identity in Azure Functions, or Connect-AzAccount locally).

.PARAMETER HuduApiKeySecretName
    Name of the Key Vault secret containing the Hudu API key.
    Default: '365Audit-HuduApiKey'.

.PARAMETER OutputRoot
    Override the root folder where per-customer output folders are created.
    Falls back to OutputRoot in config.psd1, then defaults to two levels above the toolkit.

.PARAMETER SkipCertCheck
    Skip the certificate expiry check and renewal step.

.PARAMETER SkipSync
    Skip the automatic customer list sync from Hudu. By default, the script
    runs Helpers\Sync-UnattendedCustomers.ps1 to update UnattendedCustomers.psd1
    before processing. Use this switch if the list is already current or if you
    want to avoid the extra Hudu API calls.

.PARAMETER SkipPublish
    Skip publishing reports to Hudu after each audit completes. The audit still
    runs and generates local HTML reports, but the Hudu API publish step is
    skipped entirely. Useful for dry-run testing or when validating the Azure
    deployment without affecting production Hudu data.

.EXAMPLE
    .\Invoke-AzAuditBatch.ps1 -HuduApiKey 'key123'
    Runs all customers concurrently (3 at a time) using an explicit API key.

.EXAMPLE
    .\Invoke-AzAuditBatch.ps1 -Sequential -Customers 'contoso' -HuduApiKey 'key123'
    Runs a single customer sequentially for debugging.

.EXAMPLE
    .\Invoke-AzAuditBatch.ps1 -KeyVaultName 'kv-365audit' -HuduBaseUrl 'https://hudu.example.com'
    Azure mode: fetches Hudu API key from Key Vault, runs all customers concurrently.

.EXAMPLE
    .\Invoke-AzAuditBatch.ps1 -Customers 'contoso','fabrikam' -Modules 1,2 -ThrottleLimit 2
    Runs Entra and Exchange audits for two customers, 2 at a time.

.NOTES
    Author      : Raymond Slater
    Version     : 1.0.0

.LINK
    https://github.com/razer86/365Audit
#>

#Requires -Version 7.2

[CmdletBinding()]
param (
    [string[]]$Customers,

    [ValidateSet('1', '2', '3', '4', '5', '6', '7', 'A')]
    [string[]]$Modules,

    # ── Concurrency ──────────────────────────────────────────────────────────
    [ValidateRange(1, 10)]
    [int]$ThrottleLimit = 3,

    [switch]$Sequential,

    # ── Hudu (local mode) ────────────────────────────────────────────────────
    [string]$HuduBaseUrl,
    [string]$HuduApiKey,
    [int]$HuduAssetLayoutId,
    [int]$HuduReportLayoutId,
    [string]$HuduReportAssetName,

    # ── Azure Key Vault (Azure mode) ─────────────────────────────────────────
    [string]$KeyVaultName,
    [string]$HuduApiKeySecretName = '365Audit-HuduApiKey',

    # ── Output ───────────────────────────────────────────────────────────────
    [string]$OutputRoot,
    [switch]$CleanupLocalReports,

    # ── MSP identity ────────────────────────────────────────────────────────
    [string[]]$MspDomains,

    [switch]$SkipCertCheck,
    [switch]$SkipSync,
    [switch]$SkipPublish
)

$ScriptVersion = "1.0.0"
Write-Verbose "Invoke-AzAuditBatch.ps1 loaded (v$ScriptVersion)"

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# ── Load config.psd1 fallbacks ───────────────────────────────────────────────
$_configPath = Join-Path $PSScriptRoot 'config.psd1'
if (Test-Path $_configPath) {
    try {
        $_config = Import-PowerShellDataFile -Path $_configPath
        if (-not $HuduApiKey)          { $HuduApiKey          = $_config.HuduApiKey }
        if (-not $HuduBaseUrl)         { $HuduBaseUrl         = $_config.HuduBaseUrl }
        if (-not $OutputRoot)          { $OutputRoot          = $_config.OutputRoot }
        if ($HuduAssetLayoutId  -le 0) { $HuduAssetLayoutId   = $_config.HuduAssetLayoutId }
        if ($HuduReportLayoutId -le 0) { $HuduReportLayoutId  = $_config.HuduReportLayoutId }
        if (-not $HuduReportAssetName) { $HuduReportAssetName = $_config.HuduReportAssetName }
        if (-not $CleanupLocalReports) { $CleanupLocalReports = [bool]$_config.CleanupLocalReports }
        if (-not $MspDomains)          { $MspDomains          = @($_config.MspDomains) }
    }
    catch { Write-Warning "Could not load config.psd1: $_" }
}

# ── Environment variable fallbacks ──────────────────────────────────────────
if (-not $HuduBaseUrl         -and $env:HUDU_BASE_URL)           { $HuduBaseUrl         = $env:HUDU_BASE_URL }
if ($HuduAssetLayoutId  -le 0 -and $env:HUDU_ASSET_LAYOUT_ID)   { $HuduAssetLayoutId   = [int]$env:HUDU_ASSET_LAYOUT_ID }
if ($HuduReportLayoutId -le 0 -and $env:HUDU_REPORT_LAYOUT_ID)  { $HuduReportLayoutId  = [int]$env:HUDU_REPORT_LAYOUT_ID }
if (-not $HuduReportAssetName -and $env:HUDU_REPORT_ASSET_NAME) { $HuduReportAssetName = $env:HUDU_REPORT_ASSET_NAME }
if (-not $CleanupLocalReports -and $env:CLEANUP_LOCAL_REPORTS -eq 'true') { $CleanupLocalReports = $true }
if (-not $MspDomains          -and $env:MSP_DOMAINS)             { $MspDomains          = $env:MSP_DOMAINS -split ',' }

# ── Defaults ────────────────────────────────────────────────────────────────
if (-not $HuduBaseUrl)         { $HuduBaseUrl         = 'https://neconnect.huducloud.com' }
if ($HuduAssetLayoutId  -le 0) { $HuduAssetLayoutId   = 67 }
if ($HuduReportLayoutId -le 0) { $HuduReportLayoutId  = 68 }
if (-not $HuduReportAssetName) { $HuduReportAssetName = 'M365 - Monthly Audit Report' }

# ── Azure Key Vault credential resolution ────────────────────────────────────
if (-not $HuduApiKey -and $KeyVaultName) {
    Write-Host "  Fetching Hudu API key from Key Vault '$KeyVaultName'..." -ForegroundColor DarkCyan
    try {
        if (-not (Get-Module -ListAvailable -Name Az.KeyVault)) {
            Write-Error "Az.KeyVault module is not installed. Install it with: Install-Module Az.KeyVault" -ErrorAction Stop
        }
        Import-Module Az.KeyVault -ErrorAction Stop
        $secret = Get-AzKeyVaultSecret -VaultName $KeyVaultName -Name $HuduApiKeySecretName -AsPlainText -ErrorAction Stop
        if (-not $secret) {
            Write-Error "Key Vault secret '$HuduApiKeySecretName' not found in vault '$KeyVaultName'." -ErrorAction Stop
        }
        $HuduApiKey = $secret
        Write-Host "  Hudu API key retrieved from Key Vault." -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to retrieve Hudu API key from Key Vault: $($_.Exception.Message)" -ErrorAction Stop
    }
}

# ── Environment variable fallback ────────────────────────────────────────────
if (-not $HuduApiKey -and $env:HUDU_API_KEY) {
    $HuduApiKey = $env:HUDU_API_KEY
}

if (-not $HuduApiKey) {
    Write-Error ("Hudu API key not found. Provide one of:`n" +
        "  -HuduApiKey parameter`n" +
        "  -KeyVaultName parameter (Azure Key Vault)`n" +
        "  config.psd1 HuduApiKey value`n" +
        "  HUDU_API_KEY environment variable") -ErrorAction Stop
}

# ── Validate OutputRoot early ────────────────────────────────────────────────
if ($OutputRoot) {
    try {
        $OutputRoot  = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputRoot)
        $_qualifier  = Split-Path -Qualifier $OutputRoot -ErrorAction SilentlyContinue
        if ($_qualifier -and -not (Test-Path $_qualifier)) {
            throw "Drive or UNC root is not accessible: '$_qualifier'"
        }
        New-Item -ItemType Directory -Path $OutputRoot -Force -ErrorAction Stop | Out-Null
    }
    catch {
        throw "OutputRoot '$OutputRoot' is invalid: $($_.Exception.Message)"
    }
}

# ── Sync customer list from Hudu ─────────────────────────────────────────────
# Determine customer list path — use script root first, fall back to writable
# temp location when running in a read-only filesystem (Azure WEBSITE_RUN_FROM_PACKAGE).
$_defaultCustomerList = Join-Path $PSScriptRoot 'UnattendedCustomers.psd1'
$_tempDir = if ($env:TEMP) { $env:TEMP } else { '/tmp' }
$_writableCustomerList = Join-Path $_tempDir 'UnattendedCustomers.psd1'
$customerListPath = if (Test-Path $_defaultCustomerList) { $_defaultCustomerList }
                    elseif (Test-Path $_writableCustomerList) { $_writableCustomerList }
                    else { $_writableCustomerList }  # Sync will create it here

$_syncScript = Join-Path $PSScriptRoot 'Helpers\Sync-UnattendedCustomers.ps1'
if (-not $SkipSync -and (Test-Path $_syncScript)) {
    Write-Host "  Syncing customer list from Hudu..." -ForegroundColor DarkCyan
    try {
        & $_syncScript -HuduBaseUrl $HuduBaseUrl -HuduApiKey $HuduApiKey -HuduAssetLayoutId $HuduAssetLayoutId -OutputFilePath $customerListPath -ErrorAction Stop
    }
    catch {
        Write-Warning "Customer sync failed — continuing with existing list: $($_.Exception.Message)"
    }
}

# ── Load customer list ───────────────────────────────────────────────────────
if (-not (Test-Path $customerListPath)) {
    Write-Error ("Customer list not found: $customerListPath`n" +
        "Run Helpers\Sync-UnattendedCustomers.ps1 first, or copy " +
        "UnattendedCustomers.psd1.example and add your customers.") -ErrorAction Stop
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

if ($Customers) {
    $customerList = @($customerList | Where-Object { $_.HuduCompanySlug -in $Customers })
    if ($customerList.Count -eq 0) {
        Write-Error "None of the specified customers matched entries in UnattendedCustomers.psd1." -ErrorAction Stop
    }
}

# ── Resolve script paths ────────────────────────────────────────────────────
$scriptRoot     = $PSScriptRoot
$setupScript    = Join-Path $scriptRoot 'Setup-365AuditApp.ps1'
$auditScript    = Join-Path $scriptRoot 'Start-365Audit.ps1'
$publishScript  = Join-Path $scriptRoot 'Helpers\Publish-HuduAuditReport.ps1'

foreach ($reqScript in @($setupScript, $auditScript)) {
    if (-not (Test-Path $reqScript)) {
        Write-Error "Required script not found: $reqScript" -ErrorAction Stop
    }
}

$totalCustomers = $customerList.Count
$results        = [System.Collections.Generic.List[PSCustomObject]]::new()

# ── Batch log ────────────────────────────────────────────────────────────────
$_logDir  = if ($OutputRoot) { $OutputRoot } else { $scriptRoot }
$_logFile = Join-Path $_logDir "AzAuditBatch_$(Get-Date -Format 'yyyy-MM-dd_HHmmss').log"
function Write-BatchLog ([string]$Message) {
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')  $Message"
    Add-Content -Path $_logFile -Value $line -Encoding UTF8
}

# Force sequential for single-customer runs (no benefit from job overhead)
if ($totalCustomers -eq 1) { $Sequential = $true }

Write-Host "`n$('=' * 72)" -ForegroundColor Cyan
Write-Host "  Invoke-AzAuditBatch v$ScriptVersion — $totalCustomers customer(s)" -ForegroundColor Cyan
if (-not $Sequential) {
    Write-Host "  Concurrency: $ThrottleLimit parallel jobs" -ForegroundColor Cyan
}
Write-Host "$('=' * 72)" -ForegroundColor Cyan
Write-BatchLog "=== Batch started — $totalCustomers customer(s), ThrottleLimit=$ThrottleLimit, Sequential=$Sequential ==="

# ── Hudu publish helper (runs in parent process to avoid rate limits) ────────
function Invoke-HuduPublish {
    [CmdletBinding()]
    param (
        [string]$CustomerSlug,
        [string]$CustomerOutputPath
    )
    if (-not (Test-Path $publishScript)) { return }
    if (-not $CustomerOutputPath -or -not (Test-Path $CustomerOutputPath)) { return }

    Write-Host "  Publishing report for $CustomerSlug to Hudu..." -ForegroundColor DarkCyan
    try {
        & $publishScript `
            -OutputPath       $CustomerOutputPath `
            -CompanySlug      $CustomerSlug `
            -HuduBaseUrl      $HuduBaseUrl `
            -HuduApiKey       $HuduApiKey `
            -ReportLayoutId   $HuduReportLayoutId `
            -ReportAssetName  $HuduReportAssetName `
            -CleanupLocal:$CleanupLocalReports
        Write-BatchLog "PUBLISH  $CustomerSlug  OK"
    }
    catch {
        Write-Warning "  Hudu publish failed for ${CustomerSlug}: $($_.Exception.Message)"
        Write-BatchLog "PUBLISH  $CustomerSlug  FAIL  error=$($_.Exception.Message)"
    }
}

# ══════════════════════════════════════════════════════════════════════════════
# Sequential mode — same-process execution for debugging
# ══════════════════════════════════════════════════════════════════════════════
if ($Sequential) {
    $currentIndex = 0
    foreach ($entry in $customerList) {
        $currentIndex++
        $customerId   = $entry.HuduCompanySlug
        $customerMods = if ($Modules) { $Modules } else { @($entry.Modules ?? @('1', '2', '3', '4')) }
        $customerLabel = "[$currentIndex/$totalCustomers] $customerId"

        Write-Host "`n$('=' * 72)" -ForegroundColor Cyan
        Write-Host "  $customerLabel  (modules: $($customerMods -join ','))" -ForegroundColor Cyan
        Write-Host "$('=' * 72)" -ForegroundColor Cyan
        Write-BatchLog "START  $customerLabel  modules=$($customerMods -join ',')"
        $_tenantStart = Get-Date

        $result = [PSCustomObject]@{
            Customer    = $customerId
            Modules     = $customerMods -join ','
            AuditStatus = 'Pending'
            Error       = $null
            Elapsed     = $null
        }

        try {
            if (-not $SkipCertCheck) {
                Write-Host "  Checking certificate expiry..." -ForegroundColor DarkCyan
                & $setupScript `
                    -HuduCompanyId $customerId `
                    -HuduBaseUrl   $HuduBaseUrl `
                    -HuduApiKey    $HuduApiKey `
                    -ErrorAction Stop
            }

            Write-Host "  Starting audit (modules: $($customerMods -join ','))..." -ForegroundColor DarkCyan
            $_lastOutputFile = Join-Path $_tempDir "365Audit_LastOutput_$customerId.txt"
            $auditParams = @{
                HuduCompanyId  = $customerId
                HuduBaseUrl    = $HuduBaseUrl
                HuduApiKey     = $HuduApiKey
                Modules        = $customerMods
                LastOutputFile = $_lastOutputFile
                ErrorAction    = 'Stop'
            }
            if ($OutputRoot)  { $auditParams['OutputRoot']  = $OutputRoot }
            if ($MspDomains)  { $auditParams['MspDomains']  = $MspDomains }
            & $auditScript @auditParams

            $result.AuditStatus = 'Completed'
            $_elapsed = (Get-Date) - $_tenantStart
            $result.Elapsed = [math]::Round($_elapsed.TotalMinutes, 1)
            Write-Host "  $customerLabel — DONE ($($result.Elapsed)m)" -ForegroundColor Green
            Write-BatchLog "DONE   $customerLabel  elapsed=$($result.Elapsed)m"

            # Publish to Hudu
            if (-not $SkipPublish) {
                $_customerOutputPath = if (Test-Path $_lastOutputFile) {
                    (Get-Content $_lastOutputFile -Raw -ErrorAction SilentlyContinue).Trim()
                } else { $null }
                if ($_customerOutputPath) {
                    Invoke-HuduPublish -CustomerSlug $customerId -CustomerOutputPath $_customerOutputPath
                }
            }
        }
        catch {
            $result.AuditStatus = 'Failed'
            $result.Error       = $_.Exception.Message
            $_elapsed = (Get-Date) - $_tenantStart
            $result.Elapsed = [math]::Round($_elapsed.TotalMinutes, 1)
            Write-Host "  $customerLabel — FAILED: $($_.Exception.Message)" -ForegroundColor Red
            Write-BatchLog "FAIL   $customerLabel  elapsed=$($result.Elapsed)m  error=$($_.Exception.Message)"
        }
        finally {
            Remove-Item $_lastOutputFile -Force -ErrorAction SilentlyContinue
        }

        $results.Add($result)
    }
}
# ══════════════════════════════════════════════════════════════════════════════
# Concurrent mode — each customer runs in a separate pwsh process via Start-Job
# ══════════════════════════════════════════════════════════════════════════════
else {
    $jobScriptBlock = {
        param(
            [string]$SetupScript,
            [string]$AuditScript,
            [string]$CustomerId,
            [string[]]$CustomerMods,
            [string]$HuduBaseUrl,
            [string]$HuduApiKey,
            [string]$OutputRoot,
            [bool]$SkipCertCheck,
            [string]$LastOutputFile,
            [string[]]$MspDomains
        )

        $startTime = Get-Date
        try {
            # Step 1: Cert check/renewal (non-interactive, app-only auth)
            if (-not $SkipCertCheck) {
                & $SetupScript `
                    -HuduCompanyId $CustomerId `
                    -HuduBaseUrl   $HuduBaseUrl `
                    -HuduApiKey    $HuduApiKey `
                    -ErrorAction Stop
            }

            # Step 2: Run audit with per-customer LastOutputFile
            $auditParams = @{
                HuduCompanyId  = $CustomerId
                HuduBaseUrl    = $HuduBaseUrl
                HuduApiKey     = $HuduApiKey
                Modules        = $CustomerMods
                LastOutputFile = $LastOutputFile
                ErrorAction    = 'Stop'
            }
            if ($OutputRoot) { $auditParams['OutputRoot'] = $OutputRoot }
            if ($MspDomains) { $auditParams['MspDomains'] = $MspDomains }
            & $AuditScript @auditParams

            # Step 3: Read back output path
            $outputPath = if (Test-Path $LastOutputFile) {
                (Get-Content $LastOutputFile -Raw -ErrorAction SilentlyContinue).Trim()
            } else { $null }

            [PSCustomObject]@{
                Customer    = $CustomerId
                Modules     = $CustomerMods -join ','
                AuditStatus = 'Completed'
                Error       = $null
                Elapsed     = [math]::Round(((Get-Date) - $startTime).TotalMinutes, 1)
                OutputPath  = $outputPath
            }
        }
        catch {
            [PSCustomObject]@{
                Customer    = $CustomerId
                Modules     = $CustomerMods -join ','
                AuditStatus = 'Failed'
                Error       = $_.Exception.Message
                Elapsed     = [math]::Round(((Get-Date) - $startTime).TotalMinutes, 1)
                OutputPath  = $null
            }
        }
        finally {
            Remove-Item $LastOutputFile -Force -ErrorAction SilentlyContinue
        }
    }

    $queue      = [System.Collections.Queue]::new([array]$customerList)
    $activeJobs = [ordered]@{}

    while ($queue.Count -gt 0 -or $activeJobs.Count -gt 0) {
        # ── Launch jobs up to ThrottleLimit ───────────────────────────────────
        while ($queue.Count -gt 0 -and $activeJobs.Count -lt $ThrottleLimit) {
            $entry = $queue.Dequeue()
            $slug  = $entry.HuduCompanySlug
            $mods  = if ($Modules) { $Modules } else { @($entry.Modules ?? @('1', '2', '3', '4')) }
            $lastOutputFile = Join-Path $_tempDir "365Audit_LastOutput_$slug.txt"

            $job = Start-Job -ScriptBlock $jobScriptBlock -ArgumentList @(
                $setupScript, $auditScript,
                $slug, $mods, $HuduBaseUrl, $HuduApiKey,
                $OutputRoot, [bool]$SkipCertCheck, $lastOutputFile,
                $MspDomains
            )
            $activeJobs[$job.Id] = @{ Job = $job; Customer = $slug; Started = Get-Date }
            Write-BatchLog "START  $slug  modules=$($mods -join ',')"
            Write-Host "  Started job: $slug (modules: $($mods -join ','))" -ForegroundColor DarkCyan
        }

        # ── Check for completed jobs ─────────────────────────────────────────
        $completedIds = @($activeJobs.Keys | Where-Object { $activeJobs[$_].Job.State -in 'Completed', 'Failed' })

        foreach ($id in $completedIds) {
            $info = $activeJobs[$id]
            $slug = $info.Customer

            # Collect the result PSCustomObject emitted by the job
            $jobOutput = @(Receive-Job -Job $info.Job -ErrorAction SilentlyContinue)
            $jobResult = $jobOutput | Where-Object { $_ -is [PSCustomObject] -and $_.PSObject.Properties['AuditStatus'] } | Select-Object -Last 1

            if (-not $jobResult) {
                # Job failed without emitting a result (e.g. process crash)
                $jobError = if ($info.Job.ChildJobs -and $info.Job.ChildJobs[0].JobStateInfo.Reason) {
                    $info.Job.ChildJobs[0].JobStateInfo.Reason.Message
                } else { 'Job terminated without output' }

                $jobResult = [PSCustomObject]@{
                    Customer    = $slug
                    Modules     = ''
                    AuditStatus = 'Failed'
                    Error       = $jobError
                    Elapsed     = [math]::Round(((Get-Date) - $info.Started).TotalMinutes, 1)
                    OutputPath  = $null
                }
            }

            $results.Add($jobResult)
            Remove-Job -Job $info.Job -Force
            $activeJobs.Remove($id)

            if ($jobResult.AuditStatus -eq 'Completed') {
                Write-Host "  $slug — DONE ($($jobResult.Elapsed)m)" -ForegroundColor Green
                Write-BatchLog "DONE   $slug  elapsed=$($jobResult.Elapsed)m"

                # Publish to Hudu sequentially from parent process
                if (-not $SkipPublish -and $jobResult.OutputPath) {
                    Invoke-HuduPublish -CustomerSlug $slug -CustomerOutputPath $jobResult.OutputPath
                }
            }
            else {
                Write-Host "  $slug — FAILED: $($jobResult.Error)" -ForegroundColor Red
                Write-BatchLog "FAIL   $slug  elapsed=$($jobResult.Elapsed)m  error=$($jobResult.Error)"
            }
        }

        # ── Progress display ─────────────────────────────────────────────────
        if ($activeJobs.Count -gt 0) {
            $running = ($activeJobs.Values | ForEach-Object {
                "$($_.Customer) ($([math]::Round(((Get-Date) - $_.Started).TotalMinutes, 1))m)"
            }) -join ', '
            Write-Host "`r  Running: $running  |  Queued: $($queue.Count)  |  Done: $($results.Count)/$totalCustomers    " -ForegroundColor DarkGray -NoNewline
            Start-Sleep -Seconds 5
        }
    }
    Write-Host ""  # Clear the progress line
}

# ── Final summary ────────────────────────────────────────────────────────────
Write-Host "`n$('=' * 72)" -ForegroundColor Cyan
Write-Host "  Batch complete — $($results.Count) customer(s)" -ForegroundColor Cyan
Write-Host "$('=' * 72)" -ForegroundColor Cyan

$results | Format-Table -AutoSize

$_completedCount = @($results | Where-Object { $_.AuditStatus -eq 'Completed' }).Count
$_failedCount    = @($results | Where-Object { $_.AuditStatus -eq 'Failed'    }).Count
Write-BatchLog "=== Batch finished — $_completedCount completed, $_failedCount failed ==="
$results | Format-Table -AutoSize | Out-String -Width 120 | ForEach-Object { $_.TrimEnd() } |
    Where-Object { $_ } | Add-Content -Path $_logFile -Encoding UTF8
Write-Host "  Batch log: $_logFile" -ForegroundColor DarkGray

$failed = @($results | Where-Object { $_.AuditStatus -eq 'Failed' })
if ($failed.Count -gt 0) {
    Write-Host "  $($failed.Count) customer(s) failed — review errors above." -ForegroundColor Red
    exit 1
}
