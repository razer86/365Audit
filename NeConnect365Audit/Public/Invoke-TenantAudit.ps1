function Invoke-TenantAudit {
    <#
    .SYNOPSIS
        Runs the NeConnect 365Audit toolkit against one or more customer tenants.

    .DESCRIPTION
        Single entry point for all audit operations. Supports single-customer,
        filtered batch, and full batch modes. For each customer:
          1. Fetches credentials from Hudu (or Key Vault in Azure mode)
          2. Connects to Graph, Exchange, Teams, and Security & Compliance
          3. Runs selected audit modules (Entra, Exchange, SharePoint, Mail Security, Intune, Teams, Maester)
          4. Generates HTML summary report
          5. Publishes report to Hudu (unless -SkipPublish)
          6. Disconnects all sessions and cleans up temp cert

    .PARAMETER Customers
        One or more Hudu company slugs to audit. When omitted, all customers
        are synced from Hudu and audited.

    .PARAMETER HuduBaseUrl
        Hudu instance base URL (no trailing slash).

    .PARAMETER HuduApiKey
        Hudu API key. Alternatively, use -KeyVaultName for Azure Automation.

    .PARAMETER KeyVaultName
        Azure Key Vault name containing the Hudu API key secret (365Audit-HuduApiKey).
        Used for Azure Automation runs with managed identity.

    .PARAMETER Modules
        Select specific audit modules to run. Default: all modules.
        Valid values: 1 (Entra), 2 (Exchange), 3 (SharePoint), 4 (Mail Security),
        5 (Intune), 6 (Teams), 7 (Maester CIS Baseline), A (All).

    .PARAMETER OutputRoot
        Root folder for per-customer output. Defaults to $env:TEMP.

    .PARAMETER SkipPublish
        Don't publish reports to Hudu after audit completes.

    .PARAMETER SkipCertCheck
        Skip certificate expiry check/renewal step.

    .PARAMETER MspDomains
        MSP email domains for report filtering (flags non-MSP technical contacts).

    .PARAMETER CleanupLocal
        Delete local report folder after successful Hudu upload.

    .PARAMETER HuduAssetLayoutId
        Hudu asset layout ID for credential assets. Default: 67.

    .PARAMETER HuduReportLayoutId
        Hudu asset layout ID for monthly report assets. Default: 68.

    .PARAMETER HuduReportAssetName
        Display name prefix for monthly report assets.

    .EXAMPLE
        Invoke-TenantAudit -Customers '03e7d5ad117e' -HuduApiKey $key -HuduBaseUrl $url -SkipPublish

    .EXAMPLE
        Invoke-TenantAudit -HuduApiKey $key -HuduBaseUrl $url
    #>
    [CmdletBinding()]
    param(
        [string[]]$Customers,

        [ValidateSet('1', '2', '3', '4', '5', '6', '7', 'A')]
        [string[]]$Modules = @('A'),

        [string]$HuduBaseUrl,
        [string]$HuduApiKey,
        [string]$KeyVaultName,

        [string]$OutputRoot,

        [switch]$SkipPublish,
        [switch]$SkipCertCheck,
        [switch]$CleanupLocal,

        [string[]]$MspDomains = @(),

        [int]$HuduAssetLayoutId = 67,
        [int]$HuduReportLayoutId = 68,
        [string]$HuduReportAssetName = 'M365 - Monthly Audit Report'
    )

    $ErrorActionPreference = 'Stop'

    # ── Resolve Hudu API key ────────────────────────────────────────────────
    if (-not $HuduApiKey -and $KeyVaultName) {
        Write-Host "Fetching Hudu API key from Key Vault '$KeyVaultName'..." -ForegroundColor DarkCyan
        Import-Module Az.KeyVault -ErrorAction Stop
        $HuduApiKey = Get-AzKeyVaultSecret -VaultName $KeyVaultName -Name '365Audit-HuduApiKey' -AsPlainText -ErrorAction Stop
        if (-not $HuduApiKey) {
            throw "Key Vault secret '365Audit-HuduApiKey' not found in vault '$KeyVaultName'."
        }
        Write-Host "Hudu API key retrieved." -ForegroundColor Green
    }

    if (-not $HuduApiKey) {
        throw "Hudu API key is required. Pass -HuduApiKey or -KeyVaultName."
    }
    if (-not $HuduBaseUrl) {
        throw "Hudu base URL is required. Pass -HuduBaseUrl."
    }

    # ── Resolve output root ─────────────────────────────────────────────────
    if (-not $OutputRoot) {
        $OutputRoot = Join-Path ($env:TEMP ?? $env:TMPDIR ?? '/tmp') '365audit'
    }
    New-Item -ItemType Directory -Path $OutputRoot -Force -ErrorAction Stop | Out-Null

    # ── Resolve collector list ──────────────────────────────────────────────
    $_collectorMap = @{
        '1' = @{ Name = 'Entra';         Function = 'Get-EntraInventory' }
        '2' = @{ Name = 'Exchange';       Function = 'Get-ExchangeInventory' }
        '3' = @{ Name = 'SharePoint';     Function = 'Get-SharePointInventory' }
        '4' = @{ Name = 'Mail Security';  Function = 'Get-MailSecurityInventory' }
        '5' = @{ Name = 'Intune';         Function = 'Get-IntuneInventory' }
        '6' = @{ Name = 'Teams';          Function = 'Get-TeamsInventory' }
    }

    if ($Modules -contains 'A') {
        $_selectedCollectors = @('1', '2', '3', '4', '5', '6')
        $_runMaester = $true
    }
    elseif ($Modules -contains '7') {
        $_selectedCollectors = @($Modules | Where-Object { $_ -ne '7' })
        $_runMaester = $true
    }
    else {
        $_selectedCollectors = $Modules
        $_runMaester = $false
    }

    # ── Sync customer list from Hudu ────────────────────────────────────────
    if (-not $Customers) {
        Write-Host "Syncing customer list from Hudu..." -ForegroundColor DarkCyan
        $_allCustomers = Sync-AuditCustomers -HuduBaseUrl $HuduBaseUrl -HuduApiKey $HuduApiKey `
            -HuduAssetLayoutId $HuduAssetLayoutId -PassThru
        $Customers = @($_allCustomers.HuduCompanySlug)
        Write-Host "  Found $($Customers.Count) customer(s)." -ForegroundColor Green
    }

    if ($Customers.Count -eq 0) {
        throw "No customers to audit."
    }

    # ── Batch setup ─────────────────────────────────────────────────────────
    $_logFile = Join-Path $OutputRoot "AuditBatch_$(Get-Date -Format 'yyyy-MM-dd_HHmmss').log"
    $_results = [System.Collections.Generic.List[PSCustomObject]]::new()
    $_totalCustomers = $Customers.Count
    $_currentIndex   = 0

    $_collectorNames = ($_selectedCollectors | ForEach-Object { $_collectorMap[$_].Name }) -join ', '
    $_maesterLabel   = if ($_runMaester) { ' + Maester CIS Baseline' } else { '' }

    Write-Host "`n$('=' * 72)" -ForegroundColor Cyan
    Write-Host "  NeConnect365Audit — $_totalCustomers customer(s)" -ForegroundColor Cyan
    Write-Host "  Collectors: ${_collectorNames}${_maesterLabel}" -ForegroundColor Cyan
    Write-Host "$('=' * 72)" -ForegroundColor Cyan
    Write-AuditLog -Message "=== Batch started — $_totalCustomers customer(s) ===" -LogFile $_logFile

    # ── Process each customer ───────────────────────────────────────────────
    foreach ($_slug in $Customers) {
        $_currentIndex++
        $_label = "[$_currentIndex/$_totalCustomers] $_slug"
        $_tenantStart = Get-Date

        $_result = [PSCustomObject]@{
            Customer    = $_slug
            AuditStatus = 'Pending'
            Error       = $null
            Elapsed     = $null
        }

        Write-Host "`n$('=' * 72)" -ForegroundColor Cyan
        Write-Host "  $_label" -ForegroundColor Cyan
        Write-Host "$('=' * 72)" -ForegroundColor Cyan
        Write-AuditLog -Message "START  $_label" -LogFile $_logFile

        try {
            # ── Fetch credentials from Hudu ─────────────────────────────────
            Write-Host "  Fetching credentials from Hudu..." -ForegroundColor DarkCyan
            $_creds = Resolve-HuduCredentials -CompanySlug $_slug `
                -HuduBaseUrl $HuduBaseUrl -HuduApiKey $HuduApiKey -AssetLayoutId $HuduAssetLayoutId
            Write-Host "  Credentials loaded: $($_creds.AssetName)" -ForegroundColor Green

            # ── Set audit context for this customer ─────────────────────────
            Set-AuditContext `
                -AppId               $_creds.AppId `
                -TenantId            $_creds.TenantId `
                -CertFilePath        $_creds.CertFilePath `
                -CertPassword        $_creds.CertPassword `
                -HuduBaseUrl         $HuduBaseUrl `
                -HuduApiKey          $HuduApiKey `
                -HuduCompanySlug     $_slug `
                -HuduAssetLayoutId   $HuduAssetLayoutId `
                -HuduReportLayoutId  $HuduReportLayoutId `
                -HuduReportAssetName $HuduReportAssetName `
                -MspDomains          $MspDomains

            # ── Connect to services ─────────────────────────────────────────
            Connect-AuditGraph
            $_ctx = Initialize-AuditOutput -OutputRoot $OutputRoot
            Write-Host "  Company: $($_ctx.OrgName)" -ForegroundColor Green
            Write-Host "  Output:  $($_ctx.OutputPath)" -ForegroundColor Gray

            # Connect EXO for exchange/mail security collectors and Maester
            if ($_selectedCollectors | Where-Object { $_ -in @('2', '4') }) {
                Connect-AuditExchange
            }

            # ── DATA COLLECTION — inventory CSVs ────────────────────────────
            $_collectorErrors = @()
            foreach ($_colKey in $_selectedCollectors) {
                $_col = $_collectorMap[$_colKey]
                if (-not $_col) { continue }

                Write-Host "`n  ── $($_col.Name) Inventory ──" -ForegroundColor Cyan
                try {
                    & $_col.Function
                    Write-Host "  Completed: $($_col.Name)" -ForegroundColor Green
                }
                catch {
                    Write-Warning "  Collector FAILED: $($_col.Name) — $($_.Exception.Message)"
                    $_collectorErrors += $_col.Name
                }
            }

            # ── SECURITY ASSESSMENT — Maester ───────────────────────────────
            if ($_runMaester) {
                Write-Host "`n  ── Maester CIS Baseline ──" -ForegroundColor Cyan
                try {
                    # Ensure all service connections for Maester
                    $_exoConnected = Get-ConnectionInformation -ErrorAction SilentlyContinue |
                        Where-Object { $_.State -eq 'Connected' }
                    if (-not $_exoConnected) { Connect-AuditExchange }

                    try { Connect-AuditTeams } catch {
                        Write-Warning "  Teams connection failed — Maester Teams tests will be skipped."
                    }
                    Connect-AuditCompliance

                    # Install Maester tests + copy our custom tests alongside
                    $_tempBase       = $env:TEMP ?? $env:TMPDIR ?? '/tmp'
                    $_maesterTestDir = Join-Path $_tempBase "MaesterTests_$(New-Guid)"
                    Install-MaesterTests -Path $_maesterTestDir -ErrorAction Stop

                    # Copy custom MSP tests from module
                    $_customTestsDir = Join-Path $PSScriptRoot '..' 'Tests'
                    if (Test-Path $_customTestsDir) {
                        Get-ChildItem -Path $_customTestsDir -Filter '*.Tests.ps1' | ForEach-Object {
                            Copy-Item $_.FullName -Destination $_maesterTestDir -Force
                        }
                    }

                    # Run Maester
                    $_maesterDir = Join-Path $_ctx.RawOutputPath 'Maester'
                    New-Item -ItemType Directory -Path $_maesterDir -Force | Out-Null

                    $_maesterResult = Invoke-Maester `
                        -Path           $_maesterTestDir `
                        -NonInteractive `
                        -OutputJsonFile (Join-Path $_maesterDir 'MaesterResults.json') `
                        -OutputHtmlFile (Join-Path $_maesterDir 'MaesterReport.html') `
                        -PassThru `
                        -Verbosity None

                    if ($_maesterResult) {
                        Write-Host "  Maester: $($_maesterResult.TotalCount) tests, $($_maesterResult.PassedCount) passed, $($_maesterResult.FailedCount) failed." -ForegroundColor Green
                    }

                    # Clean up temp test files
                    Remove-Item $_maesterTestDir -Recurse -Force -ErrorAction SilentlyContinue
                }
                catch {
                    Write-Warning "  Maester FAILED: $($_.Exception.Message)"
                    $_collectorErrors += 'Maester'
                }
            }

            if ($_collectorErrors.Count -gt 0) {
                Write-Warning "$($_collectorErrors.Count) component(s) failed: $($_collectorErrors -join ', ')"
            }

            # ── REPORT — inventory CSVs + Maester JSON → HTML ───────────────
            Write-Host "`n  Generating summary report..." -ForegroundColor DarkCyan
            New-AuditSummary -AuditFolder $_ctx.OutputPath -NoOpen

            # ── PUBLISH — push to Hudu ──────────────────────────────────────
            if (-not $SkipPublish) {
                Write-Host "  Publishing report to Hudu..." -ForegroundColor DarkCyan
                try {
                    Publish-AuditReport `
                        -OutputPath       $_ctx.OutputPath `
                        -CompanySlug      $_slug `
                        -HuduBaseUrl      $HuduBaseUrl `
                        -HuduApiKey       $HuduApiKey `
                        -ReportLayoutId   $HuduReportLayoutId `
                        -ReportAssetName  $HuduReportAssetName `
                        -CleanupLocal:$CleanupLocal
                }
                catch {
                    Write-Warning "  Hudu publish failed: $($_.Exception.Message)"
                }
            }

            $_result.AuditStatus = if ($_collectorErrors.Count -gt 0) { 'Partial' } else { 'Completed' }
            $_elapsed = (Get-Date) - $_tenantStart
            $_result.Elapsed = [math]::Round($_elapsed.TotalMinutes, 1)
            Write-Host "  $_label — DONE ($($_result.Elapsed)m)" -ForegroundColor Green
            Write-AuditLog -Message "DONE   $_label  elapsed=$($_result.Elapsed)m" -LogFile $_logFile
        }
        catch {
            $_result.AuditStatus = 'Failed'
            $_result.Error = $_.Exception.Message
            $_elapsed = (Get-Date) - $_tenantStart
            $_result.Elapsed = [math]::Round($_elapsed.TotalMinutes, 1)
            Write-Host "  $_label — FAILED: $($_.Exception.Message)" -ForegroundColor Red
            Write-AuditLog -Message "FAIL   $_label  error=$($_.Exception.Message)" -LogFile $_logFile
        }
        finally {
            # ── Disconnect all sessions ─────────────────────────────────────
            if (Get-MgContext -ErrorAction SilentlyContinue) {
                Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
            }
            if (Get-ConnectionInformation -ErrorAction SilentlyContinue | Where-Object { $_.State -eq 'Connected' }) {
                Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
            }
            if (Get-Module -Name MicrosoftTeams -ErrorAction SilentlyContinue) {
                Disconnect-MicrosoftTeams -ErrorAction SilentlyContinue | Out-Null
            }

            # Clean up temp cert file
            $_certPath = (Get-AuditContext -NoThrow)?.CertFilePath
            if ($_certPath -and (Test-Path $_certPath)) {
                Remove-Item $_certPath -Force -ErrorAction SilentlyContinue
            }

            # Clear context for next customer
            $script:AuditContext = $null
        }

        $_results.Add($_result)
    }

    # ── Final summary ───────────────────────────────────────────────────────
    Write-Host "`n$('=' * 72)" -ForegroundColor Cyan
    Write-Host "  Batch complete — $($_results.Count) customer(s)" -ForegroundColor Cyan
    Write-Host "$('=' * 72)" -ForegroundColor Cyan

    $_results | Format-Table -AutoSize

    $_completedCount = @($_results | Where-Object { $_.AuditStatus -in @('Completed', 'Partial') }).Count
    $_failedCount    = @($_results | Where-Object { $_.AuditStatus -eq 'Failed' }).Count
    Write-AuditLog -Message "=== Batch finished — $_completedCount completed, $_failedCount failed ===" -LogFile $_logFile
    Write-Host "  Batch log: $_logFile" -ForegroundColor DarkGray

    if ($_failedCount -gt 0) {
        Write-Warning "$_failedCount customer(s) failed — review output above."
    }

    return $_results
}
