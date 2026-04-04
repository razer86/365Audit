<#
.SYNOPSIS
    Runs the Maester M365 security configuration assessment against the current tenant.

.DESCRIPTION
    Maester (https://maester.dev) is a PowerShell 7-native security testing framework
    that evaluates Microsoft 365 tenants against CISA SCuBA, CIS v5.0, and EIDSCA
    baselines using Pester-based tests.

    Unlike ScubaGear (which requires Windows PowerShell 5.1), Maester runs natively
    in PowerShell 7 on Windows, Linux, and macOS — making it suitable for Docker
    containers and Azure Container Apps Jobs.

    This module reuses the existing Graph and Exchange connections established by
    Start-365Audit.ps1 (via Connect-MgGraphSecure / Connect-ExchangeOnlineSecure).

    Output is written to:
        <AuditFolder>\Raw\Maester\
            MaesterResults.json     — consolidated JSON (ingested by Generate-AuditSummary.ps1)
            MaesterResults.csv      — flat CSV of all test results
            MaesterReport.html      — Maester's own HTML report

    Generate-AuditSummary.ps1 detects the Maester\ folder and adds failing/warning
    controls to the action items list and a CIS Baseline section to the report.

.NOTES
    Author      : Raymond Slater
    Version     : 1.0.0
    Change Log  : See CHANGELOG.md

    Prerequisites (one-time, per app registration):
      - Microsoft Graph application permissions:
        Directory.Read.All, Policy.Read.All, RoleManagement.Read.All,
        UserAuthenticationMethod.Read.All, Organization.Read.All,
        Policy.ReadWrite.ConditionalAccess, SecurityEvents.Read.All,
        Reports.Read.All, IdentityRiskyUser.Read.All, CrossTenantInformation.ReadBasic.All,
        SharePointTenantSettings.Read.All, PrivilegedEligibilitySchedule.Read.AzureADGroup,
        RoleEligibilitySchedule.Read.Directory
      - Exchange Online: Exchange.ManageAsApp application permission
        (the service principal must be a member of the Exchange Administrator role)

.LINK
    https://github.com/razer86/365Audit
    https://maester.dev
#>

#Requires -Version 7.2

param (
    [switch]$DevMode = $false
)

if (-not $DevMode -and $MyInvocation.InvocationName -eq $MyInvocation.MyCommand.Name) {
    Write-Error "This script must be run from the 365Audit launcher. Use -DevMode for development." -ErrorAction Stop
}

$ScriptVersion = "1.0.0"
Write-Verbose "Invoke-MaesterAudit.ps1 loaded (v$ScriptVersion)"

# ── Ensure Maester module is available ──────────────────────────────────────
Write-Progress -Id 1 -Activity 'Maester CIS Baseline' -Status 'Checking Maester module...' -PercentComplete 5

if (-not (Get-Module -ListAvailable -Name Maester)) {
    Write-Host "  Required module 'Maester' not found — installing..." -ForegroundColor Yellow
    Install-Module Maester -Scope CurrentUser -Force -AllowClobber -SkipPublisherCheck -ErrorAction Stop
    $_installedMod = Get-Module -ListAvailable -Name Maester | Sort-Object Version -Descending | Select-Object -First 1
    if (-not $_installedMod) {
        throw "Installation of 'Maester' failed — module still not found after install."
    }
    Write-Host "  Installed 'Maester' v$($_installedMod.Version)." -ForegroundColor Green
}

Import-Module Maester -ErrorAction Stop -WarningAction SilentlyContinue
$_maesterVersion = (Get-Module Maester).Version.ToString()
Write-Host "  Maester v$_maesterVersion loaded." -ForegroundColor Gray

# ── Get audit context ───────────────────────────────────────────────────────
$_ctx = Initialize-AuditOutput
if (-not $_ctx) {
    Write-Error "Could not initialise audit output directory." -ErrorAction Stop
}

$_rawOutPath = $_ctx.RawOutputPath
$_maesterDir = Join-Path $_rawOutPath 'Maester'
New-Item -ItemType Directory -Path $_maesterDir -Force | Out-Null

# ── Ensure service connections are active ───────────────────────────────────
# Maester detects EXO/Teams connections independently. If earlier modules already
# connected, these are no-ops. If running module 7 standalone, this establishes them.
Write-Progress -Id 1 -Activity 'Maester CIS Baseline' -Status 'Ensuring service connections...' -PercentComplete 10

# Graph should already be connected from the launcher, but verify
if (-not (Get-MgContext)) {
    Connect-MgGraphSecure
}

# Exchange Online — Maester skips 130+ tests without this
$_exoConnected = Get-ConnectionInformation -ErrorAction SilentlyContinue | Where-Object { $_.State -eq 'Connected' }
if (-not $_exoConnected) {
    Write-Host "  Connecting to Exchange Online for Maester..." -ForegroundColor Gray
    Connect-ExchangeOnlineSecure
}

# Teams — Maester skips 11+ tests without this
try {
    Connect-TeamsSecure
} catch {
    Write-Warning "Could not connect to Teams — Maester Teams tests will be skipped: $($_.Exception.Message)"
}

# ── Run Maester ─────────────────────────────────────────────────────────────
# Maester uses the active Graph, EXO, and Teams connections.
Write-Progress -Id 1 -Activity 'Maester CIS Baseline' -Status 'Installing Maester tests...' -PercentComplete 15

# Install Maester test files to a temp directory — the module doesn't bundle them
$_maesterTestDir = Join-Path $_maesterDir 'tests'
Write-Host "  Installing Maester test files..." -ForegroundColor Gray
Install-MaesterTests -Path $_maesterTestDir -ErrorAction Stop

Write-Progress -Id 1 -Activity 'Maester CIS Baseline' -Status 'Running Maester assessment...' -PercentComplete 20
Write-Host "  Running Maester security assessment (this may take several minutes)..." -ForegroundColor Cyan

$_maesterJsonPath  = Join-Path $_maesterDir 'MaesterResults.json'
$_maesterHtmlPath  = Join-Path $_maesterDir 'MaesterReport.html'
$_maesterCsvPath   = Join-Path $_maesterDir 'MaesterResults.csv'

try {
    # Invoke-Maester runs Pester tests from the installed test directory
    # -NonInteractive suppresses browser-based prompts
    # -OutputJsonFile / -OutputHtmlFile capture results
    $_maesterParams = @{
        Path           = $_maesterTestDir
        NonInteractive = $true
        OutputJsonFile = $_maesterJsonPath
        OutputHtmlFile = $_maesterHtmlPath
        PassThru       = $true
        Verbosity      = 'None'
    }

    $_pesterResult = Invoke-Maester @_maesterParams

    if (-not $_pesterResult) {
        Write-Warning "Maester returned no results."
        Add-AuditIssue -Severity 'Warning' -Section 'CIS Baseline' -Collector 'Maester' `
            -Description 'Maester assessment returned no results.' `
            -Action 'Check Maester module version and Graph permissions.'
        return
    }

    Write-Host "  Maester assessment complete — $($_pesterResult.TotalCount) tests, $($_pesterResult.PassedCount) passed, $($_pesterResult.FailedCount) failed." -ForegroundColor Green
}
catch {
    Add-AuditIssue -Severity 'Warning' -Section 'CIS Baseline' -Collector 'Maester' `
        -Description "Maester assessment failed: $($_.Exception.Message)" `
        -Action 'Check Maester prerequisites: https://maester.dev/docs/installation'
    Write-Warning "Maester assessment failed: $($_.Exception.Message)"
    return
}

# ── Transform results to ScubaGear-compatible JSON ──────────────────────────
# Generate-AuditSummary.ps1 expects the ScubaGear JSON structure:
#   MetaData: { ToolVersion }
#   Summary:  { <Product>: { Passes, Failures, Warnings, Manual } }
#   Results:  { <Product>: [ { GroupName, GroupReferenceURL, Controls: [ { Control ID, Requirement, Result, Criticality, Details } ] } ] }
#
# We transform Maester's Pester output into this structure so the existing
# report generation works with minimal changes.
Write-Progress -Id 1 -Activity 'Maester CIS Baseline' -Status 'Processing results...' -PercentComplete 85

# ── Build CSV export ────────────────────────────────────────────────────────
$_csvRows = foreach ($_test in $_pesterResult.Tests) {
    [PSCustomObject]@{
        Name       = $_test.Name
        Result     = $_test.Result
        Block      = ($_test.Block -join ' / ')
        Duration   = $_test.Duration.TotalSeconds
        ErrorMessage = if ($_test.ErrorRecord) { $_test.ErrorRecord.Exception.Message } else { '' }
    }
}
if ($_csvRows) {
    $_csvRows | Export-Csv -Path $_maesterCsvPath -NoTypeInformation -Encoding UTF8
}

# ── Build ScubaGear-compatible JSON ─────────────────────────────────────────
# Group tests by their top-level Block (CISA, CIS, EIDSCA, etc.)
$_productMap = @{
    'CISA'   = 'CISA SCuBA'
    'CIS'    = 'CIS v5.0'
    'EIDSCA' = 'EIDSCA'
    'Maester' = 'Maester'
}

$_resultsByProduct = @{}
$_summaryByProduct = @{}

foreach ($_test in $_pesterResult.Tests) {
    # Determine product category from the test block path
    $_blocks = @($_test.Block)
    $_topBlock = if ($_blocks.Count -gt 0) { $_blocks[0] } else { 'Other' }

    # Map to product name
    $_product = 'Other'
    foreach ($_key in $_productMap.Keys) {
        if ($_topBlock -match $_key) {
            $_product = $_productMap[$_key]
            break
        }
    }

    # Determine group (second-level block)
    $_groupName = if ($_blocks.Count -gt 1) { $_blocks[1] } else { $_topBlock }

    # Map Pester result to ScubaGear-style result
    $_result = switch ($_test.Result) {
        'Passed'   { 'Pass' }
        'Failed'   { 'Fail' }
        'Skipped'  { 'N/A' }
        'Inconclusive' { 'Warning' }
        default    { $_test.Result }
    }

    # Build control entry
    $_controlEntry = @{
        'Control ID'  = $_test.Name -replace '^(MT\.\d+\.\d+\.\d+|MS\.\w+\.\d+\.\d+v\d+).*', '$1'
        'Requirement' = $_test.Name
        'Result'      = $_result
        'Criticality' = if ($_result -eq 'Fail') { 'Shall' } else { 'Should' }
        'Details'     = if ($_test.ErrorRecord) { $_test.ErrorRecord.Exception.Message } else { '' }
    }

    # Group controls by product and group
    if (-not $_resultsByProduct.ContainsKey($_product)) {
        $_resultsByProduct[$_product] = @{}
    }
    if (-not $_resultsByProduct[$_product].ContainsKey($_groupName)) {
        $_resultsByProduct[$_product][$_groupName] = @{
            GroupName         = $_groupName
            GroupReferenceURL = ''
            Controls          = [System.Collections.Generic.List[object]]::new()
        }
    }
    $_resultsByProduct[$_product][$_groupName].Controls.Add($_controlEntry)

    # Update summary counts
    if (-not $_summaryByProduct.ContainsKey($_product)) {
        $_summaryByProduct[$_product] = @{ Passes = 0; Failures = 0; Warnings = 0; Manual = 0 }
    }
    switch ($_result) {
        'Pass'    { $_summaryByProduct[$_product].Passes++ }
        'Fail'    { $_summaryByProduct[$_product].Failures++ }
        'Warning' { $_summaryByProduct[$_product].Warnings++ }
        default   { $_summaryByProduct[$_product].Manual++ }
    }
}

# Assemble final JSON structure
$_resultsObj = @{}
foreach ($_prod in $_resultsByProduct.Keys) {
    $_resultsObj[$_prod] = @($_resultsByProduct[$_prod].Values)
}

$_compatJson = @{
    MetaData = @{
        ToolVersion   = "Maester $_maesterVersion"
        TenantId      = (Get-MgContext).TenantId
        Timestamp     = (Get-Date -Format 'o')
        TotalTests    = $_pesterResult.TotalCount
        PassedTests   = $_pesterResult.PassedCount
        FailedTests   = $_pesterResult.FailedCount
        SkippedTests  = $_pesterResult.SkippedCount
    }
    Summary = $_summaryByProduct
    Results = $_resultsObj
}

# Write ScubaGear-compatible JSON to Maester directory
# Use the ScubaResults_*.json naming convention so Generate-AuditSummary.ps1 finds it
$_compatJsonPath = Join-Path $_maesterDir "ScubaResults_Maester.json"
$_compatJson | ConvertTo-Json -Depth 10 | Set-Content -Path $_compatJsonPath -Encoding UTF8

Write-Host "  Results: $_maesterJsonPath" -ForegroundColor Gray
Write-Host "  CSV:     $_maesterCsvPath" -ForegroundColor Gray
Write-Host "  Report:  $_maesterHtmlPath" -ForegroundColor Gray

Write-Progress -Id 1 -Activity 'Maester CIS Baseline' -Completed
