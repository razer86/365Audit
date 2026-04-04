<#
.SYNOPSIS
    Generates or updates UnattendedCustomers.psd1 from Hudu asset data.

.DESCRIPTION
    Queries Hudu for all companies that have an asset matching the configured
    asset layout (HuduAssetLayoutId in config.psd1), then merges the results
    into UnattendedCustomers.psd1:

      - Companies already in the file are left untouched (their Modules config is preserved)
      - New companies are appended with the DefaultModules value
      - Companies in the file but no longer found in Hudu are commented out in-place
        so they are skipped at runtime but remain visible for manual review

    Reads HuduBaseUrl, HuduApiKey, and HuduAssetLayoutId from config.psd1.

.PARAMETER DefaultModules
    Module list assigned to newly discovered customers.
    Valid values: 1=Entra  2=Exchange  3=SharePoint  4=MailSec  5=Intune  6=Teams  7=Maester  A=All
    Defaults to @('A') (Run All).

.PARAMETER WhatIf
    Show what would be added/updated without writing the file.

.EXAMPLE
    .\Helpers\Sync-UnattendedCustomers.ps1

.EXAMPLE
    .\Helpers\Sync-UnattendedCustomers.ps1 -DefaultModules '1','2','4'

.NOTES
    Author      : Raymond Slater
    Version     : 1.0.0
#>

#Requires -Version 7.2

[CmdletBinding(SupportsShouldProcess)]
param (
    [ValidateSet('1', '2', '3', '4', '5', '6', '7', 'A')]
    [string[]]$DefaultModules = @('A')
)

$ScriptVersion         = "1.0.0"
Write-Verbose "Sync-UnattendedCustomers.ps1 loaded (v$ScriptVersion)"

$ErrorActionPreference = 'Stop'

# ── Load config ────────────────────────────────────────────────────────────────

$_configPath = Join-Path $PSScriptRoot '..\config.psd1'
if (-not (Test-Path $_configPath)) {
    Write-Error "config.psd1 not found at $_configPath. Copy config.psd1.example and fill in your values."
}

try   { $_config = Import-PowerShellDataFile -Path $_configPath }
catch { Write-Error "Could not load config.psd1: $_" }

$huduBaseUrl  = $_config.HuduBaseUrl?.TrimEnd('/')
$huduApiKey   = $_config.HuduApiKey
$layoutId     = if ($_config.HuduAssetLayoutId -gt 0) { $_config.HuduAssetLayoutId } else { 67 }

if (-not $huduBaseUrl) { Write-Error "HuduBaseUrl is not set in config.psd1." }
if (-not $huduApiKey)  { Write-Error "HuduApiKey is not set in config.psd1." }

$headers = @{ 'x-api-key' = $huduApiKey }

# ── Load existing UnattendedCustomers.psd1 ────────────────────────────────────

$outputPath = Join-Path $PSScriptRoot '..\UnattendedCustomers.psd1'
$existing   = @{}   # slug → existing entry hashtable

if (Test-Path $outputPath) {
    try {
        $_existing = Import-PowerShellDataFile -Path $outputPath
        foreach ($entry in @($_existing.Customers)) {
            $existing[$entry.HuduCompanySlug] = $entry
        }
        Write-Host "Loaded $($existing.Count) existing customer(s) from UnattendedCustomers.psd1." -ForegroundColor DarkGray
    }
    catch { Write-Warning "Could not read existing UnattendedCustomers.psd1 — will treat as empty: $_" }
}

# ── Fetch all assets for the layout from Hudu (paginated) ─────────────────────

Write-Host "Querying Hudu for assets with layout ID $layoutId..." -ForegroundColor Cyan

$allAssets = [System.Collections.Generic.List[object]]::new()
$page      = 1
do {
    try {
        $response = Invoke-RestMethod `
            -Uri     "$huduBaseUrl/api/v1/assets?asset_layout_id=$layoutId&page_size=25&page=$page" `
            -Headers $headers -Method Get -ErrorAction Stop
    }
    catch { Write-Error "Hudu asset query failed (page $page): $_" }
    foreach ($a in @($response.assets)) { $allAssets.Add($a) }
    $page++
} while ($response.assets.Count -gt 0)

Write-Host "Found $($allAssets.Count) asset(s) across all companies." -ForegroundColor Cyan

# ── Resolve company slug for each unique company_id ───────────────────────────

Write-Host "Resolving company slugs..." -ForegroundColor Cyan

$companyMap = @{}   # company_id → { slug, name }
foreach ($asset in $allAssets) {
    $cid = $asset.company_id
    if (-not $cid -or $companyMap.ContainsKey($cid)) { continue }
    try {
        $company = Invoke-RestMethod -Uri "$huduBaseUrl/api/v1/companies/$cid" `
            -Headers $headers -Method Get -ErrorAction Stop
        $companyMap[$cid] = @{
            Slug = $company.company.slug
            Name = $company.company.name
        }
    }
    catch { Write-Warning "Could not resolve company ID $cid — skipping: $_" }
}

Write-Host "Resolved $($companyMap.Count) company slug(s)." -ForegroundColor DarkGray

# ── Merge: classify each found company ────────────────────────────────────────

$toAdd  = [System.Collections.Generic.List[object]]::new()
$kept   = [System.Collections.Generic.List[object]]::new()
$stale  = [System.Collections.Generic.List[object]]::new()

foreach ($cid in $companyMap.Keys) {
    $info = $companyMap[$cid]
    $slug = $info.Slug
    if (-not $slug) { continue }

    if ($existing.ContainsKey($slug)) {
        $kept.Add($existing[$slug])
    }
    else {
        $toAdd.Add([PSCustomObject]@{
            HuduCompanySlug = $slug
            HuduCompanyName = $info.Name
            Modules         = $DefaultModules
        })
    }
}

# Entries in the file that are no longer in Hudu — will be commented out
$foundSlugs = $companyMap.Values | ForEach-Object { $_.Slug }
foreach ($slug in $existing.Keys) {
    if ($slug -notin $foundSlugs) {
        $stale.Add($existing[$slug])
    }
}

# ── Report ─────────────────────────────────────────────────────────────────────

Write-Host ""
Write-Host "  Existing (unchanged) : $($kept.Count)"  -ForegroundColor DarkGray
Write-Host "  New to add           : $($toAdd.Count)"  -ForegroundColor $(if ($toAdd.Count -gt 0) { 'Green' } else { 'DarkGray' })
Write-Host "  No longer in Hudu    : $($stale.Count)" -ForegroundColor $(if ($stale.Count -gt 0) { 'Yellow' } else { 'DarkGray' })

if ($toAdd.Count -gt 0) {
    $toAdd  | ForEach-Object { Write-Host "    + $($_.HuduCompanyName) ($($_.HuduCompanySlug))" -ForegroundColor Green }
}
if ($stale.Count -gt 0) {
    $stale  | ForEach-Object { Write-Host "    # $($_.HuduCompanyName) ($($_.HuduCompanySlug)) — commented out" -ForegroundColor Yellow }
}

if ($toAdd.Count -eq 0 -and $stale.Count -eq 0) {
    Write-Host ""
    Write-Host "No changes — UnattendedCustomers.psd1 is already up to date." -ForegroundColor Cyan
    exit 0
}

# ── Build active customer list (sorted) ───────────────────────────────────────

$activeEntries = @($kept + $toAdd | Sort-Object HuduCompanyName)
$staleEntries  = @($stale        | Sort-Object HuduCompanyName)

# ── Serialise to PSD1 ──────────────────────────────────────────────────────────

$outputLines = [System.Collections.Generic.List[string]]::new()
$outputLines.Add('# UnattendedCustomers.psd1 — auto-generated by Helpers\Sync-UnattendedCustomers.ps1')
$outputLines.Add("# Last synced: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
$outputLines.Add('#')
$outputLines.Add('# Modules:  1=Entra  2=Exchange  3=SharePoint  4=MailSec  5=Intune  6=Teams  7=Maester  A=All')
$outputLines.Add('')
$outputLines.Add('@{')
$outputLines.Add('    Customers = @(')

foreach ($entry in $activeEntries) {
    $slug = $entry.HuduCompanySlug
    $name = ($entry.HuduCompanyName -replace "'", "''")
    $mods = ($entry.Modules | ForEach-Object { "'$_'" }) -join ', '
    $outputLines.Add('        @{')
    $outputLines.Add("            HuduCompanySlug = '$slug'")
    $outputLines.Add("            HuduCompanyName = '$name'")
    $outputLines.Add("            Modules         = @($mods)")
    $outputLines.Add('        }')
}

# Stale entries — commented out so they are skipped at runtime but remain visible
if ($staleEntries.Count -gt 0) {
    $outputLines.Add('')
    $outputLines.Add('        # ── No longer found in Hudu — review and remove when confirmed offboarded ──')
    foreach ($entry in $staleEntries) {
        $slug = $entry.HuduCompanySlug
        $name = ($entry.HuduCompanyName -replace "'", "''")
        $mods = ($entry.Modules | ForEach-Object { "'$_'" }) -join ', '
        $outputLines.Add("#       @{")
        $outputLines.Add("#           HuduCompanySlug = '$slug'")
        $outputLines.Add("#           HuduCompanyName = '$name'")
        $outputLines.Add("#           Modules         = @($mods)")
        $outputLines.Add("#       }")
    }
}

$outputLines.Add('    )')
$outputLines.Add('}')

# ── Write ──────────────────────────────────────────────────────────────────────

if ($PSCmdlet.ShouldProcess($outputPath, 'Write UnattendedCustomers.psd1')) {
    Set-Content -Path $outputPath -Value $outputLines -Encoding UTF8
    Write-Host ""
    Write-Host "UnattendedCustomers.psd1 updated — $($activeEntries.Count) active, $($staleEntries.Count) commented out." -ForegroundColor Cyan
}

