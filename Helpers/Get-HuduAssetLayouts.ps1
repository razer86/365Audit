<#
.SYNOPSIS
    Lists all Hudu asset layouts with their numeric IDs.

.DESCRIPTION
    Queries the Hudu API and prints every asset layout name alongside its numeric ID.
    Use this to find the correct HuduAssetLayoutId value for config.psd1.

    Hudu's web UI shows slugs in URLs — the numeric ID is only available via the API.

.PARAMETER HuduBaseUrl
    Base URL of your Hudu instance (e.g. 'https://hudu.yourcompany.com').
    Falls back to HuduBaseUrl in config.psd1 if not provided.

.PARAMETER HuduApiKey
    Hudu API key (Administrator > API Keys).
    Falls back to HuduApiKey in config.psd1 if not provided.

.EXAMPLE
    .\Helpers\Get-HuduAssetLayouts.ps1

.EXAMPLE
    .\Helpers\Get-HuduAssetLayouts.ps1 -HuduBaseUrl 'https://hudu.example.com' -HuduApiKey 'abc123'

.NOTES
    Author  : Raymond Slater
    Version : 1.0.0
#>

#Requires -Version 7.2

param (
    [string]$HuduBaseUrl,
    [string]$HuduApiKey
)

$ErrorActionPreference = 'Stop'

# ── Load config.psd1 fallbacks ─────────────────────────────────────────────────

$_configPath = Join-Path $PSScriptRoot '..' 'config.psd1'
if (Test-Path $_configPath) {
    try {
        $_config = Import-PowerShellDataFile -Path $_configPath
        if (-not $HuduBaseUrl -and $_config.HuduBaseUrl) { $HuduBaseUrl = $_config.HuduBaseUrl }
        if (-not $HuduApiKey  -and $_config.HuduApiKey)  { $HuduApiKey  = $_config.HuduApiKey  }
    }
    catch { Write-Warning "Could not load config.psd1: $_" }
}

if (-not $HuduBaseUrl) {
    Write-Error "HuduBaseUrl is required. Set it in config.psd1 or pass -HuduBaseUrl."
}
if (-not $HuduApiKey) {
    Write-Error "HuduApiKey is required. Set it in config.psd1 or pass -HuduApiKey."
}

$HuduBaseUrl = $HuduBaseUrl.TrimEnd('/')
$_headers    = @{ 'x-api-key' = $HuduApiKey }

# ── Fetch asset layouts (paginated) ───────────────────────────────────────────

Write-Host "`nFetching asset layouts from $HuduBaseUrl..." -ForegroundColor Cyan

$layouts  = [System.Collections.Generic.List[object]]::new()
$page     = 1
$pageSize = 100

do {
    try {
        $response = Invoke-RestMethod `
            -Uri "$HuduBaseUrl/api/v1/asset_layouts?page_size=$pageSize&page=$page" `
            -Headers $_headers -Method Get -ErrorAction Stop
    }
    catch {
        Write-Error "Hudu API request failed: $_"
    }
    foreach ($l in @($response.asset_layouts)) { $layouts.Add($l) }
    $page++
} while ($response.asset_layouts.Count -gt 0)

if ($layouts.Count -eq 0) {
    Write-Warning "No asset layouts returned. Check your credentials and Hudu URL."
    exit 0
}

# ── Display results ────────────────────────────────────────────────────────────

Write-Host ""
Write-Host ("  {0,-6}  {1}" -f "ID", "Name") -ForegroundColor DarkGray
Write-Host ("  {0,-6}  {1}" -f "──────", "────────────────────────────────────────") -ForegroundColor DarkGray

foreach ($layout in ($layouts | Sort-Object id)) {
    Write-Host ("  {0,-6}  {1}" -f $layout.id, $layout.name)
}

Write-Host ""
Write-Host "Set HuduAssetLayoutId in config.psd1 to the ID of your audit credential layout." -ForegroundColor Yellow
Write-Host ""
