<#
.SYNOPSIS
    Publishes a completed 365Audit report to a Hudu 'Monthly Audit Report' asset.

.DESCRIPTION
    After a successful audit run this script:
      1. Reads ActionItems.json written by Generate-AuditSummary.ps1
      2. Finds or creates the 'Monthly Audit Report' asset for the company in Hudu
      3. Populates the 'Critical Action Items' and 'Warning Action Items' RichText fields
      4. Uploads M365_HuduReport.html as an attachment to the asset
      5. Compresses the full output folder to a zip and uploads that as an attachment

.PARAMETER OutputPath
    Path to the customer's audit output folder (e.g. C:\AuditReports\ContosoPty_20260326).

.PARAMETER CompanySlug
    Hudu company slug — used to locate the company via the API.

.PARAMETER HuduBaseUrl
    Base URL of your Hudu instance (no trailing slash).

.PARAMETER HuduApiKey
    Hudu API key (Administrator > API Keys).

.PARAMETER ReportLayoutId
    Asset layout ID for the 'Monthly Audit Report' layout in Hudu. Default: 68.

.EXAMPLE
    .\Helpers\Publish-HuduAuditReport.ps1 `
        -OutputPath  'C:\AuditReports\ContosoPty_20260326' `
        -CompanySlug 'contoso-pty-ltd' `
        -HuduBaseUrl 'https://hudu.example.com' `
        -HuduApiKey  'your-api-key'

.NOTES
    Author  : Raymond Slater
    Version : 1.1.0
#>

#Requires -Version 7.2

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$OutputPath,

    [Parameter(Mandatory)]
    [string]$CompanySlug,

    [Parameter(Mandatory)]
    [string]$HuduBaseUrl,

    [Parameter(Mandatory)]
    [string]$HuduApiKey,

    [int]$ReportLayoutId = 68
)

$ErrorActionPreference = 'Stop'
$base        = $HuduBaseUrl.TrimEnd('/')
$jsonHeaders = @{ 'x-api-key' = $HuduApiKey; 'Content-Type' = 'application/json' }
$formHeaders = @{ 'x-api-key' = $HuduApiKey }

# ── 1. Locate report files ─────────────────────────────────────────────────────

$huduHtmlFile = Get-Item (Join-Path $OutputPath 'M365_HuduReport.html') -ErrorAction SilentlyContinue

if (-not $huduHtmlFile) {
    Write-Warning "Publish-HuduAuditReport: M365_HuduReport.html not found in '$OutputPath' — skipping."
    return
}

# ── 2. Resolve Hudu company ID from slug ──────────────────────────────────────

Write-Host "  [Hudu] Resolving company '$CompanySlug'..." -ForegroundColor DarkCyan

$companyId = $null

try {
    $resp  = Invoke-RestMethod `
        -Uri "$base/api/v1/companies?slug=$([Uri]::EscapeDataString($CompanySlug))&page_size=25" `
        -Headers $jsonHeaders -Method Get -ErrorAction Stop
    $match = @($resp.companies) | Where-Object { $_.slug -eq $CompanySlug } | Select-Object -First 1
    if ($match) { $companyId = $match.id }
}
catch { Write-Verbose "Slug param query failed: $_" }

if (-not $companyId) {
    try {
        $resp  = Invoke-RestMethod `
            -Uri "$base/api/v1/companies?search=$([Uri]::EscapeDataString($CompanySlug))&page_size=50" `
            -Headers $jsonHeaders -Method Get -ErrorAction Stop
        $match = @($resp.companies) | Where-Object { $_.slug -eq $CompanySlug } | Select-Object -First 1
        if ($match) { $companyId = $match.id }
    }
    catch {
        Write-Warning "Publish-HuduAuditReport: Company lookup failed — $_"
        return
    }
}

if (-not $companyId) {
    Write-Warning "Publish-HuduAuditReport: Company slug '$CompanySlug' not found in Hudu — skipping."
    return
}

Write-Verbose "Resolved company ID: $companyId"

# ── 4. Find or create monthly audit asset ────────────────────────────────────

$assetName = "M365 Audit - $(Get-Date -Format 'yyyy-MM')"
Write-Host "  [Hudu] Locating asset '$assetName'..." -ForegroundColor DarkCyan

$assetId = $null

try {
    $resp     = Invoke-RestMethod `
        -Uri "$base/api/v1/assets?asset_layout_id=$ReportLayoutId&company_id=$companyId&page_size=50" `
        -Headers $jsonHeaders -Method Get -ErrorAction Stop
    $existing = @($resp.assets) | Where-Object { $_.name -eq $assetName } | Select-Object -First 1
    if ($existing) { $assetId = $existing.id }
}
catch { Write-Verbose "Asset query failed: $_" }

if (-not $assetId) {
    Write-Host "  [Hudu] Creating asset '$assetName'..." -ForegroundColor DarkCyan
    $createBody = @{
        asset_layout_id = $ReportLayoutId
        name            = $assetName
    } | ConvertTo-Json -Depth 3 -Compress
    try {
        $resp    = Invoke-RestMethod -Uri "$base/api/v1/companies/$companyId/assets" -Method Post `
            -Headers $jsonHeaders -Body $createBody -ErrorAction Stop
        $assetId = $resp.asset.id
        Write-Host "  [Hudu] Asset created (ID $assetId)." -ForegroundColor DarkCyan
    }
    catch {
        Write-Warning "Publish-HuduAuditReport: Asset creation failed — $_"
        return
    }
}
else {
    Write-Host "  [Hudu] Existing asset found (ID $assetId)." -ForegroundColor DarkCyan
}

# ── 5. Update report_summary asset field ──────────────────────────────────────

Write-Host "  [Hudu] Updating asset fields..." -ForegroundColor DarkCyan

$updateBody = @{
    asset = @{
        custom_fields = @(
            @{ report_summary = Get-Content $huduHtmlFile.FullName -Raw -Encoding UTF8 }
        )
    }
} | ConvertTo-Json -Depth 6 -Compress

try {
    Invoke-RestMethod -Uri "$base/api/v1/companies/$companyId/assets/$assetId" -Method Put `
        -Headers $jsonHeaders -Body $updateBody -ErrorAction Stop | Out-Null
    Write-Host "  [Hudu] Asset fields updated." -ForegroundColor DarkCyan
}
catch { Write-Warning "Publish-HuduAuditReport: Field update failed — $_" }

# ── 6. Upload M365_HuduReport.html to asset ───────────────────────────────────

Write-Host "  [Hudu] Uploading HTML report to asset..." -ForegroundColor DarkCyan

try {
    Invoke-RestMethod -Uri "$base/api/v1/uploads" -Method Post -Headers $formHeaders `
        -Form @{
            file                      = (Get-Item $huduHtmlFile.FullName)
            'upload[uploadable_id]'   = "$assetId"
            'upload[uploadable_type]' = 'Asset'
        } -ErrorAction Stop | Out-Null
    Write-Host "  [Hudu] HTML report uploaded." -ForegroundColor DarkCyan
}
catch { Write-Warning "Publish-HuduAuditReport: HTML upload failed — $_" }

# ── 7. Compress output folder and attach zip to asset ─────────────────────────

$zipPath = "$OutputPath.zip"
Write-Host "  [Hudu] Compressing output folder..." -ForegroundColor DarkCyan

try {
    Compress-Archive -Path $OutputPath -DestinationPath $zipPath -Force -ErrorAction Stop
}
catch {
    Write-Warning "Publish-HuduAuditReport: Compression failed — $_"
    return
}

Write-Host "  [Hudu] Uploading zip archive to asset..." -ForegroundColor DarkCyan

try {
    Invoke-RestMethod -Uri "$base/api/v1/uploads" -Method Post -Headers $formHeaders `
        -Form @{
            file                      = (Get-Item $zipPath)
            'upload[uploadable_id]'   = "$assetId"
            'upload[uploadable_type]' = 'Asset'
        } -ErrorAction Stop | Out-Null
    Write-Host "  [Hudu] Archive uploaded: $(Split-Path $zipPath -Leaf)" -ForegroundColor Green
}
catch { Write-Warning "Publish-HuduAuditReport: Zip upload failed — $_" }
