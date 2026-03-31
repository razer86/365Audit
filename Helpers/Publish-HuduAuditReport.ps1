<#
.SYNOPSIS
    Publishes a completed 365Audit report to a Hudu 'Monthly Audit Report' asset.

.DESCRIPTION
    After a successful audit run this script:
      1. Reads ActionItems.json and AuditMetrics.json written by Generate-AuditSummary.ps1
      2. Finds or creates the 'Monthly Audit Report' asset for the company in Hudu
      3. Downloads the prior month's zip attachment, extracts AuditMetrics.json and
         ActionItems.json, and computes month-over-month deltas:
           - Tile delta markers (<!-- TILE_DELTA_* -->) are replaced with coloured
             change indicators inside the KPI tile cards
           - A "Changes Since Last Month" section (resolved/new action items + key
             metric changes) is injected at <!-- AUDIT_DELTA_INJECT -->
      4. Populates the report_summary field with the (now-enriched) M365_HuduReport.html
      5. Uploads M365_AuditSummary.html (the full interactive report) as an attachment
      6. Compresses the full output folder to a zip and uploads that as an attachment

    All delta computation is wrapped in a non-fatal try/catch — if the prior zip is
    absent or the download fails, the report is published without delta information.

.PARAMETER OutputPath
    Path to the customer's audit output folder (e.g. C:\AuditReports\ContosoPty_20260326).

.PARAMETER CompanySlug
    Hudu company slug — used to locate the company via the API.

.PARAMETER HuduBaseUrl
    Base URL of your Hudu instance (no trailing slash).
    Optional — falls back to HuduBaseUrl in config.psd1.

.PARAMETER HuduApiKey
    Hudu API key (Administrator > API Keys).
    Optional — falls back to HuduApiKey in config.psd1.

.PARAMETER ReportLayoutId
    Asset layout ID for the 'Monthly Audit Report' layout in Hudu.
    Optional — falls back to HuduReportLayoutId in config.psd1, then 68.

.PARAMETER ReportAssetName
    Prefix for the monthly report asset name. Asset is created as "<ReportAssetName> - yyyy-MM".
    Optional — falls back to HuduReportAssetName in config.psd1, then 'M365 Monthly Audit'.

.EXAMPLE
    # Minimal — Hudu connection details read from config.psd1
    .\Helpers\Publish-HuduAuditReport.ps1 `
        -OutputPath  'C:\AuditReports\ContosoPty_20260326' `
        -CompanySlug 'a1b2c3d4e5f6'

.EXAMPLE
    # Explicit — override config.psd1 values
    .\Helpers\Publish-HuduAuditReport.ps1 `
        -OutputPath  'C:\AuditReports\ContosoPty_20260326' `
        -CompanySlug 'a1b2c3d4e5f6' `
        -HuduBaseUrl 'https://hudu.example.com' `
        -HuduApiKey  'your-api-key'

.NOTES
    Author      : Raymond Slater
    Version     : 1.5.0
#>

#Requires -Version 7.2

[CmdletBinding()]
param (
    [Parameter(Mandatory)]
    [string]$OutputPath,

    [Parameter(Mandatory)]
    [string]$CompanySlug,

    [string]$HuduBaseUrl,
    [string]$HuduApiKey,
    [int]$ReportLayoutId = 0,
    [string]$ReportAssetName = '',

    # When set, the local output folder is deleted after the uploaded zip is
    # downloaded back from Hudu and verified to open correctly.
    [switch]$CleanupLocal
)

$ScriptVersion         = "1.5.0"
Write-Verbose "Publish-HuduAuditReport.ps1 loaded (v$ScriptVersion)"

$ErrorActionPreference = 'Stop'

# ── Load config.psd1 fallbacks ─────────────────────────────────────────────────

$_configPath = Join-Path $PSScriptRoot '..\config.psd1'
if (Test-Path $_configPath) {
    try {
        $_config = Import-PowerShellDataFile -Path $_configPath
        if (-not $HuduBaseUrl)       { $HuduBaseUrl    = $_config.HuduBaseUrl }
        if (-not $HuduApiKey)        { $HuduApiKey     = $_config.HuduApiKey  }
        if ($ReportLayoutId -eq 0)   { $ReportLayoutId    = if ($_config.HuduReportLayoutId -gt 0) { $_config.HuduReportLayoutId } else { 68 } }
        if (-not $ReportAssetName)   { $ReportAssetName   = $_config.HuduReportAssetName }
    }
    catch { Write-Verbose "Could not load config.psd1: $_" }
}
else {
    if ($ReportLayoutId -eq 0) { $ReportLayoutId = 68 }
}
if (-not $ReportAssetName) { $ReportAssetName = 'M365 Monthly Audit' }

if (-not $HuduBaseUrl) { Write-Error 'HuduBaseUrl is required — supply -HuduBaseUrl or set it in config.psd1.' }
if (-not $HuduApiKey)  { Write-Error 'HuduApiKey is required — supply -HuduApiKey or set it in config.psd1.'  }

$base        = $HuduBaseUrl.TrimEnd('/')
$jsonHeaders = @{ 'x-api-key' = $HuduApiKey; 'Content-Type' = 'application/json' }
$formHeaders = @{ 'x-api-key' = $HuduApiKey }

# ── 1. Locate report files ─────────────────────────────────────────────────────

$huduBodyFile   = Get-Item (Join-Path $OutputPath 'M365_HuduReport.html')    -ErrorAction SilentlyContinue
$fullReportFile = Get-Item (Join-Path $OutputPath 'M365_AuditSummary.html')  -ErrorAction SilentlyContinue

if (-not $huduBodyFile) {
    Write-Warning "Publish-HuduAuditReport: M365_HuduReport.html not found in '$OutputPath' — skipping."
    return
}

if (-not $fullReportFile) {
    Write-Warning "Publish-HuduAuditReport: M365_AuditSummary.html not found in '$OutputPath' — full report attachment will be skipped."
}

# Load the Hudu report body — delta injection modifies this in memory before upload
$huduBodyContent = Get-Content $huduBodyFile.FullName -Raw -Encoding UTF8

# Load AuditMetrics.json for KPI field population (also used later for delta)
$auditMetrics = $null
$_metricsFilePath = Join-Path $OutputPath 'AuditMetrics.json'
if (Test-Path $_metricsFilePath) {
    try { $auditMetrics = Get-Content $_metricsFilePath -Raw | ConvertFrom-Json -ErrorAction Stop }
    catch { Write-Verbose "Could not parse AuditMetrics.json: $_" }
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

# ── 3. Find or create monthly audit asset ────────────────────────────────────

$assetName      = "$ReportAssetName - $(Get-Date -Format 'yyyy-MM')"
$priorMonthName = "$ReportAssetName - $((Get-Date).AddMonths(-1).ToString('yyyy-MM'))"
Write-Host "  [Hudu] Locating asset '$assetName'..." -ForegroundColor DarkCyan

$assetId      = $null
$priorAssetId = $null

try {
    $resp     = Invoke-RestMethod `
        -Uri "$base/api/v1/assets?asset_layout_id=$ReportLayoutId&company_id=$companyId&page_size=50" `
        -Headers $jsonHeaders -Method Get -ErrorAction Stop
    $existing  = @($resp.assets) | Where-Object { $_.name -eq $assetName }       | Select-Object -First 1
    $priorAsset = @($resp.assets) | Where-Object { $_.name -eq $priorMonthName } | Select-Object -First 1
    if ($existing)   { $assetId      = $existing.id }
    if ($priorAsset) { $priorAssetId = $priorAsset.id }
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

# ── 4. Month-over-month delta computation ─────────────────────────────────────

function Format-DeltaSpan {
    [CmdletBinding()]
    param(
        [double]$Delta,
        [switch]$InvertColour,
        [string]$Suffix    = '%',
        [int]   $Decimals  = 1
    )
    if ($Delta -eq 0) { return '' }
    $positive = $Delta -gt 0
    $good     = if ($InvertColour) { -not $positive } else { $positive }
    $colour   = if ($good) { '#16a34a' } else { '#dc2626' }
    $sign     = if ($positive) { '+' } else { '' }
    $val      = [math]::Round($Delta, $Decimals)
    return "<span style='font-size:10px;color:$colour;margin-left:4px;'>${sign}${val}${Suffix}</span>"
}

try {
    # Use already-loaded metrics; load action items
    $currentMetrics = $auditMetrics
    $currentItems   = @()
    $_actionItemPath = Join-Path $OutputPath 'ActionItems.json'
    if (Test-Path $_actionItemPath) { $currentItems = Get-Content $_actionItemPath -Raw | ConvertFrom-Json -ErrorAction SilentlyContinue }
    if ($null -eq $currentItems) { $currentItems = @() }

    if ($priorAssetId) {
        Write-Host "  [Hudu] Downloading prior month data for delta..." -ForegroundColor DarkCyan

        # Fetch uploads list for prior asset
        $uploadsResp = Invoke-RestMethod `
            -Uri "$base/api/v1/uploads?uploadable_id=$priorAssetId&uploadable_type=Asset" `
            -Headers $jsonHeaders -Method Get -ErrorAction Stop
        $priorZip = @($uploadsResp.uploads) | Where-Object { $_.file_name -like '*.zip' } | Select-Object -Last 1

        if ($priorZip) {
            $tmpZip = Join-Path $env:TEMP "365Audit_prior_$([System.IO.Path]::GetRandomFileName()).zip"
            Invoke-WebRequest -Uri $priorZip.url -Headers $formHeaders -OutFile $tmpZip -ErrorAction Stop
            Write-Verbose "Prior zip downloaded: $($priorZip.file_name)"

            # Extract sidecar files from zip
            $priorMetrics = $null
            $priorItems   = @()
            Add-Type -AssemblyName System.IO.Compression.FileSystem
            $zipArchive = [System.IO.Compression.ZipFile]::OpenRead($tmpZip)
            try {
                $metricsEntry = $zipArchive.Entries | Where-Object { $_.FullName -match 'AuditMetrics\.json$' }  | Select-Object -First 1
                $itemsEntry   = $zipArchive.Entries | Where-Object { $_.FullName -match 'ActionItems\.json$'   } | Select-Object -First 1

                if ($metricsEntry) {
                    $reader = [System.IO.StreamReader]::new($metricsEntry.Open())
                    $priorMetrics = $reader.ReadToEnd() | ConvertFrom-Json -ErrorAction SilentlyContinue
                    $reader.Dispose()
                }
                if ($itemsEntry) {
                    $reader = [System.IO.StreamReader]::new($itemsEntry.Open())
                    $priorItems = $reader.ReadToEnd() | ConvertFrom-Json -ErrorAction SilentlyContinue
                    if ($null -eq $priorItems) { $priorItems = @() }
                    $reader.Dispose()
                }
            }
            finally { $zipArchive.Dispose() }
            Remove-Item $tmpZip -ErrorAction SilentlyContinue

            # ── Tile delta markers ─────────────────────────────────────────────
            $tileDeltas = @{}
            if ($currentMetrics -and $priorMetrics) {
                # MFA — higher is better
                if ($null -ne $currentMetrics.MfaCoveragePct -and $null -ne $priorMetrics.MfaCoveragePct) {
                    $tileDeltas['TILE_DELTA_MFA'] = Format-DeltaSpan ($currentMetrics.MfaCoveragePct - $priorMetrics.MfaCoveragePct)
                }
                # Secure Score — higher is better
                if ($null -ne $currentMetrics.SecureScorePct -and $null -ne $priorMetrics.SecureScorePct) {
                    $tileDeltas['TILE_DELTA_SCORE'] = Format-DeltaSpan ($currentMetrics.SecureScorePct - $priorMetrics.SecureScorePct)
                }
                # Devices — neutral count delta; colour by non-compliant direction
                if ($null -ne $currentMetrics.ManagedDeviceCount -and $null -ne $priorMetrics.ManagedDeviceCount) {
                    $deviceDelta = $currentMetrics.ManagedDeviceCount - $priorMetrics.ManagedDeviceCount
                    if ($deviceDelta -ne 0) {
                        $sign = if ($deviceDelta -gt 0) { '+' } else { '' }
                        $tileDeltas['TILE_DELTA_DEVICES'] = "<span style='font-size:10px;color:#64748b;margin-left:4px;'>${sign}$deviceDelta</span>"
                    }
                }
                # Storage — higher is bad
                if ($null -ne $currentMetrics.TenantStoragePct -and $null -ne $priorMetrics.TenantStoragePct) {
                    $tileDeltas['TILE_DELTA_STORAGE'] = Format-DeltaSpan ($currentMetrics.TenantStoragePct - $priorMetrics.TenantStoragePct) -InvertColour
                }
                # Action items total — fewer is better
                $curAi  = [int]$currentMetrics.ActionItemCritical + [int]$currentMetrics.ActionItemWarning
                $prevAi = [int]$priorMetrics.ActionItemCritical   + [int]$priorMetrics.ActionItemWarning
                if ($curAi -ne $prevAi) {
                    $tileDeltas['TILE_DELTA_AI'] = Format-DeltaSpan ($curAi - $prevAi) -Suffix '' -Decimals 0 -InvertColour
                }
            }

            # Replace tile markers
            foreach ($marker in $tileDeltas.Keys) {
                $huduBodyContent = $huduBodyContent.Replace("<!-- $marker -->", $tileDeltas[$marker])
            }

            # ── Action item diff ───────────────────────────────────────────────
            function Get-ItemKey { param($Item)
                if ($Item.CheckId) { return $Item.CheckId }
                return "$($Item.Category)|$($Item.Text -replace '<br>', ' ' -replace '<[^>]+>', '')"
            }

            $priorMap   = @{}
            $currentMap = @{}
            foreach ($item in @($priorItems))   { $priorMap[(Get-ItemKey $item)]   = $item }
            foreach ($item in @($currentItems)) { $currentMap[(Get-ItemKey $item)] = $item }

            $resolved = @($priorMap.Keys   | Where-Object { -not $currentMap.ContainsKey($_) } | ForEach-Object { $priorMap[$_] })
            $newItems  = @($currentMap.Keys | Where-Object { -not $priorMap.ContainsKey($_) }  | ForEach-Object { $currentMap[$_] })

            # ── Build delta section HTML ───────────────────────────────────────
            $deltaHtml = ''

            if ($resolved.Count -gt 0 -or $newItems.Count -gt 0 -or ($currentMetrics -and $priorMetrics)) {
                $metricRows = ''
                if ($currentMetrics -and $priorMetrics) {
                    $licAssignedDelta = '&mdash;'
                    if ($null -ne $currentMetrics.LicenseTotalAssigned -and $null -ne $priorMetrics.LicenseTotalAssigned) {
                        $d = [int]$currentMetrics.LicenseTotalAssigned - [int]$priorMetrics.LicenseTotalAssigned
                        if ($d -ne 0) {
                            $sign = if ($d -gt 0) { '+' } else { '' }
                            $licAssignedDelta = "${sign}$d"
                        }
                    }
                    $storageDelta = '&mdash;'
                    if ($null -ne $currentMetrics.TenantStorageUsedGB -and $null -ne $priorMetrics.TenantStorageUsedGB) {
                        $d = [math]::Round([double]$currentMetrics.TenantStorageUsedGB - [double]$priorMetrics.TenantStorageUsedGB, 1)
                        if ($d -ne 0) {
                            $sign = if ($d -gt 0) { '+' } else { '' }
                            $storageDelta = "${sign}${d} GB"
                        }
                    }
                    $metricRows = "<table style='width:100%;border-collapse:collapse;font-size:13px;margin-bottom:10px;'>" +
                        "<thead><tr style='border-bottom:1px solid rgba(128,128,128,0.2);'>" +
                        "<th style='text-align:left;padding:4px 8px;'>Metric</th>" +
                        "<th style='text-align:right;padding:4px 8px;'>Change</th></tr></thead><tbody>" +
                        "<tr><td style='padding:4px 8px;'>Assigned Licences</td><td style='text-align:right;padding:4px 8px;'>$licAssignedDelta</td></tr>" +
                        "<tr><td style='padding:4px 8px;'>Tenant Storage Used</td><td style='text-align:right;padding:4px 8px;'>$storageDelta</td></tr>" +
                        "</tbody></table>"
                }

                $resolvedTable = ''
                if ($resolved.Count -gt 0) {
                    $rows = ($resolved | Sort-Object { $_.Severity }, { $_.Category } | ForEach-Object {
                        "<tr><td style='padding:4px 8px;'>$($_.Category)</td>" +
                        "<td style='padding:4px 8px;'>$($_.Text -replace '<br>', ' ')</td></tr>"
                    }) -join ''
                    $resolvedTable = "<p style='font-weight:600;color:#16a34a;margin:10px 0 4px;'>Resolved ($($resolved.Count))</p>" +
                        "<table style='width:100%;border-collapse:collapse;font-size:13px;margin-bottom:10px;'>" +
                        "<thead><tr style='border-bottom:1px solid rgba(128,128,128,0.2);'>" +
                        "<th style='text-align:left;padding:4px 8px;'>Category</th>" +
                        "<th style='text-align:left;padding:4px 8px;'>Finding</th></tr></thead>" +
                        "<tbody>$rows</tbody></table>"
                }

                $newTable = ''
                if ($newItems.Count -gt 0) {
                    $rows = ($newItems | Sort-Object { $_.Severity }, { $_.Category } | ForEach-Object {
                        $sev = if ($_.Severity -eq 'critical') { '#dc2626' } else { '#d97706' }
                        "<tr><td style='padding:4px 8px;color:$sev;font-weight:600;'>$($_.Category)</td>" +
                        "<td style='padding:4px 8px;'>$($_.Text -replace '<br>', ' ')</td></tr>"
                    }) -join ''
                    $newTable = "<p style='font-weight:600;color:#dc2626;margin:10px 0 4px;'>New ($($newItems.Count))</p>" +
                        "<table style='width:100%;border-collapse:collapse;font-size:13px;margin-bottom:10px;'>" +
                        "<thead><tr style='border-bottom:1px solid rgba(128,128,128,0.2);'>" +
                        "<th style='text-align:left;padding:4px 8px;'>Category</th>" +
                        "<th style='text-align:left;padding:4px 8px;'>Finding</th></tr></thead>" +
                        "<tbody>$rows</tbody></table>"
                }

                $deltaHtml = "<details style='margin-bottom:16px;border:1px solid rgba(128,128,128,0.2);border-radius:14px;overflow:hidden;box-shadow:0 1px 3px rgba(0,0,0,.04);'>" +
                    "<summary style='padding:12px 24px;border-bottom:2px solid rgba(30,58,95,0.12);font-size:18px;font-weight:700;cursor:pointer;list-style:none;display:flex;align-items:center;gap:10px;'>" +
                    "<span style='display:inline-flex;align-items:center;justify-content:center;min-width:28px;height:28px;border-radius:8px;background:#1e3a5f;color:#fff;font-size:13px;font-weight:700;flex-shrink:0;'>&#9651;</span>" +
                    "Changes Since Last Month</summary>" +
                    "<div style='padding:20px 24px;'>$metricRows$resolvedTable$newTable</div></details>"
            }

            $huduBodyContent = $huduBodyContent.Replace('<!-- AUDIT_DELTA_INJECT -->', $deltaHtml)
        }
        else {
            Write-Verbose "No zip attachment found on prior month asset — delta skipped."
        }
    }
    else {
        Write-Verbose "No prior month asset found ('$priorMonthName') — first-run, delta skipped."
    }
}
catch {
    Write-Warning "Publish-HuduAuditReport: Delta computation failed (non-fatal) — $_"
}

# Clear any unfilled delta markers (no prior data, or delta computation skipped)
$huduBodyContent = $huduBodyContent -replace '<!-- AUDIT_DELTA_INJECT -->', ''
$huduBodyContent = $huduBodyContent -replace '<!-- TILE_DELTA_[A-Z_]+ -->', ''

# ── 5. Update report_summary asset field ──────────────────────────────────────

Write-Host "  [Hudu] Updating asset fields..." -ForegroundColor DarkCyan

# Build KPI field values from AuditMetrics.json (plain text — shown in Hudu asset table)
$_fieldMfa      = if ($null -ne $auditMetrics.MfaCoveragePct)   { "$([math]::Round([double]$auditMetrics.MfaCoveragePct, 1))%" }        else { '' }
$_fieldScore    = if ($null -ne $auditMetrics.SecureScoreCurrent -and $null -ne $auditMetrics.SecureScoreMax) {
    "$([math]::Round([double]$auditMetrics.SecureScoreCurrent, 0)) / $([math]::Round([double]$auditMetrics.SecureScoreMax, 0))"
} elseif ($null -ne $auditMetrics.SecureScoreCurrent) { "$([math]::Round([double]$auditMetrics.SecureScoreCurrent, 0))" } else { '' }
$_fieldStorage  = if ($null -ne $auditMetrics.TenantStoragePct) {
    $pct = [math]::Round([double]$auditMetrics.TenantStoragePct, 1)
    if ($null -ne $auditMetrics.TenantStorageUsedGB -and $null -ne $auditMetrics.TenantStorageTotalGB) {
        $used  = [math]::Round([double]$auditMetrics.TenantStorageUsedGB, 0)
        $total = [math]::Round([double]$auditMetrics.TenantStorageTotalGB, 0)
        "${pct}% (${used} GB / ${total} GB)"
    } else { "${pct}%" }
} else { '' }
$_fieldCritical = if ($null -ne $auditMetrics.ActionItemCritical) { "$([int]$auditMetrics.ActionItemCritical)" } else { '' }

$updateBody = @{
    asset = @{
        custom_fields = @(
            @{ report_summary = $huduBodyContent }
            @{ mfa_coverage   = $_fieldMfa      }
            @{ secure_score   = $_fieldScore    }
            @{ tenant_storage = $_fieldStorage  }
            @{ critical_items = $_fieldCritical }
        )
    }
} | ConvertTo-Json -Depth 6 -Compress

try {
    Invoke-RestMethod -Uri "$base/api/v1/companies/$companyId/assets/$assetId" -Method Put `
        -Headers $jsonHeaders -Body $updateBody -ErrorAction Stop | Out-Null
    Write-Host "  [Hudu] Asset fields updated." -ForegroundColor DarkCyan
}
catch { Write-Warning "Publish-HuduAuditReport: Field update failed — $_" }

# ── 6. Upload M365_AuditSummary.html (full report) to asset ───────────────────

if ($fullReportFile) {
    Write-Host "  [Hudu] Uploading full HTML report to asset..." -ForegroundColor DarkCyan
    try {
        Invoke-RestMethod -Uri "$base/api/v1/uploads" -Method Post -Headers $formHeaders `
            -Form @{
                file                      = (Get-Item $fullReportFile.FullName)
                'upload[uploadable_id]'   = "$assetId"
                'upload[uploadable_type]' = 'Asset'
            } -ErrorAction Stop | Out-Null
        Write-Host "  [Hudu] Full HTML report uploaded." -ForegroundColor DarkCyan
    }
    catch { Write-Warning "Publish-HuduAuditReport: Full HTML upload failed — $_" }
}

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

# Verify archive integrity before uploading
Add-Type -AssemblyName System.IO.Compression.FileSystem
try {
    $_localZip    = [System.IO.Compression.ZipFile]::OpenRead($zipPath)
    $_localCount  = $_localZip.Entries.Count
    $_localZip.Dispose()
    if ($_localCount -eq 0) {
        Write-Warning "Publish-HuduAuditReport: Archive appears empty — skipping upload."
        Remove-Item $zipPath -ErrorAction SilentlyContinue
        return
    }
    Write-Verbose "Local archive verified ($_localCount entries)."
}
catch {
    Write-Warning "Publish-HuduAuditReport: Archive verification failed — $_ — skipping upload."
    Remove-Item $zipPath -ErrorAction SilentlyContinue
    return
}

Write-Host "  [Hudu] Uploading zip archive to asset..." -ForegroundColor DarkCyan

$_uploadResp      = $null
$_uploadSucceeded = $false
try {
    $_uploadResp = Invoke-RestMethod -Uri "$base/api/v1/uploads" -Method Post -Headers $formHeaders `
        -Form @{
            file                      = (Get-Item $zipPath)
            'upload[uploadable_id]'   = "$assetId"
            'upload[uploadable_type]' = 'Asset'
        } -ErrorAction Stop
    Write-Host "  [Hudu] Archive uploaded: $(Split-Path $zipPath -Leaf)" -ForegroundColor Green
    $_uploadSucceeded = $true
}
catch { Write-Warning "Publish-HuduAuditReport: Zip upload failed — $_" }
finally { Remove-Item $zipPath -ErrorAction SilentlyContinue }

# ── 8. Verify uploaded archive and clean up local output folder ───────────────

if ($CleanupLocal -and $_uploadSucceeded) {
    Write-Host "  [Hudu] Verifying uploaded archive before local cleanup..." -ForegroundColor DarkCyan
    try {
        # Resolve the download URL — prefer the upload response, fall back to re-querying the asset
        $_zipUrl = if ($_uploadResp -and $_uploadResp.upload) { $_uploadResp.upload.url } else { $null }
        if (-not $_zipUrl) {
            $uploadsResp = Invoke-RestMethod `
                -Uri "$base/api/v1/uploads?uploadable_id=$assetId&uploadable_type=Asset" `
                -Headers $jsonHeaders -Method Get -ErrorAction Stop
            $_zipUrl = @($uploadsResp.uploads) |
                Where-Object { $_.file_name -like '*.zip' } |
                Select-Object -Last 1 |
                Select-Object -ExpandProperty url
        }

        if (-not $_zipUrl) {
            Write-Warning "Publish-HuduAuditReport: Could not resolve uploaded zip URL — local folder kept."
        }
        else {
            $tmpVerify = Join-Path $env:TEMP "365Audit_verify_$([System.IO.Path]::GetRandomFileName()).zip"
            try {
                Invoke-WebRequest -Uri $_zipUrl -Headers $formHeaders -OutFile $tmpVerify -ErrorAction Stop
                $_verifyZip   = [System.IO.Compression.ZipFile]::OpenRead($tmpVerify)
                $_verifyCount = $_verifyZip.Entries.Count
                $_verifyZip.Dispose()

                if ($_verifyCount -gt 0) {
                    Write-Host "  [Hudu] Uploaded archive verified ($_verifyCount entries) — removing local folder." -ForegroundColor DarkCyan
                    Remove-Item -Path $OutputPath -Recurse -Force -ErrorAction Stop
                    Write-Host "  [Hudu] Local report folder removed: $(Split-Path $OutputPath -Leaf)" -ForegroundColor Green
                }
                else {
                    Write-Warning "Publish-HuduAuditReport: Uploaded archive appears empty — local folder kept."
                }
            }
            finally {
                Remove-Item $tmpVerify -ErrorAction SilentlyContinue
            }
        }
    }
    catch {
        Write-Warning "Publish-HuduAuditReport: Cleanup verification failed — local folder kept: $_"
    }
}

