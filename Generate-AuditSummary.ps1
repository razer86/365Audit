<#
.SYNOPSIS
    Generates an HTML summary report from Microsoft 365 audit CSV output files.

.DESCRIPTION
    Reads CSV files produced by the audit modules and compiles them into a
    single styled HTML report with collapsible sections and links to the full
    CSV exports.

    CSVs consumed:
    - Entra_Users.csv
    - Entra_Users_Unlicensed.csv
    - Entra_GlobalAdmins.csv
    - Entra_AdminRoles.csv
    - Entra_Licenses.csv
    - Entra_SSPR.csv
    - Entra_SecurityDefaults.csv
    - Entra_SecureScore.csv
    - Entra_SecureScoreControls.csv
    - Entra_CA_Policies.csv
    - Entra_SignIns.csv
    - Entra_AccountCreations.csv
    - Entra_AccountDeletions.csv
    - Entra_AuditEvents.csv
    - Entra_Groups.csv
    - Exchange_Mailboxes.csv
    - Exchange_InboxForwardingRules.csv
    - SharePoint_TenantStorage.csv
    - SharePoint_Sites.csv
    - SharePoint_SPGroups.csv
    - SharePoint_ExternalSharing_Tenant.csv
    - SharePoint_ExternalSharing_SiteOverrides.csv
    - SharePoint_OneDriveUsage.csv
    - SharePoint_OneDrive_Unlicensed.csv
    - SharePoint_AccessControlPolicies.csv
    - MailSec_DKIM.csv
    - MailSec_DMARC.csv
    - MailSec_SPF.csv
    - Entra_PartnerRelationships.csv
    - Entra_EnterpriseApps.csv
    - Entra_RiskyUsers.csv
    - Entra_RiskySignIns.csv
    - Exchange_MailConnectors.csv
    - Intune_LicenceCheck.csv
    - Intune_Devices.csv
    - Intune_DeviceComplianceStates.csv
    - Intune_CompliancePolicies.csv
    - Intune_CompliancePolicySettings.csv
    - Intune_ConfigProfiles.csv
    - Intune_ConfigProfileSettings.csv
    - Intune_Apps.csv
    - Intune_AutopilotDevices.csv
    - Intune_EnrollmentRestrictions.csv
    - Teams_FederationConfig.csv
    - Teams_ClientConfig.csv
    - Teams_MeetingPolicies.csv
    - Teams_GuestMeetingConfig.csv
    - Teams_GuestCallingConfig.csv
    - Teams_MessagingPolicies.csv
    - Teams_AppPermissionPolicies.csv
    - Teams_AppSetupPolicies.csv
    - Teams_ChannelPolicies.csv
    - Exchange_OrgConfig.csv
    - Exchange_ExternalSenderTagging.csv
    - Exchange_ConnectionFilter.csv
    - Exchange_OwaPolicy.csv
    - Entra_OrgSettings.csv

.NOTES
    Author      : Raymond Slater
    Version     : 1.50.0
    Change Log  : See CHANGELOG.md

.LINK
    https://github.com/razer86/365Audit
#>

#Requires -Version 7.2

param (
    [Parameter(Mandatory)]
    [string]$AuditFolder,
    [switch]$DevMode = $false,
    [switch]$NoOpen,
    [int]$CertExpiryDays = -1,
    [string]$OutputPath
)

if (-not $DevMode -and $MyInvocation.InvocationName -eq $MyInvocation.MyCommand.Name) {
    Write-Error "This script must be run from the 365Audit launcher. Use -DevMode for development." -ErrorAction Stop
}

$ScriptVersion = "1.50.0"
Write-Verbose "Generate-AuditSummary.ps1 loaded (v$ScriptVersion)"

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Load config.psd1 — MSP-specific values (domains)
$_configPath    = Join-Path $PSScriptRoot 'config.psd1'
$_mspDomains    = @()
if (Test-Path $_configPath) {
    try {
        $_config = Import-PowerShellDataFile -Path $_configPath
        if ($_config.MspDomains) { $_mspDomains = @($_config.MspDomains) }
    }
    catch { Write-Warning "Could not load config.psd1: $_" }
}

if (-not (Test-Path $AuditFolder)) {
    Write-Error "Provided audit folder does not exist: $AuditFolder"
    exit 1
}

# Module output subfolders
$rawDir      = Join-Path $AuditFolder "Raw"
$entraDir    = $rawDir
$exchangeDir = $rawDir
$spDir       = $rawDir
$mailSecDir  = $rawDir
$intuneDir   = $rawDir
$teamsDir    = $rawDir

# =========================================
# ===   HTML Section Helper             ===
# =========================================
function Add-Section {
    [CmdletBinding()]
    param (
        [string]   $Title,
        [string]   $AnchorId = "",
        [string[]] $CsvFiles,
        [string]   $SummaryHtml,
        [string]   $ModuleVersion,
        [switch]   $Collapsed
    )

    if ([string]::IsNullOrWhiteSpace($AnchorId)) {
        $AnchorId = ($Title -replace '[^a-zA-Z0-9]+', '-' -replace '^-|-$', '').ToLower()
    }

    $csvLinks = ""
    if ($CsvFiles.Count -gt 0) {
        $fileItems = ""
        foreach ($file in $CsvFiles) {
            $name      = [System.IO.Path]::GetFileName($file)
            $href      = ([System.IO.Path]::GetRelativePath($script:ReportBaseDir, $file) -replace '\\', '/')
            $fileItems += "<li><a href='$href' target='_blank'>$name</a></li>"
        }
        $csvLinks = "<details class='raw-files'><summary>Raw CSV Files ($($CsvFiles.Count))</summary><ul class='raw-files-list'>$fileItems</ul></details>"
    }

    $encodedTitle   = [System.Net.WebUtility]::HtmlEncode($Title)
    $encodedVersion = if ([string]::IsNullOrWhiteSpace($ModuleVersion)) { "" } else { [System.Net.WebUtility]::HtmlEncode($ModuleVersion) }
    $versionMarkup  = if ($encodedVersion) { "<span class='section-version'>$encodedVersion</span>" } else { "" }
    $bodyStyle      = if ($Collapsed) { " style='display:none'" } else { "" }
    $toggleClass    = if ($Collapsed) { "" } else { " open" }

    return @"
<section class='module' id='$AnchorId'>
  <div class='module-hdr' onclick='toggleModule(this)'>
    <span class='module-title'>$encodedTitle</span>$versionMarkup
    <span class='module-toggle$toggleClass'>&#9658;</span>
  </div>
  <div class='module-body'$bodyStyle>
    $SummaryHtml
    $csvLinks
  </div>
</section>
"@
}

$script:ModuleVersionCache = @{}
function Get-ModuleScriptVersion {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ScriptName
    )

    if ($script:ModuleVersionCache.ContainsKey($ScriptName)) {
        return $script:ModuleVersionCache[$ScriptName]
    }

    $scriptPath = Join-Path $PSScriptRoot $ScriptName
    if (-not (Test-Path $scriptPath)) {
        $script:ModuleVersionCache[$ScriptName] = $null
        return $null
    }

    $match = Select-String -Path $scriptPath -Pattern '^\$ScriptVersion\s*=\s*"([^"]+)"' | Select-Object -First 1
    $version = if ($match -and $match.Matches.Count -gt 0) {
        "v{0}" -f $match.Matches[0].Groups[1].Value
    } else {
        $null
    }

    $script:ModuleVersionCache[$ScriptName] = $version
    return $version
}

function ConvertTo-HtmlText {
    [CmdletBinding()]
    param(
        $Value,
        [string]$NullText = '&mdash;'
    )

    if ($null -eq $Value) {
        return $NullText
    }

    $text = [string]$Value
    if ([string]::IsNullOrWhiteSpace($text)) {
        return $NullText
    }

    return [System.Net.WebUtility]::HtmlEncode($text)
}

function ConvertTo-HtmlMultilineText {
    [CmdletBinding()]
    param(
        $Value,
        [string]$NullText = '&mdash;'
    )

    $encoded = ConvertTo-HtmlText -Value $Value -NullText $NullText
    return (($encoded -replace "(`r`n|`n|`r)", '<br>'))
}

function Get-ExpandHintHtml {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Text
    )

    return "<p class='expand-hint'>{0}</p>" -f ([System.Net.WebUtility]::HtmlEncode($Text))
}

# =========================================
# ===   HTML Page Header                ===
# =========================================
if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $reportPath = Join-Path $AuditFolder "M365_AuditSummary.html"
}
else {
    $reportPath = $OutputPath
}

$script:ReportBaseDir = Split-Path -Parent $reportPath
if ([string]::IsNullOrWhiteSpace($script:ReportBaseDir)) {
    $script:ReportBaseDir = (Get-Location).Path
}

if (-not (Test-Path $script:ReportBaseDir)) {
    New-Item -ItemType Directory -Path $script:ReportBaseDir -Force | Out-Null
}

$reportDate = Get-Date -Format "dd MMMM yyyy HH:mm"

$html = [System.Collections.Generic.List[string]]::new()
$html.Add(@"
<!DOCTYPE html>
<html lang='en'>
<head>
<meta charset='UTF-8'>
<title>Microsoft 365 Audit Summary</title>
<style>
/* ── Reset & layout ── */
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0;}
body{margin:0;padding:0;font-family:'Segoe UI',system-ui,sans-serif;background:#f0f4f8;color:#1e293b;font-size:14px;}
.page-wrapper{height:100vh;overflow-y:auto;overflow-x:hidden;display:flex;flex-direction:column;scroll-behavior:smooth;}
/* ── App header ── */
.sticky-header{position:sticky;top:0;z-index:20;}
.app-header{background:linear-gradient(135deg,#0f2744 0%,#1d4ed8 100%);color:#fff;padding:0.6rem 1.25rem;display:grid;grid-template-columns:1fr auto 1fr;align-items:center;flex-shrink:0;box-shadow:0 2px 8px rgba(0,0,0,0.25);}
.app-header h1{font-size:0.97rem;font-weight:700;letter-spacing:0.01em;}
.app-header-sub{font-size:0.72rem;opacity:0.7;margin-top:0.1rem;}
/* ── KPI strip ── */
.kpi-strip{background:#fff;border-bottom:1px solid #dde3ea;display:flex;flex-shrink:0;}
.kpi-card{flex:1;padding:0.55rem 0.9rem;border-right:1px solid #e8edf3;text-align:center;}
.kpi-card:last-child{border-right:none;}
.kpi-value{font-size:1.45rem;font-weight:800;line-height:1;}
/* ── Section stat chips ── */
.section-stats{display:flex;flex-wrap:wrap;gap:0.5rem;margin:0 0 1.25rem;}
.stat-chip{background:#f5f7fa;border:1px solid #dde3ea;border-radius:6px;padding:0.45rem 0.85rem;text-align:center;min-width:88px;}
.stat-chip-value{font-size:1.1rem;font-weight:700;line-height:1.2;}
.stat-chip-label{font-size:0.72rem;color:#666;margin-top:2px;}
.stat-chip.ok .stat-chip-value{color:#2e7d32;} .stat-chip.warn .stat-chip-value{color:#e65100;} .stat-chip.critical .stat-chip-value{color:#b71c1c;} .stat-chip.neutral .stat-chip-value{color:#1565c0;}
/* ── Compliance Overview ── */
.cov-bar-wrap{margin:0.5rem 0 0.85rem;}
.cov-bar{display:flex;height:10px;border-radius:5px;overflow:hidden;}
.cov-bar-pass{background:#2e7d32;} .cov-bar-warn{background:#f9a825;} .cov-bar-fail{background:#c62828;} .cov-bar-na{background:#e0e0e0;}
.cov-legend{display:flex;flex-wrap:wrap;gap:0.8rem;font-size:0.82rem;color:#555;margin-top:0.4rem;}
.cov-legend-item{display:flex;align-items:center;gap:0.35rem;}
.cov-legend-dot{width:10px;height:10px;border-radius:50%;flex-shrink:0;}
/* ── Technical Issues ── */
.issue-sev-critical{color:#b71c1c;font-weight:bold;} .issue-sev-warning{color:#e65100;font-weight:bold;} .issue-sev-info{color:#1565c0;}
.kpi-label{font-size:0.67rem;color:#64748b;margin-top:0.2rem;}
.kpi-sub{font-size:0.63rem;color:#94a3b8;margin-top:0.05rem;}
.kpi-value.ok{color:#16a34a;} .kpi-value.warn{color:#d97706;} .kpi-value.critical{color:#dc2626;}
/* ── Layout ── */
.layout{display:flex;flex:1;}
/* ── Sidebar ── */
.sidebar{width:208px;background:#1e293b;color:#94a3b8;display:flex;flex-direction:column;flex-shrink:0;overflow-x:hidden;position:sticky;top:0;height:100vh;overflow-y:auto;}
.sidebar::-webkit-scrollbar{width:3px;} .sidebar::-webkit-scrollbar-thumb{background:#334155;}
.sb-section-label{padding:0.85rem 1rem 0.3rem;font-size:0.62rem;font-weight:700;letter-spacing:0.12em;text-transform:uppercase;color:#475569;}
.sb-item{display:flex;align-items:center;gap:0.5rem;padding:0.46rem 1rem;font-size:0.81rem;color:#94a3b8;text-decoration:none;border-left:3px solid transparent;transition:background 0.12s,color 0.12s,border-color 0.12s;white-space:nowrap;cursor:pointer;}
.sb-item:hover{background:#263548;color:#e2e8f0;border-left-color:#3b82f6;}
.sb-item.active{background:#172034;color:#93c5fd;border-left-color:#3b82f6;font-weight:600;}
.sb-dot{width:8px;height:8px;border-radius:50%;flex-shrink:0;}
.dot-ok{background:#22c55e;} .dot-warn{background:#f59e0b;} .dot-critical{background:#ef4444;} .dot-neutral{background:#475569;}
.sb-badge{margin-left:auto;font-size:0.67rem;font-weight:700;background:rgba(239,68,68,0.25);color:#fca5a5;border-radius:999px;padding:1px 6px;}
.sb-badge.warn{background:rgba(245,158,11,0.25);color:#fcd34d;}
.sb-divider{margin:0.4rem 1rem;border:none;border-top:1px solid #263548;}
/* ── Main content ── */
.main{flex:1;min-width:0;}
.content-area{padding:0.9rem 1.1rem;}
/* ── Company card ── */
.company-card{background:#fff;border:1px solid #dde3ea;border-radius:8px;padding:0.7rem 1rem;margin-bottom:0.8rem;}
.company-card h2{font-size:0.95rem;font-weight:700;color:#0f2744;margin-bottom:0.45rem;padding-bottom:0.35rem;border-bottom:1px solid #f0f4f8;}
.company-info-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:0.3rem 1.25rem;}
.ci-field{display:flex;flex-direction:column;}
.ci-key{font-size:0.67rem;font-weight:700;color:#64748b;text-transform:uppercase;letter-spacing:0.06em;}
.ci-val{font-size:0.82rem;color:#1e293b;margin-top:0.08rem;}
.company-info-severity.critical{color:#c00;font-weight:600;} .company-info-severity.warning{color:#805500;font-weight:600;}
.company-info-lines{display:flex;flex-direction:column;gap:0.25rem;}
.company-info-line{display:grid;grid-template-columns:minmax(0,1fr) auto;align-items:baseline;column-gap:1.5rem;}
.company-info-meta{color:#666;font-size:0.82rem;text-align:right;white-space:nowrap;}
/* ── Action items ── */
.ai-section{border:1px solid #e2e8f0;border-radius:8px;margin-bottom:0.9rem;overflow:hidden;}
.ai-section-hdr{display:flex;align-items:center;justify-content:space-between;padding:0.5rem 0.85rem;background:#f8fafc;cursor:pointer;user-select:none;border-bottom:1px solid #e2e8f0;}
.ai-section-hdr:hover{background:#f1f5f9;}
.ai-section-title{font-weight:700;font-size:0.82rem;letter-spacing:0.02em;color:#1e293b;}
.ai-section-body{padding:0.65rem;}
.ai-grid{display:grid;grid-template-columns:1fr 1fr;gap:0.75rem;margin-bottom:0;}
.ai-panel{border-radius:8px;overflow:hidden;border:1px solid;}
.ai-panel.critical{border-color:#fca5a5;} .ai-panel.warning{border-color:#fcd34d;}
.ai-panel-header{padding:0.48rem 0.85rem;font-weight:700;font-size:0.79rem;letter-spacing:0.02em;}
.ai-panel.critical .ai-panel-header{background:#fee2e2;color:#991b1b;border-bottom:1px solid #fca5a5;}
.ai-panel.warning  .ai-panel-header{background:#fef3c7;color:#92400e;border-bottom:1px solid #fcd34d;}
.ai-panel-body{background:#fff;}
.ai-row{padding:0.52rem 0.85rem;border-bottom:1px solid #f8fafc;}
.ai-row:last-child{border-bottom:none;}
.ai-cat{font-size:0.69rem;font-weight:700;color:#64748b;text-transform:uppercase;letter-spacing:0.04em;margin-bottom:0.12rem;}
.ai-text{font-size:0.84rem;color:#1e293b;line-height:1.45;}
.ai-doc{font-size:0.73rem;color:#2563eb;text-decoration:none;display:inline-block;margin-top:0.12rem;}
.ai-doc:hover{text-decoration:underline;}
.ai-none{color:#16a34a;font-weight:600;padding:0.5rem 0;font-size:0.88rem;}
/* ── Module sections ── */
.module{background:#fff;border:1px solid #dde3ea;border-radius:8px;margin-bottom:0.8rem;overflow:hidden;scroll-margin-top:0.6rem;}
.module-hdr{display:flex;align-items:center;padding:0.65rem 1rem;background:linear-gradient(180deg,#f8fafc 0%,#f1f5f9 100%);border-bottom:1px solid #e2e8f0;cursor:pointer;user-select:none;gap:0.55rem;}
.module-hdr:hover{background:linear-gradient(180deg,#f1f5f9 0%,#e8edf4 100%);}
.module-title{flex:1;font-weight:700;font-size:0.9rem;color:#1e293b;}
.section-version{font-size:0.74rem;font-weight:600;border-radius:999px;padding:0.15rem 0.5rem;background:#fff;border:1px solid #b6c4d4;color:#44566c;white-space:nowrap;}
.module-toggle{font-size:0.78rem;color:#64748b;transition:transform 0.18s;flex-shrink:0;}
.module-toggle.open{transform:rotate(90deg);}
.module-body{padding:0.8rem 1rem;}
/* ── Raw CSV files ── */
.raw-files{margin-top:0.9rem;padding-top:0.65rem;border-top:1px dashed #d6dee8;}
.raw-files>summary{cursor:pointer;list-style:none;position:relative;font-size:0.82rem;font-weight:600;padding:0.5rem 2rem 0.5rem 0.8rem;border-radius:8px;background:#f7f9fb;color:#516375;}
.raw-files>summary::-webkit-details-marker{display:none;}
.raw-files>summary::after{content:'▸';position:absolute;right:0.8rem;top:50%;transform:translateY(-50%);font-size:0.82rem;transition:transform 0.16s ease;}
.raw-files[open]>summary::after{transform:translateY(-50%) rotate(90deg);}
.raw-files-list{margin:0.4rem 0 0 1rem;padding-left:0.5rem;font-size:0.82rem;}
.raw-files-list li{margin:0.18rem 0;}
/* ── Tables ── */
table{border-collapse:collapse;width:100%;}
th,td{border:1px solid #dde3ea;padding:5px 9px;text-align:left;font-size:0.87rem;}
th{background:#f1f5f9;color:#334155;font-weight:700;font-size:0.79rem;}
tr:nth-child(even){background:#fafcff;}
/* ── Utility ── */
.ok{color:green;font-weight:bold;} .warn{color:darkorange;font-weight:bold;} .critical{color:red;font-weight:bold;}
.mfa-miss{background-color:#ffcccc;}
.expand-hint{display:inline-flex;align-items:center;gap:0.45rem;margin:0 0 0.6rem;padding:0.38rem 0.75rem;font-size:0.82rem;font-weight:600;color:#4d6278;background:#f5f8fb;border:1px solid #d8e2ec;border-radius:999px;}
.expand-hint::before{content:'▸';font-size:0.72rem;color:#6b7f95;}
.user-row{cursor:pointer;}
.user-row td{position:relative;transition:background-color 0.16s ease,border-color 0.16s ease;}
.user-row>td:first-child{padding-left:1.9rem;}
.user-row>td:first-child::before{content:'▸';position:absolute;left:0.7rem;top:50%;transform:translateY(-50%);font-size:0.78rem;color:#6a7d91;transition:transform 0.16s ease,color 0.16s ease;}
.user-row:hover td{background-color:#eaf3fb !important;}
.user-row.expanded td{background-color:#dbe9f6 !important;border-color:#c6d7e7;}
.user-row.expanded>td:first-child::before{transform:translateY(-50%) rotate(90deg);color:#2f4b67;}
.signin-detail{background:transparent !important;}
.signin-detail>td{padding:0.75rem 1rem !important;background:linear-gradient(180deg,#f8fbff 0%,#f1f6fb 100%) !important;border-top:none;border-bottom:1px solid #d5dfeb;}
.inner-table{width:100%;border-collapse:separate;border-spacing:0;font-size:0.85rem;margin:0;border:1px solid #cbd8e5;border-radius:8px;overflow:hidden;background:#fff;}
.inner-table th{background:#ddeaf6;color:#24384d;}
.inner-table td,.inner-table th{border-right:1px solid #cbd8e5;border-bottom:1px solid #cbd8e5;padding:5px 8px;}
.inner-table th:last-child,.inner-table td:last-child{border-right:none;}
.inner-table tbody tr:last-child td{border-bottom:none;}
.inner-table tr:nth-child(even) td{background:#fbfdff;}
.detail-empty{color:#6b7280;font-style:italic;}
.signin-ok{color:green;font-weight:bold;} .signin-fail{color:red;font-weight:bold;}
.size-warn{background-color:#fff3cd;} .size-critical{background-color:#ffcccc;}
code{background:#f1f5f9;border-radius:3px;padding:1px 4px;font-family:Consolas,monospace;font-size:0.82rem;color:#334155;}
/* ── Module body spacing & subsection headers ── */
.module-body h4{margin:1.1rem 0 0.4rem;font-size:0.88rem;font-weight:700;color:#1e293b;padding-top:0.9rem;border-top:1px solid #e8edf3;}
.module-body h4:first-child{border-top:none;padding-top:0;margin-top:0;}
.module-body details{margin:0;}
.module-body details > summary{display:block;margin:1.1rem 0 0.4rem;font-size:0.88rem;font-weight:700;color:#1e293b;padding:0.9rem 1.4rem 0 0;border-top:1px solid #e8edf3;cursor:pointer;list-style:none;position:relative;}
.module-body details > summary::-webkit-details-marker,.module-body details > summary::marker{display:none;}
.module-body details > summary::after{content:'▾';position:absolute;right:0;top:0.85rem;color:#94a3b8;font-size:0.75rem;}
.module-body details:not([open]) > summary::after{content:'▸';}
.module-body > details:first-of-type > summary{border-top:none;padding-top:0;margin-top:0;}
.module-body p{margin-bottom:0.3rem;line-height:1.55;}
.module-body p:last-child{margin-bottom:0;}
.module-body details{margin:0.3rem 0;}
.module-body table{margin-bottom:0.5rem;}
td ul, td ol{padding-left:1.4em;overflow-wrap:break-word;}
/* ── Sidebar sub-items ── */
.sb-module-group{display:flex;flex-direction:column;}
.sb-sub-group{border-left:1px solid #253245;margin-left:1.45rem;display:flex;flex-direction:column;padding-bottom:0.3rem;}
.sb-sub{display:block;padding:0.27rem 0.65rem;font-size:0.76rem;color:#4e6888;text-decoration:none;border-left:2px solid transparent;transition:color 0.12s,border-color 0.12s;white-space:nowrap;}
.sb-sub:hover{color:#94a3b8;border-left-color:#475569;}
.sb-sub.active{color:#93c5fd;border-left-color:#3b82f6;font-weight:600;}
/* ── App header company info ── */
.app-hdr-company{font-size:1.05rem;font-weight:700;letter-spacing:0.01em;text-align:center;}
.app-hdr-right{text-align:right;font-size:0.72rem;opacity:0.7;}
/* ── Stat chips ── */
a.stat-chip{text-decoration:none;display:block;}
a.stat-chip:hover{background:#edf2f7;border-color:#b0bec5;box-shadow:0 1px 4px rgba(0,0,0,0.08);}
div.stat-chip[data-scuba-filter]{cursor:pointer;transition:opacity 0.15s,border-color 0.15s,box-shadow 0.15s;user-select:none;}
div.stat-chip[data-scuba-filter]:hover{background:#edf2f7;border-color:#b0bec5;box-shadow:0 1px 4px rgba(0,0,0,0.08);}
div.stat-chip[data-scuba-filter].inactive{opacity:0.3;}
</style>
</head>
"@)


# =========================================
# ===   Company Summary                 ===
# =========================================
$script:TechnicalContactSeverity = $null

$orgInfoPath = Join-Path $AuditFolder "OrgInfo.json"
if (Test-Path $orgInfoPath) {
    $orgInfo = Get-Content $orgInfoPath -Raw | ConvertFrom-Json

    # Address
    $addrParts = @($orgInfo.Raw.Street, $orgInfo.Raw.City, $orgInfo.Raw.State, $orgInfo.Raw.PostalCode, $orgInfo.CountryLetterCode)
    $address   = ($addrParts | Where-Object { $_ }) -join ", "

    # Phone and technical contact
    $phone       = if ($orgInfo.Raw.BusinessPhones) { ($orgInfo.Raw.BusinessPhones -join ", ") } else { $null }
    $techContact = if ($orgInfo.TechnicalNotificationMails.Count -gt 0) { $orgInfo.TechnicalNotificationMails -join ", " } else { "—" }
    if ($_mspDomains.Count -gt 0) {
        $_foreignContacts = @($orgInfo.TechnicalNotificationMails | Where-Object {
            $domain = ($_ -split '@')[-1].ToLower()
            $domain -notin $_mspDomains
        })
        if ($_foreignContacts.Count -gt 0) {
            $script:TechnicalContactSeverity = 'critical'
        }
    }
    $techContactHtml = if ($script:TechnicalContactSeverity) {
        "<span class='company-info-severity $($script:TechnicalContactSeverity)'>$(ConvertTo-HtmlText $techContact)</span>"
    }
    else {
        ConvertTo-HtmlText $techContact
    }

    # Azure AD Sync status (also satisfies the "Azure AD Sync Health" checklist item)
    $syncEnabled = $orgInfo.Raw.OnPremisesSyncEnabled
    $syncErrors  = if ($orgInfo.Raw.OnPremisesProvisioningErrors) { @($orgInfo.Raw.OnPremisesProvisioningErrors).Count } else { 0 }
    if ($syncEnabled) {
        $lastSyncDt  = [datetime]$orgInfo.Raw.OnPremisesLastSyncDateTime
        $hoursSince  = [math]::Round(([datetime]::UtcNow - $lastSyncDt.ToUniversalTime()).TotalHours, 1)
        $lastSyncFmt = $lastSyncDt.ToUniversalTime().ToString("yyyy-MM-dd HH:mm") + " UTC"
        $syncClass   = if ($hoursSince -gt 24) { "critical" } elseif ($hoursSince -gt 4) { "warn" } else { "ok" }
        $syncCell    = "<span class='$syncClass'>Enabled &mdash; last sync $lastSyncFmt ($hoursSince h ago)</span>"
    } else {
        $syncCell    = $null
    }

    # Verified domains — exclude the internal EOP routing domain (*.mail.onmicrosoft.com)
    $domains    = @($orgInfo.VerifiedDomains | Where-Object { $_.Name -notlike "*.mail.onmicrosoft.com" })
    $domainRows = foreach ($d in $domains) {
        $dtype = if ($d.IsInitial) { "*.onmicrosoft.com" } elseif ($d.IsDefault) { "Default" } else { "Custom" }
        $mark = if ($d.IsDefault) { " <b>(default)</b>" } else { "" }
        $typeHtml = if ($dtype) { "<span class='company-info-meta'>$([System.Net.WebUtility]::HtmlEncode($dtype))</span>" } else { "" }
        "<div class='company-info-line'><span>$(ConvertTo-HtmlText $d.Name)$mark</span>$typeHtml</div>"
    }
    $domainsHtml = if ($domainRows.Count -gt 0) { "<div class='company-info-lines'>$($domainRows -join '')</div>" } else { '&mdash;' }

    $_addrField  = if ($address) { "<div class='ci-field'><span class='ci-key'>Address</span><span class='ci-val'>$(ConvertTo-HtmlText $address)</span></div>" } else { "" }
    $_phoneField = if ($phone)   { "<div class='ci-field'><span class='ci-key'>Phone</span><span class='ci-val'>$(ConvertTo-HtmlText $phone)</span></div>" } else { "" }

    if ($syncEnabled) {
        $syncField = "<div class='ci-field'><span class='ci-key'>Azure AD Sync</span><span class='ci-val'>$syncCell</span></div>"
        $errField  = if ($syncErrors -gt 0) { "<div class='ci-field'><span class='ci-key'>Sync Errors</span><span class='ci-val'><span class='critical'>$syncErrors error(s)</span></span></div>" } else { "" }
    } else {
        $syncField = "<div class='ci-field'><span class='ci-key'>Azure AD Sync</span><span class='ci-val'>Not configured (cloud-only)</span></div>"
        $errField  = ""
    }

    $script:companyCardHtml = @"
<div class='company-card'>
  <div class='company-info-grid'>
    <div class='ci-field'><span class='ci-key'>Tenant ID</span><span class='ci-val'><code>$($orgInfo.Id)</code></span></div>
    <div class='ci-field'><span class='ci-key'>Technical Contact</span><span class='ci-val'>$techContactHtml</span></div>
    $syncField
    <div class='ci-field'><span class='ci-key'>Domains</span><span class='ci-val'>$domainsHtml</span></div>
    $_addrField
    $_phoneField
    $errField
  </div>
</div>
"@

    $html.Add(@"
<body>
<div class='page-wrapper'>
<div class='sticky-header'>
<div class='app-header'>
  <div></div>
  <div class='app-hdr-company'>$(ConvertTo-HtmlText $orgInfo.DisplayName)</div>
  <div class='app-hdr-right'>Generated: $reportDate</div>
</div>
"@)
}
else {
    $script:companyCardHtml = ''
    $html.Add(@"
<body>
<div class='page-wrapper'>
<div class='sticky-header'>
<div class='app-header'>
  <div></div>
  <div class='app-hdr-company'>Microsoft 365 Audit</div>
  <div class='app-hdr-right'>Generated: $reportDate</div>
</div>
"@)
}


# =========================================
# ===   Action Items                    ===
# =========================================
$actionItems = [System.Collections.Generic.List[hashtable]]::new()
$script:ActionItemSequence = 0

# Load ScubaGear results if a prior run produced output in Raw\ScubaGear_*\
$_scubaResults  = $null
$_scubaHtmlPath = $null
$_scubaRunDir   = Get-ChildItem -Path (Join-Path $AuditFolder 'Raw') -Directory -Filter 'ScubaGear_*' `
    -ErrorAction SilentlyContinue | Sort-Object Name -Descending | Select-Object -First 1
if ($_scubaRunDir) {
    $_scubaJsonFile = Get-ChildItem -Path $_scubaRunDir.FullName -Filter 'ScubaResults_*.json' `
        -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($_scubaJsonFile) {
        try {
            $_scubaResults = Get-Content $_scubaJsonFile.FullName -Raw | ConvertFrom-Json
            Write-Verbose "ScubaGear results loaded: $($_scubaJsonFile.FullName)"
        } catch {
            Write-Warning "Could not parse ScubaGear results: $($_.Exception.Message)"
        }
    }
    $_scubaHtmlPath = Join-Path $_scubaRunDir.FullName 'BaselineReports.html'
    if (-not (Test-Path $_scubaHtmlPath)) { $_scubaHtmlPath = $null }
}

# Helper: add an action item
# Severity: 'critical' | 'warning'
function Add-ActionItem {
    param([string]$Severity, [string]$Category, [string]$Text, [string]$DocUrl = "", [string]$CheckId = "")
    $script:ActionItemSequence++
    $script:actionItems.Add(@{
        Severity = $Severity
        Category = $Category
        Text     = $Text
        DocUrl   = $DocUrl
        CheckId  = $CheckId
        Sequence = $script:ActionItemSequence
    })
}

function Get-ActionItemModuleSortOrder {
    [CmdletBinding()]
    param(
        [string]$Category
    )

    $moduleName = (($Category -split '/', 2)[0]).Trim()
    switch ($moduleName) {
        'Entra'         { return 10 }
        'Exchange'      { return 20 }
        'SharePoint'    { return 30 }
        'Mail Security' { return 40 }
        'Intune'        { return 50 }
        'Teams'         { return 60 }
        default         { return 90 }
    }
}

# --- Audit certificate expiry ---
if ($CertExpiryDays -ge 0 -and $CertExpiryDays -le 30) {
    Add-ActionItem -Severity 'warning' -Category 'Toolkit / Certificate' -Text "Audit app certificate expires in $CertExpiryDays day(s). Run Setup-365AuditApp.ps1 -Force (requires interactive Global Admin login) to renew before the next audit run." -CheckId 'TK-CERT-001'
}

# --- Technical Contact domain check ---
# Flags any technical notification email that is not from an MSP domain (MspDomains in config.psd1).
# A foreign address likely belongs to a previous MSP and should be removed.
if ($_mspDomains.Count -eq 0) {
    Write-Warning "MspDomains is not configured in config.psd1 — Technical Contact domain check skipped."
}
elseif ($orgInfo -and $orgInfo.TechnicalNotificationMails.Count -gt 0) {
    $_foreignContacts = @($orgInfo.TechnicalNotificationMails | Where-Object {
        $domain = ($_ -split '@')[-1].ToLower()
        $domain -notin $_mspDomains
    })
    if ($_foreignContacts.Count -gt 0) {
        $_contactList = $_foreignContacts -join ', '
        Add-ActionItem -Severity 'critical' -Category 'Tenant / Technical Contact' `
            -Text "Technical Contact address(es) are not from a recognised MSP domain: $_contactList — this may be a previous MSP's details still on the tenant. Review and update the Technical Notification email in the Microsoft 365 admin centre (Settings &rarr; Org settings &rarr; Organisation profile)." `
            -CheckId 'TENANT-CONTACT-001'
    }
}

# --- Entra checks ---

# Guest accounts in privileged roles
# #EXT# in the UPN identifies a guest account. Guests holding admin roles may be
# a previous MSP's staff who retained access after the engagement ended.
$_aiAdminRolesCsv = Join-Path $entraDir "Entra_AdminRoles.csv"
if (Test-Path $_aiAdminRolesCsv) {
    $_guestAdmins = @(Import-Csv $_aiAdminRolesCsv | Where-Object { $_.MemberUserPrincipalName -like '*#EXT#*' })
    if ($_guestAdmins.Count -gt 0) {
        $_guestList = ($_guestAdmins | ForEach-Object { "$($_.MemberDisplayName) &mdash; $($_.RoleName)" }) -join '<br>'
        Add-ActionItem -Severity 'critical' -Category 'Entra / Admins' `
            -Text "Guest account(s) hold privileged admin roles — this may indicate a previous MSP's accounts still have admin access. Review and remove if no longer required:<br>$_guestList" `
            -DocUrl 'https://learn.microsoft.com/en-us/entra/identity/role-based-access-control/best-practices' `
            -CheckId 'ADMIN-GUEST-ROLE-001'
    }
}

# MFA coverage
$_aiUsersCsv = Join-Path $entraDir "Entra_Users.csv"
if (Test-Path $_aiUsersCsv) {
    $_aiUsers   = @(Import-Csv $_aiUsersCsv)
    $_aiTotal   = $_aiUsers.Count
    $_aiEnabled = ($_aiUsers | Where-Object { $_.MFAEnabled -eq 'True' }).Count
    $_aiPct     = if ($_aiTotal -gt 0) { [math]::Round(($_aiEnabled / $_aiTotal) * 100, 1) } else { 100 }
    if ($_aiPct -lt 100) {
        $missing = $_aiTotal - $_aiEnabled
        Add-ActionItem -Severity 'critical' -Category 'Entra / MFA' -Text "MFA not enabled for $missing of $_aiTotal licensed users (${_aiPct}%). Essential Eight: Restrict privileged access — all users must have MFA. (CIS 2.2.2)" -DocUrl 'https://learn.microsoft.com/en-us/entra/identity/authentication/concept-mfa-howitworks' -CheckId 'MFA-USERS-001'
    }
}

# Security Defaults + Conditional Access enforcement
$_aiSdCsv = Join-Path $entraDir "Entra_SecurityDefaults.csv"
$_aiCaCsv = Join-Path $entraDir "Entra_CA_Policies.csv"
$_aiSdEnabled = $false
if (Test-Path $_aiSdCsv) {
    $_aiSd = Import-Csv $_aiSdCsv | Select-Object -First 1
    $_aiSdEnabled = ($_aiSd.SecurityDefaultsEnabled -eq "True")
}
if (-not $_aiSdEnabled -and (Test-Path $_aiCaCsv)) {
    $_aiCaPolicies = @(Import-Csv $_aiCaCsv)
    $_aiEnabledCa  = @($_aiCaPolicies | Where-Object { $_.State -eq "enabled" })
    if ($_aiEnabledCa.Count -eq 0 -and $_aiCaPolicies.Count -eq 0) {
        Add-ActionItem -Severity 'critical' -Category 'Entra / MFA' -Text "Security Defaults are disabled and no Conditional Access policies exist. MFA is not enforced for any user. (CIS 2.2.2)" -DocUrl 'https://learn.microsoft.com/en-us/entra/fundamentals/security-defaults' -CheckId 'MFA-NODEFAULTS-001'
    }
    elseif ($_aiEnabledCa.Count -eq 0) {
        $_reportOnly = @($_aiCaPolicies | Where-Object { $_.State -eq "enabledForReportingButNotEnforced" }).Count
        Add-ActionItem -Severity 'critical' -Category 'Entra / CA' -Text "Security Defaults disabled and no CA policies are in 'Enabled' state ($_reportOnly in report-only). MFA is not enforced. (CIS 2.2.2)" -DocUrl 'https://learn.microsoft.com/en-us/entra/identity/conditional-access/overview' -CheckId 'CA-NOTENFORCED-001'
    }
    elseif (($_aiCaPolicies | Where-Object { $_.State -eq "enabledForReportingButNotEnforced" }).Count -gt 0) {
        $_roCount = ($_aiCaPolicies | Where-Object { $_.State -eq "enabledForReportingButNotEnforced" }).Count
        Add-ActionItem -Severity 'warning' -Category 'Entra / CA' -Text "$_roCount Conditional Access policy/policies are in report-only mode and not enforcing controls. Review and enable when ready." -DocUrl 'https://learn.microsoft.com/en-us/entra/identity/conditional-access/concept-conditional-access-report-only' -CheckId 'CA-REPORTONLY-001'
    }
}

# Global Admin count
$_aiGaCsv = Join-Path $entraDir "Entra_GlobalAdmins.csv"
if (Test-Path $_aiGaCsv) {
    $_aiGaCount = @(Import-Csv $_aiGaCsv).Count
    if ($_aiGaCount -eq 0) {
        Add-ActionItem -Severity 'critical' -Category 'Entra / Admins' -Text "No Global Administrators found — this may indicate a data collection issue." -CheckId 'ADMIN-NOGA-001'
    }
    elseif ($_aiGaCount -eq 1) {
        Add-ActionItem -Severity 'warning' -Category 'Entra / Admins' -Text "Only 1 Global Administrator account. Recommend at least 2 for resilience (break-glass scenario). (CIS 1.1.3)" -DocUrl 'https://learn.microsoft.com/en-us/entra/identity/role-based-access-control/best-practices' -CheckId 'ADMIN-SINGLE-001'
    }
    elseif ($_aiGaCount -gt 4) {
        Add-ActionItem -Severity 'warning' -Category 'Entra / Admins' -Text "$_aiGaCount Global Administrator accounts. Microsoft recommends 2–4 max. Essential Eight: Restrict administrative privileges. (CIS 1.1.3)" -DocUrl 'https://learn.microsoft.com/en-us/entra/identity/role-based-access-control/best-practices' -CheckId 'ADMIN-EXCESS-001'
    }
}

# SSPR
$_aiSsprCsv = Join-Path $entraDir "Entra_SSPR.csv"
if (Test-Path $_aiSsprCsv) {
    $_aiSspr = Import-Csv $_aiSsprCsv | Select-Object -First 1
    if ($_aiSspr.SSPREnabled -ne "Enabled") {
        Add-ActionItem -Severity 'warning' -Category 'Entra / SSPR' -Text "Self-Service Password Reset is not fully enabled (current: $($_aiSspr.SSPREnabled)). Users cannot reset passwords without helpdesk intervention." -DocUrl 'https://learn.microsoft.com/en-us/entra/identity/authentication/concept-sspr-howitworks' -CheckId 'SSPR-DISABLED-001'
    }
}

# Third-party enterprise apps with admin-consented permissions
# Apps with admin consent have been granted API access to the tenant. Review for unknown or MSP-installed apps.
$_aiAppsCsv = Join-Path $entraDir "Entra_EnterpriseApps.csv"
if (Test-Path $_aiAppsCsv) {
    $_aiApps         = @(Import-Csv $_aiAppsCsv -Encoding UTF8)
    $_aiConsentedApps = @($_aiApps | Where-Object { $_.AdminConsented -eq 'True' })
    if ($_aiConsentedApps.Count -gt 0) {
        $_appList = ($_aiConsentedApps | ForEach-Object { "$($_.DisplayName) ($($_.PublisherName))" }) -join '<br>'
        Add-ActionItem -Severity 'warning' -Category 'Entra / Enterprise Apps' `
            -Text "$($_aiConsentedApps.Count) third-party app(s) have admin-consented API permissions. Review to confirm all are authorised and none were installed by a previous MSP:<br>$_appList" `
            -DocUrl 'https://learn.microsoft.com/en-us/entra/identity/enterprise-apps/manage-consent-requests' `
            -CheckId 'APPS-CONSENT-001'
    }

    # AvePoint SaaS backup detection — AvePoint registers service principals in the tenant when configured
    $_avePoint = @($_aiApps | Where-Object { $_.DisplayName -like 'AvePoint*' })
    if ($_avePoint.Count -eq 0) {
        Add-ActionItem -Severity 'warning' -Category 'Entra / Backup' `
            -Text "No AvePoint service principal detected in this tenant. Confirm that SaaS backup (Microsoft 365 data) is configured and active for this customer." `
            -DocUrl 'https://partner.avepointonlineservices.com/Dashboard#/directory' `
            -CheckId 'APPS-NOBACKUP-001'
    }
}

# Identity Protection — risky users
$_aiRiskyUsersCsv = Join-Path $entraDir "Entra_RiskyUsers.csv"
if (Test-Path $_aiRiskyUsersCsv) {
    $_aiRiskyUsers = @(Import-Csv $_aiRiskyUsersCsv | Where-Object { $_.RiskState -in @('atRisk','confirmedCompromised') -and $_.RiskLevel -in @('medium','high') })
    if ($_aiRiskyUsers.Count -gt 0) {
        $_ruList = ($_aiRiskyUsers | ForEach-Object { "$(ConvertTo-HtmlText $_.UserPrincipalName) — $($_.RiskLevel) ($($_.RiskState))" }) -join '<br>'
        Add-ActionItem -Severity 'critical' -Category 'Entra / Identity Protection' `
            -Text "$($_aiRiskyUsers.Count) user(s) flagged as at-risk or compromised by Entra Identity Protection. Investigate and remediate immediately:<br>$_ruList" `
            -DocUrl 'https://learn.microsoft.com/en-us/entra/id-protection/howto-identity-protection-investigate-risk' `
            -CheckId 'IDP-RISKYUSERS-001'
    }
}

# Admin role holders without MFA (cross-reference AdminRoles + Users)
$_aiAdminRolesCsv2 = Join-Path $entraDir "Entra_AdminRoles.csv"
if ((Test-Path $_aiAdminRolesCsv2) -and (Test-Path $_aiUsersCsv)) {
    $_adminUpns  = @(Import-Csv $_aiAdminRolesCsv2 | Select-Object -ExpandProperty MemberUserPrincipalName -Unique)
    $_adminNoMfa = @($_aiUsers | Where-Object { $_.UPN -in $_adminUpns -and $_.MFAEnabled -eq 'False' })
    if ($_adminNoMfa.Count -gt 0) {
        $_adminList = ($_adminNoMfa | ForEach-Object { ConvertTo-HtmlText $_.UPN }) -join '<br>'
        Add-ActionItem -Severity 'critical' -Category 'Entra / Admins' `
            -Text "$($_adminNoMfa.Count) admin role holder(s) do not have MFA registered. Privileged accounts without MFA are the single highest-risk attack vector for account takeover:<br>$_adminList" `
            -DocUrl 'https://learn.microsoft.com/en-us/entra/identity/role-based-access-control/best-practices' `
            -CheckId 'ADMIN-NOMFA-001'
    }
}

# No CA policy enforces MFA for all users
if (-not $_aiSdEnabled -and $null -ne $_aiCaPolicies -and @($_aiEnabledCa).Count -gt 0) {
    $_aiMfaAllUsers = @($_aiEnabledCa | Where-Object { $_.RequiresMFA -eq 'True' -and $_.IncludeUsers -match 'All users' })
    if ($_aiMfaAllUsers.Count -eq 0) {
        Add-ActionItem -Severity 'critical' -Category 'Entra / CA' `
            -Text "No enabled Conditional Access policy requires MFA for all users. Existing policies may only cover a subset of users or applications, leaving gaps in MFA enforcement." `
            -DocUrl 'https://learn.microsoft.com/en-us/entra/identity/conditional-access/howto-conditional-access-policy-all-users-mfa' `
            -CheckId 'CA-NOMFAREQUIRED-001'
    }
}

# App registration credentials expired or expiring within 30 days
$_aiAppRegCsv = Join-Path $entraDir "Entra_AppRegistrations.csv"
if (Test-Path $_aiAppRegCsv) {
    $_aiAppRegs = @(Import-Csv $_aiAppRegCsv | Where-Object { $_.CredentialType -ne 'None' -and $_.DaysUntilExpiry -ne '' })
    $_aiExpired  = @($_aiAppRegs | Where-Object { [int]$_.DaysUntilExpiry -lt 0 })
    $_aiExpiring = @($_aiAppRegs | Where-Object { [int]$_.DaysUntilExpiry -ge 0 -and [int]$_.DaysUntilExpiry -le 30 })
    if ($_aiExpired.Count -gt 0) {
        $_expList = ($_aiExpired | ForEach-Object { "$(ConvertTo-HtmlText $_.DisplayName) — $($_.CredentialType) ($($_.CredentialName)) expired $([Math]::Abs([int]$_.DaysUntilExpiry))d ago" }) -join '<br>'
        Add-ActionItem -Severity 'critical' -Category 'Entra / App Registrations' `
            -Text "$($_aiExpired.Count) app registration credential(s) have expired. Applications using these will fail to authenticate:<br>$_expList" `
            -DocUrl 'https://learn.microsoft.com/en-us/entra/identity/enterprise-apps/manage-certificates-for-federated-single-sign-on' `
            -CheckId 'APPS-EXPIREDCRED-001'
    }
    if ($_aiExpiring.Count -gt 0) {
        $_expList = ($_aiExpiring | ForEach-Object { "$(ConvertTo-HtmlText $_.DisplayName) — $($_.CredentialType) ($($_.CredentialName)) expires in $($_.DaysUntilExpiry)d" }) -join '<br>'
        Add-ActionItem -Severity 'warning' -Category 'Entra / App Registrations' `
            -Text "$($_aiExpiring.Count) app registration credential(s) expire within 30 days. Renew before expiry to avoid authentication failures:<br>$_expList" `
            -DocUrl 'https://learn.microsoft.com/en-us/entra/identity/enterprise-apps/manage-certificates-for-federated-single-sign-on' `
            -CheckId 'APPS-EXPIRINGCRED-001'
    }
}

# Authentication Methods — SMS/voice only, no modern method
$_aiAuthMethodsCsv = Join-Path $entraDir "Entra_AuthMethodsPolicy.csv"
if (Test-Path $_aiAuthMethodsCsv) {
    $_aiAuthMethods  = @(Import-Csv $_aiAuthMethodsCsv)
    $_aiSmsEnabled   = @($_aiAuthMethods | Where-Object { $_.MethodId -eq 'sms'   -and $_.State -eq 'enabled' })
    $_aiVoiceEnabled = @($_aiAuthMethods | Where-Object { $_.MethodId -eq 'voice' -and $_.State -eq 'enabled' })
    $_aiModernMfa    = @($_aiAuthMethods | Where-Object { $_.MethodId -in @('microsoftAuthenticator','fido2','temporaryAccessPass','softwareOath','x509Certificate') -and $_.State -eq 'enabled' })
    if (($_aiSmsEnabled.Count -gt 0 -or $_aiVoiceEnabled.Count -gt 0) -and $_aiModernMfa.Count -eq 0) {
        Add-ActionItem -Severity 'warning' -Category 'Entra / Auth Methods' `
            -Text "SMS and/or voice call MFA are enabled but no phishing-resistant or modern authentication methods (Authenticator app, FIDO2, Software OATH) are enabled. SMS and voice are vulnerable to SIM-swapping and interception. (CIS 2.3.5)" `
            -DocUrl 'https://learn.microsoft.com/en-us/entra/identity/authentication/concept-authentication-methods' `
            -CheckId 'AUTH-WEAKMFA-001'
    }
}

# External collaboration settings — guest invite too permissive
$_aiExtCollabCsv = Join-Path $entraDir "Entra_ExternalCollab.csv"
if (Test-Path $_aiExtCollabCsv) {
    $_aiExtCollab = Import-Csv $_aiExtCollabCsv | Select-Object -First 1
    if ($_aiExtCollab.AllowInvitesFrom -eq 'everyone') {
        Add-ActionItem -Severity 'warning' -Category 'Entra / External Collaboration' `
            -Text "Guest invitations are open to everyone — any user in the tenant can invite external guests. Restrict to admins or the Guest Inviter role to control external access. (CIS 1.6.3)" `
            -DocUrl 'https://learn.microsoft.com/en-us/entra/external-id/external-collaboration-settings-configure' `
            -CheckId 'EXTCOLAB-OPENINVITE-001'
    }
}

# PIM — permanent admin assignments on P2-licensed tenants
$_aiPimCsv = Join-Path $entraDir "Entra_PIMAssignments.csv"
if (Test-Path $_aiPimCsv) {
    $_aiPimAssignments = @(Import-Csv $_aiPimCsv)
    $_aiPermanent = @($_aiPimAssignments | Where-Object { ($_.EndDateTime -eq '' -or [string]::IsNullOrWhiteSpace($_.EndDateTime)) -and $_.AssignmentType -ne 'Eligible' })
    if ($_aiPermanent.Count -gt 0) {
        $_permList = ($_aiPermanent | Select-Object -First 10 | ForEach-Object { "$(ConvertTo-HtmlText $_.PrincipalDisplayName) — $($_.RoleName)" }) -join '<br>'
        if ($_aiPermanent.Count -gt 10) { $_permList += '<br>...' }
        Add-ActionItem -Severity 'warning' -Category 'Entra / PIM' `
            -Text "$($_aiPermanent.Count) permanent (non-time-bound) privileged role assignment(s) detected. With PIM licensed, convert all admin access to eligible (just-in-time) assignments:<br>$_permList" `
            -DocUrl 'https://learn.microsoft.com/en-us/entra/id-governance/privileged-identity-management/pim-configure' `
            -CheckId 'PIM-PERMANENT-001'
    }
}

# --- Exchange checks ---

# Inbox forwarding rules
$_aiInboxCsv = Join-Path $exchangeDir "Exchange_InboxForwardingRules.csv"
if (Test-Path $_aiInboxCsv) {
    $_aiInboxRules = @(Import-Csv $_aiInboxCsv)
    if ($_aiInboxRules.Count -gt 0) {
        Add-ActionItem -Severity 'warning' -Category 'Exchange / Rules' -Text "$($_aiInboxRules.Count) inbox rule(s) forward or redirect mail. Review to ensure these are authorised and not a sign of account compromise. (CIS 6.2.1)" -DocUrl 'https://learn.microsoft.com/en-us/microsoft-365/security/office-365-security/mail-flow-rules-transport-rules-0' -CheckId 'EXR-INBOXFWD-001'
    }
}

# Broken inbox rules
$_aiBrokenCsv = Join-Path $exchangeDir "Exchange_BrokenInboxRules.csv"
if (Test-Path $_aiBrokenCsv) {
    $_aiBrokenRules = @(Import-Csv $_aiBrokenCsv)
    if ($_aiBrokenRules.Count -gt 0) {
        Add-ActionItem -Severity 'warning' -Category 'Exchange / Rules' -Text "$($_aiBrokenRules.Count) inbox rule(s) are in a broken/non-functional state and are not processing mail. Edit or re-create them in Outlook." -DocUrl 'https://support.microsoft.com/en-us/office/manage-email-messages-by-using-rules-c24f5dea-9465-4df4-ad17-a50704d66c59' -CheckId 'EXR-BROKEN-001'
    }
}

# Remote domain auto-forwarding — only flag named (non-wildcard) domains; the default * entry is present in every tenant
$_aiRemoteCsv = Join-Path $exchangeDir "Exchange_RemoteDomainForwarding.csv"
if (Test-Path $_aiRemoteCsv) {
    $_aiRemoteNamed = @(Import-Csv $_aiRemoteCsv | Where-Object { $_.AutoForwardEnabled -eq "True" -and $_.DomainName -ne "*" })
    if ($_aiRemoteNamed.Count -gt 0) {
        $domainList = ($_aiRemoteNamed.DomainName -join ", ")
        Add-ActionItem -Severity 'warning' -Category 'Exchange / Forwarding' -Text "Auto-forwarding explicitly enabled for named external domain(s): $domainList. Confirm these are intentional. (CIS 6.2.1)" -DocUrl 'https://learn.microsoft.com/en-us/exchange/mail-flow-best-practices/remote-domains/remote-domains' -CheckId 'EXFWD-REMOTEDOMAIN-001'
    }
}

# Unified Audit Log / retention
$_aiAuditCfgCsv = Join-Path $exchangeDir "Exchange_AuditConfig.csv"
if (Test-Path $_aiAuditCfgCsv) {
    $_aiAuditCfg = Import-Csv $_aiAuditCfgCsv | Select-Object -First 1
    if ($_aiAuditCfg.UnifiedAuditLogIngestionEnabled -eq "False") {
        Add-ActionItem -Severity 'critical' -Category 'Exchange / Audit' -Text "Unified Audit Log ingestion is disabled. Security and compliance events are not being recorded. Enable immediately. (CIS 3.1.1)" -DocUrl 'https://learn.microsoft.com/en-us/purview/audit-log-enable-disable' -CheckId 'EXAUD-DISABLED-001'
    }
    else {
        # Check retention
        $_aiRetDays = try { [int]([TimeSpan]::Parse($_aiAuditCfg.AuditLogAgeLimit).Days) } catch { 90 }
        if ($_aiRetDays -lt 90) {
            Add-ActionItem -Severity 'warning' -Category 'Exchange / Audit' -Text "Audit log retention is only $_aiRetDays days. Microsoft recommends 90+ days; Essential Eight recommends 12 months for privileged actions." -DocUrl 'https://learn.microsoft.com/en-us/purview/audit-log-enable-disable' -CheckId 'EXAUD-RETENTION-001'
        }
    }
}

# Mailbox audit status
$_aiMbxAuditCsv = Join-Path $exchangeDir "Exchange_MailboxAuditStatus.csv"
if (Test-Path $_aiMbxAuditCsv) {
    $_aiMbxAudit   = @(Import-Csv $_aiMbxAuditCsv | Where-Object { $_.UserPrincipalName -notlike 'DiscoverySearchMailbox*' })
    $_aiAuditOff   = @($_aiMbxAudit | Where-Object { $_.AuditEnabled -eq "False" })
    if ($_aiAuditOff.Count -gt 0) {
        Add-ActionItem -Severity 'warning' -Category 'Exchange / Audit' -Text "$($_aiAuditOff.Count) mailbox(es) have per-mailbox auditing disabled. Actions in these mailboxes (logins, deletions, sends) will not be logged. (CIS 6.1.2)" -DocUrl 'https://learn.microsoft.com/en-us/purview/audit-mailboxes' -CheckId 'EXAUD-MAILBOX-001'
    }
}

# DKIM
$_aiDkimCsv = Join-Path $exchangeDir "Exchange_DKIM_Status.csv"
if (Test-Path $_aiDkimCsv) {
    $_aiDkim        = @(Import-Csv $_aiDkimCsv)
    $_aiDkimOff     = @($_aiDkim | Where-Object { $_.DKIMEnabled -ne "True" -and $_.Domain -notlike "*.onmicrosoft.com" })
    if ($_aiDkimOff.Count -gt 0) {
        $dkimDomains = ($_aiDkimOff.Domain -join ", ")
        Add-ActionItem -Severity 'warning' -Category 'Exchange / DKIM' -Text "DKIM signing not enabled on: $dkimDomains. DKIM helps prevent email spoofing and improves deliverability. (CIS 2.1.9)" -DocUrl 'https://learn.microsoft.com/en-us/microsoft-365/security/office-365-security/email-authentication-dkim-configure' -CheckId 'EXDKIM-DISABLED-001'
    }
}

# Anti-phish: Spoof Intelligence
$_aiPhishCsv = Join-Path $exchangeDir "Exchange_AntiPhishPolicies.csv"
if (Test-Path $_aiPhishCsv) {
    $_aiPhish      = @(Import-Csv $_aiPhishCsv)
    $_aiNoSpoof    = @($_aiPhish | Where-Object { $_.EnableSpoofIntelligence -eq "False" })
    if ($_aiNoSpoof.Count -gt 0) {
        Add-ActionItem -Severity 'warning' -Category 'Exchange / Anti-Phish' -Text "$($_aiNoSpoof.Count) anti-phishing policy/policies have Spoof Intelligence disabled. This reduces protection against email spoofing attacks. (CIS 2.1.7)" -DocUrl 'https://learn.microsoft.com/en-us/microsoft-365/security/office-365-security/anti-phishing-policies-about' -CheckId 'EXPHISH-SPOOF-001'
    }
}

# Mail connectors
# Custom connectors are often created by MSPs to route mail through their filtering/archiving infrastructure.
# Flag any enabled custom connector so the tech can confirm it still belongs here.
$_aiConnectorsCsv = Join-Path $exchangeDir "Exchange_MailConnectors.csv"
if (Test-Path $_aiConnectorsCsv) {
    $_aiCustomConnectors = @(Import-Csv $_aiConnectorsCsv | Where-Object {
        $_.Enabled -eq 'True' -and $_.ConnectorSource -ne 'HybridWizard' -and $_.ConnectorSource -ne 'Default'
    })
    if ($_aiCustomConnectors.Count -gt 0) {
        $_connList = ($_aiCustomConnectors | ForEach-Object { "$($_.Direction): $($_.Name) (source: $($_.ConnectorSource))" }) -join '<br>'
        Add-ActionItem -Severity 'warning' -Category 'Exchange / Connectors' `
            -Text "$($_aiCustomConnectors.Count) custom mail connector(s) are active. Connectors may route mail through a previous MSP's filtering or archiving infrastructure. Review to confirm all are still required:<br>$_connList" `
            -DocUrl 'https://learn.microsoft.com/en-us/exchange/mail-flow-best-practices/use-connectors-to-configure-mail-flow/set-up-connectors-to-route-mail' `
            -CheckId 'EXCONN-CUSTOM-001'
    }
}

# --- Mail Security checks (MailSec module) ---

$_aiDmarcCsv = Join-Path $mailSecDir "MailSec_DMARC.csv"
if (Test-Path $_aiDmarcCsv) {
    $_aiDmarc    = @(Import-Csv $_aiDmarcCsv)
    $_aiNoDmarc  = @($_aiDmarc | Where-Object { $_.Domain -notlike "*.onmicrosoft.com" -and ($_.DMARC -eq "Not Found" -or $_.DMARC -eq "" -or $null -eq $_.DMARC) })
    if ($_aiNoDmarc.Count -gt 0) {
        $dmarcDomains = ($_aiNoDmarc.Domain -join ", ")
        Add-ActionItem -Severity 'critical' -Category 'Mail Security' -Text "DMARC not configured for: $dmarcDomains. Without DMARC, spoofed email from your domain cannot be detected or rejected by recipients. (CIS 2.1.10)" -DocUrl 'https://learn.microsoft.com/en-us/microsoft-365/security/office-365-security/email-authentication-dmarc-configure' -CheckId 'MAILSEC-NODMARC-001'
    }
    # CIS 2.1.9 — DMARC exists but is in monitoring mode only (p=none)
    $_aiDmarcNone = @($_aiDmarc | Where-Object {
        $_.Domain -notlike "*.onmicrosoft.com" -and
        $_.DMARC -and $_.DMARC -ne "Not Found" -and $_.DMARC -match 'p=none'
    })
    if ($_aiDmarcNone.Count -gt 0) {
        $dmarcNoneDomains = ($_aiDmarcNone.Domain -join ", ")
        Add-ActionItem -Severity 'warning' -Category 'Mail Security' `
            -Text "DMARC is configured in monitoring mode only (p=none) for: $dmarcNoneDomains. Monitoring mode does not quarantine or reject spoofed email. Set p=quarantine or p=reject to enforce protection. (CIS 2.1.10)" `
            -DocUrl 'https://learn.microsoft.com/en-us/microsoft-365/security/office-365-security/email-authentication-dmarc-configure' `
            -CheckId 'MAILSEC-DMARCNONE-001'
    }
}

$_aiSpfCsv = Join-Path $mailSecDir "MailSec_SPF.csv"
if (Test-Path $_aiSpfCsv) {
    $_aiSpf   = @(Import-Csv $_aiSpfCsv)
    $_aiNoSpf = @($_aiSpf | Where-Object { $_.Domain -notlike "*.onmicrosoft.com" -and ($_.SPF -eq "DNS query failed" -or $_.SPF -eq "" -or $null -eq $_.SPF) })
    if ($_aiNoSpf.Count -gt 0) {
        $spfDomains = ($_aiNoSpf.Domain -join ", ")
        Add-ActionItem -Severity 'critical' -Category 'Mail Security' -Text "SPF not configured for: $spfDomains. SPF is required to identify authorised sending servers and prevent spoofing. (CIS 2.1.8)" -DocUrl 'https://learn.microsoft.com/en-us/microsoft-365/security/office-365-security/email-authentication-spf-configure' -CheckId 'MAILSEC-NOSPF-001'
    }
}

# Legacy auth / Basic Auth still enabled
$_aiLegacyAuthCsv = Join-Path $exchangeDir "Exchange_LegacyAuth.csv"
if (Test-Path $_aiLegacyAuthCsv) {
    $_aiLegacyPolicies = @(Import-Csv $_aiLegacyAuthCsv)
    $_aiBasicEnabled   = @($_aiLegacyPolicies | Where-Object {
        $_.AllowBasicAuthActiveSync -eq 'True' -or $_.AllowBasicAuthImap -eq 'True' -or
        $_.AllowBasicAuthPop -eq 'True'        -or $_.AllowBasicAuthSmtp -eq 'True' -or
        $_.AllowBasicAuthWebServices -eq 'True'-or $_.AllowBasicAuthRpc -eq 'True'  -or
        $_.AllowBasicAuthPowerShell -eq 'True'
    })
    if ($_aiBasicEnabled.Count -gt 0) {
        Add-ActionItem -Severity 'warning' -Category 'Exchange / Legacy Auth' `
            -Text "$($_aiBasicEnabled.Count) authentication policy/policies still have Basic Auth (legacy authentication) enabled for one or more protocols. Legacy auth bypasses MFA and is a primary attack vector. Disable all unused protocols. (CIS 6.5.1)" `
            -DocUrl 'https://learn.microsoft.com/en-us/exchange/clients-and-mobile-in-exchange-online/disable-basic-authentication-in-exchange-online' `
            -CheckId 'EXAUTH-BASICAUTH-001'
    }
}

# SMTP AUTH enabled org-wide (CIS 6.5.4)
$_aiOrgCfgCsv = Join-Path $exchangeDir "Exchange_OrgConfig.csv"
if (Test-Path $_aiOrgCfgCsv) {
    $_aiOrgCfg = Import-Csv $_aiOrgCfgCsv | Select-Object -First 1
    if ($_aiOrgCfg.SmtpClientAuthDisabled -eq 'False') {
        Add-ActionItem -Severity 'critical' -Category 'Exchange / Auth' `
            -Text "SMTP client authentication (SMTP AUTH) is enabled at the organisation level. SMTP AUTH allows legacy clients to relay mail with username and password, bypassing MFA. Disable it org-wide and enable only for specific mailboxes that require it. (CIS 6.5.4)" `
            -DocUrl 'https://learn.microsoft.com/en-us/exchange/clients-and-mobile-in-exchange-online/authenticated-client-smtp-submission' `
            -CheckId 'EXAUTH-SMTPAUTH-001'
    }

    # CIS 1.3.6 — Customer Lockbox not enabled
    if ($_aiOrgCfg.CustomerLockboxEnabled -eq 'False') {
        Add-ActionItem -Severity 'warning' -Category 'Exchange / Security' `
            -Text "Customer Lockbox is not enabled. Customer Lockbox ensures that Microsoft engineers require your explicit approval before accessing your content during support operations. Enable it under Microsoft 365 admin centre → Settings → Org settings → Security & privacy. (CIS 1.3.6)" `
            -DocUrl 'https://learn.microsoft.com/en-us/purview/customer-lockbox-requests' `
            -CheckId 'EXSEC-LOCKBOX-001'
    }
}

# CIS 2.1.3 — Malware admin notification disabled
$_aiMalwareCsv = Join-Path $exchangeDir "Exchange_MalwarePolicies.csv"
if (Test-Path $_aiMalwareCsv) {
    $_aiMalwareNotify = @(Import-Csv $_aiMalwareCsv | Where-Object { $_.EnableExternalSenderAdminNotification -eq 'False' })
    if ($_aiMalwareNotify.Count -gt 0) {
        $_malList = ($_aiMalwareNotify.Name) -join ', '
        Add-ActionItem -Severity 'warning' -Category 'Exchange / Security' `
            -Text "Admin notification for malware detected from external senders is disabled in $($_aiMalwareNotify.Count) malware filter policy/policies: $_malList. Enable administrator notifications to alert on malware detected in inbound mail. (CIS 2.1.3)" `
            -DocUrl 'https://learn.microsoft.com/en-us/microsoft-365/security/office-365-security/anti-malware-protection-about' `
            -CheckId 'EXSEC-MALWARENOTIFY-001'
    }
}

# ZAP disabled on any anti-spam policy (CIS 2.1.12 / 2.1.13)
$_aiSpamCsv = Join-Path $exchangeDir "Exchange_SpamPolicies.csv"
if (Test-Path $_aiSpamCsv) {
    $_aiSpamPolicies = @(Import-Csv $_aiSpamCsv)
    $_aiZapOff = @($_aiSpamPolicies | Where-Object {
        $_.SpamZapEnabled -eq 'False' -or $_.PhishZapEnabled -eq 'False' -or $_.ZapEnabled -eq 'False'
    })
    if ($_aiZapOff.Count -gt 0) {
        $_zapList = ($_aiZapOff.Name -join ', ')
        Add-ActionItem -Severity 'warning' -Category 'Exchange / Anti-Spam' `
            -Text "Zero-hour Auto Purge (ZAP) is disabled in $($_aiZapOff.Count) anti-spam policy/policies: $_zapList. ZAP retroactively moves malicious messages delivered to mailboxes to Junk or Quarantine. Enable both Spam ZAP and Phish ZAP." `
            -DocUrl 'https://learn.microsoft.com/en-us/microsoft-365/security/office-365-security/zero-hour-auto-purge' `
            -CheckId 'EXSPAM-ZAPOFF-001'
    }
}

# External sender tagging disabled (CIS 6.2.3)
$_aiExtSenderCsv = Join-Path $exchangeDir "Exchange_ExternalSenderTagging.csv"
if (Test-Path $_aiExtSenderCsv) {
    $_aiExtSender = Import-Csv $_aiExtSenderCsv | Select-Object -First 1
    if ($_aiExtSender.Enabled -eq 'False') {
        Add-ActionItem -Severity 'warning' -Category 'Exchange / Security' `
            -Text "External sender identification (tagging) is disabled in Outlook. Users cannot easily identify email from outside the organisation, increasing susceptibility to phishing and display-name spoofing attacks. (CIS 6.2.3)" `
            -DocUrl 'https://learn.microsoft.com/en-us/microsoft-365/security/office-365-security/anti-phishing-policies-about' `
            -CheckId 'EXSEC-EXTERNTAG-001'
    }
}

# Connection filter IP allow list not empty (CIS 2.1.12)
$_aiConnFilterCsv = Join-Path $exchangeDir "Exchange_ConnectionFilter.csv"
if (Test-Path $_aiConnFilterCsv) {
    $_aiConnFilters = @(Import-Csv $_aiConnFilterCsv)
    $_aiIpAllowList = @($_aiConnFilters | Where-Object { $_.IPAllowList -and $_.IPAllowList -ne '' })
    if ($_aiIpAllowList.Count -gt 0) {
        Add-ActionItem -Severity 'warning' -Category 'Exchange / Security' `
            -Text "The connection filter IP Allow List contains entries in $($_aiIpAllowList.Count) policy/policies. IPs on the Allow List bypass spam and malware filtering. Remove all entries unless operationally required. (CIS 2.1.12)" `
            -DocUrl 'https://learn.microsoft.com/en-us/microsoft-365/security/office-365-security/connection-filter-policies-configure' `
            -CheckId 'EXSEC-IPALLOWLIST-001'
    }

    # CIS 2.1.13 — Connection filter safe list enabled
    $_aiSafeList = @(Import-Csv $_aiConnFilterCsv | Where-Object { $_.EnableSafeList -eq 'True' })
    if ($_aiSafeList.Count -gt 0) {
        Add-ActionItem -Severity 'warning' -Category 'Exchange / Security' `
            -Text "The connection filter safe list is enabled in $($_aiSafeList.Count) policy/policies. The safe list adds Microsoft-curated IP addresses to the allow list, bypassing spam filtering for those senders. Disable the safe list to maintain consistent filtering. (CIS 2.1.13)" `
            -DocUrl 'https://learn.microsoft.com/en-us/microsoft-365/security/office-365-security/connection-filter-policies-configure' `
            -CheckId 'EXSEC-SAFELIST-001'
    }
}

# OWA additional storage providers allowed (CIS 1.3.3 style)
$_aiOwaCsv = Join-Path $exchangeDir "Exchange_OwaPolicy.csv"
if (Test-Path $_aiOwaCsv) {
    $_aiOwa = @(Import-Csv $_aiOwaCsv)
    $_aiOwaStorage = @($_aiOwa | Where-Object { $_.AdditionalStorageProvidersAvailable -eq 'True' })
    if ($_aiOwaStorage.Count -gt 0) {
        Add-ActionItem -Severity 'warning' -Category 'Exchange / Security' `
            -Text "Third-party cloud storage providers (e.g. Dropbox, Google Drive) are available in Outlook on the Web in $($_aiOwaStorage.Count) OWA policy/policies. Disabling additional storage providers prevents data exfiltration via cloud attachments. (CIS 6.5.3)" `
            -DocUrl 'https://learn.microsoft.com/en-us/exchange/clients-and-mobile-in-exchange-online/outlook-on-the-web/owa-policies' `
            -CheckId 'EXSEC-CLOUDSTORAGE-001'
    }
}

# --- SharePoint checks ---

$_aiExtShareCsv = Join-Path $spDir "SharePoint_ExternalSharing_SiteOverrides.csv"
if (Test-Path $_aiExtShareCsv) {
    $_aiExtShare    = @(Import-Csv $_aiExtShareCsv)
    $_aiPermissive  = @($_aiExtShare | Where-Object { $_.SharingCapability -eq "ExternalUserAndGuestSharing" })
    if ($_aiPermissive.Count -gt 0) {
        Add-ActionItem -Severity 'warning' -Category 'SharePoint' -Text "$($_aiPermissive.Count) site(s) allow anonymous guest sharing, overriding tenant defaults. Review to confirm these are intentional. (CIS 7.2.6)" -DocUrl 'https://learn.microsoft.com/en-us/sharepoint/turn-external-sharing-on-or-off' -CheckId 'SP-ANONSHARE-001'
    }
}

$_aiOdUnlicCsv = Join-Path $spDir "SharePoint_OneDrive_Unlicensed.csv"
if (Test-Path $_aiOdUnlicCsv) {
    $_aiOdUnlic = @(Import-Csv $_aiOdUnlicCsv)
    if ($_aiOdUnlic.Count -gt 0) {
        Add-ActionItem -Severity 'warning' -Category 'SharePoint' -Text "$($_aiOdUnlic.Count) OneDrive account(s) belong to unlicensed users. Data may be inaccessible and storage costs may be wasted." -DocUrl 'https://learn.microsoft.com/en-us/sharepoint/manage-sites-in-new-admin-center' -CheckId 'SP-ODUNLIC-001'
    }
}

# Additional SharePoint security checks using expanded ExternalSharing_Tenant.csv
$_aiSpTenantCsv = Join-Path $spDir "SharePoint_ExternalSharing_Tenant.csv"
if (Test-Path $_aiSpTenantCsv) {
    $_aiSpTenant = Import-Csv $_aiSpTenantCsv | Select-Object -First 1

    # CIS 7.2.6 — Infected file download not blocked
    if ($_aiSpTenant.DisallowInfectedFileDownload -eq 'False') {
        Add-ActionItem -Severity 'warning' -Category 'SharePoint' `
            -Text "SharePoint is configured to allow users to download files flagged as infected by the built-in virus scanner. Enable 'DisallowInfectedFileDownload' to prevent download of detected malware. (CIS 7.3.1)" `
            -DocUrl 'https://learn.microsoft.com/en-us/sharepoint/disallow-infected-file-download' `
            -CheckId 'SP-INFECTED-001'
    }

    # CIS 7.2.9 — No external user link expiry (or expiry > 30 days)
    if ($_aiSpTenant.ExternalUserExpirationRequired -eq 'False') {
        Add-ActionItem -Severity 'warning' -Category 'SharePoint' `
            -Text "Guest / external user access to SharePoint does not expire automatically. Without an expiry, former partners and contractors retain access indefinitely. Enable 'ExternalUserExpirationRequired' and set a maximum of 30 days. (CIS 7.2.9)" `
            -DocUrl 'https://learn.microsoft.com/en-us/sharepoint/external-sharing-overview' `
            -CheckId 'SP-GUESTNOEXPIRY-001'
    }
    elseif ($_aiSpTenant.ExternalUserExpireInDays) {
        $expDays = 0
        if ([int]::TryParse($_aiSpTenant.ExternalUserExpireInDays, [ref]$expDays) -and $expDays -gt 30) {
            Add-ActionItem -Severity 'warning' -Category 'SharePoint' `
                -Text "External user access expiry is set to $expDays days, which exceeds the recommended 30-day maximum. Reduce 'ExternalUserExpireInDays' to 30 or fewer. (CIS 7.2.9)" `
                -DocUrl 'https://learn.microsoft.com/en-us/sharepoint/external-sharing-overview' `
                -CheckId 'SP-GUESTEXPIRY-001'
        }
    }

    # CIS 7.2.3 — External users can reshare content
    if ($_aiSpTenant.PreventExternalUsersFromResharing -eq 'False') {
        Add-ActionItem -Severity 'warning' -Category 'SharePoint' `
            -Text "External / guest users are permitted to reshare SharePoint content with other external parties. Enable 'PreventExternalUsersFromResharing' to ensure that only internal users can grant access to shared items. (CIS 7.2.5)" `
            -DocUrl 'https://learn.microsoft.com/en-us/sharepoint/external-sharing-overview' `
            -CheckId 'SP-GUESTRESHARE-001'
    }

    # CIS 7.2.11 — Default sharing link permission should be View, not Edit
    if ($_aiSpTenant.DefaultLinkPermission -eq 'Edit') {
        Add-ActionItem -Severity 'warning' -Category 'SharePoint' `
            -Text "The default sharing link permission is set to 'Edit', granting recipients edit rights by default when sharing content. Set the default link permission to 'View' so that sharing links are read-only unless the sender explicitly upgrades permissions. (CIS 7.2.11)" `
            -DocUrl 'https://learn.microsoft.com/en-us/sharepoint/change-default-sharing-link' `
            -CheckId 'SP-DEFAULTLINK-001'
    }
}

# Legacy authentication — only check if Security Defaults disabled
if (-not $_aiSdEnabled -and (Test-Path $_aiCaCsv)) {
    $_aiCaAll = @(Import-Csv $_aiCaCsv)
    $_aiLegacyBlocked = @($_aiCaAll | Where-Object {
        $_.State -eq "enabled" -and
        ($_.ClientAppTypes -match "exchangeActiveSync|other") -and
        ($_.GrantControls -match "block")
    })
    if ($_aiLegacyBlocked.Count -eq 0) {
        Add-ActionItem -Severity 'critical' -Category 'Entra / Auth' -Text "Legacy authentication does not appear to be blocked. Security Defaults is disabled and no enabled CA policy targets legacy auth client types with a Block control. Legacy auth bypasses MFA. Essential Eight ML2. (CIS 6.5.1)" -DocUrl 'https://learn.microsoft.com/en-us/entra/identity/conditional-access/block-legacy-authentication' -CheckId 'AUTH-LEGACYAUTH-001'
    }
}

# CA risk policy checks (CIS 5.2.x — user risk + sign-in risk policies)
if (-not $_aiSdEnabled -and (Test-Path $_aiCaCsv)) {
    $_aiCaRisk = @(Import-Csv $_aiCaCsv | Where-Object { $_.State -eq 'enabled' })

    # User risk policy: CA policy that grants conditionally based on user risk level
    $_aiUserRiskPolicy = @($_aiCaRisk | Where-Object { $_.UserRiskLevels -and $_.UserRiskLevels -ne '' })
    if ($_aiUserRiskPolicy.Count -eq 0) {
        Add-ActionItem -Severity 'warning' -Category 'Entra / Conditional Access' `
            -Text "No enabled Conditional Access policy targets User Risk. A user risk policy enforces MFA or block when Identity Protection detects compromised accounts. Requires Azure AD Premium P2. (CIS 2.2.8)" `
            -DocUrl 'https://learn.microsoft.com/en-us/entra/id-protection/howto-identity-protection-configure-risk-policies' `
            -CheckId 'CA-NOUSERRISK-001'
    }

    # Sign-in risk policy: CA policy targeting sign-in risk levels
    $_aiSignInRiskPolicy = @($_aiCaRisk | Where-Object { $_.SignInRiskLevels -and $_.SignInRiskLevels -ne '' })
    if ($_aiSignInRiskPolicy.Count -eq 0) {
        Add-ActionItem -Severity 'warning' -Category 'Entra / Conditional Access' `
            -Text "No enabled Conditional Access policy targets Sign-in Risk. A sign-in risk policy challenges or blocks suspicious authentication attempts in real time. Requires Azure AD Premium P2. (CIS 2.2.8)" `
            -DocUrl 'https://learn.microsoft.com/en-us/entra/id-protection/howto-identity-protection-configure-risk-policies' `
            -CheckId 'CA-NOSIGNINRISK-001'
    }
}

# Stale licensed accounts (no sign-in for 90+ days)
if (Test-Path $_aiUsersCsv) {
    $_aiStale = @($_aiUsers | Where-Object {
        $dt = [datetime]::MinValue
        -not $_.LastSignIn -or (-not [datetime]::TryParse(($_.LastSignIn -replace ' UTC',''), [ref]$dt)) -or (([datetime]::UtcNow - $dt).TotalDays -gt 90)
    })
    if ($_aiStale.Count -gt 0) {
        Add-ActionItem -Severity 'warning' -Category 'Entra / Accounts' -Text "$($_aiStale.Count) licensed user(s) have not signed in for 90+ days or have no recorded sign-in. Review for stale/orphaned accounts." -DocUrl 'https://learn.microsoft.com/en-us/entra/identity/monitoring-health/recommendation-remove-unused-credential-from-apps' -CheckId 'ACCT-STALE-001'
    }
}

# Stale guest accounts (no sign-in for 90+ days)
$_aiGuestCsv = Join-Path $entraDir "Entra_GuestUsers.csv"
if (Test-Path $_aiGuestCsv) {
    $_aiGuests      = @(Import-Csv $_aiGuestCsv)
    $_aiStaleGuests = @($_aiGuests | Where-Object {
        $dt = [datetime]::MinValue
        -not $_.LastSignIn -or (-not [datetime]::TryParse(($_.LastSignIn -replace ' UTC',''), [ref]$dt)) -or (([datetime]::UtcNow - $dt).TotalDays -gt 90)
    })
    if ($_aiStaleGuests.Count -gt 0) {
        Add-ActionItem -Severity 'warning' -Category 'Entra / Guests' -Text "$($_aiStaleGuests.Count) guest account(s) have not signed in for 90+ days or have no recorded sign-in. Stale guests retain access to shared resources." -DocUrl 'https://learn.microsoft.com/en-us/entra/identity/users/manage-guest-access-with-access-reviews' -CheckId 'GUEST-STALE-001'
    }
}

# Entra org settings checks
$_aiOrgSettingsCsv = Join-Path $entraDir "Entra_OrgSettings.csv"
if (Test-Path $_aiOrgSettingsCsv) {
    $_aiOrgSettings = Import-Csv $_aiOrgSettingsCsv | Select-Object -First 1

    # CIS 5.1.1 — Users can register applications
    if ($_aiOrgSettings.AllowedToCreateApps -eq 'True') {
        Add-ActionItem -Severity 'warning' -Category 'Entra / Org Settings' `
            -Text "All users are permitted to register application (app registrations). Users should not be able to create app registrations — restrict this to administrators to prevent unauthorised OAuth app registration. (CIS 1.5.1)" `
            -DocUrl 'https://learn.microsoft.com/en-us/entra/identity/role-based-access-control/delegate-app-roles' `
            -CheckId 'ORGSET-APPREGALL-001'
    }

    # CIS 5.1.2 — Admin consent workflow disabled
    if ($_aiOrgSettings.AdminConsentWorkflowEnabled -eq 'False') {
        Add-ActionItem -Severity 'warning' -Category 'Entra / Org Settings' `
            -Text "The admin consent request workflow is disabled. When users need an app requiring admin consent, they have no formal approval path and may resort to workarounds. Enable the admin consent workflow so requests are reviewed rather than ignored. (CIS 1.5.2)" `
            -DocUrl 'https://learn.microsoft.com/en-us/entra/identity/enterprise-apps/configure-admin-consent-workflow' `
            -CheckId 'ORGSET-NOCONSENT-001'
    }

    # CIS 1.2.3 — Users can create M365 tenants
    if ($_aiOrgSettings.AllowedToCreateTenants -eq 'True') {
        Add-ActionItem -Severity 'warning' -Category 'Entra / Org Settings' `
            -Text "All users are permitted to create new Microsoft 365 tenants. Non-admin users should not be able to create tenants, as this can lead to shadow IT and unmanaged Microsoft 365 environments. Restrict tenant creation to administrators. (CIS 1.2.3)" `
            -DocUrl 'https://learn.microsoft.com/en-us/entra/fundamentals/default-user-permissions' `
            -CheckId 'ORGSET-TENANTCREATE-001'
    }
}

# Shared mailbox sign-in enabled
$_aiSharedSignInCsv = Join-Path $exchangeDir "Exchange_SharedMailboxSignIn.csv"
if (Test-Path $_aiSharedSignInCsv) {
    $_aiSharedEnabled = @(Import-Csv $_aiSharedSignInCsv | Where-Object { $_.AccountDisabled -eq "False" })
    if ($_aiSharedEnabled.Count -gt 0) {
        Add-ActionItem -Severity 'warning' -Category 'Exchange / Mailboxes' -Text "$($_aiSharedEnabled.Count) shared mailbox(es) have interactive sign-in enabled. Shared mailboxes should have sign-in disabled to prevent direct login and MFA bypass. (CIS 1.2.2)" -DocUrl 'https://learn.microsoft.com/en-us/microsoft-365/admin/email/about-shared-mailboxes' -CheckId 'EXMBX-SHAREDSIGNIN-001'
    }
}

# Mailboxes nearly full (>75% used) with no In-Place Archive
$_aiMbxCapCsv = Join-Path $exchangeDir "Exchange_Mailboxes.csv"
if (Test-Path $_aiMbxCapCsv) {
    $_aiMbxNearFull = @(Import-Csv $_aiMbxCapCsv | Where-Object {
        $_.LimitMB -and $_.LimitMB -ne '' -and [double]$_.LimitMB -gt 0 -and
        $_.ArchiveEnabled -eq 'False' -and
        ([double]$_.TotalSizeMB / [double]$_.LimitMB) -gt 0.75
    })
    if ($_aiMbxNearFull.Count -gt 0) {
        $_nearFullList = ($_aiMbxNearFull | ForEach-Object {
            $pct = [math]::Round(([double]$_.TotalSizeMB / [double]$_.LimitMB) * 100, 0)
            "$($_.DisplayName) ($($_.UserPrincipalName)) — $pct% used, no archive"
        }) -join '<br>'
        Add-ActionItem -Severity 'warning' -Category 'Exchange / Mailboxes' `
            -Text "$($_aiMbxNearFull.Count) mailbox(es) are over 75% full and do not have an In-Place Archive enabled. Enable archiving to prevent mail delivery failures when the quota is reached:<br>$_nearFullList" `
            -DocUrl 'https://learn.microsoft.com/en-us/microsoft-365/compliance/enable-archive-mailboxes' `
            -CheckId 'EXMBX-NEARFULL-001'
    }
}

# Outbound spam auto-forward policy
$_aiOutboundCsv = Join-Path $exchangeDir "Exchange_OutboundSpamAutoForward.csv"
if (Test-Path $_aiOutboundCsv) {
    $_aiOutboundOn = @(Import-Csv $_aiOutboundCsv | Where-Object { $_.AutoForwardingMode -eq "On" })
    if ($_aiOutboundOn.Count -gt 0) {
        Add-ActionItem -Severity 'critical' -Category 'Exchange / Forwarding' -Text "Outbound spam policy is set to always allow auto-forwarding (AutoForwardingMode = On). This permits unrestricted external mail forwarding and is a known data exfiltration vector. (CIS 6.2.1)" -DocUrl 'https://learn.microsoft.com/en-us/microsoft-365/security/office-365-security/outbound-spam-protection-about' -CheckId 'EXFWD-AUTOFWDPOLICY-001'
    }
}

# Safe Attachments
$_aiSafAttCsv = Join-Path $exchangeDir "Exchange_SafeAttachments.csv"
if (Test-Path $_aiSafAttCsv) {
    $_aiSafAtt = @(Import-Csv $_aiSafAttCsv)
    $_aiSafAttOn = @($_aiSafAtt | Where-Object { $_.Enable -eq "True" })
    if ($_aiSafAttOn.Count -eq 0) {
        Add-ActionItem -Severity 'warning' -Category 'Exchange / Defender' -Text "No Safe Attachments policy is enabled. Attachments are not being detonated/scanned before delivery. Requires Defender for Office 365 P1. (CIS 2.1.4)" -DocUrl 'https://learn.microsoft.com/en-us/microsoft-365/security/office-365-security/safe-attachments-about' -CheckId 'EXDEF-NOATTACH-001'
    }
}

# Safe Links
$_aiSafLnkCsv = Join-Path $exchangeDir "Exchange_SafeLinks.csv"
if (Test-Path $_aiSafLnkCsv) {
    $_aiSafLnk = @(Import-Csv $_aiSafLnkCsv)
    $_aiSafLnkOn = @($_aiSafLnk | Where-Object { $_.EnableSafeLinksForEmail -eq "True" })
    if ($_aiSafLnkOn.Count -eq 0) {
        Add-ActionItem -Severity 'warning' -Category 'Exchange / Defender' -Text "No Safe Links policy is enabled for email. URLs are not being rewritten or checked at time-of-click. Requires Defender for Office 365 P1. (CIS 2.1.1)" -DocUrl 'https://learn.microsoft.com/en-us/microsoft-365/security/office-365-security/safe-links-about' -CheckId 'EXDEF-NOLINKS-001'
    }
}

# SharePoint default sharing link type
$_aiSpTenantCsv = Join-Path $spDir "SharePoint_ExternalSharing_Tenant.csv"
if (Test-Path $_aiSpTenantCsv) {
    $_aiSpTenant = Import-Csv $_aiSpTenantCsv | Select-Object -First 1
    if ($_aiSpTenant.DefaultSharingLinkType -eq "AnonymousAccess") {
        Add-ActionItem -Severity 'warning' -Category 'SharePoint' -Text "Default sharing link type is set to 'Anyone' (anonymous). Every share defaults to a link accessible by anyone with the URL, with no sign-in required. (CIS 7.2.7)" -DocUrl 'https://learn.microsoft.com/en-us/sharepoint/change-default-sharing-link' -CheckId 'SP-ANONLINKTYPE-001'
    }
}

# SharePoint sync restriction
$_aiSpAcpCsv = Join-Path $spDir "SharePoint_AccessControlPolicies.csv"
if (Test-Path $_aiSpAcpCsv) {
    $_aiSpAcp = Import-Csv $_aiSpAcpCsv | Select-Object -First 1
    if ($_aiSpAcp.IsUnmanagedSyncAppForTenantRestricted -eq "False") {
        Add-ActionItem -Severity 'warning' -Category 'SharePoint' -Text "OneDrive sync is not restricted to managed/domain-joined devices. Any personal device can sync corporate data to local storage. (CIS 7.3.2)" -DocUrl 'https://learn.microsoft.com/en-us/sharepoint/control-access-from-unmanaged-devices' -CheckId 'SP-SYNCRESTRICT-001'
    }
}

# --- Intune checks ---

$_aiIntuneLicCsv = Join-Path $intuneDir "Intune_LicenceCheck.csv"
if (Test-Path $_aiIntuneLicCsv) {
    $_aiIntuneLic = Import-Csv $_aiIntuneLicCsv | Select-Object -First 1

    if ($_aiIntuneLic.HasIntune -ne 'True') {
        # No licence — suppress further Intune checks; the HTML section will render an info note
    }
    else {
        # No compliance policies
        $_aiCompPolCsv = Join-Path $intuneDir "Intune_CompliancePolicies.csv"
        if (Test-Path $_aiCompPolCsv) {
            $_aiCompPols = @(Import-Csv $_aiCompPolCsv)
            if ($_aiCompPols.Count -eq 0) {
                Add-ActionItem -Severity 'critical' -Category 'Intune / Compliance' `
                    -Text "No compliance policies have been configured in Intune. Without policies, all devices are considered compliant regardless of their actual security state." `
                    -DocUrl 'https://learn.microsoft.com/en-us/mem/intune/protect/device-compliance-get-started' `
                    -CheckId 'INTUNE-NOPOLICY-001'
            }
        }

        # Non-compliant devices
        $_aiDevicesCsv = Join-Path $intuneDir "Intune_Devices.csv"
        if (Test-Path $_aiDevicesCsv) {
            $_aiDevices        = @(Import-Csv $_aiDevicesCsv)
            $_aiNonCompliant   = @($_aiDevices | Where-Object { $_.ComplianceState -eq 'noncompliant' })
            $_aiStaleDevices   = @($_aiDevices | Where-Object {
                $dt = [datetime]::MinValue
                $_.LastSyncDateTime -and [datetime]::TryParse($_.LastSyncDateTime, [ref]$dt) -and (([datetime]::UtcNow - $dt).TotalDays -gt 30)
            })
            if ($_aiNonCompliant.Count -gt 0) {
                $_ncList = ($_aiNonCompliant | Select-Object -First 10 | ForEach-Object { $_.DeviceName }) -join ', '
                if ($_aiNonCompliant.Count -gt 10) { $_ncList += ", ..." }
                Add-ActionItem -Severity 'critical' -Category 'Intune / Compliance' `
                    -Text "$($_aiNonCompliant.Count) device(s) are currently non-compliant: $($_ncList). Non-compliant devices may retain access to corporate resources if Conditional Access is not enforcing compliance." `
                    -DocUrl 'https://learn.microsoft.com/en-us/mem/intune/protect/device-compliance-get-started' `
                    -CheckId 'INTUNE-NONCOMPLIANT-001'
            }
            if ($_aiStaleDevices.Count -gt 0) {
                Add-ActionItem -Severity 'warning' -Category 'Intune / Devices' `
                    -Text "$($_aiStaleDevices.Count) device(s) have not checked in with Intune for more than 30 days. Stale devices may not receive policy updates or be accurately reflected in compliance reports." `
                    -DocUrl 'https://learn.microsoft.com/en-us/mem/intune/remote-actions/devices-wipe' `
                    -CheckId 'INTUNE-STALEDEV-001'
            }
        }

        # Encryption not required by any compliance policy
        $_aiCompSettingsCsv = Join-Path $intuneDir "Intune_CompliancePolicySettings.csv"
        if (Test-Path $_aiCompSettingsCsv) {
            $_aiCompSettings = @(Import-Csv $_aiCompSettingsCsv)
            $_aiNoEncrypt = @($_aiCompSettings | Where-Object {
                ($_.SettingName -eq 'storageRequireEncryption' -or $_.SettingName -eq 'bitLockerEnabled') -and
                $_.SettingValue -eq 'False'
            })
            if ($_aiNoEncrypt.Count -gt 0) {
                $_encPolicies = ($_aiNoEncrypt.PolicyName | Sort-Object -Unique) -join ', '
                Add-ActionItem -Severity 'warning' -Category 'Intune / Compliance' `
                    -Text "Storage encryption is explicitly disabled in $($_aiNoEncrypt.Count) compliance policy setting(s): $_encPolicies. Devices governed by these policies are not required to have encryption enabled." `
                    -DocUrl 'https://learn.microsoft.com/en-us/mem/intune/protect/compliance-policy-create-windows' `
                    -CheckId 'INTUNE-NOENCRYPT-001'
            }
        }

        # Long grace period (>24 hours)
        if (Test-Path $_aiCompPolCsv) {
            $_aiLongGrace = @($_aiCompPols | Where-Object {
                $g = 0
                [int]::TryParse($_.GracePeriodHours, [ref]$g) | Out-Null
                $g -gt 24
            })
            if ($_aiLongGrace.Count -gt 0) {
                $_graceList = ($_aiLongGrace | ForEach-Object { "$($_.PolicyName) ($($_.GracePeriodHours)h)" }) -join ', '
                Add-ActionItem -Severity 'warning' -Category 'Intune / Compliance' `
                    -Text "$($_aiLongGrace.Count) compliance policy(ies) have a grace period exceeding 24 hours before a device is marked non-compliant: $_graceList. Devices remain in a grace state and are not blocked during this window." `
                    -DocUrl 'https://learn.microsoft.com/en-us/mem/intune/protect/actions-for-noncompliance' `
                    -CheckId 'INTUNE-GRACELONG-001'
            }
        }

        # Personal device enrolment not blocked
        $_aiEnrolCsv = Join-Path $intuneDir "Intune_EnrollmentRestrictions.csv"
        if (Test-Path $_aiEnrolCsv) {
            $_aiEnrolRows = @(Import-Csv $_aiEnrolCsv)
            $_aiPersonalAllowed = @($_aiEnrolRows | Where-Object { $_.BlockPersonalDevices -eq 'False' })
            if ($_aiPersonalAllowed.Count -gt 0) {
                Add-ActionItem -Severity 'warning' -Category 'Intune / Enrollment' `
                    -Text "$($_aiPersonalAllowed.Count) enrollment restriction(s) allow personal (BYOD) device enrolment. Personal devices enrolled in Intune may have weaker compliance controls than corporate-owned devices." `
                    -DocUrl 'https://learn.microsoft.com/en-us/mem/intune/enrollment/enrollment-restrictions-set' `
                    -CheckId 'INTUNE-BYOD-001'
            }
        }

        # Apps with install failures
        $_aiAppsCsv = Join-Path $intuneDir "Intune_Apps.csv"
        if (Test-Path $_aiAppsCsv) {
            $_aiApps        = @(Import-Csv $_aiAppsCsv)
            $_aiFailedApps  = @($_aiApps | Where-Object {
                $f = 0
                [int]::TryParse($_.FailedDeviceCount, [ref]$f) | Out-Null
                $f -gt 0
            })
            if ($_aiFailedApps.Count -gt 0) {
                $_failList = ($_aiFailedApps | Select-Object -First 5 | ForEach-Object { "$($_.AppName) ($($_.FailedDeviceCount) failed)" }) -join ', '
                if ($_aiFailedApps.Count -gt 5) { $_failList += ", ..." }
                Add-ActionItem -Severity 'warning' -Category 'Intune / Apps' `
                    -Text "$($_aiFailedApps.Count) app(s) have deployment failures: $_failList. Failed installations may leave required security or business apps absent from affected devices." `
                    -DocUrl 'https://learn.microsoft.com/en-us/mem/intune/apps/troubleshoot-app-install' `
                    -CheckId 'INTUNE-APPFAIL-001'
            }
        }
    }
}

# --- Teams checks ---

$_aiTeamsFedCsv = Join-Path $teamsDir "Teams_FederationConfig.csv"
if (Test-Path $_aiTeamsFedCsv) {
    $_aiTeamsFed = Import-Csv $_aiTeamsFedCsv | Select-Object -First 1

    # CIS 8.1.2 — Non-managed (consumer) Teams accounts can communicate with users
    if ($_aiTeamsFed.AllowPublicUsers -eq 'True') {
        Add-ActionItem -Severity 'warning' -Category 'Teams / External Access' `
            -Text "External communication with non-managed Teams accounts (personal/consumer) is enabled. Users can communicate with anyone outside the organisation who has a personal Teams account, increasing the risk of phishing and data exfiltration. (CIS 8.2.2)" `
            -DocUrl 'https://learn.microsoft.com/en-us/microsoftteams/manage-external-access' `
            -CheckId 'TEAMS-EXTUNMANAGED-001'
    }

    # CIS 8.1.3 — Teams Consumer inbound communication (unmanaged Teams → internal users)
    if ($_aiTeamsFed.AllowTeamsConsumerInbound -eq 'True') {
        Add-ActionItem -Severity 'warning' -Category 'Teams / External Access' `
            -Text "Inbound communication from non-managed Teams consumer accounts is permitted. External consumer Teams users can initiate contact with internal users, creating an uncontrolled inbound channel. (CIS 8.2.3)" `
            -DocUrl 'https://learn.microsoft.com/en-us/microsoftteams/manage-external-access' `
            -CheckId 'TEAMS-EXTCONSUMER-001'
    }

    # CIS 8.2.1 — External federation open with no domain restrictions
    if ($_aiTeamsFed.AllowFederatedUsers -eq 'True' -and [int]$_aiTeamsFed.AllowedDomainsCount -eq 0 -and [int]$_aiTeamsFed.BlockedDomainsCount -eq 0) {
        Add-ActionItem -Severity 'warning' -Category 'Teams / External Access' `
            -Text "External federation is open to all domains with no allow-list or block-list restrictions. Any Teams tenant can communicate with your organisation without restriction. Configure allowed or blocked domains. (CIS 8.2.1)" `
            -DocUrl 'https://learn.microsoft.com/en-us/microsoftteams/manage-external-access' `
            -CheckId 'TEAMS-EXTFEDOPEN-001'
    }
}

$_aiTeamsMtgCsv = Join-Path $teamsDir "Teams_MeetingPolicies.csv"
if (Test-Path $_aiTeamsMtgCsv) {
    $_aiTeamsMtg = @(Import-Csv $_aiTeamsMtgCsv)
    $_aiGlobalMtg = $_aiTeamsMtg | Where-Object { $_.Identity -eq 'Global' } | Select-Object -First 1

    if ($_aiGlobalMtg) {
        # CIS 8.5.1 — Anonymous users can start meetings
        if ($_aiGlobalMtg.AllowAnonymousUsersToStartMeeting -eq 'True') {
            Add-ActionItem -Severity 'critical' -Category 'Teams / Meetings' `
                -Text "Anonymous users are permitted to start Teams meetings (Global policy). This allows unauthenticated external parties to initiate meetings on behalf of your organisation without any identity verification. (CIS 8.5.1)" `
                -DocUrl 'https://learn.microsoft.com/en-us/microsoftteams/meeting-policies-participants-and-guests' `
                -CheckId 'TEAMS-ANONSTART-001'
        }

        # CIS 8.5.2/8.5.3 — Lobby bypass allows everyone (Everyone bypasses = no lobby)
        if ($_aiGlobalMtg.AutoAdmittedUsers -eq 'Everyone') {
            Add-ActionItem -Severity 'warning' -Category 'Teams / Meetings' `
                -Text "The global Teams meeting policy admits everyone directly without passing through the lobby (AutoAdmittedUsers = Everyone). Anonymous and external users join meetings without host approval. (CIS 8.5.3)" `
                -DocUrl 'https://learn.microsoft.com/en-us/microsoftteams/meeting-policies-participants-and-guests' `
                -CheckId 'TEAMS-AUTOADMIT-001'
        }
    }
}

# CIS 8.5.7, 8.5.8, 8.5.9 — Additional meeting policy checks
$_aiMtgCsv2 = Join-Path $teamsDir "Teams_MeetingPolicies.csv"
if (Test-Path $_aiMtgCsv2) {
    $_aiGlobalMtg2 = Import-Csv $_aiMtgCsv2 | Where-Object { $_.Identity -eq 'Global' } | Select-Object -First 1
    if (-not $_aiGlobalMtg2) { $_aiGlobalMtg2 = Import-Csv $_aiMtgCsv2 | Select-Object -First 1 }
    if ($_aiGlobalMtg2) {
        if ($_aiGlobalMtg2.AllowExternalParticipantGiveRequestControl -eq 'True') {
            Add-ActionItem -Severity 'warning' -Category 'Teams / Meetings' `
                -Text "External meeting participants are permitted to give or request presenter control. This allows external users to take control of meeting content and screen sharing. (CIS 8.5.7)" `
                -DocUrl 'https://learn.microsoft.com/en-us/microsoftteams/meeting-policies-in-teams-general' `
                -CheckId 'TEAMS-EXTCONTROL-001'
        }
        if ($_aiGlobalMtg2.AllowExternalNonTrustedMeetingChat -eq 'True') {
            Add-ActionItem -Severity 'warning' -Category 'Teams / Meetings' `
                -Text "Meeting chat is enabled for external (non-trusted) participants. External attendees can send chat messages in meetings, increasing exposure to phishing links and malicious content. (CIS 8.5.8)" `
                -DocUrl 'https://learn.microsoft.com/en-us/microsoftteams/meeting-policies-in-teams-general' `
                -CheckId 'TEAMS-EXTCHAT-001'
        }
        if ($_aiGlobalMtg2.AllowCloudRecording -eq 'True') {
            Add-ActionItem -Severity 'warning' -Category 'Teams / Meetings' `
                -Text "Cloud meeting recording is enabled by default in the global Teams meeting policy. Recordings may contain sensitive discussions and are stored in SharePoint/OneDrive. Disable by default and allow on a per-user basis as required. (CIS 8.5.9)" `
                -DocUrl 'https://learn.microsoft.com/en-us/microsoftteams/meeting-recording' `
                -CheckId 'TEAMS-CLOUDRECORD-001'
        }
    }
}

$_aiTeamsClientCsv = Join-Path $teamsDir "Teams_ClientConfig.csv"
if (Test-Path $_aiTeamsClientCsv) {
    $_aiTeamsClient = Import-Csv $_aiTeamsClientCsv | Select-Object -First 1

    # CIS 8.6.2 — Third-party cloud storage providers allowed
    $_aiStorageProviders = @()
    if ($_aiTeamsClient.AllowBox -eq 'True')        { $_aiStorageProviders += 'Box' }
    if ($_aiTeamsClient.AllowDropBox -eq 'True')    { $_aiStorageProviders += 'Dropbox' }
    if ($_aiTeamsClient.AllowEgnyte -eq 'True')     { $_aiStorageProviders += 'Egnyte' }
    if ($_aiTeamsClient.AllowGoogleDrive -eq 'True') { $_aiStorageProviders += 'Google Drive' }
    if ($_aiTeamsClient.AllowShareFile -eq 'True')  { $_aiStorageProviders += 'ShareFile' }

    if ($_aiStorageProviders.Count -gt 0) {
        Add-ActionItem -Severity 'warning' -Category 'Teams / Apps' `
            -Text "Third-party cloud storage is enabled in Teams: $($_aiStorageProviders -join ', '). Files shared through external storage services bypass organisational DLP policies and data governance controls. (CIS 8.1.1)" `
            -DocUrl 'https://learn.microsoft.com/en-us/microsoftteams/teams-client-configuration' `
            -CheckId 'TEAMS-CLOUDSTORAGE-001'
    }
}

$_aiTeamsAppSetupCsv = Join-Path $teamsDir "Teams_AppSetupPolicies.csv"
if (Test-Path $_aiTeamsAppSetupCsv) {
    $_aiTeamsAppSetup = @(Import-Csv $_aiTeamsAppSetupCsv)
    $_aiGlobalAppSetup = $_aiTeamsAppSetup | Where-Object { $_.Identity -eq 'Global' } | Select-Object -First 1

    # CIS 8.6.1 — App sideloading allowed in global policy
    if ($_aiGlobalAppSetup -and $_aiGlobalAppSetup.AllowSideloading -eq 'True') {
        Add-ActionItem -Severity 'warning' -Category 'Teams / Apps' `
            -Text "Custom app sideloading is enabled in the global Teams app setup policy. Users can install unverified third-party apps directly into Teams without IT review or approval." `
            -DocUrl 'https://learn.microsoft.com/en-us/microsoftteams/teams-app-setup-policies' `
            -CheckId 'TEAMS-SIDELOAD-001'
    }
}

# --- Build KPI strip values (uses data already loaded during action item checks) ---
$_kpiMfaPct    = if ($null -ne $_aiPct)   { $_aiPct }   else { $null }
$_kpiUserCount = if ($null -ne $_aiTotal) { $_aiTotal } else { $null }
$_kpiCritCount = @($actionItems | Where-Object { $_.Severity -eq 'critical' }).Count
$_kpiWarnCount = @($actionItems | Where-Object { $_.Severity -eq 'warning'  }).Count

$_kpiScoreVal  = '&mdash;'
$_kpiScorePct  = $null
$_kpiScorePath = Join-Path $entraDir 'Entra_SecureScore.csv'
if (Test-Path $_kpiScorePath) {
    $_kpiSs = Import-Csv $_kpiScorePath | Select-Object -First 1
    if ($_kpiSs) {
        $_kpiScoreVal = "$($_kpiSs.CurrentScore) / $($_kpiSs.MaxScore)"
        $_kpiScorePct = [double]$_kpiSs.Percentage
    }
}

$_kpiDevCount     = $null
$_kpiNonCompliant = $null
$_kpiDevPath = Join-Path $intuneDir 'Intune_Devices.csv'
if (Test-Path $_kpiDevPath) {
    $_kpiDevRows      = @(Import-Csv $_kpiDevPath)
    $_kpiDevCount     = $_kpiDevRows.Count
    $_kpiNonCompliant = @($_kpiDevRows | Where-Object { $_.ComplianceState -eq 'noncompliant' }).Count
}

$_kpiMfaClass   = if ($null -eq $_kpiMfaPct)  { 'ok' } elseif ($_kpiMfaPct -eq 100) { 'ok' } elseif ($_kpiMfaPct -ge 80) { 'warn' } else { 'critical' }
$_kpiScoreClass = if ($null -eq $_kpiScorePct) { 'ok' } elseif ($_kpiScorePct -ge 80) { 'ok' } elseif ($_kpiScorePct -ge 50) { 'warn' } else { 'critical' }
$_kpiAiClass    = if ($_kpiCritCount -gt 0)  { 'critical' } elseif ($_kpiWarnCount -gt 0) { 'warn' } else { 'ok' }
$_kpiDevClass   = if ($null -eq $_kpiNonCompliant -or $_kpiNonCompliant -eq 0) { 'ok' } else { 'critical' }

$_kpiMfaStr     = if ($null -ne $_kpiMfaPct)    { "${_kpiMfaPct}%" }            else { '&mdash;' }
$_kpiMfaSub     = if ($null -ne $_kpiUserCount)  { "$_kpiUserCount licensed users" } else { '' }
$_kpiScoreSub   = if ($null -ne $_kpiScorePct)   { "$_kpiScorePct%" }            else { '' }
$_kpiDevStr     = if ($null -ne $_kpiDevCount)    { "$_kpiDevCount" }             else { '&mdash;' }
$_kpiDevSub     = if ($null -ne $_kpiNonCompliant -and $_kpiNonCompliant -gt 0) { "$_kpiNonCompliant non-compliant" } elseif ($null -ne $_kpiDevCount) { 'all compliant' } else { '' }
$_kpiAiStr      = if ($_kpiCritCount -gt 0) { "$_kpiCritCount critical" } elseif ($_kpiWarnCount -gt 0) { "$_kpiWarnCount warnings" } else { 'All clear' }
$_kpiAiSub      = if ($_kpiCritCount -gt 0 -and $_kpiWarnCount -gt 0) { "+ $_kpiWarnCount warnings" } else { '' }

$_kpiStorageStr   = '&mdash;'
$_kpiStorageSub   = ''
$_kpiStorageClass = 'ok'
$_kpiStoPath = Join-Path $spDir 'SharePoint_TenantStorage.csv'
if (Test-Path $_kpiStoPath) {
    $_kpiSto = Import-Csv $_kpiStoPath | Select-Object -First 1
    if ($_kpiSto) {
        $_kpiStoUsedMB  = [double]$_kpiSto.StorageUsedMB
        $_kpiStoQuotaMB = [double]$_kpiSto.StorageQuotaMB
        if ($_kpiStoUsedMB -le 0 -and $_kpiStoQuotaMB -gt 0) {
            $_kpiStoSites = try { Import-Csv (Join-Path $spDir 'SharePoint_Sites.csv')        -ErrorAction Stop } catch { @() }
            $_kpiStoOd    = try { Import-Csv (Join-Path $spDir 'SharePoint_OneDriveUsage.csv') -ErrorAction Stop } catch { @() }
            $_kpiStoUsedMB = (@($_kpiStoSites) + @($_kpiStoOd) | Measure-Object -Property StorageUsedMB -Sum).Sum
        }
        if ($_kpiStoQuotaMB -gt 0) {
            $_kpiStoUsedGB  = [math]::Round($_kpiStoUsedMB  / 1024, 1)
            $_kpiStoTotalGB = [math]::Round($_kpiStoQuotaMB / 1024, 1)
            $_kpiStoPct     = [math]::Round(($_kpiStoUsedMB / $_kpiStoQuotaMB) * 100, 0)
            $_kpiStorageStr   = "$_kpiStoUsedGB / $_kpiStoTotalGB GB"
            $_kpiStorageSub   = "$_kpiStoPct% used"
            $_kpiStorageClass = if ($_kpiStoPct -ge 90) { 'critical' } elseif ($_kpiStoPct -ge 75) { 'warn' } else { 'ok' }
        }
    }
}

# --- Build sidebar nav (status dots derived from action item categories) ---
# Helper: build a sub-item only if its associated CSV exists (or no Csv key provided = always show)
function New-SbSub { param([string]$Id, [string]$Label, [string]$Csv = '', [string]$Dir = '')
    if ($Csv -and $Dir -and -not (Test-Path (Join-Path $Dir $Csv))) { return $null }
    return @{ Id = $Id; Label = $Label }
}
$_sbDefinitions = @(
    @{ Id = 'entra'; Label = 'Microsoft Entra'; Prefix = 'Entra'; Subs = @(
        New-SbSub 'entra-score'       'Secure Score'            'Entra_SecureScore.csv'        $entraDir
        New-SbSub 'entra-users'       'User Accounts'           'Entra_Users.csv'              $entraDir
        New-SbSub 'entra-licenses'    'Licences'                'Entra_Licenses.csv'           $entraDir
        New-SbSub 'entra-admins'      'Admin Roles'             'Entra_AdminRoles.csv'         $entraDir
        New-SbSub 'entra-authmethods' 'Auth Methods Policy'     'Entra_AuthMethodsPolicy.csv'  $entraDir
        New-SbSub 'entra-extcollab'   'External Collaboration'  'Entra_ExternalCollab.csv'     $entraDir
        New-SbSub 'entra-appregs'     'App Registrations'       'Entra_AppRegistrations.csv'   $entraDir
        New-SbSub 'entra-pim'         'PIM Assignments'         'Entra_PIMAssignments.csv'     $entraDir
        New-SbSub 'entra-ca'          'Conditional Access'      'Entra_CA_Policies.csv'        $entraDir
        New-SbSub 'entra-idprot'      'Identity Protection'     'Entra_RiskyUsers.csv'         $entraDir
        New-SbSub 'entra-groups'      'Groups'                  'Entra_Groups.csv'             $entraDir
        New-SbSub 'entra-apps'        'Enterprise Apps'         'Entra_EnterpriseApps.csv'     $entraDir
        New-SbSub 'entra-orgsettings' 'Org User Settings'       'Entra_OrgSettings.csv'        $entraDir
    ) | Where-Object { $null -ne $_ } },
    @{ Id = 'exchange'; Label = 'Exchange Online'; Prefix = 'Exchange'; Subs = @(
        New-SbSub 'exchange-mailboxes'  'Mailboxes'         'Exchange_Mailboxes.csv'           $exchangeDir
        New-SbSub 'exchange-forwarding' 'Forwarding Rules'  'Exchange_InboxForwardingRules.csv' $exchangeDir
        New-SbSub 'exchange-policies'   'Security Policies' 'Exchange_AntiSpamPolicies.csv'    $exchangeDir
        New-SbSub 'exchange-org'        'Org Configuration' 'Exchange_OrgConfig.csv'           $exchangeDir
    ) | Where-Object { $null -ne $_ } },
    @{ Id = 'mailsec'; Label = 'Mail Security'; Prefix = 'Mail Security'; Subs = @(
        New-SbSub 'mailsec-records' 'DNS Records' 'MailSec_SPF.csv' $mailSecDir
    ) | Where-Object { $null -ne $_ } },
    @{ Id = 'sharepoint'; Label = 'SharePoint / OneDrive'; Prefix = 'SharePoint'; Subs = @(
        New-SbSub 'sp-storage'  'Storage'          'SharePoint_TenantStorage.csv'         $spDir
        New-SbSub 'sp-sites'    'Sites'            'SharePoint_Sites.csv'                 $spDir
        New-SbSub 'sp-sharing'  'External Sharing' 'SharePoint_ExternalSharing_Tenant.csv' $spDir
        New-SbSub 'sp-onedrive' 'OneDrive'         'SharePoint_OneDriveUsage.csv'         $spDir
    ) | Where-Object { $null -ne $_ } },
    @{ Id = 'teams'; Label = 'Microsoft Teams'; Prefix = 'Teams'; Subs = @(
        New-SbSub 'teams-external'  'External Access'  'Teams_FederationConfig.csv'         $teamsDir
        New-SbSub 'teams-meetings'  'Meeting Policies' 'Teams_MeetingPolicies.csv'          $teamsDir
        New-SbSub 'teams-apps'      'App Policies'     'Teams_AppPermissionPolicies.csv'    $teamsDir
    ) | Where-Object { $null -ne $_ } },
    @{ Id = 'intune'; Label = 'Intune'; Prefix = 'Intune'; Subs = @(
        New-SbSub 'intune-devices'    'Devices'         'Intune_Devices.csv'            $intuneDir
        New-SbSub 'intune-compliance' 'Compliance'      'Intune_CompliancePolicies.csv' $intuneDir
        New-SbSub 'intune-config'     'Config Profiles' 'Intune_ConfigProfiles.csv'     $intuneDir
        New-SbSub 'intune-apps'       'Apps'            'Intune_Apps.csv'               $intuneDir
    ) | Where-Object { $null -ne $_ } }
)
if ($_scubaResults) {
    $_sgSbSubs    = [System.Collections.Generic.List[hashtable]]::new()
    $_sgSbSubs.Add(@{ Id = 'scuba-summary'; Label = 'Baseline Summary' })
    $_sgSbProdMap = @{ AAD='Identity (AAD)'; EXO='Exchange Online'; SharePoint='SharePoint'; Teams='Teams'; Defender='Defender'; PowerPlatform='Power Platform' }
    foreach ($_sgSbProd in @('AAD','EXO','SharePoint','Teams','Defender','PowerPlatform')) {
        if ($_scubaResults.Results.PSObject.Properties.Name -contains $_sgSbProd) {
            $_sgSbLabel = if ($_sgSbProdMap.ContainsKey($_sgSbProd)) { $_sgSbProdMap[$_sgSbProd] } else { $_sgSbProd }
            $_sgSbSubs.Add(@{ Id = "scuba-$($_sgSbProd.ToLower())"; Label = $_sgSbLabel })
        }
    }
    $_sbDefinitions += @(@{ Id = 'scuba'; Label = 'ScubaGear Baseline'; Prefix = 'ScubaGear'; Subs = @($_sgSbSubs) })
}
$_sbItemsHtml = foreach ($_mod in $_sbDefinitions) {
    $_mc = @($actionItems | Where-Object { $_.Severity -eq 'critical' -and $_.Category -like "$($_mod.Prefix)*" }).Count
    $_mw = @($actionItems | Where-Object { $_.Severity -eq 'warning'  -and $_.Category -like "$($_mod.Prefix)*" }).Count
    $_dotClass  = if ($_mc -gt 0) { 'dot-critical' } elseif ($_mw -gt 0) { 'dot-warn' } else { 'dot-ok' }
    $_badgeHtml = if ($_mc -gt 0) { "<span class='sb-badge'>$_mc</span>" } elseif ($_mw -gt 0) { "<span class='sb-badge warn'>$_mw</span>" } else { '' }
    $_subLinks  = ($_mod.Subs | ForEach-Object { "<a class='sb-sub' href='#$($_.Id)'>$($_.Label)</a>" }) -join ''
    "<div class='sb-module-group'><a class='sb-item' href='#$($_mod.Id)'><span class='sb-dot $_dotClass'></span>$($_mod.Label)$_badgeHtml</a><div class='sb-sub-group'>$_subLinks</div></div>"
}

# --- Emit KPI strip + layout wrapper + sidebar + main open ---
$html.Add(@"
<div class='kpi-strip'>
  <div class='kpi-card'><div class='kpi-value $_kpiMfaClass'>$_kpiMfaStr</div><div class='kpi-label'>MFA Coverage</div><div class='kpi-sub'>$_kpiMfaSub</div></div>
  <div class='kpi-card'><div class='kpi-value $_kpiScoreClass'>$_kpiScoreVal</div><div class='kpi-label'>Identity Secure Score</div><div class='kpi-sub'>$_kpiScoreSub</div></div>
  <div class='kpi-card'><div class='kpi-value $_kpiDevClass'>$_kpiDevStr</div><div class='kpi-label'>Managed Devices</div><div class='kpi-sub'>$_kpiDevSub</div></div>
  <div class='kpi-card'><div class='kpi-value $_kpiStorageClass'>$_kpiStorageStr</div><div class='kpi-label'>Tenant Storage</div><div class='kpi-sub'>$_kpiStorageSub</div></div>
  <div class='kpi-card'><div class='kpi-value $_kpiAiClass'>$_kpiAiStr</div><div class='kpi-label'>Action Items</div><div class='kpi-sub'>$_kpiAiSub</div></div>
</div>
</div><!-- /sticky-header -->
<div class='layout'>
  <nav class='sidebar'>
    <a class='sb-item' href='#action-items'><span class='sb-dot dot-neutral'></span>Action Items</a>
    <a class='sb-item' href='#compliance-overview'><span class='sb-dot dot-neutral'></span>Compliance Overview</a>
    <hr class='sb-divider'>
    <div class='sb-section-label'>Modules</div>
    $($_sbItemsHtml -join "`n    ")
    <hr class='sb-divider'>
    <a class='sb-item' href='#tech-issues'><span class='sb-dot dot-neutral'></span>Technical Issues</a>
    <a class='sb-item' href='$(([System.IO.Path]::GetRelativePath($script:ReportBaseDir, (Join-Path $AuditFolder "Raw")) -replace '\\', '/'))' target='_blank'><span class='sb-dot dot-neutral'></span>Raw CSV Files</a>
  </nav>
  <main class='main'>
    <div class='content-area'>
"@)

$_companyCard = if ($script:companyCardHtml) { $script:companyCardHtml } else { '' }
$html.Add($_companyCard)

# --- Render Action Items block ---
if ($actionItems.Count -gt 0) {
    $_critItems = @($actionItems | Where-Object { $_.Severity -eq 'critical' } | Sort-Object @{ Expression = { Get-ActionItemModuleSortOrder -Category $_.Category } }, @{ Expression = { $_.Sequence } })
    $_warnItems = @($actionItems | Where-Object { $_.Severity -eq 'warning'  } | Sort-Object @{ Expression = { Get-ActionItemModuleSortOrder -Category $_.Category } }, @{ Expression = { $_.Sequence } })

    $_critRows = foreach ($ai in $_critItems) {
        $docLink = if ($ai.DocUrl) { "<a class='ai-doc' href='$($ai.DocUrl)' target='_blank'>&#128279; Docs</a>" } else { "" }
        "<div class='ai-row'><div class='ai-cat'>$($ai.Category)</div><div class='ai-text'>$($ai.Text)</div>$docLink</div>"
    }
    $_warnRows = foreach ($ai in $_warnItems) {
        $docLink = if ($ai.DocUrl) { "<a class='ai-doc' href='$($ai.DocUrl)' target='_blank'>&#128279; Docs</a>" } else { "" }
        "<div class='ai-row'><div class='ai-cat'>$($ai.Category)</div><div class='ai-text'>$($ai.Text)</div>$docLink</div>"
    }

    $_critPanel = if ($_critItems.Count -gt 0) { @"
<div class='ai-panel critical'>
  <div class='ai-panel-header'>&#9889; Critical Issues ($($_critItems.Count))</div>
  <div class='ai-panel-body'>$($_critRows -join '')</div>
</div>
"@ } else { "" }

    $_warnPanel = if ($_warnItems.Count -gt 0) { @"
<div class='ai-panel warning'>
  <div class='ai-panel-header'>&#9888; Warnings ($($_warnItems.Count))</div>
  <div class='ai-panel-body'>$($_warnRows -join '')</div>
</div>
"@ } else { "" }

    $_aiTitle = "&#9889; Action Items"
    if ($_critItems.Count -gt 0 -and $_warnItems.Count -gt 0) {
        $_aiTitle += " &mdash; $($_critItems.Count) critical, $($_warnItems.Count) warning$(if ($_warnItems.Count -ne 1) { 's' })"
    } elseif ($_critItems.Count -gt 0) {
        $_aiTitle += " &mdash; $($_critItems.Count) critical"
    } else {
        $_aiTitle += " &mdash; $($_warnItems.Count) warning$(if ($_warnItems.Count -ne 1) { 's' })"
    }

    $html.Add(@"
<div class='ai-section' id='action-items'>
  <div class='ai-section-hdr' onclick='toggleModule(this)'>
    <span class='ai-section-title'>$_aiTitle</span>
    <span class='module-toggle open'>&#9658;</span>
  </div>
  <div class='ai-section-body'>
    <div class='ai-grid'>$_critPanel$_warnPanel</div>
  </div>
</div>
"@)
}
else {
    $html.Add("<p class='ai-none'>&#10003; No issues identified. All checked areas meet best-practice recommendations.</p>")
}


# --- ScubaGear checks ---
if ($_scubaResults) {
    $_sgCategoryMap = @{
        'AAD'          = 'ScubaGear / Identity'
        'EXO'          = 'ScubaGear / Exchange'
        'SharePoint'   = 'ScubaGear / SharePoint'
        'Teams'        = 'ScubaGear / Teams'
        'Defender'     = 'ScubaGear / Defender'
        'PowerPlatform'= 'ScubaGear / Power Platform'
    }
    foreach ($_sgProd in $_scubaResults.Results.PSObject.Properties) {
        $_sgCat = if ($_sgCategoryMap.ContainsKey($_sgProd.Name)) { $_sgCategoryMap[$_sgProd.Name] } else { "ScubaGear / $($_sgProd.Name)" }
        foreach ($_sgGroup in $_sgProd.Value) {
            foreach ($_sgCtrl in $_sgGroup.Controls) {
                $_sgSeverity = switch ($_sgCtrl.Result) {
                    'Fail'    { if ($_sgCtrl.Criticality -like 'Shall*') { 'critical' } elseif ($_sgCtrl.Criticality -like 'Should*') { 'warning' } else { $null } }
                    'Warning' { 'warning' }
                    default   { $null }
                }
                if ($_sgSeverity) {
                    $_sgId   = $_sgCtrl.'Control ID'
                    $_sgText = "$(ConvertTo-HtmlText $_sgCtrl.Requirement) [$_sgId]"
                    Add-ActionItem -Severity $_sgSeverity -Category $_sgCat -Text $_sgText
                }
            }
        }
    }
}

# =========================================
# ===   Compliance Overview             ===
# =========================================
$_covCritCount = @($actionItems | Where-Object { $_.Severity -eq 'critical' }).Count
$_covWarnCount = @($actionItems | Where-Object { $_.Severity -eq 'warning'  }).Count

# Extract CIS control IDs from action item text (pattern: CIS N.N.N)
$_covCisIds = @($actionItems | ForEach-Object {
    if ($_.Text -match '\(CIS ([\d.]+)\)') { $Matches[1] }
} | Sort-Object -Unique)

# Approximate total distinct checks. A full M365 audit covers ~150 controls across all modules.
# Ensure the total always exceeds issues found so Passed is never artificially zero.
$_covTotalChecks = [math]::Max(150, $_covCritCount + $_covWarnCount + 30)
$_covPassCount   = $_covTotalChecks - $_covCritCount - $_covWarnCount

$_covTotal  = $_covCritCount + $_covWarnCount + $_covPassCount
$_covPctPass = if ($_covTotal -gt 0) { [math]::Round(($_covPassCount / $_covTotal) * 100) } else { 100 }
$_covPctWarn = if ($_covTotal -gt 0) { [math]::Round(($_covWarnCount / $_covTotal) * 100) } else { 0 }
$_covPctFail = if ($_covTotal -gt 0) { [math]::Round(($_covCritCount / $_covTotal) * 100) } else { 0 }
$_covPctNa   = [math]::Max(0, 100 - $_covPctPass - $_covPctWarn - $_covPctFail)

$_covBarHtml = "<div class='cov-bar-wrap'><div class='cov-bar'>" +
    "<div class='cov-bar-pass' style='width:${_covPctPass}%' title='Passed: $_covPassCount'></div>" +
    "<div class='cov-bar-warn' style='width:${_covPctWarn}%' title='Warnings: $_covWarnCount'></div>" +
    "<div class='cov-bar-fail' style='width:${_covPctFail}%' title='Critical: $_covCritCount'></div>" +
    "<div class='cov-bar-na'   style='width:${_covPctNa}%'   title='Not assessed'></div>" +
    "</div><div class='cov-legend'>" +
    "<span class='cov-legend-item'><span class='cov-legend-dot' style='background:#2e7d32'></span>Passed ($_covPassCount)</span>" +
    "<span class='cov-legend-item'><span class='cov-legend-dot' style='background:#f9a825'></span>Warnings ($_covWarnCount)</span>" +
    "<span class='cov-legend-item'><span class='cov-legend-dot' style='background:#c62828'></span>Critical ($_covCritCount)</span>" +
    "</div></div>"

# Group action items by module for a per-section breakdown
$_covModules = @(
    @{ Name = 'Entra';        Prefix = 'Entra'        }
    @{ Name = 'Exchange';     Prefix = 'Exchange'      }
    @{ Name = 'SharePoint';   Prefix = 'SharePoint'    }
    @{ Name = 'Mail Security'; Prefix = 'Mail Security' }
    @{ Name = 'Intune';       Prefix = 'Intune'        }
    @{ Name = 'Teams';        Prefix = 'Teams'         }
)
$_covModuleRows = foreach ($_cm in $_covModules) {
    $_cmCrit = @($actionItems | Where-Object { $_.Severity -eq 'critical' -and $_.Category -like "$($_cm.Prefix)*" }).Count
    $_cmWarn = @($actionItems | Where-Object { $_.Severity -eq 'warning'  -and $_.Category -like "$($_cm.Prefix)*" }).Count
    if ($_cmCrit -gt 0) { $_cmCell = "<span class='issue-sev-critical'>$_cmCrit critical</span>$(if ($_cmWarn -gt 0){ ", <span class='issue-sev-warning'>$_cmWarn warning</span>" })" }
    elseif ($_cmWarn -gt 0) { $_cmCell = "<span class='issue-sev-warning'>$_cmWarn warning</span>" }
    else { $_cmCell = "<span style='color:#2e7d32'>&#10003; No issues</span>" }
    "<tr><td>$($_cm.Name)</td><td>$_cmCell</td></tr>"
}

$_covCisHtml = if ($_covCisIds.Count -gt 0) {
    "<p style='font-size:0.85rem;color:#555;margin-top:0.75rem'><b>CIS controls with findings:</b> " + ($_covCisIds -join ', ') + "</p>"
} else { "" }

$_covSectionHtml = @"
<section class='module' id='compliance-overview'>
  <div class='module-hdr' onclick='toggleModule(this)'>
    <span class='module-title'>Compliance Overview</span>
    <span class='module-toggle open'>&#9658;</span>
  </div>
  <div class='module-body'>
    <p style='margin-bottom:0.5rem'>Distribution of checked controls across all audit modules. <em>Passed</em> is an approximation based on checks that did not raise action items.</p>
    $_covBarHtml
    <table class='summary-table'>
      <thead><tr><th>Module</th><th>Status</th></tr></thead>
      <tbody>$($_covModuleRows -join "`n      ")</tbody>
    </table>
    $_covCisHtml
  </div>
</section>
"@
$html.Add($_covSectionHtml)


# =========================================
# ===   Entra Section                   ===
# =========================================
$entraFiles = @(Get-ChildItem "$entraDir\Entra_*.csv" -ErrorAction SilentlyContinue)

if ($entraFiles.Count -gt 0) {
    $entraSummary = [System.Collections.Generic.List[string]]::new()

    # --- Section header stat chips ---
    $_esUserStr  = if ($null -ne $_aiTotal)   { "$_aiTotal" }       else { '&mdash;' }
    $_esMfaStr   = if ($null -ne $_aiPct)     { "${_aiPct}%" }      else { '&mdash;' }
    $_esGaStr    = if ($null -ne $_aiGaCount) { "$_aiGaCount" }     else { '&mdash;' }
    $_esCritStr  = @($actionItems | Where-Object { $_.Category -like 'Entra*' -and $_.Severity -eq 'critical' }).Count
    $_esWarnStr  = @($actionItems | Where-Object { $_.Category -like 'Entra*' -and $_.Severity -eq 'warning'  }).Count
    $_esMfaClass = if ($null -eq $_aiPct -or $_aiPct -eq 100) { 'ok' } elseif ($_aiPct -ge 80) { 'warn' } else { 'critical' }
    $_esGaClass  = if ($null -eq $_aiGaCount -or ($_aiGaCount -ge 2 -and $_aiGaCount -le 4)) { 'ok' } else { 'warn' }
    $_esAiClass  = if ($_esCritStr -gt 0) { 'critical' } elseif ($_esWarnStr -gt 0) { 'warn' } else { 'ok' }
    $_esAiStr    = if ($_esCritStr -gt 0) { "$_esCritStr critical" } elseif ($_esWarnStr -gt 0) { "$_esWarnStr warnings" } else { 'None' }
    $entraSummary.Add(@"
<div class='section-stats'>
  <a class='stat-chip neutral' href='#entra-users'><div class='stat-chip-value'>$_esUserStr</div><div class='stat-chip-label'>Licensed Users</div></a>
  <a class='stat-chip $_esMfaClass' href='#entra-users'><div class='stat-chip-value'>$_esMfaStr</div><div class='stat-chip-label'>MFA Coverage</div></a>
  <a class='stat-chip $_esGaClass' href='#entra-admins'><div class='stat-chip-value'>$_esGaStr</div><div class='stat-chip-label'>Global Admins</div></a>
  <a class='stat-chip $_esAiClass' href='#entra'><div class='stat-chip-value'>$_esAiStr</div><div class='stat-chip-label'>Action Items</div></a>
</div>
"@)

    # Determine audit window from licence tier (drives all "last N days" labels)
    $_auditPremiumSkus = @("AAD_PREMIUM", "AAD_PREMIUM_P2", "ENTERPRISEPREMIUM", "ENTERPRISEPACK",
                           "EMS", "EMS_PREMIUM", "SPB", "O365_BUSINESS_PREMIUM", "M365_F3", "IDENTITY_GOVERNANCE")
    $auditWindowDays = 7
    $_licCheck = Join-Path $entraDir "Entra_Licenses.csv"
    if (Test-Path $_licCheck) {
        $_skuList = @(Import-Csv $_licCheck | Select-Object -ExpandProperty SkuPartNumber)
        if (($_skuList | Where-Object { $_ -in $_auditPremiumSkus }).Count -gt 0) { $auditWindowDays = 30 }
    }
    $auditWindowLabel = "last $auditWindowDays days"

    # --- Identity Secure Score ---
    $secureScoreCsv = Join-Path $entraDir "Entra_SecureScore.csv"
    if (Test-Path $secureScoreCsv) {
        $ss = Import-Csv $secureScoreCsv | Select-Object -First 1
        if ($ss) {
            $ssPct   = [double]$ss.Percentage
            $ssColor = if ($ssPct -ge 80) { '#27ae60' } elseif ($ssPct -ge 50) { '#f39c12' } else { '#e74c3c' }
            $ssBar   = "<div style='background:#e0e0e0;border-radius:4px;height:14px;width:100%;max-width:300px;overflow:hidden;display:inline-block;vertical-align:middle;margin-right:8px'><div style='background:$ssColor;width:$($ssPct)%;height:14px'></div></div>"
            $entraSummary.Add("<h4 id='entra-score'>Secure Score</h4>")
            $entraSummary.Add("<p><b>Identity Secure Score:</b> $($ss.CurrentScore) / $($ss.MaxScore) &nbsp;($($ss.Percentage)%) &nbsp;$ssBar<span style='color:#888;font-size:0.85em'>as of $($ss.Date)</span></p>")
        }
    }

    # --- Secure Score Control Breakdown ---
    $secureScoreControlsCsv = Join-Path $entraDir "Entra_SecureScoreControls.csv"
    if (Test-Path $secureScoreControlsCsv) {
        $ssControls = @(Import-Csv $secureScoreControlsCsv)
        if ($ssControls.Count -gt 0) {
            $toAction    = @($ssControls | Where-Object { [double]$_.Score -le 0 } | Sort-Object ControlName)
            $implemented = @($ssControls | Where-Object { [double]$_.Score -gt 0 } | Sort-Object { -[double]$_.Score })

            $toActionRows = foreach ($c in $toAction) {
                "<tr><td style='font-size:0.85rem'>$($c.ControlName)</td><td style='font-size:0.82rem;color:#555'>$($c.Description)</td></tr>"
            }
            $implRows = foreach ($c in $implemented) {
                $scoreColor = if ([double]$c.Score -ge 5) { 'color:#27ae60' } else { 'color:#f39c12' }
                "<tr><td style='font-size:0.85rem'>$($c.ControlName)</td><td style='text-align:right;$scoreColor;font-weight:bold;width:70px'>$($c.Score)</td><td style='font-size:0.82rem;color:#555'>$($c.Description)</td></tr>"
            }

            $toActionHtml = if ($toAction.Count -gt 0) { @"
<details open>
  <summary style='cursor:pointer;font-weight:600;font-size:0.95rem;padding:0.4rem 0;color:#c0392b;background:transparent;border:none'>&#9888; To Action ($($toAction.Count) control(s) with 0 points)</summary>
  <table style='margin-top:0.4rem'>
    <thead><tr><th>Control</th><th>Description</th></tr></thead>
    <tbody>$($toActionRows -join "`n")</tbody>
  </table>
</details>
"@ } else { "<p class='ok'>All Secure Score controls have been addressed.</p>" }

            $implHtml = if ($implemented.Count -gt 0) { @"
<details style='margin-top:0.75rem'>
  <summary style='cursor:pointer;font-weight:600;font-size:0.95rem;padding:0.4rem 0;color:#27ae60;background:transparent;border:none'>&#10003; Implemented ($($implemented.Count) control(s))</summary>
  <table style='margin-top:0.4rem'>
    <thead><tr><th>Control</th><th style='text-align:right;width:70px'>Score</th><th>Description</th></tr></thead>
    <tbody>$($implRows -join "`n")</tbody>
  </table>
</details>
"@ } else { "" }

            $entraSummary.Add(@"
<h4>Secure Score — Control Breakdown</h4>
$toActionHtml
$implHtml
"@)
        }
    }

    # --- Security Defaults + SSPR (buffered — emitted inside the CA block) ---
    $_sdHtml = ""
    $secDefaultsCsv = Join-Path $entraDir "Entra_SecurityDefaults.csv"
    if (Test-Path $secDefaultsCsv) {
        $secDef = Import-Csv $secDefaultsCsv | Select-Object -First 1
        if ($secDef.SecurityDefaultsEnabled -eq "True") {
            $_sdHtml = "<p class='ok'>Security Defaults: <b>Enabled</b></p>"
        }
        else {
            $_sdCaCount = if ($null -ne $_aiEnabledCa) { @($_aiEnabledCa).Count } else { 0 }
            if ($_sdCaCount -gt 0) {
                $_sdHtml = "<p class='ok'>Security Defaults: <b>Disabled</b> — $_sdCaCount Conditional Access polic$(if ($_sdCaCount -eq 1) { 'y' } else { 'ies' }) active</p>"
            }
            else {
                $_sdHtml = "<p class='critical'>Security Defaults: <b>Disabled</b> — no Conditional Access policies are enabled</p>"
            }
        }
    }

    $_ssprHtml = ""
    $ssprCsv = Join-Path $entraDir "Entra_SSPR.csv"
    if (Test-Path $ssprCsv) {
        $ssprData = Import-Csv $ssprCsv | Select-Object -First 1
        if ($ssprData.SSPREnabled -eq "Enabled") {
            $_ssprHtml = "<p class='ok'>Self-Service Password Reset: <b>Enabled</b></p>"
        }
        else {
            $_ssprHtml = "<p class='critical'>Self-Service Password Reset: <b>$($ssprData.SSPREnabled)</b></p>"
        }
    }

    # --- MFA and User Table ---
    # Entra_Users.csv contains licensed members only (guests and unlicensed users are separate)
    $entraUsersCsv = Join-Path $entraDir "Entra_Users.csv"
    if (Test-Path $entraUsersCsv) {
        $userSummary = Import-Csv $entraUsersCsv
        $mfaTotal    = $userSummary.Count
        $mfaEnabled  = ($userSummary | Where-Object { $_.MFAEnabled -eq 'True' }).Count
        $mfaPercent  = if ($mfaTotal -gt 0) { [math]::Round(($mfaEnabled / $mfaTotal) * 100, 1) } else { 0 }
        $mfaClass    = if ($mfaPercent -eq 100) { "ok" } elseif ($mfaPercent -gt 0) { "warn" } else { "critical" }

        $entraSummary.Add("<h4 id='entra-users'>User Accounts</h4>")
        $entraSummary.Add("<p class='$mfaClass'>MFA enabled for <b>$mfaPercent%</b> of licensed users ($mfaEnabled / $mfaTotal)</p>")

        # Load sign-in history for expandable rows (keyed by UPN)
        $signInsByUpn = @{}
        $signInsCsv   = Join-Path $entraDir "Entra_SignIns.csv"
        if (Test-Path $signInsCsv) {
            foreach ($entry in (Import-Csv $signInsCsv)) {
                if (-not $signInsByUpn.ContainsKey($entry.UPN)) {
                    $signInsByUpn[$entry.UPN] = [System.Collections.Generic.List[object]]::new()
                }
                $signInsByUpn[$entry.UPN].Add($entry)
            }
        }

        $tableRows = foreach ($user in ($userSummary | Sort-Object UPN)) {
            $mfaCell = if ($user.MFAEnabled -eq "False") {
                "<td style='background:#ffebee;color:#b71c1c;font-weight:bold'>False</td>"
            } else {
                "<td>$($user.MFAEnabled)</td>"
            }
            $statusCell = if ($user.AccountStatus -eq "Blocked") {
                "<td style='background:#ffebee;color:#b71c1c;font-weight:bold'>Blocked</td>"
            } else {
                "<td>$($user.AccountStatus)</td>"
            }

            # Stale sign-in detection — highlight Last Sign-In cell if >90 days or never
            $_siDt = [datetime]::MinValue
            $_siStale = -not $user.LastSignIn -or (-not [datetime]::TryParse(($user.LastSignIn -replace ' UTC',''), [ref]$_siDt)) -or (([datetime]::UtcNow - $_siDt).TotalDays -gt 90)
            $lastSignInCell = if ($_siStale) {
                $_siLabel = if ($user.LastSignIn) { $user.LastSignIn } else { 'Never' }
                "<td style='color:#b71c1c;font-weight:bold' title='Potential stale account — no sign-in for 90+ days. Review for deprovisioning.'>$_siLabel</td>"
            } else {
                "<td>$($user.LastSignIn)</td>"
            }

            # Main user row — clickable to expand sign-in history
            $userRow = "<tr class='user-row' onclick='toggleSignIns(this)' title='Click to show/hide sign-in history'><td>$($user.UPN)</td><td>$($user.FirstName)</td><td>$($user.LastName)</td>$statusCell<td>$($user.AssignedLicense)</td>$mfaCell<td>$($user.MFAMethods)</td><td>$($user.DisablePasswordExpiration)</td><td>$($user.LastPasswordChange)</td>$lastSignInCell</tr>"

            # Hidden sign-in detail row immediately below
            $siEntries = if ($signInsByUpn.ContainsKey($user.UPN)) { @($signInsByUpn[$user.UPN]) } else { @() }
            if ($siEntries.Count -gt 0) {
                $siRows = foreach ($si in $siEntries) {
                    $siClass  = if ($si.Success -eq "True") { "signin-ok" } else { "signin-fail" }
                    $siResult = if ($si.Success -eq "True") { "Success" } else { "Failed: $($si.FailureReason)" }
                    $locParts = @($si.City, $si.Country) | Where-Object { $_ }
                    $siLoc    = $locParts -join ", "
                    "<tr><td>$($si.Timestamp)</td><td>$($si.App)</td><td>$($si.IPAddress)</td><td>$siLoc</td><td class='$siClass'>$siResult</td></tr>"
                }
                $detailRow = "<tr class='signin-detail' style='display:none'><td colspan='10'><table class='inner-table'><thead><tr><th>Time</th><th>Application</th><th>IP Address</th><th>Location</th><th>Result</th></tr></thead><tbody>$($siRows -join '')</tbody></table></td></tr>"
            }
            else {
                $detailRow = "<tr class='signin-detail' style='display:none'><td colspan='10'><em>No sign-in data available for this user</em></td></tr>"
            }

            $userRow
            $detailRow
        }

        $entraSummary.Add(@"
$(Get-ExpandHintHtml -Text 'Click a row to expand recent sign-in history.')
<table>
  <thead><tr>
    <th>UPN</th><th>First</th><th>Last</th><th>Account Status</th><th>License</th>
    <th>MFA Enabled</th><th>MFA Methods</th><th>Pwd Expiry</th>
    <th>Last Pwd Change</th><th>Last Sign-In</th>
  </tr></thead>
  <tbody>$($tableRows -join "`n")</tbody>
</table>
"@)
    }

    # Unlicensed member accounts
    $unlicensedUsersCsv = Join-Path $entraDir "Entra_Users_Unlicensed.csv"
    if (Test-Path $unlicensedUsersCsv) {
        $unlicUsers = @(Import-Csv $unlicensedUsersCsv)
        if ($unlicUsers.Count -gt 0) {
            $ulRows = ($unlicUsers | ForEach-Object {
                "<tr><td>$($_.UPN)</td><td>$($_.FirstName) $($_.LastName)</td><td>$($_.AccountStatus)</td><td>$($_.LastSignIn)</td></tr>"
            }) -join ""
            $entraSummary.Add(@"
<details>
  <summary class='warn' style='cursor:pointer'>$($unlicUsers.Count) member account(s) have no licence assigned</summary>
  <table><thead><tr><th>UPN</th><th>Name</th><th>Account Status</th><th>Last Sign-In</th></tr></thead><tbody>$ulRows</tbody></table>
</details>
"@)
        }
    }

    # --- License Summary ---
    $licensesCsv = Join-Path $entraDir "Entra_Licenses.csv"
    if (Test-Path $licensesCsv) {
        $licenses = Import-Csv $licensesCsv
        if ($licenses.Count -gt 0) {
            $licRows = foreach ($lic in $licenses) {
                "<tr><td>$($lic.SkuFriendlyName)</td><td>$($lic.EnabledUnits)</td><td>$($lic.SuspendedUnits)</td><td>$($lic.WarningUnits)</td><td>$($lic.ConsumedUnits)</td><td>$($lic.PurchaseChannel)</td></tr>"
            }
            $entraSummary.Add(@"
<h4 id='entra-licenses'>Licence Summary</h4>
<table>
  <thead><tr><th>Licence</th><th>Total</th><th>Suspended</th><th>Warning</th><th>Assigned</th><th>Channel</th></tr></thead>
  <tbody>$($licRows -join "`n")</tbody>
</table>
"@)
        }
    }

    # --- Global Admin count ---
    $globalAdminsCsv = Join-Path $entraDir "Entra_GlobalAdmins.csv"
    if (Test-Path $globalAdminsCsv) {
        $gaData  = @(Import-Csv $globalAdminsCsv)
        $gaCount = $gaData.Count
        if ($gaCount -eq 0) {
            $entraSummary.Add("<p class='critical'>No Global Administrators found</p>")
        }
        elseif ($gaCount -eq 1) {
            $gaName = $gaData[0].MemberDisplayName
            $gaUpn  = $gaData[0].MemberUserPrincipalName
            $entraSummary.Add("<p class='warn'>Only 1 Global Administrator: <b>$gaName</b> ($gaUpn) — recommend at least 2 to avoid lockout</p>")
        }
        else {
            $gaNames = ($gaData | ForEach-Object { "$($_.MemberDisplayName) ($($_.MemberUserPrincipalName))" }) -join ", "
            $entraSummary.Add("<p class='ok'>$gaCount Global Administrators: $gaNames</p>")
        }
    }

    # --- Admin role assignments table ---
    $adminRolesCsv = Join-Path $entraDir "Entra_AdminRoles.csv"
    if (Test-Path $adminRolesCsv) {
        $adminRoles  = Import-Csv $adminRolesCsv
        $roleCount   = ($adminRoles | Select-Object -ExpandProperty RoleName -Unique).Count
        $memberCount = ($adminRoles | Select-Object -ExpandProperty MemberUserPrincipalName -Unique).Count

        $roleRows = foreach ($assignment in ($adminRoles | Sort-Object RoleName, MemberDisplayName)) {
            "<tr><td>$($assignment.RoleName)</td><td>$($assignment.MemberDisplayName)</td><td>$($assignment.MemberUserPrincipalName)</td></tr>"
        }
        $entraSummary.Add(@"
<h4 id='entra-admins'>Admin Role Assignments ($memberCount user(s) across $roleCount role(s))</h4>
<table>
  <thead><tr><th>Role</th><th>User</th><th>UPN</th></tr></thead>
  <tbody>$($roleRows -join "`n")</tbody>
</table>
"@)
    }

    # --- Authentication Methods Policy ---
    $_amCsv = Join-Path $entraDir "Entra_AuthMethodsPolicy.csv"
    if (Test-Path $_amCsv) {
        $_amMethods = @(Import-Csv $_amCsv)
        if ($_amMethods.Count -gt 0) {
            $_amRows = foreach ($method in ($_amMethods | Sort-Object MethodType)) {
                $stateClass = if ($method.State -eq 'enabled') { 'ok' } else { '' }
                "<tr><td>$(ConvertTo-HtmlText $method.MethodType)</td><td class='$stateClass'>$($method.State)</td><td>$($method.IsRegistrationRequired)</td></tr>"
            }
            $entraSummary.Add(@"
<h4 id='entra-authmethods'>Authentication Methods Policy</h4>
<table>
  <thead><tr><th>Method</th><th>State</th><th>Registration Required</th></tr></thead>
  <tbody>$($_amRows -join "`n")</tbody>
</table>
"@)
        }
    }

    # --- External Collaboration Settings ---
    $_ecCsv = Join-Path $entraDir "Entra_ExternalCollab.csv"
    if (Test-Path $_ecCsv) {
        $_ec = Import-Csv $_ecCsv | Select-Object -First 1
        if ($_ec) {
            $_inviteLabels = @{
                'everyone'                           = 'Everyone, including unauthenticated internet users'
                'adminsGuestInvitersAndAllMembers'   = 'Admins, Guest Inviters, and all members'
                'adminsAndGuestInviters'             = 'Admins and users in the Guest Inviter role only'
                'adminsAndSingleUserMemberCanInvite' = 'Admins and specific members'
                'none'                               = 'No one (most restrictive)'
            }
            $_inviteDisplay = if ($_inviteLabels.ContainsKey($_ec.AllowInvitesFrom)) {
                $_inviteLabels[$_ec.AllowInvitesFrom]
            } else { $_ec.AllowInvitesFrom }
            $_inviteClass = switch ($_ec.AllowInvitesFrom) {
                'everyone'                         { 'critical' }
                'adminsGuestInvitersAndAllMembers' { 'warn' }
                default                            { 'ok' }
            }
            $_knownRoles = @{
                '10dae51f-b6af-4016-8d66-8c2a99b929b7' = 'Guest User (limited directory access — standard)'
                '2af84b1e-32c8-42b7-82bc-daa82404023b' = 'Restricted Guest User (minimal access — most secure)'
                'a0b1b346-4d3e-4e8b-98f8-753987be4970' = 'Member (full directory access — least secure)'
            }
            $_roleDisplay = if ($_knownRoles.ContainsKey($_ec.GuestUserRoleName)) {
                $_knownRoles[$_ec.GuestUserRoleName]
            } elseif ($_ec.GuestUserRoleName -match '^[0-9a-fA-F]{8}-') {
                # Raw GUID — not in known map; resolve via role name if available
                "Unknown role ID: $($_ec.GuestUserRoleName)"
            } else {
                $_ec.GuestUserRoleName
            }
            $_emailSub  = if ($_ec.AllowedToSignUpEmailBasedSubscriptions)       { $_ec.AllowedToSignUpEmailBasedSubscriptions }       else { '<span style=''color:#888''>Not available</span>' }
            $_emailJoin = if ($_ec.AllowEmailVerifiedUsersToJoinOrganization)     { $_ec.AllowEmailVerifiedUsersToJoinOrganization }     else { '<span style=''color:#888''>Not available</span>' }
            $entraSummary.Add(@"
<h4 id='entra-extcollab'>External Collaboration Settings</h4>
<table style='max-width:700px'>
  <thead><tr><th>Setting</th><th>Value</th></tr></thead>
  <tbody>
    <tr><td>Who can invite guests</td><td class='$_inviteClass'>$(ConvertTo-HtmlText $_inviteDisplay)</td></tr>
    <tr><td>Guest user permissions level</td><td>$(ConvertTo-HtmlText $_roleDisplay)</td></tr>
    <tr><td>Email-based subscription sign-up</td><td>$_emailSub</td></tr>
    <tr><td>Email-verified users can join org</td><td>$_emailJoin</td></tr>
  </tbody>
</table>
"@)
        }
    }

    # --- App Registrations ---
    $_arCsv = Join-Path $entraDir "Entra_AppRegistrations.csv"
    if (Test-Path $_arCsv) {
        $_arAll = @(Import-Csv $_arCsv)
        if ($_arAll.Count -gt 0) {
            # Load permissions lookup (keyed by AppDisplayName)
            $_arPermCsv  = Join-Path $entraDir "Entra_AppRegistrationPermissions.csv"
            $_arPermByApp = @{}
            if (Test-Path $_arPermCsv) {
                foreach ($_p in (Import-Csv $_arPermCsv)) {
                    if (-not $_arPermByApp.ContainsKey($_p.AppDisplayName)) { $_arPermByApp[$_p.AppDisplayName] = [System.Collections.Generic.List[object]]::new() }
                    $_arPermByApp[$_p.AppDisplayName].Add($_p)
                }
            }
            $_arUniqueNames = ($_arAll | Select-Object -ExpandProperty DisplayName -Unique)
            $_arRows = foreach ($reg in ($_arAll | Sort-Object DisplayName, CredentialExpiry)) {
                $daysVal  = if ($reg.DaysUntilExpiry -ne '') { [int]$reg.DaysUntilExpiry } else { $null }
                $daysCell = if ($null -eq $daysVal) {
                    "<td>&mdash;</td>"
                } elseif ($daysVal -lt 0) {
                    "<td style='background:#ffebee;color:#b71c1c;font-weight:bold'>Expired ($([Math]::Abs($daysVal))d ago)</td>"
                } elseif ($daysVal -le 30) {
                    "<td style='background:#fff8e1;color:#e65100;font-weight:bold'>$daysVal days</td>"
                } else {
                    "<td>$daysVal days</td>"
                }
                # Permissions cell — only rendered on the first row for each app name
                $_permCell = if ($_arPermByApp.ContainsKey($reg.DisplayName)) {
                    $_perms = @($_arPermByApp[$reg.DisplayName])
                    $_permLines = ($_perms | Sort-Object PermissionType, ResourceApp, PermissionName | ForEach-Object {
                        "<span style='font-size:0.78rem;display:block'><b>$($_.PermissionType)</b> &mdash; $($_.ResourceApp) / $(ConvertTo-HtmlText $_.PermissionName)</span>"
                    }) -join ''
                    "<td><details><summary style='cursor:pointer;font-size:0.8rem;color:#475569'>$($_perms.Count) permission(s)</summary><div style='padding-top:4px'>$_permLines</div></details></td>"
                } else {
                    "<td style='color:#888;font-size:0.82rem'>—</td>"
                }
                "<tr><td>$(ConvertTo-HtmlText $reg.DisplayName)</td><td><code>$($reg.AppId)</code></td><td>$($reg.CredentialType)</td><td>$(ConvertTo-HtmlText $reg.CredentialName)</td><td>$($reg.CredentialExpiry)</td>$daysCell$_permCell</tr>"
            }
            $entraSummary.Add(@"
<h4 id='entra-appregs'>App Registrations ($( ($_arUniqueNames).Count ))</h4>
<table>
  <thead><tr><th>Name</th><th>App ID</th><th>Credential Type</th><th>Credential Name</th><th>Expiry</th><th>Days Until Expiry</th><th>Permissions</th></tr></thead>
  <tbody>$($_arRows -join "`n")</tbody>
</table>
"@)
        }
    }

    # --- PIM Role Assignments ---
    $_pimCsvHtml = Join-Path $entraDir "Entra_PIMAssignments.csv"
    if (Test-Path $_pimCsvHtml) {
        $_pimAll = @(Import-Csv $_pimCsvHtml)
        if ($_pimAll.Count -gt 0) {
            $_pimRows = foreach ($assignment in ($_pimAll | Sort-Object RoleName, PrincipalDisplayName)) {
                $typeClass = if ($assignment.AssignmentType -eq 'Eligible') { 'ok' } else { 'warn' }
                "<tr><td>$(ConvertTo-HtmlText $assignment.RoleName)</td><td>$(ConvertTo-HtmlText $assignment.PrincipalDisplayName)</td><td>$(ConvertTo-HtmlText $assignment.PrincipalUPN)</td><td class='$typeClass'>$($assignment.AssignmentType)</td><td>$($assignment.MemberType)</td><td>$($assignment.StartDateTime)</td><td>$($assignment.EndDateTime)</td></tr>"
            }
            $entraSummary.Add(@"
<h4 id='entra-pim'>PIM Role Assignments ($($_pimAll.Count))</h4>
<table>
  <thead><tr><th>Role</th><th>Principal</th><th>UPN</th><th>Assignment Type</th><th>Member Type</th><th>Start</th><th>End</th></tr></thead>
  <tbody>$($_pimRows -join "`n")</tbody>
</table>
"@)
        }
    }

    # --- Conditional Access Policies ---
    $caPoliciesCsv  = Join-Path $entraDir "Entra_CA_Policies.csv"
    # SKUs that include Azure AD Premium P1 or higher (required for Conditional Access)
    $caCapableSkus  = @("AAD_PREMIUM", "AAD_PREMIUM_P2", "ENTERPRISEPREMIUM", "ENTERPRISEPACK",
                        "EMS", "EMS_PREMIUM", "SPB", "O365_BUSINESS_PREMIUM", "M365_F3", "IDENTITY_GOVERNANCE")

    if (Test-Path $caPoliciesCsv) {
        $caPolicies = @(Import-Csv $caPoliciesCsv)

        if ($caPolicies.Count -gt 0) {
            $caEnabled  = ($caPolicies | Where-Object { $_.State -eq "enabled" }).Count
            $caReport   = ($caPolicies | Where-Object { $_.State -eq "enabledForReportingButNotEnforced" }).Count
            $caDisabled = ($caPolicies | Where-Object { $_.State -eq "disabled" }).Count
            $caClass    = if ($caEnabled -gt 0) { "ok" } else { "warn" }

            $entraSummary.Add("<h4 id='entra-ca'>Conditional Access</h4>")
            if ($_sdHtml)   { $entraSummary.Add($_sdHtml) }
            if ($_ssprHtml) { $entraSummary.Add($_ssprHtml) }
            $entraSummary.Add("<p class='$caClass'>$($caPolicies.Count) Conditional Access policies: <b>$caEnabled enabled</b>, $caReport report-only, $caDisabled disabled</p>")

            $entraSummary.Add((Get-ExpandHintHtml -Text 'Click a row to expand policy scope and conditions.'))
            $caRows = foreach ($policy in ($caPolicies | Sort-Object Name)) {
                $stateClass = switch ($policy.State) {
                    "enabled"                           { "ok" }
                    "enabledForReportingButNotEnforced" { "warn" }
                    "disabled"                          { "critical" }
                    default                             { "" }
                }
                $stateLabel = switch ($policy.State) {
                    "enabled"                           { "Enabled" }
                    "enabledForReportingButNotEnforced" { "Report Only" }
                    "disabled"                          { "Disabled" }
                    default                             { $policy.State }
                }
                $mfaIcon = if ($policy.RequiresMFA -eq "True") { "Yes" } else { "-" }
                $grantControls     = if ($policy.GrantControls       -and $policy.GrantControls       -ne '') { $policy.GrantControls } else { '—' }
                $includeUsers      = if ($policy.IncludeUsers        -and $policy.IncludeUsers        -ne '') { $policy.IncludeUsers } else { '—' }
                $excludeUsers      = if ($policy.ExcludeUsers        -and $policy.ExcludeUsers        -ne '') { $policy.ExcludeUsers } else { '—' }
                $includeGroups     = if ($policy.IncludeGroups       -and $policy.IncludeGroups       -ne '') { $policy.IncludeGroups } else { '—' }
                $excludeGroups     = if ($policy.ExcludeGroups       -and $policy.ExcludeGroups       -ne '') { $policy.ExcludeGroups } else { '—' }
                $includeRoles      = if ($policy.IncludeRoles        -and $policy.IncludeRoles        -ne '') { $policy.IncludeRoles } else { '—' }
                $excludeRoles      = if ($policy.ExcludeRoles        -and $policy.ExcludeRoles        -ne '') { $policy.ExcludeRoles } else { '—' }
                $includeApps       = if ($policy.IncludeApplications -and $policy.IncludeApplications -ne '') { $policy.IncludeApplications } else { '—' }
                $excludeApps       = if ($policy.ExcludeApplications -and $policy.ExcludeApplications -ne '') { $policy.ExcludeApplications } else { '—' }
                $userActions       = if ($policy.UserActions         -and $policy.UserActions         -ne '') { $policy.UserActions } else { '—' }
                $clientTypes       = if ($policy.ClientAppTypes      -and $policy.ClientAppTypes      -ne '') { $policy.ClientAppTypes } else { 'All client apps' }
                $includePlatforms  = if ($policy.IncludePlatforms    -and $policy.IncludePlatforms    -ne '') { $policy.IncludePlatforms } else { '—' }
                $excludePlatforms  = if ($policy.ExcludePlatforms    -and $policy.ExcludePlatforms    -ne '') { $policy.ExcludePlatforms } else { '—' }
                $includeLocations  = if ($policy.IncludeLocations    -and $policy.IncludeLocations    -ne '') { $policy.IncludeLocations } else { '—' }
                $excludeLocations  = if ($policy.ExcludeLocations    -and $policy.ExcludeLocations    -ne '') { $policy.ExcludeLocations } else { '—' }
                $signInRiskLevels  = if ($policy.SignInRiskLevels    -and $policy.SignInRiskLevels    -ne '') { $policy.SignInRiskLevels } else { '—' }
                $userRiskLevels    = if ($policy.UserRiskLevels      -and $policy.UserRiskLevels      -ne '') { $policy.UserRiskLevels } else { '—' }
                $deviceFilter      = if ($policy.DeviceFilter        -and $policy.DeviceFilter        -ne '') { $policy.DeviceFilter } else { '—' }
                $grantOperator     = if ($policy.GrantOperator       -and $policy.GrantOperator       -ne '') { $policy.GrantOperator } else { '—' }

                $mainRow = "<tr class='user-row' onclick='togglePerms(this)' title='Click to expand policy details'><td>$(ConvertTo-HtmlText $policy.Name)</td><td class='$stateClass'>$(ConvertTo-HtmlText $stateLabel)</td><td>$(ConvertTo-HtmlText $grantControls)</td><td>$mfaIcon</td></tr>"

                $detailFields = @(
                    @{ Label = 'Include Users';        Value = $includeUsers }
                    @{ Label = 'Exclude Users';        Value = $excludeUsers }
                    @{ Label = 'Include Groups';       Value = $includeGroups }
                    @{ Label = 'Exclude Groups';       Value = $excludeGroups }
                    @{ Label = 'Include Roles';        Value = $includeRoles }
                    @{ Label = 'Exclude Roles';        Value = $excludeRoles }
                    @{ Label = 'Include Applications'; Value = $includeApps }
                    @{ Label = 'Exclude Applications'; Value = $excludeApps }
                    @{ Label = 'User Actions';         Value = $userActions }
                    @{ Label = 'Client App Types';     Value = $clientTypes }
                    @{ Label = 'Include Platforms';    Value = $includePlatforms }
                    @{ Label = 'Exclude Platforms';    Value = $excludePlatforms }
                    @{ Label = 'Include Locations';    Value = $includeLocations }
                    @{ Label = 'Exclude Locations';    Value = $excludeLocations }
                    @{ Label = 'Sign-In Risk Levels';  Value = $signInRiskLevels }
                    @{ Label = 'User Risk Levels';     Value = $userRiskLevels }
                    @{ Label = 'Device Filter';        Value = $deviceFilter }
                    @{ Label = 'Grant Controls';       Value = $grantControls }
                    @{ Label = 'Grant Requirement';    Value = $grantOperator }
                )

                $detailRows = foreach ($field in $detailFields) {
                    if (-not [string]::IsNullOrWhiteSpace($field.Value) -and $field.Value -ne '—') {
                        "<tr><td>$(ConvertTo-HtmlText $field.Label)</td><td>$(ConvertTo-HtmlMultilineText $field.Value)</td></tr>"
                    }
                }

                if (-not $detailRows -or @($detailRows).Count -eq 0) {
                    $detailRows = "<tr><td colspan='2'><span class='detail-empty'>No additional policy conditions exported.</span></td></tr>"
                }

                $detailRow = "<tr class='signin-detail' style='display:none'><td colspan='4'><table class='inner-table'>
  <thead><tr><th style='width:160px'>Setting</th><th>Value</th></tr></thead>
  <tbody>
    $($detailRows -join "`n")
  </tbody>
</table></td></tr>"
                $mainRow
                $detailRow
            }
            $entraSummary.Add(@"
<table>
  <thead><tr><th>Policy Name</th><th>State</th><th>Grant Controls</th><th>Requires MFA</th></tr></thead>
  <tbody>$($caRows -join "`n")</tbody>
</table>
"@)
        }
        else {
            # File exists but no policies — check whether tenant has a CA-capable licence
            $licensesCsv  = Join-Path $entraDir "Entra_Licenses.csv"
            $hasCALicense = $false
            if (Test-Path $licensesCsv) {
                $tenantSkus   = @(Import-Csv $licensesCsv | Select-Object -ExpandProperty SkuPartNumber)
                $hasCALicense = ($tenantSkus | Where-Object { $_ -in $caCapableSkus }).Count -gt 0
            }

            $entraSummary.Add("<h4 id='entra-ca'>Conditional Access</h4>")
            if ($_sdHtml)   { $entraSummary.Add($_sdHtml) }
            if ($_ssprHtml) { $entraSummary.Add($_ssprHtml) }
            if ($hasCALicense) {
                $entraSummary.Add("<p class='critical'>No Conditional Access policies configured — tenant has a CA-capable licence; policies are strongly recommended</p>")
            }
            else {
                $entraSummary.Add("<p class='warn'>No Conditional Access policies found — tenant may not have Azure AD Premium P1 or higher (required for CA). Consider upgrading to M365 Business Premium or an EMS plan.</p>")
            }
        }
    }

    # --- Account Creations ---
    $creationsCsv = Join-Path $entraDir "Entra_AccountCreations.csv"
    if (Test-Path $creationsCsv) {
        $creations = @(Import-Csv $creationsCsv)
        if ($creations.Count -gt 0) {
            $creationRows = foreach ($c in $creations) {
                "<tr><td>$($c.Timestamp)</td><td>$($c.InitiatedBy)</td><td>$($c.TargetUPN)</td><td>$($c.TargetName)</td><td>$($c.Result)</td></tr>"
            }
            $entraSummary.Add("<p class='warn'>$($creations.Count) account(s) created in the $auditWindowLabel — please verify all are expected</p>")
            $entraSummary.Add(@"
<table>
  <thead><tr><th>Timestamp</th><th>Created By</th><th>Target UPN</th><th>Display Name</th><th>Result</th></tr></thead>
  <tbody>$($creationRows -join "`n")</tbody>
</table>
"@)
        }
        else {
            $entraSummary.Add("<p class='ok'>No new accounts created in the $auditWindowLabel</p>")
        }
    }

    # --- Account Deletions ---
    $deletionsCsv = Join-Path $entraDir "Entra_AccountDeletions.csv"
    if (Test-Path $deletionsCsv) {
        $deletions = @(Import-Csv $deletionsCsv)
        if ($deletions.Count -gt 0) {
            $deletionRows = foreach ($d in $deletions) {
                "<tr><td>$($d.Timestamp)</td><td>$($d.InitiatedBy)</td><td>$($d.TargetUPN)</td><td>$($d.TargetName)</td><td>$($d.Result)</td></tr>"
            }
            $entraSummary.Add("<p class='warn'>$($deletions.Count) account(s) deleted in the $auditWindowLabel — please verify all are expected</p>")
            $entraSummary.Add(@"
<table>
  <thead><tr><th>Timestamp</th><th>Deleted By</th><th>Target UPN</th><th>Display Name</th><th>Result</th></tr></thead>
  <tbody>$($deletionRows -join "`n")</tbody>
</table>
"@)
        }
        else {
            $entraSummary.Add("<p class='ok'>No accounts deleted in the $auditWindowLabel</p>")
        }
    }

    # --- Notable Audit Events (role changes + MFA/security info changes) ---
    $auditEventsCsv = Join-Path $entraDir "Entra_AuditEvents.csv"
    if (Test-Path $auditEventsCsv) {
        $auditEvts = @(Import-Csv $auditEventsCsv)
        if ($auditEvts.Count -gt 0) {
            $roleEvtCount = ($auditEvts | Where-Object { $_.Category -eq 'RoleManagement' }).Count
            $secEvtCount  = $auditEvts.Count - $roleEvtCount
            $auditRows = foreach ($evt in $auditEvts) {
                "<tr><td>$($evt.Timestamp)</td><td>$($evt.Category)</td><td>$($evt.Activity)</td><td>$($evt.InitiatedBy)</td><td>$($evt.TargetUPN)</td><td>$($evt.TargetRole)</td><td>$($evt.Result)</td></tr>"
            }
            $entraSummary.Add("<p class='warn'>$($auditEvts.Count) notable audit event(s) in the ${auditWindowLabel}: $roleEvtCount role change(s), $secEvtCount security info change(s)</p>")
            $entraSummary.Add(@"
<table>
  <thead><tr><th>Timestamp</th><th>Category</th><th>Activity</th><th>Initiated By</th><th>Target UPN</th><th>Target Role</th><th>Result</th></tr></thead>
  <tbody>$($auditRows -join "`n")</tbody>
</table>
"@)
        }
        else {
            $entraSummary.Add("<p class='ok'>No notable audit events (role changes, security info changes) in the $auditWindowLabel</p>")
        }
    }

    # --- Identity Protection ---
    $_riskyUsersCsv   = Join-Path $entraDir "Entra_RiskyUsers.csv"
    $_riskySignInsCsv = Join-Path $entraDir "Entra_RiskySignIns.csv"
    $_hasRiskyData    = (Test-Path $_riskyUsersCsv) -or (Test-Path $_riskySignInsCsv)
    if ($_hasRiskyData) {
        $entraSummary.Add("<h4 id='entra-idprot'>Identity Protection</h4>")

        if (Test-Path $_riskyUsersCsv) {
            $_ruAll     = @(Import-Csv $_riskyUsersCsv)
            $_ruAtRisk  = @($_ruAll | Where-Object { $_.RiskState -in @('atRisk','confirmedCompromised') })
            $_ruStateColor = @{ 'atRisk' = 'color:#e65100;font-weight:bold'; 'confirmedCompromised' = 'color:#b71c1c;font-weight:bold'; 'remediated' = 'color:#2e7d32'; 'dismissed' = 'color:#888' }
            $_ruLevelColor = @{ 'high' = 'color:#b71c1c;font-weight:bold'; 'medium' = 'color:#e65100;font-weight:bold'; 'low' = 'color:#f9a825' }

            if ($_ruAtRisk.Count -gt 0) {
                $entraSummary.Add("<p class='critical'>$($_ruAtRisk.Count) user(s) currently at risk or confirmed compromised.</p>")
            } else {
                $entraSummary.Add("<p class='ok'>No users currently flagged as at-risk or compromised ($($_ruAll.Count) total assessed).</p>")
            }

            if ($_ruAll.Count -gt 0) {
                $_ruRows = foreach ($_ru in ($_ruAll | Sort-Object RiskLevel, UserPrincipalName)) {
                    $_ruLStyle = if ($_ruLevelColor.ContainsKey($_ru.RiskLevel)) { " style='$($_ruLevelColor[$_ru.RiskLevel])'" } else { "" }
                    $_ruSStyle = if ($_ruStateColor.ContainsKey($_ru.RiskState)) { " style='$($_ruStateColor[$_ru.RiskState])'" } else { "" }
                    "<tr><td>$(ConvertTo-HtmlText $_ru.UserPrincipalName)</td><td>$(ConvertTo-HtmlText $_ru.DisplayName)</td><td$_ruLStyle>$($_ru.RiskLevel)</td><td$_ruSStyle>$($_ru.RiskState)</td><td>$(ConvertTo-HtmlText $_ru.RiskDetail)</td><td>$($_ru.RiskLastUpdated)</td></tr>"
                }
                $entraSummary.Add(@"
<table>
  <thead><tr><th>UPN</th><th>Name</th><th>Risk Level</th><th>Risk State</th><th>Detail</th><th>Last Updated</th></tr></thead>
  <tbody>$($_ruRows -join "`n")</tbody>
</table>
"@)
            }
        }

        if (Test-Path $_riskySignInsCsv) {
            $_rsi = @(Import-Csv $_riskySignInsCsv)
            if ($_rsi.Count -gt 0) {
                $entraSummary.Add("<h5 style='margin-top:1rem'>Risky Sign-ins ($($_rsi.Count))</h5>")
                $_rsiLevelColor = @{ 'high' = 'color:#b71c1c;font-weight:bold'; 'medium' = 'color:#e65100;font-weight:bold'; 'low' = 'color:#f9a825' }
                $_rsiRows = foreach ($_rs in ($_rsi | Sort-Object { $_.CreatedDateTime } -Descending)) {
                    $_rsLStyle = if ($_rsiLevelColor.ContainsKey($_rs.RiskLevel)) { " style='$($_rsiLevelColor[$_rs.RiskLevel])'" } else { "" }
                    "<tr><td>$(ConvertTo-HtmlText $_rs.UserPrincipalName)</td><td$_rsLStyle>$($_rs.RiskLevel)</td><td>$(ConvertTo-HtmlText $_rs.RiskState)</td><td>$(ConvertTo-HtmlText $_rs.RiskEventTypes)</td><td>$(ConvertTo-HtmlText $_rs.IPAddress)</td><td>$(ConvertTo-HtmlText "$($_rs.City), $($_rs.CountryOrRegion)")</td><td>$(ConvertTo-HtmlText $_rs.AppDisplayName)</td><td>$($_rs.CreatedDateTime)</td></tr>"
                }
                $entraSummary.Add(@"
<table>
  <thead><tr><th>UPN</th><th>Risk Level</th><th>State</th><th>Event Types</th><th>IP</th><th>Location</th><th>App</th><th>Time</th></tr></thead>
  <tbody>$($_rsiRows -join "`n")</tbody>
</table>
"@)
            } else {
                $entraSummary.Add("<p class='ok'>No risky sign-ins recorded.</p>")
            }
        }
    }

    # --- Groups ---
    $_grpCsv = Join-Path $entraDir "Entra_Groups.csv"
    if (Test-Path $_grpCsv) {
        $_grps = @(Import-Csv $_grpCsv)
        if ($_grps.Count -gt 0) {
            $_grpM365     = @($_grps | Where-Object { $_.GroupType -eq 'Microsoft 365' }).Count
            $_grpSec      = @($_grps | Where-Object { $_.GroupType -eq 'Security' }).Count
            $_grpDynamic  = @($_grps | Where-Object { $_.MembershipType -eq 'Dynamic' }).Count
            $_grpAssigned = @($_grps | Where-Object { $_.MembershipType -eq 'Assigned' }).Count
            $_grpOnPrem   = @($_grps | Where-Object { $_.OnPremSyncEnabled -eq 'True' }).Count
            $_grpRoleAssignable = @($_grps | Where-Object { $_.IsAssignableToRole -eq 'True' })
            $_grpNoOwners       = @($_grps | Where-Object { -not $_.Owners })

            $entraSummary.Add("<h4 id='entra-groups'>Groups ($($_grps.Count) total)</h4>")
            $entraSummary.Add("<p>Microsoft 365: <b>$_grpM365</b> &nbsp;|&nbsp; Security: <b>$_grpSec</b> &nbsp;|&nbsp; Dynamic: <b>$_grpDynamic</b> &nbsp;|&nbsp; Assigned: <b>$_grpAssigned</b>$(if ($_grpOnPrem -gt 0) { " &nbsp;|&nbsp; On-Prem Synced: <b>$_grpOnPrem</b>" })</p>")

            if ($_grpNoOwners.Count -gt 0) {
                $entraSummary.Add("<p class='warn'>$($_grpNoOwners.Count) group(s) have no owner assigned — these groups are unmanaged and may accumulate stale members.</p>")
            }
            if ($_grpRoleAssignable.Count -gt 0) {
                $_raNames = ($_grpRoleAssignable | ForEach-Object { ConvertTo-HtmlText $_.DisplayName }) -join ', '
                $entraSummary.Add("<p class='warn'>$($_grpRoleAssignable.Count) role-assignable group(s) — membership grants Entra directory roles: <b>$_raNames</b></p>")
            }
            $entraSummary.Add("<p style='font-size:0.82rem;color:#64748b;margin-bottom:0.5rem'><b>Role-Assignable</b> and <b>Dynamic</b> are independent properties. Role-Assignable means group membership can be used to assign Entra directory roles (a security boundary). Dynamic means membership is rule-based and automatically maintained. Both badges may appear on the same group.</p>")

            $_grpRows = foreach ($_grp in ($_grps | Sort-Object DisplayName)) {
                $_grpMemberList = if ($_grp.Members) {
                    ($_grp.Members -split '; ' | Where-Object { $_ } | ForEach-Object {
                        $m = $_
                        $mStyle = if ($m -like '*#EXT#*') { " style='color:#e65100'" } elseif ($m -match '@') { " style='color:#2e7d32'" } else { "" }
                        "<span$mStyle>$(ConvertTo-HtmlText $m)</span>"
                    }) -join '<br>'
                } else { '<em style="color:#888">No members</em>' }

                $_grpOwnerCell = if ($_grp.Owners) { ConvertTo-HtmlText $_grp.Owners } else { "<span style='color:#e65100'>None</span>" }
                $_grpRaBadge   = if ($_grp.IsAssignableToRole -eq 'True') { " <span style='background:#fff3e0;color:#e65100;border:1px solid #ffcc02;border-radius:3px;padding:1px 5px;font-size:0.78rem'>Role-Assignable</span>" } else { "" }
                $_grpDynBadge  = if ($_grp.MembershipType -eq 'Dynamic') { " <span style='background:#e3f2fd;color:#1565c0;border:1px solid #90caf9;border-radius:3px;padding:1px 5px;font-size:0.78rem'>Dynamic</span>" } else { "" }
                $_grpOpBadge   = if ($_grp.Source -eq 'On-Premises') { " <span style='background:#f3e5f5;color:#6a1b9a;border:1px solid #ce93d8;border-radius:3px;padding:1px 5px;font-size:0.78rem'>On-Prem</span>" } elseif ($_grp.Source -eq 'Cloud (Sync Stopped)') { " <span style='background:#fff3e0;color:#e65100;border:1px solid #ffcc80;border-radius:3px;padding:1px 5px;font-size:0.78rem'>Sync Stopped</span>" } else { "" }

                $memberCount = if ($_grp.Members) { @($_grp.Members -split '; ' | Where-Object { $_ }).Count } else { 0 }

                $_grpEmail = if ($_grp.Email) { "<span style='color:#888;font-size:0.88em'>$(ConvertTo-HtmlText $_grp.Email)</span>" } else { "" }
                $_grpSrc   = switch ($_grp.Source) {
                    'On-Premises'        { "<span style='color:#6a1b9a'>On-Premises</span>" }
                    'Cloud (Sync Stopped)' { "<span style='color:#e65100'>Cloud (Sync Stopped)</span>" }
                    default              { "Cloud" }
                }

                "<tr class='user-row' onclick='togglePerms(this)' title='Click to show/hide members'>" +
                "<td>$(ConvertTo-HtmlText $_grp.DisplayName)$_grpRaBadge$_grpDynBadge$_grpOpBadge<br>$_grpEmail</td>" +
                "<td>$(ConvertTo-HtmlText $_grp.GroupType)</td>" +
                "<td>$_grpSrc</td>" +
                "<td>$_grpOwnerCell</td>" +
                "<td>$memberCount</td>" +
                "</tr>" +
                "<tr class='signin-detail' style='display:none'><td colspan='5'><div style='padding:0.5rem 1rem'>$_grpMemberList</div></td></tr>"
            }
            $entraSummary.Add(@"
<table>
  <thead><tr><th>Group Name</th><th>Type</th><th>Source</th><th>Owner(s)</th><th>Members</th></tr></thead>
  <tbody>$($_grpRows -join "`n")</tbody>
</table>
"@)
        }
    }

    # --- Enterprise Apps ---
    $_eaHtmlCsv = Join-Path $entraDir "Entra_EnterpriseApps.csv"
    if (Test-Path $_eaHtmlCsv) {
        $_eaApps = @(Import-Csv $_eaHtmlCsv -Encoding UTF8)
        if ($_eaApps.Count -gt 0) {
            # Load Enterprise App permissions lookup (keyed by AppDisplayName)
            $_eaPermCsv  = Join-Path $entraDir "Entra_EnterpriseAppPermissions.csv"
            $_eaPermByApp = @{}
            if (Test-Path $_eaPermCsv) {
                foreach ($_ep in (Import-Csv $_eaPermCsv)) {
                    if (-not $_eaPermByApp.ContainsKey($_ep.AppDisplayName)) { $_eaPermByApp[$_ep.AppDisplayName] = [System.Collections.Generic.List[object]]::new() }
                    $_eaPermByApp[$_ep.AppDisplayName].Add($_ep)
                }
            }
            $_eaAvePoint = @($_eaApps | Where-Object { $_.DisplayName -like 'AvePoint*' })
            $_eaRows = foreach ($_ea in ($_eaApps | Sort-Object DisplayName)) {
                $_eaIsAvePoint = $_ea.DisplayName -like 'AvePoint*'
                $_eaStyle      = if ($_eaIsAvePoint) { " style='background:#e8f5e9'" } else { "" }
                $_eaConsented  = if ($_ea.AdminConsented -eq 'True') { "<span style='color:#c62828;font-weight:bold'>Yes</span>" } else { "No" }
                $_eaEnabled    = if ($_ea.Enabled -eq 'True') { "Yes" } else { "<span style='color:#888'>No</span>" }
                $publisher     = if ($_ea.PublisherName) { $(ConvertTo-HtmlText $_ea.PublisherName) } elseif ($_ea.PublisherDomain) { $(ConvertTo-HtmlText $_ea.PublisherDomain) } else { '<span style=''color:#888''>Unknown</span>' }
                # Permissions dropdown
                $_eaPermCell = if ($_eaPermByApp.ContainsKey($_ea.DisplayName)) {
                    $_eps = @($_eaPermByApp[$_ea.DisplayName])
                    $_epLines = ($_eps | Sort-Object PermissionType, ResourceApp, PermissionName | ForEach-Object {
                        "<span style='font-size:0.78rem;display:block'><b>$($_.PermissionType)</b> &mdash; $($_.ResourceApp) / $(ConvertTo-HtmlText $_.PermissionName)</span>"
                    }) -join ''
                    "<td><details><summary style='cursor:pointer;font-size:0.8rem;color:#475569'>$($_eps.Count) permission(s)</summary><div style='padding-top:4px'>$_epLines</div></details></td>"
                } else {
                    "<td style='color:#888;font-size:0.82rem'>—</td>"
                }
                "<tr$_eaStyle><td>$(ConvertTo-HtmlText $_ea.DisplayName)</td><td>$publisher</td><td>$_eaEnabled</td><td>$_eaConsented</td>$_eaPermCell</tr>"
            }
            $_eaAveStatus = if ($_eaAvePoint.Count -gt 0) {
                "<p class='ok'>AvePoint detected — SaaS backup service principal is present in this tenant.</p>"
            } else {
                "<p class='critical'>AvePoint not detected — no AvePoint service principal found. Confirm SaaS backup is configured.</p>"
            }
            $entraSummary.Add("<h4 id='entra-apps'>Enterprise Apps ($($_eaApps.Count) third-party)</h4>")
            $entraSummary.Add($_eaAveStatus)
            $entraSummary.Add(@"
<table>
  <thead><tr><th>App Name</th><th>Publisher</th><th>Enabled</th><th>Admin Consented</th><th>Permissions</th></tr></thead>
  <tbody>$($_eaRows -join "`n")</tbody>
</table>
"@)
        }
        else {
            $entraSummary.Add("<p class='ok'>No third-party enterprise apps found in this tenant.</p>")
        }
    }

    # Org Settings (user consent, app registration, admin consent workflow)
    $orgSettingsCsv = Join-Path $entraDir "Entra_OrgSettings.csv"
    if (Test-Path $orgSettingsCsv) {
        $_orgSettings = Import-Csv $orgSettingsCsv | Select-Object -First 1
        $entraSummary.Add("<hr class='section-divider'>")
        $entraSummary.Add("<h4 id='entra-orgsettings'>Organisation User Settings</h4>")
        $entraSummary.Add("<p style='font-size:0.82rem;color:#64748b'>Recommended values based on <a href='https://www.cisecurity.org/benchmark/microsoft_365' target='_blank'>CIS Microsoft 365 Foundations Benchmark</a> (controls 2.1.x) and the <a href='https://learn.microsoft.com/en-us/entra/identity/enterprise-apps/configure-user-consent' target='_blank'>Microsoft Security Baseline</a>.</p>")
        $entraSummary.Add("<table class='summary-table'><thead><tr><th>Setting</th><th>Value</th><th>Recommended</th></tr></thead><tbody>")
        $_orgRows = @(
            @{ Label = 'Users can register applications'; Value = $_orgSettings.AllowedToCreateApps; GoodVal = 'False'; Recommend = 'False' },
            @{ Label = 'Users can create tenants'; Value = $_orgSettings.AllowedToCreateTenants; GoodVal = 'False'; Recommend = 'False' },
            @{ Label = 'Users can create security groups'; Value = $_orgSettings.AllowedToCreateSecurityGroups; GoodVal = 'False'; Recommend = 'False (Admins only)' },
            @{ Label = 'Admin consent workflow enabled'; Value = $_orgSettings.AdminConsentWorkflowEnabled; GoodVal = 'True'; Recommend = 'True' }
        )
        foreach ($_or in $_orgRows) {
            $_isGood = $_or.Value -eq $_or.GoodVal
            $_vc = if (-not $_isGood -and $_or.Value) { "style='color:#e65100'" } else { '' }
            $entraSummary.Add("<tr><td>$($_or.Label)</td><td $_vc>$(ConvertTo-HtmlText $_or.Value)</td><td>$(ConvertTo-HtmlText $_or.Recommend)</td></tr>")
        }
        $entraSummary.Add("</tbody></table>")
    }

    $html.Add((Add-Section -Title "Microsoft Entra" -AnchorId 'entra' -CsvFiles $entraFiles.FullName -SummaryHtml ($entraSummary -join "`n") -ModuleVersion (Get-ModuleScriptVersion -ScriptName 'Invoke-EntraAudit.ps1')))
}


# =========================================
# ===   Exchange Section                ===
# =========================================
$exchangeFiles = @(Get-ChildItem "$exchangeDir\Exchange_*.csv" -ErrorAction SilentlyContinue)

if ($exchangeFiles.Count -gt 0) {
    $exchangeSummary = [System.Collections.Generic.List[string]]::new()

    # --- Section header stat chips ---
    $_exMbxCount    = 0
    $_exSharedCount = 0
    $_exMbxCsvStat  = Join-Path $exchangeDir "Exchange_Mailboxes.csv"
    if (Test-Path $_exMbxCsvStat) {
        $_exMbxAll      = @(Import-Csv $_exMbxCsvStat)
        $_exMbxCount    = $_exMbxAll.Count
        $_exSharedCount = @($_exMbxAll | Where-Object { $_.RecipientType -eq 'SharedMailbox' }).Count
    }
    $_exFwdCount  = if ($null -ne $_aiInboxRules) { $_aiInboxRules.Count } else { 0 }
    $_exAiCrit    = @($actionItems | Where-Object { $_.Category -like 'Exchange*' -and $_.Severity -eq 'critical' }).Count
    $_exAiWarn    = @($actionItems | Where-Object { $_.Category -like 'Exchange*' -and $_.Severity -eq 'warning'  }).Count
    $_exAiClass   = if ($_exAiCrit -gt 0) { 'critical' } elseif ($_exAiWarn -gt 0) { 'warn' } else { 'ok' }
    $_exAiStr     = if ($_exAiCrit -gt 0) { "$_exAiCrit critical" } elseif ($_exAiWarn -gt 0) { "$_exAiWarn warnings" } else { 'None' }
    $_exFwdClass  = if ($_exFwdCount -gt 0) { 'warn' } else { 'ok' }
    $exchangeSummary.Add(@"
<div class='section-stats'>
  <a class='stat-chip neutral' href='#exchange-mailboxes'><div class='stat-chip-value'>$_exMbxCount</div><div class='stat-chip-label'>Mailboxes</div></a>
  <a class='stat-chip neutral' href='#exchange-mailboxes'><div class='stat-chip-value'>$_exSharedCount</div><div class='stat-chip-label'>Shared Mailboxes</div></a>
  <a class='stat-chip $_exFwdClass' href='#exchange-forwarding'><div class='stat-chip-value'>$_exFwdCount</div><div class='stat-chip-label'>Forwarding Rules</div></a>
  <a class='stat-chip $_exAiClass' href='#exchange'><div class='stat-chip-value'>$_exAiStr</div><div class='stat-chip-label'>Action Items</div></a>
</div>
"@)

    $mbxCsv          = Join-Path $exchangeDir "Exchange_Mailboxes.csv"
    $forwardingCsv   = Join-Path $exchangeDir "Exchange_InboxForwardingRules.csv"
    $fullAccessCsv   = Join-Path $exchangeDir "Exchange_Permissions_FullAccess.csv"
    $sendAsCsv       = Join-Path $exchangeDir "Exchange_Permissions_SendAs.csv"
    $sendOnBehalfCsv = Join-Path $exchangeDir "Exchange_Permissions_SendOnBehalf.csv"

    # Build permission lookups keyed by MailboxUPN
    $permsByUpn = @{}
    if (Test-Path $fullAccessCsv) {
        foreach ($perm in (Import-Csv $fullAccessCsv)) {
            $key = $perm.MailboxUPN
            if (-not $permsByUpn.ContainsKey($key)) { $permsByUpn[$key] = @{ FullAccess = [System.Collections.Generic.List[string]]::new(); SendAs = [System.Collections.Generic.List[string]]::new(); SendOnBehalf = [System.Collections.Generic.List[string]]::new() } }
            $permsByUpn[$key].FullAccess.Add($perm.User)
        }
    }
    if (Test-Path $sendAsCsv) {
        foreach ($perm in (Import-Csv $sendAsCsv)) {
            $key = $perm.MailboxUPN
            if (-not $key) { continue }
            if (-not $permsByUpn.ContainsKey($key)) { $permsByUpn[$key] = @{ FullAccess = [System.Collections.Generic.List[string]]::new(); SendAs = [System.Collections.Generic.List[string]]::new(); SendOnBehalf = [System.Collections.Generic.List[string]]::new() } }
            $permsByUpn[$key].SendAs.Add($perm.Trustee)
        }
    }
    if (Test-Path $sendOnBehalfCsv) {
        foreach ($perm in (Import-Csv $sendOnBehalfCsv)) {
            $key = if ($perm.MailboxUPN) { $perm.MailboxUPN } else { $perm.Mailbox }
            if (-not $key) { continue }
            if (-not $permsByUpn.ContainsKey($key)) { $permsByUpn[$key] = @{ FullAccess = [System.Collections.Generic.List[string]]::new(); SendAs = [System.Collections.Generic.List[string]]::new(); SendOnBehalf = [System.Collections.Generic.List[string]]::new() } }
            if ($perm.Delegate) { $permsByUpn[$key].SendOnBehalf.Add($perm.Delegate) }
        }
    }

    if (Test-Path $mbxCsv) {
        $mailboxes = @(Import-Csv $mbxCsv)

        # Build UPN set for row highlighting — mailboxes >75% full with no archive
        $_nearFullUpns = @{}
        foreach ($_m in $mailboxes) {
            if ($_m.LimitMB -and $_m.LimitMB -ne '' -and [double]$_m.LimitMB -gt 0 -and
                $_m.ArchiveEnabled -eq 'False' -and
                ([double]$_m.TotalSizeMB / [double]$_m.LimitMB) -gt 0.75) {
                $_nearFullUpns[$_m.UserPrincipalName] = $true
            }
        }

        $exchangeSummary.Add("<h4 id='exchange-mailboxes'>Mailboxes</h4>")
        $exchangeSummary.Add((Get-ExpandHintHtml -Text 'Click a row to expand delegated permissions.'))

        $mbxRows = foreach ($mbx in ($mailboxes | Sort-Object @{Expression={ switch ($_.RecipientType) { 'UserMailbox' { 0 } 'SharedMailbox' { 1 } default { 2 } } }}, DisplayName)) {
            $upn    = $mbx.UserPrincipalName
            $usedMB = [double]$mbx.TotalSizeMB
            $limitMB = if ($mbx.LimitMB -and $mbx.LimitMB -ne '') { [double]$mbx.LimitMB } else { 0 }

            # Usage bar — only rendered when limit is known
            if ($limitMB -gt 0) {
                $usedPct  = [math]::Round(($usedMB / $limitMB) * 100, 1)
                $barWidth = [math]::Min($usedPct, 100)
                $barColor = if ($usedPct -gt 95) { "#e53935" } elseif ($usedPct -gt 80) { "#ff9800" } else { "#4caf50" }
                $limitGB  = [math]::Round($limitMB / 1024, 1)
                $usageCell = "<td><div style='display:flex;align-items:center;gap:6px'><div style='background:#e0e0e0;border-radius:4px;width:80px;height:12px;flex-shrink:0'><div style='background:$barColor;width:${barWidth}%;height:12px;border-radius:4px'></div></div><span style='font-size:0.8rem;color:#555'>$usedPct% of ${limitGB} GB</span></div></td>"
            }
            else {
                $usageCell = "<td><span style='font-size:0.85rem;color:#666'>$usedMB MB</span></td>"
            }

            $archiveSizeCell = if ($mbx.ArchiveSizeMB -and $mbx.ArchiveSizeMB -ne '') { "$($mbx.ArchiveSizeMB) MB" } else { "—" }

            # Near-full / no-archive highlighting
            $_isNearFull  = $_nearFullUpns.ContainsKey($upn)
            $_rowStyle    = if ($_isNearFull) { " style='border-left:3px solid #ff9800;background:#fff8f0'" } else { "" }
            $archiveCell  = if ($_isNearFull) {
                "<td><span style='color:#e65100;font-weight:600' title='Mailbox is over 75% full with no archive enabled'>No Archive</span></td>"
            } else {
                "<td>$($mbx.ArchiveEnabled)</td>"
            }

            # Main mailbox row — clickable
            $mainRow = "<tr class='user-row'$_rowStyle onclick='togglePerms(this)' title='Click to show/hide delegated permissions'><td>$($mbx.DisplayName)</td><td>$upn</td><td>$($mbx.RecipientType)</td>$usageCell$archiveCell<td>$archiveSizeCell</td></tr>"

            # Permissions detail row (hidden by default)
            $perms = if ($permsByUpn.ContainsKey($upn)) { $permsByUpn[$upn] } else { $null }

            if ($perms -and ($perms.FullAccess.Count -gt 0 -or $perms.SendAs.Count -gt 0 -or $perms.SendOnBehalf.Count -gt 0)) {
                $faRows  = ($perms.FullAccess   | ForEach-Object { "<tr><td>Full Access</td><td>$_</td></tr>" }) -join ""
                $saRows  = ($perms.SendAs       | ForEach-Object { "<tr><td>Send As</td><td>$_</td></tr>" })    -join ""
                $sobRows = ($perms.SendOnBehalf | ForEach-Object { "<tr><td>Send on Behalf</td><td>$_</td></tr>" }) -join ""
                $detailRow = "<tr class='signin-detail' style='display:none'><td colspan='6'><table class='inner-table'><thead><tr><th>Permission Type</th><th>Trustee / Delegate</th></tr></thead><tbody>$faRows$saRows$sobRows</tbody></table></td></tr>"
            }
            else {
                $detailRow = "<tr class='signin-detail' style='display:none'><td colspan='6'><em>No delegated permissions on this mailbox</em></td></tr>"
            }

            $mainRow
            $detailRow
        }

        $exchangeSummary.Add(@"
<table>
  <thead><tr>
    <th>Display Name</th><th>UPN</th><th>Type</th>
    <th>Usage</th>
    <th>Archive</th><th>Archive Size</th>
  </tr></thead>
  <tbody>$($mbxRows -join "`n")</tbody>
</table>
"@)

        # --- Shared Mailbox Sign-In Status ---
        $sharedSignInCsv = Join-Path $exchangeDir "Exchange_SharedMailboxSignIn.csv"
        if (Test-Path $sharedSignInCsv) {
            $sharedMbx     = @(Import-Csv $sharedSignInCsv)
            $signInEnabled = @($sharedMbx | Where-Object { $_.AccountDisabled -eq "False" })
            if ($signInEnabled.Count -eq 0) {
                $exchangeSummary.Add("<p class='ok'>Shared mailbox interactive sign-in: all $($sharedMbx.Count) disabled.</p>")
            }
            else {
                $exchangeSummary.Add("<p class='warn'>$($signInEnabled.Count) of $($sharedMbx.Count) shared mailbox(es) have interactive sign-in <b>enabled</b> — should be disabled to prevent direct login.</p>")
                $siRows = foreach ($m in ($signInEnabled | Sort-Object DisplayName)) {
                    "<tr class='warn'><td>$($m.DisplayName)</td><td>$($m.UserPrincipalName)</td></tr>"
                }
                $exchangeSummary.Add(@"
<table>
  <thead><tr><th>Display Name</th><th>UPN</th></tr></thead>
  <tbody>$($siRows -join "`n")</tbody>
</table>
"@)
            }
        }
    }

    # --- Inbox Forwarding Rules ---
    if (Test-Path $forwardingCsv) {
        $forwarding = @(Import-Csv $forwardingCsv)
        if ($forwarding.Count -gt 0) {
            $fwdRows = foreach ($r in $forwarding) {
                "<tr><td>$($r.Mailbox)</td><td>$($r.RuleName)</td><td>$($r.ForwardTo)</td><td>$($r.RedirectTo)</td></tr>"
            }
            $exchangeSummary.Add("<h4 id='exchange-forwarding'>Inbox Forwarding Rules</h4>")
            $exchangeSummary.Add("<p class='warn'>$($forwarding.Count) inbox rule(s) with external forwarding detected</p>")
            $exchangeSummary.Add(@"
<table>
  <thead><tr><th>Mailbox</th><th>Rule Name</th><th>Forward To</th><th>Redirect To</th></tr></thead>
  <tbody>$($fwdRows -join "`n")</tbody>
</table>
"@)
        }
        else {
            $exchangeSummary.Add("<p class='ok'>No external forwarding inbox rules detected</p>")
        }
    }

    # --- Broken Inbox Rules ---
    $brokenCsv = Join-Path $exchangeDir "Exchange_BrokenInboxRules.csv"
    if (Test-Path $brokenCsv) {
        $brokenRules = @(Import-Csv $brokenCsv)
        if ($brokenRules.Count -gt 0) {
            $brokenRows = foreach ($r in $brokenRules) {
                "<tr><td>$($r.Mailbox)</td><td>$($r.RuleName)</td><td class='warn'>Broken — not processing mail</td></tr>"
            }
            $exchangeSummary.Add("<h4>Broken Inbox Rules</h4>")
            $exchangeSummary.Add("<p class='warn'>$($brokenRules.Count) inbox rule(s) are in a broken/non-functional state. Edit or re-create them in Outlook.</p>")
            $exchangeSummary.Add(@"
<table>
  <thead><tr><th>Mailbox</th><th>Rule Name</th><th>Status</th></tr></thead>
  <tbody>$($brokenRows -join "`n")</tbody>
</table>
"@)
        }
    }

    # --- Remote Domain Auto-Forwarding ---
    # The wildcard (*) entry is present in every tenant by default — only flag named domains with auto-forward enabled
    $remoteDomainCsv = Join-Path $exchangeDir "Exchange_RemoteDomainForwarding.csv"
    if (Test-Path $remoteDomainCsv) {
        $remoteDomains   = @(Import-Csv $remoteDomainCsv)
        $namedFwdDomains = @($remoteDomains | Where-Object { $_.DomainName -ne '*' -and $_.AutoForwardEnabled -eq 'True' })
        $fwdClass        = if ($namedFwdDomains.Count -gt 0) { "warn" } else { "ok" }
        $fwdMsg          = if ($namedFwdDomains.Count -gt 0) {
            "Auto-forwarding explicitly enabled for $($namedFwdDomains.Count) named external domain(s) — confirm these are intentional."
        } else { "No named external domains have auto-forwarding enabled" }
        $exchangeSummary.Add("<p class='$fwdClass'>Remote Domains: $fwdMsg</p>")
        $rdRows = foreach ($rd in $remoteDomains) {
            $rdClass     = if ($rd.DomainName -ne '*' -and $rd.AutoForwardEnabled -eq 'True') { " class='warn'" } else { "" }
            $domainLabel = if ($rd.DomainName -eq '*') { "* <span style='color:#888;font-weight:normal'>(default — all external domains)</span>" } else { $rd.DomainName }
            "<tr$rdClass><td>$domainLabel</td><td>$($rd.AutoForwardEnabled)</td></tr>"
        }
        $exchangeSummary.Add(@"
<table>
  <thead><tr><th>Remote Domain</th><th>Auto-Forward Enabled</th></tr></thead>
  <tbody>$($rdRows -join "`n")</tbody>
</table>
"@)
    }

    # --- Audit Configuration ---
    $auditConfigCsv = Join-Path $exchangeDir "Exchange_AuditConfig.csv"
    if (Test-Path $auditConfigCsv) {
        $auditCfg      = Import-Csv $auditConfigCsv | Select-Object -First 1
        $retentionDays = try { [TimeSpan]::Parse($auditCfg.AuditLogAgeLimit).Days } catch { $auditCfg.AuditLogAgeLimit }
        $ualClass      = if ($auditCfg.UnifiedAuditLogIngestionEnabled -eq 'True') { "ok" } else { "critical" }
        $aalClass      = if ($auditCfg.AdminAuditLogEnabled -eq 'True') { "ok" } else { "warn" }
        $retClass      = if ([int]$retentionDays -lt 90) { "warn" } else { "ok" }
        $exchangeSummary.Add("<h4>Audit Configuration</h4>")
        $exchangeSummary.Add(@"
<table>
  <thead><tr><th>Setting</th><th>Value</th><th>What this means</th></tr></thead>
  <tbody>
    <tr><td>Unified Audit Log</td><td class='$ualClass'>$($auditCfg.UnifiedAuditLogIngestionEnabled)</td><td style='color:#666;font-size:0.85rem'>Captures all user and admin activity across M365 services. Required for security investigations, compliance, and incident response. Should always be enabled.</td></tr>
    <tr><td>Admin Audit Log</td><td class='$aalClass'>$($auditCfg.AdminAuditLogEnabled)</td><td style='color:#666;font-size:0.85rem'>Records Exchange admin actions (PowerShell commands, admin centre changes). Useful for tracking configuration changes.</td></tr>
    <tr><td>Log Retention</td><td class='$retClass'>$retentionDays days</td><td style='color:#666;font-size:0.85rem'>How long audit records are kept. Standard tenants retain 90 days. Microsoft 365 E3/E5 can extend to 1 year or more via compliance add-ons. Below 90 days limits investigation capability.</td></tr>
  </tbody>
</table>
"@)
    }

    # --- Mailbox Audit Status ---
    $mbxAuditCsv = Join-Path $exchangeDir "Exchange_MailboxAuditStatus.csv"
    if (Test-Path $mbxAuditCsv) {
        # Exclude system mailboxes (Discovery Search is an internal EXO system account, not a user mailbox)
        $mbxAudit      = @(Import-Csv $mbxAuditCsv | Where-Object { $_.UserPrincipalName -notlike 'DiscoverySearchMailbox*' })
        $auditDisabled = @($mbxAudit | Where-Object { $_.AuditEnabled -ne 'True' })
        $exchangeSummary.Add("<p style='font-size:0.85rem;color:#666;margin:0.5rem 0 0.4rem'>Mailbox auditing records actions taken by owners, delegates, and admins — including logins, reads, moves, and deletions. Mailboxes with auditing disabled will not appear in Microsoft Purview compliance searches or security investigations.</p>")
        if ($auditDisabled.Count -gt 0) {
            $auditRows = foreach ($m in $auditDisabled) {
                "<tr><td>$($m.DisplayName)</td><td>$($m.UserPrincipalName)</td></tr>"
            }
            $exchangeSummary.Add("<p class='warn'>$($auditDisabled.Count) of $($mbxAudit.Count) mailbox(es) have auditing disabled</p>")
            $exchangeSummary.Add(@"
<table>
  <thead><tr><th>Display Name</th><th>UPN</th></tr></thead>
  <tbody>$($auditRows -join "`n")</tbody>
</table>
"@)
        }
        else {
            $exchangeSummary.Add("<p class='ok'>Mailbox auditing enabled on all $($mbxAudit.Count) mailboxes</p>")
        }
    }

    # --- DKIM Status ---
    $dkimExCsv = Join-Path $exchangeDir "Exchange_DKIM_Status.csv"
    if (Test-Path $dkimExCsv) {
        # Exclude Microsoft-managed onmicrosoft.com domains — DKIM on these is controlled by Microsoft, not the customer
        $dkimEx      = @(Import-Csv $dkimExCsv | Where-Object { $_.Domain -notlike "*.onmicrosoft.com" })
        $dkimOff     = @($dkimEx | Where-Object { $_.DKIMEnabled -ne 'True' })
        $dkimClass   = if ($dkimOff.Count -eq 0) { "ok" } else { "warn" }
        $dkimRows    = foreach ($d in $dkimEx) {
            $cls = if ($d.DKIMEnabled -ne 'True') { " class='warn'" } else { "" }
            "<tr$cls><td>$($d.Domain)</td><td>$($d.DKIMEnabled)</td><td style='font-size:0.8rem;word-break:break-all'>$($d.Selector1CNAME)</td></tr>"
        }
        $exchangeSummary.Add("<h4>DKIM Signing</h4>")
        $exchangeSummary.Add("<p class='$dkimClass'>DKIM enabled on $($dkimEx.Count - $dkimOff.Count) / $($dkimEx.Count) custom domains</p>")
        $exchangeSummary.Add(@"
<table>
  <thead><tr><th>Domain</th><th>DKIM Enabled</th><th>Selector 1 CNAME</th></tr></thead>
  <tbody>$($dkimRows -join "`n")</tbody>
</table>
"@)
    }

    # --- Anti-Phish Policies ---
    $antiPhishCsv = Join-Path $exchangeDir "Exchange_AntiPhishPolicies.csv"
    if (Test-Path $antiPhishCsv) {
        $antiPhish = @(Import-Csv $antiPhishCsv)
        $exchangeSummary.Add("<h4 id='exchange-policies'>Anti-Phish Policies</h4>")
        $exchangeSummary.Add((Get-ExpandHintHtml -Text 'Click a row to expand setting descriptions.'))
        $apRows = foreach ($p in $antiPhish) {
            $spoofClass  = if ($p.EnableSpoofIntelligence -ne 'True') { " class='warn'" } else { "" }
            $mbxIntClass = if ($p.EnableMailboxIntelligence -ne 'True') { " class='warn'" } else { "" }
            $mainRow = "<tr class='user-row' onclick='togglePerms(this)' title='Click to expand'><td>$($p.Name)</td><td$spoofClass>$($p.EnableSpoofIntelligence)</td><td$mbxIntClass>$($p.EnableMailboxIntelligence)</td><td>$($p.EnableTargetedUserProtection)</td><td>$($p.EnableATPForSpoof)</td></tr>"
            $detailRow = "<tr class='signin-detail' style='display:none'><td colspan='5'><table class='inner-table'>
  <thead><tr><th>Setting</th><th>Description</th></tr></thead>
  <tbody>
    <tr><td>Spoof Intelligence</td><td>Detects when external senders forge your domain in the From address. Disabling this allows classic spoofing attacks to bypass filters entirely.</td></tr>
    <tr><td>Mailbox Intelligence</td><td>Builds a contact graph for each mailbox and flags messages from senders impersonating frequent contacts (e.g. a fake CFO email to finance staff).</td></tr>
    <tr><td>Targeted User Protection</td><td>Adds impersonation protection for specific high-value accounts (executives, IT admins) defined in the policy. Requires Defender for Office 365 Plan 1+.</td></tr>
    <tr><td>ATP for Spoof</td><td>Applies Advanced Threat Protection verdict and action to messages that fail spoof intelligence checks, rather than the standard spam action.</td></tr>
  </tbody>
</table></td></tr>"
            $mainRow
            $detailRow
        }
        $exchangeSummary.Add(@"
<table>
  <thead><tr><th>Policy</th><th>Spoof Intelligence</th><th>Mailbox Intelligence</th><th>Targeted User Protection</th><th>ATP for Spoof</th></tr></thead>
  <tbody>$($apRows -join "`n")</tbody>
</table>
"@)
    }

    # --- Spam Policies ---
    $spamCsv = Join-Path $exchangeDir "Exchange_SpamPolicies.csv"
    if (Test-Path $spamCsv) {
        $spamPolicies = @(Import-Csv $spamCsv)
        $exchangeSummary.Add("<h4>Spam Filter Policies</h4>")
        $exchangeSummary.Add((Get-ExpandHintHtml -Text 'Click a row to expand action descriptions.'))
        $spamRows = foreach ($p in $spamPolicies) {
            $spamClass = if ($p.SpamAction -eq 'NoAction') { " class='warn'" } else { "" }
            $hcsClass  = if ($p.HighConfidenceSpamAction -eq 'NoAction' -or $p.HighConfidenceSpamAction -eq 'MoveToJmf') { " class='warn'" } else { "" }
            $mainRow = "<tr class='user-row' onclick='togglePerms(this)' title='Click to expand'><td>$($p.Name)</td><td$spamClass>$($p.SpamAction)</td><td$hcsClass>$($p.HighConfidenceSpamAction)</td><td>$($p.BulkSpamAction)</td></tr>"
            $detailRow = "<tr class='signin-detail' style='display:none'><td colspan='4'><table class='inner-table'>
  <thead><tr><th>Action</th><th>Description</th></tr></thead>
  <tbody>
    <tr><td>MoveToJmf</td><td>Move message to the recipient's Junk Email folder. User can still access it. Lowest friction but relies on users reporting missed spam.</td></tr>
    <tr><td>Quarantine</td><td>Hold message in Microsoft quarantine. Admin or user (if policy allows) must release it. Recommended for high-confidence spam and phishing.</td></tr>
    <tr><td>NoAction</td><td>Deliver the message normally. Not recommended — messages that match this filter will reach the inbox unmodified.</td></tr>
    <tr><td>Delete</td><td>Silently delete the message. Sender receives no bounce. Use with caution as legitimate mail can be lost without trace.</td></tr>
  </tbody>
</table></td></tr>"
            $mainRow
            $detailRow
        }
        $exchangeSummary.Add(@"
<table>
  <thead><tr><th>Policy</th><th>Spam Action</th><th>High Confidence Spam</th><th>Bulk Spam</th></tr></thead>
  <tbody>$($spamRows -join "`n")</tbody>
</table>
"@)
    }

    # --- Malware Policies ---
    $malwareCsv = Join-Path $exchangeDir "Exchange_MalwarePolicies.csv"
    if (Test-Path $malwareCsv) {
        $malware = @(Import-Csv $malwareCsv)
        $exchangeSummary.Add("<h4>Malware Filter Policies</h4>")
        $exchangeSummary.Add((Get-ExpandHintHtml -Text 'Click a row to expand setting descriptions.'))
        $mwRows = foreach ($p in $malware) {
            $notifyClass = if ($p.EnableExternalSenderAdminNotification -eq 'True') { "" } else { " class='warn'" }
            $mainRow = "<tr class='user-row' onclick='togglePerms(this)' title='Click to expand'><td>$($p.Name)</td><td>$($p.Action)</td><td$notifyClass>$($p.EnableExternalSenderAdminNotification)</td></tr>"
            $detailRow = "<tr class='signin-detail' style='display:none'><td colspan='3'><table class='inner-table'>
  <thead><tr><th>Setting</th><th>Description</th></tr></thead>
  <tbody>
    <tr><td>Action</td><td>What happens when malware is detected. In EXO with Defender, the default action is to quarantine the entire message. The Action field in this policy is largely superseded by Defender for Office 365 Safe Attachments.</td></tr>
    <tr><td>External Sender Admin Notification</td><td>Whether to notify an admin when an external sender's message is quarantined for malware. Useful for tracking inbound malware volume but can generate noise in high-volume environments.</td></tr>
  </tbody>
</table></td></tr>"
            $mainRow
            $detailRow
        }
        $exchangeSummary.Add(@"
<table>
  <thead><tr><th>Policy</th><th>Action</th><th>External Sender Admin Notification</th></tr></thead>
  <tbody>$($mwRows -join "`n")</tbody>
</table>
"@)
    }

    # --- Transport Rules ---
    $transportCsv = Join-Path $exchangeDir "Exchange_TransportRules.csv"
    if (Test-Path $transportCsv) {
        $transportRules = @(Import-Csv $transportCsv)
        if ($transportRules.Count -gt 0) {
            $disabledRules = @($transportRules | Where-Object { $_.State -ne 'Enabled' })
            $trClass       = if ($disabledRules.Count -gt 0) { "warn" } else { "ok" }
            $exchangeSummary.Add("<h4>Transport Rules</h4>")
            $exchangeSummary.Add("<p class='$trClass'>$($transportRules.Count) transport rule(s) — $($disabledRules.Count) disabled.</p>")
            $exchangeSummary.Add((Get-ExpandHintHtml -Text 'Click a row to expand conditions and actions.'))
            $trRows = foreach ($r in ($transportRules | Sort-Object { [int]$_.Priority })) {
                $stateClass = if ($r.State -ne 'Enabled') { " class='warn'" } else { "" }
                $mainRow    = "<tr class='user-row' onclick='togglePerms(this)' title='Click to expand'><td>$($r.Priority)</td><td>$($r.Name)</td><td$stateClass>$($r.State)</td><td>$($r.Mode)</td></tr>"

                # Build condition/action detail rows — only show populated fields
                $details = [System.Collections.Generic.List[string]]::new()
                $modeDesc = switch ($r.Mode) {
                    'Enforce' { 'Rule is active and actions are applied to matching messages.' }
                    'Audit'   { 'Rule is in audit mode — conditions are evaluated and results logged, but no action is taken.' }
                    'AuditAndNotify' { 'Audit mode with policy tip notification shown to the sender.' }
                    default   { $r.Mode }
                }
                $details.Add("<tr><td>Mode</td><td>$($r.Mode) <span style='color:#666;font-size:0.85rem'>— $modeDesc</span></td></tr>")
                if ($r.FromAddressContainsWords) { $details.Add("<tr><td>From Contains Words</td><td>$($r.FromAddressContainsWords)</td></tr>") }
                if ($r.SentTo)                   { $details.Add("<tr><td>Sent To</td><td>$($r.SentTo)</td></tr>") }
                if ($r.RedirectMessageTo)         { $details.Add("<tr><td>Redirect To</td><td>$($r.RedirectMessageTo)</td></tr>") }
                if ($r.BlindCopyTo)              { $details.Add("<tr><td>BCC To</td><td>$($r.BlindCopyTo)</td></tr>") }
                if ($r.ApplyHtmlDisclaimerText)  {
                    $disclaimerLoc = if ($r.ApplyHtmlDisclaimerLocation) { " ($($r.ApplyHtmlDisclaimerLocation))" } else { "" }
                    $details.Add("<tr><td>Disclaimer$disclaimerLoc</td><td style='max-width:400px;overflow:hidden;white-space:nowrap;text-overflow:ellipsis;font-size:0.82rem'>$($r.ApplyHtmlDisclaimerText)</td></tr>")
                }

                $detailInner = "<table class='inner-table'><thead><tr><th style='width:160px'>Condition / Action</th><th>Value</th></tr></thead><tbody>$($details -join '')</tbody></table>"
                $detailRow = "<tr class='signin-detail' style='display:none'><td colspan='4'>$detailInner</td></tr>"
                $mainRow
                $detailRow
            }
            $exchangeSummary.Add(@"
<table>
  <thead><tr><th>Priority</th><th>Rule Name</th><th>State</th><th>Mode</th></tr></thead>
  <tbody>$($trRows -join "`n")</tbody>
</table>
"@)
        }
        else {
            $exchangeSummary.Add("<p>No transport rules configured</p>")
        }
    }

    # --- Distribution Lists ---
    $dlCsv = Join-Path $exchangeDir "Exchange_DistributionLists.csv"
    if (Test-Path $dlCsv) {
        $dls      = @(Import-Csv $dlCsv)
        $emptyDls = @($dls | Where-Object { [int]$_.MemberCount -eq 0 })
        $exchangeSummary.Add("<h4>Distribution Lists ($($dls.Count) total)</h4>")
        if ($emptyDls.Count -gt 0) {
            $exchangeSummary.Add("<p class='warn'>$($emptyDls.Count) distribution list(s) have no members</p>")
        }
        $exchangeSummary.Add((Get-ExpandHintHtml -Text 'Click a row to expand the member list.'))
        $dlRows = foreach ($dl in ($dls | Sort-Object DisplayName)) {
            $warnClass  = if ([int]$dl.MemberCount -eq 0) { " warn" } else { "" }
            $mainRow    = "<tr class='user-row$warnClass' onclick='togglePerms(this)' title='Click to expand members'><td>$($dl.DisplayName)</td><td>$($dl.EmailAddress)</td><td>$($dl.GroupType)</td><td>$($dl.MemberCount)</td></tr>"

            if ($dl.Members -and [int]$dl.MemberCount -gt 0) {
                $memberRows = ($dl.Members -split '; ' | Where-Object { $_ } | Sort-Object | ForEach-Object { "<tr><td>$_</td></tr>" }) -join ""
                $detailRow  = "<tr class='signin-detail' style='display:none'><td colspan='4'><table class='inner-table'><thead><tr><th>Member</th></tr></thead><tbody>$memberRows</tbody></table></td></tr>"
            }
            else {
                $detailRow = "<tr class='signin-detail' style='display:none'><td colspan='4'><em>No members</em></td></tr>"
            }
            $mainRow
            $detailRow
        }
        $exchangeSummary.Add(@"
<table>
  <thead><tr><th>Name</th><th>Email</th><th>Type</th><th>Members</th></tr></thead>
  <tbody>$($dlRows -join "`n")</tbody>
</table>
"@)
    }

    # --- Resource Mailboxes ---
    $resourceCsv = Join-Path $exchangeDir "Exchange_ResourceMailboxes.csv"
    if (Test-Path $resourceCsv) {
        $resources = @(Import-Csv $resourceCsv)
        if ($resources.Count -gt 0) {
            $exchangeSummary.Add("<h4>Resource Mailboxes ($($resources.Count) total)</h4>")
            $resRows = foreach ($r in ($resources | Sort-Object DisplayName)) {
                $conflictClass = if ($r.AllowConflicts -eq 'True') { " class='warn'" } else { "" }
                "<tr><td>$($r.DisplayName)</td><td>$($r.ResourceType)</td><td>$($r.Email)</td><td>$($r.BookingWindowDays)</td><td$conflictClass>$($r.AllowConflicts)</td><td>$($r.BookingDelegates)</td></tr>"
            }
            $exchangeSummary.Add(@"
<table>
  <thead><tr><th>Name</th><th>Type</th><th>Email</th><th>Booking Window (days)</th><th>Allow Conflicts</th><th>Delegates</th></tr></thead>
  <tbody>$($resRows -join "`n")</tbody>
</table>
"@)
        }
    }

    # --- Outbound Spam Auto-Forward ---
    $outboundFwdCsv = Join-Path $exchangeDir "Exchange_OutboundSpamAutoForward.csv"
    if (Test-Path $outboundFwdCsv) {
        $outboundFwd  = @(Import-Csv $outboundFwdCsv)
        $fwdOnCount   = ($outboundFwd | Where-Object { $_.AutoForwardingMode -eq "On" }).Count
        $fwdAutoCount = ($outboundFwd | Where-Object { $_.AutoForwardingMode -eq "Automatic" }).Count
        $fwdOffCount  = ($outboundFwd | Where-Object { $_.AutoForwardingMode -eq "Off" }).Count
        $fwdClass     = if ($fwdOnCount -gt 0) { "critical" } elseif ($fwdOffCount -eq $outboundFwd.Count) { "ok" } else { "warn" }
        $fwdSummary   = "On: $fwdOnCount (unrestricted), Automatic: $fwdAutoCount (tenant-controlled), Off: $fwdOffCount (blocked)"
        $exchangeSummary.Add("<h4>Outbound Spam Auto-Forward Policy</h4>")
        $exchangeSummary.Add("<p class='$fwdClass'>$fwdSummary</p>")
        $exchangeSummary.Add("<p style='font-size:0.85rem;color:#555'>Recommended: <b>Off</b> (blocked) or <b>Automatic</b> (follows remote domain settings). <b>On</b> allows unrestricted forwarding.</p>")
        $fwdRows = foreach ($p in $outboundFwd) {
            $cls = switch ($p.AutoForwardingMode) { "On" { " class='critical'" } "Off" { " class='ok'" } default { "" } }
            "<tr$cls><td>$($p.Name)</td><td>$($p.AutoForwardingMode)</td></tr>"
        }
        $exchangeSummary.Add(@"
<table>
  <thead><tr><th>Policy</th><th>Auto-Forward Mode</th></tr></thead>
  <tbody>$($fwdRows -join "`n")</tbody>
</table>
"@)
    }

    # --- Safe Attachments ---
    $safeAttCsv = Join-Path $exchangeDir "Exchange_SafeAttachments.csv"
    $exchangeSummary.Add("<h4>Defender for Office 365 — Safe Attachments</h4>")
    if (Test-Path $safeAttCsv) {
        $safeAtt    = @(Import-Csv $safeAttCsv)
        $attEnabled = ($safeAtt | Where-Object { $_.Enable -eq "True" }).Count
        $attClass   = if ($attEnabled -gt 0) { "ok" } else { "warn" }
        $exchangeSummary.Add("<p class='$attClass'>$attEnabled of $($safeAtt.Count) Safe Attachment policy/policies enabled</p>")
        $exchangeSummary.Add((Get-ExpandHintHtml -Text 'Click a row to expand setting descriptions.'))
        $attRows = foreach ($p in $safeAtt) {
            $enableClass = if ($p.Enable -eq "True") { "ok" } else { "warn" }
            $mainRow = "<tr class='user-row' onclick='togglePerms(this)' title='Click to expand'><td>$($p.Name)</td><td class='$enableClass'>$($p.Enable)</td><td>$($p.Action)</td><td>$($p.ActionOnError)</td></tr>"
            $detailRow = "<tr class='signin-detail' style='display:none'><td colspan='4'><table class='inner-table'>
  <thead><tr><th style='width:160px'>Setting</th><th>Description</th></tr></thead>
  <tbody>
    <tr><td>Enabled</td><td>Whether this Safe Attachments policy is active. Disabled policies do not scan attachments, even if a rule applies the policy to users.</td></tr>
    <tr><td>Action: $($p.Action)</td><td>$(switch ($p.Action) { 'Block' { 'Quarantines the entire message when a malicious attachment is detected. The recipient never receives the message.' } 'Replace' { 'Strips the attachment from the email and delivers the message body with a notice. The user receives the email but not the file.' } 'DynamicDelivery' { 'Delivers the email immediately with a placeholder attachment while Safe Attachments scans in the background. The real attachment replaces the placeholder if clean. Recommended for minimal delay.' } 'Allow' { 'Delivers the message without scanning. Not recommended.' } default { $p.Action } })</td></tr>
    <tr><td>Action on Error</td><td>$(if ($p.ActionOnError -eq 'True') { 'Block — if the scanning service errors or times out, the message is held. Recommended to prevent bypass via scanner failure.' } else { 'Pass — if scanning errors, the message is delivered unscanned. Increases risk but reduces false-positive delays.' })</td></tr>
  </tbody>
</table></td></tr>"
            $mainRow
            $detailRow
        }
        $exchangeSummary.Add(@"
<table>
  <thead><tr><th>Policy</th><th>Enabled</th><th>Action</th><th>Action on Error</th></tr></thead>
  <tbody>$($attRows -join "`n")</tbody>
</table>
"@)
    }
    else {
        $exchangeSummary.Add("<p style='color:#888'>Safe Attachments data not collected — Defender for Office 365 P1 may not be licensed on this tenant.</p>")
    }

    # --- Safe Links ---
    $safeLinkCsv = Join-Path $exchangeDir "Exchange_SafeLinks.csv"
    $exchangeSummary.Add("<h4>Defender for Office 365 — Safe Links</h4>")
    if (Test-Path $safeLinkCsv) {
        $safeLink     = @(Import-Csv $safeLinkCsv)
        $linkEnabled  = ($safeLink | Where-Object { $_.EnableSafeLinksForEmail -eq "True" }).Count
        $linkClass    = if ($linkEnabled -gt 0) { "ok" } else { "warn" }
        $exchangeSummary.Add("<p class='$linkClass'>$linkEnabled of $($safeLink.Count) Safe Links policy/policies enabled for email</p>")
        $exchangeSummary.Add((Get-ExpandHintHtml -Text 'Click a row to expand setting descriptions.'))
        $linkRows = foreach ($p in $safeLink) {
            $enableClass = if ($p.EnableSafeLinksForEmail -eq "True") { "ok" } else { "warn" }
            $mainRow = "<tr class='user-row' onclick='togglePerms(this)' title='Click to expand'><td>$($p.Name)</td><td class='$enableClass'>$($p.EnableSafeLinksForEmail)</td><td>$($p.EnableSafeLinksForTeams)</td><td>$($p.DisableUrlRewrite)</td><td>$($p.TrackClicks)</td></tr>"
            $detailRow = "<tr class='signin-detail' style='display:none'><td colspan='5'><table class='inner-table'>
  <thead><tr><th style='width:200px'>Setting</th><th>Description</th></tr></thead>
  <tbody>
    <tr><td>Safe Links for Email</td><td>URLs in incoming emails are rewritten through Microsoft's detonation service and checked at time-of-click. If the link destination is malicious, the user sees a warning page and access is blocked.</td></tr>
    <tr><td>Safe Links for Teams</td><td>Same time-of-click URL protection applied to links shared in Teams messages and channel posts.</td></tr>
    <tr><td>URL Rewrite Disabled: $($p.DisableUrlRewrite)</td><td>$(if ($p.DisableUrlRewrite -eq 'True') { '<span class=''warn''>URLs are NOT rewritten in email bodies. Time-of-click protection still occurs via client-side hooks in Outlook, but protection is weaker for forwarded emails or non-Outlook clients.</span>' } else { 'URLs are rewritten through Microsoft — time-of-click protection works across all email clients, including forwarded messages.' })</td></tr>
    <tr><td>Track Clicks: $($p.TrackClicks)</td><td>Records which URLs users click and whether they were blocked or allowed. Data appears in Threat Explorer and URL trace reports in Microsoft Defender.</td></tr>
  </tbody>
</table></td></tr>"
            $mainRow
            $detailRow
        }
        $exchangeSummary.Add(@"
<table>
  <thead><tr><th>Policy</th><th>Email</th><th>Teams</th><th>URL Rewrite Disabled</th><th>Track Clicks</th></tr></thead>
  <tbody>$($linkRows -join "`n")</tbody>
</table>
"@)
    }
    else {
        $exchangeSummary.Add("<p style='color:#888'>Safe Links data not collected — Defender for Office 365 P1 may not be licensed on this tenant.</p>")
    }

    # --- Org Config (SMTP AUTH, Customer Lockbox) ---
    $orgCfgCsv = Join-Path $exchangeDir "Exchange_OrgConfig.csv"
    if (Test-Path $orgCfgCsv) {
        $_exoOrgCfg = Import-Csv $orgCfgCsv | Select-Object -First 1
        $exchangeSummary.Add("<hr class='section-divider'>")
        $exchangeSummary.Add("<h4 id='exchange-org'>Organisation Configuration</h4>")
        $exchangeSummary.Add("<p style='font-size:0.82rem;color:#64748b'>References: <a href='https://www.cisecurity.org/benchmark/microsoft_365' target='_blank'>CIS M365 Benchmark</a> (6.1–6.5) · <a href='https://learn.microsoft.com/en-us/exchange/clients-and-mobile-in-exchange-online/client-access/exchange-admin-center' target='_blank'>Microsoft Exchange Security Baseline</a>.</p>")
        $exchangeSummary.Add("<table class='summary-table'><thead><tr><th>Setting</th><th>Value</th></tr></thead><tbody>")
        $_smtpClass = if ($_exoOrgCfg.SmtpClientAuthDisabled -eq 'False') { "style='color:#c62828;font-weight:bold'" } else { '' }
        $exchangeSummary.Add("<tr><td>SMTP AUTH disabled org-wide</td><td $_smtpClass>$(ConvertTo-HtmlText $_exoOrgCfg.SmtpClientAuthDisabled)</td></tr>")
        $exchangeSummary.Add("<tr><td>Modern Auth (OAuth) enabled</td><td>$(ConvertTo-HtmlText $_exoOrgCfg.ModernAuthEnabled)</td></tr>")
        $exchangeSummary.Add("<tr><td>Customer Lockbox enabled</td><td>$(ConvertTo-HtmlText $_exoOrgCfg.CustomerLockboxEnabled)</td></tr>")
        $exchangeSummary.Add("<tr><td>Mailbox audit disabled</td><td>$(ConvertTo-HtmlText $_exoOrgCfg.AuditDisabled)</td></tr>")
        $exchangeSummary.Add("</tbody></table>")
    }

    # --- External Sender Tagging + Connection Filter ---
    $extSenderCsv = Join-Path $exchangeDir "Exchange_ExternalSenderTagging.csv"
    $connFilterCsv = Join-Path $exchangeDir "Exchange_ConnectionFilter.csv"
    if ((Test-Path $extSenderCsv) -or (Test-Path $connFilterCsv)) {
        $exchangeSummary.Add("<hr class='section-divider'>")
        $exchangeSummary.Add("<h4>Transport Security</h4>")
        $exchangeSummary.Add("<table class='summary-table'><thead><tr><th>Setting</th><th>Value</th></tr></thead><tbody>")
        if (Test-Path $extSenderCsv) {
            $_exoEst = Import-Csv $extSenderCsv | Select-Object -First 1
            $_estClass = if ($_exoEst.Enabled -eq 'False') { "style='color:#e65100'" } else { '' }
            $exchangeSummary.Add("<tr><td>External sender tagging enabled</td><td $_estClass>$(ConvertTo-HtmlText $_exoEst.Enabled)</td></tr>")
        }
        if (Test-Path $connFilterCsv) {
            $_exoConn = @(Import-Csv $connFilterCsv)
            foreach ($_cf in $_exoConn) {
                $_ipAllowClass = if ($_cf.IPAllowList) { "style='color:#e65100'" } else { '' }
                $exchangeSummary.Add("<tr><td>Connection filter '$($_cf.PolicyName)' — IP allow list</td><td $_ipAllowClass>$(if ($_cf.IPAllowList) { ConvertTo-HtmlText $_cf.IPAllowList } else { 'Empty' })</td></tr>")
                $exchangeSummary.Add("<tr><td>Connection filter '$($_cf.PolicyName)' — safe list enabled</td><td>$(ConvertTo-HtmlText $_cf.EnableSafeList)</td></tr>")
            }
        }
        $exchangeSummary.Add("</tbody></table>")
    }

    # --- OWA Policy ---
    $owaPolicyCsv = Join-Path $exchangeDir "Exchange_OwaPolicy.csv"
    if (Test-Path $owaPolicyCsv) {
        $_exoOwa = @(Import-Csv $owaPolicyCsv)
        $exchangeSummary.Add("<hr class='section-divider'>")
        $exchangeSummary.Add("<h4>Outlook on the Web (OWA) Policies ($($_exoOwa.Count))</h4>")
        $exchangeSummary.Add("<table class='summary-table'><thead><tr><th>Policy</th><th>External Storage</th><th>Third-party Storage</th><th>Personal Calendars</th></tr></thead><tbody>")
        foreach ($_op in $_exoOwa) {
            $_storageClass = if ($_op.AdditionalStorageProvidersAvailable -eq 'True') { "style='color:#e65100'" } else { '' }
            $exchangeSummary.Add("<tr><td>$(ConvertTo-HtmlText $_op.PolicyName)</td><td $_storageClass>$(ConvertTo-HtmlText $_op.AdditionalStorageProvidersAvailable)</td><td>$(ConvertTo-HtmlText $_op.ThirdPartyAttachmentsEnabled)</td><td>$(ConvertTo-HtmlText $_op.PersonalAccountCalendarsEnabled)</td></tr>")
        }
        $exchangeSummary.Add("</tbody></table>")
    }

    $html.Add((Add-Section -Title "Exchange Online" -AnchorId 'exchange' -CsvFiles $exchangeFiles.FullName -SummaryHtml ($exchangeSummary -join "`n") -ModuleVersion (Get-ModuleScriptVersion -ScriptName 'Invoke-ExchangeAudit.ps1')))
}


# =========================================
# ===   Mail Security Section           ===
# =========================================
$mailSecFiles = @(Get-ChildItem "$mailSecDir\MailSec_*.csv" -ErrorAction SilentlyContinue | Sort-Object Name)

if ($mailSecFiles.Count -gt 0) {
    $mailSecSummary = [System.Collections.Generic.List[string]]::new()

    # --- Section header stat chips ---
    $_msSpfStr = '&mdash;'; $_msDmarcStr = '&mdash;'; $_msDkimStr = '&mdash;'
    $_msSpfClass = 'neutral'; $_msDmarcClass = 'neutral'; $_msDkimClass = 'neutral'
    $_msSpfCsvStat = Join-Path $mailSecDir "MailSec_SPF.csv"
    if (Test-Path $_msSpfCsvStat) {
        $_msSpfAll   = @(Import-Csv $_msSpfCsvStat | Where-Object { $_.Domain -notlike '*.onmicrosoft.com' })
        $_msSpfPass  = @($_msSpfAll | Where-Object { $_.SPF -ne 'Not Found' -and $_.SPF -ne '' }).Count
        $_msSpfTotal = $_msSpfAll.Count
        if ($_msSpfTotal -gt 0) {
            $_msSpfPct   = [math]::Round(($_msSpfPass / $_msSpfTotal) * 100)
            $_msSpfStr   = "${_msSpfPct}%"
            $_msSpfClass = if ($_msSpfPct -eq 100) { 'ok' } elseif ($_msSpfPct -ge 80) { 'warn' } else { 'critical' }
        }
    }
    $_msDmarcCsvStat = Join-Path $mailSecDir "MailSec_DMARC.csv"
    if (Test-Path $_msDmarcCsvStat) {
        $_msDmarcAll   = @(Import-Csv $_msDmarcCsvStat | Where-Object { $_.Domain -notlike '*.onmicrosoft.com' })
        $_msDmarcPass  = @($_msDmarcAll | Where-Object { $_.DMARC -ne 'Not Found' -and $_.DMARC -ne '' }).Count
        $_msDmarcTotal = $_msDmarcAll.Count
        if ($_msDmarcTotal -gt 0) {
            $_msDmarcPct   = [math]::Round(($_msDmarcPass / $_msDmarcTotal) * 100)
            $_msDmarcStr   = "${_msDmarcPct}%"
            $_msDmarcClass = if ($_msDmarcPct -eq 100) { 'ok' } elseif ($_msDmarcPct -ge 80) { 'warn' } else { 'critical' }
        }
    }
    $_msDkimCsvStat = Join-Path $mailSecDir "MailSec_DKIM.csv"
    if (Test-Path $_msDkimCsvStat) {
        $_msDkimAll   = @(Import-Csv $_msDkimCsvStat | Where-Object { $_.Domain -notlike '*.onmicrosoft.com' })
        $_msDkimPass  = @($_msDkimAll | Where-Object { $_.DKIMEnabled -eq 'True' }).Count
        $_msDkimTotal = $_msDkimAll.Count
        if ($_msDkimTotal -gt 0) {
            $_msDkimPct   = [math]::Round(($_msDkimPass / $_msDkimTotal) * 100)
            $_msDkimStr   = "${_msDkimPct}%"
            $_msDkimClass = if ($_msDkimPct -eq 100) { 'ok' } elseif ($_msDkimPct -ge 80) { 'warn' } else { 'critical' }
        }
    }
    $_msAiCrit  = @($actionItems | Where-Object { $_.Category -like 'Mail Security*' -and $_.Severity -eq 'critical' }).Count
    $_msAiWarn  = @($actionItems | Where-Object { $_.Category -like 'Mail Security*' -and $_.Severity -eq 'warning'  }).Count
    $_msAiClass = if ($_msAiCrit -gt 0) { 'critical' } elseif ($_msAiWarn -gt 0) { 'warn' } else { 'ok' }
    $_msAiStr   = if ($_msAiCrit -gt 0) { "$_msAiCrit critical" } elseif ($_msAiWarn -gt 0) { "$_msAiWarn warnings" } else { 'None' }
    $mailSecSummary.Add(@"
<div class='section-stats'>
  <a class='stat-chip $_msSpfClass' href='#mailsec-records'><div class='stat-chip-value'>$_msSpfStr</div><div class='stat-chip-label'>SPF Coverage</div></a>
  <a class='stat-chip $_msDmarcClass' href='#mailsec-records'><div class='stat-chip-value'>$_msDmarcStr</div><div class='stat-chip-label'>DMARC Coverage</div></a>
  <a class='stat-chip $_msDkimClass' href='#mailsec-records'><div class='stat-chip-value'>$_msDkimStr</div><div class='stat-chip-label'>DKIM Coverage</div></a>
  <a class='stat-chip $_msAiClass' href='#mailsec'><div class='stat-chip-value'>$_msAiStr</div><div class='stat-chip-label'>Action Items</div></a>
</div>
"@)

    $dkimCsv  = Join-Path $mailSecDir "MailSec_DKIM.csv"
    $dmarcCsv = Join-Path $mailSecDir "MailSec_DMARC.csv"
    $spfCsv   = Join-Path $mailSecDir "MailSec_SPF.csv"

    # Load all three into hashtables keyed by domain, excluding onmicrosoft.com
    $dkimByDomain  = @{}
    $dmarcByDomain = @{}
    $spfByDomain   = @{}

    if (Test-Path $dkimCsv) {
        foreach ($row in (Import-Csv $dkimCsv | Where-Object { $_.Domain -notlike "*.onmicrosoft.com" })) {
            $dkimByDomain[$row.Domain] = $row
        }
    }
    if (Test-Path $dmarcCsv) {
        foreach ($row in (Import-Csv $dmarcCsv | Where-Object { $_.Domain -notlike "*.onmicrosoft.com" })) {
            $dmarcByDomain[$row.Domain] = $row
        }
    }
    if (Test-Path $spfCsv) {
        foreach ($row in (Import-Csv $spfCsv | Where-Object { $_.Domain -notlike "*.onmicrosoft.com" })) {
            $spfByDomain[$row.Domain] = $row
        }
    }

    # Union of all domains across the three CSVs
    $allDomains = @($dkimByDomain.Keys + $dmarcByDomain.Keys + $spfByDomain.Keys) | Select-Object -Unique | Sort-Object

    # Summary counts
    $dkimOkCount  = ($dkimByDomain.Values  | Where-Object { $_.DKIMEnabled -eq "True" }).Count
    $dmarcOkCount = ($dmarcByDomain.Values | Where-Object { $_.DMARC -ne "Not Found" -and $_.DMARC }).Count
    $spfOkCount   = ($spfByDomain.Values   | Where-Object { $_.SPF -ne "DNS query failed" -and $_.SPF }).Count
    $total        = $allDomains.Count

    $mailSecSummary.Add("<h4 id='mailsec-records'>DNS Records</h4>")
    $mailSecSummary.Add("<p class='$(if ($dkimOkCount -eq $total) {"ok"} else {"warn"})'>DKIM: <b>$dkimOkCount / $total</b> domains signing enabled</p>")
    $mailSecSummary.Add("<p class='$(if ($dmarcOkCount -eq $total) {"ok"} else {"warn"})'>DMARC: <b>$dmarcOkCount / $total</b> domains configured</p>")
    $mailSecSummary.Add("<p class='$(if ($spfOkCount -eq $total) {"ok"} else {"warn"})'>SPF: <b>$spfOkCount / $total</b> domains configured</p>")

    # Per-domain tables
    foreach ($domain in $allDomains) {
        $dkimRow  = $dkimByDomain[$domain]
        $dmarcRow = $dmarcByDomain[$domain]
        $spfRow   = $spfByDomain[$domain]

        # SPF row
        $spfVal   = if ($spfRow -and $spfRow.SPF -and $spfRow.SPF -ne "DNS query failed") { $spfRow.SPF } else { $null }
        $spfClass = if ($spfVal) { "" } else { " class='warn'" }
        $spfDisp  = if ($spfVal) { "<code style='font-size:0.8rem;word-break:break-all'>$spfVal</code>" } else { "<span class='warn'>Not configured</span>" }

        # DMARC row
        $dmarcVal   = if ($dmarcRow -and $dmarcRow.DMARC -and $dmarcRow.DMARC -ne "Not Found") { $dmarcRow.DMARC } else { $null }
        $dmarcClass = if ($dmarcVal) { "" } else { " class='warn'" }
        $dmarcDisp  = if ($dmarcVal) { "<code style='font-size:0.8rem;word-break:break-all'>$dmarcVal</code>" } else { "<span class='warn'>Not configured</span>" }

        # DKIM row
        $dkimEnabled = $dkimRow -and $dkimRow.DKIMEnabled -eq "True"
        $dkimClass   = if ($dkimEnabled) { "" } else { " class='warn'" }
        $dkimDisp    = if ($dkimEnabled) {
            "Enabled &mdash; <span style='font-size:0.8rem;color:#555'>$($dkimRow.Selector1CNAME)</span>"
        } else {
            "<span class='warn'>$(if ($dkimRow) { $dkimRow.DKIMEnabled } else { 'Not configured' })</span>"
        }

        $mailSecSummary.Add(@"
<h4>$domain</h4>
<table>
  <thead><tr><th style='width:80px'>Record</th><th>Value</th></tr></thead>
  <tbody>
    <tr$spfClass><td>SPF</td><td>$spfDisp</td></tr>
    <tr$dmarcClass><td>DMARC</td><td>$dmarcDisp</td></tr>
    <tr$dkimClass><td>DKIM</td><td>$dkimDisp</td></tr>
  </tbody>
</table>
"@)
    }

    $html.Add((Add-Section -Title "Mail Security" -AnchorId 'mailsec' -CsvFiles $mailSecFiles.FullName -SummaryHtml ($mailSecSummary -join "`n") -ModuleVersion (Get-ModuleScriptVersion -ScriptName 'Invoke-MailSecurityAudit.ps1')))
}


# =========================================
# ===   SharePoint / OneDrive Section   ===
# =========================================
$spFiles = @(Get-ChildItem "$spDir\SharePoint_*.csv" -ErrorAction SilentlyContinue | Sort-Object Name)

if ($spFiles.Count -gt 0) {
    $spSummary = [System.Collections.Generic.List[string]]::new()

    # --- Section header stat chips ---
    $_spSiteCount   = 0
    $_spExtCount    = 0
    $_spStorageStr  = '&mdash;'
    $_spSitesCsvStat = Join-Path $spDir "SharePoint_Sites.csv"
    if (Test-Path $_spSitesCsvStat) {
        $_spSiteCount = @(Import-Csv $_spSitesCsvStat).Count
    }
    $_spExtCsvStat = Join-Path $spDir "SharePoint_ExternalSharing_SiteOverrides.csv"
    if (Test-Path $_spExtCsvStat) {
        $_spExtCount = @(Import-Csv $_spExtCsvStat | Where-Object { $_.SharingCapability -ne 'Disabled' -and $_.SharingCapability -ne '' }).Count
    }
    $_spStorageCsvStat = Join-Path $spDir "SharePoint_TenantStorage.csv"
    if (Test-Path $_spStorageCsvStat) {
        $_spSto = Import-Csv $_spStorageCsvStat | Select-Object -First 1
        if ($_spSto) {
            $_spUsedMbStat  = [double]$_spSto.StorageUsedMB
            $_spQuotaMbStat = [double]$_spSto.StorageQuotaMB
            # Fallback: if API returned 0 for used, sum per-site + OD CSVs (same logic as section below)
            if ($_spUsedMbStat -le 0 -and $_spQuotaMbStat -gt 0) {
                $_spSitesCsv = Join-Path $spDir "SharePoint_Sites.csv"
                $_spOdCsv    = Join-Path $spDir "SharePoint_OneDriveUsage.csv"
                $_spUsedMbStat  = 0
                if (Test-Path $_spSitesCsv) { $_spUsedMbStat += [double](Import-Csv $_spSitesCsv | Measure-Object -Property StorageUsedMB -Sum).Sum }
                if (Test-Path $_spOdCsv)    { $_spUsedMbStat += [double](Import-Csv $_spOdCsv    | Measure-Object -Property StorageUsedMB -Sum).Sum }
            }
            $_spUsedGB  = [math]::Round($_spUsedMbStat  / 1024, 1)
            $_spTotalGB = [math]::Round($_spQuotaMbStat / 1024, 1)
            $_spStorageStr = "$_spUsedGB / $_spTotalGB GB"
        }
    }
    $_spAiCrit  = @($actionItems | Where-Object { $_.Category -like 'SharePoint*' -and $_.Severity -eq 'critical' }).Count
    $_spAiWarn  = @($actionItems | Where-Object { $_.Category -like 'SharePoint*' -and $_.Severity -eq 'warning'  }).Count
    $_spAiClass = if ($_spAiCrit -gt 0) { 'critical' } elseif ($_spAiWarn -gt 0) { 'warn' } else { 'ok' }
    $_spAiStr   = if ($_spAiCrit -gt 0) { "$_spAiCrit critical" } elseif ($_spAiWarn -gt 0) { "$_spAiWarn warnings" } else { 'None' }
    $_spExtClass = if ($_spExtCount -gt 0) { 'warn' } else { 'ok' }
    $spSummary.Add(@"
<div class='section-stats'>
  <a class='stat-chip neutral' href='#sp-sites'><div class='stat-chip-value'>$_spSiteCount</div><div class='stat-chip-label'>Sites</div></a>
  <a class='stat-chip neutral' href='#sp-storage'><div class='stat-chip-value'>$_spStorageStr</div><div class='stat-chip-label'>Storage Used</div></a>
  <a class='stat-chip $_spExtClass' href='#sp-sharing'><div class='stat-chip-value'>$_spExtCount</div><div class='stat-chip-label'>External Sharing On</div></a>
  <a class='stat-chip $_spAiClass' href='#sharepoint'><div class='stat-chip-value'>$_spAiStr</div><div class='stat-chip-label'>Action Items</div></a>
</div>
"@)

    $storageCsv     = Join-Path $spDir "SharePoint_TenantStorage.csv"
    $sitesCsv       = Join-Path $spDir "SharePoint_Sites.csv"
    $groupsCsv      = Join-Path $spDir "SharePoint_SPGroups.csv"
    $tenantShareCsv = Join-Path $spDir "SharePoint_ExternalSharing_Tenant.csv"
    $overridesCsv   = Join-Path $spDir "SharePoint_ExternalSharing_SiteOverrides.csv"
    $odUsageCsv     = Join-Path $spDir "SharePoint_OneDriveUsage.csv"
    $unlicensedCsv  = Join-Path $spDir "SharePoint_OneDrive_Unlicensed.csv"
    $acpCsv         = Join-Path $spDir "SharePoint_AccessControlPolicies.csv"

    # --- 1. Tenant Storage ---
    if (Test-Path $storageCsv) {
        $storage   = Import-Csv $storageCsv | Select-Object -First 1
        $quotaMB   = [double]$storage.StorageQuotaMB
        $usedMB    = [double]$storage.StorageUsedMB
        $freeMB    = [double]$storage.AvailableStorageMB
        $odQuotaGB = [math]::Round([double]$storage.OneDriveQuotaMB / 1024, 1)

        # StorageQuotaUsed is sometimes null from Get-PnPTenant — fall back to summing per-site + OD data
        if ($usedMB -le 0 -and $quotaMB -gt 0) {
            $sitesUsedMB = 0
            if (Test-Path $sitesCsv) {
                $sitesUsedMB = [double](Import-Csv $sitesCsv |
                    Measure-Object -Property StorageUsedMB -Sum).Sum
            }
            $odUsedMB = 0
            if (Test-Path $odUsageCsv) {
                $odUsedMB = [double](Import-Csv $odUsageCsv |
                    Measure-Object -Property StorageUsedMB -Sum).Sum
            }
            $usedMB = $sitesUsedMB + $odUsedMB
            $freeMB = $quotaMB - $usedMB
        }

        $usedPct  = if ($quotaMB -gt 0) { [math]::Round(($usedMB / $quotaMB) * 100, 1) } else { 0 }
        $barPct   = [math]::Min($usedPct, 100)
        $barColor = if ($usedPct -gt 90) { "#e53935" } elseif ($usedPct -gt 75) { "#ff9800" } else { "#4caf50" }
        $usedGB   = [math]::Round($usedMB  / 1024, 1)
        $quotaGB  = [math]::Round($quotaMB / 1024, 1)
        $freeGB   = [math]::Round($freeMB  / 1024, 1)

        $storageBar = "<div style='display:flex;align-items:center;gap:10px;margin-bottom:0.4rem'><div style='background:#e0e0e0;border-radius:4px;width:240px;height:16px;flex-shrink:0;overflow:hidden'><div style='background:$barColor;width:${barPct}%;height:16px'></div></div><span style='font-size:0.85rem;color:#555'>$usedPct% used &mdash; $usedGB GB of $quotaGB GB ($freeGB GB free)</span></div>"
        $spSummary.Add("<h4 id='sp-storage'>Tenant Storage</h4>$storageBar<p style='font-size:0.85rem;color:#666;margin-top:0'>OneDrive storage quota per user: $odQuotaGB GB</p>")
    }

    # --- 2. Site Collections ---
    if (Test-Path $sitesCsv) {
        $sites = @(Import-Csv $sitesCsv)

        # Build groups lookup keyed by site URL
        $groupsBySite = @{}
        if (Test-Path $groupsCsv) {
            foreach ($g in (Import-Csv $groupsCsv)) {
                if (-not $groupsBySite.ContainsKey($g.Site)) {
                    $groupsBySite[$g.Site] = [System.Collections.Generic.List[PSCustomObject]]::new()
                }
                $groupsBySite[$g.Site].Add($g)
            }
        }

        $templateLabels = @{
            'SITEPAGEPUBLISHING#0'      = 'Communication'
            'GROUP#0'                   = 'Team (M365)'
            'STS#0'                     = 'Classic Team'
            'STS#3'                     = 'Team'
            'GLOBAL#0'                  = 'Root Site'
            'EHS#1'                     = 'Team Site'
            'SPSPERS#0'                 = 'OneDrive'
            'SPSMSITEHOST#0'            = 'MySite Host'
            'APPCATALOG#0'              = 'App Catalog'
            'SRCHCEN#0'                 = 'Search Center'
            'SRCHCENTERLITE#0'          = 'Search Center'
            'EDISC#0'                   = 'eDiscovery'
            'TEAMCHANNEL#0'             = 'Teams Channel'
            'TEAMCHANNEL#1'             = 'Teams Channel'
            'PWA#0'                     = 'Project Web App'
            'RedirectSite#0'            = 'Redirect'
            'POINTPUBLISHINGTOPIC#0'    = 'Publishing Topic'
            'POINTPUBLISHINGPERSONAL#2' = 'Publishing Personal'
            'BLANKINTERNET#0'           = 'Publishing'
            'BLANKINTERNETCONTAINER#0'  = 'Publishing Portal'
            'ENTERWIKI#0'               = 'Enterprise Wiki'
        }

        # Count sharing link group types across all sites for the summary line
        $_sharingLinkCounts = @{ Anonymous = 0; Organisation = 0; SpecificPeople = 0; Other = 0 }
        foreach ($_slg in ($groupsBySite.Values | ForEach-Object { $_ })) {
            if ($_slg.GroupName -match '^SharingLinks\..+?\.(Anonymous|Organization|Flexible)') {
                switch -Regex ($Matches[1]) {
                    'Anonymous'    { $_sharingLinkCounts.Anonymous++ }
                    'Organization' { $_sharingLinkCounts.Organisation++ }
                    'Flexible'     { $_sharingLinkCounts.SpecificPeople++ }
                    default        { $_sharingLinkCounts.Other++ }
                }
            }
        }
        $_totalSharingLinks = $_sharingLinkCounts.Anonymous + $_sharingLinkCounts.Organisation + $_sharingLinkCounts.SpecificPeople + $_sharingLinkCounts.Other

        # Derive admin centre URL from first site URL (e.g. https://contoso.sharepoint.com → https://contoso-admin.sharepoint.com)
        $_spAdminUrl = if ($sites.Count -gt 0) {
            $sites[0].Url -replace 'https://([^.]+)\.sharepoint\.com.*', 'https://$1-admin.sharepoint.com'
        } else { 'https://admin.microsoft.com' }
        $_sharingReportUrl = "$_spAdminUrl/_layouts/15/online/ExternalSharingReportManager.aspx"

        $spSummary.Add("<h4 id='sp-sites'>Site Collections ($($sites.Count))</h4>")
        $spSummary.Add((Get-ExpandHintHtml -Text 'Click a row to expand SharePoint groups for that site.'))

        $siteRows = foreach ($site in ($sites | Sort-Object Title)) {
            $storMB = if ($site.StorageUsedMB -and $site.StorageUsedMB -ne '') { [double]$site.StorageUsedMB } else { 0 }
            $storGB = [math]::Round($storMB / 1024, 2)
            $storCell = if ($storMB -gt 0) { "<span style='font-size:0.85rem'>$storGB GB</span>" } else { "<span style='font-size:0.85rem;color:#888'>—</span>" }

            $hubBadge = if ($site.IsHubSite -eq 'True') { " <span style='background:#e3f2fd;color:#1565c0;border:1px solid #90caf9;border-radius:3px;font-size:0.72rem;padding:1px 5px'>Hub</span>" } else { "" }

            $templateLabel = if ($templateLabels.ContainsKey($site.Template)) { $templateLabels[$site.Template] } `
                             elseif ($site.Template -like 'TEAMCHANNEL*') { 'Teams Channel' } `
                             elseif ($site.Template -like 'SPSPERS*')     { 'OneDrive' } `
                             elseif ($site.Template -like 'GROUP*')       { 'Team (M365)' } `
                             else { $site.Template }

            $urlPath  = $site.Url -replace '^https://[^/]+', ''
            $urlCell  = "<a href='$($site.Url)' target='_blank' style='font-size:0.8rem;color:#1565c0;word-break:break-all'>$urlPath</a>"

            $mainRow = "<tr class='user-row' onclick='togglePerms(this)' title='Click to show/hide groups'><td>$($site.Title)$hubBadge</td><td>$urlCell</td><td style='font-size:0.85rem'>$templateLabel</td><td>$storCell</td><td style='font-size:0.85rem'>$($site.Owner)</td></tr>"

            # Groups detail row
            $siteGroups = if ($groupsBySite.ContainsKey($site.Url)) { @($groupsBySite[$site.Url]) } else { @() }
            # Filter out system-generated and sharing link groups — counted separately below
            $visibleGroups = @($siteGroups | Where-Object {
                $_.GroupName -notmatch '^Limited Access System Group For (List|Web) ' -and
                $_.GroupName -notmatch '^SharingLinks\.'
            })
            if ($visibleGroups.Count -gt 0) {
                $groupRows = ($visibleGroups | ForEach-Object {
                    # Resolve SharePoint claim tokens to readable labels
                    $cleanMembers = if ($_.Members) {
                        ($_.Members -split ';\s*' | ForEach-Object {
                            $m = $_.Trim()
                            $resolved = if     ($m -match '^i:0#\.f\|membership\|(.+)$')               { $Matches[1] }
                                        elseif ($m -match '^c:0t\.c\|tenant\|')                        { '[All org users]' }
                                        elseif ($m -match '^c:0-\.f\|rolemanager\|spo-grid-all-users') { '[Everyone]' }
                                        elseif ($m -match '^c:0o\.c\|federateddirectoryclaimprovider\|') { '[M365 Group]' }
                                        elseif ($m -match '^c:0\(\.s\|true')                           { '[All authenticated users]' }
                                        elseif ($m -eq 'SHAREPOINT\system')                            { '[SharePoint System]' }
                                        elseif ($m)                                                    { ($m -split '\|')[-1] }
                            if (-not $resolved) { return }
                            if ($resolved -match '^[0-9a-f]{40,}$') { return }
                            $color = if     ($resolved -match '^\[')                         { $null }
                                     elseif ($resolved -match '#ext#')                       { 'darkorange' }
                                     elseif ($resolved -match '^[^@\s]+@[^@\s]+\.[^@\s]+$') { 'green' }
                                     else                                                    { $null }
                            if ($color) { "<span style='color:$color'>$resolved</span>" } else { $resolved }
                        } | Where-Object { $_ }) -join ', '
                    } else { '—' }
                    "<tr><td>$(ConvertTo-HtmlText $_.GroupName)</td><td>$($_.Owner)</td><td style='text-align:center'>$($_.MemberCount)</td><td style='font-size:0.8rem;word-break:break-all'>$cleanMembers</td></tr>"
                }) -join ""
                $detailRow = "<tr class='signin-detail' style='display:none'><td colspan='5'><table class='inner-table'><thead><tr><th>Group</th><th>Owner</th><th>Members</th><th>Member Accounts</th></tr></thead><tbody>$groupRows</tbody></table></td></tr>"
            }
            else {
                $detailRow = "<tr class='signin-detail' style='display:none'><td colspan='5'><em style='color:#888'>No groups retrieved for this site — access may have been denied during the audit.</em></td></tr>"
            }

            $mainRow
            $detailRow
        }

        $spSummary.Add(@"
<table>
  <thead><tr>
    <th>Site Title</th><th>URL Path</th><th>Template</th><th>Storage Used</th><th>Owner</th>
  </tr></thead>
  <tbody>$($siteRows -join "`n")</tbody>
</table>
"@)

        # Sharing links summary line
        if ($_totalSharingLinks -gt 0) {
            $_slParts = [System.Collections.Generic.List[string]]::new()
            if ($_sharingLinkCounts.Anonymous     -gt 0) { $_slParts.Add("<span style='color:#e65100;font-weight:bold'>$($_sharingLinkCounts.Anonymous) anonymous</span>") }
            if ($_sharingLinkCounts.Organisation  -gt 0) { $_slParts.Add("<span style='color:#1565c0'>$($_sharingLinkCounts.Organisation) organisation-wide</span>") }
            if ($_sharingLinkCounts.SpecificPeople -gt 0) { $_slParts.Add("<span style='color:#6a1b9a'>$($_sharingLinkCounts.SpecificPeople) specific people</span>") }
            if ($_sharingLinkCounts.Other         -gt 0) { $_slParts.Add("$($_sharingLinkCounts.Other) other") }
            $_slColor = if ($_sharingLinkCounts.Anonymous -gt 0) { 'warn' } else { '' }
            $_slStyle = if ($_slColor) { " class='$_slColor'" } else { '' }
            $spSummary.Add("<p$_slStyle>$_totalSharingLinks sharing link(s) detected across all sites: $($_slParts -join ', ') &mdash; <a href='$_sharingReportUrl' target='_blank'>Generate External Sharing Report in SharePoint Admin Centre</a></p>")
        }
    }

    # --- 3. External Sharing ---
    $spSummary.Add("<h4 id='sp-sharing'>External Sharing</h4>")
    $spSummary.Add("<p style='font-size:0.82rem;color:#64748b'>References: <a href='https://www.cisecurity.org/benchmark/microsoft_365' target='_blank'>CIS M365 Benchmark</a> (7.1–7.3) · <a href='https://learn.microsoft.com/en-us/sharepoint/turn-external-sharing-on-or-off' target='_blank'>Microsoft SharePoint External Sharing</a>.</p>")

    if (Test-Path $tenantShareCsv) {
        $ts = Import-Csv $tenantShareCsv | Select-Object -First 1

        $sharingLabel = switch ($ts.SharingCapability) {
            'Disabled'                          { "<span class='ok'>Disabled — no external sharing</span>" }
            'ExternalUserSharingOnly'           { "<span class='warn'>Authenticated guests only</span>" }
            'ExternalUserAndGuestSharing'       { "<span class='critical'>Anyone (including anonymous)</span>" }
            'ExistingExternalUserSharingOnly'   { "<span class='warn'>Existing external users only</span>" }
            default                             { $ts.SharingCapability }
        }

        $linkTypeLabel = switch ($ts.DefaultSharingLinkType) {
            'AnonymousAccess' { "<span class='critical'>Anyone (anonymous) — no sign-in required</span>" }
            'Internal'        { "<span class='ok'>Only people in your organisation</span>" }
            'Direct'          { "<span class='ok'>Specific people</span>" }
            default           { $ts.DefaultSharingLinkType }
        }

        $anonExpiry = if ([int]$ts.RequireAnonymousLinksExpireInDays -gt 0) {
            "<span class='ok'>$($ts.RequireAnonymousLinksExpireInDays) days</span>"
        } else {
            "<span class='warn'>No expiry configured</span>"
        }

        $domainRestrict = switch ($ts.SharingDomainRestrictionMode) {
            'AllowList' { "Allow-list: $($ts.SharingAllowedDomainList)" }
            'BlockList' { "Block-list configured" }
            default     { "<span style='color:#888'>None</span>" }
        }

        $infectedClass   = if ($ts.DisallowInfectedFileDownload -eq 'False') { " class='warn'" } else { '' }
        $reshareClass    = if ($ts.PreventExternalUsersFromResharing -eq 'False') { " class='warn'" } else { '' }
        $expiryReqClass  = if ($ts.ExternalUserExpirationRequired -eq 'False') { " class='warn'" } else { '' }

        $spSummary.Add(@"
<table style='max-width:720px'>
  <thead><tr><th>Setting</th><th>Value</th></tr></thead>
  <tbody>
    <tr><td>Tenant Sharing Capability</td><td>$sharingLabel</td></tr>
    <tr><td>Default Sharing Link Type</td><td>$linkTypeLabel</td></tr>
    <tr><td>Anonymous Link Expiry</td><td>$anonExpiry</td></tr>
    <tr><td>Domain Restrictions</td><td>$domainRestrict</td></tr>
    <tr$infectedClass><td>Infected file download blocked</td><td>$(ConvertTo-HtmlText $ts.DisallowInfectedFileDownload)</td></tr>
    <tr$reshareClass><td>External users can reshare</td><td>$(ConvertTo-HtmlText $ts.PreventExternalUsersFromResharing)</td></tr>
    <tr$expiryReqClass><td>Guest link expiry required</td><td>$(ConvertTo-HtmlText $ts.ExternalUserExpirationRequired)</td></tr>
    <tr><td>Guest link expiry (days)</td><td>$(ConvertTo-HtmlText $ts.ExternalUserExpireInDays)</td></tr>
    <tr><td>Accepting account must match invite</td><td>$(ConvertTo-HtmlText $ts.RequireAcceptingAccountMatchInvitedAccount)</td></tr>
  </tbody>
</table>
"@)
    }

    if (Test-Path $overridesCsv) {
        $overrides = @(Import-Csv $overridesCsv)
        if ($overrides.Count -gt 0) {
            $ovRows = ($overrides | ForEach-Object {
                $sharingClass = if ($_.SharingCapability -eq 'ExternalUserAndGuestSharing') { " class='warn'" } else { "" }
                $storGB = [math]::Round([double]$_.SiteStorageMB / 1024, 2)
                $urlPath = $_.Url -replace '^https://[^/]+', ''
                "<tr><td><a href='$($_.Url)' target='_blank' style='font-size:0.85rem'>$($_.Title)</a><br><span style='font-size:0.78rem;color:#888'>$urlPath</span></td><td$sharingClass>$($_.SharingCapability)</td><td>$storGB GB</td></tr>"
            }) -join ""
            $spSummary.Add("<p class='warn' style='margin-top:0.75rem'>$($overrides.Count) site(s) override the tenant sharing policy:</p>")
            $spSummary.Add("<table style='max-width:720px'><thead><tr><th>Site</th><th>Sharing Setting</th><th>Storage</th></tr></thead><tbody>$ovRows</tbody></table>")
        }
        else {
            $spSummary.Add("<p class='ok'>No site-level external sharing overrides.</p>")
        }
    }

    # --- 4. Access Control Policies ---
    if (Test-Path $acpCsv) {
        $acp = Import-Csv $acpCsv | Select-Object -First 1
        $spSummary.Add("<h4>Access Control Policies</h4>")
        $spSummary.Add("<p style='font-size:0.82rem;color:#64748b'>References: <a href='https://www.cisecurity.org/benchmark/microsoft_365' target='_blank'>CIS M365 Benchmark</a> (7.3–7.4) · <a href='https://learn.microsoft.com/en-us/sharepoint/control-access-from-unmanaged-devices' target='_blank'>Microsoft Unmanaged Device Policy</a>.</p>")

        $syncRestrict = if ($acp.IsUnmanagedSyncAppForTenantRestricted -eq 'True') {
            "<span class='ok'>Restricted — only managed/domain-joined devices can sync</span>"
        } else {
            "<span class='warn'>Not restricted — any device can sync SharePoint/OneDrive data</span>"
        }

        $caPolicy = switch ($acp.ConditionalAccessPolicy) {
            'AllowFullAccess'    { "<span class='warn'>No restriction (full access from any device)</span>" }
            'AllowLimitedAccess' { "<span class='ok'>Limited access — unmanaged devices get browser-only</span>" }
            'BlockAccess'        { "<span class='ok'>Block — unmanaged device access denied</span>" }
            default              { if ($acp.ConditionalAccessPolicy) { $acp.ConditionalAccessPolicy } else { "<span style='color:#888'>—</span>" } }
        }

        $idleSignOut = if ($acp.IdleSessionSignOutEnabled -eq 'True') {
            "<span class='ok'>Enabled — sign out after $($acp.SignOutAfterMinutesOfInactivity) minutes</span>"
        } else {
            "<span class='warn'>Disabled</span>"
        }

        $ipEnforce = if ($acp.IPAddressEnforcement -eq 'True') { "<span class='ok'>Enabled</span>" } else { "<span style='color:#888'>Disabled</span>" }
        $macSync   = if ($acp.BlockMacSync -eq 'True') { "Blocked (legacy Mac sync client)" } else { "Allowed" }

        $spSummary.Add(@"
<table style='max-width:720px'>
  <thead><tr><th>Policy</th><th>Status</th></tr></thead>
  <tbody>
    <tr><td>Sync — managed devices only</td><td>$syncRestrict</td></tr>
    <tr><td>Conditional Access policy</td><td>$caPolicy</td></tr>
    <tr><td>Idle session sign-out</td><td>$idleSignOut</td></tr>
    <tr><td>IP address restriction</td><td>$ipEnforce</td></tr>
    <tr><td>Legacy Mac sync client</td><td>$macSync</td></tr>
  </tbody>
</table>
"@)
    }

    # --- 5. OneDrive ---
    $spSummary.Add("<h4 id='sp-onedrive'>OneDrive</h4>")

    if (Test-Path $odUsageCsv) {
        $odDrives  = @(Import-Csv $odUsageCsv)
        $totalOdGB = [math]::Round(($odDrives | Measure-Object -Property StorageUsedMB -Sum).Sum / 1024, 1)
        $spSummary.Add("<p>$($odDrives.Count) OneDrive account(s) — $totalOdGB GB total in use</p>")
        if ($odDrives.Count -gt 0) {
            $_odRows = ($odDrives | Sort-Object { -[double]$_.StorageUsedMB } | ForEach-Object {
                $odGB    = [math]::Round([double]$_.StorageUsedMB / 1024, 2)
                $urlPath = $_.OneDriveUrl -replace '^https://[^/]+', ''
                "<tr><td>$(ConvertTo-HtmlText $_.OwnerUPN)</td><td><a href='$($_.OneDriveUrl)' target='_blank' style='font-size:0.8rem'>$urlPath</a></td><td style='text-align:right'>$odGB GB</td></tr>"
            }) -join ""
            $spSummary.Add(@"
<details>
  <summary style='cursor:pointer;font-size:0.84rem;color:#475569;margin-bottom:0.4rem'>Show all $($odDrives.Count) OneDrive accounts (sorted by size)</summary>
  <table style='margin-top:0.4rem'><thead><tr><th>Owner UPN</th><th>OneDrive Path</th><th style='text-align:right'>Size</th></tr></thead>
  <tbody>$_odRows</tbody></table>
</details>
"@)
        }
    }

    if (Test-Path $unlicensedCsv) {
        $unlicensed = @(Import-Csv $unlicensedCsv)
        if ($unlicensed.Count -gt 0) {
            $ulRows = ($unlicensed | ForEach-Object {
                $odGB = [math]::Round([double]$_.StorageUsedMB / 1024, 2)
                $urlPath = $_.OneDriveUrl -replace '^https://[^/]+', ''
                "<tr><td>$($_.OwnerUPN)</td><td><a href='$($_.OneDriveUrl)' target='_blank' style='font-size:0.8rem'>$urlPath</a></td><td>$odGB GB</td></tr>"
            }) -join ""
            $spSummary.Add("<p class='warn'>$($unlicensed.Count) OneDrive account(s) belong to unlicensed users — data may be inaccessible and storage costs wasted:</p>")
            $spSummary.Add("<table><thead><tr><th>UPN</th><th>OneDrive Path</th><th>Storage Used</th></tr></thead><tbody>$ulRows</tbody></table>")
        }
        else {
            $spSummary.Add("<p class='ok'>All OneDrive accounts belong to licensed users.</p>")
        }
    }

    $html.Add((Add-Section -Title "SharePoint / OneDrive" -AnchorId 'sharepoint' -CsvFiles $spFiles.FullName -SummaryHtml ($spSummary -join "`n") -ModuleVersion (Get-ModuleScriptVersion -ScriptName 'Invoke-SharePointAudit.ps1')))
}


# =========================================
# ===   Microsoft Teams Section         ===
# =========================================
$teamsFiles = @(Get-ChildItem -Path $teamsDir -Filter "Teams_*.csv" -ErrorAction SilentlyContinue)
if ($teamsFiles.Count -gt 0) {
    $teamsSummary = [System.Collections.Generic.List[string]]::new()

    # --- Section header stat chips ---
    $_tmFedStr   = '&mdash;'; $_tmFedClass = 'neutral'
    $_tmGuestStr = '&mdash;'; $_tmGuestClass = 'neutral'
    $_tmFedCsvStat = Join-Path $teamsDir "Teams_FederationConfig.csv"
    if (Test-Path $_tmFedCsvStat) {
        $_tmFed      = Import-Csv $_tmFedCsvStat | Select-Object -First 1
        $_tmFedStr   = if ($_tmFed.AllowFederatedUsers -eq 'True') { 'Enabled' } else { 'Disabled' }
        $_tmFedClass = if ($_tmFed.AllowFederatedUsers -eq 'True' -and [int]$_tmFed.AllowedDomainsCount -eq 0 -and [int]$_tmFed.BlockedDomainsCount -eq 0) { 'warn' } else { 'ok' }
    }
    $_tmClientCsvStat = Join-Path $teamsDir "Teams_ClientConfig.csv"
    if (Test-Path $_tmClientCsvStat) {
        $_tmClient   = Import-Csv $_tmClientCsvStat | Select-Object -First 1
        $_tmGuestStr = if ($_tmClient.AllowGuestUser -eq 'True') { 'Enabled' } else { 'Disabled' }
        $_tmGuestClass = 'neutral'
    }
    $_tmMtgPoliciesCount = '&mdash;'
    $_tmMtgCsvStat = Join-Path $teamsDir "Teams_MeetingPolicies.csv"
    if (Test-Path $_tmMtgCsvStat) { $_tmMtgPoliciesCount = @(Import-Csv $_tmMtgCsvStat).Count }
    $_tmAiCrit  = @($actionItems | Where-Object { $_.Category -like 'Teams*' -and $_.Severity -eq 'critical' }).Count
    $_tmAiWarn  = @($actionItems | Where-Object { $_.Category -like 'Teams*' -and $_.Severity -eq 'warning'  }).Count
    $_tmAiClass = if ($_tmAiCrit -gt 0) { 'critical' } elseif ($_tmAiWarn -gt 0) { 'warn' } else { 'ok' }
    $_tmAiStr   = if ($_tmAiCrit -gt 0) { "$_tmAiCrit critical" } elseif ($_tmAiWarn -gt 0) { "$_tmAiWarn warnings" } else { 'None' }
    $teamsSummary.Add(@"
<div class='section-stats'>
  <a class='stat-chip $_tmFedClass' href='#teams-external'><div class='stat-chip-value'>$_tmFedStr</div><div class='stat-chip-label'>External Federation</div></a>
  <a class='stat-chip $_tmGuestClass' href='#teams-external'><div class='stat-chip-value'>$_tmGuestStr</div><div class='stat-chip-label'>Guest Access</div></a>
  <a class='stat-chip neutral' href='#teams-meetings'><div class='stat-chip-value'>$_tmMtgPoliciesCount</div><div class='stat-chip-label'>Meeting Policies</div></a>
  <a class='stat-chip $_tmAiClass' href='#teams'><div class='stat-chip-value'>$_tmAiStr</div><div class='stat-chip-label'>Action Items</div></a>
</div>
"@)

    $teamsFedCsv       = Join-Path $teamsDir "Teams_FederationConfig.csv"
    $teamsClientCsv    = Join-Path $teamsDir "Teams_ClientConfig.csv"
    $teamsMtgCsv       = Join-Path $teamsDir "Teams_MeetingPolicies.csv"
    $teamsGuestMtgCsv  = Join-Path $teamsDir "Teams_GuestMeetingConfig.csv"
    $teamsGuestCallCsv = Join-Path $teamsDir "Teams_GuestCallingConfig.csv"
    $teamsAppPermCsv   = Join-Path $teamsDir "Teams_AppPermissionPolicies.csv"
    $teamsAppSetupCsv  = Join-Path $teamsDir "Teams_AppSetupPolicies.csv"
    $teamsChanCsv      = Join-Path $teamsDir "Teams_ChannelPolicies.csv"

    # Federation / External Access
    if (Test-Path $teamsFedCsv) {
        $_teamsFed = Import-Csv $teamsFedCsv | Select-Object -First 1
        $teamsSummary.Add("<h4 id='teams-external'>External Access / Federation</h4>")
        $teamsSummary.Add("<p style='font-size:0.82rem;color:#64748b'>References: <a href='https://www.cisecurity.org/benchmark/microsoft_365' target='_blank'>CIS M365 Benchmark</a> (8.2.1–8.2.3) · <a href='https://learn.microsoft.com/en-us/microsoftteams/manage-external-access' target='_blank'>Microsoft Teams External Access</a>.</p>")
        $teamsSummary.Add("<table class='summary-table'><thead><tr><th>Setting</th><th>Value</th></tr></thead><tbody>")

        $_fedRows = @(
            @{ Label = 'Federated (managed) users allowed';       Value = $_teamsFed.AllowFederatedUsers },
            @{ Label = 'Consumer (non-managed) users allowed';    Value = $_teamsFed.AllowPublicUsers;             Warn = 'True' },
            @{ Label = 'Teams Consumer outbound';                 Value = $_teamsFed.AllowTeamsConsumer },
            @{ Label = 'Teams Consumer inbound';                  Value = $_teamsFed.AllowTeamsConsumerInbound;    Warn = 'True' },
            @{ Label = 'Treat discovered partners as unverified'; Value = $_teamsFed.TreatDiscoveredPartnersAsUnverified },
            @{ Label = 'Allowed domain count';                    Value = $_teamsFed.AllowedDomainsCount },
            @{ Label = 'Blocked domain count';                    Value = $_teamsFed.BlockedDomainsCount }
        )
        foreach ($_fr in $_fedRows) {
            $_color = if ($_fr.Warn -and $_fr.Value -eq 'True') { "style='color:#e65100'" } else { '' }
            $teamsSummary.Add("<tr><td>$($_fr.Label)</td><td $_color>$(ConvertTo-HtmlText $_fr.Value)</td></tr>")
        }
        $teamsSummary.Add("</tbody></table>")

        if ($_teamsFed.AllowedDomainsList) {
            $teamsSummary.Add("<p><strong>Allowed domains:</strong> $(ConvertTo-HtmlText $_teamsFed.AllowedDomainsList)</p>")
        }
    }

    # Client Config (guest access + cloud storage)
    if (Test-Path $teamsClientCsv) {
        $_teamsClient = Import-Csv $teamsClientCsv | Select-Object -First 1
        $teamsSummary.Add("<hr class='section-divider'>")
        $teamsSummary.Add("<h4>Teams Client Configuration</h4>")
        $teamsSummary.Add("<p style='font-size:0.82rem;color:#64748b'>References: <a href='https://www.cisecurity.org/benchmark/microsoft_365' target='_blank'>CIS M365 Benchmark</a> (8.5.x) · <a href='https://learn.microsoft.com/en-us/microsoftteams/teams-client-configuration' target='_blank'>Microsoft Teams Client Configuration</a>.</p>")
        $teamsSummary.Add("<table class='summary-table'><thead><tr><th>Setting</th><th>Value</th></tr></thead><tbody>")

        $_clientRows = @(
            @{ Label = 'Guest access enabled';          Value = $_teamsClient.AllowGuestUser },
            @{ Label = 'Skype for Business interop';    Value = $_teamsClient.AllowSkypeBusinessInterop },
            @{ Label = 'Box storage allowed';           Value = $_teamsClient.AllowBox;          Warn = 'True' },
            @{ Label = 'Dropbox storage allowed';       Value = $_teamsClient.AllowDropBox;      Warn = 'True' },
            @{ Label = 'Egnyte storage allowed';        Value = $_teamsClient.AllowEgnyte;       Warn = 'True' },
            @{ Label = 'Google Drive storage allowed';  Value = $_teamsClient.AllowGoogleDrive;  Warn = 'True' },
            @{ Label = 'ShareFile storage allowed';     Value = $_teamsClient.AllowShareFile;    Warn = 'True' }
        )
        foreach ($_cr in $_clientRows) {
            $_color = if ($_cr.Warn -and $_cr.Value -eq 'True') { "style='color:#e65100'" } else { '' }
            $teamsSummary.Add("<tr><td>$($_cr.Label)</td><td $_color>$(ConvertTo-HtmlText $_cr.Value)</td></tr>")
        }
        $teamsSummary.Add("</tbody></table>")
    }

    # Meeting Policies
    if (Test-Path $teamsMtgCsv) {
        $_teamsMtg = @(Import-Csv $teamsMtgCsv)
        $teamsSummary.Add("<hr class='section-divider'>")
        $teamsSummary.Add("<h4 id='teams-meetings'>Meeting Policies ($($_teamsMtg.Count))</h4>")
        $teamsSummary.Add("<table class='summary-table'><thead><tr><th>Policy</th><th>Anon Can Start</th><th>Auto Admit Users</th><th>Cloud Recording</th></tr></thead><tbody>")
        foreach ($_mp in ($_teamsMtg | Sort-Object { if ($_.Identity -eq 'Global') { 0 } else { 1 } }, Identity)) {
            $_anonColor  = if ($_mp.AllowAnonymousUsersToStartMeeting -eq 'True') { "style='color:#c62828;font-weight:bold'" } else { '' }
            $_admitColor = if ($_mp.AutoAdmittedUsers -eq 'Everyone') { "style='color:#e65100'" } else { '' }
            $teamsSummary.Add("<tr><td>$(ConvertTo-HtmlText $_mp.Identity)</td><td $_anonColor>$(ConvertTo-HtmlText $_mp.AllowAnonymousUsersToStartMeeting)</td><td $_admitColor>$(ConvertTo-HtmlText $_mp.AutoAdmittedUsers)</td><td>$(ConvertTo-HtmlText $_mp.AllowCloudRecording)</td></tr>")
        }
        $teamsSummary.Add("</tbody></table>")
    }

    # App Policies
    if ((Test-Path $teamsAppSetupCsv) -or (Test-Path $teamsAppPermCsv)) {
        $teamsSummary.Add("<hr class='section-divider'>")
        $teamsSummary.Add("<h4 id='teams-apps'>App Policies</h4>")

        if (Test-Path $teamsAppSetupCsv) {
            $_teamsAppSetup = @(Import-Csv $teamsAppSetupCsv)
            $teamsSummary.Add("<p><strong>App Setup Policies ($($_teamsAppSetup.Count))</strong></p>")
            $teamsSummary.Add("<table class='summary-table'><thead><tr><th>Policy</th><th>Sideloading Allowed</th><th>User Pinning</th></tr></thead><tbody>")
            foreach ($_asp in ($_teamsAppSetup | Sort-Object { if ($_.Identity -eq 'Global') { 0 } else { 1 } }, Identity)) {
                $_sideColor = if ($_asp.AllowSideloading -eq 'True') { "style='color:#e65100'" } else { '' }
                $teamsSummary.Add("<tr><td>$(ConvertTo-HtmlText $_asp.Identity)</td><td $_sideColor>$(ConvertTo-HtmlText $_asp.AllowSideloading)</td><td>$(ConvertTo-HtmlText $_asp.AllowUserPinning)</td></tr>")
            }
            $teamsSummary.Add("</tbody></table>")
        }

        if (Test-Path $teamsAppPermCsv) {
            $_teamsAppPerm = @(Import-Csv $teamsAppPermCsv)
            $teamsSummary.Add("<p><strong>App Permission Policies ($($_teamsAppPerm.Count))</strong></p>")
            $teamsSummary.Add("<table class='summary-table'><thead><tr><th>Policy</th><th>Microsoft Apps</th><th>Third-Party Apps</th><th>Custom Apps</th></tr></thead><tbody>")
            foreach ($_app in ($_teamsAppPerm | Sort-Object { if ($_.Identity -eq 'Global') { 0 } else { 1 } }, Identity)) {
                $teamsSummary.Add("<tr><td>$(ConvertTo-HtmlText $_app.Identity)</td><td>$(ConvertTo-HtmlText $_app.DefaultCatalogApps)</td><td>$(ConvertTo-HtmlText $_app.GlobalCatalogApps)</td><td>$(ConvertTo-HtmlText $_app.PrivateCatalogApps)</td></tr>")
            }
            $teamsSummary.Add("</tbody></table>")
        }
    }

    # Channel Policies
    if (Test-Path $teamsChanCsv) {
        $_teamsChan = @(Import-Csv $teamsChanCsv)
        $teamsSummary.Add("<hr class='section-divider'>")
        $teamsSummary.Add("<h4>Channel Policies ($($_teamsChan.Count))</h4>")
        $teamsSummary.Add("<table class='summary-table'><thead><tr><th>Policy</th><th>Org-wide Team Creation</th><th>Shared Channels</th><th>Private Channels</th></tr></thead><tbody>")
        foreach ($_cp in ($_teamsChan | Sort-Object { if ($_.Identity -eq 'Global') { 0 } else { 1 } }, Identity)) {
            $teamsSummary.Add("<tr><td>$(ConvertTo-HtmlText $_cp.Identity)</td><td>$(ConvertTo-HtmlText $_cp.AllowOrgWideTeamCreation)</td><td>$(ConvertTo-HtmlText $_cp.AllowSharedChannels)</td><td>$(ConvertTo-HtmlText $_cp.AllowPrivateChannels)</td></tr>")
        }
        $teamsSummary.Add("</tbody></table>")
    }

    # Guest Meeting + Calling
    if ((Test-Path $teamsGuestMtgCsv) -or (Test-Path $teamsGuestCallCsv)) {
        $teamsSummary.Add("<hr class='section-divider'>")
        $teamsSummary.Add("<h4>Guest Access Settings</h4>")
        $teamsSummary.Add("<table class='summary-table'><thead><tr><th>Setting</th><th>Value</th></tr></thead><tbody>")
        if (Test-Path $teamsGuestMtgCsv) {
            $_gm = Import-Csv $teamsGuestMtgCsv | Select-Object -First 1
            $teamsSummary.Add("<tr><td>Guest — IP video allowed</td><td>$(ConvertTo-HtmlText $_gm.AllowIPVideo)</td></tr>")
            $teamsSummary.Add("<tr><td>Guest — screen sharing mode</td><td>$(ConvertTo-HtmlText $_gm.ScreenSharingMode)</td></tr>")
            $teamsSummary.Add("<tr><td>Guest — Meet Now allowed</td><td>$(ConvertTo-HtmlText $_gm.AllowMeetNow)</td></tr>")
        }
        if (Test-Path $teamsGuestCallCsv) {
            $_gc = Import-Csv $teamsGuestCallCsv | Select-Object -First 1
            $teamsSummary.Add("<tr><td>Guest — private calling allowed</td><td>$(ConvertTo-HtmlText $_gc.AllowPrivateCalling)</td></tr>")
        }
        $teamsSummary.Add("</tbody></table>")
    }

    $html.Add((Add-Section -Title "Microsoft Teams" -AnchorId 'teams' -CsvFiles $teamsFiles.FullName -SummaryHtml ($teamsSummary -join "`n") -ModuleVersion (Get-ModuleScriptVersion -ScriptName 'Invoke-TeamsAudit.ps1')))
}


# =========================================
# ===   Intune / Endpoint Section       ===
# =========================================
$intuneFiles = @(Get-ChildItem "$intuneDir\Intune_*.csv" -ErrorAction SilentlyContinue | Sort-Object Name)

if ($intuneFiles.Count -gt 0) {
    $intuneSummary = [System.Collections.Generic.List[string]]::new()

    # --- Section header stat chips (reuse KPI values already computed) ---
    $_itDevStr      = if ($null -ne $_kpiDevCount)     { "$_kpiDevCount" }       else { '&mdash;' }
    $_itNcStr       = if ($null -ne $_kpiNonCompliant) { "$_kpiNonCompliant" }   else { '&mdash;' }
    $_itNcClass     = if ($null -eq $_kpiNonCompliant -or $_kpiNonCompliant -eq 0) { 'ok' } else { 'critical' }
    $_itComplPct    = if ($null -ne $_kpiDevCount -and $_kpiDevCount -gt 0 -and $null -ne $_kpiNonCompliant) {
        [math]::Round((($_kpiDevCount - $_kpiNonCompliant) / $_kpiDevCount) * 100)
    } else { $null }
    $_itComplStr    = if ($null -ne $_itComplPct) { "${_itComplPct}%" } else { '&mdash;' }
    $_itComplClass  = if ($null -eq $_itComplPct -or $_itComplPct -eq 100) { 'ok' } elseif ($_itComplPct -ge 80) { 'warn' } else { 'critical' }
    $_itAiCrit      = @($actionItems | Where-Object { $_.Category -like 'Intune*' -and $_.Severity -eq 'critical' }).Count
    $_itAiWarn      = @($actionItems | Where-Object { $_.Category -like 'Intune*' -and $_.Severity -eq 'warning'  }).Count
    $_itAiClass     = if ($_itAiCrit -gt 0) { 'critical' } elseif ($_itAiWarn -gt 0) { 'warn' } else { 'ok' }
    $_itAiStr       = if ($_itAiCrit -gt 0) { "$_itAiCrit critical" } elseif ($_itAiWarn -gt 0) { "$_itAiWarn warnings" } else { 'None' }
    $intuneSummary.Add(@"
<div class='section-stats'>
  <a class='stat-chip neutral' href='#intune-devices'><div class='stat-chip-value'>$_itDevStr</div><div class='stat-chip-label'>Managed Devices</div></a>
  <a class='stat-chip $_itComplClass' href='#intune-compliance'><div class='stat-chip-value'>$_itComplStr</div><div class='stat-chip-label'>Compliant</div></a>
  <a class='stat-chip $_itNcClass' href='#intune-compliance'><div class='stat-chip-value'>$_itNcStr</div><div class='stat-chip-label'>Non-Compliant</div></a>
  <a class='stat-chip $_itAiClass' href='#intune'><div class='stat-chip-value'>$_itAiStr</div><div class='stat-chip-label'>Action Items</div></a>
</div>
"@)

    $intLicCsv         = Join-Path $intuneDir "Intune_LicenceCheck.csv"
    $intDevCsv         = Join-Path $intuneDir "Intune_Devices.csv"
    $intDevStatesCsv   = Join-Path $intuneDir "Intune_DeviceComplianceStates.csv"
    $intPolCsv         = Join-Path $intuneDir "Intune_CompliancePolicies.csv"
    $intPolSetCsv  = Join-Path $intuneDir "Intune_CompliancePolicySettings.csv"
    $intProfCsv    = Join-Path $intuneDir "Intune_ConfigProfiles.csv"
    $intProfSetCsv = Join-Path $intuneDir "Intune_ConfigProfileSettings.csv"
    $intAppCsv     = Join-Path $intuneDir "Intune_Apps.csv"
    $intApCsv      = Join-Path $intuneDir "Intune_AutopilotDevices.csv"
    $intEnrolCsv   = Join-Path $intuneDir "Intune_EnrollmentRestrictions.csv"

    # Licence check
    $_intLicRow = $null
    if (Test-Path $intLicCsv) { $_intLicRow = Import-Csv $intLicCsv | Select-Object -First 1 }

    if ($null -eq $_intLicRow -or $_intLicRow.HasIntune -ne 'True') {
        $intuneSummary.Add("<p><em>No Intune-capable licence was detected on this tenant. Intune device management data was not collected.</em></p>")
    }
    else {
        $_skuFriendlyNames = @{
            'SPB'                  = 'Microsoft 365 Business Premium'
            'BUSINESS_PREMIUM'     = 'Microsoft 365 Business Premium'
            'ENTERPRISEPREMIUM'    = 'Microsoft 365 E5'
            'ENTERPRISEPACK'       = 'Microsoft 365 E3'
            'M365_F1'              = 'Microsoft 365 F1'
            'M365_F3'              = 'Microsoft 365 F3'
            'INTUNE_A'             = 'Microsoft Intune Plan 1'
            'INTUNE_A_D'           = 'Microsoft Intune Plan 1 for Education'
            'INTUNE_P2'            = 'Microsoft Intune Plan 2'
            'EMS'                  = 'Enterprise Mobility + Security E3'
            'EMS_S_1'              = 'Enterprise Mobility + Security E3'
            'EMS_S_3'              = 'Enterprise Mobility + Security E3'
            'EMS_S_5'              = 'Enterprise Mobility + Security E5'
            'EMSPREMIUM'           = 'Enterprise Mobility + Security E5'
        }
        $_skuDisplay = ($($_intLicRow.LicencedSKUs) -split ',\s*' | ForEach-Object {
            $sku = $_.Trim()
            if ($_skuFriendlyNames.ContainsKey($sku)) { "$($_skuFriendlyNames[$sku]) ($sku)" } else { $sku }
        }) -join ', '
        $intuneSummary.Add("<p><strong>Licenced SKUs:</strong> $_skuDisplay</p>")

        # Device inventory
        if (Test-Path $intDevCsv) {
            $_intDevices  = @(Import-Csv $intDevCsv)
            $_totalDev    = $_intDevices.Count
            $_osCounts    = $_intDevices | Group-Object OS | Sort-Object Count -Descending
            $_compCounts  = $_intDevices | Group-Object ComplianceState | Sort-Object Count -Descending
            $_corpDev     = @($_intDevices | Where-Object { $_.OwnerType -eq 'company' }).Count
            $_persDev     = @($_intDevices | Where-Object { $_.OwnerType -eq 'personal' }).Count

            $intuneSummary.Add("<h4 id='intune-devices'>Device Inventory ($_totalDev total)</h4>")
            $intuneSummary.Add("<p>Corporate: $_corpDev &nbsp;|&nbsp; Personal (BYOD): $_persDev</p>")

            $intuneSummary.Add("<table class='summary-table'><thead><tr><th>Operating System</th><th>Count</th></tr></thead><tbody>")
            foreach ($_osGroup in $_osCounts) {
                $intuneSummary.Add("<tr><td>$(ConvertTo-HtmlText $_osGroup.Name)</td><td>$($_osGroup.Count)</td></tr>")
            }
            $intuneSummary.Add("</tbody></table>")

            # Build per-device compliance state lookup
            $_devStatesMap = @{}
            if (Test-Path $intDevStatesCsv) {
                foreach ($_ds in (Import-Csv $intDevStatesCsv)) {
                    if (-not $_devStatesMap.ContainsKey($_ds.DeviceName)) {
                        $_devStatesMap[$_ds.DeviceName] = [System.Collections.Generic.List[object]]::new()
                    }
                    $_devStatesMap[$_ds.DeviceName].Add($_ds)
                }
            }

            $intuneSummary.Add("<h4>Managed Devices</h4>")
            $intuneSummary.Add((Get-ExpandHintHtml -Text 'Click a row to expand per-policy compliance states for that device.'))
            $intuneSummary.Add("<table class='summary-table'><thead><tr><th>Device</th><th>OS</th><th>OS Version</th><th>Type</th><th>Owner</th><th>Compliance</th><th>Assigned User</th><th>Manufacturer</th><th>Model</th><th>Last Sync</th><th>Agent</th></tr></thead><tbody>")
            foreach ($_dev in ($_intDevices | Sort-Object DeviceName, OS, LastSyncDateTime)) {
                $_complianceStyle = switch ($_dev.ComplianceState) {
                    'compliant'    { 'color:#388e3c;font-weight:bold' }
                    'noncompliant' { 'color:#c62828;font-weight:bold' }
                    default        { '' }
                }
                $_syncDt = [datetime]::MinValue
                $_syncStale = $_dev.LastSyncDateTime -and [datetime]::TryParse($_dev.LastSyncDateTime, [ref]$_syncDt) -and (([datetime]::UtcNow - $_syncDt).TotalDays -gt 30)
                $_syncStyle = if ($_syncStale) { "background:#ffebee;color:#b71c1c;font-weight:bold" } else { "" }
                $intuneSummary.Add("<tr class='user-row' onclick='togglePerms(this)' title='Click to show/hide compliance policy states'><td>$(ConvertTo-HtmlText $_dev.DeviceName)</td><td>$(ConvertTo-HtmlText $_dev.OS)</td><td>$(ConvertTo-HtmlText $_dev.OSVersion)</td><td>$(ConvertTo-HtmlText $_dev.DeviceType)</td><td>$(ConvertTo-HtmlText $_dev.OwnerType)</td><td style='$_complianceStyle'>$(ConvertTo-HtmlText $_dev.ComplianceState)</td><td>$(ConvertTo-HtmlText $_dev.AssignedUser)</td><td>$(ConvertTo-HtmlText $_dev.Manufacturer)</td><td>$(ConvertTo-HtmlText $_dev.Model)</td><td style='$_syncStyle'>$(ConvertTo-HtmlText $_dev.LastSyncDateTime)</td><td>$(ConvertTo-HtmlText $_dev.ManagementAgent)</td></tr>")

                # Detail row — per-policy compliance states (deduplicated: worst state wins per policy name)
                $_devPolicyStates = if ($_devStatesMap.ContainsKey($_dev.DeviceName)) { @($_devStatesMap[$_dev.DeviceName]) } else { @() }
                if ($_devPolicyStates.Count -gt 0) {
                    $_stateRank = @{ 'error' = 4; 'nonCompliant' = 3; 'unknown' = 2; 'notApplicable' = 1; 'compliant' = 0 }
                    $_deduped = $_devPolicyStates |
                        Group-Object PolicyName |
                        ForEach-Object {
                            $_.Group | Sort-Object { if ($_stateRank.ContainsKey($_.State)) { $_stateRank[$_.State] } else { -1 } } -Descending | Select-Object -First 1
                        } | Sort-Object @{ Expression = { if ($_stateRank.ContainsKey($_.State)) { $_stateRank[$_.State] } else { -1 } }; Descending = $true }, @{ Expression = 'PolicyName'; Descending = $false }
                    $_policyStateRows = ($_deduped | ForEach-Object {
                        $_ps = $_
                        $_psStyle = switch ($_ps.State) {
                            'compliant'    { 'color:#388e3c;font-weight:bold' }
                            'nonCompliant' { 'color:#c62828;font-weight:bold' }
                            'error'        { 'color:#e65100;font-weight:bold' }
                            default        { '' }
                        }
                        "<tr><td>$(ConvertTo-HtmlText $_ps.PolicyName)</td><td style='$_psStyle'>$(ConvertTo-HtmlText $_ps.State)</td><td>$(ConvertTo-HtmlText $_ps.LastReportedDateTime)</td></tr>"
                    }) -join ""
                    $intuneSummary.Add("<tr class='signin-detail' style='display:none'><td colspan='11'><table class='inner-table'><thead><tr><th>Policy</th><th>State</th><th>Last Reported</th></tr></thead><tbody>$_policyStateRows</tbody></table></td></tr>")
                }
                else {
                    $intuneSummary.Add("<tr class='signin-detail' style='display:none'><td colspan='11'><em style='color:#888'>No per-policy compliance state data available for this device.</em></td></tr>")
                }
            }
            $intuneSummary.Add("</tbody></table>")

            $_compLine = ($_compCounts | ForEach-Object {
                $s = switch ($_.Name) { 'compliant' { 'color:#388e3c;font-weight:bold' } 'noncompliant' { 'color:#c62828;font-weight:bold' } default { '' } }
                if ($s) { "<span style='$s'>$($_.Count) $($_.Name)</span>" } else { "$($_.Count) $($_.Name)" }
            }) -join ' &nbsp;&bull;&nbsp; '
            $intuneSummary.Add("<p style='margin-top:0.4rem;font-size:0.87rem;color:#444'>Compliance breakdown: $_compLine</p>")
        }

        # Compliance policies
        if (Test-Path $intPolCsv) {
            $_intPols = @(Import-Csv $intPolCsv)
            $_intPolSettings = @()
            if (Test-Path $intPolSetCsv) { $_intPolSettings = @(Import-Csv $intPolSetCsv) }

            $intuneSummary.Add("<h4 id='intune-compliance'>Compliance Policies ($($_intPols.Count))</h4>")
            if ($_intPols.Count -gt 0) {
                $intuneSummary.Add((Get-ExpandHintHtml -Text 'Click a row to expand policy details and settings.'))
                $intuneSummary.Add("<table class='summary-table'><thead><tr><th>Policy</th><th>Platform</th><th>Assigned To</th><th>Grace Period (h)</th><th>Last Modified</th></tr></thead><tbody>")
                $_policyRows = foreach ($_pol in ($_intPols | Sort-Object Platform, PolicyName)) {
                    $_grace = 0
                    [int]::TryParse($_pol.GracePeriodHours, [ref]$_grace) | Out-Null
                    $_graceColor = if ($_grace -gt 24) { " style='color:#e65100;font-weight:bold'" } else { '' }
                    $_polSettings = if ($_pol.PolicyId) {
                        @($_intPolSettings | Where-Object { $_.PolicyId -eq $_pol.PolicyId })
                    } else {
                        @($_intPolSettings | Where-Object { $_.PolicyName -eq $_pol.PolicyName })
                    }

                    $_mainRow = "<tr class='user-row' onclick='togglePerms(this)' title='Click to expand policy details'><td>$(ConvertTo-HtmlText $_pol.PolicyName)</td><td>$(ConvertTo-HtmlText $_pol.Platform)</td><td>$(ConvertTo-HtmlText $_pol.AssignedTo)</td><td$_graceColor>$($_pol.GracePeriodHours)</td><td>$(ConvertTo-HtmlText $_pol.LastModifiedDateTime)</td></tr>"

                    $_detailInner = [System.Collections.Generic.List[string]]::new()
                    $_detailInner.Add("<tr><td style='width:160px'>Policy Type</td><td>$(ConvertTo-HtmlText $_pol.PolicyType)</td></tr>")
                    $_detailInner.Add("<tr><td>Description</td><td>$(ConvertTo-HtmlMultilineText $_pol.Description)</td></tr>")
                    $_detailInner.Add("<tr><td>Assignments</td><td>$(ConvertTo-HtmlMultilineText $_pol.AssignmentDetails)</td></tr>")
                    if ($_polSettings.Count -gt 0) {
                        $_settingsRows = foreach ($_s in $_polSettings) {
                            "<tr><td>$(ConvertTo-HtmlText $_s.SettingName)</td><td>$(ConvertTo-HtmlMultilineText $_s.SettingValue)</td></tr>"
                        }
                        $_detailInner.Add("<tr><td>Settings</td><td><table class='inner-table'><thead><tr><th>Setting</th><th>Value</th></tr></thead><tbody>$($_settingsRows -join '')</tbody></table></td></tr>")
                    }
                    else {
                        $_detailInner.Add("<tr><td>Settings</td><td><span class='detail-empty'>No policy settings exported.</span></td></tr>")
                    }

                    $_detailRow = "<tr class='signin-detail' style='display:none'><td colspan='5'><table class='inner-table'><thead><tr><th style='width:160px'>Field</th><th>Value</th></tr></thead><tbody>$($_detailInner -join '')</tbody></table></td></tr>"
                    $_mainRow
                    $_detailRow
                }
                $intuneSummary.Add($($_policyRows -join "`n"))
                $intuneSummary.Add("</tbody></table>")
            } else {
                $intuneSummary.Add("<p style='color:#b71c1c'><strong>No compliance policies found.</strong> All devices are considered compliant by default.</p>")
            }
        }

        # Configuration profiles / policies
        if (Test-Path $intProfCsv) {
            $_intProfs = @(Import-Csv $intProfCsv)
            $_intProfSettings = @()
            if (Test-Path $intProfSetCsv) { $_intProfSettings = @(Import-Csv $intProfSetCsv) }

            $intuneSummary.Add("<h4 id='intune-config'>Configuration Policies / Profiles ($($_intProfs.Count))</h4>")
            if ($_intProfs.Count -gt 0) {
                $intuneSummary.Add((Get-ExpandHintHtml -Text 'Click a row to expand profile details and settings.'))
                $intuneSummary.Add("<table class='summary-table'><thead><tr><th>Profile</th><th>Platform</th><th>Type</th><th>Source</th><th>Last Modified</th><th>Assigned To</th></tr></thead><tbody>")
                $_profileRows = foreach ($_prof in ($_intProfs | Sort-Object Platform, ProfileName)) {
                    $_profSettings = if ($_prof.ProfileId) {
                        @($_intProfSettings | Where-Object { $_.ProfileId -eq $_prof.ProfileId })
                    } else {
                        @($_intProfSettings | Where-Object { $_.ProfileName -eq $_prof.ProfileName })
                    }

                    $_mainRow = "<tr class='user-row' onclick='togglePerms(this)' title='Click to expand profile details'><td>$(ConvertTo-HtmlText $_prof.ProfileName)</td><td>$(ConvertTo-HtmlText $_prof.Platform)</td><td>$(ConvertTo-HtmlText $_prof.ProfileType)</td><td>$(ConvertTo-HtmlText $_prof.Source)</td><td>$(ConvertTo-HtmlText $_prof.LastModifiedDateTime)</td><td>$(ConvertTo-HtmlText $_prof.AssignedTo)</td></tr>"

                    $_detailInner = [System.Collections.Generic.List[string]]::new()
                    $_detailInner.Add("<tr><td style='width:160px'>Description</td><td>$(ConvertTo-HtmlMultilineText $_prof.Description)</td></tr>")
                    $_detailInner.Add("<tr><td>Assignments</td><td>$(ConvertTo-HtmlMultilineText $_prof.AssignmentDetails)</td></tr>")
                    if ($_profSettings.Count -gt 0) {
                        $_settingsRows = foreach ($_s in $_profSettings) {
                            "<tr><td>$(ConvertTo-HtmlText $_s.SettingName)</td><td>$(ConvertTo-HtmlMultilineText $_s.SettingValue)</td></tr>"
                        }
                        $_detailInner.Add("<tr><td>Settings</td><td><table class='inner-table'><thead><tr><th>Setting</th><th>Value</th></tr></thead><tbody>$($_settingsRows -join '')</tbody></table></td></tr>")
                    } else {
                        $_detailInner.Add("<tr><td>Settings</td><td><span class='detail-empty'>No profile settings exported.</span></td></tr>")
                    }
                    $_detailRow = "<tr class='signin-detail' style='display:none'><td colspan='6'><table class='inner-table'><thead><tr><th style='width:160px'>Field</th><th>Value</th></tr></thead><tbody>$($_detailInner -join '')</tbody></table></td></tr>"
                    $_mainRow
                    $_detailRow
                }
                $intuneSummary.Add($($_profileRows -join "`n"))
                $intuneSummary.Add("</tbody></table>")
            }
        }

        # Apps
        if (Test-Path $intAppCsv) {
            $_intApps = @(Import-Csv $intAppCsv)
            $intuneSummary.Add("<h4 id='intune-apps'>Apps ($($_intApps.Count))</h4>")
            if ($_intApps.Count -gt 0) {
                $intuneSummary.Add((Get-ExpandHintHtml -Text 'Click a row to expand app deployment details.'))
                $intuneSummary.Add("<table class='summary-table'><thead><tr><th>App</th><th>Type</th><th>Assigned To</th><th>Installed</th><th>Failed</th><th>Pending</th></tr></thead><tbody>")
                $_appRows = foreach ($_app in ($_intApps | Sort-Object AppName)) {
                    $_failedCount = 0
                    [int]::TryParse($_app.FailedDeviceCount, [ref]$_failedCount) | Out-Null
                    $_failColor = if ($_failedCount -gt 0) { " style='color:#c62828;font-weight:bold'" } else { '' }

                    $_mainRow = "<tr class='user-row' onclick='togglePerms(this)' title='Click to expand app details'><td>$(ConvertTo-HtmlText $_app.AppName)</td><td>$(ConvertTo-HtmlText $_app.AppType)</td><td>$(ConvertTo-HtmlText $_app.AssignedTo)</td><td>$(ConvertTo-HtmlText $_app.InstalledDeviceCount)</td><td$_failColor>$(ConvertTo-HtmlText $_app.FailedDeviceCount)</td><td>$(ConvertTo-HtmlText $_app.PendingInstallCount)</td></tr>"
                    $_detailInner = @(
                        "<tr><td style='width:160px'>Publisher</td><td>$(ConvertTo-HtmlText $_app.Publisher)</td></tr>",
                        "<tr><td>Description</td><td>$(ConvertTo-HtmlMultilineText $_app.Description)</td></tr>",
                        "<tr><td>Assignments</td><td>$(ConvertTo-HtmlMultilineText $_app.AssignmentDetails)</td></tr>",
                        "<tr><td>Installation Summary</td><td><table class='inner-table'><thead><tr><th>Installed</th><th>Failed</th><th>Pending</th></tr></thead><tbody><tr><td>$(ConvertTo-HtmlText $_app.InstalledDeviceCount)</td><td>$(ConvertTo-HtmlText $_app.FailedDeviceCount)</td><td>$(ConvertTo-HtmlText $_app.PendingInstallCount)</td></tr></tbody></table></td></tr>"
                    )
                    $_detailRow = "<tr class='signin-detail' style='display:none'><td colspan='6'><table class='inner-table'><thead><tr><th style='width:160px'>Field</th><th>Value</th></tr></thead><tbody>$($_detailInner -join '')</tbody></table></td></tr>"
                    $_mainRow
                    $_detailRow
                }
                $intuneSummary.Add($($_appRows -join "`n"))
                $intuneSummary.Add("</tbody></table>")
            }
        }

        # Autopilot
        if (Test-Path $intApCsv) {
            $_intAp = @(Import-Csv $intApCsv)
            if ($_intAp.Count -gt 0) {
                $intuneSummary.Add("<p><strong>Windows Autopilot:</strong> $($_intAp.Count) device(s) registered.</p>")
            } else {
                $intuneSummary.Add("<p><strong>Windows Autopilot:</strong> No devices registered.</p>")
            }
        }

        # Enrollment restrictions
        if (Test-Path $intEnrolCsv) {
            $_intEnrol = @(Import-Csv $intEnrolCsv)
            $intuneSummary.Add("<h4>Enrollment Restrictions ($($_intEnrol.Count))</h4>")
            if ($_intEnrol.Count -gt 0) {
                $intuneSummary.Add("<table class='summary-table'><thead><tr><th>Config</th><th>Platform</th><th>Block Personal</th><th>Max Devices/User</th><th>Priority</th><th>Assigned To</th></tr></thead><tbody>")
                foreach ($_er in $_intEnrol) {
                    $_blockColor = if ($_er.BlockPersonalDevices -eq 'False') { "color:#e65100" } else { "" }
                    $intuneSummary.Add("<tr><td>$($_er.ConfigName)</td><td>$($_er.Platform)</td><td style='$_blockColor'>$($_er.BlockPersonalDevices)</td><td>$($_er.MaxDevicesPerUser)</td><td>$($_er.Priority)</td><td>$($_er.AssignedTo)</td></tr>")
                }
                $intuneSummary.Add("</tbody></table>")
            }
        }
    }

    $html.Add((Add-Section -Title "Intune / Endpoint Management" -AnchorId 'intune' -CsvFiles $intuneFiles.FullName -SummaryHtml ($intuneSummary -join "`n") -ModuleVersion (Get-ModuleScriptVersion -ScriptName 'Invoke-IntuneAudit.ps1')))
}


# =========================================
# ===   Technical Issues Section        ===
# =========================================
$_techIssuesCsvPath = Join-Path $AuditFolder "AuditIssues.csv"
if (Test-Path $_techIssuesCsvPath) {
    $_techIssues = @(Import-Csv $_techIssuesCsvPath)

    if ($_techIssues.Count -gt 0) {
        $_tiCritCount = @($_techIssues | Where-Object { $_.Severity -eq 'Critical' }).Count
        $_tiWarnCount = @($_techIssues | Where-Object { $_.Severity -eq 'Warning'  }).Count
        $_tiInfoCount = @($_techIssues | Where-Object { $_.Severity -eq 'Info'     }).Count

        $_tiSummary = [System.Collections.Generic.List[string]]::new()

        # Stat chips
        $_tiCritClass = if ($_tiCritCount -gt 0) { 'critical' } else { 'ok' }
        $_tiWarnClass = if ($_tiWarnCount -gt 0) { 'warn'     } else { 'ok' }
        $_tiSummary.Add(@"
<div class='section-stats'>
  <div class='stat-chip $_tiCritClass'><div class='stat-chip-value'>$_tiCritCount</div><div class='stat-chip-label'>Critical</div></div>
  <div class='stat-chip $_tiWarnClass'><div class='stat-chip-value'>$_tiWarnCount</div><div class='stat-chip-label'>Warnings</div></div>
  <div class='stat-chip neutral'><div class='stat-chip-value'>$_tiInfoCount</div><div class='stat-chip-label'>Info</div></div>
  <div class='stat-chip neutral'><div class='stat-chip-value'>$($_techIssues.Count)</div><div class='stat-chip-label'>Total Issues</div></div>
</div>
"@)

        $_tiRows = foreach ($_ti in ($_techIssues | Sort-Object Severity, Section, Timestamp)) {
            $_tiSevClass = switch ($_ti.Severity) {
                'Critical' { 'issue-sev-critical' }
                'Warning'  { 'issue-sev-warning'  }
                default    { 'issue-sev-info'      }
            }
            $_tiAction = if ($_ti.Action) { "<br><span style='color:#555;font-size:0.88em'>&#8594; $([System.Net.WebUtility]::HtmlEncode($_ti.Action))</span>" } else { "" }
            "<tr><td>$($_ti.Timestamp)</td><td><span class='$_tiSevClass'>$($_ti.Severity)</span></td><td>$($_ti.Section)</td><td><code>$([System.Net.WebUtility]::HtmlEncode($_ti.Collector))</code></td><td>$([System.Net.WebUtility]::HtmlEncode($_ti.Description))$_tiAction</td></tr>"
        }

        $_tiSummary.Add("<p>Collection errors and permission issues encountered during this audit run. Resolve these to ensure complete coverage.</p>")
        $_tiSummary.Add(@"
<table>
  <thead><tr><th>Timestamp</th><th>Severity</th><th>Section</th><th>Collector</th><th>Description / Action</th></tr></thead>
  <tbody>$($_tiRows -join "`n  ")</tbody>
</table>
"@)

        $html.Add((Add-Section -Title "Technical Issues" -AnchorId 'tech-issues' -CsvFiles @($_techIssuesCsvPath) -SummaryHtml ($_tiSummary -join "`n")))
    }
}


# =========================================
# ===   ScubaGear Baseline Section      ===
# =========================================
if ($_scubaResults) {
    $_sgSummary = [System.Collections.Generic.List[string]]::new()

    # Count result types directly from control data (accurate for filter chip labels)
    $_sgByResult = @{}
    foreach ($_p in $_scubaResults.Results.PSObject.Properties) {
        foreach ($_grp in $_p.Value) {
            foreach ($_ctrl in $_grp.Controls) {
                $_r = [string]$_ctrl.Result
                if ($_sgByResult.ContainsKey($_r)) { $_sgByResult[$_r]++ } else { $_sgByResult[$_r] = 1 }
            }
        }
    }
    $_sgTotal = 0; foreach ($_v in $_sgByResult.Values) { $_sgTotal += $_v }

    # Also tally from Summary for the per-product table
    $_sgTotalPass = 0; $_sgTotalFail = 0; $_sgTotalWarn = 0; $_sgTotalManual = 0
    foreach ($_p in $_scubaResults.Summary.PSObject.Properties) {
        $_sgTotalPass   += [int]$_p.Value.Passes
        $_sgTotalFail   += [int]$_p.Value.Failures
        $_sgTotalWarn   += [int]$_p.Value.Warnings
        $_sgTotalManual += [int]$_p.Value.Manual
    }

    # Filter chips — clicking hides/shows that result type across all product tables
    $_sgChipDefs = [ordered]@{
        'Fail'    = if (($_sgByResult['Fail']    -gt 0)) { 'critical' } else { 'ok' }
        'Warning' = if (($_sgByResult['Warning'] -gt 0)) { 'warn'     } else { 'ok' }
        'Pass'    = 'ok'
        'Manual'  = 'neutral'
        'N/A'     = 'neutral'
        'Error'   = 'critical'
    }
    $_sgChipsHtml = foreach ($_res in $_sgChipDefs.Keys) {
        $_cnt = if ($_sgByResult.ContainsKey($_res)) { $_sgByResult[$_res] } else { 0 }
        if ($_cnt -eq 0) { continue }
        "<div class='stat-chip $($_sgChipDefs[$_res])' data-scuba-filter='$_res' onclick='scubaFilter(this)'><div class='stat-chip-value'>$_cnt</div><div class='stat-chip-label'>$_res</div></div>"
    }
    # "All Controls" chip resets filters
    $_sgChipsHtml += "<div class='stat-chip neutral' data-scuba-filter='all' onclick='scubaFilter(this)' title='Show all results'><div class='stat-chip-value'>$_sgTotal</div><div class='stat-chip-label'>All Controls</div></div>"

    $_sgSummary.Add("<p style='font-size:0.82rem;color:#64748b;margin-bottom:0.75rem'>Assessment provided by <a href='https://www.cisa.gov/resources-tools/services/secure-cloud-business-applications-scuba-project' target='_blank'>CISA ScubaGear</a> — an open-source tool maintained by the Cybersecurity and Infrastructure Security Agency (CISA) that evaluates Microsoft 365 tenants against the <a href='https://www.cisa.gov/resources-tools/resources/secure-cloud-business-applications-scuba-project' target='_blank'>SCuBA M365 Security Configuration Baselines</a>.</p>")
    $_sgSummary.Add("<div class='section-stats'>$($_sgChipsHtml -join '')</div>")

    # Per-product summary table with links to detail subsections
    $_sgProductLabels    = @{ AAD='Identity (AAD)'; EXO='Exchange Online'; SharePoint='SharePoint'; Teams='Teams'; Defender='Defender'; PowerPlatform='Power Platform' }
    $_sgProductAnchorIds = @{ AAD='scuba-aad'; EXO='scuba-exo'; SharePoint='scuba-sharepoint'; Teams='scuba-teams'; Defender='scuba-defender'; PowerPlatform='scuba-powerplatform' }

    $_sgRows = foreach ($_p in $_scubaResults.Summary.PSObject.Properties) {
        $_label    = if ($_sgProductLabels.ContainsKey($_p.Name)) { $_sgProductLabels[$_p.Name] } else { $_p.Name }
        $_anchor   = if ($_sgProductAnchorIds.ContainsKey($_p.Name)) { $_sgProductAnchorIds[$_p.Name] } else { 'scuba-summary' }
        $_passes   = [int]$_p.Value.Passes
        $_fails    = [int]$_p.Value.Failures
        $_warns    = [int]$_p.Value.Warnings
        $_manual   = [int]$_p.Value.Manual
        $_failHtml = if ($_fails -gt 0) { "<span class='critical'>$_fails</span>" } else { "$_fails" }
        $_warnHtml = if ($_warns -gt 0) { "<span class='warn'>$_warns</span>" } else { "$_warns" }
        "<tr><td><a href='#$_anchor'>$_label</a></td><td style='color:#2e7d32'>$_passes</td><td>$_failHtml</td><td>$_warnHtml</td><td style='color:#666'>$_manual</td></tr>"
    }

    $_sgSummary.Add(@"
<h4 id='scuba-summary'>CIS M365 Foundations Baseline — Per-Product Summary</h4>
<table><thead><tr><th>Product</th><th>Passed</th><th>Failed</th><th>Warnings</th><th>Manual / N/A</th></tr></thead>
<tbody>$($_sgRows -join '')</tbody></table>
"@)

    if ($_scubaHtmlPath) {
        $_scubaRelPath = [System.IO.Path]::GetRelativePath($script:ReportBaseDir, $_scubaHtmlPath) -replace '\\', '/'
        $_sgSummary.Add("<p style='font-size:0.85rem;margin-top:0.5rem'><a href='$_scubaRelPath' target='_blank'>&#128196; Open full ScubaGear report</a></p>")
    }

    # ── Per-product control detail tables ──────────────────────────────────────
    foreach ($_sgProdName in @('AAD','EXO','SharePoint','Teams','Defender','PowerPlatform')) {
        if ($_scubaResults.Results.PSObject.Properties.Name -notcontains $_sgProdName) { continue }
        $_sgProdData = $_scubaResults.Results.$_sgProdName
        if (-not $_sgProdData) { continue }

        $_sgProdLabel    = if ($_sgProductLabels.ContainsKey($_sgProdName)) { $_sgProductLabels[$_sgProdName] } else { $_sgProdName }
        $_sgProdAnchorId = if ($_sgProductAnchorIds.ContainsKey($_sgProdName)) { $_sgProductAnchorIds[$_sgProdName] } else { "scuba-$($_sgProdName.ToLower())" }

        $_sgCtrlRows = [System.Collections.Generic.List[string]]::new()

        foreach ($_sgGrp in $_sgProdData) {
            # Group header row
            $_sgGrpName = [System.Net.WebUtility]::HtmlEncode($_sgGrp.GroupName)
            $_sgGrpRef  = if ($_sgGrp.GroupReferenceURL) {
                " &nbsp;<a href='$([System.Net.WebUtility]::HtmlEncode($_sgGrp.GroupReferenceURL))' target='_blank' style='font-weight:normal;font-size:0.75rem;color:#2563eb'>baseline &#8599;</a>"
            } else { '' }
            $_sgCtrlRows.Add("<tr class='scuba-group-hdr' style='background:#edf0f7'><td colspan='5' style='font-weight:700;font-size:0.82rem;color:#1e3a5f;padding:6px 9px'>$_sgGrpName$_sgGrpRef</td></tr>")

            foreach ($_sgCtrl in $_sgGrp.Controls) {
                $_sgResult   = [string]$_sgCtrl.Result
                $_sgCritFull = [string]$_sgCtrl.Criticality
                $_sgCrit     = ($_sgCritFull -split '/')[0]
                $_sgCritSub  = if ($_sgCritFull -match '/(.+)') { $Matches[1] } else { '' }

                # Strip HTML from Details — ScubaGear embeds internal report anchors that are meaningless outside their own report
                $_sgDetails = [System.Text.RegularExpressions.Regex]::Replace([string]$_sgCtrl.Details, '<[^>]+>', ' ')
                $_sgDetails = ($($_sgDetails -replace '\s+', ' ')).Trim()

                # Row background tint by result + criticality
                $_sgRowBg = switch ($_sgResult) {
                    'Fail'    { if ($_sgCrit -eq 'Shall') { "background:#fff1f0" } else { "background:#fffbe6" } }
                    'Warning' { "background:#fffbe6" }
                    'Manual'  { "background:#f0f7ff" }
                    'N/A'     { "background:#f7f7f7" }
                    default   { "" }
                }

                # Result cell
                $_sgResultHtml = switch ($_sgResult) {
                    'Fail'    { "<span class='critical'>Fail</span>" }
                    'Warning' { "<span class='warn'>Warning</span>" }
                    'Pass'    { "<span class='ok'>Pass</span>" }
                    'Manual'  { "<span style='color:#1565c0;font-weight:bold'>Manual</span>" }
                    default   { "<span style='color:#999'>$([System.Net.WebUtility]::HtmlEncode($_sgResult))</span>" }
                }

                # Criticality cell — show sub-classification (Not-Implemented / 3rd Party) in smaller text
                $_sgCritHtml = [System.Net.WebUtility]::HtmlEncode($_sgCrit)
                if ($_sgCritSub) {
                    $_sgCritHtml += "<br><span style='font-size:0.72rem;color:#999'>$([System.Net.WebUtility]::HtmlEncode($_sgCritSub))</span>"
                }

                # Control ID — link to the group's baseline docs section
                $_sgCtrlId     = [System.Net.WebUtility]::HtmlEncode($_sgCtrl.'Control ID')
                $_sgCtrlIdHtml = if ($_sgGrp.GroupReferenceURL) {
                    "<a href='$([System.Net.WebUtility]::HtmlEncode($_sgGrp.GroupReferenceURL))' target='_blank' style='font-family:monospace;font-size:0.8rem;white-space:nowrap'>$_sgCtrlId</a>"
                } else {
                    "<code>$_sgCtrlId</code>"
                }

                $_sgReqEnc = [System.Net.WebUtility]::HtmlEncode($_sgCtrl.Requirement)
                $_sgDetEnc = [System.Net.WebUtility]::HtmlEncode($_sgDetails)

                $_sgCtrlRows.Add("<tr data-scuba-result='$_sgResult' style='$_sgRowBg'><td style='white-space:nowrap;vertical-align:top'>$_sgCtrlIdHtml</td><td style='vertical-align:top'>$_sgReqEnc</td><td style='white-space:nowrap;vertical-align:top'>$_sgCritHtml</td><td style='white-space:nowrap;vertical-align:top'>$_sgResultHtml</td><td style='font-size:0.8rem;color:#555;vertical-align:top'>$_sgDetEnc</td></tr>")
            }
        }

        $_sgSummary.Add("<h4 id='$_sgProdAnchorId'>$_sgProdLabel</h4>")
        $_sgSummary.Add("<table><thead><tr><th style='white-space:nowrap'>Control ID</th><th>Requirement</th><th>Criticality</th><th>Result</th><th>Details</th></tr></thead><tbody>$($_sgCtrlRows -join '')</tbody></table>")
    }

    $_sgMeta    = $_scubaResults.MetaData
    $_sgVersion = if ($_sgMeta.ToolVersion) { "ScubaGear $($_sgMeta.ToolVersion)" } else { "ScubaGear" }

    $html.Add((Add-Section -Title "ScubaGear CIS Baseline" -AnchorId 'scuba' -CsvFiles @() -SummaryHtml ($_sgSummary -join "`n") -ModuleVersion $_sgVersion -Collapsed))
}

# =========================================
# ===   Close and Write Report          ===
# =========================================
$html.Add(@"
    </div><!-- /content-area -->
  </main>
</div><!-- /layout -->
<script>
// ScubaGear result filter — chips toggle visibility of rows by result type
function scubaFilter(chip) {
    if (chip.dataset.scubaFilter === 'all') {
        document.querySelectorAll('div.stat-chip[data-scuba-filter]').forEach(function(c) { c.classList.remove('inactive'); });
        document.querySelectorAll('tr[data-scuba-result]').forEach(function(r) { r.style.display = ''; });
        document.querySelectorAll('tr.scuba-group-hdr').forEach(function(r) { r.style.display = ''; });
        return;
    }
    chip.classList.toggle('inactive');
    var hidden = Array.from(document.querySelectorAll('div.stat-chip[data-scuba-filter].inactive')).map(function(c) { return c.dataset.scubaFilter; });
    document.querySelectorAll('tr[data-scuba-result]').forEach(function(r) {
        r.style.display = hidden.indexOf(r.dataset.scubaResult) !== -1 ? 'none' : '';
    });
    document.querySelectorAll('tr.scuba-group-hdr').forEach(function(hdr) {
        var next = hdr.nextElementSibling;
        var anyVisible = false;
        while (next && next.tagName === 'TR') {
            if (next.classList.contains('scuba-group-hdr')) break;
            if (next.dataset.scubaResult && next.style.display !== 'none') { anyVisible = true; break; }
            next = next.nextElementSibling;
        }
        hdr.style.display = anyVisible ? '' : 'none';
    });
}
// Row expand (sign-in history / permissions)
function toggleSignIns(row) {
    var d = row.nextElementSibling;
    if (d && d.classList.contains('signin-detail')) {
        var hidden = (d.style.display === 'none' || d.style.display === '');
        d.style.display = hidden ? 'table-row' : 'none';
        row.classList.toggle('expanded', hidden);
    }
}
function togglePerms(row) {
    var d = row.nextElementSibling;
    if (d && d.classList.contains('signin-detail')) {
        var hidden = (d.style.display === 'none' || d.style.display === '');
        d.style.display = hidden ? 'table-row' : 'none';
        row.classList.toggle('expanded', hidden);
    }
}
// Module section collapse/expand
function toggleModule(hdr) {
    var body   = hdr.nextElementSibling;
    var toggle = hdr.querySelector('.module-toggle');
    var isOpen = body.style.display !== 'none';
    body.style.display = isOpen ? 'none' : '';
    if (toggle) toggle.classList.toggle('open', !isOpen);
}
// Make h4 subsections inside module-body collapsible
(function() {
  document.querySelectorAll('.module-body').forEach(function(body) {
    Array.from(body.querySelectorAll(':scope > h4')).forEach(function(h4) {
      var det = document.createElement('details');
      det.open = true;
      if (h4.id) { det.id = h4.id; h4.removeAttribute('id'); }
      var sum = document.createElement('summary');
      sum.innerHTML = h4.innerHTML;
      det.appendChild(sum);
      var next = h4.nextElementSibling;
      while (next && next.tagName !== 'H4') {
        var after = next.nextElementSibling;
        det.appendChild(next);
        next = after;
      }
      h4.replaceWith(det);
    });
  });
})();
// Layout mode: app-shell when standalone, flow when embedded in a CMS (e.g. Hudu)
(function() {
    var wrapper  = document.querySelector('.page-wrapper');
    var hdr      = document.querySelector('.sticky-header');
    var sb       = document.querySelector('.sidebar');
    // If our wrapper is a direct child of <body> we are standalone; otherwise we are embedded
    var embedded = !wrapper || wrapper.parentNode.nodeName !== 'BODY';

    if (embedded) {
        // Inside a CMS — remove fixed-height constraints and let the host page scroll
        if (wrapper) { wrapper.style.height = 'auto'; wrapper.style.overflowY = 'visible'; }
        if (hdr)     { hdr.style.position = 'static'; }
        if (sb)      { sb.style.position = 'static'; sb.style.height = 'auto'; }
    } else {
        // Standalone — apply app-shell: wrapper scrolls, sidebar tracks header bottom
        function sizeIt() {
            if (!hdr || !sb) return;
            var h = hdr.offsetHeight;
            sb.style.top    = h + 'px';
            sb.style.height = 'calc(100vh - ' + h + 'px)';
        }
        sizeIt();
        window.addEventListener('resize', sizeIt);
    }

    // Scroll-spy — precompute stable offsets so hidden elements (display:none collapsed
    // sections) never pollute the active calculation via zero-value getBoundingClientRect.
    var links = document.querySelectorAll('.sb-item[href^="#"], .sb-sub[href^="#"]');
    if (!links.length) return;
    var scroller    = embedded ? window : (wrapper || window);
    var triggerPx   = function() { return embedded ? 20 : (hdr ? hdr.offsetHeight + 20 : 20); };
    var scrollTop   = function() { return embedded ? (window.pageYOffset || document.documentElement.scrollTop) : (wrapper ? wrapper.scrollTop : 0); };

    // Build sorted array of { id, pos } where pos is the element's distance from the
    // top of the scroll container at the time the page finishes loading.
    // Elements inside display:none ancestors have a zero bounding rect — exclude them.
    function buildAnchors() {
        var base = scrollTop();
        var data = [];
        links.forEach(function(link) {
            var id = link.getAttribute('href').substring(1);
            var el = document.getElementById(id);
            if (!el) return;
            var r = el.getBoundingClientRect();
            if (r.width === 0 && r.height === 0) return; // hidden — skip
            data.push({ id: id, pos: r.top + base });
        });
        data.sort(function(a, b) { return a.pos - b.pos; });
        return data;
    }
    var anchors = buildAnchors();

    // Rebuild if user expands/collapses a module (positions shift)
    var origToggle = window.toggleModule;
    window.toggleModule = function(h) { if (origToggle) origToggle(h); anchors = buildAnchors(); applyActive(); };

    function applyActive() {
        var threshold = scrollTop() + triggerPx();
        var current   = '';
        for (var i = anchors.length - 1; i >= 0; i--) {
            if (anchors[i].pos <= threshold) { current = anchors[i].id; break; }
        }
        links.forEach(function(link) {
            link.classList.toggle('active', link.getAttribute('href').substring(1) === current);
        });
        document.querySelectorAll('.sb-module-group').forEach(function(group) {
            var parent = group.querySelector('.sb-item');
            var activeSub = group.querySelector('.sb-sub.active');
            if (parent && !parent.classList.contains('active')) {
                parent.classList.toggle('active', !!activeSub);
            }
        });
    }
    scroller.addEventListener('scroll', applyActive, { passive: true });
    applyActive();
})();
</script>
</div><!-- /page-wrapper -->
</body></html>
"@)
$html -join "`n" | Set-Content -Path $reportPath -Encoding UTF8

# Write structured action items sidecar for downstream integrations (e.g. Hudu publish)
if ($actionItems.Count -gt 0) {
    $sidecarPath = Join-Path $AuditFolder 'ActionItems.json'
    $actionItems | ForEach-Object { [PSCustomObject]$_ } | ConvertTo-Json -Depth 3 |
        Set-Content -Path $sidecarPath -Encoding UTF8
    Write-Verbose "Action items sidecar written: $sidecarPath"
}

# Write AuditMetrics.json — snapshot of key health metrics for month-over-month delta tracking
$_licTotalAssigned  = 0
$_licTotalAvailable = 0
$_licMetricsPath = Join-Path $entraDir 'Entra_Licenses.csv'
if (Test-Path $_licMetricsPath) {
    foreach ($_licRow in @(Import-Csv $_licMetricsPath)) {
        $consumed = 0; $enabled = 0
        [int]::TryParse($_licRow.ConsumedUnits, [ref]$consumed) | Out-Null
        [int]::TryParse($_licRow.EnabledUnits,  [ref]$enabled)  | Out-Null
        $_licTotalAssigned  += $consumed
        $_licTotalAvailable += [math]::Max(0, $enabled - $consumed)
    }
}

$_metricsObj = [ordered]@{
    RunDate                  = (Get-Date -Format 'yyyy-MM-dd')
    MfaCoveragePct           = if ($null -ne $_kpiMfaPct)       { [double]$_kpiMfaPct }       else { $null }
    MfaUserCount             = if ($null -ne $_kpiUserCount)    { [int]$_kpiUserCount }        else { $null }
    SecureScorePct           = if ($null -ne $_kpiScorePct)     { [double]$_kpiScorePct }      else { $null }
    SecureScoreCurrent       = if ($null -ne $_kpiSs)           { try { [double]$_kpiSs.CurrentScore } catch { $null } } else { $null }
    SecureScoreMax           = if ($null -ne $_kpiSs)           { try { [double]$_kpiSs.MaxScore }      catch { $null } } else { $null }
    ManagedDeviceCount       = if ($null -ne $_kpiDevCount)     { [int]$_kpiDevCount }         else { $null }
    NonCompliantDeviceCount  = if ($null -ne $_kpiNonCompliant) { [int]$_kpiNonCompliant }     else { $null }
    TenantStorageUsedGB      = if ($null -ne (Get-Variable '_kpiStoUsedGB'  -ErrorAction SilentlyContinue).Value) { [double]$_kpiStoUsedGB  } else { $null }
    TenantStorageTotalGB     = if ($null -ne (Get-Variable '_kpiStoTotalGB' -ErrorAction SilentlyContinue).Value) { [double]$_kpiStoTotalGB } else { $null }
    TenantStoragePct         = if ($null -ne (Get-Variable '_kpiStoPct'     -ErrorAction SilentlyContinue).Value) { [double]$_kpiStoPct     } else { $null }
    LicenseTotalAssigned     = $_licTotalAssigned
    LicenseTotalAvailable    = $_licTotalAvailable
    ActionItemCritical       = $_kpiCritCount
    ActionItemWarning        = $_kpiWarnCount
}
$_metricsPath = Join-Path $AuditFolder 'AuditMetrics.json'
$_metricsObj | ConvertTo-Json -Depth 2 | Set-Content -Path $_metricsPath -Encoding UTF8
Write-Verbose "Audit metrics written: $_metricsPath"

# Write Hudu-compatible inline-styled report (M365_HuduReport.html)
# No JavaScript. Inline styles throughout. Collapsible sections via <details>/<summary>.
$_huduReportPath = Join-Path $AuditFolder 'M365_HuduReport.html'
$_huduCompany    = if ($orgInfo -and $orgInfo.DisplayName) { [System.Net.WebUtility]::HtmlEncode($orgInfo.DisplayName) } else { 'Microsoft 365 Tenant' }
$_huduColour     = @{ ok = '#16a34a'; warn = '#d97706'; critical = '#dc2626' }

# ── Builder helpers ────────────────────────────────────────────────────────────

function New-HuduKpiTile { param([string]$Label, [string]$Value, [string]$Sub, [string]$Colour, [string]$DeltaMarkerId = '')
    $subHtml   = if ($Sub)            { "<div style='font-size:11px;color:#94a3b8;margin-top:3px;'>$Sub</div>" } else { '' }
    $deltaHtml = if ($DeltaMarkerId)  { "<!-- TILE_DELTA_$DeltaMarkerId -->" }                                   else { '' }
    return "<div style='flex:1;min-width:155px;background:rgba(128,128,128,0.05);border:1px solid rgba(128,128,128,0.2);border-radius:8px;padding:14px 16px;'>" +
           "<div style='font-size:11px;text-transform:uppercase;color:#64748b;letter-spacing:.05em;margin-bottom:4px;'>$Label</div>" +
           "<div style='font-size:22px;font-weight:700;color:$Colour;'>$Value</div>$subHtml$deltaHtml</div>"
}

function New-HuduSection { param([string]$Title, [string]$Content, [switch]$Open, [string]$Accent = '#1849a9')
    $openAttr = if ($Open) { ' open' } else { '' }
    return "<details$openAttr style='margin-bottom:10px;border:1px solid rgba(128,128,128,0.2);border-radius:8px;overflow:hidden;'>" +
           "<summary style='padding:11px 16px;background:$Accent;color:#fff;font-weight:600;font-size:13px;cursor:pointer;list-style:none;display:flex;align-items:center;justify-content:space-between;'>" +
           "<span>$Title</span><span style='font-size:10px;opacity:.65;'>&#9660;</span></summary>" +
           "<div style='padding:14px 16px;'>$Content</div></details>"
}

function New-HuduStatGrid { param([hashtable[]]$Stats)
    $tiles = foreach ($s in $Stats) {
        $c = if ($s.Colour) { $s.Colour } else { 'inherit' }
        "<div style='flex:1;min-width:110px;padding:10px 12px;background:rgba(128,128,128,0.05);border-radius:6px;border:1px solid rgba(128,128,128,0.2);'>" +
        "<div style='font-size:10px;text-transform:uppercase;color:#64748b;letter-spacing:.04em;'>$($s.Label)</div>" +
        "<div style='font-size:17px;font-weight:700;color:$c;margin-top:3px;'>$($s.Value)</div></div>"
    }
    return "<div style='display:flex;gap:8px;flex-wrap:wrap;margin-bottom:12px;'>$($tiles -join '')</div>"
}

function New-HuduTable { param([string[]]$Headers, [string[][]]$Rows, [int]$MaxRows = 0)
    if (-not $Rows -or $Rows.Count -eq 0) { return "<p style='font-size:12px;color:#94a3b8;margin:4px 0;'>No records found.</p>" }
    $show     = if ($MaxRows -gt 0 -and $Rows.Count -gt $MaxRows) { $Rows | Select-Object -First $MaxRows } else { $Rows }
    $truncMsg = if ($MaxRows -gt 0 -and $Rows.Count -gt $MaxRows) { "<p style='font-size:11px;color:#94a3b8;margin:6px 0 0;'>Showing $MaxRows of $($Rows.Count) — full list in M365_AuditSummary.html.</p>" } else { '' }
    $hCells   = $Headers | ForEach-Object { "<th style='padding:6px 10px;text-align:left;font-size:11px;text-transform:uppercase;color:#64748b;letter-spacing:.04em;background:rgba(128,128,128,0.05);white-space:nowrap;'>$_</th>" }
    $bRows    = $show | ForEach-Object { $cells = $_ | ForEach-Object { "<td style='padding:6px 10px;font-size:12px;border-top:1px solid rgba(128,128,128,0.1);'>$_</td>" }; "<tr>$($cells -join '')</tr>" }
    return "<table style='width:100%;border-collapse:collapse;border:1px solid rgba(128,128,128,0.2);border-radius:6px;overflow:hidden;margin-bottom:4px;'>" +
           "<thead><tr>$($hCells -join '')</tr></thead><tbody>$($bRows -join '')</tbody></table>$truncMsg"
}

function New-HuduAiTable { param([array]$Items, [string]$Heading, [string]$AccentColour)
    if (-not $Items -or $Items.Count -eq 0) { return '' }
    $rows = foreach ($ai in $Items) {
        $_cleanText = $ai.Text -replace '\s*\(CIS\s[\d.]+\)', ''
        "<tr style='border-bottom:1px solid rgba(128,128,128,0.1);'>" +
        "<td style='padding:8px 10px;font-size:12px;font-weight:600;white-space:nowrap;'>$($ai.Category)</td>" +
        "<td style='padding:8px 10px;font-size:13px;'>$_cleanText</td></tr>"
    }
    return "<div style='margin-bottom:12px;'>" +
           "<div style='padding:9px 14px;background:$AccentColour;border-radius:6px 6px 0 0;'>" +
           "<span style='color:#fff;font-weight:700;font-size:13px;'>$Heading ($($Items.Count))</span></div>" +
           "<div style='border:1px solid rgba(128,128,128,0.2);border-top:none;border-radius:0 0 6px 6px;overflow:hidden;'>" +
           "<table style='width:100%;border-collapse:collapse;'>" +
           "<thead><tr style='background:rgba(128,128,128,0.05);'>" +
           "<th style='padding:7px 10px;text-align:left;font-size:11px;text-transform:uppercase;color:#64748b;letter-spacing:.04em;width:160px;'>Category</th>" +
           "<th style='padding:7px 10px;text-align:left;font-size:11px;text-transform:uppercase;color:#64748b;letter-spacing:.04em;'>Finding</th>" +
           "</tr></thead>" +
           "<tbody>$($rows -join '')</tbody></table></div></div>"
}

function New-HuduModuleAi {
    # Compact action item sub-panel for a module section, filtered by category prefix
    param([string[]]$Prefixes)
    $_mItems = @($actionItems | Where-Object {
        $_cat = $_.Category; $Prefixes | Where-Object { $_cat -like "$_*" }
    } | Sort-Object { $_.Sequence })
    if ($_mItems.Count -eq 0) { return "<p style='font-size:12px;color:#16a34a;margin:10px 0 0;'>&#10003; No action items for this module.</p>" }
    $_mc = @($_mItems | Where-Object { $_.Severity -eq 'critical' })
    $_mw = @($_mItems | Where-Object { $_.Severity -eq 'warning'  })
    $rows = foreach ($ai in $_mItems) {
        $sev  = if ($ai.Severity -eq 'critical') { '#dc2626' } else { '#d97706' }
        $icon = if ($ai.Severity -eq 'critical') { '&#9889;' } else { '&#9888;' }
        $doc  = if ($ai.DocUrl) { "<td style='padding:5px 8px;white-space:nowrap;'><a href='$($ai.DocUrl)' style='color:#3b82f6;font-size:11px;'>Docs</a></td>" } else { '<td></td>' }
        "<tr style='border-top:1px solid rgba(128,128,128,0.1);'>" +
        "<td style='padding:5px 8px;white-space:nowrap;font-size:11px;font-weight:700;color:$sev;'>$icon $(($ai.Severity).ToUpper())</td>" +
        "<td style='padding:5px 8px;font-size:11px;white-space:nowrap;'>$($ai.Category)</td>" +
        "<td style='padding:5px 8px;font-size:12px;'>$($ai.Text)</td>$doc</tr>"
    }
    $badge = @()
    if ($_mc.Count -gt 0) { $badge += "<span style='color:#dc2626;font-weight:700;'>$($_mc.Count) critical</span>" }
    if ($_mw.Count -gt 0) { $badge += "<span style='color:#d97706;font-weight:700;'>$($_mw.Count) warning$(if ($_mw.Count -ne 1) {'s'})</span>" }
    return "<div style='margin-top:12px;border-top:2px solid rgba(128,128,128,0.1);padding-top:10px;'>" +
           "<div style='font-size:12px;font-weight:700;margin-bottom:6px;'>Action Items &mdash; $($badge -join ' &bull; ')</div>" +
           "<table style='width:100%;border-collapse:collapse;'><tbody>$($rows -join '')</tbody></table></div>"
}

# ── KPI row ────────────────────────────────────────────────────────────────────

# Tenant storage — read from SharePoint_TenantStorage.csv if available
$_kpiStorageStr   = '&mdash;'
$_kpiStorageSub   = $null
$_kpiStorageClass = 'ok'
$_h_tenantSto = try { Import-Csv (Join-Path $rawDir 'SharePoint_TenantStorage.csv') -ErrorAction Stop | Select-Object -First 1 } catch { $null }
if ($_h_tenantSto) {
    $_stoUsedMB  = [double]$_h_tenantSto.StorageUsedMB
    $_stoQuotaMB = [double]$_h_tenantSto.StorageQuotaMB
    if ($_stoUsedMB -le 0 -and $_stoQuotaMB -gt 0) {
        # Fallback: sum per-site + OneDrive CSVs if tenant-level value is zero
        $_h_fallbackSites = try { Import-Csv (Join-Path $rawDir 'SharePoint_Sites.csv')        -ErrorAction Stop } catch { @() }
        $_h_fallbackOd    = try { Import-Csv (Join-Path $rawDir 'SharePoint_OneDriveUsage.csv') -ErrorAction Stop } catch { @() }
        $_stoUsedMB  = (@($_h_fallbackSites) + @($_h_fallbackOd) | Measure-Object -Property StorageUsedMB -Sum).Sum
    }
    if ($_stoQuotaMB -gt 0) {
        $_stoUsedGB  = [math]::Round($_stoUsedMB  / 1024, 1)
        $_stoTotalGB = [math]::Round($_stoQuotaMB / 1024, 1)
        $_stoPct     = [math]::Round(($_stoUsedMB / $_stoQuotaMB) * 100, 0)
        $_kpiStorageStr   = "$_stoUsedGB / $_stoTotalGB GB"
        $_kpiStorageSub   = "$_stoPct% used"
        $_kpiStorageClass = if ($_stoPct -ge 90) { 'critical' } elseif ($_stoPct -ge 75) { 'warn' } else { 'ok' }
    }
}

$_huduKpiRow = "<div style='display:flex;gap:12px;flex-wrap:wrap;margin-bottom:20px;'>" +
    (New-HuduKpiTile 'MFA Coverage'      $_kpiMfaStr       $_kpiMfaSub       $_huduColour[$_kpiMfaClass]     -DeltaMarkerId 'MFA')     +
    (New-HuduKpiTile 'Secure Score'      $_kpiScoreVal     $_kpiScoreSub     $_huduColour[$_kpiScoreClass]   -DeltaMarkerId 'SCORE')   +
    (New-HuduKpiTile 'Managed Devices'   $_kpiDevStr       $_kpiDevSub       $_huduColour[$_kpiDevClass]     -DeltaMarkerId 'DEVICES') +
    (New-HuduKpiTile 'Tenant Storage'    $_kpiStorageStr   $_kpiStorageSub   $_huduColour[$_kpiStorageClass] -DeltaMarkerId 'STORAGE') +
    (New-HuduKpiTile 'Action Items'      $_kpiAiStr        $_kpiAiSub        $_huduColour[$_kpiAiClass]      -DeltaMarkerId 'AI')      +
    "</div>"

# ── Section: Action Items (ScubaGear excluded — same as main report) ───────────
$_huduCritItems = @($actionItems | Where-Object { $_.Severity -eq 'critical' -and $_.Category -notlike 'ScubaGear*' } |
    Sort-Object @{ Expression = { Get-ActionItemModuleSortOrder -Category $_.Category } }, @{ Expression = { $_.Sequence } })
$_huduWarnItems = @($actionItems | Where-Object { $_.Severity -eq 'warning'  -and $_.Category -notlike 'ScubaGear*' } |
    Sort-Object @{ Expression = { Get-ActionItemModuleSortOrder -Category $_.Category } }, @{ Expression = { $_.Sequence } })

$_huduAiContent = (New-HuduAiTable -Items $_huduCritItems -Heading '&#9889; Critical Issues' -AccentColour '#dc2626') +
                  (New-HuduAiTable -Items $_huduWarnItems -Heading '&#9888; Warnings'        -AccentColour '#d97706')
if (-not $_huduAiContent) {
    $_huduAiContent = "<p style='padding:14px;background:rgba(5,150,105,0.1);border:1px solid rgba(5,150,105,0.3);border-radius:6px;color:#15803d;font-size:13px;margin:0;'>&#10003; No action items — all checks passed.</p>"
}
$_huduAiParts = @()
if ($_huduCritItems.Count -gt 0) { $_huduAiParts += "$($_huduCritItems.Count) critical" }
if ($_huduWarnItems.Count -gt 0) { $_huduAiParts += "$($_huduWarnItems.Count) warning$(if ($_huduWarnItems.Count -ne 1) {'s'})" }
$_huduAiTitle = if ($_huduAiParts.Count -gt 0) { "Action Items &mdash; $($_huduAiParts -join ', ')" } else { "Action Items" }

$_secActionItems = New-HuduSection -Title $_huduAiTitle -Content $_huduAiContent -Open

# ── Section: Microsoft Entra ───────────────────────────────────────────────────
$_entraContent = ''
$_h_users = @(try { Import-Csv (Join-Path $rawDir 'Entra_Users.csv') -ErrorAction Stop } catch { @() })
if ($_h_users.Count -gt 0) {
    $_h_enabled  = @($_h_users | Where-Object { $_.AccountStatus -eq 'Enabled' })
    $_h_licensed = @($_h_users | Where-Object { $_.AssignedLicense -and $_.AssignedLicense -ne '' -and $_.AssignedLicense -ne 'None' -and $_.AccountStatus -eq 'Enabled' })
    $_entraContent += New-HuduStatGrid -Stats @(
        @{ Label = 'Enabled Users';   Value = $_h_enabled.Count;  Colour = 'inherit' }
        @{ Label = 'Licensed';        Value = $_h_licensed.Count; Colour = 'inherit' }
        @{ Label = 'MFA Coverage';    Value = $_kpiMfaStr;         Colour = $_huduColour[$_kpiMfaClass] }
        @{ Label = 'Secure Score';    Value = $_kpiScoreVal;       Colour = $_huduColour[$_kpiScoreClass] }
    )
}
$_h_admins = @(try { Import-Csv (Join-Path $rawDir 'Entra_GlobalAdmins.csv') -ErrorAction Stop } catch { @() })
$_h_ca     = @(try { Import-Csv (Join-Path $rawDir 'Entra_CA_Policies.csv')  -ErrorAction Stop } catch { @() })
$_h_sd     = @(try { Import-Csv (Join-Path $rawDir 'Entra_SecurityDefaults.csv') -ErrorAction Stop } catch { @() })
if ($_h_admins.Count -gt 0 -or $_h_ca.Count -gt 0 -or $_h_sd.Count -gt 0) {
    $_h_caEnabled  = @($_h_ca | Where-Object { $_.State -eq 'enabled' })
    $_h_caRO       = @($_h_ca | Where-Object { $_.State -eq 'enabledForReportingButNotEnforcing' })
    $_h_sdEnabled  = if ($_h_sd.Count -gt 0) { $_h_sd[0].SecurityDefaultsEnabled } else { '—' }
    $_adminColour  = if ($_h_admins.Count -gt 4 -or $_h_admins.Count -eq 0) { '#dc2626' } elseif ($_h_admins.Count -eq 1) { '#d97706' } else { '#16a34a' }
    $_entraContent += New-HuduStatGrid -Stats @(
        @{ Label = 'Global Admins';      Value = $_h_admins.Count;    Colour = $_adminColour }
        @{ Label = 'CA Policies';        Value = $_h_caEnabled.Count; Colour = if ($_h_caEnabled.Count -gt 0) { '#16a34a' } else { '#d97706' } }
        @{ Label = 'Report-Only CA';     Value = $_h_caRO.Count;      Colour = if ($_h_caRO.Count -gt 0) { '#d97706' } else { 'inherit' } }
        @{ Label = 'Security Defaults';  Value = $_h_sdEnabled;       Colour = if ($_h_sdEnabled -eq 'True') { '#16a34a' } elseif ($_h_sdEnabled -eq 'False') { '#d97706' } else { 'inherit' } }
    )
}
if ($_h_users.Count -gt 0) {
    $_thS  = "padding:6px 8px;text-align:left;font-size:11px;text-transform:uppercase;color:#64748b;letter-spacing:.04em;background:rgba(128,128,128,0.05);white-space:nowrap;"
    $_uRows = @($_h_users | Sort-Object UPN | Select-Object -First 30 | ForEach-Object {
        $_mfaStyle = if ($_.MFAEnabled -eq 'False') { 'background:rgba(220,38,38,0.1);color:#991b1b;font-weight:600;white-space:nowrap;' } else { 'color:#15803d;font-weight:600;white-space:nowrap;' }
        "<tr style='border-bottom:1px solid rgba(128,128,128,0.1);'>" +
        "<td style='padding:5px 8px;font-size:12px;'>$($_.UPN)</td>" +
        "<td style='padding:5px 8px;font-size:12px;white-space:nowrap;'>$("$($_.FirstName) $($_.LastName)".Trim())</td>" +
        "<td style='padding:5px 8px;font-size:12px;white-space:nowrap;'>$($_.AccountStatus)</td>" +
        "<td style='padding:5px 8px;font-size:12px;'>$($_.AssignedLicense)</td>" +
        "<td style='padding:5px 8px;font-size:12px;text-align:center;${_mfaStyle}'>$($_.MFAEnabled)</td>" +
        "<td style='padding:5px 8px;font-size:12px;white-space:nowrap;'>$($_.LastSignIn)</td>" +
        "</tr>"
    })
    $_truncNote = if ($_h_users.Count -gt 30) { "<p style='font-size:11px;color:#94a3b8;margin:4px 0 0;'>Showing 30 of $($_h_users.Count) — full list in M365_AuditSummary.html.</p>" } else { '' }
    $_entraContent += "<table style='width:100%;border-collapse:collapse;border:1px solid rgba(128,128,128,0.2);border-radius:6px;overflow:hidden;margin-bottom:4px;'>" +
        "<thead><tr><th style='$_thS'>UPN</th><th style='$_thS'>Name</th><th style='$_thS'>Status</th>" +
        "<th style='$_thS'>License</th><th style='$_thS'>MFA</th><th style='$_thS'>Last Sign-In</th>" +
        "</tr></thead><tbody>$($_uRows -join '')</tbody></table>$_truncNote"
}
$_secEntra = New-HuduSection -Title 'Microsoft Entra' -Content $_entraContent

# ── Section: Exchange ──────────────────────────────────────────────────────────
$_exchContent = ''
$_h_mbx = @(try { Import-Csv (Join-Path $rawDir 'Exchange_Mailboxes.csv') -ErrorAction Stop } catch { @() })
if ($_h_mbx.Count -gt 0) {
    $_h_userMbx     = @($_h_mbx | Where-Object { $_.RecipientType -eq 'UserMailbox' })
    $_h_sharedMbx   = @($_h_mbx | Where-Object { $_.RecipientType -eq 'SharedMailbox' })
    $_h_resourceMbx = @($_h_mbx | Where-Object { $_.RecipientType -notin @('UserMailbox','SharedMailbox') })
    $_exchContent += New-HuduStatGrid -Stats @(
        @{ Label = 'User Mailboxes';    Value = $_h_userMbx.Count;     Colour = 'inherit' }
        @{ Label = 'Shared Mailboxes';  Value = $_h_sharedMbx.Count;   Colour = 'inherit' }
        @{ Label = 'Resource / Other';  Value = $_h_resourceMbx.Count; Colour = 'inherit' }
    )
}
$_h_fwdRules = @(try { Import-Csv (Join-Path $rawDir 'Exchange_InboxForwardingRules.csv') -ErrorAction Stop } catch { @() })
if ($_h_fwdRules.Count -gt 0) {
    $_exchContent += "<p style='font-size:12px;color:#d97706;margin:0 0 10px;font-weight:600;'>&#9888; $($_h_fwdRules.Count) inbox forwarding rule(s) detected — see full report for details.</p>"
}
if ($_h_mbx.Count -gt 0) {
    $_thS = "padding:6px 8px;text-align:left;font-size:11px;text-transform:uppercase;color:#64748b;letter-spacing:.04em;background:rgba(128,128,128,0.05);white-space:nowrap;"
    $_mRows = @($_h_mbx | Sort-Object @{E={switch ($_.RecipientType){'UserMailbox'{0}'SharedMailbox'{1}default{2}}}}, DisplayName | Select-Object -First 30 | ForEach-Object {
        $_usedMB  = if ($_.TotalSizeMB  -and $_.TotalSizeMB  -ne '') { [double]$_.TotalSizeMB  } else { 0 }
        $_limitMB = if ($_.LimitMB      -and $_.LimitMB      -ne '') { [double]$_.LimitMB      } else { 0 }
        if ($_limitMB -gt 0) {
            $_pct    = [math]::Round(($_usedMB / $_limitMB) * 100, 0)
            $_barW   = [math]::Min($_pct, 100)
            $_barClr = if ($_pct -gt 95) { '#dc2626' } elseif ($_pct -gt 80) { '#d97706' } else { '#16a34a' }
            $_limitGB  = [math]::Round($_limitMB / 1024, 1)
            $_sizeCell = "<div style='display:flex;align-items:center;gap:5px;'><div style='background:#e2e8f0;border-radius:3px;width:50px;height:7px;flex-shrink:0;overflow:hidden;'><div style='background:$_barClr;width:${_barW}%;height:7px;'></div></div><span style='font-size:11px;color:#64748b;white-space:nowrap;'>$_pct% of ${_limitGB} GB</span></div>"
        } else {
            $_sizeCell = if ($_usedMB -gt 0) { "<span style='font-size:12px;color:#64748b;'>$([math]::Round($_usedMB)) MB</span>" } else { "<span style='color:#94a3b8;'>—</span>" }
        }
        $_archiveCell = if ($_.ArchiveEnabled -eq 'True') {
            "<span style='color:#15803d;font-weight:700;'>&#10003;</span>"
        } elseif ($_limitMB -gt 0 -and $_usedMB -gt 0 -and ($_usedMB / $_limitMB) -gt 0.75) {
            "<span style='color:#d97706;font-weight:600;' title='No archive, mailbox over 75% full'>&#9888; None</span>"
        } else {
            "<span style='color:#94a3b8;'>—</span>"
        }
        "<tr style='border-bottom:1px solid rgba(128,128,128,0.1);'>" +
        "<td style='padding:5px 8px;font-size:12px;'>$($_.DisplayName)</td>" +
        "<td style='padding:5px 8px;font-size:11px;color:#64748b;'>$($_.UserPrincipalName)</td>" +
        "<td style='padding:5px 8px;font-size:12px;'>$($_.RecipientType)</td>" +
        "<td style='padding:5px 8px;min-width:90px;'>$_sizeCell</td>" +
        "<td style='padding:5px 8px;text-align:center;'>$_archiveCell</td>" +
        "</tr>"
    })
    $_truncNote = if ($_h_mbx.Count -gt 30) { "<p style='font-size:11px;color:#94a3b8;margin:4px 0 0;'>Showing 30 of $($_h_mbx.Count) — full list in M365_AuditSummary.html.</p>" } else { '' }
    $_exchContent += "<table style='width:100%;border-collapse:collapse;border:1px solid #e2e8f0;border-radius:6px;overflow:hidden;margin-bottom:4px;'>" +
        "<thead><tr><th style='$_thS'>Display Name</th><th style='$_thS'>UPN</th><th style='$_thS'>Type</th>" +
        "<th style='$_thS'>Size</th><th style='$_thS'>Archive</th>" +
        "</tr></thead><tbody>$($_mRows -join '')</tbody></table>$_truncNote"
}
$_secExchange = New-HuduSection -Title 'Exchange' -Content $_exchContent

# ── Section: SharePoint ────────────────────────────────────────────────────────
$_spContent = ''
$_h_sites = @(try { Import-Csv (Join-Path $rawDir 'SharePoint_Sites.csv') -ErrorAction Stop } catch { @() })
if ($_h_sites.Count -gt 0) {
    $_h_hubSites = @($_h_sites | Where-Object { $_.IsHubSite -eq 'True' })
    $_h_spSizeMB = ($_h_sites | Measure-Object -Property StorageUsedMB -Sum -ErrorAction SilentlyContinue).Sum
    $_spSizeLabel = if ($_h_spSizeMB -ge 1024) { "$([math]::Round($_h_spSizeMB / 1024, 1)) GB" } else { "$([math]::Round($_h_spSizeMB)) MB" }
    $_spContent += New-HuduStatGrid -Stats @(
        @{ Label = 'Storage Used'; Value = $_spSizeLabel;     Colour = 'inherit' }
        @{ Label = 'Total Sites';  Value = $_h_sites.Count;   Colour = 'inherit' }
        @{ Label = 'Hub Sites';    Value = $_h_hubSites.Count; Colour = 'inherit' }
    )
}
if ($_h_sites.Count -gt 0) {
    $_spTmpl = @{
        'SITEPAGEPUBLISHING#0' = 'Communication'; 'GROUP#0' = 'Team (M365)'; 'STS#0' = 'Classic Team'
        'STS#3' = 'Team'; 'GLOBAL#0' = 'Root Site'; 'SPSPERS#0' = 'OneDrive'; 'EHS#1' = 'Team Site'
        'SPSMSITEHOST#0' = 'MySite Host'; 'APPCATALOG#0' = 'App Catalog'; 'SRCHCEN#0' = 'Search Center'
        'SRCHCENTERLITE#0' = 'Search Center'; 'EDISC#0' = 'eDiscovery'; 'TEAMCHANNEL#0' = 'Teams Channel'
        'TEAMCHANNEL#1' = 'Teams Channel'; 'PWA#0' = 'Project Web App'; 'RedirectSite#0' = 'Redirect'
        'BLANKINTERNET#0' = 'Publishing'; 'BLANKINTERNETCONTAINER#0' = 'Publishing Portal'
        'ENTERWIKI#0' = 'Enterprise Wiki'
    }
    $_thS = "padding:6px 8px;text-align:left;font-size:11px;text-transform:uppercase;color:#64748b;letter-spacing:.04em;background:rgba(128,128,128,0.05);white-space:nowrap;"
    $_sRows = @($_h_sites | Sort-Object Title | Select-Object -First 30 | ForEach-Object {
        $_sg      = if ($_.StorageUsedMB -and [double]$_.StorageUsedMB -gt 0) { "$([math]::Round([double]$_.StorageUsedMB/1024,2)) GB" } else { '—' }
        $_urlPath = $_.Url -replace '^https://[^/]+', ''
        $_tmLabel = if ($_spTmpl.ContainsKey($_.Template)) { $_spTmpl[$_.Template] } `
                    elseif ($_.Template -like 'TEAMCHANNEL*') { 'Teams Channel' } `
                    elseif ($_.Template -like 'SPSPERS*')     { 'OneDrive' } `
                    elseif ($_.Template -like 'GROUP*')       { 'Team (M365)' } `
                    else { $_.Template }
        "<tr style='border-bottom:1px solid rgba(128,128,128,0.1);'>" +
        "<td style='padding:5px 8px;font-size:12px;'>$($_.Title)</td>" +
        "<td style='padding:5px 8px;font-size:11px;color:#64748b;word-break:break-all;'>$_urlPath</td>" +
        "<td style='padding:5px 8px;font-size:12px;white-space:nowrap;'>$_tmLabel</td>" +
        "<td style='padding:5px 8px;font-size:12px;white-space:nowrap;'>$_sg</td>" +
        "<td style='padding:5px 8px;font-size:11px;color:#64748b;'>$($_.Owner)</td>" +
        "</tr>"
    })
    $_truncNote = if ($_h_sites.Count -gt 30) { "<p style='font-size:11px;color:#94a3b8;margin:4px 0 0;'>Showing 30 of $($_h_sites.Count) — full list in M365_AuditSummary.html.</p>" } else { '' }
    $_spContent += "<table style='width:100%;border-collapse:collapse;border:1px solid #e2e8f0;border-radius:6px;overflow:hidden;margin-bottom:4px;'>" +
        "<thead><tr><th style='$_thS'>Title</th><th style='$_thS'>URL</th><th style='$_thS'>Template</th>" +
        "<th style='$_thS'>Storage</th><th style='$_thS'>Owner</th>" +
        "</tr></thead><tbody>$($_sRows -join '')</tbody></table>$_truncNote"
}
$_secSharePoint = New-HuduSection -Title 'SharePoint' -Content $_spContent

# ── Section: Mail Security ─────────────────────────────────────────────────────
$_mailContent = ''
$_h_dmarc = @(try { Import-Csv (Join-Path $rawDir 'MailSec_DMARC.csv') -ErrorAction Stop } catch { @() })
$_h_spf   = @(try { Import-Csv (Join-Path $rawDir 'MailSec_SPF.csv')   -ErrorAction Stop } catch { @() })
$_h_dkim  = @(try { Import-Csv (Join-Path $rawDir 'MailSec_DKIM.csv')  -ErrorAction Stop } catch { @() })
if ($_h_dmarc.Count -gt 0 -or $_h_spf.Count -gt 0) {
    $_h_allDomains = @(($_h_dmarc | Select-Object -ExpandProperty Domain) + ($_h_spf | Select-Object -ExpandProperty Domain) |
        Where-Object { $_ -notlike '*.onmicrosoft.com' } | Sort-Object -Unique)
    if ($_h_allDomains.Count -gt 0) {
        $_mailRows = foreach ($_dom in $_h_allDomains) {
            $_d = $_h_dmarc | Where-Object { $_.Domain -eq $_dom } | Select-Object -First 1
            $_s = $_h_spf   | Where-Object { $_.Domain -eq $_dom } | Select-Object -First 1
            $_k = $_h_dkim  | Where-Object { $_.Domain -eq $_dom } | Select-Object -First 1
            $_dmarcRaw = if ($_d) { $_d.DMARC } else { '' }
            $_dmarcVal = if ($_dmarcRaw -and $_dmarcRaw -ne 'Not Found' -and $_dmarcRaw -match 'v=DMARC1') {
                if ($_dmarcRaw -match 'p=none') { "<span style='color:#d97706;font-weight:600;font-size:11px;'>$_dmarcRaw</span>" }
                else { "<span style='color:#15803d;font-weight:600;font-size:11px;'>$_dmarcRaw</span>" }
            } else { "<span style='color:#dc2626;font-weight:600;'>Not Found</span>" }
            $_spfRaw  = if ($_s) { $_s.SPF } else { '' }
            $_spfVal  = if ($_spfRaw -and $_spfRaw -ne 'Not Found') {
                "<span style='color:#15803d;font-weight:600;font-size:11px;'>$_spfRaw</span>"
            } else { "<span style='color:#dc2626;font-weight:600;'>Not Found</span>" }
            $_dkimVal = if ($_k) {
                switch ($_k.DKIMEnabled) {
                    'True'           { "<span style='color:#15803d;font-weight:600;'>Enabled</span>" }
                    'False'          { "<span style='color:#dc2626;font-weight:600;'>Not Enabled</span>" }
                    'Not Configured' { "<span style='color:#dc2626;font-weight:600;'>Not Configured</span>" }
                    default          { "<span style='color:#94a3b8;'>$($_k.DKIMEnabled)</span>" }
                }
            } else { "<span style='color:#dc2626;font-weight:600;'>Not Found</span>" }
            ,@($_dom, $_dmarcVal, $_spfVal, $_dkimVal)
        }
        $_mailContent += New-HuduTable -Headers @('Domain', 'DMARC', 'SPF', 'DKIM') -Rows $_mailRows
    }
}
$_secMailSec = New-HuduSection -Title 'Mail Security' -Content $_mailContent

# ── Section: Intune (only if data present) ────────────────────────────────────
$_secIntune = ''
if (Test-Path (Join-Path $rawDir 'Intune_Devices.csv')) {
    $_intuneContent = ''
    $_h_devs = @(try { Import-Csv (Join-Path $rawDir 'Intune_Devices.csv') -ErrorAction Stop } catch { @() })
    if ($_h_devs.Count -gt 0) {
        $_h_compliant    = @($_h_devs | Where-Object { $_.ComplianceState -eq 'compliant' })
        $_h_nonCompliant = @($_h_devs | Where-Object { $_.ComplianceState -eq 'noncompliant' })
        $_h_stale        = @($_h_devs | Where-Object { try { ([datetime]::Now - [datetime]$_.LastSyncDateTime).TotalDays -gt 30 } catch { $false } })
        $_intuneContent += New-HuduStatGrid -Stats @(
            @{ Label = 'Total Devices';   Value = $_h_devs.Count;         Colour = 'inherit' }
            @{ Label = 'Compliant';       Value = $_h_compliant.Count;    Colour = '#16a34a' }
            @{ Label = 'Non-Compliant';   Value = $_h_nonCompliant.Count; Colour = if ($_h_nonCompliant.Count -gt 0) { '#dc2626' } else { '#16a34a' } }
            @{ Label = 'Stale (>30 d)';   Value = $_h_stale.Count;        Colour = if ($_h_stale.Count -gt 0) { '#d97706' } else { 'inherit' } }
        )
        $_devRows = $_h_devs | Sort-Object DeviceName | ForEach-Object {
            $_syncDt = [datetime]::MinValue
            $_isStale = $_.LastSyncDateTime -and [datetime]::TryParse($_.LastSyncDateTime, [ref]$_syncDt) -and (([datetime]::UtcNow - $_syncDt).TotalDays -gt 30)
            $_rowBg = switch ($_.ComplianceState) {
                'compliant'    { 'background:rgba(5,150,105,0.1);' }
                'noncompliant' { 'background:rgba(220,38,38,0.1);' }
                default        { '' }
            }
            $_compColour = switch ($_.ComplianceState) {
                'compliant'    { 'color:#15803d;font-weight:600;' }
                'noncompliant' { 'color:#991b1b;font-weight:600;' }
                default        { 'color:#64748b;' }
            }
            $_syncStyle = if ($_isStale) { 'color:#991b1b;font-weight:600;' } else { '' }
            "<tr style='${_rowBg}border-bottom:1px solid rgba(128,128,128,0.1);'>" +
            "<td style='padding:5px 8px;font-size:12px;'>$($_.DeviceName)</td>" +
            "<td style='padding:5px 8px;font-size:12px;'>$($_.OS)</td>" +
            "<td style='padding:5px 8px;font-size:12px;'>$($_.OSVersion)</td>" +
            "<td style='padding:5px 8px;font-size:12px;'>$($_.OwnerType)</td>" +
            "<td style='padding:5px 8px;font-size:12px;${_compColour}'>$($_.ComplianceState)</td>" +
            "<td style='padding:5px 8px;font-size:12px;'>$($_.AssignedUser)</td>" +
            "<td style='padding:5px 8px;font-size:12px;${_syncStyle}'>$($_.LastSyncDateTime)</td>" +
            "</tr>"
        }
        $_thStyle = "padding:6px 8px;text-align:left;font-size:11px;text-transform:uppercase;color:#64748b;letter-spacing:.04em;background:rgba(128,128,128,0.05);white-space:nowrap;"
        $_intuneContent += "<table style='width:100%;border-collapse:collapse;border:1px solid #e2e8f0;border-radius:6px;overflow:hidden;margin-bottom:4px;'>" +
            "<thead><tr>" +
            "<th style='$_thStyle'>Device</th><th style='$_thStyle'>OS</th><th style='$_thStyle'>Version</th>" +
            "<th style='$_thStyle'>Owner</th><th style='$_thStyle'>Compliance</th>" +
            "<th style='$_thStyle'>Assigned User</th><th style='$_thStyle'>Last Sync</th>" +
            "</tr></thead><tbody>$($_devRows -join '')</tbody></table>"
    }
    $_secIntune = New-HuduSection -Title 'Intune / Endpoint Management' -Content $_intuneContent
}

# ── Section: Teams (only if data present) ────────────────────────────────────
$_secTeams = ''
$_h_teamsFed = @(try { Import-Csv (Join-Path $rawDir 'Teams_FederationConfig.csv') -ErrorAction Stop } catch { @() })
if ($_h_teamsFed.Count -gt 0) {
    $_teamsContent = ''
    $_tf = $_h_teamsFed[0]
    $_fedColour  = if ($_tf.AllowFederatedUsers -eq 'True') { '#d97706' } else { '#16a34a' }
    $_consColour = if ($_tf.AllowTeamsConsumer  -eq 'True') { '#d97706' } else { '#16a34a' }
    $_teamsContent += New-HuduStatGrid -Stats @(
        @{ Label = 'External Federation'; Value = $_tf.AllowFederatedUsers; Colour = $_fedColour  }
        @{ Label = 'Teams Consumer';      Value = $_tf.AllowTeamsConsumer;  Colour = $_consColour }
        @{ Label = 'Skype (Public)';      Value = $_tf.AllowPublicUsers;    Colour = if ($_tf.AllowPublicUsers -eq 'True') { '#d97706' } else { 'inherit' } }
    )
    $_teamsContent += New-HuduModuleAi -Prefixes @('Teams /')
    $_secTeams = New-HuduSection -Title 'Microsoft Teams' -Content $_teamsContent
}

# ── Section: ScubaGear (only if results loaded) ────────────────────────────────
$_secScuba = ''
if ($_scubaResults) {
    $_sgContent = ''
    $_sgAiItems = @($actionItems | Where-Object { $_.Category -like 'ScubaGear*' })
    $_sgRows    = foreach ($_sgP in $_scubaResults.Results.PSObject.Properties) {
        $_allCtrls = @($_sgP.Value | ForEach-Object { $_.Controls }) | Where-Object { $_ }
        $_sgPass   = @($_allCtrls | Where-Object { $_.Result -eq 'Pass'    }).Count
        $_sgFail   = @($_allCtrls | Where-Object { $_.Result -eq 'Fail'    }).Count
        $_sgWarn   = @($_allCtrls | Where-Object { $_.Result -eq 'Warning' }).Count
        $_sgNA     = @($_allCtrls | Where-Object { $_.Result -notin @('Pass','Fail','Warning') }).Count
        $_sgColour = if ($_sgFail -gt 0) { '#dc2626' } elseif ($_sgWarn -gt 0) { '#d97706' } else { '#16a34a' }
        ,@("<span style='font-weight:600;color:$_sgColour;'>$($_sgP.Name)</span>", $_sgPass, $_sgFail, $_sgWarn, $_sgNA)
    }
    if ($_sgRows.Count -gt 0) {
        $_sgContent += New-HuduTable -Headers @('Product', 'Pass', 'Fail', 'Warning', 'N/A') -Rows $_sgRows
    }
    $_sgContent += New-HuduModuleAi -Prefixes @('ScubaGear /')
    $_sgAiCrits = @($_sgAiItems | Where-Object { $_.Severity -eq 'critical' }).Count
    $_sgAiWarns = @($_sgAiItems | Where-Object { $_.Severity -eq 'warning'  }).Count
    $_sgParts   = @(); if ($_sgAiCrits -gt 0) { $_sgParts += "$_sgAiCrits critical" }; if ($_sgAiWarns -gt 0) { $_sgParts += "$_sgAiWarns warning$(if ($_sgAiWarns -ne 1){'s'})" }
    $_scubaTitle = if ($_sgParts.Count -gt 0) { "ScubaGear CIS Baseline &mdash; $($_sgParts -join ', ')" } else { "ScubaGear CIS Baseline" }
    $_secScuba = New-HuduSection -Title $_scubaTitle -Content $_sgContent -Accent '#334155'
}

# ── Assemble HTML fragment (rendered inside Hudu's rich-text field) ────────────
$_huduHtml = @"
<div style="font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;">

<div style="background:#1849a9;color:#fff;padding:18px 24px;border-radius:8px;margin-bottom:20px;display:flex;align-items:baseline;justify-content:space-between;flex-wrap:wrap;gap:8px;">
  <div>
    <div style="font-size:11px;text-transform:uppercase;letter-spacing:.08em;opacity:.7;margin-bottom:4px;">Microsoft 365 Audit Report</div>
    <div style="font-size:18px;font-weight:700;">$_huduCompany</div>
  </div>
  <div style="font-size:12px;opacity:.75;">Generated: $reportDate</div>
</div>

$_huduKpiRow
<!-- AUDIT_DELTA_INJECT -->
$_secActionItems
$_secEntra
$_secExchange
$_secSharePoint
$_secMailSec
$_secIntune
$_secTeams
$_secScuba

<div style="margin-top:20px;padding-top:12px;border-top:1px solid rgba(128,128,128,0.2);font-size:11px;color:#94a3b8;">
  Full detail: M365_AuditSummary.html &bull; 365Audit v$ScriptVersion
</div>

</div>
"@

$_huduHtml | Set-Content -Path $_huduReportPath -Encoding UTF8
Write-Verbose "Hudu report written: $_huduReportPath"

Write-Host "Summary report written to: $reportPath" -ForegroundColor Green
if (-not $NoOpen) {
    if ($IsLinux) {
        xdg-open $reportPath
    } elseif ($IsMacOS) {
        open $reportPath
    } else {
        Start-Process $reportPath
    }
}
