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

.NOTES
    Author      : Raymond Slater
    Version     : 1.32.0
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

$ScriptVersion = "1.32.0"
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
$rawDir      = Join-Path $AuditFolder "Raw Files"
$entraDir    = $rawDir
$exchangeDir = $rawDir
$spDir       = $rawDir
$mailSecDir  = $rawDir
$intuneDir   = $rawDir

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
        [string]   $ModuleVersion
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

    return @"
<section class='module' id='$AnchorId'>
  <div class='module-hdr' onclick='toggleModule(this)'>
    <span class='module-title'>$encodedTitle</span>$versionMarkup
    <span class='module-toggle open'>&#9658;</span>
  </div>
  <div class='module-body'>
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
body{font-family:'Segoe UI',system-ui,sans-serif;background:#f0f4f8;color:#1e293b;height:100vh;display:flex;flex-direction:column;overflow:hidden;font-size:14px;}
/* ── App header ── */
.app-header{background:linear-gradient(135deg,#0f2744 0%,#1d4ed8 100%);color:#fff;padding:0.6rem 1.25rem;display:flex;align-items:center;justify-content:space-between;flex-shrink:0;box-shadow:0 2px 8px rgba(0,0,0,0.25);}
.app-header h1{font-size:0.97rem;font-weight:700;letter-spacing:0.01em;}
.app-header-sub{font-size:0.72rem;opacity:0.7;margin-top:0.1rem;}
/* ── KPI strip ── */
.kpi-strip{background:#fff;border-bottom:1px solid #dde3ea;display:flex;flex-shrink:0;}
.kpi-card{flex:1;padding:0.55rem 0.9rem;border-right:1px solid #e8edf3;text-align:center;}
.kpi-card:last-child{border-right:none;}
.kpi-value{font-size:1.45rem;font-weight:800;line-height:1;}
.kpi-label{font-size:0.67rem;color:#64748b;margin-top:0.2rem;}
.kpi-sub{font-size:0.63rem;color:#94a3b8;margin-top:0.05rem;}
.kpi-value.ok{color:#16a34a;} .kpi-value.warn{color:#d97706;} .kpi-value.critical{color:#dc2626;}
/* ── Layout ── */
.layout{display:flex;flex:1;overflow:hidden;}
/* ── Sidebar ── */
.sidebar{width:208px;background:#1e293b;color:#94a3b8;display:flex;flex-direction:column;flex-shrink:0;overflow-y:auto;overflow-x:hidden;}
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
.main{flex:1;overflow-y:auto;scroll-behavior:smooth;}
.main::-webkit-scrollbar{width:6px;} .main::-webkit-scrollbar-thumb{background:#c8d3de;border-radius:3px;}
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
.ai-grid{display:grid;grid-template-columns:1fr 1fr;gap:0.75rem;margin-bottom:0.8rem;}
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
</style>
</head>
<body>
<div class='app-header'>
  <div>
    <h1>Microsoft 365 Audit Summary</h1>
    <div class='app-header-sub'>Generated: $reportDate &nbsp;|&nbsp; Folder: $(Split-Path $AuditFolder -Leaf)</div>
  </div>
</div>
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
  <h2>$($orgInfo.DisplayName)</h2>
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
}


# =========================================
# ===   Action Items                    ===
# =========================================
$actionItems = [System.Collections.Generic.List[hashtable]]::new()
$script:ActionItemSequence = 0

# Helper: add an action item
# Severity: 'critical' | 'warning'
function Add-ActionItem {
    param([string]$Severity, [string]$Category, [string]$Text, [string]$DocUrl = "")
    $script:ActionItemSequence++
    $script:actionItems.Add(@{
        Severity = $Severity
        Category = $Category
        Text     = $Text
        DocUrl   = $DocUrl
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
        default         { return 90 }
    }
}

# --- Audit certificate expiry ---
if ($CertExpiryDays -ge 0 -and $CertExpiryDays -le 30) {
    Add-ActionItem -Severity 'warning' -Category 'Toolkit / Certificate' -Text "Audit app certificate expires in $CertExpiryDays day(s). Run Setup-365AuditApp.ps1 -Force (requires interactive Global Admin login) to renew before the next audit run."
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
            -Text "Technical Contact address(es) are not from a recognised MSP domain: $_contactList — this may be a previous MSP's details still on the tenant. Review and update the Technical Notification email in the Microsoft 365 admin centre (Settings &rarr; Org settings &rarr; Organisation profile)."
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
            -DocUrl 'https://learn.microsoft.com/en-us/entra/identity/role-based-access-control/best-practices'
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
        Add-ActionItem -Severity 'critical' -Category 'Entra / MFA' -Text "MFA not enabled for $missing of $_aiTotal licensed users (${_aiPct}%). Essential Eight: Restrict privileged access — all users must have MFA." -DocUrl 'https://learn.microsoft.com/en-us/entra/identity/authentication/concept-mfa-howitworks'
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
        Add-ActionItem -Severity 'critical' -Category 'Entra / MFA' -Text "Security Defaults are disabled and no Conditional Access policies exist. MFA is not enforced for any user." -DocUrl 'https://learn.microsoft.com/en-us/entra/fundamentals/security-defaults'
    }
    elseif ($_aiEnabledCa.Count -eq 0) {
        $_reportOnly = @($_aiCaPolicies | Where-Object { $_.State -eq "enabledForReportingButNotEnforced" }).Count
        Add-ActionItem -Severity 'critical' -Category 'Entra / CA' -Text "Security Defaults disabled and no CA policies are in 'Enabled' state ($_reportOnly in report-only). MFA is not enforced." -DocUrl 'https://learn.microsoft.com/en-us/entra/identity/conditional-access/overview'
    }
    elseif (($_aiCaPolicies | Where-Object { $_.State -eq "enabledForReportingButNotEnforced" }).Count -gt 0) {
        $_roCount = ($_aiCaPolicies | Where-Object { $_.State -eq "enabledForReportingButNotEnforced" }).Count
        Add-ActionItem -Severity 'warning' -Category 'Entra / CA' -Text "$_roCount Conditional Access policy/policies are in report-only mode and not enforcing controls. Review and enable when ready." -DocUrl 'https://learn.microsoft.com/en-us/entra/identity/conditional-access/concept-conditional-access-report-only'
    }
}

# Global Admin count
$_aiGaCsv = Join-Path $entraDir "Entra_GlobalAdmins.csv"
if (Test-Path $_aiGaCsv) {
    $_aiGaCount = @(Import-Csv $_aiGaCsv).Count
    if ($_aiGaCount -eq 0) {
        Add-ActionItem -Severity 'critical' -Category 'Entra / Admins' -Text "No Global Administrators found — this may indicate a data collection issue."
    }
    elseif ($_aiGaCount -eq 1) {
        Add-ActionItem -Severity 'warning' -Category 'Entra / Admins' -Text "Only 1 Global Administrator account. Recommend at least 2 for resilience (break-glass scenario)." -DocUrl 'https://learn.microsoft.com/en-us/entra/identity/role-based-access-control/best-practices'
    }
    elseif ($_aiGaCount -gt 4) {
        Add-ActionItem -Severity 'warning' -Category 'Entra / Admins' -Text "$_aiGaCount Global Administrator accounts. Microsoft recommends 2–4 max. Essential Eight: Restrict administrative privileges." -DocUrl 'https://learn.microsoft.com/en-us/entra/identity/role-based-access-control/best-practices'
    }
}

# SSPR
$_aiSsprCsv = Join-Path $entraDir "Entra_SSPR.csv"
if (Test-Path $_aiSsprCsv) {
    $_aiSspr = Import-Csv $_aiSsprCsv | Select-Object -First 1
    if ($_aiSspr.SSPREnabled -ne "Enabled") {
        Add-ActionItem -Severity 'warning' -Category 'Entra / SSPR' -Text "Self-Service Password Reset is not fully enabled (current: $($_aiSspr.SSPREnabled)). Users cannot reset passwords without helpdesk intervention." -DocUrl 'https://learn.microsoft.com/en-us/entra/identity/authentication/concept-sspr-howitworks'
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
            -DocUrl 'https://learn.microsoft.com/en-us/entra/identity/enterprise-apps/manage-consent-requests'
    }

    # AvePoint SaaS backup detection — AvePoint registers service principals in the tenant when configured
    $_avePoint = @($_aiApps | Where-Object { $_.DisplayName -like 'AvePoint*' })
    if ($_avePoint.Count -eq 0) {
        Add-ActionItem -Severity 'warning' -Category 'Entra / Backup' `
            -Text "No AvePoint service principal detected in this tenant. Confirm that SaaS backup (Microsoft 365 data) is configured and active for this customer." `
            -DocUrl 'https://partner.avepointonlineservices.com/Dashboard#/directory'
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
            -DocUrl 'https://learn.microsoft.com/en-us/entra/id-protection/howto-identity-protection-investigate-risk'
    }
}

# --- Exchange checks ---

# Inbox forwarding rules
$_aiInboxCsv = Join-Path $exchangeDir "Exchange_InboxForwardingRules.csv"
if (Test-Path $_aiInboxCsv) {
    $_aiInboxRules = @(Import-Csv $_aiInboxCsv)
    if ($_aiInboxRules.Count -gt 0) {
        Add-ActionItem -Severity 'warning' -Category 'Exchange / Rules' -Text "$($_aiInboxRules.Count) inbox rule(s) forward or redirect mail. Review to ensure these are authorised and not a sign of account compromise." -DocUrl 'https://learn.microsoft.com/en-us/microsoft-365/security/office-365-security/mail-flow-rules-transport-rules-0'
    }
}

# Broken inbox rules
$_aiBrokenCsv = Join-Path $exchangeDir "Exchange_BrokenInboxRules.csv"
if (Test-Path $_aiBrokenCsv) {
    $_aiBrokenRules = @(Import-Csv $_aiBrokenCsv)
    if ($_aiBrokenRules.Count -gt 0) {
        Add-ActionItem -Severity 'warning' -Category 'Exchange / Rules' -Text "$($_aiBrokenRules.Count) inbox rule(s) are in a broken/non-functional state and are not processing mail. Edit or re-create them in Outlook." -DocUrl 'https://support.microsoft.com/en-us/office/manage-email-messages-by-using-rules-c24f5dea-9465-4df4-ad17-a50704d66c59'
    }
}

# Remote domain auto-forwarding — only flag named (non-wildcard) domains; the default * entry is present in every tenant
$_aiRemoteCsv = Join-Path $exchangeDir "Exchange_RemoteDomainForwarding.csv"
if (Test-Path $_aiRemoteCsv) {
    $_aiRemoteNamed = @(Import-Csv $_aiRemoteCsv | Where-Object { $_.AutoForwardEnabled -eq "True" -and $_.DomainName -ne "*" })
    if ($_aiRemoteNamed.Count -gt 0) {
        $domainList = ($_aiRemoteNamed.DomainName -join ", ")
        Add-ActionItem -Severity 'warning' -Category 'Exchange / Forwarding' -Text "Auto-forwarding explicitly enabled for named external domain(s): $domainList. Confirm these are intentional." -DocUrl 'https://learn.microsoft.com/en-us/exchange/mail-flow-best-practices/remote-domains/remote-domains'
    }
}

# Unified Audit Log / retention
$_aiAuditCfgCsv = Join-Path $exchangeDir "Exchange_AuditConfig.csv"
if (Test-Path $_aiAuditCfgCsv) {
    $_aiAuditCfg = Import-Csv $_aiAuditCfgCsv | Select-Object -First 1
    if ($_aiAuditCfg.UnifiedAuditLogIngestionEnabled -eq "False") {
        Add-ActionItem -Severity 'critical' -Category 'Exchange / Audit' -Text "Unified Audit Log ingestion is disabled. Security and compliance events are not being recorded. Enable immediately." -DocUrl 'https://learn.microsoft.com/en-us/purview/audit-log-enable-disable'
    }
    else {
        # Check retention
        $_aiRetDays = try { [int]([TimeSpan]::Parse($_aiAuditCfg.AuditLogAgeLimit).Days) } catch { 90 }
        if ($_aiRetDays -lt 90) {
            Add-ActionItem -Severity 'warning' -Category 'Exchange / Audit' -Text "Audit log retention is only $_aiRetDays days. Microsoft recommends 90+ days; Essential Eight recommends 12 months for privileged actions." -DocUrl 'https://learn.microsoft.com/en-us/purview/audit-log-enable-disable'
        }
    }
}

# Mailbox audit status
$_aiMbxAuditCsv = Join-Path $exchangeDir "Exchange_MailboxAuditStatus.csv"
if (Test-Path $_aiMbxAuditCsv) {
    $_aiMbxAudit   = @(Import-Csv $_aiMbxAuditCsv | Where-Object { $_.UserPrincipalName -notlike 'DiscoverySearchMailbox*' })
    $_aiAuditOff   = @($_aiMbxAudit | Where-Object { $_.AuditEnabled -eq "False" })
    if ($_aiAuditOff.Count -gt 0) {
        Add-ActionItem -Severity 'warning' -Category 'Exchange / Audit' -Text "$($_aiAuditOff.Count) mailbox(es) have per-mailbox auditing disabled. Actions in these mailboxes (logins, deletions, sends) will not be logged." -DocUrl 'https://learn.microsoft.com/en-us/purview/audit-mailboxes'
    }
}

# DKIM
$_aiDkimCsv = Join-Path $exchangeDir "Exchange_DKIM_Status.csv"
if (Test-Path $_aiDkimCsv) {
    $_aiDkim        = @(Import-Csv $_aiDkimCsv)
    $_aiDkimOff     = @($_aiDkim | Where-Object { $_.DKIMEnabled -ne "True" -and $_.Domain -notlike "*.onmicrosoft.com" })
    if ($_aiDkimOff.Count -gt 0) {
        $dkimDomains = ($_aiDkimOff.Domain -join ", ")
        Add-ActionItem -Severity 'warning' -Category 'Exchange / DKIM' -Text "DKIM signing not enabled on: $dkimDomains. DKIM helps prevent email spoofing and improves deliverability." -DocUrl 'https://learn.microsoft.com/en-us/microsoft-365/security/office-365-security/email-authentication-dkim-configure'
    }
}

# Anti-phish: Spoof Intelligence
$_aiPhishCsv = Join-Path $exchangeDir "Exchange_AntiPhishPolicies.csv"
if (Test-Path $_aiPhishCsv) {
    $_aiPhish      = @(Import-Csv $_aiPhishCsv)
    $_aiNoSpoof    = @($_aiPhish | Where-Object { $_.EnableSpoofIntelligence -eq "False" })
    if ($_aiNoSpoof.Count -gt 0) {
        Add-ActionItem -Severity 'warning' -Category 'Exchange / Anti-Phish' -Text "$($_aiNoSpoof.Count) anti-phishing policy/policies have Spoof Intelligence disabled. This reduces protection against email spoofing attacks." -DocUrl 'https://learn.microsoft.com/en-us/microsoft-365/security/office-365-security/anti-phishing-policies-about'
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
            -DocUrl 'https://learn.microsoft.com/en-us/exchange/mail-flow-best-practices/use-connectors-to-configure-mail-flow/set-up-connectors-to-route-mail'
    }
}

# --- Mail Security checks (MailSec module) ---

$_aiDmarcCsv = Join-Path $mailSecDir "MailSec_DMARC.csv"
if (Test-Path $_aiDmarcCsv) {
    $_aiDmarc    = @(Import-Csv $_aiDmarcCsv)
    $_aiNoDmarc  = @($_aiDmarc | Where-Object { $_.Domain -notlike "*.onmicrosoft.com" -and ($_.DMARC -eq "Not Found" -or $_.DMARC -eq "" -or $null -eq $_.DMARC) })
    if ($_aiNoDmarc.Count -gt 0) {
        $dmarcDomains = ($_aiNoDmarc.Domain -join ", ")
        Add-ActionItem -Severity 'critical' -Category 'Mail Security' -Text "DMARC not configured for: $dmarcDomains. Without DMARC, spoofed email from your domain cannot be detected or rejected by recipients." -DocUrl 'https://learn.microsoft.com/en-us/microsoft-365/security/office-365-security/email-authentication-dmarc-configure'
    }
}

$_aiSpfCsv = Join-Path $mailSecDir "MailSec_SPF.csv"
if (Test-Path $_aiSpfCsv) {
    $_aiSpf   = @(Import-Csv $_aiSpfCsv)
    $_aiNoSpf = @($_aiSpf | Where-Object { $_.Domain -notlike "*.onmicrosoft.com" -and ($_.SPF -eq "DNS query failed" -or $_.SPF -eq "" -or $null -eq $_.SPF) })
    if ($_aiNoSpf.Count -gt 0) {
        $spfDomains = ($_aiNoSpf.Domain -join ", ")
        Add-ActionItem -Severity 'critical' -Category 'Mail Security' -Text "SPF not configured for: $spfDomains. SPF is required to identify authorised sending servers and prevent spoofing." -DocUrl 'https://learn.microsoft.com/en-us/microsoft-365/security/office-365-security/email-authentication-spf-configure'
    }
}

# --- SharePoint checks ---

$_aiExtShareCsv = Join-Path $spDir "SharePoint_ExternalSharing_SiteOverrides.csv"
if (Test-Path $_aiExtShareCsv) {
    $_aiExtShare    = @(Import-Csv $_aiExtShareCsv)
    $_aiPermissive  = @($_aiExtShare | Where-Object { $_.SharingCapability -eq "ExternalUserAndGuestSharing" })
    if ($_aiPermissive.Count -gt 0) {
        Add-ActionItem -Severity 'warning' -Category 'SharePoint' -Text "$($_aiPermissive.Count) site(s) allow anonymous guest sharing, overriding tenant defaults. Review to confirm these are intentional." -DocUrl 'https://learn.microsoft.com/en-us/sharepoint/turn-external-sharing-on-or-off'
    }
}

$_aiOdUnlicCsv = Join-Path $spDir "SharePoint_OneDrive_Unlicensed.csv"
if (Test-Path $_aiOdUnlicCsv) {
    $_aiOdUnlic = @(Import-Csv $_aiOdUnlicCsv)
    if ($_aiOdUnlic.Count -gt 0) {
        Add-ActionItem -Severity 'warning' -Category 'SharePoint' -Text "$($_aiOdUnlic.Count) OneDrive account(s) belong to unlicensed users. Data may be inaccessible and storage costs may be wasted." -DocUrl 'https://learn.microsoft.com/en-us/sharepoint/manage-sites-in-new-admin-center'
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
        Add-ActionItem -Severity 'critical' -Category 'Entra / Auth' -Text "Legacy authentication does not appear to be blocked. Security Defaults is disabled and no enabled CA policy targets legacy auth client types with a Block control. Legacy auth bypasses MFA. Essential Eight ML2." -DocUrl 'https://learn.microsoft.com/en-us/entra/identity/conditional-access/block-legacy-authentication'
    }
}

# Stale licensed accounts (no sign-in for 90+ days)
if (Test-Path $_aiUsersCsv) {
    $_aiStale = @($_aiUsers | Where-Object {
        $dt = [datetime]::MinValue
        -not $_.LastSignIn -or (-not [datetime]::TryParse(($_.LastSignIn -replace ' UTC',''), [ref]$dt)) -or (([datetime]::UtcNow - $dt).TotalDays -gt 90)
    })
    if ($_aiStale.Count -gt 0) {
        Add-ActionItem -Severity 'warning' -Category 'Entra / Accounts' -Text "$($_aiStale.Count) licensed user(s) have not signed in for 90+ days or have no recorded sign-in. Review for stale/orphaned accounts." -DocUrl 'https://learn.microsoft.com/en-us/entra/identity/monitoring-health/recommendation-remove-unused-credential-from-apps'
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
        Add-ActionItem -Severity 'warning' -Category 'Entra / Guests' -Text "$($_aiStaleGuests.Count) guest account(s) have not signed in for 90+ days or have no recorded sign-in. Stale guests retain access to shared resources." -DocUrl 'https://learn.microsoft.com/en-us/entra/identity/users/manage-guest-access-with-access-reviews'
    }
}

# Shared mailbox sign-in enabled
$_aiSharedSignInCsv = Join-Path $exchangeDir "Exchange_SharedMailboxSignIn.csv"
if (Test-Path $_aiSharedSignInCsv) {
    $_aiSharedEnabled = @(Import-Csv $_aiSharedSignInCsv | Where-Object { $_.AccountDisabled -eq "False" })
    if ($_aiSharedEnabled.Count -gt 0) {
        Add-ActionItem -Severity 'warning' -Category 'Exchange / Mailboxes' -Text "$($_aiSharedEnabled.Count) shared mailbox(es) have interactive sign-in enabled. Shared mailboxes should have sign-in disabled to prevent direct login and MFA bypass." -DocUrl 'https://learn.microsoft.com/en-us/microsoft-365/admin/email/about-shared-mailboxes'
    }
}

# Outbound spam auto-forward policy
$_aiOutboundCsv = Join-Path $exchangeDir "Exchange_OutboundSpamAutoForward.csv"
if (Test-Path $_aiOutboundCsv) {
    $_aiOutboundOn = @(Import-Csv $_aiOutboundCsv | Where-Object { $_.AutoForwardingMode -eq "On" })
    if ($_aiOutboundOn.Count -gt 0) {
        Add-ActionItem -Severity 'critical' -Category 'Exchange / Forwarding' -Text "Outbound spam policy is set to always allow auto-forwarding (AutoForwardingMode = On). This permits unrestricted external mail forwarding and is a known data exfiltration vector." -DocUrl 'https://learn.microsoft.com/en-us/microsoft-365/security/office-365-security/outbound-spam-protection-about'
    }
}

# Safe Attachments
$_aiSafAttCsv = Join-Path $exchangeDir "Exchange_SafeAttachments.csv"
if (Test-Path $_aiSafAttCsv) {
    $_aiSafAtt = @(Import-Csv $_aiSafAttCsv)
    $_aiSafAttOn = @($_aiSafAtt | Where-Object { $_.Enable -eq "True" })
    if ($_aiSafAttOn.Count -eq 0) {
        Add-ActionItem -Severity 'warning' -Category 'Exchange / Defender' -Text "No Safe Attachments policy is enabled. Attachments are not being detonated/scanned before delivery. Requires Defender for Office 365 P1." -DocUrl 'https://learn.microsoft.com/en-us/microsoft-365/security/office-365-security/safe-attachments-about'
    }
}

# Safe Links
$_aiSafLnkCsv = Join-Path $exchangeDir "Exchange_SafeLinks.csv"
if (Test-Path $_aiSafLnkCsv) {
    $_aiSafLnk = @(Import-Csv $_aiSafLnkCsv)
    $_aiSafLnkOn = @($_aiSafLnk | Where-Object { $_.EnableSafeLinksForEmail -eq "True" })
    if ($_aiSafLnkOn.Count -eq 0) {
        Add-ActionItem -Severity 'warning' -Category 'Exchange / Defender' -Text "No Safe Links policy is enabled for email. URLs are not being rewritten or checked at time-of-click. Requires Defender for Office 365 P1." -DocUrl 'https://learn.microsoft.com/en-us/microsoft-365/security/office-365-security/safe-links-about'
    }
}

# SharePoint default sharing link type
$_aiSpTenantCsv = Join-Path $spDir "SharePoint_ExternalSharing_Tenant.csv"
if (Test-Path $_aiSpTenantCsv) {
    $_aiSpTenant = Import-Csv $_aiSpTenantCsv | Select-Object -First 1
    if ($_aiSpTenant.DefaultSharingLinkType -eq "AnonymousAccess") {
        Add-ActionItem -Severity 'warning' -Category 'SharePoint' -Text "Default sharing link type is set to 'Anyone' (anonymous). Every share defaults to a link accessible by anyone with the URL, with no sign-in required." -DocUrl 'https://learn.microsoft.com/en-us/sharepoint/change-default-sharing-link'
    }
}

# SharePoint sync restriction
$_aiSpAcpCsv = Join-Path $spDir "SharePoint_AccessControlPolicies.csv"
if (Test-Path $_aiSpAcpCsv) {
    $_aiSpAcp = Import-Csv $_aiSpAcpCsv | Select-Object -First 1
    if ($_aiSpAcp.IsUnmanagedSyncAppForTenantRestricted -eq "False") {
        Add-ActionItem -Severity 'warning' -Category 'SharePoint' -Text "OneDrive sync is not restricted to managed/domain-joined devices. Any personal device can sync corporate data to local storage." -DocUrl 'https://learn.microsoft.com/en-us/sharepoint/control-access-from-unmanaged-devices'
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
                    -DocUrl 'https://learn.microsoft.com/en-us/mem/intune/protect/device-compliance-get-started'
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
                    -DocUrl 'https://learn.microsoft.com/en-us/mem/intune/protect/device-compliance-get-started'
            }
            if ($_aiStaleDevices.Count -gt 0) {
                Add-ActionItem -Severity 'warning' -Category 'Intune / Devices' `
                    -Text "$($_aiStaleDevices.Count) device(s) have not checked in with Intune for more than 30 days. Stale devices may not receive policy updates or be accurately reflected in compliance reports." `
                    -DocUrl 'https://learn.microsoft.com/en-us/mem/intune/remote-actions/devices-wipe'
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
                    -DocUrl 'https://learn.microsoft.com/en-us/mem/intune/protect/compliance-policy-create-windows'
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
                    -DocUrl 'https://learn.microsoft.com/en-us/mem/intune/protect/actions-for-noncompliance'
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
                    -DocUrl 'https://learn.microsoft.com/en-us/mem/intune/enrollment/enrollment-restrictions-set'
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
                    -DocUrl 'https://learn.microsoft.com/en-us/mem/intune/apps/troubleshoot-app-install'
            }
        }
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

# --- Build sidebar nav (status dots derived from action item categories) ---
$_sbModules = @(
    @{ Id = 'entra';       Label = 'Entra / Identity';      Prefix = 'Entra' },
    @{ Id = 'exchange';    Label = 'Exchange Online';        Prefix = 'Exchange' },
    @{ Id = 'sharepoint';  Label = 'SharePoint / OneDrive';  Prefix = 'SharePoint' },
    @{ Id = 'mailsec';     Label = 'Mail Security';          Prefix = 'Mail Security' },
    @{ Id = 'intune';      Label = 'Intune';                 Prefix = 'Intune' }
)
$_sbItemsHtml = foreach ($_mod in $_sbModules) {
    $_mc = @($actionItems | Where-Object { $_.Severity -eq 'critical' -and $_.Category -like "$($_mod.Prefix)*" }).Count
    $_mw = @($actionItems | Where-Object { $_.Severity -eq 'warning'  -and $_.Category -like "$($_mod.Prefix)*" }).Count
    $_dotClass   = if ($_mc -gt 0) { 'dot-critical' } elseif ($_mw -gt 0) { 'dot-warn' } else { 'dot-ok' }
    $_badgeHtml  = if ($_mc -gt 0) { "<span class='sb-badge'>$_mc</span>" } elseif ($_mw -gt 0) { "<span class='sb-badge warn'>$_mw</span>" } else { '' }
    "<a class='sb-item' href='#$($_mod.Id)'><span class='sb-dot $_dotClass'></span>$($_mod.Label)$_badgeHtml</a>"
}

# --- Emit KPI strip + layout wrapper + sidebar + main open ---
$_companyCard = if ($script:companyCardHtml) { $script:companyCardHtml } else { '' }

$html.Add(@"
<div class='kpi-strip'>
  <div class='kpi-card'><div class='kpi-value $_kpiMfaClass'>$_kpiMfaStr</div><div class='kpi-label'>MFA Coverage</div><div class='kpi-sub'>$_kpiMfaSub</div></div>
  <div class='kpi-card'><div class='kpi-value $_kpiScoreClass'>$_kpiScoreVal</div><div class='kpi-label'>Identity Secure Score</div><div class='kpi-sub'>$_kpiScoreSub</div></div>
  <div class='kpi-card'><div class='kpi-value $_kpiDevClass'>$_kpiDevStr</div><div class='kpi-label'>Managed Devices</div><div class='kpi-sub'>$_kpiDevSub</div></div>
  <div class='kpi-card'><div class='kpi-value $_kpiAiClass'>$_kpiAiStr</div><div class='kpi-label'>Action Items</div><div class='kpi-sub'>$_kpiAiSub</div></div>
</div>
<div class='layout'>
  <nav class='sidebar'>
    <div class='sb-section-label'>Modules</div>
    $($_sbItemsHtml -join "`n    ")
    <hr class='sb-divider'>
    <div class='sb-section-label'>Report</div>
    <a class='sb-item' href='$(([System.IO.Path]::GetRelativePath($script:ReportBaseDir, (Join-Path $AuditFolder "Raw Files")) -replace '\\', '/'))' target='_blank'><span class='sb-dot dot-neutral'></span>Raw CSV Files</a>
  </nav>
  <main class='main'>
    <div class='content-area'>
    $_companyCard
"@)

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

    $html.Add("<div class='ai-grid'>$_critPanel$_warnPanel</div>")
}
else {
    $html.Add("<p class='ai-none'>&#10003; No issues identified. All checked areas meet best-practice recommendations.</p>")
}


# =========================================
# ===   Entra Section                   ===
# =========================================
$entraFiles = @(Get-ChildItem "$entraDir\Entra_*.csv" -ErrorAction SilentlyContinue)

if ($entraFiles.Count -gt 0) {
    $entraSummary = [System.Collections.Generic.List[string]]::new()

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
            $ssBar   = "<div style='background:#e0e0e0;border-radius:4px;height:14px;width:100%;max-width:300px;overflow:hidden;display:inline-block;vertical-align:middle'><div style='background:$ssColor;width:$($ssPct)%;height:14px'></div></div>"
            $entraSummary.Add("<p><b>Identity Secure Score:</b> $($ss.CurrentScore) / $($ss.MaxScore) &nbsp;($($ss.Percentage)%) &nbsp;$ssBar &nbsp;<span style='color:#888;font-size:0.85em'>as of $($ss.Date)</span></p>")
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

    # --- Security Defaults ---
    $secDefaultsCsv = Join-Path $entraDir "Entra_SecurityDefaults.csv"
    if (Test-Path $secDefaultsCsv) {
        $secDef = Import-Csv $secDefaultsCsv | Select-Object -First 1
        if ($secDef.SecurityDefaultsEnabled -eq "True") {
            $entraSummary.Add("<p class='ok'>Security Defaults: <b>Enabled</b></p>")
        }
        else {
            $_sdCaCount = if ($null -ne $_aiEnabledCa) { @($_aiEnabledCa).Count } else { 0 }
            if ($_sdCaCount -gt 0) {
                $entraSummary.Add("<p class='ok'>Security Defaults: <b>Disabled</b> — $_sdCaCount Conditional Access polic$(if ($_sdCaCount -eq 1) { 'y' } else { 'ies' }) active</p>")
            }
            else {
                $entraSummary.Add("<p class='critical'>Security Defaults: <b>Disabled</b> — no Conditional Access policies are enabled</p>")
            }
        }
    }

    # --- SSPR ---
    $ssprCsv = Join-Path $entraDir "Entra_SSPR.csv"
    if (Test-Path $ssprCsv) {
        $ssprData = Import-Csv $ssprCsv | Select-Object -First 1
        if ($ssprData.SSPREnabled -eq "Enabled") {
            $entraSummary.Add("<p class='ok'>Self-Service Password Reset: <b>Enabled</b></p>")
        }
        else {
            $entraSummary.Add("<p class='critical'>Self-Service Password Reset: <b>$($ssprData.SSPREnabled)</b></p>")
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

            # Main user row — clickable to expand sign-in history
            $userRow = "<tr class='user-row' onclick='toggleSignIns(this)' title='Click to show/hide sign-in history'><td>$($user.UPN)</td><td>$($user.FirstName)</td><td>$($user.LastName)</td>$statusCell<td>$($user.AssignedLicense)</td>$mfaCell<td>$($user.MFAMethods)</td><td>$($user.DisablePasswordExpiration)</td><td>$($user.LastPasswordChange)</td><td>$($user.LastSignIn)</td></tr>"

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
<h4>Licence Summary</h4>
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
<h4>Admin Role Assignments ($memberCount user(s) across $roleCount role(s))</h4>
<table>
  <thead><tr><th>Role</th><th>User</th><th>UPN</th></tr></thead>
  <tbody>$($roleRows -join "`n")</tbody>
</table>
"@)
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

    # --- Stale Licensed Accounts ---
    if (Test-Path $entraUsersCsv) {
        $staleUsers = @(Import-Csv $entraUsersCsv | Where-Object {
            $dt = [datetime]::MinValue
            -not $_.LastSignIn -or (-not [datetime]::TryParse(($_.LastSignIn -replace ' UTC',''), [ref]$dt)) -or (([datetime]::UtcNow - $dt).TotalDays -gt 90)
        })
        if ($staleUsers.Count -gt 0) {
            $staleRows = foreach ($u in ($staleUsers | Sort-Object UPN)) {
                "<tr><td>$($u.UPN)</td><td>$($u.AssignedLicense)</td><td>$(if ($u.LastSignIn) { $u.LastSignIn } else { '<span class=''warn''>Never</span>' })</td></tr>"
            }
            $entraSummary.Add("<h4>Stale Licensed Accounts (no sign-in for 90+ days)</h4>")
            $entraSummary.Add("<p class='warn'>$($staleUsers.Count) licensed account(s) have not signed in for 90+ days or have no recorded sign-in. Review these accounts for deprovisioning.</p>")
            $entraSummary.Add(@"
<table>
  <thead><tr><th>UPN</th><th>License</th><th>Last Sign-In</th></tr></thead>
  <tbody>$($staleRows -join "`n")</tbody>
</table>
"@)
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
        $entraSummary.Add("<h4>Identity Protection</h4>")

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

            $entraSummary.Add("<h4>Groups ($($_grps.Count) total)</h4>")
            $entraSummary.Add("<p>Microsoft 365: <b>$_grpM365</b> &nbsp;|&nbsp; Security: <b>$_grpSec</b> &nbsp;|&nbsp; Dynamic: <b>$_grpDynamic</b> &nbsp;|&nbsp; Assigned: <b>$_grpAssigned</b>$(if ($_grpOnPrem -gt 0) { " &nbsp;|&nbsp; On-Prem Synced: <b>$_grpOnPrem</b>" })</p>")

            if ($_grpNoOwners.Count -gt 0) {
                $entraSummary.Add("<p class='warn'>$($_grpNoOwners.Count) group(s) have no owner assigned — these groups are unmanaged and may accumulate stale members.</p>")
            }
            if ($_grpRoleAssignable.Count -gt 0) {
                $_raNames = ($_grpRoleAssignable | ForEach-Object { ConvertTo-HtmlText $_.DisplayName }) -join ', '
                $entraSummary.Add("<p class='warn'>$($_grpRoleAssignable.Count) role-assignable group(s) — membership grants Entra directory roles: <b>$_raNames</b></p>")
            }

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
            $_eaAvePoint = @($_eaApps | Where-Object { $_.DisplayName -like 'AvePoint*' })
            $_eaRows = foreach ($_ea in ($_eaApps | Sort-Object DisplayName)) {
                $_eaIsAvePoint = $_ea.DisplayName -like 'AvePoint*'
                $_eaStyle      = if ($_eaIsAvePoint) { " style='background:#e8f5e9'" } else { "" }
                $_eaConsented  = if ($_ea.AdminConsented -eq 'True') { "<span style='color:#c62828;font-weight:bold'>Yes</span>" } else { "No" }
                $_eaEnabled    = if ($_ea.Enabled -eq 'True') { "Yes" } else { "<span style='color:#888'>No</span>" }
                $_eaRoles      = if ($_ea.ConsentedRoles -and [int]$_ea.ConsentedRoles -gt 0) { "<b>$($_ea.ConsentedRoles)</b>" } else { $_ea.ConsentedRoles }
                $publisher     = if ($_ea.PublisherName) { $(ConvertTo-HtmlText $_ea.PublisherName) } elseif ($_ea.PublisherDomain) { $(ConvertTo-HtmlText $_ea.PublisherDomain) } else { '<span style=''color:#888''>Unknown</span>' }
                "<tr$_eaStyle><td>$(ConvertTo-HtmlText $_ea.DisplayName)</td><td>$publisher</td><td>$_eaEnabled</td><td>$_eaConsented</td><td>$_eaRoles</td></tr>"
            }
            $_eaAveStatus = if ($_eaAvePoint.Count -gt 0) {
                "<p class='ok'>AvePoint detected — SaaS backup service principal is present in this tenant.</p>"
            } else {
                "<p class='critical'>AvePoint not detected — no AvePoint service principal found. Confirm SaaS backup is configured.</p>"
            }
            $entraSummary.Add("<h4>Enterprise Apps ($($_eaApps.Count) third-party)</h4>")
            $entraSummary.Add($_eaAveStatus)
            $entraSummary.Add(@"
<table>
  <thead><tr><th>App Name</th><th>Publisher</th><th>Enabled</th><th>Admin Consented</th><th>Consented Roles</th></tr></thead>
  <tbody>$($_eaRows -join "`n")</tbody>
</table>
"@)
        }
        else {
            $entraSummary.Add("<p class='ok'>No third-party enterprise apps found in this tenant.</p>")
        }
    }

    $html.Add((Add-Section -Title "Microsoft Entra" -AnchorId 'entra' -CsvFiles $entraFiles.FullName -SummaryHtml ($entraSummary -join "`n") -ModuleVersion (Get-ModuleScriptVersion -ScriptName 'Invoke-EntraAudit.ps1')))
}


# =========================================
# ===   Exchange Section                ===
# =========================================
$exchangeFiles = @(Get-ChildItem "$exchangeDir\Exchange_*.csv" -ErrorAction SilentlyContinue)

if ($exchangeFiles.Count -gt 0) {
    $exchangeSummary = [System.Collections.Generic.List[string]]::new()

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

            # Main mailbox row — clickable
            $mainRow = "<tr class='user-row' onclick='togglePerms(this)' title='Click to show/hide delegated permissions'><td>$($mbx.DisplayName)</td><td>$upn</td><td>$($mbx.RecipientType)</td>$usageCell<td>$($mbx.ArchiveEnabled)</td><td>$archiveSizeCell</td></tr>"

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
    }

    # --- Inbox Forwarding Rules ---
    if (Test-Path $forwardingCsv) {
        $forwarding = @(Import-Csv $forwardingCsv)
        if ($forwarding.Count -gt 0) {
            $fwdRows = foreach ($r in $forwarding) {
                "<tr><td>$($r.Mailbox)</td><td>$($r.RuleName)</td><td>$($r.ForwardTo)</td><td>$($r.RedirectTo)</td></tr>"
            }
            $exchangeSummary.Add("<h4>Inbox Forwarding Rules</h4>")
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
        $exchangeSummary.Add("<h4>Anti-Phish Policies</h4>")
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

    # --- Shared Mailbox Sign-In Status ---
    $sharedSignInCsv = Join-Path $exchangeDir "Exchange_SharedMailboxSignIn.csv"
    if (Test-Path $sharedSignInCsv) {
        $sharedMbx     = @(Import-Csv $sharedSignInCsv)
        $signInEnabled = @($sharedMbx | Where-Object { $_.AccountDisabled -eq "False" })
        $siClass       = if ($signInEnabled.Count -eq 0) { "ok" } else { "warn" }
        $exchangeSummary.Add("<h4>Shared Mailbox Sign-In</h4>")
        if ($signInEnabled.Count -eq 0) {
            $exchangeSummary.Add("<p class='ok'>All $($sharedMbx.Count) shared mailbox(es) have interactive sign-in disabled.</p>")
        }
        else {
            $exchangeSummary.Add("<p class='warn'>$($signInEnabled.Count) of $($sharedMbx.Count) shared mailbox(es) have interactive sign-in <b>enabled</b>. Shared mailboxes should have sign-in disabled to prevent direct login.</p>")
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

    $html.Add((Add-Section -Title "Exchange Online" -AnchorId 'exchange' -CsvFiles $exchangeFiles.FullName -SummaryHtml ($exchangeSummary -join "`n") -ModuleVersion (Get-ModuleScriptVersion -ScriptName 'Invoke-ExchangeAudit.ps1')))
}


# =========================================
# ===   SharePoint / OneDrive Section   ===
# =========================================
$spFiles = @(Get-ChildItem "$spDir\SharePoint_*.csv" -ErrorAction SilentlyContinue | Sort-Object Name)

if ($spFiles.Count -gt 0) {
    $spSummary = [System.Collections.Generic.List[string]]::new()

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
        $spSummary.Add("<h4>Tenant Storage</h4>$storageBar<p style='font-size:0.85rem;color:#666;margin-top:0'>OneDrive storage quota per user: $odQuotaGB GB</p>")
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

        $spSummary.Add("<h4>Site Collections ($($sites.Count))</h4>")
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
    $spSummary.Add("<h4>External Sharing</h4>")

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

        $spSummary.Add(@"
<table style='max-width:720px'>
  <thead><tr><th>Setting</th><th>Value</th></tr></thead>
  <tbody>
    <tr><td>Tenant Sharing Capability</td><td>$sharingLabel</td></tr>
    <tr><td>Default Sharing Link Type</td><td>$linkTypeLabel</td></tr>
    <tr><td>Anonymous Link Expiry</td><td>$anonExpiry</td></tr>
    <tr><td>Domain Restrictions</td><td>$domainRestrict</td></tr>
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
    $spSummary.Add("<h4>OneDrive</h4>")

    if (Test-Path $odUsageCsv) {
        $odDrives  = @(Import-Csv $odUsageCsv)
        $totalOdGB = [math]::Round(($odDrives | Measure-Object -Property StorageUsedMB -Sum).Sum / 1024, 1)
        $spSummary.Add("<p>$($odDrives.Count) OneDrive account(s) — $totalOdGB GB total in use</p>")
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
# ===   Mail Security Section           ===
# =========================================
$mailSecFiles = @(Get-ChildItem "$mailSecDir\MailSec_*.csv" -ErrorAction SilentlyContinue | Sort-Object Name)

if ($mailSecFiles.Count -gt 0) {
    $mailSecSummary = [System.Collections.Generic.List[string]]::new()

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
<h4 style='margin-top:1.2rem;margin-bottom:0.3rem'>$domain</h4>
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
# ===   Intune / Endpoint Section       ===
# =========================================
$intuneFiles = @(Get-ChildItem "$intuneDir\Intune_*.csv" -ErrorAction SilentlyContinue | Sort-Object Name)

if ($intuneFiles.Count -gt 0) {
    $intuneSummary = [System.Collections.Generic.List[string]]::new()

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

            $intuneSummary.Add("<h4 style='margin:1rem 0 0.25rem'>Device Inventory ($_totalDev total)</h4>")
            $intuneSummary.Add("<p>Corporate: $_corpDev &nbsp;|&nbsp; Personal (BYOD): $_persDev</p>")

            $intuneSummary.Add("<table class='summary-table'><thead><tr><th>Operating System</th><th>Count</th></tr></thead><tbody>")
            foreach ($_osGroup in $_osCounts) {
                $intuneSummary.Add("<tr><td>$(ConvertTo-HtmlText $_osGroup.Name)</td><td>$($_osGroup.Count)</td></tr>")
            }
            $intuneSummary.Add("</tbody></table>")

            $intuneSummary.Add("<h4 style='margin:1rem 0 0.25rem'>Compliance State Summary</h4>")
            $intuneSummary.Add("<table class='summary-table'><thead><tr><th>State</th><th>Count</th></tr></thead><tbody>")
            foreach ($_compGroup in $_compCounts) {
                $_stateColor = switch ($_compGroup.Name) {
                    'compliant'    { 'color:#388e3c' }
                    'noncompliant' { 'color:#c62828;font-weight:bold' }
                    default        { '' }
                }
                $intuneSummary.Add("<tr><td style='$_stateColor'>$(ConvertTo-HtmlText $_compGroup.Name)</td><td>$($_compGroup.Count)</td></tr>")
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

            $intuneSummary.Add("<h4 style='margin:1rem 0 0.25rem'>Managed Devices</h4>")
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
        }

        # Compliance policies
        if (Test-Path $intPolCsv) {
            $_intPols = @(Import-Csv $intPolCsv)
            $_intPolSettings = @()
            if (Test-Path $intPolSetCsv) { $_intPolSettings = @(Import-Csv $intPolSetCsv) }

            $intuneSummary.Add("<h4 style='margin:1rem 0 0.25rem'>Compliance Policies ($($_intPols.Count))</h4>")
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

            $intuneSummary.Add("<h4 style='margin:1rem 0 0.25rem'>Configuration Policies / Profiles ($($_intProfs.Count))</h4>")
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
            $intuneSummary.Add("<h4 style='margin:1rem 0 0.25rem'>Apps ($($_intApps.Count))</h4>")
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
            $intuneSummary.Add("<h4 style='margin:1rem 0 0.25rem'>Enrollment Restrictions ($($_intEnrol.Count))</h4>")
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
# ===   Close and Write Report          ===
# =========================================
$html.Add(@"
    </div><!-- /content-area -->
  </main>
</div><!-- /layout -->
<script>
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
// Sidebar scroll-spy
(function() {
    var main  = document.querySelector('.main');
    var links = document.querySelectorAll('.sb-item[href^="#"]');
    if (!main || !links.length) return;
    main.addEventListener('scroll', function() {
        var scrollTop = main.scrollTop + 80;
        var current   = '';
        links.forEach(function(link) {
            var id  = link.getAttribute('href').substring(1);
            var sec = document.getElementById(id);
            if (sec && sec.offsetTop <= scrollTop) current = id;
        });
        links.forEach(function(link) {
            var id = link.getAttribute('href').substring(1);
            link.classList.toggle('active', id === current);
        });
    }, { passive: true });
})();
</script>
</body></html>
"@)
$html -join "`n" | Set-Content -Path $reportPath -Encoding UTF8

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
