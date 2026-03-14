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

.NOTES
    Author      : Raymond Slater
    Version     : 1.19.0
    Change Log  : See CHANGELOG.md

.LINK
    https://github.com/razer86/365Audit
#>

#Requires -Version 7.2

param (
    [Parameter(Mandatory)]
    [string]$AuditFolder,
    [switch]$DevMode = $false
)

if (-not $DevMode -and $MyInvocation.InvocationName -eq $MyInvocation.MyCommand.Name) {
    Write-Error "This script must be run from the 365Audit launcher. Use -DevMode for development." -ErrorAction Stop
}

$ScriptVersion = "1.19.0"
Write-Verbose "Generate-AuditSummary.ps1 loaded (v$ScriptVersion)"

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

if (-not (Test-Path $AuditFolder)) {
    Write-Error "Provided audit folder does not exist: $AuditFolder"
    exit 1
}

# =========================================
# ===   HTML Section Helper             ===
# =========================================
function Add-Section {
    [CmdletBinding()]
    param (
        [string]   $Title,
        [string[]] $CsvFiles,
        [string]   $SummaryHtml
    )

    $csvLinks = ""
    if ($CsvFiles.Count -gt 0) {
        $fileItems = ""
        foreach ($file in $CsvFiles) {
            $name      = [System.IO.Path]::GetFileName($file)
            $fileItems += "<li><a href='$name' target='_blank'>$name</a></li>"
        }
        $csvLinks = "<details style='margin-top:1rem'><summary style='cursor:pointer;color:#555;font-size:0.85rem;list-style:disclosure-closed'>Raw CSV Files ($($CsvFiles.Count))</summary><ul style='margin:0.4rem 0 0 1rem;font-size:0.85rem'>$fileItems</ul></details>"
    }

    return @"
<details class='section'>
  <summary>$Title</summary>
  <div class='content'>
    $SummaryHtml
    $csvLinks
  </div>
</details>
"@
}

# =========================================
# ===   HTML Page Header                ===
# =========================================
$reportPath = Join-Path $AuditFolder "M365_AuditSummary.html"
$reportDate = Get-Date -Format "dd MMMM yyyy HH:mm"

$html = [System.Collections.Generic.List[string]]::new()
$html.Add(@"
<!DOCTYPE html>
<html lang='en'>
<head>
<meta charset='UTF-8'>
<title>Microsoft 365 Audit Summary</title>
<style>
  body        { font-family: Segoe UI, sans-serif; background: #f7f7f7; color: #333; margin: 2rem; }
  h1          { text-align: center; }
  .subtitle   { text-align: center; color: #666; margin-top: -0.5rem; margin-bottom: 2rem; }
  .section    { margin-bottom: 1.5rem; border: 1px solid #ccc; border-radius: 6px; background: #fff; box-shadow: 0 2px 5px rgba(0,0,0,0.05); }
  summary     { font-size: 1.1rem; font-weight: bold; padding: 1rem; cursor: pointer; background: #eaeaea; border-bottom: 1px solid #ccc; border-radius: 6px 6px 0 0; }
  .content    { padding: 1rem; overflow-x: auto; }
  table       { border-collapse: collapse; width: 100%; }
  th, td      { border: 1px solid #ccc; padding: 6px 10px; text-align: left; font-size: 0.9rem; }
  th          { background: #f0f0f0; }
  tr:nth-child(even) { background: #fafafa; }
  .ok         { color: green;      font-weight: bold; }
  .warn       { color: darkorange; font-weight: bold; }
  .critical   { color: red;        font-weight: bold; }
  .mfa-miss          { background-color: #ffcccc; }
  .user-row          { cursor: pointer; }
  .user-row:hover    { background-color: #e8f4fd !important; }
  .user-row.expanded { background-color: #cce4f7 !important; }
  .signin-detail     { background: #f0f7ff !important; }
  .signin-detail > td { padding: 0.5rem 1rem; }
  .inner-table       { width: 100%; border-collapse: collapse; font-size: 0.85rem; margin: 0; }
  .inner-table th    { background: #d0e8f8; }
  .inner-table td, .inner-table th { border: 1px solid #c0d8ec; padding: 4px 8px; }
  .signin-ok         { color: green;  font-weight: bold; }
  .signin-fail       { color: red;    font-weight: bold; }
  .size-warn         { background-color: #fff3cd; }
  .size-critical     { background-color: #ffcccc; }
  .company-info      { background: #fff; border: 1px solid #ccc; border-radius: 6px; padding: 1rem 1.5rem; margin: 0 auto 1.5rem; max-width: 700px; box-shadow: 0 2px 5px rgba(0,0,0,0.05); }
  .company-info h2   { margin: 0 0 0.75rem; font-size: 1.15rem; color: #333; }
  .company-info table { width: auto; min-width: 500px; }
  .company-info th   { background: transparent; font-weight: 600; padding: 4px 1.5rem 4px 0; width: 160px; border: none; border-bottom: 1px solid #eee; vertical-align: top; }
  .company-info td   { border: none; border-bottom: 1px solid #eee; padding: 4px 0; }
  .action-items      { background: #fff; border: 1px solid #ccc; border-radius: 6px; padding: 1rem 1.5rem; margin-bottom: 1.5rem; box-shadow: 0 2px 5px rgba(0,0,0,0.05); }
  .action-items h2   { margin: 0 0 0.75rem; font-size: 1.1rem; color: #333; }
  .action-item       { display: flex; align-items: flex-start; gap: 0.6rem; padding: 0.45rem 0; border-bottom: 1px solid #f0f0f0; }
  .action-item:last-child { border-bottom: none; }
  .action-badge      { display: inline-block; font-size: 0.7rem; font-weight: bold; padding: 2px 7px; border-radius: 3px; white-space: nowrap; flex-shrink: 0; margin-top: 2px; }
  .action-badge.critical { background: #ffcccc; color: #c00; border: 1px solid #f5a0a0; }
  .action-badge.warning  { background: #fff3cd; color: #805500; border: 1px solid #ffe082; }
  .action-cat        { font-weight: 600; white-space: nowrap; flex-shrink: 0; min-width: 110px; color: #555; font-size: 0.85rem; margin-top: 2px; }
  .action-text       { font-size: 0.9rem; color: #333; }
  .action-none       { color: #4caf50; font-weight: bold; font-size: 0.9rem; }
</style>
</head>
<body>
<h1>Microsoft 365 Audit Summary</h1>
<p class='subtitle'>Generated: $reportDate &nbsp;|&nbsp; Folder: $(Split-Path $AuditFolder -Leaf)</p>
"@)


# =========================================
# ===   Company Summary                 ===
# =========================================
$orgInfoPath = Join-Path $AuditFolder "OrgInfo.json"
if (Test-Path $orgInfoPath) {
    $orgInfo = Get-Content $orgInfoPath -Raw | ConvertFrom-Json

    # Address
    $addrParts = @($orgInfo.Raw.Street, $orgInfo.Raw.City, $orgInfo.Raw.State, $orgInfo.Raw.PostalCode, $orgInfo.CountryLetterCode)
    $address   = ($addrParts | Where-Object { $_ }) -join ", "

    # Phone and technical contact
    $phone       = if ($orgInfo.Raw.BusinessPhones) { ($orgInfo.Raw.BusinessPhones -join ", ") } else { $null }
    $techContact = if ($orgInfo.TechnicalNotificationMails.Count -gt 0) { $orgInfo.TechnicalNotificationMails -join ", " } else { "—" }

    # Azure AD Sync status (also satisfies the "Azure AD Sync Health" checklist item)
    $syncEnabled = $orgInfo.Raw.OnPremisesSyncEnabled
    $syncErrors  = if ($orgInfo.Raw.OnPremisesProvisioningErrors) { @($orgInfo.Raw.OnPremisesProvisioningErrors).Count } else { 0 }
    if ($syncEnabled) {
        $lastSyncDt  = [datetime]$orgInfo.Raw.OnPremisesLastSyncDateTime
        $hoursSince  = [math]::Round(([datetime]::UtcNow - $lastSyncDt.ToUniversalTime()).TotalHours, 1)
        $lastSyncFmt = $lastSyncDt.ToUniversalTime().ToString("yyyy-MM-dd HH:mm") + " UTC"
        $syncClass   = if ($hoursSince -gt 24) { "critical" } elseif ($hoursSince -gt 4) { "warn" } else { "ok" }
        $syncCell    = "<span class='$syncClass'>Enabled &mdash; last sync $lastSyncFmt ($hoursSince h ago)</span>"
        $errCell     = if ($syncErrors -gt 0) { "<span class='critical'>$syncErrors error(s)</span>" } else { "<span class='ok'>None</span>" }
        $syncRows    = "<tr><th>Azure AD Sync</th><td>$syncCell</td></tr><tr><th>Sync Errors</th><td>$errCell</td></tr>"
    } else {
        $syncRows = "<tr><th>Azure AD Sync</th><td>Not configured (cloud-only)</td></tr>"
    }

    # Verified domains — exclude the internal EOP routing domain (*.mail.onmicrosoft.com)
    $domains    = @($orgInfo.VerifiedDomains | Where-Object { $_.Name -notlike "*.mail.onmicrosoft.com" })
    $domainRows = foreach ($d in $domains) {
        $dtype = if ($d.IsInitial) { "*.onmicrosoft.com" } elseif ($d.IsDefault) { "Default" } else { "Custom" }
        $mark  = if ($d.IsDefault) { " <b>(default)</b>" } else { "" }
        "<tr><td>$($d.Name)$mark</td><td style='color:#666;font-size:0.85rem'>$dtype</td></tr>"
    }
    $domainsHtml = "<table class='inner-table' style='margin-top:2px'><tbody>$($domainRows -join '')</tbody></table>"

    $addrRow  = if ($address)  { "<tr><th>Address</th><td>$address</td></tr>" } else { "" }
    $phoneRow = if ($phone)    { "<tr><th>Phone</th><td>$phone</td></tr>" }    else { "" }

    $html.Add(@"
<div class='company-info'>
  <h2>$($orgInfo.DisplayName)</h2>
  <table>
    <tr><th>Tenant ID</th><td><code>$($orgInfo.Id)</code></td></tr>
    $addrRow
    $phoneRow
    <tr><th>Technical Contact</th><td>$techContact</td></tr>
    $syncRows
    <tr><th>Domains</th><td>$domainsHtml</td></tr>
  </table>
</div>
"@)
}


# =========================================
# ===   Action Items                    ===
# =========================================
$actionItems = [System.Collections.Generic.List[hashtable]]::new()

# Helper: add an action item
# Severity: 'critical' | 'warning'
function Add-ActionItem {
    param([string]$Severity, [string]$Category, [string]$Text, [string]$DocUrl = "")
    $script:actionItems.Add(@{ Severity = $Severity; Category = $Category; Text = $Text; DocUrl = $DocUrl })
}

# --- Entra checks ---

# MFA coverage
$_aiUsersCsv = Join-Path $AuditFolder "Entra_Users.csv"
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
$_aiSdCsv = Join-Path $AuditFolder "Entra_SecurityDefaults.csv"
$_aiCaCsv = Join-Path $AuditFolder "Entra_CA_Policies.csv"
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
$_aiGaCsv = Join-Path $AuditFolder "Entra_GlobalAdmins.csv"
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
$_aiSsprCsv = Join-Path $AuditFolder "Entra_SSPR.csv"
if (Test-Path $_aiSsprCsv) {
    $_aiSspr = Import-Csv $_aiSsprCsv | Select-Object -First 1
    if ($_aiSspr.SSPREnabled -ne "Enabled") {
        Add-ActionItem -Severity 'warning' -Category 'Entra / SSPR' -Text "Self-Service Password Reset is not fully enabled (current: $($_aiSspr.SSPREnabled)). Users cannot reset passwords without helpdesk intervention." -DocUrl 'https://learn.microsoft.com/en-us/entra/identity/authentication/concept-sspr-howitworks'
    }
}

# --- Exchange checks ---

# Inbox forwarding rules
$_aiInboxCsv = Join-Path $AuditFolder "Exchange_InboxForwardingRules.csv"
if (Test-Path $_aiInboxCsv) {
    $_aiInboxRules = @(Import-Csv $_aiInboxCsv)
    if ($_aiInboxRules.Count -gt 0) {
        Add-ActionItem -Severity 'warning' -Category 'Exchange / Rules' -Text "$($_aiInboxRules.Count) inbox rule(s) forward or redirect mail. Review to ensure these are authorised and not a sign of account compromise." -DocUrl 'https://learn.microsoft.com/en-us/microsoft-365/security/office-365-security/mail-flow-rules-transport-rules-0'
    }
}

# Broken inbox rules
$_aiBrokenCsv = Join-Path $AuditFolder "Exchange_BrokenInboxRules.csv"
if (Test-Path $_aiBrokenCsv) {
    $_aiBrokenRules = @(Import-Csv $_aiBrokenCsv)
    if ($_aiBrokenRules.Count -gt 0) {
        Add-ActionItem -Severity 'warning' -Category 'Exchange / Rules' -Text "$($_aiBrokenRules.Count) inbox rule(s) are in a broken/non-functional state and are not processing mail. Edit or re-create them in Outlook." -DocUrl 'https://support.microsoft.com/en-us/office/manage-email-messages-by-using-rules-c24f5dea-9465-4df4-ad17-a50704d66c59'
    }
}

# Remote domain auto-forwarding — only flag named (non-wildcard) domains; the default * entry is present in every tenant
$_aiRemoteCsv = Join-Path $AuditFolder "Exchange_RemoteDomainForwarding.csv"
if (Test-Path $_aiRemoteCsv) {
    $_aiRemoteNamed = @(Import-Csv $_aiRemoteCsv | Where-Object { $_.AutoForwardEnabled -eq "True" -and $_.DomainName -ne "*" })
    if ($_aiRemoteNamed.Count -gt 0) {
        $domainList = ($_aiRemoteNamed.DomainName -join ", ")
        Add-ActionItem -Severity 'warning' -Category 'Exchange / Forwarding' -Text "Auto-forwarding explicitly enabled for named external domain(s): $domainList. Confirm these are intentional." -DocUrl 'https://learn.microsoft.com/en-us/exchange/mail-flow-best-practices/remote-domains/remote-domains'
    }
}

# Unified Audit Log / retention
$_aiAuditCfgCsv = Join-Path $AuditFolder "Exchange_AuditConfig.csv"
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
$_aiMbxAuditCsv = Join-Path $AuditFolder "Exchange_MailboxAuditStatus.csv"
if (Test-Path $_aiMbxAuditCsv) {
    $_aiMbxAudit   = @(Import-Csv $_aiMbxAuditCsv | Where-Object { $_.UserPrincipalName -notlike 'DiscoverySearchMailbox*' })
    $_aiAuditOff   = @($_aiMbxAudit | Where-Object { $_.AuditEnabled -eq "False" })
    if ($_aiAuditOff.Count -gt 0) {
        Add-ActionItem -Severity 'warning' -Category 'Exchange / Audit' -Text "$($_aiAuditOff.Count) mailbox(es) have per-mailbox auditing disabled. Actions in these mailboxes (logins, deletions, sends) will not be logged." -DocUrl 'https://learn.microsoft.com/en-us/purview/audit-mailboxes'
    }
}

# DKIM
$_aiDkimCsv = Join-Path $AuditFolder "Exchange_DKIM_Status.csv"
if (Test-Path $_aiDkimCsv) {
    $_aiDkim        = @(Import-Csv $_aiDkimCsv)
    $_aiDkimOff     = @($_aiDkim | Where-Object { $_.DKIMEnabled -ne "True" -and $_.Domain -notlike "*.onmicrosoft.com" })
    if ($_aiDkimOff.Count -gt 0) {
        $dkimDomains = ($_aiDkimOff.Domain -join ", ")
        Add-ActionItem -Severity 'warning' -Category 'Exchange / DKIM' -Text "DKIM signing not enabled on: $dkimDomains. DKIM helps prevent email spoofing and improves deliverability." -DocUrl 'https://learn.microsoft.com/en-us/microsoft-365/security/office-365-security/email-authentication-dkim-configure'
    }
}

# Anti-phish: Spoof Intelligence
$_aiPhishCsv = Join-Path $AuditFolder "Exchange_AntiPhishPolicies.csv"
if (Test-Path $_aiPhishCsv) {
    $_aiPhish      = @(Import-Csv $_aiPhishCsv)
    $_aiNoSpoof    = @($_aiPhish | Where-Object { $_.EnableSpoofIntelligence -eq "False" })
    if ($_aiNoSpoof.Count -gt 0) {
        Add-ActionItem -Severity 'warning' -Category 'Exchange / Anti-Phish' -Text "$($_aiNoSpoof.Count) anti-phishing policy/policies have Spoof Intelligence disabled. This reduces protection against email spoofing attacks." -DocUrl 'https://learn.microsoft.com/en-us/microsoft-365/security/office-365-security/anti-phishing-policies-about'
    }
}

# --- Mail Security checks (MailSec module) ---

$_aiDmarcCsv = Join-Path $AuditFolder "MailSec_DMARC.csv"
if (Test-Path $_aiDmarcCsv) {
    $_aiDmarc    = @(Import-Csv $_aiDmarcCsv)
    $_aiNoDmarc  = @($_aiDmarc | Where-Object { $_.Domain -notlike "*.onmicrosoft.com" -and ($_.DMARC -eq "Not Found" -or $_.DMARC -eq "" -or $null -eq $_.DMARC) })
    if ($_aiNoDmarc.Count -gt 0) {
        $dmarcDomains = ($_aiNoDmarc.Domain -join ", ")
        Add-ActionItem -Severity 'critical' -Category 'Mail Security' -Text "DMARC not configured for: $dmarcDomains. Without DMARC, spoofed email from your domain cannot be detected or rejected by recipients." -DocUrl 'https://learn.microsoft.com/en-us/microsoft-365/security/office-365-security/email-authentication-dmarc-configure'
    }
}

$_aiSpfCsv = Join-Path $AuditFolder "MailSec_SPF.csv"
if (Test-Path $_aiSpfCsv) {
    $_aiSpf   = @(Import-Csv $_aiSpfCsv)
    $_aiNoSpf = @($_aiSpf | Where-Object { $_.Domain -notlike "*.onmicrosoft.com" -and ($_.SPF -eq "DNS query failed" -or $_.SPF -eq "" -or $null -eq $_.SPF) })
    if ($_aiNoSpf.Count -gt 0) {
        $spfDomains = ($_aiNoSpf.Domain -join ", ")
        Add-ActionItem -Severity 'critical' -Category 'Mail Security' -Text "SPF not configured for: $spfDomains. SPF is required to identify authorised sending servers and prevent spoofing." -DocUrl 'https://learn.microsoft.com/en-us/microsoft-365/security/office-365-security/email-authentication-spf-configure'
    }
}

# --- SharePoint checks ---

$_aiExtShareCsv = Join-Path $AuditFolder "SharePoint_ExternalSharing_SiteOverrides.csv"
if (Test-Path $_aiExtShareCsv) {
    $_aiExtShare    = @(Import-Csv $_aiExtShareCsv)
    $_aiPermissive  = @($_aiExtShare | Where-Object { $_.SharingCapability -eq "ExternalUserAndGuestSharing" })
    if ($_aiPermissive.Count -gt 0) {
        Add-ActionItem -Severity 'warning' -Category 'SharePoint' -Text "$($_aiPermissive.Count) site(s) allow anonymous guest sharing, overriding tenant defaults. Review to confirm these are intentional." -DocUrl 'https://learn.microsoft.com/en-us/sharepoint/turn-external-sharing-on-or-off'
    }
}

$_aiOdUnlicCsv = Join-Path $AuditFolder "SharePoint_OneDrive_Unlicensed.csv"
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
$_aiGuestCsv = Join-Path $AuditFolder "Entra_GuestUsers.csv"
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
$_aiSharedSignInCsv = Join-Path $AuditFolder "Exchange_SharedMailboxSignIn.csv"
if (Test-Path $_aiSharedSignInCsv) {
    $_aiSharedEnabled = @(Import-Csv $_aiSharedSignInCsv | Where-Object { $_.AccountDisabled -eq "False" })
    if ($_aiSharedEnabled.Count -gt 0) {
        Add-ActionItem -Severity 'warning' -Category 'Exchange / Mailboxes' -Text "$($_aiSharedEnabled.Count) shared mailbox(es) have interactive sign-in enabled. Shared mailboxes should have sign-in disabled to prevent direct login and MFA bypass." -DocUrl 'https://learn.microsoft.com/en-us/microsoft-365/admin/email/about-shared-mailboxes'
    }
}

# Outbound spam auto-forward policy
$_aiOutboundCsv = Join-Path $AuditFolder "Exchange_OutboundSpamAutoForward.csv"
if (Test-Path $_aiOutboundCsv) {
    $_aiOutboundOn = @(Import-Csv $_aiOutboundCsv | Where-Object { $_.AutoForwardingMode -eq "On" })
    if ($_aiOutboundOn.Count -gt 0) {
        Add-ActionItem -Severity 'critical' -Category 'Exchange / Forwarding' -Text "Outbound spam policy is set to always allow auto-forwarding (AutoForwardingMode = On). This permits unrestricted external mail forwarding and is a known data exfiltration vector." -DocUrl 'https://learn.microsoft.com/en-us/microsoft-365/security/office-365-security/outbound-spam-protection-about'
    }
}

# Safe Attachments
$_aiSafAttCsv = Join-Path $AuditFolder "Exchange_SafeAttachments.csv"
if (Test-Path $_aiSafAttCsv) {
    $_aiSafAtt = @(Import-Csv $_aiSafAttCsv)
    $_aiSafAttOn = @($_aiSafAtt | Where-Object { $_.Enable -eq "True" })
    if ($_aiSafAttOn.Count -eq 0) {
        Add-ActionItem -Severity 'warning' -Category 'Exchange / Defender' -Text "No Safe Attachments policy is enabled. Attachments are not being detonated/scanned before delivery. Requires Defender for Office 365 P1." -DocUrl 'https://learn.microsoft.com/en-us/microsoft-365/security/office-365-security/safe-attachments-about'
    }
}

# Safe Links
$_aiSafLnkCsv = Join-Path $AuditFolder "Exchange_SafeLinks.csv"
if (Test-Path $_aiSafLnkCsv) {
    $_aiSafLnk = @(Import-Csv $_aiSafLnkCsv)
    $_aiSafLnkOn = @($_aiSafLnk | Where-Object { $_.EnableSafeLinksForEmail -eq "True" })
    if ($_aiSafLnkOn.Count -eq 0) {
        Add-ActionItem -Severity 'warning' -Category 'Exchange / Defender' -Text "No Safe Links policy is enabled for email. URLs are not being rewritten or checked at time-of-click. Requires Defender for Office 365 P1." -DocUrl 'https://learn.microsoft.com/en-us/microsoft-365/security/office-365-security/safe-links-about'
    }
}

# SharePoint default sharing link type
$_aiSpTenantCsv = Join-Path $AuditFolder "SharePoint_ExternalSharing_Tenant.csv"
if (Test-Path $_aiSpTenantCsv) {
    $_aiSpTenant = Import-Csv $_aiSpTenantCsv | Select-Object -First 1
    if ($_aiSpTenant.DefaultSharingLinkType -eq "AnonymousAccess") {
        Add-ActionItem -Severity 'warning' -Category 'SharePoint' -Text "Default sharing link type is set to 'Anyone' (anonymous). Every share defaults to a link accessible by anyone with the URL, with no sign-in required." -DocUrl 'https://learn.microsoft.com/en-us/sharepoint/change-default-sharing-link'
    }
}

# SharePoint sync restriction
$_aiSpAcpCsv = Join-Path $AuditFolder "SharePoint_AccessControlPolicies.csv"
if (Test-Path $_aiSpAcpCsv) {
    $_aiSpAcp = Import-Csv $_aiSpAcpCsv | Select-Object -First 1
    if ($_aiSpAcp.IsUnmanagedSyncAppForTenantRestricted -eq "False") {
        Add-ActionItem -Severity 'warning' -Category 'SharePoint' -Text "OneDrive sync is not restricted to managed/domain-joined devices. Any personal device can sync corporate data to local storage." -DocUrl 'https://learn.microsoft.com/en-us/sharepoint/control-access-from-unmanaged-devices'
    }
}

# --- Render Action Items block ---
if ($actionItems.Count -gt 0) {
    $aiRows = foreach ($ai in $actionItems) {
        $badgeClass = $ai.Severity
        $badgeLabel = if ($ai.Severity -eq 'critical') { 'CRITICAL' } else { 'WARNING' }
        $docLink    = if ($ai.DocUrl) { " <a href='$($ai.DocUrl)' target='_blank' style='font-size:0.8rem;color:#1565c0;white-space:nowrap;text-decoration:none' title='Microsoft documentation'>&#128279; Docs</a>" } else { "" }
        "<div class='action-item'><span class='action-badge $badgeClass'>$badgeLabel</span><span class='action-cat'>$($ai.Category)</span><span class='action-text'>$($ai.Text)$docLink</span></div>"
    }
    $html.Add(@"
<div class='action-items'>
  <h2>&#9888; Action Items ($($actionItems.Count))</h2>
  $($aiRows -join "`n  ")
</div>
"@)
}
else {
    $html.Add(@"
<div class='action-items'>
  <h2>Action Items</h2>
  <p class='action-none'>&#10003; No issues identified. All checked areas meet best-practice recommendations.</p>
</div>
"@)
}


# =========================================
# ===   Entra Section                   ===
# =========================================
$entraFiles = @(Get-ChildItem "$AuditFolder\Entra_*.csv" -ErrorAction SilentlyContinue)

if ($entraFiles.Count -gt 0) {
    $entraSummary = [System.Collections.Generic.List[string]]::new()

    # Determine audit window from licence tier (drives all "last N days" labels)
    $_auditPremiumSkus = @("AAD_PREMIUM", "AAD_PREMIUM_P2", "ENTERPRISEPREMIUM", "ENTERPRISEPACK",
                           "EMS", "EMS_PREMIUM", "SPB", "O365_BUSINESS_PREMIUM", "M365_F3", "IDENTITY_GOVERNANCE")
    $auditWindowDays = 7
    $_licCheck = Join-Path $AuditFolder "Entra_Licenses.csv"
    if (Test-Path $_licCheck) {
        $_skuList = @(Import-Csv $_licCheck | Select-Object -ExpandProperty SkuPartNumber)
        if (($_skuList | Where-Object { $_ -in $_auditPremiumSkus }).Count -gt 0) { $auditWindowDays = 30 }
    }
    $auditWindowLabel = "last $auditWindowDays days"

    # --- Identity Secure Score ---
    $secureScoreCsv = Join-Path $AuditFolder "Entra_SecureScore.csv"
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
    $secureScoreControlsCsv = Join-Path $AuditFolder "Entra_SecureScoreControls.csv"
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
    $secDefaultsCsv = Join-Path $AuditFolder "Entra_SecurityDefaults.csv"
    if (Test-Path $secDefaultsCsv) {
        $secDef = Import-Csv $secDefaultsCsv | Select-Object -First 1
        if ($secDef.SecurityDefaultsEnabled -eq "True") {
            $entraSummary.Add("<p class='ok'>Security Defaults: <b>Enabled</b></p>")
        }
        else {
            $entraSummary.Add("<p class='warn'>Security Defaults: <b>Disabled</b> — ensure Conditional Access policies are in place</p>")
        }
    }

    # --- SSPR ---
    $ssprCsv = Join-Path $AuditFolder "Entra_SSPR.csv"
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
    $entraUsersCsv = Join-Path $AuditFolder "Entra_Users.csv"
    if (Test-Path $entraUsersCsv) {
        $userSummary = Import-Csv $entraUsersCsv
        $mfaTotal    = $userSummary.Count
        $mfaEnabled  = ($userSummary | Where-Object { $_.MFAEnabled -eq 'True' }).Count
        $mfaPercent  = if ($mfaTotal -gt 0) { [math]::Round(($mfaEnabled / $mfaTotal) * 100, 1) } else { 0 }
        $mfaClass    = if ($mfaPercent -eq 100) { "ok" } elseif ($mfaPercent -gt 0) { "warn" } else { "critical" }

        $entraSummary.Add("<p class='$mfaClass'>MFA enabled for <b>$mfaPercent%</b> of licensed users ($mfaEnabled / $mfaTotal)</p>")

        # Load sign-in history for expandable rows (keyed by UPN)
        $signInsByUpn = @{}
        $signInsCsv   = Join-Path $AuditFolder "Entra_SignIns.csv"
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
                "<td class='mfa-miss'>$($user.MFAEnabled)</td>"
            } else {
                "<td>$($user.MFAEnabled)</td>"
            }
            $statusCell = if ($user.AccountStatus -eq "Blocked") {
                "<td class='warn'>Blocked</td>"
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
<p style='font-size:0.85rem;color:#666;margin-bottom:0.5rem'>Click a row to expand recent sign-in history</p>
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
    $unlicensedUsersCsv = Join-Path $AuditFolder "Entra_Users_Unlicensed.csv"
    if (Test-Path $unlicensedUsersCsv) {
        $unlicCount = @(Import-Csv $unlicensedUsersCsv).Count
        if ($unlicCount -gt 0) {
            $entraSummary.Add("<p class='warn'>$unlicCount member account(s) have no licence assigned (see Entra_Users_Unlicensed.csv)</p>")
        }
    }

    # --- License Summary ---
    $licensesCsv = Join-Path $AuditFolder "Entra_Licenses.csv"
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
    $globalAdminsCsv = Join-Path $AuditFolder "Entra_GlobalAdmins.csv"
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
    $adminRolesCsv = Join-Path $AuditFolder "Entra_AdminRoles.csv"
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
    $caPoliciesCsv  = Join-Path $AuditFolder "Entra_CA_Policies.csv"
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

            $entraSummary.Add("<p style='font-size:0.85rem;color:#666;margin-bottom:0.5rem'>Click a row to expand policy scope and conditions.</p>")
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
                $mainRow = "<tr class='user-row' onclick='togglePerms(this)' title='Click to expand policy details'><td>$($policy.Name)</td><td class='$stateClass'>$stateLabel</td><td>$($policy.GrantControls)</td><td>$mfaIcon</td></tr>"

                $incUsers    = if ($policy.IncludeUsers  -and $policy.IncludeUsers  -ne '') { $policy.IncludeUsers  } else { '—' }
                $excUsers    = if ($policy.ExcludeUsers  -and $policy.ExcludeUsers  -ne '') { $policy.ExcludeUsers  } else { '—' }
                $incGroups   = if ($policy.IncludeGroups -and $policy.IncludeGroups -ne '') { $policy.IncludeGroups } else { '—' }
                $clientTypes = if ($policy.ClientAppTypes -and $policy.ClientAppTypes -ne '') { $policy.ClientAppTypes } else { 'All' }
                $detailRow = "<tr class='signin-detail' style='display:none'><td colspan='4'><table class='inner-table'>
  <thead><tr><th style='width:160px'>Setting</th><th>Value</th></tr></thead>
  <tbody>
    <tr><td>Include Users</td><td>$incUsers</td></tr>
    <tr><td>Exclude Users</td><td>$excUsers</td></tr>
    <tr><td>Include Groups</td><td>$incGroups</td></tr>
    <tr><td>Client App Types</td><td>$clientTypes</td></tr>
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
            $licensesCsv  = Join-Path $AuditFolder "Entra_Licenses.csv"
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
    $creationsCsv = Join-Path $AuditFolder "Entra_AccountCreations.csv"
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
    $deletionsCsv = Join-Path $AuditFolder "Entra_AccountDeletions.csv"
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
    $auditEventsCsv = Join-Path $AuditFolder "Entra_AuditEvents.csv"
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

    $html.Add((Add-Section -Title "Microsoft Entra" -CsvFiles $entraFiles.FullName -SummaryHtml ($entraSummary -join "`n")))
}


# =========================================
# ===   Exchange Section                ===
# =========================================
$exchangeFiles = @(Get-ChildItem "$AuditFolder\Exchange_*.csv" -ErrorAction SilentlyContinue)

if ($exchangeFiles.Count -gt 0) {
    $exchangeSummary = [System.Collections.Generic.List[string]]::new()

    $mbxCsv          = Join-Path $AuditFolder "Exchange_Mailboxes.csv"
    $forwardingCsv   = Join-Path $AuditFolder "Exchange_InboxForwardingRules.csv"
    $fullAccessCsv   = Join-Path $AuditFolder "Exchange_Permissions_FullAccess.csv"
    $sendAsCsv       = Join-Path $AuditFolder "Exchange_Permissions_SendAs.csv"
    $sendOnBehalfCsv = Join-Path $AuditFolder "Exchange_Permissions_SendOnBehalf.csv"

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

        $exchangeSummary.Add("<p style='font-size:0.85rem;color:#666;margin-bottom:0.5rem'>Click a row to expand delegated permissions.</p>")

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
    $brokenCsv = Join-Path $AuditFolder "Exchange_BrokenInboxRules.csv"
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
    $remoteDomainCsv = Join-Path $AuditFolder "Exchange_RemoteDomainForwarding.csv"
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
    $auditConfigCsv = Join-Path $AuditFolder "Exchange_AuditConfig.csv"
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
    $mbxAuditCsv = Join-Path $AuditFolder "Exchange_MailboxAuditStatus.csv"
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
    $dkimExCsv = Join-Path $AuditFolder "Exchange_DKIM_Status.csv"
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
    $antiPhishCsv = Join-Path $AuditFolder "Exchange_AntiPhishPolicies.csv"
    if (Test-Path $antiPhishCsv) {
        $antiPhish = @(Import-Csv $antiPhishCsv)
        $exchangeSummary.Add("<h4>Anti-Phish Policies</h4>")
        $exchangeSummary.Add("<p style='font-size:0.85rem;color:#666;margin-bottom:0.5rem'>Click a row to expand setting descriptions.</p>")
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
    $spamCsv = Join-Path $AuditFolder "Exchange_SpamPolicies.csv"
    if (Test-Path $spamCsv) {
        $spamPolicies = @(Import-Csv $spamCsv)
        $exchangeSummary.Add("<h4>Spam Filter Policies</h4>")
        $exchangeSummary.Add("<p style='font-size:0.85rem;color:#666;margin-bottom:0.5rem'>Click a row to expand action descriptions.</p>")
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
    $malwareCsv = Join-Path $AuditFolder "Exchange_MalwarePolicies.csv"
    if (Test-Path $malwareCsv) {
        $malware = @(Import-Csv $malwareCsv)
        $exchangeSummary.Add("<h4>Malware Filter Policies</h4>")
        $exchangeSummary.Add("<p style='font-size:0.85rem;color:#666;margin-bottom:0.5rem'>Click a row to expand setting descriptions.</p>")
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
    $transportCsv = Join-Path $AuditFolder "Exchange_TransportRules.csv"
    if (Test-Path $transportCsv) {
        $transportRules = @(Import-Csv $transportCsv)
        if ($transportRules.Count -gt 0) {
            $disabledRules = @($transportRules | Where-Object { $_.State -ne 'Enabled' })
            $trClass       = if ($disabledRules.Count -gt 0) { "warn" } else { "ok" }
            $exchangeSummary.Add("<h4>Transport Rules</h4>")
            $exchangeSummary.Add("<p class='$trClass'>$($transportRules.Count) transport rule(s) — $($disabledRules.Count) disabled. Click a row to expand conditions and actions.</p>")
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
    $dlCsv = Join-Path $AuditFolder "Exchange_DistributionLists.csv"
    if (Test-Path $dlCsv) {
        $dls      = @(Import-Csv $dlCsv)
        $emptyDls = @($dls | Where-Object { [int]$_.MemberCount -eq 0 })
        $exchangeSummary.Add("<h4>Distribution Lists ($($dls.Count) total)</h4>")
        if ($emptyDls.Count -gt 0) {
            $exchangeSummary.Add("<p class='warn'>$($emptyDls.Count) distribution list(s) have no members</p>")
        }
        $exchangeSummary.Add("<p style='font-size:0.85rem;color:#666;margin-bottom:0.5rem'>Click a row to expand the member list.</p>")
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
    $resourceCsv = Join-Path $AuditFolder "Exchange_ResourceMailboxes.csv"
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
    $outboundFwdCsv = Join-Path $AuditFolder "Exchange_OutboundSpamAutoForward.csv"
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
    $sharedSignInCsv = Join-Path $AuditFolder "Exchange_SharedMailboxSignIn.csv"
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
    $safeAttCsv = Join-Path $AuditFolder "Exchange_SafeAttachments.csv"
    $exchangeSummary.Add("<h4>Defender for Office 365 — Safe Attachments</h4>")
    if (Test-Path $safeAttCsv) {
        $safeAtt    = @(Import-Csv $safeAttCsv)
        $attEnabled = ($safeAtt | Where-Object { $_.Enable -eq "True" }).Count
        $attClass   = if ($attEnabled -gt 0) { "ok" } else { "warn" }
        $exchangeSummary.Add("<p class='$attClass'>$attEnabled of $($safeAtt.Count) Safe Attachment policy/policies enabled</p>")
        $exchangeSummary.Add("<p style='font-size:0.85rem;color:#666;margin-bottom:0.5rem'>Click a row to expand setting descriptions.</p>")
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
    $safeLinkCsv = Join-Path $AuditFolder "Exchange_SafeLinks.csv"
    $exchangeSummary.Add("<h4>Defender for Office 365 — Safe Links</h4>")
    if (Test-Path $safeLinkCsv) {
        $safeLink     = @(Import-Csv $safeLinkCsv)
        $linkEnabled  = ($safeLink | Where-Object { $_.EnableSafeLinksForEmail -eq "True" }).Count
        $linkClass    = if ($linkEnabled -gt 0) { "ok" } else { "warn" }
        $exchangeSummary.Add("<p class='$linkClass'>$linkEnabled of $($safeLink.Count) Safe Links policy/policies enabled for email</p>")
        $exchangeSummary.Add("<p style='font-size:0.85rem;color:#666;margin-bottom:0.5rem'>Click a row to expand setting descriptions.</p>")
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

    $html.Add((Add-Section -Title "Exchange Online" -CsvFiles $exchangeFiles.FullName -SummaryHtml ($exchangeSummary -join "`n")))
}


# =========================================
# ===   SharePoint / OneDrive Section   ===
# =========================================
$spFiles = @(Get-ChildItem "$AuditFolder\SharePoint_*.csv" -ErrorAction SilentlyContinue | Sort-Object Name)

if ($spFiles.Count -gt 0) {
    $spSummary = [System.Collections.Generic.List[string]]::new()

    $storageCsv     = Join-Path $AuditFolder "SharePoint_TenantStorage.csv"
    $sitesCsv       = Join-Path $AuditFolder "SharePoint_Sites.csv"
    $groupsCsv      = Join-Path $AuditFolder "SharePoint_SPGroups.csv"
    $tenantShareCsv = Join-Path $AuditFolder "SharePoint_ExternalSharing_Tenant.csv"
    $overridesCsv   = Join-Path $AuditFolder "SharePoint_ExternalSharing_SiteOverrides.csv"
    $odUsageCsv     = Join-Path $AuditFolder "SharePoint_OneDriveUsage.csv"
    $unlicensedCsv  = Join-Path $AuditFolder "SharePoint_OneDrive_Unlicensed.csv"
    $acpCsv         = Join-Path $AuditFolder "SharePoint_AccessControlPolicies.csv"

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

        $spSummary.Add("<h4>Site Collections ($($sites.Count))</h4>")
        $spSummary.Add("<p style='font-size:0.85rem;color:#666;margin-bottom:0.5rem'>Click a row to expand SharePoint groups for that site.</p>")

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
            if ($siteGroups.Count -gt 0) {
                $groupRows = ($siteGroups | ForEach-Object {
                    # Resolve SharePoint claim tokens to readable labels
                    $cleanMembers = if ($_.Members) {
                        ($_.Members -split ';\s*' | ForEach-Object {
                            $m = $_.Trim()
                            if     ($m -match '^i:0#\.f\|membership\|(.+)$')           { $Matches[1] }
                            elseif ($m -match '^c:0t\.c\|tenant\|')                    { '[All org users]' }
                            elseif ($m -match '^c:0-\.f\|rolemanager\|spo-grid-all-users') { '[Everyone]' }
                            elseif ($m -match '^c:0o\.c\|federateddirectoryclaimprovider\|') { '[M365 Group]' }
                            elseif ($m -match '^c:0\(\.s\|true')                       { '[All authenticated users]' }
                            elseif ($m -eq 'SHAREPOINT\system')                        { '[SharePoint System]' }
                            elseif ($m)                                                 { ($m -split '\|')[-1] }
                        } | Where-Object { $_ }) -join ', '
                    } else { '—' }
                    "<tr><td>$($_.GroupName)</td><td>$($_.Owner)</td><td style='text-align:center'>$($_.MemberCount)</td><td style='font-size:0.8rem;word-break:break-all'>$cleanMembers</td></tr>"
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

    $html.Add((Add-Section -Title "SharePoint / OneDrive" -CsvFiles $spFiles.FullName -SummaryHtml ($spSummary -join "`n")))
}


# =========================================
# ===   Mail Security Section           ===
# =========================================
$mailSecFiles = @(Get-ChildItem "$AuditFolder\MailSec_*.csv" -ErrorAction SilentlyContinue | Sort-Object Name)

if ($mailSecFiles.Count -gt 0) {
    $mailSecSummary = [System.Collections.Generic.List[string]]::new()

    $dkimCsv  = Join-Path $AuditFolder "MailSec_DKIM.csv"
    $dmarcCsv = Join-Path $AuditFolder "MailSec_DMARC.csv"
    $spfCsv   = Join-Path $AuditFolder "MailSec_SPF.csv"

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

    $html.Add((Add-Section -Title "Mail Security" -CsvFiles $mailSecFiles.FullName -SummaryHtml ($mailSecSummary -join "`n")))
}


# =========================================
# ===   Close and Write Report          ===
# =========================================
$html.Add(@"
<script>
function toggleSignIns(row) {
    var detail = row.nextElementSibling;
    if (detail && detail.classList.contains('signin-detail')) {
        var hidden = (detail.style.display === 'none' || detail.style.display === '');
        detail.style.display = hidden ? 'table-row' : 'none';
        row.classList.toggle('expanded', hidden);
    }
}
function togglePerms(row) {
    var detail = row.nextElementSibling;
    if (detail && detail.classList.contains('signin-detail')) {
        var hidden = (detail.style.display === 'none' || detail.style.display === '');
        detail.style.display = hidden ? 'table-row' : 'none';
        row.classList.toggle('expanded', hidden);
    }
}
</script>
</body></html>
"@)
$html -join "`n" | Set-Content -Path $reportPath -Encoding UTF8

Write-Host "Summary report written to: $reportPath" -ForegroundColor Green
if ($IsLinux) {
    xdg-open $reportPath
} elseif ($IsMacOS) {
    open $reportPath
} else {
    Start-Process $reportPath
}
