<#
.SYNOPSIS
    Interactive launcher for the Microsoft 365 Audit toolkit.

.DESCRIPTION
    Presents a menu of available audit modules. Modules are loaded from local disk;
    missing modules are downloaded from GitHub as a fallback. On startup, compares
    local script versions against the GitHub version manifest and warns if any are outdated.

.PARAMETER AppId
    Azure AD application (client) ID for app-only Entra/Exchange authentication.
    When provided alongside -AppSecret and -TenantId, Entra and Exchange modules
    will authenticate silently using the app credentials.

.PARAMETER AppSecret
    Client secret for the app registration specified by -AppId.

.PARAMETER TenantId
    Azure AD tenant ID (GUID or .onmicrosoft.com domain) for app-only auth.

.PARAMETER PnPAppId
    Azure AD application (client) ID of the PnP interactive auth app registered
    by Setup-365AuditApp.ps1. Required for the SharePoint Online audit module.

.NOTES
    Author      : Raymond Slater
    Version     : 1.9.0
    Change Log  :
        1.0.0 - Initial release
        1.0.1 - Standardised comments; pass folder to summary
        1.1.0 - Removed duplicate helper functions (moved to Audit-Common.ps1);
                fixed menu option 9 script name; added version check on startup;
                added Invoke-MailSecurityAudit.ps1 as option 4
        1.2.0 - Added optional -AppId/-AppSecret/-TenantId parameters for
                app-only SharePoint authentication (MSP cross-tenant support)
        1.3.0 - SharePoint module now skipped with setup guidance when app
                credentials are not supplied at launch
        1.4.0 - Added -CertThumbprint parameter; SharePoint audit now requires a
                certificate (SharePoint admin APIs reject client-secret tokens)
        1.5.0 - Reverted SharePoint to interactive auth; removed -CertThumbprint
                parameter and SharePoint skip block
        1.6.0 - Added -PnPAppId parameter for the dedicated PnP interactive auth
                app registered by Setup-365AuditApp.ps1 (Register-PnPEntraIDAppForInteractiveLogin)
        1.7.0 - Added Generate-AuditSummary.ps1 to option 9 script list so "Run All"
                automatically generates the summary report after all modules complete
        1.9.0 - Generate-AuditSummary.ps1 removed from menu Script arrays; summary
                now runs once after all selected modules complete to avoid
                multiple report generations when selecting e.g. "1,2"
        1.8.0 - AppId, AppSecret, and TenantId are now mandatory parameters with
                HelpMessage guidance; PnPAppId warning displayed at startup when
                not provided; removed unnecessary Start-Sleep stalls

.LINK
    https://github.com/razer86/365Audit
#>

#Requires -Version 7.2

param (
    [Parameter(Mandatory, HelpMessage = 'Azure AD application (client) ID. Run Setup-365AuditApp.ps1 to obtain.')]
    [string]$AppId,

    [Parameter(Mandatory, HelpMessage = 'Client secret for the app registration. Run Setup-365AuditApp.ps1 to obtain.')]
    [string]$AppSecret,

    [Parameter(Mandatory, HelpMessage = 'Azure AD tenant ID (GUID or .onmicrosoft.com domain). Run Setup-365AuditApp.ps1 to obtain.')]
    [string]$TenantId,

    [Parameter(HelpMessage = 'PnP interactive auth app ID. Optional — SharePoint audit is skipped if not provided. Run Setup-365AuditApp.ps1 to obtain.')]
    [string]$PnPAppId
)

$ScriptVersion = "1.9.0"
Write-Verbose "Start-365Audit.ps1 loaded (v$ScriptVersion)"

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# === Load shared helper functions ===
$commonPath = Join-Path $PSScriptRoot "common\Audit-Common.ps1"
if (Test-Path $commonPath) {
    . $commonPath
}
else {
    Write-Error "Required helper script not found: $commonPath"
    exit 1
}

# === Config ===
$localPath = $PSScriptRoot

# Expose app credentials so dot-sourced modules can access them.
# Variables are empty strings when not supplied; modules check for non-empty values before using.
$AuditAppId     = $AppId
$AuditAppSecret = $AppSecret
$AuditTenantId  = $TenantId
$AuditPnPAppId  = $PnPAppId

# === Check for updates ===
Invoke-VersionCheck -ScriptRoot $PSScriptRoot

if (-not $PnPAppId) {
    Write-Host "`n[!] -PnPAppId not provided — SharePoint Online audit will be skipped." -ForegroundColor Yellow
    Write-Host "    Run Setup-365AuditApp.ps1 to register the PnP app, then re-launch with -PnPAppId <id>.`n" -ForegroundColor Yellow
}

# === Define Menu Items ===
$menu = @{
    1 = @{ Name = "Microsoft Entra Audit";      Script = @("Invoke-EntraAudit.ps1") }
    2 = @{ Name = "Exchange Online Audit";      Script = @("Invoke-ExchangeAudit.ps1") }
    3 = @{ Name = "SharePoint Online Audit";    Script = @("Invoke-SharePointAudit.ps1") }
    4 = @{ Name = "Mail Security Audit";        Script = @("Invoke-MailSecurityAudit.ps1") }
    9 = @{ Name = "Run All Modules (1,2,3,4)";  Script = @("Invoke-EntraAudit.ps1", "Invoke-ExchangeAudit.ps1", "Invoke-SharePointAudit.ps1", "Invoke-MailSecurityAudit.ps1") }
    0 = @{ Name = "Exit";                       Script = $null }
}

# === Display Menu ===
Write-Host "`n╔════════════════════════════════════╗"
Write-Host "║    Microsoft 365 Audit Launcher    ║"
Write-Host "╚════════════════════════════════════╝"

foreach ($key in ($menu.Keys | Sort-Object { [int]$_ })) {
    Write-Host "$key. $($menu[$key].Name)"
}

# === User Selection ===
$selection = Read-Host "`nSelect one or more modules (comma separated, e.g. 1,2)"
if ($selection -eq "0") {
    Write-Host "Exiting. Goodbye!"
    return
}

# === Parse Selection ===
$selectedIndexes = $selection -split "," |
    ForEach-Object { $_.Trim() } |
    Where-Object    { $_ -match '^\d+$' } |
    ForEach-Object  { [int]$_ }

# === Execute Selected Modules ===
foreach ($index in $selectedIndexes) {
    if (-not $menu.ContainsKey($index)) {
        Write-Warning "Invalid selection: $index"
        continue
    }

    $module = $menu[$index]
    if (-not $module.Script) { continue }

    $scriptsToRun = @($module.Script)

    foreach ($scriptName in $scriptsToRun) {
        $localScriptPath = Join-Path $localPath $scriptName
        $remoteScriptUrl = "$RemoteBaseUrl/$scriptName"

        Write-Host "`n================================================================"
        Write-Host "  Starting: $scriptName" -ForegroundColor Cyan
        Write-Host "================================================================"

        if (Test-Path $localScriptPath) {
            Write-Host "Loading local script: $localScriptPath`n"
            . $localScriptPath
        }
        else {
            Write-Host "Local script not found. Downloading from GitHub..."
            Write-Host "Fetching: $remoteScriptUrl`n"
            try {
                Invoke-Expression (Invoke-RestMethod $remoteScriptUrl)
            }
            catch {
                Write-Warning "Failed to download or run ${scriptName}: $_"
            }
        }

        Write-Host "Completed: $scriptName" -ForegroundColor Green
    }
}

# === Generate Summary Report (once, after all modules) ===
$summaryScript = Join-Path $localPath "Generate-AuditSummary.ps1"
$auditContext  = Initialize-AuditOutput
if ($auditContext -and (Test-Path $summaryScript)) {
    Write-Host "`n================================================================"
    Write-Host "  Starting: Generate-AuditSummary.ps1" -ForegroundColor Cyan
    Write-Host "================================================================"
    & $summaryScript -AuditFolder $auditContext.OutputPath
}
else {
    Write-Warning "No audit output context found — summary report skipped."
}
