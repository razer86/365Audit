<#
.SYNOPSIS
    Interactive launcher for the Microsoft 365 Audit toolkit.

.DESCRIPTION
    Presents a menu of available audit modules. Modules are loaded from local disk;
    missing modules are downloaded from GitHub as a fallback. On startup, compares
    local script versions against the GitHub version manifest and warns if any are outdated.

.PARAMETER AppId
    Azure AD application (client) ID for app-only authentication.
    Run Setup-365AuditApp.ps1 to create the app registration.

.PARAMETER TenantId
    Azure AD tenant ID (GUID or .onmicrosoft.com domain) of the customer tenant.

.PARAMETER CertBase64
    Base64-encoded contents of the .pfx certificate file.
    Paste the value from Hudu (output by Setup-365AuditApp.ps1).
    The script writes a temp .pfx to $env:TEMP and deletes it on exit.
    If omitted, the script will prompt for it interactively.

.PARAMETER CertPassword
    Password for the .pfx certificate file.

.EXAMPLE
    .\Start-365Audit.ps1 -AppId '<guid>' -TenantId '<guid>' -CertPassword (Read-Host -AsSecureString)
    Prompts interactively for the certificate Base64 (paste from Hudu) and certificate password.

.NOTES
    Author      : Raymond Slater
    Version     : 2.5.0
    Change Log  : See CHANGELOG.md

.LINK
    https://github.com/razer86/365Audit
#>

#Requires -Version 7.2

param (
    [Parameter(Mandatory, HelpMessage = 'Azure AD application (client) ID. Run Setup-365AuditApp.ps1 to obtain.')]
    [string]$AppId,

    [Parameter(Mandatory, HelpMessage = 'Azure AD tenant ID (GUID or .onmicrosoft.com domain). Run Setup-365AuditApp.ps1 to obtain.')]
    [string]$TenantId,

    [Parameter(HelpMessage = 'Base64-encoded .pfx certificate (output by Setup-365AuditApp.ps1, stored in Hudu). Omit to be prompted.')]
    [string]$CertBase64,

    [Parameter(Mandatory, HelpMessage = 'Password for the .pfx certificate file.')]
    [SecureString]$CertPassword
)

$ScriptVersion = "2.5.0"
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

# === Decode base64 cert to a temp .pfx (deleted on exit) ===
if (-not $CertBase64) {
    $CertBase64 = Read-Host 'Paste certificate Base64'
}
$_tempCertPath = Join-Path $env:TEMP "365Audit-$(New-Guid).pfx"
[System.IO.File]::WriteAllBytes($_tempCertPath, [Convert]::FromBase64String($CertBase64))
$CertFilePath  = $_tempCertPath
Write-Verbose "Certificate decoded from base64 to temp file: $_tempCertPath"

# Expose app credentials so dot-sourced modules can access them.
$AuditAppId        = $AppId
$AuditTenantId     = $TenantId
$AuditCertFilePath = $CertFilePath
$AuditCertPassword = $CertPassword

# === Check for updates ===
Invoke-VersionCheck -ScriptRoot $PSScriptRoot

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
try {
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
}
finally {
    if (Test-Path $_tempCertPath) {
        Remove-Item $_tempCertPath -Force -ErrorAction SilentlyContinue
        Write-Verbose "Temp certificate file removed: $_tempCertPath"
    }
}
