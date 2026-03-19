<#
.SYNOPSIS
    Uninstalls all PowerShell modules used by 365Audit.

.DESCRIPTION
    Removes every installed version of the modules required by 365Audit — including all
    Microsoft.Graph sub-modules, ExchangeOnlineManagement, and PnP.PowerShell.

    Useful for:
      - Testing a clean first-run install experience
      - Clearing conflicting module versions before re-installing
      - Decommissioning a machine

    Modules are removed from all scopes (CurrentUser and AllUsers) where found.
    Run as Administrator if you need to remove AllUsers-scoped installs.

.PARAMETER WhatIf
    Show which modules would be removed without actually removing them.

.EXAMPLE
    .\Helpers\Uninstall-AuditModules.ps1

.EXAMPLE
    .\Helpers\Uninstall-AuditModules.ps1 -WhatIf

.NOTES
    Author  : Raymond Slater
    Version : 1.0.0
#>

#Requires -Version 7.2

[CmdletBinding(SupportsShouldProcess)]
param()

$ErrorActionPreference = 'Continue'

# ── Full module list ───────────────────────────────────────────────────────────

$_modules = @(
    # Microsoft Graph — core
    'Microsoft.Graph.Authentication'
    'Microsoft.Graph.Identity.DirectoryManagement'

    # Microsoft Graph — Entra audit
    'Microsoft.Graph.Users'
    'Microsoft.Graph.Groups'
    'Microsoft.Graph.Reports'
    'Microsoft.Graph.Identity.SignIns'
    'Microsoft.Graph.Applications'

    # Microsoft Graph — Intune audit
    'Microsoft.Graph.DeviceManagement'
    'Microsoft.Graph.Devices.CorporateManagement'
    'Microsoft.Graph.DeviceManagement.Enrollment'

    # Exchange / Mail Security audit
    'ExchangeOnlineManagement'

    # SharePoint audit
    'PnP.PowerShell'
)

# ── Pre-flight: block if any modules are currently loaded ─────────────────────

$_loaded = $_modules | Where-Object { Get-Module -Name $_ -All }
if ($_loaded) {
    Write-Host ""
    Write-Host "Cannot uninstall — the following modules are currently loaded in this session:" -ForegroundColor Red
    $_loaded | ForEach-Object { Write-Host "  - $_" -ForegroundColor Yellow }
    Write-Host ""
    Write-Host "Close ALL PowerShell/VS Code terminals, open a single fresh session, and try again." -ForegroundColor Red
    exit 1
}

# ── Remove each module ─────────────────────────────────────────────────────────

$removed = 0
$skipped = 0

foreach ($modName in $_modules) {
    $installed = @(Get-Module -ListAvailable -Name $modName)

    if ($installed.Count -eq 0) {
        Write-Host "  SKIP  $modName — not installed" -ForegroundColor DarkGray
        $skipped++
        continue
    }

    foreach ($mod in $installed) {
        $label = "$modName $($mod.Version)"
        if ($PSCmdlet.ShouldProcess($label, 'Uninstall-Module')) {
            try {
                Uninstall-Module -Name $modName -RequiredVersion $mod.Version -Force -ErrorAction Stop
                Write-Host "  OK    $label" -ForegroundColor Green
                $removed++
            }
            catch {
                Write-Warning "  FAIL  $label — $_"
            }
        }
    }
}

# ── Summary ────────────────────────────────────────────────────────────────────

Write-Host ""
if ($WhatIfPreference) {
    Write-Host "WhatIf: no modules were removed." -ForegroundColor Yellow
}
else {
    Write-Host "Done. Removed: $removed   Skipped (not installed): $skipped" -ForegroundColor Cyan
    if ($removed -gt 0) {
        Write-Host "Open a new PowerShell session before re-running 365Audit." -ForegroundColor Yellow
    }
}
