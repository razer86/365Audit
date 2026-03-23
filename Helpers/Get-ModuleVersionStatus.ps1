<#
.SYNOPSIS
    Checks installed versions of all 365Audit modules against the latest available in PSGallery.

.DESCRIPTION
    For each module required by 365Audit, displays:
      - Currently installed version(s)
      - Latest version available in PSGallery
      - Whether an update is available

    Useful for identifying outdated or conflicting module versions before running an audit.

.EXAMPLE
    .\Helpers\Get-ModuleVersionStatus.ps1

.NOTES
    Author  : Raymond Slater
    Version : 1.0.0
#>

#Requires -Version 7.2

$ErrorActionPreference = 'Continue'

# ScubaGear installs into the Windows PowerShell 5.1 module path.
# Inject it into PSModulePath so Get-Module -ListAvailable and Find-Module can reach it.
$_ps51ModulePath = Join-Path $env:USERPROFILE 'Documents\WindowsPowerShell\Modules'
if ((Test-Path $_ps51ModulePath) -and ($env:PSModulePath -notlike "*$_ps51ModulePath*")) {
    $env:PSModulePath += ";$_ps51ModulePath"
}

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

    # Teams audit
    'MicrosoftTeams'

    # ScubaGear audit (Windows PowerShell 5.1 — installed to Documents\WindowsPowerShell\Modules)
    'ScubaGear'
)

Write-Host ""
Write-Host "Checking module versions..." -ForegroundColor Cyan

# Single bulk PSGallery lookup — much faster than one Find-Module call per module
Write-Host "  Querying PSGallery..." -ForegroundColor DarkGray
$_galleryResults = @{}
try {
    Find-Module -Name $_modules -ErrorAction Stop | ForEach-Object {
        $_galleryResults[$_.Name] = $_.Version
    }
}
catch {
    Write-Warning "PSGallery lookup failed: $_"
}

Write-Host ""

$results = foreach ($modName in $_modules) {
    $installed = @(Get-Module -ListAvailable -Name $modName | Sort-Object Version -Descending)

    $installedVersions = if ($installed.Count -gt 0) {
        ($installed | ForEach-Object { $_.Version.ToString() }) -join ', '
    } else {
        $null
    }

    $latestVersion = if ($_galleryResults.ContainsKey($modName)) { $_galleryResults[$modName] } else { '(unavailable)' }
    $newestInstalled = if ($installed.Count -gt 0) { $installed[0].Version } else { $null }

    $status = if (-not $newestInstalled) {
        'NOT INSTALLED'
    } elseif ($latestVersion -eq '(unavailable)') {
        'UNKNOWN'
    } elseif ([version]$newestInstalled -lt [version]$latestVersion) {
        'UPDATE AVAILABLE'
    } elseif ($installed.Count -gt 1) {
        'MULTIPLE VERSIONS'
    } else {
        'OK'
    }

    [PSCustomObject]@{
        Module    = $modName
        Installed = if ($installedVersions) { $installedVersions } else { '—' }
        Latest    = $latestVersion
        Status    = $status
    }
}

# ── Display ────────────────────────────────────────────────────────────────────

foreach ($row in $results) {
    $colour = switch ($row.Status) {
        'OK'                { 'Green'  }
        'NOT INSTALLED'     { 'DarkGray' }
        'UPDATE AVAILABLE'  { 'Yellow' }
        'MULTIPLE VERSIONS' { 'Cyan'   }
        default             { 'White'  }
    }

    $icon = switch ($row.Status) {
        'OK'                { '  OK  ' }
        'NOT INSTALLED'     { ' MISS  ' }
        'UPDATE AVAILABLE'  { '  UP  ' }
        'MULTIPLE VERSIONS' { ' MULTI ' }
        default             { '  ??  ' }
    }

    Write-Host "$icon" -ForegroundColor $colour -NoNewline
    Write-Host "$($row.Module.PadRight(48))" -NoNewline
    Write-Host "installed: $($row.Installed.PadRight(10))  latest: $($row.Latest)" -ForegroundColor DarkGray
}

Write-Host ""

# ── Known incompatibility warnings ─────────────────────────────────────────────

$_knownIssues = @()

# EXO 3.8.0 ships an msalruntime.dll that conflicts with Microsoft.Graph.Authentication.
# The dll is placed in a subdirectory that EXO's assembly resolver doesn't add to the load
# path, so MSAL fails to initialise when Graph and EXO are both loaded in the same session.
# Safe known-good pairing: Graph SDK 2.35.x + ExchangeOnlineManagement 3.7.1.
$_exoInstalled = Get-Module -ListAvailable -Name ExchangeOnlineManagement |
    Sort-Object Version -Descending | Select-Object -First 1
if ($_exoInstalled -and $_exoInstalled.Version -ge [version]'3.8.0') {
    $_knownIssues += "ExchangeOnlineManagement $($_exoInstalled.Version) has known MSAL conflicts with Microsoft.Graph.Authentication. " +
        "Downgrade to 3.7.1: Uninstall-Module ExchangeOnlineManagement -AllVersions -Force; Install-Module ExchangeOnlineManagement -RequiredVersion 3.7.1 -Scope CurrentUser"

    # Check whether msalruntime.dll is also missing from the flat netCore folder
    $_exoBase    = Split-Path $_exoInstalled.ModuleBase
    $_msalTarget = Join-Path $_exoBase "$($_exoInstalled.Version)\netCore\msalruntime.dll"
    $_msalSource = Join-Path $_exoBase "$($_exoInstalled.Version)\netCore\runtimes\win-x64\native\msalruntime.dll"
    if (-not (Test-Path $_msalTarget) -and (Test-Path $_msalSource)) {
        $_knownIssues += "msalruntime.dll is missing from EXO module load path. Fix: " +
            "Copy-Item '$_msalSource' '$_msalTarget'"
    }
}

if ($_knownIssues.Count -gt 0) {
    Write-Host "+----------------------------------------------------------+" -ForegroundColor Yellow
    Write-Host "  Module Issue" -ForegroundColor Yellow
    Write-Host "+----------------------------------------------------------+" -ForegroundColor Yellow
    foreach ($issue in $_knownIssues) {
        Write-Host "  ! $issue" -ForegroundColor Yellow
    }
    Write-Host ""
    Write-Host "  Known compatible combo: Graph SDK 2.35.x + EXO 3.7.1" -ForegroundColor DarkGray
    Write-Host "  Also: Always connect Graph BEFORE EXO in the same session." -ForegroundColor DarkGray
    Write-Host ""
}

# ── Summary ────────────────────────────────────────────────────────────────────

$missing  = @($results | Where-Object Status -eq 'NOT INSTALLED')
$outdated = @($results | Where-Object Status -eq 'UPDATE AVAILABLE')
$multi    = @($results | Where-Object Status -eq 'MULTIPLE VERSIONS')

if ($missing.Count -eq 0 -and $outdated.Count -eq 0 -and $multi.Count -eq 0) {
    Write-Host "All modules are installed and up to date." -ForegroundColor Green
} else {
    if ($missing.Count  -gt 0) { Write-Host "Not installed    : $($missing.Count)  — run 365Audit to install automatically" -ForegroundColor DarkGray }
    if ($outdated.Count -gt 0) { Write-Host "Updates available: $($outdated.Count)  — run Install-Module <name> -Force to update" -ForegroundColor Yellow }
    if ($multi.Count    -gt 0) { Write-Host "Multiple versions: $($multi.Count)  — run .\Helpers\Uninstall-AuditModules.ps1 to clean up" -ForegroundColor Cyan }
}

Write-Host ""
