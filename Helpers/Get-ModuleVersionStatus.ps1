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
