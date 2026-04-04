#Requires -Version 7.2

# ── NeConnect365Audit Module ────────────────────────────────────────────────
# Root module — loads all public and private functions, initialises the
# Microsoft Graph SDK assembly preloading to prevent MSAL version conflicts.

$ErrorActionPreference = 'Stop'

# ── Module-scoped state ─────────────────────────────────────────────────────
# Audit context stores per-tenant credentials and config for the current run.
# Set by Set-AuditContext, read by Get-AuditContext and all connection functions.
$script:AuditContext = $null

# Graph SDK version cache — resolved once per session
$script:GraphModuleVersion = $null
$script:GraphDependencyDirectories = $null

# ── Load private functions ──────────────────────────────────────────────────
$privatePath = Join-Path $PSScriptRoot 'Private'
if (Test-Path $privatePath) {
    Get-ChildItem -Path $privatePath -Filter '*.ps1' -Recurse | ForEach-Object {
        . $_.FullName
    }
}

# ── Load audit module functions ─────────────────────────────────────────────
$modulesPath = Join-Path $PSScriptRoot 'Modules'
if (Test-Path $modulesPath) {
    Get-ChildItem -Path $modulesPath -Filter '*.ps1' -Recurse | ForEach-Object {
        . $_.FullName
    }
}

# ── Load public functions ───────────────────────────────────────────────────
$publicPath = Join-Path $PSScriptRoot 'Public'
if (Test-Path $publicPath) {
    Get-ChildItem -Path $publicPath -Filter '*.ps1' -Recurse | ForEach-Object {
        . $_.FullName
    }
}

# ── Initialise Graph SDK assemblies ─────────────────────────────────────────
# Must happen at module import time, BEFORE any Az.* module loads, to ensure
# the correct Azure.Identity / Microsoft.Kiota assembly versions are in the
# AppDomain. If Az.Accounts loads first, its MSAL assemblies conflict with
# the Graph SDK's expected versions.
try {
    Initialize-GraphSdk
}
catch {
    Write-Warning "Graph SDK assembly preloading failed: $($_.Exception.Message). Graph operations may encounter assembly conflicts."
}
