<#
.SYNOPSIS
    Creates the M365 Audit Toolkit asset layout in Hudu.

.DESCRIPTION
    Creates the asset layout required by Setup-365AuditApp.ps1 to store audit
    credentials (Application ID, Tenant ID, certificate, etc.) per company.

    The layout name defaults to the HuduAssetName value in config.psd1
    (typically 'M365 Audit Toolkit'). After a successful creation, the new
    layout ID is printed so you can set HuduAssetLayoutId in config.psd1.

    IMPORTANT: Creating asset layouts requires Hudu Administrator or Super
    Administrator privileges. A standard user API key will receive a 422 error.

    Reads HuduBaseUrl, HuduApiKey, and HuduAssetName from config.psd1.

.PARAMETER LayoutName
    Override the layout name. Defaults to HuduAssetName in config.psd1,
    or 'M365 Audit Toolkit' if not set.

.PARAMETER Icon
    Font Awesome icon class for the layout (e.g. 'fas fa-shield-halved').
    Defaults to 'fab fa-microsoft'.

.PARAMETER Color
    Background colour for the layout icon (hex, e.g. '#1849a9').
Defaults to '#1849a9' (blue).

.PARAMETER IconColor
    Icon foreground colour (hex). Defaults to '#FFFFFF' (white).

.PARAMETER HuduBaseUrl
    Base URL of your Hudu instance. Falls back to HuduBaseUrl in config.psd1.

.PARAMETER HuduApiKey
    Hudu API key. Falls back to HuduApiKey in config.psd1.

.PARAMETER WhatIf
    Show what would be created without actually calling the Hudu API.

.EXAMPLE
    .\Helpers\New-HuduAssetLayout.ps1

.EXAMPLE
    .\Helpers\New-HuduAssetLayout.ps1 -LayoutName 'M365 Audit Toolkit'

.NOTES
    Author      : Raymond Slater
    Version     : 1.0.0

    Requires Hudu Administrator or Super Administrator to create asset layouts.
#>

#Requires -Version 7.2

[CmdletBinding(SupportsShouldProcess)]
param (
    [string]$LayoutName,
    [string]$Icon      = 'fab fa-microsoft',
    [string]$Color     = '#1849a9',
    [string]$IconColor = '#FFFFFF',
    [string]$HuduBaseUrl,
    [string]$HuduApiKey
)

$ScriptVersion         = "1.0.0"
Write-Verbose "New-HuduAssetLayout.ps1 loaded (v$ScriptVersion)"

$ErrorActionPreference = 'Stop'

# ── Load config.psd1 fallbacks ─────────────────────────────────────────────────

$_configPath = Join-Path $PSScriptRoot '..' 'config.psd1'
if (Test-Path $_configPath) {
    try {
        $_config = Import-PowerShellDataFile -Path $_configPath
        if (-not $HuduBaseUrl -and $_config.HuduBaseUrl)  { $HuduBaseUrl = $_config.HuduBaseUrl }
        if (-not $HuduApiKey  -and $_config.HuduApiKey)   { $HuduApiKey  = $_config.HuduApiKey }
        if (-not $LayoutName  -and $_config.HuduAssetName) { $LayoutName  = $_config.HuduAssetName }
    }
    catch { Write-Warning "Could not load config.psd1: $_" }
}

if (-not $LayoutName)  { $LayoutName  = 'M365 Audit Toolkit' }
if (-not $HuduBaseUrl) { Write-Error "HuduBaseUrl is required. Set it in config.psd1 or pass -HuduBaseUrl." }
if (-not $HuduApiKey)  { Write-Error "HuduApiKey is required. Set it in config.psd1 or pass -HuduApiKey." }

$HuduBaseUrl = $HuduBaseUrl.TrimEnd('/')
$_headers    = @{ 'x-api-key' = $HuduApiKey; 'Content-Type' = 'application/json' }

# ── Field definitions ──────────────────────────────────────────────────────────
# These match exactly the fields read/written by Setup-365AuditApp.ps1.
# Field labels must match what Push-HuduAuditAsset and Get-HuduAuditCredentials expect.

$fields = @(
    @{ label = 'Application ID';           field_type = 'Text';     required = $true;  show_in_list = $false; position = 1 }
    @{ label = 'Tenant ID';                field_type = 'Text';     required = $true;  show_in_list = $false; position = 2 }
    @{ label = 'Cert Base64';              field_type = 'Password'; required = $true;  show_in_list = $false; position = 3 }
    @{ label = 'Cert Password';            field_type = 'Password'; required = $true;  show_in_list = $false; position = 4 }
    @{ label = 'Cert Expiry';              field_type = 'Date';     required = $true;  show_in_list = $true;  expiration = $true; position = 5 }
    @{ label = 'Powershell Launch Command'; field_type = 'RichText'; required = $false; show_in_list = $false; position = 6 }
)

# ── Summary ────────────────────────────────────────────────────────────────────

Write-Host ""
Write-Host "  Hudu Asset Layout to be created" -ForegroundColor Cyan
Write-Host "  ────────────────────────────────────────────────────────" -ForegroundColor DarkGray
Write-Host ("  {0,-22} {1}" -f "Name:",       $LayoutName)
Write-Host ("  {0,-22} {1}" -f "Icon:",       $Icon)
Write-Host ("  {0,-22} {1}" -f "Color:",      $Color)
Write-Host ("  {0,-22} {1}" -f "Icon color:", $IconColor)
Write-Host ("  {0,-22} {1}" -f "Hudu URL:",   $HuduBaseUrl)
Write-Host ""
Write-Host "  Fields:" -ForegroundColor DarkGray
Write-Host ("  {0,-30} {1,-12} {2,-9} {3}" -f "Label", "Type", "Required", "In List") -ForegroundColor DarkGray

foreach ($f in $fields) {
    $req    = if ($f.required)      { 'Yes' } else { 'No' }
    $inList = if ($f.show_in_list)  { 'Yes' } else { 'No' }
    Write-Host ("  {0,-30} {1,-12} {2,-9} {3}" -f $f.label, $f.field_type, $req, $inList)
}

Write-Host ""
Write-Host "  NOTE: Creating asset layouts requires Hudu Administrator or Super Administrator." -ForegroundColor Yellow
Write-Host ""

# ── WhatIf early exit ──────────────────────────────────────────────────────────

if (-not $PSCmdlet.ShouldProcess("$HuduBaseUrl/api/v1/asset_layouts", "Create asset layout '$LayoutName'")) {
    Write-Host "WhatIf: No changes made." -ForegroundColor DarkGray
    exit 0
}

# ── Confirmation ───────────────────────────────────────────────────────────────

$answer = Read-Host "  Create this layout in Hudu? [y/N]"
if ($answer -notmatch '^[Yy]$') {
    Write-Host "  Cancelled — no changes made." -ForegroundColor DarkGray
    exit 0
}

# ── Build payload ──────────────────────────────────────────────────────────────

$payload = @{
    asset_layout = @{
        name              = $LayoutName
        icon              = $Icon
        color             = $Color
        icon_color        = $IconColor
        active            = $true
        include_passwords = $false
        include_photos    = $false
        include_comments  = $false
        include_files     = $false
        fields            = $fields
    }
} | ConvertTo-Json -Depth 5

# ── Create ─────────────────────────────────────────────────────────────────────

Write-Host "  Creating asset layout..." -ForegroundColor Cyan

try {
    $response = Invoke-RestMethod `
        -Uri     "$HuduBaseUrl/api/v1/asset_layouts" `
        -Headers $_headers `
        -Method  Post `
        -Body    $payload `
        -ErrorAction Stop
}
catch {
    $statusCode = $null
    if ($_.Exception.Response) {
        $statusCode = [int]$_.Exception.Response.StatusCode
    }

    switch ($statusCode) {
        401 {
            Write-Error ("Hudu returned 401 Unauthorized.`n" +
                "Check that your HuduApiKey in config.psd1 is correct and has not expired.")
        }
        404 {
            Write-Error ("Hudu returned 404 Not Found.`n" +
                "Check that HuduBaseUrl in config.psd1 is correct and the Hudu instance is reachable.")
        }
        422 {
            Write-Error ("Hudu returned 422 Unprocessable Entity.`n" +
                "This usually means the API key belongs to an account without Administrator or " +
                "Super Administrator privileges, or a layout with this name already exists.")
        }
        default {
            Write-Error "Hudu API request failed (HTTP $statusCode): $_"
        }
    }
}

$newId   = $response.asset_layout.id
$newName = $response.asset_layout.name

Write-Host ""
Write-Host "  Asset layout created successfully." -ForegroundColor Green
Write-Host ("  {0,-22} {1}" -f "Name:", $newName)
Write-Host ("  {0,-22} {1}" -f "Layout ID:", $newId) -ForegroundColor Green
Write-Host ""
Write-Host "  Next step: update HuduAssetLayoutId in config.psd1:" -ForegroundColor Yellow
Write-Host "    HuduAssetLayoutId = $newId" -ForegroundColor Cyan
Write-Host ""

