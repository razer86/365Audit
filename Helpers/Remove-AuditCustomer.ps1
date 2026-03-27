<#
.SYNOPSIS
    Removes a customer's 365Audit app registration from Entra ID and archives
    the corresponding asset in Hudu.

.DESCRIPTION
    Used when a customer offboards or when you need to fully reset a customer's
    365Audit configuration.

    When run with -HuduCompanyId or -HuduCompanyName, the script:
      1. Looks up the customer's AppId and TenantId from their Hudu asset.
      2. Connects to Microsoft Graph interactively (browser sign-in).
      3. Deletes the app registration from Entra ID (soft-delete by default;
         use -PermanentDelete to purge from the recycle bin immediately).
      4. Archives the Hudu asset (preserves history; asset is hidden from active views).

    When run with -AppId and -TenantId directly (no Hudu), only the Entra app
    registration is removed.

    IMPORTANT: Soft-deleted apps can be restored from the Entra recycle bin for
    up to 30 days. Use -PermanentDelete only when you are certain the customer
    will not be re-onboarded.

.PARAMETER HuduCompanyId
    Hudu company slug (12-character hex, e.g. 'a1b2c3d4e5f6') or numeric company
    ID. Used to look up the customer's AppId and TenantId from Hudu automatically.

.PARAMETER HuduCompanyName
    Exact Hudu company name. Alternative to -HuduCompanyId.

.PARAMETER AppId
    Azure AD application (client) ID to remove. Use when not fetching from Hudu.
    Requires -TenantId.

.PARAMETER TenantId
    Tenant ID of the app registration. Required when using -AppId directly.

.PARAMETER HuduBaseUrl
    Hudu instance base URL. Falls back to config.psd1 in the script root, then
    the HUDU_BASE_URL environment variable.

.PARAMETER HuduApiKey
    Hudu API key. Falls back to config.psd1, then the HUDU_API_KEY environment
    variable.

.PARAMETER PermanentDelete
    Permanently purge the app registration from the Entra recycle bin immediately
    after soft-deleting it. Cannot be undone.

.EXAMPLE
    .\Helpers\Remove-AuditCustomer.ps1 -HuduCompanyId 'a1b2c3d4e5f6'
    Looks up the customer in Hudu, removes the Entra app, and archives the Hudu asset.

.EXAMPLE
    .\Helpers\Remove-AuditCustomer.ps1 -HuduCompanyId 'a1b2c3d4e5f6' -PermanentDelete
    Removes and immediately purges the app — use when certain the customer won't be re-onboarded.

.EXAMPLE
    .\Helpers\Remove-AuditCustomer.ps1 -AppId '00000000-0000-0000-0000-000000000000' -TenantId '00000000-0000-0000-0000-000000000000'
    Removes the Entra app directly without Hudu interaction.

.EXAMPLE
    .\Helpers\Remove-AuditCustomer.ps1 -HuduCompanyId 'a1b2c3d4e5f6' -WhatIf
    Preview what would be removed without making any changes.

.NOTES
    Author  : Raymond Slater
    Version : 1.0.0
#>

#Requires -Version 7.2

[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High', DefaultParameterSetName = 'Manual')]
param (
    # ── Manual credential parameters ──────────────────────────────────────────
    [Parameter(Mandatory, ParameterSetName = 'Manual',
        HelpMessage = 'Azure AD application (client) ID of the app registration to remove. Run Setup-365AuditApp.ps1 to obtain.')]
    [string]$AppId,

    [Parameter(Mandatory, ParameterSetName = 'Manual',
        HelpMessage = 'Azure AD tenant ID (GUID or .onmicrosoft.com domain).')]
    [string]$TenantId,

    # ── Hudu parameters ────────────────────────────────────────────────────────
    [Parameter(Mandatory, ParameterSetName = 'HuduById',
        HelpMessage = 'Hudu company slug (12-character hex) or numeric ID. AppId and TenantId are fetched from the Hudu asset automatically.')]
    [string]$HuduCompanyId,

    [Parameter(Mandatory, ParameterSetName = 'HuduByName',
        HelpMessage = 'Exact Hudu company name. AppId and TenantId are fetched from the Hudu asset automatically.')]
    [string]$HuduCompanyName,

    [Parameter(ParameterSetName = 'HuduById')]
    [Parameter(ParameterSetName = 'HuduByName')]
    [string]$HuduBaseUrl,

    [Parameter(ParameterSetName = 'HuduById')]
    [Parameter(ParameterSetName = 'HuduByName')]
    [string]$HuduApiKey,

    # ── Options ────────────────────────────────────────────────────────────────
    [Parameter(ParameterSetName = 'Manual')]
    [Parameter(ParameterSetName = 'HuduById')]
    [Parameter(ParameterSetName = 'HuduByName')]
    [switch]$PermanentDelete
)

$ScriptVersion = '1.0.0'
Write-Verbose "Remove-AuditCustomer.ps1 loaded (v$ScriptVersion)"

$ErrorActionPreference = 'Stop'

# ── Helpers ────────────────────────────────────────────────────────────────────

function Write-Status {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory, Position = 0)] [string]$Message,
        [ValidateSet('Info', 'Success', 'Error', 'Warning')]
        [string]$Type = 'Info'
    )
    $map = @{
        Info    = @{ Prefix = '[INFO]';    Color = 'Cyan' }
        Success = @{ Prefix = '[SUCCESS]'; Color = 'Green' }
        Error   = @{ Prefix = '[ERROR]';   Color = 'Red' }
        Warning = @{ Prefix = '[WARNING]'; Color = 'Yellow' }
    }
    Write-Host "$($map[$Type].Prefix) $Message" -ForegroundColor $map[$Type].Color
}

# ── Load config.psd1 ───────────────────────────────────────────────────────────

$script:HuduAssetLayoutId = 67
$script:HuduAssetName     = 'M365 Audit Toolkit'

$_configPath = Join-Path $PSScriptRoot '..\config.psd1'
if (Test-Path $_configPath) {
    try {
        $_config = Import-PowerShellDataFile -Path $_configPath
        if (-not $HuduApiKey  -and $_config.HuduApiKey)  { $HuduApiKey               = $_config.HuduApiKey }
        if (-not $HuduBaseUrl -and $_config.HuduBaseUrl) { $HuduBaseUrl              = $_config.HuduBaseUrl }
        if ($_config.HuduAssetLayoutId -gt 0)            { $script:HuduAssetLayoutId = $_config.HuduAssetLayoutId }
        if ($_config.HuduAssetName)                      { $script:HuduAssetName     = $_config.HuduAssetName }
    }
    catch { Write-Warning "Could not load config.psd1: $_" }
}

if (-not $HuduBaseUrl) { $HuduBaseUrl = $env:HUDU_BASE_URL }
if (-not $HuduApiKey)  { $HuduApiKey  = $env:HUDU_API_KEY }

# ── Derived state ─────────────────────────────────────────────────────────────

$usingHudu = $PSCmdlet.ParameterSetName -in 'HuduById', 'HuduByName'

if ($usingHudu -and -not $HuduApiKey) {
    throw "HUDU_API_KEY is required when using -HuduCompanyId or -HuduCompanyName. Set the environment variable, config.psd1, or pass -HuduApiKey."
}

# ── Hudu lookup ────────────────────────────────────────────────────────────────

$huduAssetId  = $null
$companyLabel = $HuduCompanyId ?? $HuduCompanyName

if ($usingHudu) {
    $huduUrl = $HuduBaseUrl.TrimEnd('/')
    $headers = @{ 'x-api-key' = $HuduApiKey }

    Write-Status "Looking up Hudu company '$companyLabel'..."

    $company = $null
    if ($HuduCompanyId) {
        if ($HuduCompanyId -match '^\d+$') {
            $company = (Invoke-RestMethod -Uri "$huduUrl/api/v1/companies/$HuduCompanyId" -Headers $headers -Method Get -ErrorAction Stop).company
        }
        else {
            $encoded = [uri]::EscapeDataString($HuduCompanyId)
            $company = @((Invoke-RestMethod -Uri "$huduUrl/api/v1/companies?slug=$encoded&page_size=1" -Headers $headers -Method Get -ErrorAction Stop).companies) | Select-Object -First 1
        }
    }
    else {
        $encoded = [uri]::EscapeDataString($HuduCompanyName)
        $company = @((Invoke-RestMethod -Uri "$huduUrl/api/v1/companies?search=$encoded&page_size=25" -Headers $headers -Method Get -ErrorAction Stop).companies) |
            Where-Object { $_.name -eq $HuduCompanyName } | Select-Object -First 1
    }

    if (-not $company) { throw "No Hudu company found for '$companyLabel'." }
    $companyLabel = $company.name
    Write-Status "Company: $companyLabel" -Type Success

    $asset = @((Invoke-RestMethod -Uri "$huduUrl/api/v1/assets?company_id=$($company.id)&asset_layout_id=$($script:HuduAssetLayoutId)&page_size=5" `
        -Headers $headers -Method Get -ErrorAction Stop).assets) | Sort-Object updated_at -Descending | Select-Object -First 1

    if (-not $asset) {
        Write-Status "No '$($script:HuduAssetName)' asset found in Hudu for '$companyLabel' — Hudu step will be skipped." -Type Warning
    }
    else {
        $huduAssetId = $asset.id
        $fieldMap    = @{}
        foreach ($f in $asset.fields) { $fieldMap[$f.label] = "$($f.value)" }

        if (-not $AppId    -and $fieldMap['Application ID']) { $AppId    = $fieldMap['Application ID'] }
        if (-not $TenantId -and $fieldMap['Tenant ID'])      { $TenantId = $fieldMap['Tenant ID'] }

        Write-Status "Found asset '$($asset.name)' (ID: $huduAssetId)" -Type Success
    }

    if (-not $AppId -or -not $TenantId) {
        throw "Could not determine AppId/TenantId from Hudu asset. Pass -AppId and -TenantId explicitly."
    }
}

# ── Summary before acting ──────────────────────────────────────────────────────

$sep = '=' * 72
Write-Host "`n$sep" -ForegroundColor Yellow
Write-Host '  365Audit Customer Removal' -ForegroundColor Yellow
Write-Host $sep -ForegroundColor Yellow
Write-Host "  Company     : $companyLabel"
Write-Host "  App ID      : $AppId"
Write-Host "  Tenant ID   : $TenantId"
if ($huduAssetId) {
    Write-Host "  Hudu asset  : will be ARCHIVED (ID: $huduAssetId)"
}
else {
    Write-Host "  Hudu asset  : none found / not applicable"
}
if ($PermanentDelete) {
    Write-Host "  Entra app   : will be PERMANENTLY DELETED (not recoverable)" -ForegroundColor Red
}
else {
    Write-Host "  Entra app   : will be soft-deleted (recoverable for 30 days)"
}
Write-Host "$sep`n" -ForegroundColor Yellow

# ── Graph connection ───────────────────────────────────────────────────────────

Write-Status 'Connecting to Microsoft Graph (browser window will open)...'
try {
    Connect-MgGraph -Scopes 'Application.ReadWrite.All' -TenantId $TenantId -NoWelcome -ErrorAction Stop
}
catch {
    if ($_.Exception.Message -match 'WithBroker|BrokerExtension|MsalCacheHelper|InteractiveBrowserCredential') {
        throw (
            "Interactive authentication failed due to a MSAL version conflict. " +
            "Open a standalone PowerShell 7 window (not an IDE terminal) and re-run, " +
            "or run: Update-Module Microsoft.Graph -Force  then restart PowerShell."
        )
    }
    throw
}
Write-Status 'Connected.' -Type Success

# ── Remove Entra app ───────────────────────────────────────────────────────────

$app = Get-MgApplication -Filter "appId eq '$AppId'" -ErrorAction Stop | Select-Object -First 1

if (-not $app) {
    Write-Status "No app registration found for AppId '$AppId' in this tenant — may have already been removed." -Type Warning
}
else {
    Write-Status "Found app: '$($app.DisplayName)' (Object ID: $($app.Id))"

    if ($PSCmdlet.ShouldProcess("App '$($app.DisplayName)' ($AppId)", 'Remove from Entra ID')) {
        Remove-MgApplication -ApplicationId $app.Id -ErrorAction Stop
        Write-Status "App '$($app.DisplayName)' soft-deleted from Entra ID." -Type Success

        if ($PermanentDelete) {
            if ($PSCmdlet.ShouldProcess("App '$($app.DisplayName)' ($AppId)", 'Permanently purge from Entra recycle bin')) {
                Start-Sleep -Seconds 3   # brief wait for soft-delete to propagate
                Remove-MgDirectoryDeletedItem -DirectoryObjectId $app.Id -ErrorAction Stop
                Write-Status "App permanently purged from Entra recycle bin." -Type Success
            }
        }
        else {
            Write-Status "App is in the Entra recycle bin — recoverable for 30 days via the Azure portal." -Type Info
        }
    }
}

# ── Archive Hudu asset ─────────────────────────────────────────────────────────

if ($huduAssetId) {
    if ($PSCmdlet.ShouldProcess("Hudu asset ID $huduAssetId for '$companyLabel'", 'Archive in Hudu')) {
        $huduUrl = $HuduBaseUrl.TrimEnd('/')
        $headers = @{ 'x-api-key' = $HuduApiKey }
        Invoke-RestMethod -Uri "$huduUrl/api/v1/assets/$huduAssetId/archive" -Headers $headers -Method Put -ErrorAction Stop | Out-Null
        Write-Status "Hudu asset archived for '$companyLabel'." -Type Success
    }
}

Write-Host ""
Write-Status "Customer removal complete for '$companyLabel'." -Type Success
