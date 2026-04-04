function Remove-AuditCustomer {
    <#
    .SYNOPSIS
        Removes a customer's 365Audit app registration from Entra ID and archives
        the corresponding asset in Hudu.

    .DESCRIPTION
        Used when a customer offboards or when you need to fully reset a customer's
        365Audit configuration.

        When run with -HuduCompanyId or -HuduCompanyName, the function:
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
        Hudu instance base URL. Falls back to audit context (Set-AuditContext), then
        the HUDU_BASE_URL environment variable.

    .PARAMETER HuduApiKey
        Hudu API key. Falls back to audit context (Set-AuditContext), then the
        HUDU_API_KEY environment variable.

    .PARAMETER HuduAssetLayoutId
        Asset layout ID for the audit toolkit credential assets. Default: 67.

    .PARAMETER HuduAssetName
        Display name of the Hudu asset to look up. Default: 'M365 Audit Toolkit'.

    .PARAMETER PermanentDelete
        Permanently purge the app registration from the Entra recycle bin immediately
        after soft-deleting it. Cannot be undone.

    .EXAMPLE
        Remove-AuditCustomer -HuduCompanyId 'a1b2c3d4e5f6'
        Looks up the customer in Hudu, removes the Entra app, and archives the Hudu asset.

    .EXAMPLE
        Remove-AuditCustomer -HuduCompanyId 'a1b2c3d4e5f6' -PermanentDelete
        Removes and immediately purges the app -- use when certain the customer won't be re-onboarded.

    .EXAMPLE
        Remove-AuditCustomer -AppId '00000000-0000-0000-0000-000000000000' -TenantId '00000000-0000-0000-0000-000000000000'
        Removes the Entra app directly without Hudu interaction.

    .EXAMPLE
        Remove-AuditCustomer -HuduCompanyId 'a1b2c3d4e5f6' -WhatIf
        Preview what would be removed without making any changes.
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High', DefaultParameterSetName = 'Manual')]
    param(
        # -- Manual credential parameters ----------------------------------------
        [Parameter(Mandatory, ParameterSetName = 'Manual',
            HelpMessage = 'Azure AD application (client) ID of the app registration to remove.')]
        [string]$AppId,

        [Parameter(Mandatory, ParameterSetName = 'Manual',
            HelpMessage = 'Azure AD tenant ID (GUID or .onmicrosoft.com domain).')]
        [string]$TenantId,

        # -- Hudu parameters -----------------------------------------------------
        [Parameter(Mandatory, ParameterSetName = 'HuduById',
            HelpMessage = 'Hudu company slug (12-character hex) or numeric ID.')]
        [string]$HuduCompanyId,

        [Parameter(Mandatory, ParameterSetName = 'HuduByName',
            HelpMessage = 'Exact Hudu company name.')]
        [string]$HuduCompanyName,

        [Parameter(ParameterSetName = 'HuduById')]
        [Parameter(ParameterSetName = 'HuduByName')]
        [string]$HuduBaseUrl,

        [Parameter(ParameterSetName = 'HuduById')]
        [Parameter(ParameterSetName = 'HuduByName')]
        [string]$HuduApiKey,

        [Parameter(ParameterSetName = 'HuduById')]
        [Parameter(ParameterSetName = 'HuduByName')]
        [int]$HuduAssetLayoutId = 67,

        [Parameter(ParameterSetName = 'HuduById')]
        [Parameter(ParameterSetName = 'HuduByName')]
        [string]$HuduAssetName = 'M365 Audit Toolkit',

        # -- Options -------------------------------------------------------------
        [Parameter(ParameterSetName = 'Manual')]
        [Parameter(ParameterSetName = 'HuduById')]
        [Parameter(ParameterSetName = 'HuduByName')]
        [switch]$PermanentDelete
    )

    # ── Resolve Hudu credentials from context / environment ───────────────────

    $usingHudu = $PSCmdlet.ParameterSetName -in 'HuduById', 'HuduByName'

    if ($usingHudu) {
        if (-not $HuduBaseUrl -or -not $HuduApiKey) {
            $ctx = Get-AuditContext -NoThrow
            if ($ctx) {
                if (-not $HuduBaseUrl) { $HuduBaseUrl = $ctx.HuduBaseUrl }
                if (-not $HuduApiKey)  { $HuduApiKey  = $ctx.HuduApiKey }
            }
        }

        if (-not $HuduBaseUrl) { $HuduBaseUrl = $env:HUDU_BASE_URL }
        if (-not $HuduApiKey)  { $HuduApiKey  = $env:HUDU_API_KEY }

        if (-not $HuduApiKey) {
            throw "Hudu API key is required when using -HuduCompanyId or -HuduCompanyName. " +
                  "Set via Set-AuditContext, the HUDU_API_KEY environment variable, or pass -HuduApiKey."
        }
        if (-not $HuduBaseUrl) {
            throw "Hudu base URL is required when using -HuduCompanyId or -HuduCompanyName. " +
                  "Set via Set-AuditContext, the HUDU_BASE_URL environment variable, or pass -HuduBaseUrl."
        }
    }

    # ── Hudu lookup ───────────────────────────────────────────────────────────

    $huduAssetId  = $null
    $companyLabel = $HuduCompanyId ?? $HuduCompanyName

    if ($usingHudu) {
        Write-Host "[INFO] Looking up Hudu company '$companyLabel'..." -ForegroundColor Cyan

        $company = $null
        if ($HuduCompanyId) {
            if ($HuduCompanyId -match '^\d+$') {
                $result = Invoke-HuduRequest -Endpoint "api/v1/companies/$HuduCompanyId" `
                    -HuduBaseUrl $HuduBaseUrl -HuduApiKey $HuduApiKey
                $company = if ($result.PSObject.Properties.Name -contains 'company') { $result.company } else { $result }
            }
            else {
                $encoded = [uri]::EscapeDataString($HuduCompanyId)
                $result  = Invoke-HuduRequest -Endpoint "api/v1/companies?slug=$encoded&page_size=1" `
                    -HuduBaseUrl $HuduBaseUrl -HuduApiKey $HuduApiKey
                $company = @(if ($result.PSObject.Properties.Name -contains 'companies') { $result.companies } else { $result }) |
                    Select-Object -First 1
            }
        }
        else {
            $encoded = [uri]::EscapeDataString($HuduCompanyName)
            $result  = Invoke-HuduRequest -Endpoint "api/v1/companies?search=$encoded&page_size=25" `
                -HuduBaseUrl $HuduBaseUrl -HuduApiKey $HuduApiKey
            $company = @(if ($result.PSObject.Properties.Name -contains 'companies') { $result.companies } else { $result }) |
                Where-Object { $_.name -eq $HuduCompanyName } | Select-Object -First 1
        }

        if (-not $company) { throw "No Hudu company found for '$companyLabel'." }
        $companyLabel = $company.name
        Write-Host "[SUCCESS] Company: $companyLabel" -ForegroundColor Green

        $assetResult = Invoke-HuduRequest -Endpoint "api/v1/assets?company_id=$($company.id)&asset_layout_id=$HuduAssetLayoutId&page_size=5" `
            -HuduBaseUrl $HuduBaseUrl -HuduApiKey $HuduApiKey
        $assets = @(if ($assetResult.PSObject.Properties.Name -contains 'assets') { $assetResult.assets } else { $assetResult })
        $asset  = $assets | Sort-Object updated_at -Descending | Select-Object -First 1

        if (-not $asset) {
            Write-Warning "No '$HuduAssetName' asset found in Hudu for '$companyLabel' -- Hudu step will be skipped."
        }
        else {
            $huduAssetId = $asset.id
            $fieldMap    = @{}
            foreach ($f in $asset.fields) { $fieldMap[$f.label] = "$($f.value)" }

            if (-not $AppId    -and $fieldMap['Application ID']) { $AppId    = $fieldMap['Application ID'] }
            if (-not $TenantId -and $fieldMap['Tenant ID'])      { $TenantId = $fieldMap['Tenant ID'] }

            Write-Host "[SUCCESS] Found asset '$($asset.name)' (ID: $huduAssetId)" -ForegroundColor Green
        }

        if (-not $AppId -or -not $TenantId) {
            throw "Could not determine AppId/TenantId from Hudu asset. Pass -AppId and -TenantId explicitly."
        }
    }

    # ── Summary before acting ─────────────────────────────────────────────────

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

    # ── Graph connection ──────────────────────────────────────────────────────

    Write-Host "[INFO] Connecting to Microsoft Graph (browser window will open)..." -ForegroundColor Cyan
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
    Write-Host "[SUCCESS] Connected." -ForegroundColor Green

    # ── Remove Entra app ──────────────────────────────────────────────────────

    $app = Get-MgApplication -Filter "appId eq '$AppId'" -ErrorAction Stop | Select-Object -First 1

    if (-not $app) {
        Write-Warning "No app registration found for AppId '$AppId' in this tenant -- may have already been removed."
    }
    else {
        Write-Host "[INFO] Found app: '$($app.DisplayName)' (Object ID: $($app.Id))" -ForegroundColor Cyan

        if ($PSCmdlet.ShouldProcess("App '$($app.DisplayName)' ($AppId)", 'Remove from Entra ID')) {
            Remove-MgApplication -ApplicationId $app.Id -ErrorAction Stop
            Write-Host "[SUCCESS] App '$($app.DisplayName)' soft-deleted from Entra ID." -ForegroundColor Green

            if ($PermanentDelete) {
                if ($PSCmdlet.ShouldProcess("App '$($app.DisplayName)' ($AppId)", 'Permanently purge from Entra recycle bin')) {
                    Start-Sleep -Seconds 3   # brief wait for soft-delete to propagate
                    Remove-MgDirectoryDeletedItem -DirectoryObjectId $app.Id -ErrorAction Stop
                    Write-Host "[SUCCESS] App permanently purged from Entra recycle bin." -ForegroundColor Green
                }
            }
            else {
                Write-Host "[INFO] App is in the Entra recycle bin -- recoverable for 30 days via the Azure portal." -ForegroundColor Cyan
            }
        }
    }

    # ── Archive Hudu asset ────────────────────────────────────────────────────

    if ($huduAssetId) {
        if ($PSCmdlet.ShouldProcess("Hudu asset ID $huduAssetId for '$companyLabel'", 'Archive in Hudu')) {
            Invoke-HuduRequest -Endpoint "api/v1/assets/$huduAssetId/archive" -Method PUT `
                -HuduBaseUrl $HuduBaseUrl -HuduApiKey $HuduApiKey | Out-Null
            Write-Host "[SUCCESS] Hudu asset archived for '$companyLabel'." -ForegroundColor Green
        }
    }

    Write-Host ""
    Write-Host "[SUCCESS] Customer removal complete for '$companyLabel'." -ForegroundColor Green
}
