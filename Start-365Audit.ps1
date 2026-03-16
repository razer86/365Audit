<#
.SYNOPSIS
    Interactive launcher for the Microsoft 365 Audit toolkit.

.DESCRIPTION
    Presents a menu of available audit modules. Modules are loaded from local disk;
    missing modules are downloaded from GitHub as a fallback. On startup, compares
    local script versions against the GitHub version manifest and warns if any are outdated.

    Credentials can be supplied manually or retrieved automatically from Hudu using
    -HuduCompanyId (slug or numeric ID) or -HuduCompanyName (exact match).

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
    Password for the .pfx certificate file. If omitted, you will be prompted.

.PARAMETER HuduCompanyId
    Hudu company slug (alphanumeric) or numeric database ID.
    All credentials are retrieved automatically from the 'NeConnect Audit Toolkit'
    asset for that company. Requires HUDU_API_KEY in the environment.

.PARAMETER HuduCompanyName
    Exact Hudu company name (case-sensitive).
    All credentials are retrieved automatically from the 'NeConnect Audit Toolkit'
    asset for that company. Requires HUDU_API_KEY in the environment.

.PARAMETER HuduBaseUrl
    Hudu instance base URL. Falls back to HUDU_BASE_URL env var, then
    'https://neconnect.huducloud.com'. Only used with -HuduCompanyId or -HuduCompanyName.

.PARAMETER HuduApiKey
    Hudu API key. Falls back to HUDU_API_KEY env var.
    Only used with -HuduCompanyId or -HuduCompanyName.

.EXAMPLE
    .\Start-365Audit.ps1 -AppId '<guid>' -TenantId '<guid>' -CertBase64 '<base64>' -CertPassword (Read-Host -AsSecureString 'Cert Password')
    Supply all credentials on the command line.

.EXAMPLE
    .\Start-365Audit.ps1 -AppId '<guid>' -TenantId '<guid>'
    Prompts interactively for the certificate Base64 and password.

.EXAMPLE
    .\Start-365Audit.ps1 -HuduCompanyId '44706357047c'
    Fetches all credentials from Hudu using the company slug. Requires HUDU_API_KEY env var.

.EXAMPLE
    .\Start-365Audit.ps1 -HuduCompanyName 'Contoso Ltd'
    Fetches all credentials from Hudu using the exact company name. Requires HUDU_API_KEY env var.

.NOTES
    Author      : Raymond Slater
    Version     : 2.9.2
    Change Log  : See CHANGELOG.md

.LINK
    https://github.com/razer86/365Audit
#>

#Requires -Version 7.2

[CmdletBinding(DefaultParameterSetName = 'Manual')]
param (
    # ── Manual credential parameters ──────────────────────────────────────────
    [Parameter(Mandatory, ParameterSetName = 'Manual',
        HelpMessage = 'Azure AD application (client) ID. Run Setup-365AuditApp.ps1 to obtain.')]
    [string]$AppId,

    [Parameter(Mandatory, ParameterSetName = 'Manual',
        HelpMessage = 'Azure AD tenant ID (GUID or .onmicrosoft.com domain).')]
    [string]$TenantId,

    [Parameter(ParameterSetName = 'Manual',
        HelpMessage = 'Base64-encoded .pfx certificate. Omit to be prompted.')]
    [string]$CertBase64,

    [Parameter(ParameterSetName = 'Manual',
        HelpMessage = 'Password for the .pfx certificate file. Omit to be prompted.')]
    [SecureString]$CertPassword,

    # ── Hudu parameters ────────────────────────────────────────────────────────
    [Parameter(Mandatory, ParameterSetName = 'HuduById',
        HelpMessage = 'Hudu company slug or numeric ID. Credentials fetched from NeConnect Audit Toolkit asset.')]
    [string]$HuduCompanyId,

    [Parameter(Mandatory, ParameterSetName = 'HuduByName',
        HelpMessage = 'Exact Hudu company name. Credentials fetched from NeConnect Audit Toolkit asset.')]
    [string]$HuduCompanyName,

    [Parameter(ParameterSetName = 'HuduById')]
    [Parameter(ParameterSetName = 'HuduByName')]
    [string]$HuduBaseUrl = ($env:HUDU_BASE_URL ?? 'https://neconnect.huducloud.com'),

    [Parameter(ParameterSetName = 'HuduById')]
    [Parameter(ParameterSetName = 'HuduByName')]
    [string]$HuduApiKey = $env:HUDU_API_KEY,

    # ── Automation ────────────────────────────────────────────────────────────
    # Provide module numbers to skip the menu and run non-interactively.
    # The HTML summary is generated but not opened automatically.
    # Example: -Modules 1,2,3,4  or  -Modules 9
    [Parameter(ParameterSetName = 'Manual')]
    [Parameter(ParameterSetName = 'HuduById')]
    [Parameter(ParameterSetName = 'HuduByName')]
    [ValidateSet(1, 2, 3, 4, 5, 9)]
    [int[]]$Modules
)

$ScriptVersion = "2.9.2"
Write-Verbose "Start-365Audit.ps1 loaded (v$ScriptVersion)"

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# === Transcript logging ===
# Capture all console output to a temp file; moved to the audit output folder as AuditLog.txt on exit.
$_transcriptActive = $false
$_transcriptPath   = $null
$_logTempDir       = $env:TEMP ?? $env:TMPDIR ?? '/tmp'
$_transcriptPath   = Join-Path $_logTempDir "365Audit-$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
try {
    Start-Transcript -Path $_transcriptPath -UseMinimalHeader | Out-Null
    $_transcriptActive = $true
}
catch { Write-Verbose "Could not start transcript: $_" }

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
$localPath     = $PSScriptRoot
$_tempCertPath = $null   # populated after cert decode; used in finally block


# === Fetch credentials from Hudu =============================================
if ($PSCmdlet.ParameterSetName -in 'HuduById', 'HuduByName') {

    if (-not $HuduApiKey) {
        Write-Error ("HUDU_API_KEY environment variable is not set.`n" +
            "  Set it with: `$env:HUDU_API_KEY = (Read-Host 'Hudu API key')") -ErrorAction Stop
    }

    $HuduBaseUrl  = $HuduBaseUrl.TrimEnd('/')
    $huduHeaders  = @{ 'x-api-key' = $HuduApiKey; 'Content-Type' = 'application/json' }

    Write-Host "`n  Fetching credentials from Hudu..." -ForegroundColor Cyan

    # Resolve company
    $huduCompany = $null
    if ($PSCmdlet.ParameterSetName -eq 'HuduById') {
        try {
            if ($HuduCompanyId -match '^\d+$') {
                $r = Invoke-RestMethod -Uri "$HuduBaseUrl/api/v1/companies/$HuduCompanyId" `
                    -Headers $huduHeaders -Method Get -ErrorAction Stop
                $huduCompany = $r.company
            }
            else {
                $encoded = [uri]::EscapeDataString($HuduCompanyId)
                $r = Invoke-RestMethod -Uri "$HuduBaseUrl/api/v1/companies?slug=$encoded&page_size=1" `
                    -Headers $huduHeaders -Method Get -ErrorAction Stop
                $huduCompany = @($r.companies) | Select-Object -First 1
            }
        }
        catch { Write-Error "Hudu company lookup failed: $_" -ErrorAction Stop }

        if (-not $huduCompany) {
            Write-Error "No Hudu company found for ID/slug '$HuduCompanyId'." -ErrorAction Stop
        }
    }
    else {
        # HuduByName — search then require exact match
        try {
            $encoded = [uri]::EscapeDataString($HuduCompanyName)
            $r = Invoke-RestMethod -Uri "$HuduBaseUrl/api/v1/companies?search=$encoded&page_size=25" `
                -Headers $huduHeaders -Method Get -ErrorAction Stop
            $huduCompany = @($r.companies) | Where-Object { $_.name -eq $HuduCompanyName } |
                Select-Object -First 1
        }
        catch { Write-Error "Hudu company lookup failed: $_" -ErrorAction Stop }

        if (-not $huduCompany) {
            Write-Error "No Hudu company found with exact name '$HuduCompanyName'." -ErrorAction Stop
        }
    }

    Write-Host "  Company : $($huduCompany.name) (id: $($huduCompany.id))" -ForegroundColor Green

    # Find the NeConnect Audit Toolkit asset (layout ID 67)
    try {
        $assetsResult = Invoke-RestMethod `
            -Uri     "$HuduBaseUrl/api/v1/assets?company_id=$($huduCompany.id)&asset_layout_id=67&page_size=5" `
            -Headers $huduHeaders -Method Get -ErrorAction Stop
    }
    catch { Write-Error "Hudu asset lookup failed: $_" -ErrorAction Stop }

    $huduAsset = @($assetsResult.assets) | Sort-Object updated_at -Descending | Select-Object -First 1
    if (-not $huduAsset) {
        Write-Error ("No '365Audit' asset found for '$($huduCompany.name)' in Hudu.`n" +
            "  Run Setup-365AuditApp.ps1 to create the app registration and populate the asset.") -ErrorAction Stop
    }

    # Map field labels to values
    $fieldMap = @{}
    foreach ($f in $huduAsset.fields) { $fieldMap[$f.label] = "$($f.value)" }

    $AppId      = $fieldMap['Application ID']
    $TenantId   = $fieldMap['Tenant ID']
    $CertBase64 = $fieldMap['Cert Base64']
    $plainPwd   = $fieldMap['Cert Password']

    foreach ($pair in @(@('Application ID', $AppId), @('Tenant ID', $TenantId),
                        @('Cert Base64', $CertBase64), @('Cert Password', $plainPwd))) {
        if (-not $pair[1]) {
            Write-Error ("Hudu asset '$($huduAsset.name)' is missing field: $($pair[0]).`n" +
                "  Run Setup-365AuditApp.ps1 to repopulate the asset.") -ErrorAction Stop
        }
    }

    $CertPassword = ConvertTo-SecureString $plainPwd -AsPlainText -Force
    Write-Host "  Credentials loaded from Hudu asset: $($huduAsset.name)" -ForegroundColor Green
}


# === System clock drift check ================================================
# Certificate-based auth fails when the local clock differs from Microsoft's servers
# by more than ~5 minutes. Warn at >60 s, stop at >300 s.
try {
    $response    = Invoke-WebRequest -Uri 'https://login.microsoftonline.com' -Method Head `
                       -UseBasicParsing -TimeoutSec 10 -ErrorAction Stop
    $serverUtc   = [datetime]::ParseExact(
        $response.Headers['Date'],
        'ddd, dd MMM yyyy HH:mm:ss \G\M\T',
        [System.Globalization.CultureInfo]::InvariantCulture,
        [System.Globalization.DateTimeStyles]::AssumeUniversal).ToUniversalTime()
    $driftSec = [math]::Abs(([datetime]::UtcNow - $serverUtc).TotalSeconds)

    if ($driftSec -gt 300) {
        Write-Error ("System clock is $([math]::Round($driftSec))s out of sync with Microsoft servers " +
            "(limit: 300s). Certificate authentication will fail — correct the system time and retry.") -ErrorAction Stop
    }
    elseif ($driftSec -gt 60) {
        Write-Warning "System clock is $([math]::Round($driftSec))s out of sync with Microsoft servers. Authentication may fail if drift exceeds 300s."
    }
    else {
        Write-Verbose "Clock drift: $([math]::Round($driftSec))s (OK)."
    }
}
catch [System.Net.WebException] {
    Write-Warning "Could not check clock drift (no network): $_"
}
catch [System.Management.Automation.ErrorRecord] {
    # Already handled above (Stop error from drift > 300s)
    throw
}
catch {
    Write-Warning "Clock drift check skipped: $_"
}


# === Decode base64 cert to a temp .pfx (deleted on exit) =====================
if (-not $CertBase64) {
    $CertBase64 = Read-Host 'Paste certificate Base64'
}

# Validate the base64 string decodes correctly before writing anything to disk.
try {
    $certBytes = [Convert]::FromBase64String($CertBase64)
}
catch {
    Write-Error "Certificate Base64 is invalid and could not be decoded. Verify the value copied from Hudu is complete." -ErrorAction Stop
}

$_tempDir      = $env:TEMP ?? $env:TMPDIR ?? '/tmp'
$_tempCertPath = Join-Path $_tempDir "365Audit-$(New-Guid).pfx"
[System.IO.File]::WriteAllBytes($_tempCertPath, $certBytes)
$CertFilePath  = $_tempCertPath
Write-Verbose "Certificate decoded from base64 to temp file: $_tempCertPath"

# Prompt for cert password if not supplied (Hudu path pre-populates this)
if (-not $CertPassword) {
    $CertPassword = Read-Host 'Cert Password' -AsSecureString
}

# Check certificate expiry and warn if renewal is needed within 30 days.
# EphemeralKeySet keeps the private key in memory only (Windows).
# Linux/macOS do not support EphemeralKeySet — use Exportable|PersistKeySet instead.
$keyStorageFlags = if ($IsLinux -or $IsMacOS) {
    [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::Exportable -bor
    [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::PersistKeySet
} else {
    [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::EphemeralKeySet
}
$certDaysRemaining = -1
try {
    $certObj = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new(
        $certBytes,
        $CertPassword,
        $keyStorageFlags
    )
    $certDaysRemaining = ($certObj.NotAfter - (Get-Date)).Days
    if ($certDaysRemaining -le 0) {
        Write-Error ("The audit certificate EXPIRED $([math]::Abs($certDaysRemaining)) day(s) ago. " +
            "Authentication will fail. Run Setup-365AuditApp.ps1 -Force (requires interactive Global Admin login) to renew.") -ErrorAction Stop
    }
    elseif ($certDaysRemaining -le 30) {
        Write-Warning "The audit certificate expires in $certDaysRemaining day(s) ($($certObj.NotAfter.ToString('yyyy-MM-dd'))). Run Setup-365AuditApp.ps1 -Force soon to renew."
    }
    else {
        Write-Verbose "Certificate valid until $($certObj.NotAfter.ToString('yyyy-MM-dd')) ($certDaysRemaining days remaining)."
    }
    $certObj.Dispose()
}
catch {
    Write-Warning "Could not read certificate expiry: $_"
}

# Expose app credentials so dot-sourced modules can access them via Get-Variable.
$AuditAppId        = $AppId
$AuditTenantId     = $TenantId
$AuditCertFilePath = $CertFilePath
$AuditCertPassword = $CertPassword
Write-Verbose "Audit credentials set in launcher scope (AppId=$AuditAppId, TenantId=$AuditTenantId, CertFilePath=$AuditCertFilePath, CertPassword=$(if ($AuditCertPassword) {'set'} else {'not set'}))"

# === Drop any existing sessions from a prior run in this PS session ===
# The connect helpers skip reconnecting if a session is already active, so we must
# disconnect before setting new credentials — otherwise the wrong tenant is audited.
if (Get-MgContext -ErrorAction SilentlyContinue) {
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    Write-Verbose "Disconnected existing Microsoft Graph session."
}
if (Get-ConnectionInformation -ErrorAction SilentlyContinue | Where-Object { $_.State -eq 'Connected' }) {
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
    Write-Verbose "Disconnected existing Exchange Online session."
}
try {
    Disconnect-PnPOnline -ErrorAction SilentlyContinue | Out-Null
    Write-Verbose "Disconnected existing SharePoint Online session."
} catch {}

# === Check for updates ===
Invoke-VersionCheck -ScriptRoot $PSScriptRoot

# === Define Menu Items ===
$menu = @{
    1 = @{ Name = "Microsoft Entra Audit";      Script = @("Invoke-EntraAudit.ps1") }
    2 = @{ Name = "Exchange Online Audit";      Script = @("Invoke-ExchangeAudit.ps1") }
    3 = @{ Name = "SharePoint Online Audit";    Script = @("Invoke-SharePointAudit.ps1") }
    4 = @{ Name = "Mail Security Audit";        Script = @("Invoke-MailSecurityAudit.ps1") }
    5 = @{ Name = "Intune / Endpoint Audit";    Script = @("Invoke-IntuneAudit.ps1") }
    9 = @{ Name = "Run All Modules (1-5)";      Script = @("Invoke-EntraAudit.ps1", "Invoke-ExchangeAudit.ps1", "Invoke-SharePointAudit.ps1", "Invoke-MailSecurityAudit.ps1", "Invoke-IntuneAudit.ps1") }
    0 = @{ Name = "Exit";                       Script = $null }
}

# === Select Modules ===
if ($Modules) {
    # Non-interactive: use the provided module list directly
    $selectedIndexes = $Modules
}
else {
    # Interactive: display menu and prompt
    Write-Host "`n╔════════════════════════════════════╗"
    Write-Host "║    Microsoft 365 Audit Launcher    ║"
    Write-Host "╚════════════════════════════════════╝"

    foreach ($key in ($menu.Keys | Sort-Object { [int]$_ })) {
        Write-Host "$key. $($menu[$key].Name)"
    }

    $selection = Read-Host "`nSelect one or more modules (comma separated, e.g. 1,2)"
    if ($selection -eq "0") {
        Write-Host "Exiting. Goodbye!"
        return
    }

    $selectedIndexes = $selection -split "," |
        ForEach-Object { $_.Trim() } |
        Where-Object    { $_ -match '^\d+$' } |
        ForEach-Object  { [int]$_ }
}

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
        $summaryParams = @{ AuditFolder = $auditContext.OutputPath }
        if ($Modules) { $summaryParams['NoOpen'] = $true }
        if ($certDaysRemaining -ge 0 -and $certDaysRemaining -le 30) { $summaryParams['CertExpiryDays'] = $certDaysRemaining }
        & $summaryScript @summaryParams
    }
    else {
        Write-Warning "No audit output context found — summary report skipped."
    }
}
finally {
    if ($_tempCertPath -and (Test-Path $_tempCertPath)) {
        Remove-Item $_tempCertPath -Force -ErrorAction SilentlyContinue
        Write-Verbose "Temp certificate file removed: $_tempCertPath"
    }
    if ($_transcriptActive) {
        try { Stop-Transcript | Out-Null } catch {}
        $logCtx = try { Initialize-AuditOutput } catch { $null }
        if ($logCtx -and $_transcriptPath -and (Test-Path $_transcriptPath -ErrorAction SilentlyContinue)) {
            $logDir  = Join-Path $logCtx.OutputPath 'Logs'
            New-Item -ItemType Directory -Path $logDir -Force -ErrorAction SilentlyContinue | Out-Null
            $logDest = Join-Path $logDir 'AuditLog.txt'
            Move-Item -Path $_transcriptPath -Destination $logDest -Force -ErrorAction SilentlyContinue
            Write-Verbose "Audit log saved: $logDest"
        }
    }
    if (Get-MgContext -ErrorAction SilentlyContinue) {
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    }
    if (Get-ConnectionInformation -ErrorAction SilentlyContinue | Where-Object { $_.State -eq 'Connected' }) {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
    }
    try { Disconnect-PnPOnline -ErrorAction SilentlyContinue | Out-Null } catch {}
}
