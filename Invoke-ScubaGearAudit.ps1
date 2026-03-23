<#
.SYNOPSIS
    Runs the CISA ScubaGear M365 Foundations Benchmark assessment against the current tenant.

.DESCRIPTION
    Obtains audit credentials and imports the certificate into the current user's certificate
    store (required by ScubaGear), then spawns a Windows PowerShell 5.1 subprocess to install
    or update ScubaGear and execute Invoke-SCuBA.  Using a separate WinPS 5.1 process avoids
    module-version conflicts with the PS 7 modules already loaded in the 365Audit session.
    The certificate is removed from the store when the assessment finishes.

    Output is written to:
        <AuditFolder>\Raw\ScubaGear_<timestamp>\
            BaselineReports.html      — ScubaGear's own HTML report
            ScubaResults_<uuid>.json  — consolidated JSON (ingested by Generate-AuditSummary.ps1)
            ScubaResults.csv          — flat CSV of all controls
            ActionPlan.csv            — failing Shall controls with blank remediation fields
            IndividualReports\        — per-product HTML + JSON files

    Generate-AuditSummary.ps1 automatically detects the ScubaGear_* folder and adds
    failing/warning controls to the action items list and a CIS Baseline section to the report.

.NOTES
    Author      : Raymond Slater
    Version     : 1.3.3
    Change Log  : See CHANGELOG.md

    Prerequisites (one-time, per app registration):
      - App registration must have the additional Graph/SPO/EXO permissions listed in
        the ScubaGear non-interactive prerequisites:
        https://github.com/cisagov/ScubaGear/blob/main/docs/prerequisites/noninteractive.md
      - The service principal must be assigned the Global Reader Entra role.
      - Power Platform assessment requires an interactive one-time setup step and is
        excluded from this script's default product list.

.LINK
    https://github.com/razer86/365Audit
    https://github.com/cisagov/ScubaGear
#>

#Requires -Version 7.2

param (
    [string]$AuditFolder,
    [switch]$DevMode = $false
)

if (-not $DevMode -and $MyInvocation.InvocationName -eq $MyInvocation.MyCommand.Name) {
    Write-Error "This script must be run from the 365Audit launcher. Use -DevMode for development." -ErrorAction Stop
}

$ScriptVersion = "1.3.3"
Write-Verbose "Invoke-ScubaGearAudit.ps1 loaded (v$ScriptVersion)"

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# ─── Verify Windows PowerShell 5.1 is available ───────────────────────────────
$_ps51Path = (Get-Command powershell.exe -ErrorAction SilentlyContinue)?.Source
if (-not $_ps51Path) {
    Add-AuditIssue -Severity 'Warning' -Section 'ScubaGear' -Collector 'Prerequisites' `
        -Description 'powershell.exe (Windows PowerShell 5.1) not found — ScubaGear requires WinPS 5.1.' `
        -Action 'Ensure Windows PowerShell 5.1 is installed on this machine.'
    Write-Warning "powershell.exe not found. ScubaGear requires Windows PowerShell 5.1."
    return
}

# ─── Step 1 — Get audit context and credentials ───────────────────────────────
Write-Progress -Id 1 -Activity 'ScubaGear Audit' -Status 'Connecting to Microsoft Graph...' -PercentComplete 10

$_ctx = Initialize-AuditOutput

$_appId    = Get-Variable -Name AuditAppId        -ValueOnly -ErrorAction SilentlyContinue
$_certPath = Get-Variable -Name AuditCertFilePath -ValueOnly -ErrorAction SilentlyContinue
$_certPwd  = Get-Variable -Name AuditCertPassword -ValueOnly -ErrorAction SilentlyContinue

if (-not ($_appId -and $_certPath -and $_certPwd)) {
    Write-Error "Audit credentials (AuditAppId / AuditCertFilePath / AuditCertPassword) not found in scope. Run from Start-365Audit.ps1." -ErrorAction Stop
}

# Resolve the initial .onmicrosoft.com domain — ScubaGear requires -Organization in domain form
$_initialDomain = (Get-GraphOrganizationSafe).VerifiedDomains |
    Where-Object { $_.IsInitial -eq $true } |
    Select-Object -ExpandProperty Name -First 1

if (-not $_initialDomain) {
    Add-AuditIssue -Severity 'Warning' -Section 'ScubaGear' -Collector 'Domain resolution' `
        -Description 'Could not determine the initial .onmicrosoft.com domain for the tenant.'
    Write-Error "Could not resolve initial .onmicrosoft.com domain." -ErrorAction Stop
}

Write-Verbose "ScubaGear -Organization: $_initialDomain"

# ─── Step 2 — Import certificate into CurrentUser\My store ───────────────────
# The cert store is shared between PS 7 and PS 5.1, so the subprocess can use it by thumbprint.
Write-Progress -Id 1 -Activity 'ScubaGear Audit' -Status 'Importing certificate for ScubaGear...' -PercentComplete 25

$_thumbprint      = $null
$_scubaTempScript = $null

try {
    $_certBytes = [System.IO.File]::ReadAllBytes($_certPath)
    $_bstr      = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($_certPwd)
    $_plainPwd  = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($_bstr)
    [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($_bstr)

    # PersistKeySet is required so the cert store retains the private key across processes
    $_certImport = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new(
        $_certBytes, $_plainPwd,
        [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::PersistKeySet -bor
        [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::Exportable)
    $_plainPwd  = $null
    $_certBytes = $null

    $_store = [System.Security.Cryptography.X509Certificates.X509Store]::new('My', 'CurrentUser')
    $_store.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)
    $_store.Add($_certImport)
    $_store.Close()
    $_thumbprint = $_certImport.Thumbprint
    $_certImport.Dispose()
    Write-Verbose "Certificate imported (Thumbprint: $_thumbprint)"
}
catch {
    Add-AuditIssue -Severity 'Warning' -Section 'ScubaGear' -Collector 'Certificate import' `
        -Description "Failed to import certificate into Cert:\CurrentUser\My: $($_.Exception.Message)"
    Write-Error "Certificate import failed: $($_.Exception.Message)" -ErrorAction Stop
}

# ─── Step 3 — Write runner script and spawn Windows PowerShell 5.1 ───────────
# ScubaGear and its module-version dependencies must run in a fresh WinPS 5.1 process.
# All values are passed as named parameters to the temp script — no string interpolation needed.
Write-Progress -Id 1 -Activity 'ScubaGear Audit' -Status 'Launching ScubaGear in Windows PowerShell 5.1...' -PercentComplete 35

try {
    $_scubaTempScript = [System.IO.Path]::Combine(
        [System.IO.Path]::GetTempPath(),
        "ScubaGearRun_$([System.Guid]::NewGuid().ToString('N')).ps1")

    @'
param (
    [string]$AppId,
    [string]$Thumbprint,
    [string]$Organization,
    [string]$OutPath
)

$ErrorActionPreference = 'Continue'
$DebugPreference       = 'SilentlyContinue'

# Replace the inherited PS 7 PSModulePath with a clean WinPS 5.1 module path.
# PS 7 does not include Documents\WindowsPowerShell\Modules in its path, so filtering
# alone leaves Install-Module installing to a path that Import-Module cannot find.
# We explicitly build the three standard WinPS 5.1 locations and ensure the user
# folder exists so Install-Module has a writable destination.
$_userModPath = Join-Path $env:USERPROFILE 'Documents\WindowsPowerShell\Modules'
if (-not (Test-Path $_userModPath)) {
    New-Item -ItemType Directory -Path $_userModPath -Force | Out-Null
}
$env:PSModulePath = @(
    $_userModPath,
    'C:\Program Files\WindowsPowerShell\Modules',
    'C:\Windows\System32\WindowsPowerShell\v1.0\Modules'
) -join ';'

# Install / update ScubaGear
$installed = Get-Module ScubaGear -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1
if (-not $installed) {
    Write-Host "  ScubaGear not installed. Installing from PSGallery..." -ForegroundColor Cyan
    Install-Module ScubaGear -Force -AllowClobber -Scope CurrentUser -Repository PSGallery
    $installed = Get-Module ScubaGear -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1
    Write-Host "  ScubaGear $($installed.Version) installed." -ForegroundColor Green
} else {
    try {
        $latest = (Find-Module ScubaGear -Repository PSGallery -ErrorAction Stop).Version
        if ([version]$latest -gt [version]$installed.Version) {
            Write-Host "  Updating ScubaGear $($installed.Version) -> $latest..." -ForegroundColor Cyan
            Update-Module ScubaGear -Force -ErrorAction Stop
            Write-Host "  ScubaGear updated to $latest." -ForegroundColor Green
        } else {
            Write-Host "  ScubaGear $($installed.Version) is current." -ForegroundColor Gray
        }
    } catch {
        Write-Warning "Could not check for ScubaGear updates: $_"
    }
}

Import-Module ScubaGear -Force -ErrorAction Stop 5>$null 6>$null

try   { Initialize-SCuBA -ErrorAction Stop 5>$null 6>$null }
catch { Write-Warning "Initialize-SCuBA: $_" }

# Clear any stale auth sessions / MSAL token cache before ScubaGear connects.
# A cached Graph session or on-disk token from a previous tenant will cause
# "Authentication needed. Please call Connect-MgGraph" inside ScubaGear even
# when valid cert-based credentials are supplied.  Pattern from Galvnyz/M365-Assess.
try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch {}
try { Remove-Item "$env:USERPROFILE\.graph" -Recurse -Force -ErrorAction SilentlyContinue } catch {}
try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch {}

# OPA strips the drive letter from absolute paths when the PowerShell working directory is
# on a different drive (e.g. D:\) than the ScubaGear rego files (C:\Users\...\ScubaGear\).
# Setting the working directory to the same drive as the module install resolves this.
Set-Location $env:USERPROFILE

Invoke-SCuBA `
    -AppID                 $AppId `
    -CertificateThumbprint $Thumbprint `
    -Organization          $Organization `
    -OutPath               $OutPath `
    -OutFolderName         'ScubaGear' `
    -ProductNames          'aad', 'defender', 'exo', 'sharepoint', 'teams' `
    -M365Environment       'commercial' `
    -SilenceBODWarnings `
    -Quiet
'@ | Set-Content -Path $_scubaTempScript -Encoding UTF8

    Write-Host "  Launching ScubaGear in Windows PowerShell 5.1 (this may take several minutes)..." -ForegroundColor Cyan
    Write-Progress -Id 1 -Activity 'ScubaGear Audit' -Completed

    $_rawOutPath = [System.IO.Path]::GetFullPath($_ctx.RawOutputPath)

    $proc = Start-Process -FilePath 'powershell.exe' `
        -ArgumentList @(
            '-NoProfile', '-ExecutionPolicy', 'Bypass',
            '-File', $_scubaTempScript,
            '-AppId',        $_appId,
            '-Thumbprint',   $_thumbprint,
            '-Organization', $_initialDomain,
            '-OutPath',      $_rawOutPath
        ) `
        -Wait -PassThru -NoNewWindow -ErrorAction Stop

    if ($proc.ExitCode -ne 0) {
        Add-AuditIssue -Severity 'Warning' -Section 'ScubaGear' -Collector 'Invoke-SCuBA' `
            -Description "ScubaGear subprocess exited with code $($proc.ExitCode)." `
            -Action 'Review prerequisites: https://github.com/cisagov/ScubaGear/blob/main/docs/prerequisites/noninteractive.md'
        Write-Warning "ScubaGear subprocess exited with code $($proc.ExitCode)"
    } else {
        Write-Progress -Id 1 -Activity 'ScubaGear Audit' -Status 'Assessment complete.' -PercentComplete 90
        Write-Host "  ScubaGear assessment complete. Output: $_rawOutPath\ScubaGear_*" -ForegroundColor Green
    }
}
catch {
    Add-AuditIssue -Severity 'Warning' -Section 'ScubaGear' -Collector 'Invoke-SCuBA' `
        -Description $_.Exception.Message `
        -Action 'Review ScubaGear prerequisites: https://github.com/cisagov/ScubaGear/blob/main/docs/prerequisites/noninteractive.md'
    Write-Warning "ScubaGear assessment failed: $($_.Exception.Message)"
}
finally {
    # ─── Clean up temp script ─────────────────────────────────────────────────
    if ($_scubaTempScript -and (Test-Path $_scubaTempScript)) {
        Remove-Item $_scubaTempScript -Force -ErrorAction SilentlyContinue
    }

    # ─── Remove certificate from store ───────────────────────────────────────
    if ($_thumbprint) {
        Write-Progress -Id 1 -Activity 'ScubaGear Audit' -Status 'Cleaning up certificate...' -PercentComplete 95
        try {
            $_cleanStore = [System.Security.Cryptography.X509Certificates.X509Store]::new('My', 'CurrentUser')
            $_cleanStore.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)
            $_certToRemove = $_cleanStore.Certificates | Where-Object { $_.Thumbprint -eq $_thumbprint }
            if ($_certToRemove) {
                $_cleanStore.Remove($_certToRemove)
                Write-Verbose "Certificate removed from Cert:\CurrentUser\My."
            }
            $_cleanStore.Close()
        }
        catch {
            Write-Warning "Failed to remove certificate from store: $($_.Exception.Message)"
        }
    }
}

Write-Progress -Id 1 -Activity 'ScubaGear Audit' -Completed
