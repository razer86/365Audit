<#
.SYNOPSIS
    One-time tenant setup for the 365Audit toolkit.

.DESCRIPTION
    First run  : Creates the 'NeConnect MSA Audit Toolkit' app registration with all
                 required permissions (Graph, Exchange, SharePoint), grants admin consent,
                 generates a self-signed certificate (.pfx), and prints all credentials
                 (AppId, TenantId, CertBase64, CertPassword) for storage in Hudu.

    Subsequent : Checks the existing certificate's expiry. If expiring within 30 days
                 (or -Force is used), generates and uploads a new certificate.

    At audit runtime, techs provide -AppId, -TenantId, -CertBase64, and -CertPassword
    to Start-365Audit.ps1. The script decodes the Base64 cert to a temporary .pfx,
    uses it for all module connections, and deletes the temp file on exit.

.PARAMETER AppName
    Display name for the Azure app registration.
    Default: 'NeConnect MSA Audit Toolkit'

.PARAMETER CertExpiryYears
    Validity period for the generated certificate (1–5 years). Default: 2.

.PARAMETER Force
    Generate a new certificate even when the existing one is not near expiry.

.PARAMETER AppId
    Existing app registration client ID. When combined with -TenantId, -CertBase64,
    and -CertPassword, skips the browser login and renews the certificate using the
    app's own credentials (requires Application.ReadWrite.OwnedBy on the app).
    Use this for fully automated cert renewal — no Global Admin login required.

.PARAMETER TenantId
    Tenant ID for non-interactive cert renewal. Required with -AppId.

.PARAMETER CertBase64
    Base64-encoded current .pfx for non-interactive cert renewal. Required with -AppId.

.PARAMETER CertPassword
    Password for the current .pfx. Required with -AppId.

.PARAMETER HuduCompanyId
    Hudu company slug (alphanumeric) or numeric ID. When provided, credentials are pushed
    to the matching Hudu company's 'NeConnect Audit Toolkit' asset automatically without
    prompting. Requires HuduApiKey in config.psd1 (or supplied via -HuduApiKey).

.PARAMETER HuduCompanyName
    Exact Hudu company name. Alternative to -HuduCompanyId for pre-specifying the company.

.PARAMETER HuduBaseUrl
    Hudu instance base URL. Falls back to config.psd1 in the script root, then
    'https://neconnect.huducloud.com'.

.PARAMETER HuduApiKey
    Hudu API key. Falls back to config.psd1 in the script root.

.EXAMPLE
    .\Setup-365AuditApp.ps1
    Interactive setup in the customer's tenant.

.EXAMPLE
    .\Setup-365AuditApp.ps1 -Force
    Force certificate renewal even when the current certificate is healthy.

.NOTES
    Author      : Raymond Slater
    Version     : 2.12.0
    Change Log  : See CHANGELOG.md

.LINK
    https://github.com/razer86/365Audit
#>

#Requires -Version 7.4

[CmdletBinding(SupportsShouldProcess)]
param (
    [string]$AppName = 'NeConnect MSA Audit Toolkit',

    [ValidateRange(1, 5)]
    [int]$CertExpiryYears = 2,

    [switch]$Force,

    # ── Non-interactive cert renewal (optional) ────────────────────────────────
    # When all four are supplied, the browser login is skipped entirely.
    [string]$AppId,
    [string]$TenantId,
    [string]$CertBase64,
    [SecureString]$CertPassword,

    # ── Hudu integration (optional) ────────────────────────────────────────────
    [string]$HuduCompanyId,
    [string]$HuduCompanyName,
    [string]$HuduBaseUrl,
    [string]$HuduApiKey
)

$ScriptVersion         = "2.12.0"
$ErrorActionPreference = 'Stop'
$ProgressPreference    = 'SilentlyContinue'
Write-Verbose "Setup-365AuditApp.ps1 loaded (v$ScriptVersion)"

# Load config.psd1 from the script root — fallback for HuduApiKey / HuduBaseUrl.
# Explicit command-line parameters always take precedence over config file values.
$_configPath = Join-Path $PSScriptRoot 'config.psd1'
$script:HuduAssetLayoutId = 67                  # default; overridden by config.psd1
$script:HuduAssetName     = 'M365 Audit Toolkit'  # default; overridden by config.psd1
if (Test-Path $_configPath) {
    try {
        $_config = Import-PowerShellDataFile -Path $_configPath
        if (-not $HuduApiKey  -and $_config.HuduApiKey)   { $HuduApiKey                 = $_config.HuduApiKey }
        if (-not $HuduBaseUrl -and $_config.HuduBaseUrl)  { $HuduBaseUrl                = $_config.HuduBaseUrl }
        if ($_config.HuduAssetLayoutId -gt 0)             { $script:HuduAssetLayoutId   = $_config.HuduAssetLayoutId }
        if (-not $AppName -and $_config.AuditAppName)     { $AppName                    = $_config.AuditAppName }
        if ($_config.HuduAssetName)                       { $script:HuduAssetName       = $_config.HuduAssetName }
    }
    catch { Write-Warning "Could not load config.psd1: $_" }
}
if (-not $HuduBaseUrl) { $HuduBaseUrl = 'https://neconnect.huducloud.com' }

# Microsoft Graph service principal app ID (constant in all Azure tenants)
$script:GraphResourceAppId = '00000003-0000-0000-c000-000000000000'

# Required Microsoft Graph application permissions for app-only audit authentication
$script:GraphPermissions = @(
    'Organization.Read.All',
    'Directory.Read.All',
    'User.Read.All',
    'Reports.Read.All',
    'Policy.Read.All',
    'UserAuthenticationMethod.Read.All',
    'RoleManagement.Read.Directory',
    'Group.Read.All',
    'AuditLog.Read.All',
    'SecurityEvents.Read.All',
    'DeviceManagementManagedDevices.Read.All',
    'DeviceManagementConfiguration.Read.All',
    'DeviceManagementApps.Read.All',
    'DeviceManagementServiceConfig.Read.All',
    'IdentityRiskyUser.Read.All',
    'IdentityRiskEvent.Read.All',
    'Application.Read.All',
    'Application.ReadWrite.OwnedBy',
    'TeamSettings.Read.All',
    'PrivilegedAccess.Read.AzureAD'
)

# Office 365 Exchange Online service principal app ID (constant in all Azure tenants)
$script:ExchangeResourceAppId = '00000002-0000-0ff1-ce00-000000000000'

# Required Exchange Online application permission for app-only PowerShell authentication
$script:ExchangePermissions = @('Exchange.ManageAsApp')

# SharePoint Online service principal app ID (constant in all Azure tenants)
$script:SharePointResourceAppId = '00000003-0000-0ff1-ce00-000000000000'

# Required SharePoint Online application permission for app-only access.
# Sites.FullControl.All is the minimum required for SharePoint tenant admin API calls
# (Get-PnPTenant, Get-PnPTenantSite) when using app-only auth — even for read operations.
$script:SharePointPermissions = @('Sites.FullControl.All')

# Days before certificate expiry to trigger a warning / offer rotation
$script:ExpiryWarnDays = 30

# Output directory for generated .pfx files (alongside this script)
$script:CertOutputDir = $PSScriptRoot

# Tracks every .pfx written this run so they can be deleted in finally{}
$script:GeneratedPfxPaths = [System.Collections.Generic.List[string]]::new()
$script:SetupNeedsManualConsent = $false
$script:SetupGraphModules = @(
    'Microsoft.Graph.Authentication',
    'Microsoft.Graph.Applications',
    'Microsoft.Graph.Identity.DirectoryManagement'
)


# ============================================================
# Helper: write formatted status messages
# ============================================================
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


function ConvertTo-NormalizedVersion {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [version]$Version
    )

    $revision = if ($Version.Revision -lt 0) { 0 } else { $Version.Revision }
    return [version]::new($Version.Major, $Version.Minor, $Version.Build, $revision)
}


function Resolve-SetupGraphModuleVersion {
    [CmdletBinding()]
    param()

    if ($script:SetupGraphModuleVersion) {
        return $script:SetupGraphModuleVersion
    }

    $commonVersions = $null

    foreach ($moduleName in $script:SetupGraphModules) {
        $versions = @(Get-Module -ListAvailable -Name $moduleName |
                Select-Object -ExpandProperty Version -Unique |
                Sort-Object -Descending)

        if (-not $versions) {
            throw "Required module '$moduleName' is not installed."
        }

        if ($null -eq $commonVersions) {
            $commonVersions = @($versions)
            continue
        }

        $commonVersions = @($commonVersions | Where-Object { $_ -in $versions })
    }

    if (-not $commonVersions) {
        $available = foreach ($moduleName in $script:SetupGraphModules) {
            $moduleVersions = @(Get-Module -ListAvailable -Name $moduleName |
                    Select-Object -ExpandProperty Version -Unique |
                    Sort-Object -Descending |
                    ForEach-Object { $_.ToString() })
            "{0}: {1}" -f $moduleName, ($moduleVersions -join ', ')
        }
        throw ("No common Microsoft Graph module version is installed for setup.`n" + ($available -join "`n"))
    }

    $script:SetupGraphModuleVersion = $commonVersions | Sort-Object -Descending | Select-Object -First 1
    Write-Verbose "Resolved setup Microsoft.Graph module version: $script:SetupGraphModuleVersion"
    return $script:SetupGraphModuleVersion
}


function Get-SetupGraphModuleInfo {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ModuleName,

        [Parameter(Mandatory)]
        [version]$ModuleVersion
    )

    $normalizedVersion = ConvertTo-NormalizedVersion -Version $ModuleVersion

    return Get-Module -ListAvailable -Name $ModuleName |
        Where-Object { (ConvertTo-NormalizedVersion -Version $_.Version) -eq $normalizedVersion } |
        Select-Object -First 1
}


function Initialize-SetupGraphDependencies {
    [CmdletBinding()]
    param()

    $targetVersion = Resolve-SetupGraphModuleVersion
    $authModule = Get-SetupGraphModuleInfo -ModuleName 'Microsoft.Graph.Authentication' -ModuleVersion $targetVersion
    if (-not $authModule) {
        throw "Unable to locate Microsoft.Graph.Authentication $targetVersion."
    }

    $dependencyDirs = @(
        (Join-Path $authModule.ModuleBase 'Dependencies\Core'),
        (Join-Path $authModule.ModuleBase 'Dependencies\Desktop'),
        (Join-Path $authModule.ModuleBase 'Dependencies'),
        $authModule.ModuleBase
    ) | Where-Object { Test-Path $_ }

    $alreadyLoaded = [AppDomain]::CurrentDomain.GetAssemblies() |
        Group-Object { $_.GetName().Name } -AsHashTable -AsString

    foreach ($assemblyName in @(
            'Microsoft.Graph.Core',
            'Azure.Core',
            'Azure.Identity',
            'Microsoft.Kiota.Abstractions',
            'Microsoft.Kiota.Authentication.Azure',
            'Microsoft.Kiota.Http.HttpClientLibrary',
            'Microsoft.Kiota.Serialization.Json',
            'Microsoft.Kiota.Serialization.Form',
            'Microsoft.Kiota.Serialization.Text')) {
        if ($alreadyLoaded.ContainsKey($assemblyName)) {
            continue
        }

        foreach ($dir in $dependencyDirs) {
            $candidatePath = Join-Path $dir ($assemblyName + '.dll')
            if (-not (Test-Path $candidatePath)) {
                continue
            }

            try {
                [System.Runtime.Loader.AssemblyLoadContext]::Default.LoadFromAssemblyPath($candidatePath) | Out-Null
                break
            }
            catch {
                Write-Verbose "Could not preload setup assembly '$assemblyName' from '$candidatePath': $_"
            }
        }
    }
}


function Import-SetupGraphModule {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ModuleName
    )

    $targetVersion = Resolve-SetupGraphModuleVersion
    $normalizedVersion = ConvertTo-NormalizedVersion -Version $targetVersion
    $moduleInfo = Get-SetupGraphModuleInfo -ModuleName $ModuleName -ModuleVersion $targetVersion

    if (-not $moduleInfo) {
        throw "Unable to locate module '$ModuleName' version $targetVersion."
    }

    $loadedModules = @(Get-Module -Name $ModuleName -All)
    $mismatched = @($loadedModules | Where-Object {
            (ConvertTo-NormalizedVersion -Version $_.Version) -ne $normalizedVersion
        })

    if ($mismatched) {
        $loadedSummary = $mismatched | ForEach-Object { "{0} ({1})" -f $_.Name, $_.Version }
        throw "Microsoft Graph setup module mismatch for '$ModuleName': $($loadedSummary -join ', '). Start a new PowerShell session and rerun setup."
    }

    if (-not ($loadedModules | Where-Object {
                (ConvertTo-NormalizedVersion -Version $_.Version) -eq $normalizedVersion
            })) {
        Import-Module -Name $moduleInfo.Path -ErrorAction Stop
    }

    return $moduleInfo
}


function Initialize-SetupGraphModules {
    [CmdletBinding()]
    param()

    Initialize-SetupGraphDependencies

    foreach ($moduleName in $script:SetupGraphModules) {
        Import-SetupGraphModule -ModuleName $moduleName | Out-Null
    }
}


# ============================================================
# Connect to Microsoft Graph with admin scopes
# Returns the tenant ID string
# ============================================================
function Connect-GraphForSetup {
    [CmdletBinding()]
    param()

    if (Get-MgContext) {
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    }

    $scopes = @(
        'Application.ReadWrite.All'
        'AppRoleAssignment.ReadWrite.All'
        'Directory.Read.All'
        'RoleManagement.ReadWrite.Directory'
    )

    Write-Status 'Connecting to Microsoft Graph (browser window will open for interactive sign-in)...'
    Connect-MgGraph -Scopes $scopes -NoWelcome -ErrorAction Stop

    $ctx = Get-MgContext
    if (-not $ctx) { throw "Authentication completed but no Graph context was established." }
    Write-Status "Connected — Tenant: $($ctx.TenantId)" -Type Success
    return $ctx.TenantId
}


# ============================================================
# Resolve AppRole IDs by name from a service principal
# Returns objects with Id, Name, ResourceSpId
# ============================================================
function Resolve-AppRoleIds {
    [CmdletBinding()]
    param (
        [string]   $ResourceAppId,
        [string[]] $PermissionNames
    )

    $sp = Get-MgServicePrincipal -Filter "appId eq '$ResourceAppId'" -ErrorAction Stop
    if (-not $sp) {
        throw "Service principal for appId '$ResourceAppId' not found in tenant."
    }

    foreach ($name in $PermissionNames) {
        $role = $sp.AppRoles | Where-Object { $_.Value -eq $name -and $_.AllowedMemberTypes -contains 'Application' }
        if (-not $role) {
            throw "Application permission '$name' not found on service principal '$ResourceAppId'."
        }
        [PSCustomObject]@{
            Id           = $role.Id
            Name         = $name
            ResourceSpId = $sp.Id
        }
    }
}


# ============================================================
# Ensure a service principal exists for our app
# ============================================================
function Resolve-ServicePrincipal {
    [CmdletBinding()]
    param ([string]$AppId)

    $sp = Get-MgServicePrincipal -Filter "appId eq '$AppId'" -ErrorAction SilentlyContinue
    if (-not $sp) {
        Write-Status 'Creating service principal for app...'
        $sp = New-MgServicePrincipal -AppId $AppId -ErrorAction Stop
    }
    return $sp
}


# ============================================================
# Grant admin consent for a list of resolved permissions
# ============================================================
function Grant-AdminConsent {
    [CmdletBinding()]
    param (
        [string]   $OurSpId,
        [object[]] $Permissions
    )

    foreach ($perm in $Permissions) {
        $existing = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $OurSpId -ErrorAction SilentlyContinue |
            Where-Object { $_.AppRoleId -eq $perm.Id }

        if ($existing) {
            Write-Verbose "Permission '$($perm.Name)' already granted — skipping."
            continue
        }

        try {
            New-MgServicePrincipalAppRoleAssignment `
                -ServicePrincipalId $OurSpId `
                -PrincipalId        $OurSpId `
                -ResourceId         $perm.ResourceSpId `
                -AppRoleId          $perm.Id `
                -ErrorAction Stop | Out-Null

            Write-Verbose "Granted: $($perm.Name)"
        }
        catch {
            $script:SetupNeedsManualConsent = $true
            Write-Status ("Automatic admin consent failed for '{0}': {1}" -f $perm.Name, $_.Exception.Message) -Type Warning
        }
    }
}


# ============================================================
# Assign Exchange Administrator Entra role to service principal
# Required for Exchange Online PowerShell app-only authentication
# ============================================================
function Set-ExchangeAdminRole {
    [CmdletBinding()]
    param ([string]$ServicePrincipalId)

    try {
        # Activate the role in the tenant if not yet activated (lazy-loaded roles)
        $role = Get-MgDirectoryRole -Filter "displayName eq 'Exchange Administrator'" -ErrorAction SilentlyContinue
        if (-not $role) {
            $template = Get-MgDirectoryRoleTemplate -ErrorAction Stop |
                Where-Object { $_.DisplayName -eq 'Exchange Administrator' }
            if (-not $template) {
                Write-Warning "Exchange Administrator role template not found — assign the role manually in Entra ID."
                return
            }
            $role = New-MgDirectoryRole -RoleTemplateId $template.Id -ErrorAction Stop
        }

        # Skip if already assigned (-All prevents pagination from missing members in large tenants)
        $existing = Get-MgDirectoryRoleMember -DirectoryRoleId $role.Id -All -ErrorAction SilentlyContinue |
            Where-Object { $_.Id -eq $ServicePrincipalId }
        if ($existing) {
            Write-Verbose 'Exchange Administrator role already assigned — skipping.'
            return
        }

        $body = @{ '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$ServicePrincipalId" }
        New-MgDirectoryRoleMemberByRef -DirectoryRoleId $role.Id -BodyParameter $body -ErrorAction Stop
        Write-Verbose "Exchange Administrator role assigned to service principal."
    }
    catch {
        if ($_.Exception.Message -match 'already exist') {
            Write-Verbose 'Exchange Administrator role already assigned — skipping.'
        }
        else {
            Write-Warning "Could not assign Exchange Administrator role automatically: $($_.Exception.Message)"
            Write-Warning "ACTION REQUIRED: In Entra ID, assign the 'Exchange Administrator' role to the '$(Get-Variable -Name AppName -ValueOnly -ErrorAction SilentlyContinue)' service principal manually."
        }
    }
}


# ============================================================
# Assign Global Reader Entra role to service principal
# Required by ScubaGear for non-interactive M365 baseline assessment
# ============================================================
function Set-GlobalReaderRole {
    [CmdletBinding()]
    param ([string]$ServicePrincipalId)

    try {
        # Activate the role in the tenant if not yet activated (lazy-loaded roles)
        $role = Get-MgDirectoryRole -Filter "displayName eq 'Global Reader'" -ErrorAction SilentlyContinue
        if (-not $role) {
            $template = Get-MgDirectoryRoleTemplate -ErrorAction Stop |
                Where-Object { $_.DisplayName -eq 'Global Reader' }
            if (-not $template) {
                Write-Warning "Global Reader role template not found — assign the role manually in Entra ID."
                return
            }
            $role = New-MgDirectoryRole -RoleTemplateId $template.Id -ErrorAction Stop
        }

        # Skip if already assigned (-All prevents pagination from missing members in large tenants)
        $existing = Get-MgDirectoryRoleMember -DirectoryRoleId $role.Id -All -ErrorAction SilentlyContinue |
            Where-Object { $_.Id -eq $ServicePrincipalId }
        if ($existing) {
            Write-Verbose 'Global Reader role already assigned — skipping.'
            return
        }

        $body = @{ '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$ServicePrincipalId" }
        New-MgDirectoryRoleMemberByRef -DirectoryRoleId $role.Id -BodyParameter $body -ErrorAction Stop
        Write-Verbose "Global Reader role assigned to service principal."
    }
    catch {
        if ($_.Exception.Message -match 'already exist') {
            Write-Verbose 'Global Reader role already assigned — skipping.'
        }
        else {
            Write-Warning "Could not assign Global Reader role automatically: $($_.Exception.Message)"
            Write-Warning "ACTION REQUIRED: In Entra ID, assign the 'Global Reader' role to the '$(Get-Variable -Name AppName -ValueOnly -ErrorAction SilentlyContinue)' service principal manually. This role is required by ScubaGear."
        }
    }
}


# ============================================================
# Analyse existing key (certificate) credentials on the app
# ============================================================
function Get-CertificateStatus {
    [CmdletBinding()]
    param (
        [Microsoft.Graph.PowerShell.Models.MicrosoftGraphApplication]$App
    )

    $now    = Get-Date
    $active = @($App.KeyCredentials | Where-Object { $_.Usage -eq 'Verify' -and $_.EndDateTime -gt $now })
    $soon   = @($active | Where-Object { $_.EndDateTime -lt $now.AddDays($script:ExpiryWarnDays) })
    $next   = $active | Sort-Object EndDateTime | Select-Object -First 1

    [PSCustomObject]@{
        HasActive           = $active.Count -gt 0
        ExpiresWithin30Days = $soon.Count -gt 0
        Soonest             = $next
    }
}


# ============================================================
# Generate a self-signed certificate, export the .pfx, and
# upload the public key to the Entra app registration.
# Returns: [PSCustomObject] PfxPath, PlainPassword, ExpiryDate
# ============================================================
function New-AuditCertificate {
    [CmdletBinding()]
    param (
        [string] $AppObjectId,
        [string] $AppId,
        [int]    $ExpiryYears
    )

    $certName  = "365Audit-$AppId"
    $notAfter  = [System.DateTimeOffset]::UtcNow.AddYears($ExpiryYears)

    # Generate a random 32-char password for the .pfx
    $chars     = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnpqrstuvwxyz23456789!@#%&*'
    $plainPwd  = -join ((1..32) | ForEach-Object { $chars[(Get-Random -Maximum $chars.Length)] })
    $securePwd = ConvertTo-SecureString $plainPwd -AsPlainText -Force

    # Export .pfx alongside this script
    $pfxPath = Join-Path $script:CertOutputDir "$certName.pfx"
    $days    = $ExpiryYears * 365

    if ($IsLinux -or $IsMacOS) {
        # New-SelfSignedCertificate and Cert:\ are Windows-only.
        # Use openssl (must be installed: apt install openssl / brew install openssl).
        $tmpKey  = [System.IO.Path]::GetTempFileName() + '.key'
        $tmpCert = [System.IO.Path]::GetTempFileName() + '.crt'
        try {
            & openssl req -x509 -newkey rsa:2048 -keyout $tmpKey -out $tmpCert `
                -days $days -nodes -subj "/CN=$certName" 2>$null
            if ($LASTEXITCODE -ne 0) { throw "openssl certificate generation failed (exit $LASTEXITCODE)" }

            # -legacy required on OpenSSL 3.x to produce a .pfx readable by .NET
            & openssl pkcs12 -export -legacy -out $pfxPath `
                -inkey $tmpKey -in $tmpCert -passout "pass:$plainPwd" 2>$null
            if ($LASTEXITCODE -ne 0) { throw "openssl pfx export failed (exit $LASTEXITCODE)" }

            $rawData = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new(
                [System.IO.File]::ReadAllBytes($tmpCert)
            ).RawData
        }
        finally {
            Remove-Item $tmpKey, $tmpCert -ErrorAction SilentlyContinue
        }
    }
    else {
        # -KeySpec KeyExchange forces the legacy CSP provider (not CNG).
        # EXO Connect-ExchangeOnline -CertificateFilePath requires a CSP cert.
        $cert = New-SelfSignedCertificate `
            -Subject           "CN=$certName" `
            -CertStoreLocation 'Cert:\CurrentUser\My' `
            -NotAfter          $notAfter.LocalDateTime `
            -KeySpec           KeyExchange `
            -ErrorAction Stop

        $cert | Export-PfxCertificate -FilePath $pfxPath -Password $securePwd -ErrorAction Stop | Out-Null
        $rawData = [byte[]]$cert.RawData

        # Remove from local cert store — the .pfx is the portable copy
        Remove-Item "Cert:\CurrentUser\My\$($cert.Thumbprint)" -ErrorAction SilentlyContinue
    }

    # Upload the public key to the Entra app registration.
    # Graph SDK requires DateTimeOffset (not DateTime) for key credential dates.
    Update-MgApplication -ApplicationId $AppObjectId -KeyCredentials @(
        @{
            type          = 'AsymmetricX509Cert'
            usage         = 'Verify'
            key           = $rawData
            displayName   = $certName
            startDateTime = [System.DateTimeOffset]::UtcNow
            endDateTime   = $notAfter
        }
    ) -ErrorAction Stop

    # Encode the .pfx as base64 so it can be stored as a Hudu secret
    $certBase64 = [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($pfxPath))
    $script:GeneratedPfxPaths.Add($pfxPath)

    return [PSCustomObject]@{
        PfxPath       = $pfxPath
        PlainPassword = $plainPwd
        ExpiryDate    = $notAfter.LocalDateTime
        CertBase64    = $certBase64
    }
}


# ============================================================
# Open Azure portal to API permissions page for admin consent
# ============================================================
function Request-AdminConsent {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)] [string]$ApplicationId,
        [Parameter(Mandatory)] [string]$TenantName
    )

    $portalUrl = "https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/~/CallAnAPI/appId/$ApplicationId/isMSAApp~/false"

    $sep = '=' * 72
    Write-Host "`n$sep" -ForegroundColor Yellow
    Write-Host '  ACTION REQUIRED: Grant admin consent in the Azure portal' -ForegroundColor Yellow
    Write-Host $sep -ForegroundColor Yellow
    Write-Host ''
    Write-Host '  Waiting for Entra ID to finish replicating the application...' -ForegroundColor DarkYellow
    Write-Host '  This usually takes 5-10 seconds.' -ForegroundColor DarkYellow

    Start-Sleep -Seconds 10

    Write-Host ''
    Write-Host '  Opening Azure portal — API Permissions page for this app...' -ForegroundColor Cyan
    Write-Host "  In the portal, click: " -NoNewline -ForegroundColor White
    Write-Host "'Grant admin consent for $TenantName'" -ForegroundColor Yellow
    Write-Host "  then click 'Yes' to confirm." -ForegroundColor White
    Write-Host ''
    Write-Host "  If the browser doesn't open automatically, use this URL:" -ForegroundColor DarkCyan
    Write-Host "  $portalUrl" -ForegroundColor Cyan
    Write-Host "$sep`n" -ForegroundColor Yellow

    try {
        if ($IsLinux) {
            xdg-open $portalUrl
        } elseif ($IsMacOS) {
            open $portalUrl
        } else {
            Start-Process $portalUrl -ErrorAction Stop
        }
    }
    catch {
        Write-Warning "Unable to open browser automatically: $($_.Exception.Message)"
    }
}



# ============================================================
# Print credentials in a clearly formatted block
# ============================================================
function Write-CredentialSummary {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingPlainTextForPassword', 'CertPassword',
        Justification = 'Intentionally plain text — displayed once so the operator can store it in the password manager.')]
    [CmdletBinding()]
    param (
        [string]   $AppId,
        [string]   $TenantId,
        [string]   $CertBase64,
        [string]   $CertPassword,
        [datetime] $CertExpiry
    )

    $sep = '=' * 72
    Write-Host "`n$sep" -ForegroundColor Cyan
    Write-Host '  NeConnect MSA Audit Toolkit — Store these credentials in Hudu' -ForegroundColor Cyan
    Write-Host $sep -ForegroundColor Cyan
    Write-Host "  App ID (Client ID) : $AppId"
    Write-Host "  Tenant ID          : $TenantId"
    Write-Host "  Cert Base64        : $CertBase64" -ForegroundColor Yellow
    Write-Host "  Cert Password      : $CertPassword" -ForegroundColor Yellow
    Write-Host "  Cert Expires       : $($CertExpiry.ToString('yyyy-MM-dd'))"
    Write-Host ''
    Write-Host '  Run the audit with:' -ForegroundColor DarkCyan
    Write-Host "  .\Start-365Audit.ps1 -AppId '$AppId' -TenantId '$TenantId' -CertBase64 '<paste base64>' -CertPassword (Read-Host -AsSecureString 'Cert Password')" -ForegroundColor Cyan
    Write-Host "$sep`n" -ForegroundColor Cyan
}


# ============================================================
# Fetch audit credentials from a Hudu asset
# Returns: [PSCustomObject] AppId, TenantId, CertBase64, CertPassword, CompanyName
# ============================================================
function Get-HuduAuditCredentials {
    [CmdletBinding()]
    param (
        [string]$HuduCompanyId,
        [string]$HuduCompanyName,
        [string]$HuduBaseUrl,
        [string]$HuduApiKey
    )

    if (-not $HuduApiKey) {
        throw "HUDU_API_KEY is not set. Provide -HuduApiKey or set the HUDU_API_KEY environment variable."
    }

    $huduUrl = $HuduBaseUrl.TrimEnd('/')
    $headers = @{ 'x-api-key' = $HuduApiKey; 'Content-Type' = 'application/json' }

    # Resolve company
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
    elseif ($HuduCompanyName) {
        $encoded = [uri]::EscapeDataString($HuduCompanyName)
        $company = @((Invoke-RestMethod -Uri "$huduUrl/api/v1/companies?search=$encoded&page_size=25" -Headers $headers -Method Get -ErrorAction Stop).companies) |
            Where-Object { $_.name -eq $HuduCompanyName } | Select-Object -First 1
    }
    else {
        throw "Either -HuduCompanyId or -HuduCompanyName must be provided."
    }

    if (-not $company) { throw "No Hudu company found for '$($HuduCompanyId ?? $HuduCompanyName)'." }

    $asset = @((Invoke-RestMethod -Uri "$huduUrl/api/v1/assets?company_id=$($company.id)&asset_layout_id=$($script:HuduAssetLayoutId)&page_size=5" `
        -Headers $headers -Method Get -ErrorAction Stop).assets) | Sort-Object updated_at -Descending | Select-Object -First 1

    if (-not $asset) {
        Write-Status "No existing 365Audit asset found for '$($company.name)' — will run first-time setup." -Type Info
        return $null
    }

    $fieldMap = @{}
    foreach ($f in $asset.fields) { $fieldMap[$f.label] = "$($f.value)" }

    foreach ($required in @('Application ID', 'Tenant ID', 'Cert Base64', 'Cert Password')) {
        if (-not $fieldMap[$required]) {
            Write-Status "Hudu asset '$($asset.name)' is missing field '$required' — will run first-time setup." -Type Warning
            return $null
        }
    }

    return [PSCustomObject]@{
        AppId        = $fieldMap['Application ID']
        TenantId     = $fieldMap['Tenant ID']
        CertBase64   = $fieldMap['Cert Base64']
        CertPassword = $fieldMap['Cert Password']
        CompanyName  = $company.name
        CompanyId    = $company.id
    }
}


# ============================================================
# Push credentials to Hudu
# Requires HUDU_BASE_URL and HUDU_API_KEY environment variables.
# Prompts for the Hudu company URL or numeric company ID.
# Finds the existing 'NeConnect Audit Toolkit' asset (layout 67)
# for that company and updates it, or creates one if absent.
# ============================================================
function Push-HuduAuditAsset {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingPlainTextForPassword', 'CertPassword',
        Justification = 'Plain text required — value is written to Hudu password field via REST API.')]
    [CmdletBinding()]
    param (
        [string]   $AppId,
        [string]   $TenantId,
        [string]   $CertBase64,
        [string]   $CertPassword,
        [datetime] $CertExpiry,
        [string]   $HuduCompanyId,
        [string]   $HuduCompanyName,
        [string]   $HuduBaseUrl,
        [string]   $HuduApiKey
    )

    $huduUrl = $HuduBaseUrl.TrimEnd('/')
    $huduKey = $HuduApiKey

    if (-not $huduUrl -or -not $huduKey) {
        Write-Warning 'Hudu credentials not configured — skipping Hudu push.'
        Write-Host '  Populate HuduBaseUrl and HuduApiKey in config.psd1 (see config.psd1.example), or pass -HuduApiKey on the command line.' -ForegroundColor DarkGray
        return
    }

    $headers = @{ 'x-api-key' = $huduKey; 'Content-Type' = 'application/json' }

    # --- Resolve company ---
    $company = $null

    if ($HuduCompanyId) {
        # Slug or numeric ID supplied via parameter
        $companyId = $HuduCompanyId
        try {
            if ($companyId -match '^\d+$') {
                $company = (Invoke-RestMethod -Uri "$huduUrl/api/v1/companies/$companyId" `
                    -Headers $headers -Method Get).company
            }
            else {
                $encoded = [uri]::EscapeDataString($companyId)
                $company = @((Invoke-RestMethod -Uri "$huduUrl/api/v1/companies?slug=$encoded&page_size=1" `
                    -Headers $headers -Method Get).companies) | Select-Object -First 1
            }
        }
        catch { Write-Warning "Hudu company lookup failed for '$companyId': $_"; return }
    }
    elseif ($HuduCompanyName) {
        # Exact name supplied via parameter
        try {
            $encoded = [uri]::EscapeDataString($HuduCompanyName)
            $company = @((Invoke-RestMethod -Uri "$huduUrl/api/v1/companies?search=$encoded&page_size=25" `
                -Headers $headers -Method Get).companies) |
                Where-Object { $_.name -eq $HuduCompanyName } | Select-Object -First 1
        }
        catch { Write-Warning "Hudu company lookup failed for '$HuduCompanyName': $_"; return }
    }
    else {
        # Interactive prompt — accept URL, slug, or numeric ID
        Write-Host ''
        Write-Host '  Paste the Hudu company URL (e.g. https://hudu.example.com/c/contoso),' -ForegroundColor DarkCyan
        Write-Host '  company slug, or numeric ID:' -ForegroundColor DarkCyan
        $companyInput = (Read-Host '  Company URL, slug, or ID').Trim()

        $companyId = if ($companyInput -match '://') {
            try { ([System.Uri]$companyInput).Segments[-1].TrimEnd('/') }
            catch { Write-Warning "Could not parse URL '$companyInput': $_"; return }
        } else { $companyInput }

        try {
            if ($companyId -match '^\d+$') {
                $company = (Invoke-RestMethod -Uri "$huduUrl/api/v1/companies/$companyId" `
                    -Headers $headers -Method Get).company
            }
            else {
                $encoded = [uri]::EscapeDataString($companyId)
                $company = @((Invoke-RestMethod -Uri "$huduUrl/api/v1/companies?slug=$encoded&page_size=1" `
                    -Headers $headers -Method Get).companies) | Select-Object -First 1
            }
        }
        catch { Write-Warning "Hudu company lookup failed for '$companyId': $_"; return }
    }

    if (-not $company) {
        Write-Warning "No Hudu company found — skipping asset push."
        return
    }

    $companyName = $company.name
    $companyId   = $company.id   # normalise to numeric DB id for subsequent calls
    Write-Host "  Company: $companyName (id: $companyId)" -ForegroundColor Green

    # Build field payload using snake_case custom field names (Hudu API requirement).
    # Field names are the asset layout field labels lowercased with spaces replaced by underscores.
    $companySlug = if ($company.slug) { $company.slug } else { $companyId }
    $launchCmd = "<p><strong>Without Hudu API key:</strong><br>" +
                 ".\Start-365Audit.ps1 -AppId '$AppId' -TenantId '$TenantId' -CertPassword (Read-Host -AsSecureString 'Cert Password')</p>" +
                 "<p><strong>With Hudu API key:</strong><br>" +
                 ".\Start-365Audit.ps1 -HuduCompanyId '$companySlug'</p>"

    $body = @{
        name            = "$($script:HuduAssetName) - $companyName"
        asset_layout_id = $script:HuduAssetLayoutId
        custom_fields   = @(
            @{ application_id            = $AppId }
            @{ tenant_id                 = $TenantId }
            @{ cert_base64               = $CertBase64 }
            @{ cert_password             = $CertPassword }
            @{ cert_expiry               = $CertExpiry.ToString('yyyy/MM/dd') }
            @{ powershell_launch_command = $launchCmd }
        )
    } | ConvertTo-Json -Depth 5

    # Find existing asset for this company in the layout
    try {
        $existingResult = Invoke-RestMethod `
            -Uri     "$huduUrl/api/v1/assets?company_id=$companyId&asset_layout_id=$($script:HuduAssetLayoutId)&page_size=5" `
            -Headers $headers -Method Get -ErrorAction Stop
        $existingAsset = @($existingResult.assets) | Select-Object -First 1
    }
    catch {
        Write-Warning "Could not query Hudu assets: $_"
        $existingAsset = $null
    }

    try {
        if ($existingAsset) {
            Invoke-RestMethod -Uri "$huduUrl/api/v1/assets/$($existingAsset.id)" `
                -Headers $headers -Method Put -Body $body -ErrorAction Stop | Out-Null
            Write-Host "  Hudu asset updated: $($existingAsset.name) (id: $($existingAsset.id))" -ForegroundColor Green
        }
        else {
            # Assets must be created under the company-scoped endpoint
            $created = Invoke-RestMethod -Uri "$huduUrl/api/v1/companies/$companyId/assets" `
                -Headers $headers -Method Post -Body $body -ErrorAction Stop
            Write-Host "  Hudu asset created: $($created.asset.name) (id: $($created.asset.id))" -ForegroundColor Green
        }
    }
    catch {
        Write-Warning "Hudu asset push failed — credentials were printed above and must be saved manually: $_"
    }
}



# ============================================================
# Shared sequences — called from multiple execution paths
# ============================================================

# Invoke-PermissionCheck
# Diffs required vs present permissions and applies any missing ones.
# Returns $true if any permissions were added (caller may then open admin consent portal).
function Invoke-PermissionCheck {
    param ([object]$App)

    $gPerms  = @(Resolve-AppRoleIds -ResourceAppId $script:GraphResourceAppId      -PermissionNames $script:GraphPermissions)
    $ePerms  = @(Resolve-AppRoleIds -ResourceAppId $script:ExchangeResourceAppId   -PermissionNames $script:ExchangePermissions)
    $spPerms = @(Resolve-AppRoleIds -ResourceAppId $script:SharePointResourceAppId -PermissionNames $script:SharePointPermissions)
    $sp      = Resolve-ServicePrincipal -AppId $App.AppId

    $currentIds  = @($App.RequiredResourceAccess.ResourceAccess | ForEach-Object { $_.Id })
    $missingG    = @($gPerms  | Where-Object { $_.Id -notin $currentIds })
    $missingE    = @($ePerms  | Where-Object { $_.Id -notin $currentIds })
    $missingSP   = @($spPerms | Where-Object { $_.Id -notin $currentIds })

    $grantedIds       = @(Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id -All -ErrorAction SilentlyContinue | ForEach-Object { $_.AppRoleId })
    $missingConsentG  = @($gPerms  | Where-Object { $_.Id -notin $grantedIds })
    $missingConsentE  = @($ePerms  | Where-Object { $_.Id -notin $grantedIds })
    $missingConsentSP = @($spPerms | Where-Object { $_.Id -notin $grantedIds })

    if ($missingG.Count -gt 0 -or $missingE.Count -gt 0 -or $missingSP.Count -gt 0) {
        $names = ($missingG + $missingE + $missingSP | ForEach-Object { $_.Name }) -join ', '
        Write-Status "Missing application permissions detected: $names" -Type Warning

        $reqResourceAccess = @(
            @{ resourceAppId = $script:GraphResourceAppId;      resourceAccess = @($gPerms  | ForEach-Object { @{ id = $_.Id.ToString(); type = 'Role' } }) },
            @{ resourceAppId = $script:ExchangeResourceAppId;   resourceAccess = @($ePerms  | ForEach-Object { @{ id = $_.Id.ToString(); type = 'Role' } }) },
            @{ resourceAppId = $script:SharePointResourceAppId; resourceAccess = @($spPerms | ForEach-Object { @{ id = $_.Id.ToString(); type = 'Role' } }) }
        )
        try {
            Update-MgApplication -ApplicationId $App.Id -RequiredResourceAccess $reqResourceAccess -ErrorAction Stop
        }
        catch {
            if ($_.Exception.Message -match 'Authorization_RequestDenied|Insufficient privileges') {
                $ctx = Get-MgContext
                if ($ctx -and $ctx.AuthType -eq 'Delegated') {
                    # Already interactive — the failure is a real permissions issue, not an auth type mismatch.
                    # Reconnecting would just prompt the user again for no benefit.
                    throw
                }
                # App-only auth cannot update its own permissions — reconnect interactively and retry.
                Write-Status "App-only auth cannot update permissions — reconnecting interactively to apply changes..." -Type Warning
                Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
                Connect-GraphForSetup | Out-Null
                Update-MgApplication -ApplicationId $App.Id -RequiredResourceAccess $reqResourceAccess -ErrorAction Stop
            }
            else { throw }
        }
    }

    if ($missingConsentG.Count -gt 0 -or $missingConsentE.Count -gt 0 -or $missingConsentSP.Count -gt 0) {
        $consentNames = ($missingConsentG + $missingConsentE + $missingConsentSP | ForEach-Object { $_.Name }) -join ', '
        Write-Status "Missing admin consent detected: $consentNames" -Type Warning
    }

    if ($missingG.Count -gt 0 -or $missingE.Count -gt 0 -or $missingSP.Count -gt 0 -or
        $missingConsentG.Count -gt 0 -or $missingConsentE.Count -gt 0 -or $missingConsentSP.Count -gt 0) {
        if ($missingConsentG.Count  -gt 0) { Grant-AdminConsent -OurSpId $sp.Id -Permissions $missingConsentG }
        if ($missingConsentSP.Count -gt 0) { Grant-AdminConsent -OurSpId $sp.Id -Permissions $missingConsentSP }
        if ($missingConsentE.Count  -gt 0) {
            Grant-AdminConsent -OurSpId $sp.Id -Permissions $missingConsentE
            Write-Status 'Assigning Exchange Administrator role...'
            Set-ExchangeAdminRole -ServicePrincipalId $sp.Id
        }
        Write-Status 'Assigning Global Reader role (required by ScubaGear)...'
        Set-GlobalReaderRole -ServicePrincipalId $sp.Id
        Write-Status 'Permissions validated.' -Type Success
        return $true
    }

    Write-Status 'All required permissions and admin consent are present.' -Type Success
    return $false
}


function Resolve-PermissionState {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$App,

        [string]$TenantId
    )

    $orgName = $null
    $script:SetupNeedsManualConsent = $false
    $permissionChanges = Invoke-PermissionCheck -App $App

    if ($script:SetupNeedsManualConsent) {
        $ctx = Get-MgContext
        if ($ctx -and $ctx.AuthType -eq 'Delegated') {
            # Already in an interactive session — a reconnect would just prompt again for no benefit.
            # Go straight to manual portal consent.
            if (-not $orgName) {
                $orgName = (Get-MgOrganization -ErrorAction SilentlyContinue | Select-Object -First 1).DisplayName
            }
            Request-AdminConsent -ApplicationId $App.AppId -TenantName ($orgName ?? $TenantId)
        }
        else {
            Write-Status 'Automatic admin consent was not sufficient — reconnecting interactively to complete consent...' -Type Warning

            # Capture org name while still connected app-only, before disconnecting
            if (-not $orgName) {
                $orgName = (Get-MgOrganization -ErrorAction SilentlyContinue | Select-Object -First 1).DisplayName
            }

            Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null

            $interactiveOk = $false
            try {
                $TenantId     = Connect-GraphForSetup
                $orgName      = (Get-MgOrganization -ErrorAction Stop | Select-Object -First 1).DisplayName
                Write-Status "Tenant: $orgName ($TenantId)" -Type Success
                $interactiveOk = $true
            }
            catch {
                Write-Status "Interactive sign-in unavailable: $($_.Exception.Message)" -Type Warning
                Write-Status 'Falling back to manual portal consent...' -Type Warning
            }

            if ($interactiveOk) {
                $App = Get-MgApplication -Filter "appId eq '$($App.AppId)'" -ErrorAction Stop | Select-Object -First 1
                if (-not $App) {
                    throw "No app registration found for AppId '$($App.AppId)' after interactive reconnect."
                }

                $script:SetupNeedsManualConsent = $false
                $permissionChanges = Invoke-PermissionCheck -App $App
            }

            if ($script:SetupNeedsManualConsent) {
                Request-AdminConsent -ApplicationId $App.AppId -TenantName ($orgName ?? $TenantId)
            }
        }
    }

    return [PSCustomObject]@{
        App               = $App
        TenantId          = $TenantId
        OrgName           = $orgName
        PermissionChanges = $permissionChanges
    }
}

# Invoke-OwnerCheck
# Ensures the app's own service principal is listed as an owner of the app registration.
# Required for Application.ReadWrite.OwnedBy to permit non-interactive cert renewal.
function Invoke-OwnerCheck {
    param ([object]$App)

    if (-not (Get-MgContext)) {
        Write-Verbose 'No active Graph session — skipping owner check.'
        return
    }

    $sp = Resolve-ServicePrincipal -AppId $App.AppId
    try {
        New-MgApplicationOwnerByRef -ApplicationId $App.Id -BodyParameter @{
            '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$($sp.Id)"
        } -ErrorAction Stop
        Write-Status 'Service principal confirmed as app owner.' -Type Success
    }
    catch {
        if ($_.Exception.Message -match 'already exist') {
            Write-Verbose 'SP already an owner.'
        }
        else {
            Write-Warning "Could not confirm SP as app owner (non-fatal): $_"
        }
    }
}

# New-EntraApp
# Creates a brand-new app registration with all required permissions, service
# principal, owner assignment, programmatic admin consent, and Exchange Admin role.
# Opens the Azure portal for the operator to confirm consent visually.
# Returns the created app object.
function New-EntraApp {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)] [string]$AppName,
        [Parameter(Mandatory)] [string]$TenantId,
        [Parameter(Mandatory)] [string]$OrgName
    )

    Write-Status "Creating app registration: '$AppName'..."

    $gPerms  = @(Resolve-AppRoleIds -ResourceAppId $script:GraphResourceAppId      -PermissionNames $script:GraphPermissions)
    $ePerms  = @(Resolve-AppRoleIds -ResourceAppId $script:ExchangeResourceAppId   -PermissionNames $script:ExchangePermissions)
    $spPerms = @(Resolve-AppRoleIds -ResourceAppId $script:SharePointResourceAppId -PermissionNames $script:SharePointPermissions)

    $app = New-MgApplication -DisplayName $AppName -RequiredResourceAccess @(
        @{ resourceAppId = $script:GraphResourceAppId;      resourceAccess = @($gPerms  | ForEach-Object { @{ id = $_.Id.ToString(); type = 'Role' } }) },
        @{ resourceAppId = $script:ExchangeResourceAppId;   resourceAccess = @($ePerms  | ForEach-Object { @{ id = $_.Id.ToString(); type = 'Role' } }) },
        @{ resourceAppId = $script:SharePointResourceAppId; resourceAccess = @($spPerms | ForEach-Object { @{ id = $_.Id.ToString(); type = 'Role' } }) }
    ) -ErrorAction Stop
    Write-Status "App created — App ID: $($app.AppId)" -Type Success

    Write-Status 'Waiting for app to replicate across Entra ID...'
    Start-Sleep -Seconds 10

    $sp = Resolve-ServicePrincipal -AppId $app.AppId

    # Register the SP as an owner so it can renew its own certificate non-interactively
    # (required for Application.ReadWrite.OwnedBy to work on subsequent unattended runs)
    try {
        New-MgApplicationOwnerByRef -ApplicationId $app.Id -BodyParameter @{
            '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$($sp.Id)"
        } -ErrorAction Stop
        Write-Status 'Service principal registered as app owner.' -Type Success
    }
    catch {
        if ($_.Exception.Message -match 'already exist') { Write-Verbose 'SP already an owner.' }
        else { Write-Warning "Could not register SP as app owner (non-fatal): $_" }
    }

    Write-Status 'Granting admin consent for all permissions...'
    Grant-AdminConsent -OurSpId $sp.Id -Permissions $gPerms
    Grant-AdminConsent -OurSpId $sp.Id -Permissions $spPerms
    Grant-AdminConsent -OurSpId $sp.Id -Permissions $ePerms
    Write-Status 'Assigning Exchange Administrator role...'
    Set-ExchangeAdminRole -ServicePrincipalId $sp.Id
    Write-Status 'Assigning Global Reader role (required by ScubaGear)...'
    Set-GlobalReaderRole -ServicePrincipalId $sp.Id

    Request-AdminConsent -ApplicationId $app.AppId -TenantName $OrgName

    return $app
}


# ============================================================
# Main
# ============================================================
try {
    Write-Host "`n365Audit App Setup v$ScriptVersion`n" -ForegroundColor Cyan

    # ── Ensure required Graph modules ─────────────────────────────────────────
    Write-Status 'Checking required PowerShell modules... (this can take up to 30 seconds)' -Type Info
    foreach ($mod in $script:SetupGraphModules) {
        if (-not (Get-Module -ListAvailable -Name $mod)) {
            Write-Status "Installing $mod..." -Type Warning
            Install-Module $mod -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        }
    }
    Initialize-SetupGraphModules

    # Platform-appropriate key storage flags for X509Certificate2 import
    $certKeyFlags = if ($IsLinux -or $IsMacOS) {
        [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::Exportable -bor
        [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::PersistKeySet
    } else {
        [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::EphemeralKeySet
    }

    # ══════════════════════════════════════════════════════════════════════════
    # MODE 3 — Explicit AppId + TenantId supplied
    # App-only connect using the provided cert. No Hudu push.
    # ══════════════════════════════════════════════════════════════════════════
    if ($AppId -and $TenantId) {
        if (-not $CertBase64) {
            $CertBase64 = Read-Host 'Paste certificate Base64'
        }
        if (-not $CertPassword) {
            $CertPassword = Read-Host 'Certificate password' -AsSecureString
        }

        $certBytes   = try { [Convert]::FromBase64String($CertBase64) }
                       catch { throw "CertBase64 is not valid base64: $_" }
        $connectCert = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($certBytes, $CertPassword, $certKeyFlags)

        Write-Status 'Connecting to Microsoft Graph (app-only — no browser required)...'
        Connect-MgGraph -ClientId $AppId -TenantId $TenantId -Certificate $connectCert -NoWelcome -ErrorAction Stop
        $connectCert = $null   # MgGraph holds the CNG key handle; do not Dispose here
        Write-Status "Connected." -Type Success

        $app = Get-MgApplication -Filter "appId eq '$AppId'" -ErrorAction Stop | Select-Object -First 1
        if (-not $app) { throw "No app registration found for AppId '$AppId'." }
        Write-Status "App: $($app.DisplayName) | AppId: $($app.AppId)" -Type Info

        $permResult = Resolve-PermissionState -App $app -TenantId $TenantId
        $app        = $permResult.App
        $TenantId   = $permResult.TenantId
        Invoke-OwnerCheck -App $app

        # Cert check uses the live app registration (more authoritative than the passed-in cert)
        $certStatus = Get-CertificateStatus -App $app
        if (-not $certStatus.HasActive) {
            Write-Status 'No active certificate on app — generating new certificate.' -Type Warning
            $newCert = New-AuditCertificate -AppObjectId $app.Id -AppId $AppId -ExpiryYears $CertExpiryYears
            Write-Status "Certificate generated — expires $($newCert.ExpiryDate.ToString('yyyy-MM-dd'))." -Type Success
            Write-CredentialSummary -AppId $AppId -TenantId $TenantId -CertBase64 $newCert.CertBase64 -CertPassword $newCert.PlainPassword -CertExpiry $newCert.ExpiryDate
        }
        elseif ($certStatus.ExpiresWithin30Days -or $Force) {
            $reason = if ($Force) { '-Force specified' } else { "expiring within $script:ExpiryWarnDays days" }
            Write-Status "Certificate $reason — generating new certificate." -Type Warning
            $newCert = New-AuditCertificate -AppObjectId $app.Id -AppId $AppId -ExpiryYears $CertExpiryYears
            Write-Status "Certificate renewed — expires $($newCert.ExpiryDate.ToString('yyyy-MM-dd'))." -Type Success
            Write-CredentialSummary -AppId $AppId -TenantId $TenantId -CertBase64 $newCert.CertBase64 -CertPassword $newCert.PlainPassword -CertExpiry $newCert.ExpiryDate
        }
        else {
            $expiry = $certStatus.Soonest.EndDateTime
            Write-Status "Certificate healthy — expires $($expiry.ToString('yyyy-MM-dd'))." -Type Success
        }
    }

    # ══════════════════════════════════════════════════════════════════════════
    # MODE 2 — Hudu company context supplied
    # ══════════════════════════════════════════════════════════════════════════
    elseif ($HuduCompanyId -or $HuduCompanyName) {
        if (-not $HuduApiKey) {
            throw "HUDU_API_KEY is required when using -HuduCompanyId or -HuduCompanyName. Set the environment variable or pass -HuduApiKey."
        }

        Write-Status "Looking up Hudu company '$($HuduCompanyId ?? $HuduCompanyName)'..."
        $huduCreds = Get-HuduAuditCredentials `
            -HuduCompanyId   $HuduCompanyId `
            -HuduCompanyName $HuduCompanyName `
            -HuduBaseUrl     $HuduBaseUrl `
            -HuduApiKey      $HuduApiKey

        if ($huduCreds) {
            # ── Mode 2b: Asset found — app-only connect ───────────────────────────
            Write-Status "Company: $($huduCreds.CompanyName)" -Type Success

            # Load cert once — reused directly for Connect-MgGraph to avoid the Windows CNG
            # double-import error that occurs when the same EphemeralKeySet PFX bytes are
            # decoded twice in the same process (second import fails with ERROR_PATH_NOT_FOUND).
            $certBytes   = [Convert]::FromBase64String($huduCreds.CertBase64)
            $certPwd     = ConvertTo-SecureString $huduCreds.CertPassword -AsPlainText -Force
            $connectCert = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($certBytes, $certPwd, $certKeyFlags)

            $daysRemaining = ($connectCert.NotAfter - (Get-Date)).Days
            $needsRenewal  = ($daysRemaining -le 0) -or ($daysRemaining -le $script:ExpiryWarnDays) -or $Force

            if ($daysRemaining -le 0) {
                Write-Status "Certificate EXPIRED $([math]::Abs($daysRemaining)) day(s) ago — will renew after connecting." -Type Warning
            } elseif ($daysRemaining -le $script:ExpiryWarnDays -or $Force) {
                $reason = if ($Force) { '-Force' } else { "$daysRemaining day(s) remaining" }
                Write-Status "Certificate expiring ($reason) — will renew after connecting." -Type Warning
            } else {
                Write-Status "Certificate valid — $daysRemaining day(s) remaining." -Type Success
            }

            $AppId    = $huduCreds.AppId
            $TenantId = $huduCreds.TenantId

            Write-Status 'Connecting to Microsoft Graph (app-only — no browser required)...'
            Connect-MgGraph -ClientId $AppId -TenantId $TenantId -Certificate $connectCert -NoWelcome -ErrorAction Stop
            $connectCert = $null   # MgGraph holds the CNG key handle; do not Dispose here
            Write-Status "Connected." -Type Success

            $app = Get-MgApplication -Filter "appId eq '$AppId'" -ErrorAction Stop | Select-Object -First 1
            if (-not $app) { throw "No app registration found for AppId '$AppId'." }
            Write-Status "App: $($app.DisplayName) | AppId: $($app.AppId)" -Type Info

            $permResult = Resolve-PermissionState -App $app -TenantId $TenantId
            $app        = $permResult.App
            $TenantId   = $permResult.TenantId
            Invoke-OwnerCheck -App $app

            if ($needsRenewal) {
                Write-Status "Generating $CertExpiryYears-year replacement certificate..."
                $newCert = New-AuditCertificate -AppObjectId $app.Id -AppId $AppId -ExpiryYears $CertExpiryYears
                Write-Status "Certificate renewed — expires $($newCert.ExpiryDate.ToString('yyyy-MM-dd'))." -Type Success
                Write-CredentialSummary -AppId $AppId -TenantId $TenantId -CertBase64 $newCert.CertBase64 -CertPassword $newCert.PlainPassword -CertExpiry $newCert.ExpiryDate
                Push-HuduAuditAsset     -AppId $AppId -TenantId $TenantId -CertBase64 $newCert.CertBase64 -CertPassword $newCert.PlainPassword -CertExpiry $newCert.ExpiryDate `
                    -HuduCompanyId $HuduCompanyId -HuduCompanyName $HuduCompanyName -HuduBaseUrl $HuduBaseUrl -HuduApiKey $HuduApiKey
            } else {
                Write-Status 'All checks passed — no changes needed.' -Type Success
            }
        }
        else {
            # ── Mode 2a: No asset found — first-time interactive setup ────────────
            # HuduCompanyId/Name remain set so Push-HuduAuditAsset fires after cert generation.
            Write-Status 'No existing Hudu asset — running first-time interactive setup...' -Type Info

            $TenantId = Connect-GraphForSetup
            $orgName  = (Get-MgOrganization -ErrorAction Stop | Select-Object -First 1).DisplayName
            Write-Status "Tenant: $orgName ($TenantId)" -Type Success

            Write-Status "Searching for existing app: '$AppName'..."
            $app = Get-MgApplication -Filter "displayName eq '$AppName'" -ErrorAction Stop | Select-Object -First 1

            if ($app) {
                Write-Status "App found — App ID: $($app.AppId)" -Type Info
                $permResult = Resolve-PermissionState -App $app -TenantId $TenantId
                $app        = $permResult.App
                $TenantId   = $permResult.TenantId
                if ($permResult.OrgName) { $orgName = $permResult.OrgName }
                Invoke-OwnerCheck -App $app
            } else {
                $app = New-EntraApp -AppName $AppName -TenantId $TenantId -OrgName $orgName
            }

            $certStatus = Get-CertificateStatus -App $app
            if (-not $certStatus.HasActive -or $certStatus.ExpiresWithin30Days -or $Force) {
                $reason = if (-not $certStatus.HasActive) { 'no active certificate' } elseif ($Force) { '-Force' } else { 'expiring soon' }
                Write-Status "Generating certificate ($reason)..." -Type Warning
                $newCert = New-AuditCertificate -AppObjectId $app.Id -AppId $app.AppId -ExpiryYears $CertExpiryYears
                Write-Status "Certificate generated — expires $($newCert.ExpiryDate.ToString('yyyy-MM-dd'))." -Type Success
                Write-CredentialSummary -AppId $app.AppId -TenantId $TenantId -CertBase64 $newCert.CertBase64 -CertPassword $newCert.PlainPassword -CertExpiry $newCert.ExpiryDate
                Push-HuduAuditAsset     -AppId $app.AppId -TenantId $TenantId -CertBase64 $newCert.CertBase64 -CertPassword $newCert.PlainPassword -CertExpiry $newCert.ExpiryDate `
                    -HuduCompanyId $HuduCompanyId -HuduCompanyName $HuduCompanyName -HuduBaseUrl $HuduBaseUrl -HuduApiKey $HuduApiKey
            } else {
                $expiry = $certStatus.Soonest.EndDateTime
                Write-Status "Certificate healthy — expires $($expiry.ToString('yyyy-MM-dd'))." -Type Success
            }
        }
    }

    # ══════════════════════════════════════════════════════════════════════════
    # MODE 1 — No params: fully interactive, no Hudu push
    # ══════════════════════════════════════════════════════════════════════════
    else {
        $TenantId = Connect-GraphForSetup
        $orgName  = (Get-MgOrganization -ErrorAction Stop | Select-Object -First 1).DisplayName
        Write-Status "Tenant: $orgName ($TenantId)" -Type Success

        Write-Status "Searching for existing app: '$AppName'..."
        $app = Get-MgApplication -Filter "displayName eq '$AppName'" -ErrorAction Stop | Select-Object -First 1

        if ($app) {
            Write-Status "App found — App ID: $($app.AppId)" -Type Info
            $permResult = Resolve-PermissionState -App $app -TenantId $TenantId
            $app        = $permResult.App
            $TenantId   = $permResult.TenantId
            if ($permResult.OrgName) { $orgName = $permResult.OrgName }
            Invoke-OwnerCheck -App $app
        } else {
            $app = New-EntraApp -AppName $AppName -TenantId $TenantId -OrgName $orgName
        }

        $certStatus = Get-CertificateStatus -App $app
        if (-not $certStatus.HasActive) {
            Write-Status 'No active certificate — generating new certificate.' -Type Warning
            $newCert = New-AuditCertificate -AppObjectId $app.Id -AppId $app.AppId -ExpiryYears $CertExpiryYears
            Write-Status "Certificate generated — expires $($newCert.ExpiryDate.ToString('yyyy-MM-dd'))." -Type Success
            Write-CredentialSummary -AppId $app.AppId -TenantId $TenantId -CertBase64 $newCert.CertBase64 -CertPassword $newCert.PlainPassword -CertExpiry $newCert.ExpiryDate
        }
        elseif ($certStatus.ExpiresWithin30Days -or $Force) {
            $reason = if ($Force) { '-Force specified' } else { "expiring within $script:ExpiryWarnDays days" }
            Write-Status "Certificate $reason — generating new certificate." -Type Warning
            $newCert = New-AuditCertificate -AppObjectId $app.Id -AppId $app.AppId -ExpiryYears $CertExpiryYears
            Write-Status "Certificate renewed — expires $($newCert.ExpiryDate.ToString('yyyy-MM-dd'))." -Type Success
            Write-CredentialSummary -AppId $app.AppId -TenantId $TenantId -CertBase64 $newCert.CertBase64 -CertPassword $newCert.PlainPassword -CertExpiry $newCert.ExpiryDate
        }
        else {
            $expiry = $certStatus.Soonest.EndDateTime
            $sep = '=' * 72
            Write-Host "`n$sep" -ForegroundColor Cyan
            Write-Host "  App ID (Client ID) : $($app.AppId)"
            Write-Host "  Tenant ID          : $TenantId"
            Write-Host "  Cert Expires       : $($expiry.ToString('yyyy-MM-dd'))"
            Write-Host "  Use -Force to rotate the certificate regardless of expiry." -ForegroundColor DarkCyan
            Write-Host "$sep`n" -ForegroundColor Cyan
        }
    }
}
catch {
    Write-Status "Setup failed: $($_.Exception.Message)" -Type Error
    if ($VerbosePreference -eq 'Continue') {
        Write-Host $_.ScriptStackTrace -ForegroundColor Red
    }
    exit 1
}
finally {
    if (Get-MgContext) {
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    }
    foreach ($pfx in $script:GeneratedPfxPaths) {
        if (Test-Path $pfx) {
            Remove-Item $pfx -Force -ErrorAction SilentlyContinue
            Write-Verbose "Deleted temporary certificate file: $pfx"
        }
    }
}

