<#
.SYNOPSIS
    Shared helper functions for Microsoft 365 Audit modules.

.DESCRIPTION
    Provides common functionality used across all audit modules:
    - Connect-MgGraphSecure       : Connects to Microsoft Graph (app-only or interactive)
    - Connect-ExchangeOnlineSecure: Connects to Exchange Online (app-only or interactive)
    - Initialize-AuditOutput      : Creates and caches the per-run output folder
    - Invoke-VersionCheck         : Compares local script versions against the GitHub manifest

.NOTES
    Author      : Raymond Slater
    Version     : 1.24.0
    Change Log  : See CHANGELOG.md

.LINK
    https://github.com/razer86/365Audit
#>

#Requires -Version 7.2

$ScriptVersion = "1.24.0"
$RemoteBaseUrl = "https://raw.githubusercontent.com/razer86/365Audit/refs/heads/main"
Write-Verbose "Audit-Common.ps1 loaded (v$ScriptVersion)"

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8


# =====================================================================
# ===   Graph sub-module bootstrap (file scope — runs once on load) ===
# =====================================================================
# Runs when the launcher dot-sources this file, before any function calls.
# Install-only: never Import-Module here. Sub-module imports are deferred to
# Connect-MgGraphSecure, which runs them AFTER Connect-MgGraph so that
# Microsoft.Graph.Authentication is already in Get-Module when sub-modules
# resolve their RequiredModules — preventing any attempt to re-load the Auth DLL.
# Core Graph modules required by all audit runs:
#   Authentication       — the base SDK module
#   Identity.DirectoryManagement — needed by Initialize-AuditOutput (Get-MgOrganization)
#                                  and Connect-ExchangeOnlineSecure (org domain lookup)
# Audit-specific sub-modules are installed and imported on demand by each Invoke-* script
# via Ensure-GraphSubModules — keeping startup fast when only one module is selected.
$script:GraphCoreModules = @(
    'Microsoft.Graph.Authentication',
    'Microsoft.Graph.Identity.DirectoryManagement'
)
foreach ($_mod in $script:GraphCoreModules) {
    if (-not (Get-Module -ListAvailable -Name $_mod)) {
        Write-Host "Required module '$_mod' not found — installing latest..." -ForegroundColor Yellow
        Install-Module $_mod -Scope CurrentUser -Force -AllowClobber -SkipPublisherCheck -ErrorAction Stop
        $_installedMod = Get-Module -ListAvailable -Name $_mod | Sort-Object Version -Descending | Select-Object -First 1
        if (-not $_installedMod) {
            throw "Installation of '$_mod' failed — module still not found after install."
        }
        Write-Host "  Installed '$_mod' v$($_installedMod.Version)." -ForegroundColor Green
    }
}
Remove-Variable _mod -ErrorAction SilentlyContinue


function Resolve-GraphModuleVersion {
    [CmdletBinding()]
    param()

    if ($script:GraphModuleVersion) {
        return $script:GraphModuleVersion
    }

    # All Microsoft.Graph.* sub-modules are versioned in lockstep with the Authentication
    # module, so we use its installed version as the target for every sub-module install
    # and import. This avoids iterating every sub-module (many of which aren't installed
    # until their audit module runs) and keeps the resolver fast.
    $versions = @(Get-Module -ListAvailable -Name 'Microsoft.Graph.Authentication' |
        Select-Object -ExpandProperty Version -Unique |
        Sort-Object -Descending)

    if (-not $versions) {
        throw "Microsoft.Graph.Authentication is not installed. Run Start-365Audit.ps1 to install required modules."
    }

    $script:GraphModuleVersion = $versions | Select-Object -First 1
    Write-Verbose "Resolved Microsoft.Graph module version: $script:GraphModuleVersion"
    return $script:GraphModuleVersion
}


function Get-GraphModuleInfo {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ModuleName,

        [Parameter(Mandatory)]
        [version]$ModuleVersion
    )

    $normalizedModuleVersion = ConvertTo-NormalizedVersion -Version $ModuleVersion

    return Get-Module -ListAvailable -Name $ModuleName |
        Where-Object { (ConvertTo-NormalizedVersion -Version $_.Version) -eq $normalizedModuleVersion } |
        Select-Object -First 1
}


function Get-GraphDependencyAssemblyVersion {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ModuleName,

        [Parameter(Mandatory)]
        [version]$ModuleVersion,

        [Parameter(Mandatory)]
        [string]$AssemblyName
    )

    $moduleInfo = Get-GraphModuleInfo -ModuleName $ModuleName -ModuleVersion $ModuleVersion

    if (-not $moduleInfo) {
        return $null
    }

    $assemblyPath = Get-ChildItem -Path $moduleInfo.ModuleBase -Recurse -Filter "$AssemblyName.dll" -ErrorAction SilentlyContinue |
        Sort-Object FullName |
        Select-Object -First 1 -ExpandProperty FullName

    if (-not $assemblyPath) {
        return $null
    }

    return [Reflection.AssemblyName]::GetAssemblyName($assemblyPath).Version
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


function Get-GraphDependencyDirectories {
    [CmdletBinding()]
    param()

    if ($script:GraphDependencyDirectories) {
        return $script:GraphDependencyDirectories
    }

    $targetVersion = Resolve-GraphModuleVersion
    $authModule = Get-GraphModuleInfo -ModuleName 'Microsoft.Graph.Authentication' -ModuleVersion $targetVersion

    if (-not $authModule) {
        throw "Unable to locate Microsoft.Graph.Authentication $targetVersion."
    }

    $candidateDirs = @(
        (Join-Path $authModule.ModuleBase 'Dependencies\Core'),
        (Join-Path $authModule.ModuleBase 'Dependencies\Desktop'),
        (Join-Path $authModule.ModuleBase 'Dependencies'),
        $authModule.ModuleBase
    ) | Where-Object { Test-Path $_ }

    $script:GraphDependencyDirectories = $candidateDirs
    return $script:GraphDependencyDirectories
}


function Initialize-GraphDependencies {
    [CmdletBinding()]
    param()

    $assembliesToPrime = @(
        'Microsoft.Graph.Core',
        'Azure.Core',
        'Azure.Identity',
        'Microsoft.Kiota.Abstractions',
        'Microsoft.Kiota.Authentication.Azure',
        'Microsoft.Kiota.Http.HttpClientLibrary',
        'Microsoft.Kiota.Serialization.Json',
        'Microsoft.Kiota.Serialization.Form',
        'Microsoft.Kiota.Serialization.Text'
    )

    $alreadyLoaded = [AppDomain]::CurrentDomain.GetAssemblies() |
        Group-Object { $_.GetName().Name } -AsHashTable -AsString

    foreach ($assemblyName in $assembliesToPrime) {
        if ($alreadyLoaded.ContainsKey($assemblyName)) {
            continue
        }

        foreach ($dir in (Get-GraphDependencyDirectories)) {
            $candidatePath = Join-Path $dir ($assemblyName + '.dll')
            if (-not (Test-Path $candidatePath)) {
                continue
            }

            try {
                [System.Runtime.Loader.AssemblyLoadContext]::Default.LoadFromAssemblyPath($candidatePath) | Out-Null
                break
            }
            catch {
                Write-Verbose "Could not preload assembly '$assemblyName' from '$candidatePath': $_"
            }
        }
    }
}


function Import-GraphModuleVersioned {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ModuleName
    )

    $targetVersion = Resolve-GraphModuleVersion
    $normalizedTargetVersion = ConvertTo-NormalizedVersion -Version $targetVersion
    $targetModuleInfo = Get-GraphModuleInfo -ModuleName $ModuleName -ModuleVersion $targetVersion

    if (-not $targetModuleInfo) {
        throw "Unable to locate module '$ModuleName' version $targetVersion."
    }

    $loadedModules = @(Get-Module -Name $ModuleName -All)
    $mismatchedLoad = @($loadedModules | Where-Object {
            (ConvertTo-NormalizedVersion -Version $_.Version) -ne $normalizedTargetVersion
        })

    if ($mismatchedLoad) {
        $loadedSummary = $mismatchedLoad |
            Sort-Object Version -Descending |
            ForEach-Object { "{0} ({1})" -f $_.Name, $_.Version }
        throw ((("Microsoft Graph module version mismatch detected for '{0}': {1}. " +
            "Close this PowerShell session and start a new one so 365Audit can load Microsoft.Graph {2} consistently.") -f
            $ModuleName, ($loadedSummary -join ', '), $targetVersion))
    }

    if (-not ($loadedModules | Where-Object {
                (ConvertTo-NormalizedVersion -Version $_.Version) -eq $normalizedTargetVersion
            })) {
        Import-Module -Name $targetModuleInfo.Path -ErrorAction Stop
    }
}


function Import-GraphSubModules {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string[]]$Modules
    )

    $targetVersion = Resolve-GraphModuleVersion
    foreach ($modName in $Modules) {
        $installed = Get-Module -ListAvailable -Name $modName |
            Where-Object { (ConvertTo-NormalizedVersion $_.Version) -eq (ConvertTo-NormalizedVersion $targetVersion) }
        if (-not $installed) {
            Write-Host "  Required module '$modName' not found — installing v$targetVersion..." -ForegroundColor Yellow
            Install-Module $modName -RequiredVersion $targetVersion -Scope CurrentUser -Force -AllowClobber -SkipPublisherCheck -WarningAction SilentlyContinue -ErrorAction Stop
            $verified = Get-Module -ListAvailable -Name $modName |
                Where-Object { (ConvertTo-NormalizedVersion $_.Version) -eq (ConvertTo-NormalizedVersion $targetVersion) }
            if (-not $verified) {
                throw "Installation of '$modName' v$targetVersion failed — module still not found after install."
            }
            Write-Host "  Installed '$modName' v$targetVersion." -ForegroundColor Green
        }
        Import-GraphModuleVersioned -ModuleName $modName
    }
}


function Initialize-GraphSdk {
    [CmdletBinding()]
    param()

    $targetVersion = Resolve-GraphModuleVersion
    Initialize-GraphDependencies
    $expectedCore = Get-GraphDependencyAssemblyVersion -ModuleName 'Microsoft.Graph.Authentication' `
        -ModuleVersion $targetVersion `
        -AssemblyName 'Microsoft.Graph.Core'
    $expectedMsal = Get-GraphDependencyAssemblyVersion -ModuleName 'Microsoft.Graph.Authentication' `
        -ModuleVersion $targetVersion `
        -AssemblyName 'Microsoft.Identity.Client'

    $loadedCore = [AppDomain]::CurrentDomain.GetAssemblies() |
        Where-Object { $_.GetName().Name -eq 'Microsoft.Graph.Core' } |
        Select-Object -First 1
    $loadedMsal = [AppDomain]::CurrentDomain.GetAssemblies() |
        Where-Object { $_.GetName().Name -eq 'Microsoft.Identity.Client' } |
        Select-Object -First 1

    if ($loadedCore -and $expectedCore -and $loadedCore.GetName().Version -ne $expectedCore) {
        throw ((("Microsoft.Graph.Core {0} is already loaded in this PowerShell session, but 365Audit requires {1}. " +
            "Close this PowerShell session and start a new one before running 365Audit again.") -f
            $loadedCore.GetName().Version, $expectedCore))
    }

    if ($loadedMsal -and $expectedMsal -and $loadedMsal.GetName().Version -ne $expectedMsal) {
        throw ((("Microsoft.Identity.Client {0} is already loaded in this PowerShell session, but 365Audit requires {1} " +
            "from Microsoft.Graph.Authentication {2}. Another module likely loaded an incompatible MSAL version first " +
            "(commonly PnP.PowerShell or ExchangeOnlineManagement). Close this PowerShell session and start a new one before running 365Audit again.") -f
            $loadedMsal.GetName().Version, $expectedMsal, $targetVersion))
    }

    Import-GraphModuleVersioned -ModuleName 'Microsoft.Graph.Authentication'
}


function Get-GraphOrganizationSafe {
    [CmdletBinding()]
    param()

    try {
        $orgList = @(Get-MgOrganization -ErrorAction Stop)
    }
    catch {
        throw "Unable to query Microsoft Graph organization details: $($_.Exception.Message)"
    }

    if (-not $orgList) {
        throw "Microsoft Graph returned no organization objects. Verify the tenant connection and Organization.Read.All permission."
    }

    $primaryOrg = $orgList | Select-Object -First 1
    if (-not $primaryOrg.Id) {
        throw "Microsoft Graph organization lookup returned an object without an Id."
    }

    return $orgList
}


# ===============================================
# ===   Connect to Microsoft Graph Securely   ===
# ===============================================
function Connect-MgGraphSecure {
    [CmdletBinding()]
    param()

    Initialize-GraphSdk

    $requiredScopes = @(
        "User.Read.All",
        "Directory.Read.All",
        "Reports.Read.All",
        "Organization.Read.All",
        "Policy.Read.All",
        "Policy.ReadWrite.ConditionalAccess",
        "UserAuthenticationMethod.Read.All",
        "RoleManagement.Read.Directory",
        "Group.Read.All",
        "AuditLog.Read.All",
        "SecurityEvents.Read.All",
        "DeviceManagementManagedDevices.Read.All",
        "DeviceManagementConfiguration.Read.All",
        "DeviceManagementApps.Read.All",
        "DeviceManagementServiceConfig.Read.All"
    )

    # Auto-detect app credentials set by the launcher (-AppId/-TenantId/-CertBase64/-CertPassword params).
    # Get-Variable searches the current scope and all parent scopes.
    $appId        = Get-Variable -Name AuditAppId        -ValueOnly -ErrorAction SilentlyContinue
    $certFilePath = Get-Variable -Name AuditCertFilePath -ValueOnly -ErrorAction SilentlyContinue
    $certPassword = Get-Variable -Name AuditCertPassword -ValueOnly -ErrorAction SilentlyContinue
    $tenantId     = Get-Variable -Name AuditTenantId     -ValueOnly -ErrorAction SilentlyContinue

    $useAppAuth = $appId -and $certFilePath -and $tenantId

    if (-not (Get-MgContext)) {
        try {
            if ($useAppAuth) {
                Write-Host "Connecting to Microsoft Graph (app-only auth)..." -ForegroundColor Cyan
                $_cert = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($certFilePath, $certPassword)
                Connect-MgGraph -ClientId $appId -TenantId $tenantId -Certificate $_cert -NoWelcome -ErrorAction Stop
            }
            else {
                Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
                Connect-MgGraph -Scopes $requiredScopes -NoWelcome -ErrorAction Stop
            }
        }
        catch {
            Write-Error "Failed to connect to Microsoft Graph: $_"
            throw
        }

        # Scope validation only applies to delegated (interactive) auth.
        # App-only auth uses application permissions granted via admin consent — no scopes in context.
        if (-not $useAppAuth) {
            $context       = Get-MgContext
            $missingScopes = $requiredScopes | Where-Object { $_ -notin $context.Scopes }

            if ($missingScopes.Count -gt 0) {
                Write-Warning "Missing required Microsoft Graph scopes:"
                $missingScopes | ForEach-Object { Write-Warning "  - $_" }
                throw "Please re-run and ensure consent is granted for the missing scopes."
            }
        }

        Write-Host "Connected to Microsoft Graph." -ForegroundColor Green
    }
    else {
        Write-Verbose "Already connected to Microsoft Graph."
    }

    # Import the core sub-module required by Initialize-AuditOutput (Get-MgOrganization)
    # and Connect-ExchangeOnlineSecure (org domain lookup). This is the only module
    # imported unconditionally — audit-specific modules are loaded on demand by each
    # Invoke-* script via Import-GraphSubModules.
    Import-GraphModuleVersioned -ModuleName 'Microsoft.Graph.Identity.DirectoryManagement'
}


# =======================================================
# ===   Connect to Exchange Online (app-only or UI)   ===
# =======================================================
# Must be called AFTER Import-Module ExchangeOnlineManagement.
# Detects $AuditAppId/$AuditCertFilePath/$AuditCertPassword from the launcher scope.
# When present: connects via certificate-based app-only auth (-CertificateFilePath).
# Requires Exchange.ManageAsApp permission + Exchange Administrator Entra role on the SP.
# Falls back to interactive browser auth when credentials are absent.
function Connect-ExchangeOnlineSecure {
    [CmdletBinding()]
    param()

    # Skip if already connected
    $_exoConn = Get-ConnectionInformation -ErrorAction SilentlyContinue |
        Where-Object { $_.State -eq 'Connected' } |
        Select-Object -First 1
    if ($_exoConn) {
        Write-Verbose "Already connected to Exchange Online."
        return
    }

    # Auto-detect app credentials from the launcher scope
    $appId        = Get-Variable -Name AuditAppId        -ValueOnly -ErrorAction SilentlyContinue
    $certFilePath = Get-Variable -Name AuditCertFilePath -ValueOnly -ErrorAction SilentlyContinue
    $certPassword = Get-Variable -Name AuditCertPassword -ValueOnly -ErrorAction SilentlyContinue

    if ($appId -and $certFilePath) {
        Write-Host "Connecting to Exchange Online (app-only auth)..." -ForegroundColor Cyan

        # Resolve the tenant's initial .onmicrosoft.com domain for the -Organization parameter
        $_orgDomain = (Get-GraphOrganizationSafe).VerifiedDomains |
            Where-Object { $_.IsInitial -eq $true } |
            Select-Object -ExpandProperty Name -First 1

        Connect-ExchangeOnline `
            -AppId               $appId `
            -Organization        $_orgDomain `
            -CertificateFilePath $certFilePath `
            -CertificatePassword $certPassword `
            -ShowBanner:$false `
            -ErrorAction Stop

        Remove-Variable _orgDomain -ErrorAction SilentlyContinue
        Write-Host "Connected to Exchange Online." -ForegroundColor Green
    }
    else {
        Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
        Connect-ExchangeOnline -ShowBanner:$false
        Write-Host "Connected to Exchange Online." -ForegroundColor Green
    }
}


# ==========================================
# ===   Connect-TeamsSecure             ===
# ==========================================
# Connects to Microsoft Teams PowerShell using certificate-based app-only auth.
# Must be called after the MicrosoftTeams module is available.
# Detects $AuditAppId/$AuditTenantId/$AuditCertFilePath/$AuditCertPassword from launcher scope.
# When present: connects via -ApplicationId and -Certificate (no cert-store import required).
# Requires TeamSettings.Read.All Graph permission granted on the app registration.
# Falls back to interactive browser auth when credentials are absent.
function Connect-TeamsSecure {
    [CmdletBinding()]
    param()

    # Ensure MicrosoftTeams module is available
    if (-not (Get-Module -ListAvailable -Name 'MicrosoftTeams')) {
        Write-Host "Required module 'MicrosoftTeams' not found — installing latest..." -ForegroundColor Yellow
        Install-Module MicrosoftTeams -Scope CurrentUser -Force -AllowClobber -SkipPublisherCheck -ErrorAction Stop
        $_installedMod = Get-Module -ListAvailable -Name 'MicrosoftTeams' | Sort-Object Version -Descending | Select-Object -First 1
        if (-not $_installedMod) {
            throw "Installation of 'MicrosoftTeams' failed — module still not found after install."
        }
        Write-Host "  Installed 'MicrosoftTeams' v$($_installedMod.Version)." -ForegroundColor Green
    }

    Import-Module MicrosoftTeams -ErrorAction Stop -WarningAction SilentlyContinue

    # Check if already connected
    try {
        Get-CsTenant -ErrorAction Stop | Out-Null
        Write-Verbose "Already connected to Microsoft Teams."
        return
    }
    catch { <# not connected — continue #> }

    # Auto-detect app credentials from the launcher scope
    $appId        = Get-Variable -Name AuditAppId        -ValueOnly -ErrorAction SilentlyContinue
    $tenantId     = Get-Variable -Name AuditTenantId     -ValueOnly -ErrorAction SilentlyContinue
    $certFilePath = Get-Variable -Name AuditCertFilePath -ValueOnly -ErrorAction SilentlyContinue
    $certPassword = Get-Variable -Name AuditCertPassword -ValueOnly -ErrorAction SilentlyContinue

    if ($appId -and $tenantId -and $certFilePath) {
        Write-Host "Connecting to Microsoft Teams (app-only auth)..." -ForegroundColor Cyan

        # Load the X509Certificate2 directly from the .pfx — no cert-store import needed
        $certBytes = [System.IO.File]::ReadAllBytes($certFilePath)
        $_bstr     = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($certPassword)
        try {
            $plainPwd = [Runtime.InteropServices.Marshal]::PtrToStringAuto($_bstr)
            $certObj  = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new(
                $certBytes, $plainPwd,
                [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::Exportable)
        }
        finally {
            [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($_bstr)
        }

        Connect-MicrosoftTeams `
            -ApplicationId $appId `
            -TenantId      $tenantId `
            -Certificate   $certObj `
            -ErrorAction   Stop

        $certObj.Dispose()
        Write-Host "Connected to Microsoft Teams." -ForegroundColor Green
    }
    else {
        Write-Host "Connecting to Microsoft Teams (interactive)..." -ForegroundColor Cyan
        Connect-MicrosoftTeams -ErrorAction Stop
        Write-Host "Connected to Microsoft Teams." -ForegroundColor Green
    }
}


# ==========================================
# ===   Initialize Audit Output Folder   ===
# ==========================================
function Initialize-AuditOutput {
    [CmdletBinding()]
    param(
        # Optional override for the root folder where per-customer output folders are created.
        # Takes precedence over config.psd1 OutputRoot and the default (two levels above the toolkit).
        # Persisted in $script:AuditOutputRoot so subsequent calls from module scripts use it automatically.
        [string]$OutputRoot = ''
    )

    if ($OutputRoot) { $script:AuditOutputRoot = $OutputRoot }

    if ($script:AuditOutputContext) {
        return $script:AuditOutputContext
    }

    if (-not (Get-MgContext)) {
        Connect-MgGraphSecure
    }

    $orgList = Get-GraphOrganizationSafe

    if ($orgList.Count -gt 1) {
        $primaryDomain = $orgList[0].VerifiedDomains |
            Where-Object { $_.IsInitial -eq $true -and $_.Name -like "*.onmicrosoft.com" } |
            Select-Object -ExpandProperty Name
        Write-Warning "Multiple organisations detected ($($orgList.Count)). Using first: $($orgList[0].DisplayName) ($primaryDomain)"
    }

    $org      = $orgList | Select-Object -First 1
    $branding = Get-MgOrganizationBranding -OrganizationId $org.Id -ErrorAction SilentlyContinue

    $orgExpanded = [PSCustomObject]@{
        Id                          = $org.Id
        DisplayName                 = $org.DisplayName
        VerifiedDomains             = $org.VerifiedDomains
        TechnicalNotificationMails  = $org.TechnicalNotificationMails
        MarketingNotificationEmails = $org.MarketingNotificationEmails
        DefaultDomain               = $org.DefaultDomain
        CountryLetterCode           = $org.CountryLetterCode
        PreferredLanguageTag        = $org.PreferredLanguageTag
        ProvisionedPlans            = $org.ProvisionedPlans
        AssignedPlans               = $org.AssignedPlans
        Branding                    = $branding
        Raw                         = $org
    }

    $cleanDisplayName = $orgExpanded.DisplayName -replace '[^a-zA-Z0-9]', ''
    $folderName       = "${cleanDisplayName}_$(Get-Date -Format 'yyyyMMdd')"
    $outputDir        = if ($script:AuditOutputRoot) {
        Join-Path $script:AuditOutputRoot $folderName
    } else {
        Join-Path $PSScriptRoot "..\..\$folderName"
    }
    $rawOutputDir     = Join-Path $outputDir "Raw"
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
    New-Item -ItemType Directory -Path $rawOutputDir -Force | Out-Null

    $orgExpanded | ConvertTo-Json -Depth 10 |
        Set-Content -Path (Join-Path $outputDir "OrgInfo.json") -Encoding UTF8

    $script:AuditOutputContext = @{
        OrgName       = $orgExpanded.DisplayName
        FolderName    = $folderName
        OutputPath    = $outputDir
        RawOutputPath = $rawOutputDir
    }

    return $script:AuditOutputContext
}


# ==========================================
# ===   GitHub Version Check             ===
# ==========================================
function Invoke-VersionCheck {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$ScriptRoot
    )

    Write-Host "Checking for script updates..." -ForegroundColor Cyan

    try {
        $manifest = Invoke-RestMethod -Uri "$RemoteBaseUrl/version.json" -ErrorAction Stop
    }
    catch {
        Write-Warning "Unable to check for updates (network unavailable): $_"
        return
    }

    $outdated = @()

    foreach ($entry in $manifest.PSObject.Properties) {
        $relativeName  = $entry.Name
        $latestVersion = $entry.Value
        $localPath     = Join-Path $ScriptRoot $relativeName

        if (-not (Test-Path $localPath)) { continue }

        $versionLine = Select-String -Path $localPath -Pattern '^\$ScriptVersion\s*=\s*"(.+)"' |
            Select-Object -First 1

        if (-not $versionLine) { continue }

        $localVersion = $versionLine.Matches[0].Groups[1].Value

        try {
            if ([System.Version]$localVersion -lt [System.Version]$latestVersion) {
                $outdated += "  - $relativeName  (installed: $localVersion -> latest: $latestVersion)"
            }
        }
        catch {
            Write-Verbose "Could not compare versions for ${relativeName}: $_"
        }
    }

    if ($outdated.Count -gt 0) {
        Write-Host "Updates available for the following scripts:" -ForegroundColor Yellow
        $outdated | ForEach-Object { Write-Host $_ -ForegroundColor Yellow }
        Write-Host "  Download the latest at: https://github.com/razer86/365Audit`n" -ForegroundColor Yellow
    }
    else {
        Write-Host "All scripts are up to date.`n" -ForegroundColor Green
    }
}


# ==========================================
# ===   Audit Issue Logger              ===
# ==========================================
# Call from catch blocks in audit scripts to record collection failures,
# permission errors, and module issues into AuditIssues.csv in the output folder.
# Generate-AuditSummary.ps1 reads this file and renders it as a Technical Issues section.
function Add-AuditIssue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateSet('Critical', 'Warning', 'Info')]
        [string]$Severity,

        [Parameter(Mandatory)]
        [string]$Section,

        [Parameter(Mandatory)]
        [string]$Collector,

        [Parameter(Mandatory)]
        [string]$Description,

        [string]$Action = ''
    )

    if (-not $script:AuditOutputContext) {
        Write-Warning "Add-AuditIssue called before Initialize-AuditOutput — issue not recorded."
        return
    }

    $issuesPath = Join-Path $script:AuditOutputContext.OutputPath 'AuditIssues.csv'
    $row = [PSCustomObject]@{
        Timestamp   = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        Severity    = $Severity
        Section     = $Section
        Collector   = $Collector
        Description = $Description
        Action      = $Action
    }

    if (-not (Test-Path $issuesPath)) {
        $row | Export-Csv -Path $issuesPath -NoTypeInformation -Encoding UTF8
    } else {
        $row | Export-Csv -Path $issuesPath -NoTypeInformation -Encoding UTF8 -Append
    }
}
