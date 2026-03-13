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
    Version     : 1.11.0
    Change Log  :
        1.0.0 - Initial creation and migration of shared helpers from launcher
        1.1.0 - Added CmdletBinding, Invoke-VersionCheck, centralised RemoteBaseUrl
        1.2.0 - Import required Microsoft Graph sub-modules in Connect-MgGraphSecure
        1.3.0 - Moved Graph sub-module loading to file scope (wrong approach)
        1.4.0 - Reverted to install-only at file scope; never explicitly import sub-modules.
                SilentlyContinue does not catch terminating .NET FileLoadException so
                any Import-Module on a conflicting assembly caches the failure and blocks
                auto-loading. Modules are now installed if missing so PowerShell's own
                auto-loading can resolve them after Connect-MgGraph registers the Auth module.
        1.5.0 - Auto-loading sub-modules still fails because RequiredModules resolution
                re-triggers a DLL load even when the DLL is in the AppDomain. Fix: import
                sub-modules explicitly inside Connect-MgGraphSecure AFTER Connect-MgGraph,
                at which point Microsoft.Graph.Authentication is in Get-Module and the
                RequiredModules check finds it without attempting to reload the DLL.
        1.6.0 - Add AuditLog.Read.All to required Graph scopes (needed for sign-in logs)
        1.7.0 - Connect-MgGraphSecure auto-detects $AuditAppId/$AuditAppSecret/$AuditTenantId
                from the launcher scope; uses app-only (ClientSecretCredential) auth when all
                three are present, falls back to interactive delegated auth otherwise
        1.8.0 - Add Connect-ExchangeOnlineSecure: uses client-credentials OAuth token for
                Exchange Online when app credentials are present; falls back to interactive
        1.11.0 - Initialize-AuditOutput: move output folder from repo root to parent
                 directory (../.. from common/) to avoid git conflicts on update
        1.10.0 - Connect-ExchangeOnlineSecure: add missing -AppId to Connect-ExchangeOnline
                 -AccessToken call; EXO v3 requires -AppId alongside -AccessToken to
                 recognise the connection as app-only context (omitting it causes UnAuthorized)
        1.9.0 - Change "Already connected" Write-Host calls to Write-Verbose so they do
                not add scroll lines when modules reuse an existing session

.LINK
    https://github.com/razer86/365Audit
#>

#Requires -Version 7.2

$ScriptVersion = "1.11.0"
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
$_graphSubModules = @(
    'Microsoft.Graph.Authentication',
    'Microsoft.Graph.Identity.DirectoryManagement',
    'Microsoft.Graph.Users',
    'Microsoft.Graph.Groups',
    'Microsoft.Graph.Reports',
    'Microsoft.Graph.Identity.SignIns'
)
foreach ($_mod in $_graphSubModules) {
    if (-not (Get-Module -ListAvailable -Name $_mod)) {
        Write-Host "Installing $_mod..." -ForegroundColor Yellow
        Install-Module $_mod -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
    }
}
Remove-Variable _graphSubModules, _mod -ErrorAction SilentlyContinue


# ===============================================
# ===   Connect to Microsoft Graph Securely   ===
# ===============================================
function Connect-MgGraphSecure {
    [CmdletBinding()]
    param()

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
        "AuditLog.Read.All"
    )

    # Auto-detect app credentials set by the launcher (-AppId/-AppSecret/-TenantId params).
    # Get-Variable searches the current scope and all parent scopes.
    $appId     = Get-Variable -Name AuditAppId     -ValueOnly -ErrorAction SilentlyContinue
    $appSecret = Get-Variable -Name AuditAppSecret -ValueOnly -ErrorAction SilentlyContinue
    $tenantId  = Get-Variable -Name AuditTenantId  -ValueOnly -ErrorAction SilentlyContinue

    $useAppAuth = $appId -and $appSecret -and $tenantId

    if (-not (Get-MgContext)) {
        try {
            if ($useAppAuth) {
                Write-Host "Connecting to Microsoft Graph (app-only auth)..." -ForegroundColor Cyan
                $secureSecret = ConvertTo-SecureString $appSecret -AsPlainText -Force
                $credential   = New-Object System.Management.Automation.PSCredential($appId, $secureSecret)
                Connect-MgGraph -TenantId $tenantId -ClientSecretCredential $credential -NoWelcome -ErrorAction Stop
            }
            else {
                Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
                Connect-MgGraph -Scopes $requiredScopes -ErrorAction Stop
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

    # Import Graph sub-modules now that Microsoft.Graph.Authentication is registered in
    # Get-Module. Sub-modules resolve RequiredModules against the already-loaded Auth entry
    # and do NOT attempt to re-load its DLL — so no FileLoadException can occur.
    foreach ($mod in @(
            'Microsoft.Graph.Identity.DirectoryManagement',
            'Microsoft.Graph.Users',
            'Microsoft.Graph.Groups',
            'Microsoft.Graph.Reports',
            'Microsoft.Graph.Identity.SignIns')) {
        if (-not (Get-Module -Name $mod)) {
            Import-Module $mod -ErrorAction SilentlyContinue
        }
    }
}


# =======================================================
# ===   Connect to Exchange Online (app-only or UI)   ===
# =======================================================
# Must be called AFTER Import-Module ExchangeOnlineManagement.
# Detects $AuditAppId/$AuditAppSecret/$AuditTenantId from the launcher scope.
# When present: obtains an OAuth2 client-credentials token and connects via -AccessToken/-AppId.
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
    $appId     = Get-Variable -Name AuditAppId     -ValueOnly -ErrorAction SilentlyContinue
    $appSecret = Get-Variable -Name AuditAppSecret -ValueOnly -ErrorAction SilentlyContinue
    $tenantId  = Get-Variable -Name AuditTenantId  -ValueOnly -ErrorAction SilentlyContinue

    if ($appId -and $appSecret -and $tenantId) {
        Write-Host "Connecting to Exchange Online (app-only auth)..." -ForegroundColor Cyan

        # Resolve the tenant's initial .onmicrosoft.com domain for the -Organization parameter
        $_orgDomain = (Get-MgOrganization).VerifiedDomains |
            Where-Object { $_.IsInitial -eq $true } |
            Select-Object -ExpandProperty Name -First 1

        # Obtain an OAuth2 token for Exchange Online using client credentials.
        # -AppId must be passed alongside -AccessToken in EXO v3 to declare an app-only context;
        # omitting it causes an UnAuthorized rejection even with a valid token.
        $_tokenBody = @{
            grant_type    = 'client_credentials'
            scope         = 'https://outlook.office365.com/.default'
            client_id     = $appId
            client_secret = $appSecret
        }
        $_tokenResponse = Invoke-RestMethod `
            -Uri    "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" `
            -Method POST `
            -Body   $_tokenBody `
            -ErrorAction Stop

        $_secureToken = ConvertTo-SecureString $_tokenResponse.access_token -AsPlainText -Force
        Connect-ExchangeOnline -AccessToken $_secureToken -AppId $appId -Organization $_orgDomain -ShowBanner:$false -ErrorAction Stop

        Remove-Variable _orgDomain, _tokenBody, _tokenResponse, _secureToken -ErrorAction SilentlyContinue
    }
    else {
        Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
        Connect-ExchangeOnline -ShowBanner:$false
    }
}


# ==========================================
# ===   Initialize Audit Output Folder   ===
# ==========================================
function Initialize-AuditOutput {
    [CmdletBinding()]
    param()

    if ($script:AuditOutputContext) {
        return $script:AuditOutputContext
    }

    if (-not (Get-MgContext)) {
        Connect-MgGraphSecure
    }

    $orgList = Get-MgOrganization

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
    $outputDir        = Join-Path $PSScriptRoot "..\..\$folderName"
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null

    $orgExpanded | ConvertTo-Json -Depth 10 |
        Set-Content -Path (Join-Path $outputDir "OrgInfo.json") -Encoding UTF8

    $script:AuditOutputContext = @{
        OrgName    = $orgExpanded.DisplayName
        FolderName = $folderName
        OutputPath = $outputDir
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
