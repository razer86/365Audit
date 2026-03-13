<#
.SYNOPSIS
    Performs a SharePoint Online and OneDrive audit.

.DESCRIPTION
    Connects to SharePoint Online using PnP.PowerShell with interactive authentication
    and exports:
    - Tenant storage summary
    - Site collection list (URL, template, storage, owner)
    - SharePoint groups and members per site
    - Site users and admin status per site
    - External sharing configuration (tenant-wide and per-site overrides)
    - Per-user OneDrive storage usage
    - Unlicensed OneDrive accounts
    - Access control policies

    Output CSVs are written to the shared audit output folder.

.NOTES
    Author      : Raymond Slater
    Version     : 2.7.0
    Change Log  :
        1.0.0 - Initial release
        1.0.1 - Refactor output directory initialisation
        1.0.2 - Helper function refactor
        1.1.0 - Removed duplicate guard clause; fixed outputDir override;
                replaced deprecated Get-MsolUser with Microsoft Graph;
                removed alias usage; added CmdletBinding
        1.2.0 - Fixed Connect-PnPOnline (removed invalid -Scopes param);
                derive SharePoint admin URL from tenant .onmicrosoft.com domain;
                replaced Get-SPOSite/Get-SPOTenant (SPO Management Shell) with
                PnP equivalents (Get-PnPTenantSite, Get-PnPTenant)
        1.3.0 - PnP.PowerShell v2+ requires -ClientId with -Interactive;
                using PnP Management Shell public app (31359c7f-bd7e-475c-86db-fdb8c937548e)
        1.4.0 - Conditional auth: app-only (ClientId/Secret) when launcher provides
                -AppId/-AppSecret/-TenantId; falls back to interactive PnP Management Shell
        1.5.0 - Extend ExternalSharing_Tenant CSV with DefaultSharingLinkType and
                RequireAnonymousLinksExpireInDays; extend AccessControlPolicies CSV
                with IsUnmanagedSyncAppForTenantRestricted and BlockMacSync
        1.6.0 - Replaced PnP.PowerShell with OAuth2 client-credentials token approach
        1.7.0 - Replaced PnP.PowerShell entirely with Microsoft.Online.SharePoint.PowerShell
                (the official Microsoft SPO admin module); Connect-SPOService -AccessToken
                correctly handles tenant admin APIs; all PnP cmdlets replaced with SPO equivalents
        1.8.0 - Reverted to PnP.PowerShell; connect to admin URL for tenant-wide operations;
                reconnect per-site for group/user queries
        1.9.0 - Replaced all module dependencies with direct SharePoint REST API calls;
                enum mapping functions for string compatibility
        2.0.0 - Root cause identified: SharePoint admin APIs block tokens with azpacr=0
                (client-secret credentials) with "Unsupported app only token" regardless of
                REST vs CSOM; switched to PnP.PowerShell with certificate-based app-only auth
                which produces azpacr=1 tokens; requires $AuditCertThumbprint from launcher
        2.1.0 - Reverted to interactive authentication; certificate-based app-only auth is
                not portable across technician machines; interactive sign-in is sufficient
                for manual monthly audit runs; MSAL token cache prevents repeated prompts
                for per-site reconnections
        2.2.0 - Added pre-flight check: verifies PnP Management Shell is registered in the
                tenant before attempting interactive auth; prints setup guidance if missing
        2.3.0 - PnP Management Shell app deprecated in PnP.PowerShell v2; switched to
                using $AuditPnPAppId (registered via Register-PnPEntraIDAppForInteractiveLogin
                in Setup-365AuditApp.ps1) as the ClientId for Connect-PnPOnline -Interactive;
                requires Setup to have been run and -PnPAppId provided at launch;
                bumped #Requires to 7.4 (PnP.PowerShell v3 requires PowerShell 7.4+);
                updated PnP module check to enforce MinimumVersion 3.0.0
        2.4.0 - Replaced per-section Write-Host progress lines with Write-Progress
                for cleaner terminal output
        2.5.0 - Per-site connections use -ReturnConnection so each connection is a
                named object; admin cmdlets pin to $adminConn, per-site cmdlets pin
                to $siteConn; guarantees exactly one browser sign-in (MSAL reuses
                the cached token for all subsequent site connections silently)
        2.6.0 - Added Step X/Y counter to Write-Progress status strings
        2.7.0 - Replaced -ReturnConnection MSAL caching strategy with explicit
                -AccessToken pass-through: authenticate interactively once to the
                admin URL, capture the SPO access token via Get-PnPAccessToken, then
                connect to each site with -AccessToken (no browser prompt per site);
                also removed Disconnect-PnPOnline -Connection which is not a valid
                parameter in PnP.PowerShell v3

.LINK
    https://github.com/razer86/365Audit
#>

#Requires -Version 7.4

param (
    [string]$AuditFolder,
    [switch]$DevMode = $false
)

if (-not $DevMode -and $MyInvocation.InvocationName -eq $MyInvocation.MyCommand.Name) {
    Write-Error "This script must be run from the 365Audit launcher. Use -DevMode for development." -ErrorAction Stop
}

$ScriptVersion = "2.7.0"
Write-Verbose "Invoke-SharePointAudit.ps1 loaded (v$ScriptVersion)"

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# === Retrieve shared output folder ===
try {
    $context   = Initialize-AuditOutput
    $outputDir = $context.OutputPath
}
catch {
    Write-Error "Failed to initialise audit output directory: $_"
    exit 1
}

# === Pre-flight: verify PnP interactive auth app ID is available ===
# Setup-365AuditApp.ps1 must be run once per tenant to register the PnP interactive
# auth app (via Register-PnPEntraIDAppForInteractiveLogin) and save the App ID.
# Pass it via Start-365Audit.ps1 -PnPAppId (sets $AuditPnPAppId in launcher scope).
$pnpClientId = Get-Variable -Name AuditPnPAppId -ValueOnly -ErrorAction SilentlyContinue
if (-not $pnpClientId) {
    Write-Host ""
    Write-Host "  ┌───────────────────────────────────────────────────────────────────────┐" -ForegroundColor Yellow
    Write-Host "  │  SharePoint Audit — one-time setup required for this tenant           │" -ForegroundColor Yellow
    Write-Host "  │                                                                       │" -ForegroundColor Yellow
    Write-Host "  │  A registered PnP interactive auth app is required.                  │" -ForegroundColor Yellow
    Write-Host "  │  Run the following as a Global Administrator in this tenant:          │" -ForegroundColor Yellow
    Write-Host "  │    .\Setup-365AuditApp.ps1                                            │" -ForegroundColor Yellow
    Write-Host "  │                                                                       │" -ForegroundColor Yellow
    Write-Host "  │  Then provide all credentials at audit runtime:                      │" -ForegroundColor Yellow
    Write-Host "  │    .\Start-365Audit.ps1 -AppId <id> -AppSecret <secret> \            │" -ForegroundColor Yellow
    Write-Host "  │       -TenantId <id> -PnPAppId <pnp-id>                              │" -ForegroundColor Yellow
    Write-Host "  └───────────────────────────────────────────────────────────────────────┘" -ForegroundColor Yellow
    Write-Host ""
    return
}

# === Ensure PnP.PowerShell v3+ is available ===
if (-not (Get-Module -ListAvailable -Name PnP.PowerShell | Where-Object Version -ge '3.0.0')) {
    Write-Host "Installing PnP.PowerShell v3+..." -ForegroundColor Yellow
    Install-Module PnP.PowerShell -Scope CurrentUser -Force -AllowClobber
}
Import-Module PnP.PowerShell -WarningAction SilentlyContinue

# === Derive SharePoint admin URL from the tenant's .onmicrosoft.com domain ===
$initialDomain = (Get-MgOrganization).VerifiedDomains |
    Where-Object { $_.IsInitial -eq $true } |
    Select-Object -ExpandProperty Name -First 1
$spoAdminUrl = "https://$(($initialDomain -split '\.')[0])-admin.sharepoint.com"

# === Connect once — capture the SPO access token for silent per-site reuse ===
# Interactive auth prompts the browser exactly once. Get-PnPAccessToken then captures
# the resulting SharePoint token (aud: https://<tenant>.sharepoint.com) which is valid
# for all site URLs in the tenant. Each per-site Connect-PnPOnline uses this token
# directly — no additional browser prompts regardless of site count.
Write-Host "Connecting to SharePoint Online ($spoAdminUrl)..." -ForegroundColor Cyan
Write-Host "  Sign in once — site connections reuse this token silently." -ForegroundColor DarkCyan
Connect-PnPOnline `
    -Url      $spoAdminUrl `
    -ClientId $pnpClientId `
    -Interactive `
    -ErrorAction Stop

$spAccessToken = Get-PnPAccessToken -ResourceTypeName SharePoint

Write-Host "`nStarting SharePoint Online Audit for $($context.OrgName)..." -ForegroundColor Cyan

$step       = 0
$totalSteps = 8
$activity   = "SharePoint Audit — $($context.OrgName)"


# === 1. Tenant Storage ===
$step++
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Gathering tenant storage information..." -PercentComplete ([int]($step / $totalSteps * 100))

# $tenantConfig is reused across sections 1, 5, 7, and 8
$tenantConfig = Get-PnPTenant
[PSCustomObject]@{
    StorageQuotaMB     = $tenantConfig.StorageQuota
    StorageUsedMB      = $tenantConfig.StorageQuotaUsed
    AvailableStorageMB = $tenantConfig.StorageQuota - $tenantConfig.StorageQuotaUsed
    OneDriveQuotaMB    = $tenantConfig.OneDriveStorageQuota
} | Export-Csv "$outputDir\SharePoint_TenantStorage.csv" -NoTypeInformation -Encoding UTF8
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Gathering tenant storage information..." -CurrentOperation "Saved: SharePoint_TenantStorage.csv" -PercentComplete ([int]($step / $totalSteps * 100))


# === 2. Site Collection List ===
$step++
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Gathering site collection details..." -PercentComplete ([int]($step / $totalSteps * 100))

$sites      = Get-PnPTenantSite
$siteReport = foreach ($site in $sites) {
    [PSCustomObject]@{
        Title         = $site.Title
        Url           = $site.Url
        Template      = $site.Template
        StorageUsedMB = $site.StorageUsageCurrent
        IsHubSite     = $site.IsHubSite
        Owner         = $site.Owner
    }
}
$siteReport | Export-Csv "$outputDir\SharePoint_Sites.csv" -NoTypeInformation -Encoding UTF8
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Gathering site collection details..." -CurrentOperation "Saved: SharePoint_Sites.csv" -PercentComplete ([int]($step / $totalSteps * 100))

# Pre-fetch OneDrive sites while still on the admin connection
$oneDriveSites = Get-PnPTenantSite -IncludeOneDriveSites |
    Where-Object { $_.Template -like 'SPSPERS*' }

# Disconnect from admin — per-site connections use the captured access token
Disconnect-PnPOnline


# === 3 & 4. SharePoint Groups, Members, and Site Users (combined per-site loop) ===
$step++
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Retrieving SharePoint groups, members, and site users..." -PercentComplete ([int]($step / $totalSteps * 100))

$groupReport      = [System.Collections.Generic.List[PSCustomObject]]::new()
$permissionReport = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($site in $sites) {
    try {
        # Connects silently using the captured token — no browser prompt
        Connect-PnPOnline -Url $site.Url -AccessToken $spAccessToken -ErrorAction Stop

        # Groups and members
        $groups = Get-PnPGroup -ErrorAction Stop
        foreach ($group in $groups) {
            $members = Get-PnPGroupMember -Group $group -ErrorAction SilentlyContinue
            $groupReport.Add([PSCustomObject]@{
                Site        = $site.Url
                GroupName   = $group.Title
                Owner       = $group.OwnerTitle
                MemberCount = if ($members) { @($members).Count } else { 0 }
                Members     = if ($members) { ($members.LoginName -join '; ') } else { '' }
            })
        }

        # Site users
        $users = Get-PnPUser -ErrorAction Stop
        foreach ($user in $users) {
            $permissionReport.Add([PSCustomObject]@{
                Site        = $site.Url
                Principal   = $user.Title
                LoginName   = $user.LoginName
                IsSiteAdmin = $user.IsSiteAdmin
                Type        = if ($user.PrincipalType -eq 'User') { 'User' } else { 'Group' }
            })
        }

        Disconnect-PnPOnline
    }
    catch {
        Write-Warning "Could not get groups/users for site: $($site.Url)"
    }
}

$groupReport      | Export-Csv "$outputDir\SharePoint_SPGroups.csv"       -NoTypeInformation -Encoding UTF8
$permissionReport | Export-Csv "$outputDir\SharePoint_SitePermissions.csv" -NoTypeInformation -Encoding UTF8
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Retrieving SharePoint groups, members, and site users..." -CurrentOperation "Saved: SharePoint_SPGroups.csv, SharePoint_SitePermissions.csv" -PercentComplete ([int]($step / $totalSteps * 100))


# === 5. Tenant-Wide External Sharing ===
$step++
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Checking external sharing configuration..." -PercentComplete ([int]($step / $totalSteps * 100))

$tenantConfig |
    Select-Object -Property SharingCapability, SharingDomainRestrictionMode, SharingAllowedDomainList, DefaultSharingLinkType, RequireAnonymousLinksExpireInDays |
    Export-Csv "$outputDir\SharePoint_ExternalSharing_Tenant.csv" -NoTypeInformation -Encoding UTF8
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Checking external sharing configuration..." -CurrentOperation "Saved: SharePoint_ExternalSharing_Tenant.csv" -PercentComplete ([int]($step / $totalSteps * 100))


# === 6. Per-User OneDrive Storage ===
$step++
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Gathering OneDrive storage usage per user..." -PercentComplete ([int]($step / $totalSteps * 100))

$oneDriveReport = foreach ($site in $oneDriveSites) {
    [PSCustomObject]@{
        OwnerUPN      = $site.Owner
        OneDriveUrl   = $site.Url
        StorageUsedMB = $site.StorageUsageCurrent
    }
}
$oneDriveReport | Export-Csv "$outputDir\SharePoint_OneDriveUsage.csv" -NoTypeInformation -Encoding UTF8
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Gathering OneDrive storage usage per user..." -CurrentOperation "Saved: SharePoint_OneDriveUsage.csv" -PercentComplete ([int]($step / $totalSteps * 100))


# === 7. Site-Level External Sharing Overrides ===
$step++
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Checking per-site external sharing overrides..." -PercentComplete ([int]($step / $totalSteps * 100))

$siteOverrides     = $sites | Where-Object { $_.SharingCapability -ne $tenantConfig.SharingCapability }
$siteSharingReport = foreach ($site in $siteOverrides) {
    [PSCustomObject]@{
        Url               = $site.Url
        Title             = $site.Title
        SharingCapability = $site.SharingCapability
        SiteStorageMB     = $site.StorageUsageCurrent
    }
}
$siteSharingReport | Export-Csv "$outputDir\SharePoint_ExternalSharing_SiteOverrides.csv" -NoTypeInformation -Encoding UTF8
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Checking per-site external sharing overrides..." -CurrentOperation "Saved: SharePoint_ExternalSharing_SiteOverrides.csv" -PercentComplete ([int]($step / $totalSteps * 100))


# === 8. Access Control Policies + Unlicensed OneDrive Accounts ===
$step++
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Collecting access control policies..." -PercentComplete ([int]($step / $totalSteps * 100))

[PSCustomObject]@{
    AllowEditing                          = $tenantConfig.AllowEditing
    ConditionalAccessPolicy               = $tenantConfig.ConditionalAccessPolicy
    LimitedAccessFileType                 = $tenantConfig.LimitedAccessFileType
    IPAddressEnforcement                  = $tenantConfig.IPAddressEnforcement
    BypassAppsForManagedDevices           = $tenantConfig.BypassAppLockerForManagedDevices
    IdleSessionSignOutEnabled             = $tenantConfig.IdleSessionSignOut
    SignOutAfterMinutesOfInactivity       = $tenantConfig.InactiveBrowserSessionTimeout
    IsUnmanagedSyncAppForTenantRestricted = $tenantConfig.IsUnmanagedSyncAppForTenantRestricted
    BlockMacSync                          = $tenantConfig.BlockMacSync
} | Export-Csv "$outputDir\SharePoint_AccessControlPolicies.csv" -NoTypeInformation -Encoding UTF8
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Collecting access control policies..." -CurrentOperation "Saved: SharePoint_AccessControlPolicies.csv" -PercentComplete ([int]($step / $totalSteps * 100))

$licensedUPNs = (Get-MgUser -All -Property UserPrincipalName, AssignedLicenses |
    Where-Object { $_.AssignedLicenses.Count -gt 0 }).UserPrincipalName

$unlicensedOneDrives = $oneDriveReport |
    Where-Object { $_.OwnerUPN -and ($_.OwnerUPN -notin $licensedUPNs) }

$unlicensedOneDrives | Export-Csv "$outputDir\SharePoint_OneDrive_Unlicensed.csv" -NoTypeInformation -Encoding UTF8

if ($unlicensedOneDrives.Count -gt 0) {
    Write-Warning "Found $($unlicensedOneDrives.Count) OneDrive accounts without active licences."
}
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Collecting access control policies..." -CurrentOperation "Saved: SharePoint_OneDrive_Unlicensed.csv" -PercentComplete 100


# ================================
# ===   Done                    ===
# ================================
Write-Progress -Id 1 -Activity $activity -Completed
Write-Host "`nSharePoint Online Audit complete. Results saved to: $outputDir`n" -ForegroundColor Green
