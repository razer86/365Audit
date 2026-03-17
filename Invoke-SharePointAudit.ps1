<#
.SYNOPSIS
    Performs a SharePoint Online and OneDrive audit.

.DESCRIPTION
    Connects to SharePoint Online using PnP.PowerShell with certificate-based app-only
    authentication and exports:
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
    Version     : 2.10.0
    Change Log  : See CHANGELOG.md

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

$ScriptVersion = "2.10.0"
Write-Verbose "Invoke-SharePointAudit.ps1 loaded (v$ScriptVersion)"

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# === Retrieve shared output folder ===
try {
    $context   = Initialize-AuditOutput
    $outputDir = $context.RawOutputPath
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
}
catch {
    Write-Error "Failed to initialise audit output directory: $_"
    exit 1
}

# === Ensure PnP.PowerShell v3+ is available ===
if (-not (Get-Module -ListAvailable -Name PnP.PowerShell | Where-Object Version -ge '3.0.0')) {
    Write-Host "Installing PnP.PowerShell v3+..." -ForegroundColor Yellow
    Install-Module PnP.PowerShell -Scope CurrentUser -Force -AllowClobber
}
Import-Module PnP.PowerShell -WarningAction SilentlyContinue

# === Auto-detect app credentials from the launcher scope ===
$_spAppId        = Get-Variable -Name AuditAppId        -ValueOnly -ErrorAction SilentlyContinue
$_spCertFilePath = Get-Variable -Name AuditCertFilePath -ValueOnly -ErrorAction SilentlyContinue
$_spCertPassword = Get-Variable -Name AuditCertPassword -ValueOnly -ErrorAction SilentlyContinue
$_useAppAuth     = $_spAppId -and $_spCertFilePath

# === Derive SharePoint admin URL from the tenant's .onmicrosoft.com domain ===
$initialDomain = (Get-MgOrganization).VerifiedDomains |
    Where-Object { $_.IsInitial -eq $true } |
    Select-Object -ExpandProperty Name -First 1
$spoAdminUrl = "https://$(($initialDomain -split '\.')[0])-admin.sharepoint.com"

# === Connect to SharePoint admin URL ===
Write-Host "Connecting to SharePoint Online ($spoAdminUrl)..." -ForegroundColor Cyan
if ($_useAppAuth) {
    Connect-PnPOnline `
        -Url                 $spoAdminUrl `
        -ClientId            $_spAppId `
        -Tenant              $initialDomain `
        -CertificatePath     $_spCertFilePath `
        -CertificatePassword $_spCertPassword `
        -ErrorAction Stop
}
else {
    Write-Host "  No app credentials found — falling back to interactive sign-in." -ForegroundColor DarkCyan
    Connect-PnPOnline -Url $spoAdminUrl -Interactive -ErrorAction Stop
}
Write-Host "Connected to SharePoint Online." -ForegroundColor Green

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

Disconnect-PnPOnline


# === 3 & 4. SharePoint Groups, Members, and Site Users (combined per-site loop) ===
$step++
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Retrieving SharePoint groups, members, and site users..." -PercentComplete ([int]($step / $totalSteps * 100))

$groupReport      = [System.Collections.Generic.List[PSCustomObject]]::new()
$permissionReport = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($site in $sites) {
    try {
        if ($_useAppAuth) {
            Connect-PnPOnline -Url $site.Url -ClientId $_spAppId -Tenant $initialDomain `
                -CertificatePath $_spCertFilePath -CertificatePassword $_spCertPassword -ErrorAction Stop
        }
        else {
            Connect-PnPOnline -Url $site.Url -Interactive -ErrorAction Stop
        }

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
