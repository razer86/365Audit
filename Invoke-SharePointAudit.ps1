param (
    [string]$AuditFolder,
    [switch]$DevMode = $false
)

if (-not $DevMode -and $MyInvocation.InvocationName -eq $MyInvocation.MyCommand.Name) {
    Write-Error "This script must be run from the 365Audit launcher. Use -DevMode for development." -ErrorAction Stop
}


if (-not $DevMode -and $MyInvocation.InvocationName -eq $MyInvocation.MyCommand.Name) {
    Write-Error "This script must be run from the 365Audit launcher. Use -DevMode for development." -ErrorAction Stop
}

<#
.SYNOPSIS
    Performs a SharePoint Online audit

.DESCRIPTION
    This script connects to SharePoint Online using the PnP.PowerShell module and exports the following:
    - Total available and used storage across the tenant
    - List of SharePoint sites (URL, type, storage used)
    - SharePoint groups and their members/owners
    - Site-level permissions
    - Default external sharing settings

    Results are saved in CSV format to a timestamped folder named after the organization.

.NOTES
    Author      : Raymond Slater
    Version     : 1.0.1
    Change Log  :
        1.0.0 - Initial release
        1.0.1 - Refactor Output directory initialisation
        1.0.2 - Helper function refactor
        
.LINK
    https://github.com/razer86/365Audit
#>

# === Retrieve Output Folder ===
try {
    $context = Initialize-AuditOutput
    $outputDir = $context.OutputPath
}
catch {
    Write-Error "❌ Failed to locate audit output directory: $_"
    exit 1
}

# === Check for PnP.PowerShell Module ===
if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    Write-Host "📦 Installing PnP.PowerShell module..." -ForegroundColor Yellow
    Install-Module PnP.PowerShell -Scope CurrentUser -Force
}
Import-Module PnP.PowerShell

# === Connect to SharePoint Online ===
Write-Host "🔐 Connecting to SharePoint Online..."
Connect-PnPOnline -Scopes "Sites.Read.All", "Group.Read.All" -Interactive

# === Setup output folder ===
$tenant = Get-PnPTenant
$companyName = $tenant.TenantId -replace '[^a-zA-Z0-9]', '_'
$timestamp = Get-Date -Format "yyyyMMdd-HHmm"
$outputDir = "${companyName}_SharePointAudit_$timestamp"
New-Item -ItemType Directory -Path $outputDir -Force | Out-Null

Write-Host "`n📁 Starting SharePoint Online Audit for $companyName`n"

# === 1. Tenant Storage Info ===
Write-Host "✔ Gathering tenant storage information..."
$tenantStorage = Get-PnPTenant

[PSCustomObject]@{
    StorageQuotaMB     = $tenantStorage.StorageQuota
    StorageUsedMB      = $tenantStorage.StorageQuotaUsed
    AvailableStorageMB = $tenantStorage.StorageQuota - $tenantStorage.StorageQuotaUsed
    OneDriveQuotaMB    = $tenantStorage.OneDriveStorageQuota
} | Export-Csv "$outputDir\Sharepoint_TenantStorage.csv" -NoTypeInformation

# === 2. Site List ===
Write-Host "✔ Gathering site collection details..."
$sites = Get-PnPTenantSite

$siteReport = foreach ($site in $sites) {
    [PSCustomObject]@{
        Title        = $site.Title
        Url          = $site.Url
        Template     = $site.Template
        StorageUsedMB = $site.StorageUsage
        IsHubSite    = $site.IsHubSite
        Owner        = $site.Owner
    }
}
$siteReport | Export-Csv "$outputDir\Sharepoint_Sites.csv" -NoTypeInformation

# === 3. SharePoint Groups and Members ===
Write-Host "✔ Retrieving SharePoint groups and members..."

$groupReport = @()
foreach ($site in $sites) {
    try {
        Connect-PnPOnline -Url $site.Url -Interactive
        $spGroups = Get-PnPGroup
        foreach ($group in $spGroups) {
            $members = Get-PnPGroupMember -Identity $group.Title -ErrorAction SilentlyContinue
            $groupReport += [PSCustomObject]@{
                Site       = $site.Url
                GroupName  = $group.Title
                Owner      = $group.OwnerTitle
                MemberCount = $members.Count
                Members    = ($members.Title -join "; ")
            }
        }
    }
    catch {
        Write-Warning "Could not get groups for site: $($site.Url)"
    }
}
$groupReport | Export-Csv "$outputDir\Sharepoint_SPGroups.csv" -NoTypeInformation

# === 4. Site-Level Permissions ===
Write-Host "✔ Collecting site-level permissions..."

$permissionReport = @()
foreach ($site in $sites) {
    try {
        Connect-PnPOnline -Url $site.Url -Interactive
        $roles = Get-PnPRoleAssignment
        foreach ($role in $roles) {
            $permissionReport += [PSCustomObject]@{
                Site     = $site.Url
                Principal = $role.PrincipalName
                Roles    = ($role.RoleDefinitionBindings.Name -join ", ")
            }
        }
    }
    catch {
        Write-Warning "Could not get permissions for site: $($site.Url)"
    }
}
$permissionReport | Export-Csv "$outputDir\Sharepoint_SitePermissions.csv" -NoTypeInformation

# === 5. External Sharing Settings ===
Write-Host "✔ Checking external sharing configuration..."
$sharingSettings = Get-PnPTenant | Select SharingCapability, SharingDomainRestrictionMode, SharingAllowedDomainList
$sharingSettings | Export-Csv "$outputDir\Sharepoint_ExternalSharingSettings.csv" -NoTypeInformation


# === 6. Per-User OneDrive Storage ===
Write-Host "✔ Gathering OneDrive storage usage per user..."

# OneDrive sites use the "SPSPERS" template
$oneDriveSites = Get-SPOSite -IncludePersonalSite $true -Limit All | Where-Object { $_.Template -eq "SPSPERS" }

$oneDriveReport = foreach ($site in $oneDriveSites) {
    [PSCustomObject]@{
        OwnerUPN       = $site.Owner
        OneDriveUrl    = $site.Url
        StorageUsedMB  = $site.StorageUsageCurrent
    }
}

$oneDriveReport | Export-Csv "$outputDir\Sharepoint_OneDriveUsage.csv" -NoTypeInformation

# === 7. External Sharing Configuration ===
Write-Host "✔ Collecting external sharing configuration..."

# Tenant-wide sharing
$tenantSharing = Get-SPOTenant | Select SharingCapability, SharingDomainRestrictionMode, SharingAllowedDomainList
$tenantSharing | Export-Csv "$outputDir\Sharepoint_ExternalSharing_Tenant.csv" -NoTypeInformation

# Site-level overrides (only sites where sharing differs from tenant policy)
$siteOverrides = Get-SPOSite -Limit All | Where-Object { $_.SharingCapability -ne $tenantSharing.SharingCapability }

$siteSharingReport = foreach ($site in $siteOverrides) {
    [PSCustomObject]@{
        Url                = $site.Url
        Title              = $site.Title
        SharingCapability  = $site.SharingCapability
        SiteStorageMB      = $site.StorageUsageCurrent
    }
}
$siteSharingReport | Export-Csv "$outputDir\Sharepoint_ExternalSharing_SiteOverrides.csv" -NoTypeInformation



# === 8. SharePoint Access Control Policies ===
Write-Host "✔ Collecting access control policies..."

$tenantConfig = Get-SPOTenant

[PSCustomObject]@{
    AllowEditing                          = $tenantConfig.AllowEditing
    ConditionalAccessPolicy               = $tenantConfig.ConditionalAccessPolicy
    LimitedAccessFileType                 = $tenantConfig.LimitedAccessFileType
    IPAddressEnforcement                  = $tenantConfig.IPAddressEnforcement
    BypassAppsForManagedDevices           = $tenantConfig.BypassAppLockerForManagedDevices
    IdleSessionSignOutEnabled             = $tenantConfig.UseIdleSessionSignOut
    SignOutAfterMinutesOfInactivity       = $tenantConfig.InactiveBrowserSessionTimeout
} | Export-Csv "$outputDir\Sharepoint_AccessControlPolicies.csv" -NoTypeInformation



# === 9. Unlicensed OneDrive Accounts ===
Write-Host "✔ Checking for unlicensed OneDrive users..."

# Get all licensed users
$licensedUsers = Get-MsolUser -All | Where-Object { $_.isLicensed } | Select-Object UserPrincipalName
$licensedUPNs = $licensedUsers.UserPrincipalName

# Compare against OneDrive owners
$unlicensedOneDrives = $oneDriveReport | Where-Object {
    $_.OwnerUPN -and ($_.OwnerUPN -notin $licensedUPNs)
}

$unlicensedOneDrives | Export-Csv "$outputDir\Sharepoint_OneDrive_Unlicensed.csv" -NoTypeInformation

if ($unlicensedOneDrives.Count -gt 0) {
    Write-Host "`n⚠ Found $($unlicensedOneDrives.Count) OneDrive accounts without active licenses." -ForegroundColor Yellow
}













Write-Host "`n✅ SharePoint Online Audit Complete. Results saved to: $outputDir`n"