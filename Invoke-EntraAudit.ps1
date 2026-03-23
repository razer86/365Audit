<#
.SYNOPSIS
    Performs a security-focused audit of Microsoft Entra ID.

.DESCRIPTION
    Connects to Microsoft Graph and collects identity-related data including:
    - User summary (UPN, name, license, MFA methods, password policy, last sign-in)
    - License inventory with friendly names
    - Admin role assignments and Global Administrator list
    - Guest user list
    - SSPR (Self-Service Password Reset) configuration
    - Group membership and ownership
    - Conditional Access policies
    - Named/Trusted locations
    - Security Defaults status

    Output CSVs:
    - Entra_Users.csv
    - Entra_Users_Unlicensed.csv
    - Entra_Licenses.csv
    - Entra_SSPR.csv
    - Entra_AdminRoles.csv
    - Entra_GlobalAdmins.csv
    - Entra_GuestUsers.csv
    - Entra_Groups.csv
    - Entra_CA_Policies.csv
    - Entra_TrustedLocations.csv
    - Entra_SecurityDefaults.csv
    - Entra_SecureScore.csv
    - Entra_SecureScoreControls.csv
    - Entra_SignIns.csv
    - Entra_AccountCreations.csv
    - Entra_AccountDeletions.csv
    - Entra_AuditEvents.csv
    - Entra_EnterpriseApps.csv
    - Entra_RiskyUsers.csv
    - Entra_RiskySignIns.csv

.NOTES
    Author      : Raymond Slater
    Version     : 1.17.2
    Change Log  : See CHANGELOG.md

.LINK
    https://github.com/razer86/365Audit
#>

#Requires -Version 7.2

param (
    [string]$AuditFolder,
    [switch]$DevMode = $false
)

if (-not $DevMode -and $MyInvocation.InvocationName -eq $MyInvocation.MyCommand.Name) {
    Write-Error "This script must be run from the 365Audit launcher. Use -DevMode for development." -ErrorAction Stop
}

$ScriptVersion = "1.17.2"
Write-Verbose "Invoke-EntraAudit.ps1 loaded (v$ScriptVersion)"

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# === Convert SKU part numbers to friendly names ===
function Get-FriendlySkuName {
    [CmdletBinding()]
    param (
        [string]$Sku
    )

    $skuMap = @{
        # Microsoft 365 Business Plans
        "O365_BUSINESS_PREMIUM"       = "Microsoft 365 Business Premium"
        "O365_BUSINESS_STANDARD"      = "Microsoft 365 Business Standard"
        "O365_BUSINESS_ESSENTIALS"    = "Microsoft 365 Business Basic"
        "M365_APPS_FOR_BUSINESS"      = "Microsoft 365 Apps for Business"
        "M365_APPS_FOR_ENTERPRISE"    = "Microsoft 365 Apps for Enterprise"

        # Microsoft 365 Enterprise Plans
        "ENTERPRISEPREMIUM"           = "Microsoft 365 E5"
        "ENTERPRISEPACK"              = "Microsoft 365 E3"
        "STANDARDPACK"                = "Microsoft 365 E1"
        "M365_F3"                     = "Microsoft 365 F3"
        "DESKLESSPACK"                = "Microsoft 365 F1"

        # Office 365 Plans
        "O365_E1"                     = "Office 365 E1"
        "O365_E3"                     = "Office 365 E3"
        "O365_E5"                     = "Office 365 E5"

        # Exchange Plans
        "EXCHANGESTANDARD"            = "Exchange Online Plan 1"
        "EXCHANGEENTERPRISE"          = "Exchange Online Plan 2"
        "EXCHANGEESSENTIALS"          = "Exchange Online Essentials"

        # Project / Planner
        "PROJECTESSENTIALS"           = "Project Plan 1"
        "PROJECTPREMIUM"              = "Project Plan 3"
        "PROJECTPROFESSIONAL"         = "Project Professional"
        "PROJECT_PLAN1"               = "Project Online Essentials"
        "PROJECT_PLAN2"               = "Project Online Professional"
        "PROJECT_PLAN3"               = "Project Online Premium"
        "PLANNERSTANDALONE"           = "Microsoft Planner"

        # Power Platform
        "POWER_BI_STANDARD"           = "Power BI (Free)"
        "POWER_BI_PRO"                = "Power BI Pro"
        "POWERAPPS_VIRAL"             = "PowerApps (Free)"
        "POWERAPPS_P1"                = "PowerApps Plan 1"
        "POWERAPPS_P2"                = "PowerApps Plan 2"
        "FLOW_FREE"                   = "Power Automate Free"
        "FLOW_P1"                     = "Power Automate Plan 1"
        "FLOW_P2"                     = "Power Automate Plan 2"

        # Misc
        "SMB_APPS"                               = "Business Apps (Free)"
        "SPB"                                    = "Microsoft 365 Business Premium"
        "RMSBASIC"                               = "Azure Rights Management (Free)"
        "Microsoft_Teams_Rooms_Basic"            = "Microsoft Teams Rooms Basic"
        "Microsoft_Teams_Rooms_Pro"              = "Microsoft Teams Rooms Pro"
        "MEETING_ROOM"                           = "Microsoft Teams Rooms Standard"
    }

    if ($skuMap.ContainsKey($Sku)) { return $skuMap[$Sku] }
    return $Sku
}

$script:ConditionalAccessUserCache               = @{}
$script:ConditionalAccessGroupCache              = @{}
$script:ConditionalAccessServicePrincipalCache   = @{}
$script:ConditionalAccessRoleCache               = @{}
$script:ConditionalAccessRoleCacheInitialized    = $false
$script:ConditionalAccessNamedLocationCache      = @{}
$script:ConditionalAccessNamedLocationCacheReady = $false
$script:ConditionalAccessNamedLocations          = @()

function Format-DirectoryUserLabel {
    [CmdletBinding()]
    param(
        [string]$DisplayName,
        [string]$UserPrincipalName,
        [string]$Fallback
    )

    if (-not [string]::IsNullOrWhiteSpace($DisplayName) -and -not [string]::IsNullOrWhiteSpace($UserPrincipalName) -and $DisplayName -ne $UserPrincipalName) {
        return "{0} ({1})" -f $DisplayName, $UserPrincipalName
    }
    if (-not [string]::IsNullOrWhiteSpace($UserPrincipalName)) { return $UserPrincipalName }
    if (-not [string]::IsNullOrWhiteSpace($DisplayName)) { return $DisplayName }
    return $Fallback
}

function Join-ConditionalAccessValues {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [object[]]$Values,
        [string]$Default = '—'
    )

    $items = @(
        $Values |
            Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } |
            Select-Object -Unique
    )

    if ($items.Count -eq 0) {
        return $Default
    }

    return ($items -join '; ')
}

function Convert-ConditionalAccessTokenToLabel {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Kind,
        [Parameter(Mandatory)]
        [string]$Value
    )

    switch ($Kind) {
        'User' {
            switch ($Value) {
                'All'                   { return 'All users' }
                'None'                  { return 'None' }
                'GuestsOrExternalUsers' { return 'All guest and external users' }
            }
        }
        'Group' {
            switch ($Value) {
                'All'  { return 'All groups' }
                'None' { return 'None' }
            }
        }
        'Role' {
            switch ($Value) {
                'All'  { return 'All roles' }
                'None' { return 'None' }
            }
        }
        'Application' {
            switch ($Value) {
                'All'                   { return 'All cloud apps' }
                'Office365'             { return 'Office 365' }
                'MicrosoftAdminPortals' { return 'Microsoft Admin Portals' }
                'None'                  { return 'None' }
            }
        }
        'Location' {
            switch ($Value) {
                'All'        { return 'All locations' }
                'AllTrusted' { return 'All trusted locations' }
                'None'       { return 'None' }
            }
        }
    }

    return $null
}

function Convert-ConditionalAccessBuiltInControlToLabel {
    [CmdletBinding()]
    param(
        [string]$Control
    )

    switch ($Control) {
        'mfa'                 { return 'Require MFA' }
        'block'               { return 'Block access' }
        'compliantDevice'     { return 'Require compliant device' }
        'domainJoinedDevice'  { return 'Require Microsoft Entra hybrid joined device' }
        'approvedApplication' { return 'Require approved client app' }
        'compliantApplication'{ return 'Require app protection policy' }
        'passwordChange'      { return 'Require password change' }
        default               { return $Control }
    }
}

function Convert-ConditionalAccessClientAppTypeToLabel {
    [CmdletBinding()]
    param(
        [string]$Type
    )

    switch ($Type) {
        'all'                           { return 'All client apps' }
        'browser'                       { return 'Browser' }
        'mobileAppsAndDesktopClients'   { return 'Mobile apps and desktop clients' }
        'exchangeActiveSync'            { return 'Exchange ActiveSync' }
        'easSupported'                  { return 'Exchange ActiveSync clients' }
        'other'                         { return 'Other clients' }
        default                         { return $Type }
    }
}

function Convert-ConditionalAccessPlatformToLabel {
    [CmdletBinding()]
    param(
        [string]$Platform
    )

    switch ($Platform) {
        'all'     { return 'All platforms' }
        'android' { return 'Android' }
        'iOS'     { return 'iOS' }
        'windows' { return 'Windows' }
        'macOS'   { return 'macOS' }
        'linux'   { return 'Linux' }
        default   { return $Platform }
    }
}

function Convert-ConditionalAccessRiskLevelToLabel {
    [CmdletBinding()]
    param(
        [string]$RiskLevel
    )

    switch ($RiskLevel) {
        'low'    { return 'Low' }
        'medium' { return 'Medium' }
        'high'   { return 'High' }
        'hidden' { return 'Hidden' }
        'none'   { return 'None' }
        default  { return $RiskLevel }
    }
}

function Convert-ConditionalAccessUserActionToLabel {
    [CmdletBinding()]
    param(
        [string]$Action
    )

    switch ($Action) {
        'urn:user:registersecurityinfo' { return 'Register security information' }
        default                         { return $Action }
    }
}

function Convert-ConditionalAccessGrantOperatorToLabel {
    [CmdletBinding()]
    param(
        [string]$Operator
    )

    switch ($Operator) {
        'AND' { return 'All selected controls required' }
        'OR'  { return 'One of the selected controls required' }
        default { return $Operator }
    }
}

function Initialize-ConditionalAccessRoleCache {
    [CmdletBinding()]
    param()

    if ($script:ConditionalAccessRoleCacheInitialized) {
        return
    }

    try {
        if (Get-Command Get-MgDirectoryRoleTemplate -ErrorAction SilentlyContinue) {
            foreach ($template in (Get-MgDirectoryRoleTemplate -All -ErrorAction Stop)) {
                if ($template.Id -and $template.DisplayName) {
                    $script:ConditionalAccessRoleCache[$template.Id] = $template.DisplayName
                }
            }
        }
    }
    catch {
        Write-Verbose "Unable to resolve Conditional Access role templates: $_"
    }

    try {
        foreach ($role in (Get-MgDirectoryRole -All -ErrorAction Stop)) {
            if ($role.Id -and $role.DisplayName) {
                $script:ConditionalAccessRoleCache[$role.Id] = $role.DisplayName
            }
            if ($role.RoleTemplateId -and $role.DisplayName) {
                $script:ConditionalAccessRoleCache[$role.RoleTemplateId] = $role.DisplayName
            }
        }
    }
    catch {
        Write-Verbose "Unable to resolve active directory roles for Conditional Access: $_"
    }

    $script:ConditionalAccessRoleCacheInitialized = $true
}

function Initialize-ConditionalAccessNamedLocationCache {
    [CmdletBinding()]
    param()

    if ($script:ConditionalAccessNamedLocationCacheReady) {
        return
    }

    try {
        if (-not $script:ConditionalAccessNamedLocations -or $script:ConditionalAccessNamedLocations.Count -eq 0) {
            $script:ConditionalAccessNamedLocations = @(Get-MgIdentityConditionalAccessNamedLocation -All -ErrorAction Stop)
        }

        foreach ($location in $script:ConditionalAccessNamedLocations) {
            if ($location.Id -and $location.DisplayName) {
                $script:ConditionalAccessNamedLocationCache[$location.Id] = $location.DisplayName
            }
        }
    }
    catch {
        Write-Verbose "Unable to resolve Conditional Access named locations: $_"
    }

    $script:ConditionalAccessNamedLocationCacheReady = $true
}

function Resolve-ConditionalAccessObjectId {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateSet('User', 'Group', 'Role', 'Application', 'Location')]
        [string]$Kind,
        [Parameter(Mandatory)]
        [string]$Id
    )

    if ([string]::IsNullOrWhiteSpace($Id)) {
        return $null
    }

    $wellKnown = Convert-ConditionalAccessTokenToLabel -Kind $Kind -Value $Id
    if ($wellKnown) {
        return $wellKnown
    }

    switch ($Kind) {
        'User' {
            if ($script:ConditionalAccessUserCache.ContainsKey($Id)) {
                return $script:ConditionalAccessUserCache[$Id]
            }

            try {
                $user = Get-MgUser -UserId $Id -Property Id,DisplayName,UserPrincipalName -ErrorAction Stop
                $resolved = Format-DirectoryUserLabel -DisplayName $user.DisplayName -UserPrincipalName $user.UserPrincipalName -Fallback $Id
            }
            catch {
                $resolved = $Id
            }

            $script:ConditionalAccessUserCache[$Id] = $resolved
            return $resolved
        }
        'Group' {
            if ($script:ConditionalAccessGroupCache.ContainsKey($Id)) {
                return $script:ConditionalAccessGroupCache[$Id]
            }

            try {
                $group = Get-MgGroup -GroupId $Id -Property Id,DisplayName -ErrorAction Stop
                $resolved = if ($group.DisplayName) { $group.DisplayName } else { $Id }
            }
            catch {
                $resolved = $Id
            }

            $script:ConditionalAccessGroupCache[$Id] = $resolved
            return $resolved
        }
        'Role' {
            Initialize-ConditionalAccessRoleCache
            if ($script:ConditionalAccessRoleCache.ContainsKey($Id)) {
                return $script:ConditionalAccessRoleCache[$Id]
            }

            return $Id
        }
        'Application' {
            if ($script:ConditionalAccessServicePrincipalCache.ContainsKey($Id)) {
                return $script:ConditionalAccessServicePrincipalCache[$Id]
            }

            try {
                $servicePrincipal = Get-MgServicePrincipal -ServicePrincipalId $Id -Property Id,DisplayName,AppId -ErrorAction Stop
                if ($servicePrincipal.DisplayName -and $servicePrincipal.AppId) {
                    $resolved = "{0} ({1})" -f $servicePrincipal.DisplayName, $servicePrincipal.AppId
                }
                elseif ($servicePrincipal.DisplayName) {
                    $resolved = $servicePrincipal.DisplayName
                }
                else {
                    $resolved = $Id
                }
            }
            catch {
                $resolved = $Id
            }

            $script:ConditionalAccessServicePrincipalCache[$Id] = $resolved
            return $resolved
        }
        'Location' {
            Initialize-ConditionalAccessNamedLocationCache
            if ($script:ConditionalAccessNamedLocationCache.ContainsKey($Id)) {
                return $script:ConditionalAccessNamedLocationCache[$Id]
            }

            return $Id
        }
    }
}

function Resolve-ConditionalAccessValueList {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [object[]]$Values,
        [Parameter(Mandatory)]
        [ValidateSet('User', 'Group', 'Role', 'Application', 'Location')]
        [string]$Kind
    )

    $resolvedValues = foreach ($value in @($Values)) {
        if ($null -eq $value) { continue }

        $textValue = [string]$value
        if ([string]::IsNullOrWhiteSpace($textValue)) { continue }

        Resolve-ConditionalAccessObjectId -Kind $Kind -Id $textValue
    }

    return @($resolvedValues | Where-Object { $_ } | Select-Object -Unique)
}

# === Ensure helper functions are loaded ===
if (-not (Get-Command Connect-MgGraphSecure -ErrorAction SilentlyContinue)) {
    Write-Error "Connect-MgGraphSecure is not loaded. Please run from the 365Audit launcher."
    exit 1
}
if (-not (Get-Command Initialize-AuditOutput -ErrorAction SilentlyContinue)) {
    Write-Error "Initialize-AuditOutput is not loaded. Please run from the 365Audit launcher."
    exit 1
}

# === Initialise output folder ===
try {
    $context   = Initialize-AuditOutput
    $outputDir = $context.RawOutputPath
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
}
catch {
    Write-Error "Failed to initialise audit output directory: $_"
    exit 1
}

# === Connect to Microsoft Graph and load Entra-specific sub-modules ===
try {
    Connect-MgGraphSecure
    Import-GraphSubModules @(
        'Microsoft.Graph.Users',
        'Microsoft.Graph.Groups',
        'Microsoft.Graph.Reports',
        'Microsoft.Graph.Identity.SignIns',
        'Microsoft.Graph.Applications'          # Get-MgServicePrincipal / Get-MgServicePrincipalAppRoleAssignment
    )
}
catch {
    Write-Error "Microsoft Graph connection failed: $_"
    exit 1
}

Write-Host "`nStarting Entra Audit for $($context.OrgName)..." -ForegroundColor Cyan

$step       = 0
$totalSteps = 19
$activity   = "Entra Audit — $($context.OrgName)"


# ================================
# ===   Sign-in Logs            ===
# ================================
$subscribedSkus    = Get-MgSubscribedSku -All
$premiumSignInSkus = @("AAD_PREMIUM", "AAD_PREMIUM_P2", "ENTERPRISEPREMIUM", "ENTERPRISEPACK",
                       "EMS", "EMS_PREMIUM", "SPB", "O365_BUSINESS_PREMIUM", "M365_F3", "IDENTITY_GOVERNANCE")
$retentionDays   = if (($subscribedSkus.SkuPartNumber | Where-Object { $_ -in $premiumSignInSkus }).Count -gt 0) { 30 } else { 7 }
$signInRetention = if ($retentionDays -eq 30) { "30 days (AAD Premium)" } else { "7 days (AAD Free)" }

$step++
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Retrieving sign-in logs ($signInRetention)..." -PercentComplete ([int]($step / $totalSteps * 100))

$signIns   = @{}
$signInMap = @{}

try {
    $rawSignIns = Get-MgAuditLogSignIn -All -ErrorAction Stop
    foreach ($entry in $rawSignIns) {
        $upn = $entry.UserPrincipalName
        if (-not $upn) { continue }

        if (-not $signIns.ContainsKey($upn)) {
            $signIns[$upn] = $entry.CreatedDateTime
        }

        if (-not $signInMap.ContainsKey($upn)) {
            $signInMap[$upn] = [System.Collections.Generic.List[object]]::new()
        }

        if ($signInMap[$upn].Count -lt 10) {
            $loc = $entry.Location
            $signInMap[$upn].Add([PSCustomObject]@{
                UPN           = $upn
                Timestamp     = $entry.CreatedDateTime.ToUniversalTime().ToString("yyyy-MM-dd HH:mm") + " UTC"
                App           = $entry.AppDisplayName
                IPAddress     = $entry.IpAddress
                City          = $loc.City
                Country       = $loc.CountryOrRegion
                Success       = ($entry.Status.ErrorCode -eq 0)
                FailureReason = if ($entry.Status.ErrorCode -ne 0) { $entry.Status.FailureReason } else { "" }
            })
        }
    }

    $signInExport = foreach ($entries in $signInMap.Values) { $entries }
    $signInExport | Export-Csv "$outputDir\Entra_SignIns.csv" -NoTypeInformation -Encoding UTF8
    Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Retrieving sign-in logs..." -CurrentOperation "Saved: Entra_SignIns.csv ($($signInMap.Count) users)" -PercentComplete ([int]($step / $totalSteps * 100))
}
catch {
    Add-AuditIssue -Severity 'Warning' -Section 'Entra' -Collector 'Get-MgAuditLogSignIn' -Description ($_.Exception.Message ?? "$_") -Action 'Check AuditLog.Read.All permissions or re-run Setup-365AuditApp.ps1'
    Write-Warning "Failed to retrieve sign-in logs: $_"
}


# ================================
# ===   Directory Audit Events  ===
# ================================
$step++
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Retrieving directory audit events (last $retentionDays days)..." -PercentComplete ([int]($step / $totalSteps * 100))

$auditFrom  = (Get-Date).AddDays(-$retentionDays).ToString("yyyy-MM-ddTHH:mm:ssZ")
$dateFilter = "activityDateTime ge $auditFrom"

function Get-AuditInitiator {
    [CmdletBinding()]
    param ($Entry)
    if ($Entry.InitiatedBy.User.UserPrincipalName) { return $Entry.InitiatedBy.User.UserPrincipalName }
    if ($Entry.InitiatedBy.App.DisplayName)        { return "$($Entry.InitiatedBy.App.DisplayName) [app]" }
    return "System"
}

# --- Account Creations ---
try {
    $rawCreations  = Get-MgAuditLogDirectoryAudit -Filter "$dateFilter and activityDisplayName eq 'Add user'" -All -ErrorAction Stop
    $acctCreations = foreach ($entry in $rawCreations) {
        $target = $entry.TargetResources | Select-Object -First 1
        [PSCustomObject]@{
            Timestamp   = $entry.ActivityDateTime.ToUniversalTime().ToString("yyyy-MM-dd HH:mm") + " UTC"
            InitiatedBy = Get-AuditInitiator $entry
            TargetUPN   = $target.UserPrincipalName
            TargetName  = $target.DisplayName
            Result      = $entry.Result
        }
    }
    $acctCreations | Export-Csv "$outputDir\Entra_AccountCreations.csv" -NoTypeInformation -Encoding UTF8
    Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Retrieving directory audit events..." -CurrentOperation "Saved: Entra_AccountCreations.csv ($($acctCreations.Count) events)" -PercentComplete ([int]($step / $totalSteps * 100))
}
catch {
    Add-AuditIssue -Severity 'Warning' -Section 'Entra' -Collector 'Get-MgAuditLogDirectoryAudit' -Description ($_.Exception.Message ?? "$_") -Action 'Check AuditLog.Read.All permissions or re-run Setup-365AuditApp.ps1'
    Write-Warning "Failed to retrieve account creation events: $_"
}

# --- Account Deletions ---
try {
    $rawDeletions  = Get-MgAuditLogDirectoryAudit -Filter "$dateFilter and activityDisplayName eq 'Delete user'" -All -ErrorAction Stop
    $acctDeletions = foreach ($entry in $rawDeletions) {
        $target = $entry.TargetResources | Select-Object -First 1
        [PSCustomObject]@{
            Timestamp   = $entry.ActivityDateTime.ToUniversalTime().ToString("yyyy-MM-dd HH:mm") + " UTC"
            InitiatedBy = Get-AuditInitiator $entry
            TargetUPN   = $target.UserPrincipalName
            TargetName  = $target.DisplayName
            Result      = $entry.Result
        }
    }
    $acctDeletions | Export-Csv "$outputDir\Entra_AccountDeletions.csv" -NoTypeInformation -Encoding UTF8
    Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Retrieving directory audit events..." -CurrentOperation "Saved: Entra_AccountDeletions.csv ($($acctDeletions.Count) events)" -PercentComplete ([int]($step / $totalSteps * 100))
}
catch {
    Add-AuditIssue -Severity 'Warning' -Section 'Entra' -Collector 'Get-MgAuditLogDirectoryAudit' -Description ($_.Exception.Message ?? "$_") -Action 'Check AuditLog.Read.All permissions or re-run Setup-365AuditApp.ps1'
    Write-Warning "Failed to retrieve account deletion events: $_"
}

# --- Notable Audit Events ---
$securityActivityNames = @(
    "Reset user password",
    "User registered security info",
    "User deleted security info",
    "User changed default security info",
    "Admin registered security info for a user",
    "Admin deleted security info for a user",
    "Admin updated security info for a user"
)

try {
    $roleEvents     = @(Get-MgAuditLogDirectoryAudit -Filter "$dateFilter and category eq 'RoleManagement'" -All -ErrorAction Stop)
    $rawUserMgmt    = @(Get-MgAuditLogDirectoryAudit -Filter "$dateFilter and category eq 'UserManagement'" -All -ErrorAction Stop)
    $securityEvents = @($rawUserMgmt | Where-Object { $_.ActivityDisplayName -in $securityActivityNames })

    $auditEvents = foreach ($entry in ($roleEvents + $securityEvents)) {
        $targetUser = $entry.TargetResources | Where-Object { $_.Type -eq "User" } | Select-Object -First 1
        $targetRole = $entry.TargetResources | Where-Object { $_.Type -eq "Role" } | Select-Object -First 1
        if (-not $targetUser) { $targetUser = $entry.TargetResources | Select-Object -First 1 }

        [PSCustomObject]@{
            Timestamp   = $entry.ActivityDateTime.ToUniversalTime().ToString("yyyy-MM-dd HH:mm") + " UTC"
            Category    = $entry.Category
            Activity    = $entry.ActivityDisplayName
            InitiatedBy = Get-AuditInitiator $entry
            TargetUPN   = $targetUser.UserPrincipalName
            TargetName  = $targetUser.DisplayName
            TargetRole  = if ($targetRole) { $targetRole.DisplayName } else { "" }
            Result      = $entry.Result
        }
    }

    $auditEvents | Sort-Object Timestamp -Descending |
        Export-Csv "$outputDir\Entra_AuditEvents.csv" -NoTypeInformation -Encoding UTF8
    Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Retrieving directory audit events..." -CurrentOperation "Saved: Entra_AuditEvents.csv ($($auditEvents.Count) events)" -PercentComplete ([int]($step / $totalSteps * 100))
}
catch {
    Add-AuditIssue -Severity 'Warning' -Section 'Entra' -Collector 'Get-MgAuditLogDirectoryAudit' -Description ($_.Exception.Message ?? "$_") -Action 'Check AuditLog.Read.All permissions or re-run Setup-365AuditApp.ps1'
    Write-Warning "Failed to retrieve directory audit events: $_"
}


# ================================
# ===   Licence Summary         ===
# ================================
$step++
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Collecting licence summary..." -PercentComplete ([int]($step / $totalSteps * 100))

$licenseDetails = foreach ($sku in $subscribedSkus) {
    [PSCustomObject]@{
        SkuPartNumber    = $sku.SkuPartNumber
        SkuFriendlyName  = Get-FriendlySkuName $sku.SkuPartNumber
        SkuId            = $sku.SkuId
        EnabledUnits     = $sku.PrepaidUnits.Enabled
        SuspendedUnits   = $sku.PrepaidUnits.Suspended
        WarningUnits     = $sku.PrepaidUnits.Warning
        ConsumedUnits    = $sku.ConsumedUnits
        CapabilityStatus = $sku.CapabilityStatus
        SubscriptionIds  = ($sku.SubscriptionIds -join ", ")
        PurchaseChannel  = if ($sku.AppliesTo -eq "User") { "Direct" } else { "Partner" }
    }
}

$licenseDetails | Export-Csv -Path "$outputDir\Entra_Licenses.csv" -NoTypeInformation -Encoding UTF8
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Collecting licence summary..." -CurrentOperation "Saved: Entra_Licenses.csv" -PercentComplete ([int]($step / $totalSteps * 100))


# ================================
# ===   SSPR Configuration      ===
# ================================
$step++
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Collecting SSPR configuration..." -PercentComplete ([int]($step / $totalSteps * 100))

try {
    $authPolicy = Get-MgPolicyAuthenticationMethodPolicy
    $ssprState  = $authPolicy.RegistrationEnforcement.AuthenticationMethodsRegistrationCampaign.State

    $friendlySspr = switch ($ssprState) {
        "enabled"  { "Enabled" }
        "disabled" { "Disabled" }
        "default"  { "Not Enforced (Default)" }
        default    { "Unknown" }
    }

    [PSCustomObject]@{ SSPREnabled = $friendlySspr } |
        Export-Csv "$outputDir\Entra_SSPR.csv" -NoTypeInformation -Encoding UTF8
    Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Collecting SSPR configuration..." -CurrentOperation "Saved: Entra_SSPR.csv" -PercentComplete ([int]($step / $totalSteps * 100))
}
catch {
    Add-AuditIssue -Severity 'Warning' -Section 'Entra' -Collector 'Get-MgPolicyAuthenticationMethodPolicy' -Description ($_.Exception.Message ?? "$_") -Action 'Check Policy.Read.All permissions or re-run Setup-365AuditApp.ps1'
    Write-Warning "Unable to retrieve SSPR configuration: $_"
}


# ================================
# ===   User Summary            ===
# ================================
$step++
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Collecting user summary..." -PercentComplete ([int]($step / $totalSteps * 100))

$users = Get-MgUser -All -Filter "userType eq 'Member'" -Property DisplayName, GivenName, Surname, UserPrincipalName, Id, AccountEnabled, PasswordPolicies, LastPasswordChangeDateTime

$skuLookup = @{}
$subscribedSkus | ForEach-Object { $skuLookup[$_.SkuId] = $_.SkuPartNumber }

$mfaMap = @{}
foreach ($user in $users) {
    try {
        $methods       = Get-MgUserAuthenticationMethod -UserId $user.Id
        $types         = $methods | ForEach-Object { $_.AdditionalProperties['@odata.type'] }
        $filteredTypes = $types | Where-Object { $_ -ne "#microsoft.graph.passwordAuthenticationMethod" }

        $friendlyTypes = $filteredTypes | ForEach-Object {
            switch ($_ -replace "#microsoft.graph.", "") {
                "phoneAuthenticationMethod"                   { "Phone (SMS/Call)" }
                "microsoftAuthenticatorAuthenticationMethod"  { "Authenticator App" }
                "fido2AuthenticationMethod"                   { "FIDO2 Key" }
                "windowsHelloForBusinessAuthenticationMethod" { "Windows Hello" }
                "emailAuthenticationMethod"                   { "Email" }
                "softwareOathAuthenticationMethod"            { "Software OTP" }
                default { $_ }
            }
        }

        $uniqueTypes = $friendlyTypes | Sort-Object -Unique
        $mfaMap[$user.UserPrincipalName] = @{ Types = $uniqueTypes; Count = $uniqueTypes.Count }
    }
    catch {
        Write-Warning "Unable to get MFA methods for $($user.UserPrincipalName): $_"
        $mfaMap[$user.UserPrincipalName] = @{ Types = @(); Count = 0 }
    }
}

$licenseMap = @{}
foreach ($user in $users) {
    $licenses = @()
    try {
        $details = Get-MgUserLicenseDetail -UserId $user.Id
        foreach ($detail in $details) {
            $licenses += Get-FriendlySkuName $detail.SkuPartNumber
        }
    }
    catch {
        Write-Verbose "Failed to retrieve licence details for $($user.UserPrincipalName): $_"
    }
    if ($user.UserPrincipalName) {
        $licenseMap[$user.UserPrincipalName] = $licenses -join ", "
    }
}

$userReport = foreach ($user in $users) {
    $upn        = $user.UserPrincipalName
    $mfaEnabled = $mfaMap.ContainsKey($upn) -and ($mfaMap[$upn]['Count'] -gt 0)
    $mfaTypes   = if ($mfaMap[$upn]['Count'] -gt 0) { $mfaMap[$upn].Types -join ", " } else { "None" }
    $lastSignIn = if ($signIns.ContainsKey($upn)) { $signIns[$upn].ToUniversalTime().ToString("yyyy-MM-dd HH:mm") + " UTC" } else { "Unavailable" }

    [PSCustomObject]@{
        UserId                    = $user.Id
        UPN                       = $upn
        FirstName                 = $user.GivenName
        LastName                  = $user.Surname
        AccountStatus             = if ($user.AccountEnabled) { "Enabled" } else { "Blocked" }
        AssignedLicense           = if ($licenseMap.ContainsKey($upn) -and $licenseMap[$upn]) { $licenseMap[$upn] } else { "None" }
        MFAEnabled                = $mfaEnabled
        MFAMethods                = $mfaTypes
        MFACount                  = $mfaMap[$upn]['Count']
        DisablePasswordExpiration = if ($user.PasswordPolicies -notmatch "DisablePasswordExpiration") { "Enabled" } else { "Disabled" }
        LastPasswordChange        = $user.LastPasswordChangeDateTime
        LastSignIn                = $lastSignIn
    }
}

foreach ($user in $users) {
    if ($user.Id) {
        $script:ConditionalAccessUserCache[$user.Id] = Format-DirectoryUserLabel -DisplayName $user.DisplayName -UserPrincipalName $user.UserPrincipalName -Fallback $user.Id
    }
}

$licensedUsers   = @($userReport | Where-Object { $_.AssignedLicense -ne "None" })
$unlicensedUsers = @($userReport | Where-Object { $_.AssignedLicense -eq "None" })

$licensedUsers   | Export-Csv "$outputDir\Entra_Users.csv"            -NoTypeInformation -Encoding UTF8
$unlicensedUsers | Export-Csv "$outputDir\Entra_Users_Unlicensed.csv" -NoTypeInformation -Encoding UTF8
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Collecting user summary..." -CurrentOperation "Saved: Entra_Users.csv ($($licensedUsers.Count) licensed), Entra_Users_Unlicensed.csv ($($unlicensedUsers.Count) unlicensed)" -PercentComplete ([int]($step / $totalSteps * 100))


# ================================
# ===   Admin Role Assignments  ===
# ================================
$step++
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Collecting admin role assignments..." -PercentComplete ([int]($step / $totalSteps * 100))

$roles        = Get-MgDirectoryRole
$adminReport  = @()
$globalAdmins = @()

foreach ($role in $roles) {
    try {
        $members = Get-MgDirectoryRoleMember -DirectoryRoleId $role.Id

        foreach ($member in $members) {
            $entry = [PSCustomObject]@{
                RoleName                = $role.DisplayName
                MemberDisplayName       = $member.AdditionalProperties.displayName
                MemberUserPrincipalName = $member.AdditionalProperties.userPrincipalName
            }
            $adminReport += $entry
            if ($role.DisplayName -eq "Global Administrator") {
                $globalAdmins += $entry
            }
        }
    }
    catch {
        Write-Warning "Could not retrieve members for role: $($role.DisplayName)"
    }
}

$adminReport  | Export-Csv "$outputDir\Entra_AdminRoles.csv"   -NoTypeInformation -Encoding UTF8
$globalAdmins | Export-Csv "$outputDir\Entra_GlobalAdmins.csv" -NoTypeInformation -Encoding UTF8
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Collecting admin role assignments..." -CurrentOperation "Saved: Entra_AdminRoles.csv, Entra_GlobalAdmins.csv" -PercentComplete ([int]($step / $totalSteps * 100))

if ($globalAdmins.Count -eq 1) {
    Write-Warning "Only ONE Global Administrator found. Best practice is at least two to avoid lockout."
}


# ================================
# ===   Guest Users             ===
# ================================
$step++
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Collecting guest user summary..." -PercentComplete ([int]($step / $totalSteps * 100))

$guestData = foreach ($guest in (Get-MgUser -Filter "UserType eq 'Guest'" -All -Property Id,DisplayName,UserPrincipalName,CreatedDateTime,SignInActivity -ErrorAction SilentlyContinue)) {
    if ($guest.Id) {
        $script:ConditionalAccessUserCache[$guest.Id] = Format-DirectoryUserLabel -DisplayName $guest.DisplayName -UserPrincipalName $guest.UserPrincipalName -Fallback $guest.Id
    }

    [PSCustomObject]@{
        UserId            = $guest.Id
        DisplayName       = $guest.DisplayName
        UserPrincipalName = $guest.UserPrincipalName
        CreatedDateTime   = $guest.CreatedDateTime
        LastSignIn        = if ($guest.SignInActivity.LastSignInDateTime) { $guest.SignInActivity.LastSignInDateTime.ToString("yyyy-MM-dd HH:mm") + " UTC" } else { $null }
    }
}
$guestData | Export-Csv "$outputDir\Entra_GuestUsers.csv" -NoTypeInformation -Encoding UTF8
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Collecting guest user summary..." -CurrentOperation "Saved: Entra_GuestUsers.csv" -PercentComplete ([int]($step / $totalSteps * 100))


# ================================
# ===   Groups                  ===
# ================================
$step++
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Collecting group summary..." -PercentComplete ([int]($step / $totalSteps * 100))

$groupData = foreach ($group in (Get-MgGroup -All)) {
    $owners  = (Get-MgGroupOwner -GroupId $group.Id -ErrorAction SilentlyContinue |
        ForEach-Object { $_.AdditionalProperties['userPrincipalName'] }) -join "; "
    $members = (Get-MgGroupMember -GroupId $group.Id -ErrorAction SilentlyContinue |
        ForEach-Object { $_.AdditionalProperties['userPrincipalName'] }) -join "; "

    [PSCustomObject]@{
        DisplayName        = $group.DisplayName
        GroupId            = $group.Id
        GroupType          = if ($group.GroupTypes -contains "Unified") { "Microsoft 365" } else { "Security" }
        MembershipType     = if ($group.MembershipRule) { "Dynamic" } else { "Assigned" }
        Email              = $group.Mail
        Source             = if ($group.OnPremisesSyncEnabled -eq $true) { "On-Premises" } elseif ($group.OnPremisesSyncEnabled -eq $false) { "Cloud (Sync Stopped)" } else { "Cloud" }
        Owners             = $owners
        Members            = $members
        MailEnabled        = $group.MailEnabled
        SecurityEnabled    = $group.SecurityEnabled
        IsAssignableToRole = $group.IsAssignableToRole
        Visibility         = $group.Visibility
        OnPremSyncEnabled  = $group.OnPremisesSyncEnabled
    }
}

foreach ($group in $groupData) {
    if ($group.GroupId -and $group.DisplayName) {
        $script:ConditionalAccessGroupCache[$group.GroupId] = $group.DisplayName
    }
}

$groupData | Export-Csv -Path "$outputDir\Entra_Groups.csv" -NoTypeInformation -Encoding UTF8
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Collecting group summary..." -CurrentOperation "Saved: Entra_Groups.csv" -PercentComplete ([int]($step / $totalSteps * 100))


# ================================
# ===   Conditional Access      ===
# ================================
$step++
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Collecting Conditional Access policies..." -PercentComplete ([int]($step / $totalSteps * 100))

try {
    Initialize-ConditionalAccessNamedLocationCache

    $caPolicyData = foreach ($policy in (Get-MgIdentityConditionalAccessPolicy -All)) {
        $userConditions         = $policy.Conditions.Users
        $applicationConditions  = $policy.Conditions.Applications
        $platformConditions     = $policy.Conditions.Platforms
        $locationConditions     = $policy.Conditions.Locations
        $deviceConditions       = $policy.Conditions.Devices
        $deviceFilter           = if ($deviceConditions) { $deviceConditions.DeviceFilter } else { $null }
        $grantControls          = $policy.GrantControls

        $includeUsers           = Join-ConditionalAccessValues (Resolve-ConditionalAccessValueList -Values $userConditions.IncludeUsers -Kind 'User')
        $excludeUsers           = Join-ConditionalAccessValues (Resolve-ConditionalAccessValueList -Values $userConditions.ExcludeUsers -Kind 'User')
        $includeGroups          = Join-ConditionalAccessValues (Resolve-ConditionalAccessValueList -Values $userConditions.IncludeGroups -Kind 'Group')
        $excludeGroups          = Join-ConditionalAccessValues (Resolve-ConditionalAccessValueList -Values $userConditions.ExcludeGroups -Kind 'Group')
        $includeRoles           = Join-ConditionalAccessValues (Resolve-ConditionalAccessValueList -Values $userConditions.IncludeRoles -Kind 'Role')
        $excludeRoles           = Join-ConditionalAccessValues (Resolve-ConditionalAccessValueList -Values $userConditions.ExcludeRoles -Kind 'Role')
        $includeApplications    = Join-ConditionalAccessValues (Resolve-ConditionalAccessValueList -Values $applicationConditions.IncludeApplications -Kind 'Application')
        $excludeApplications    = Join-ConditionalAccessValues (Resolve-ConditionalAccessValueList -Values $applicationConditions.ExcludeApplications -Kind 'Application')
        $includeUserActions     = Join-ConditionalAccessValues (@($applicationConditions.IncludeUserActions | ForEach-Object { Convert-ConditionalAccessUserActionToLabel $_ }))
        $grantControlsLabel     = Join-ConditionalAccessValues (@($grantControls.BuiltInControls | ForEach-Object { Convert-ConditionalAccessBuiltInControlToLabel $_ }))
        $grantOperatorLabel     = if ($grantControls.Operator) { Convert-ConditionalAccessGrantOperatorToLabel -Operator $grantControls.Operator } else { '—' }
        $clientAppTypes         = Join-ConditionalAccessValues (@($policy.Conditions.ClientAppTypes | ForEach-Object { Convert-ConditionalAccessClientAppTypeToLabel $_ })) -Default 'All client apps'
        $includePlatforms       = Join-ConditionalAccessValues (@($platformConditions.IncludePlatforms | ForEach-Object { Convert-ConditionalAccessPlatformToLabel $_ }))
        $excludePlatforms       = Join-ConditionalAccessValues (@($platformConditions.ExcludePlatforms | ForEach-Object { Convert-ConditionalAccessPlatformToLabel $_ }))
        $includeLocations       = Join-ConditionalAccessValues (Resolve-ConditionalAccessValueList -Values $locationConditions.IncludeLocations -Kind 'Location')
        $excludeLocations       = Join-ConditionalAccessValues (Resolve-ConditionalAccessValueList -Values $locationConditions.ExcludeLocations -Kind 'Location')
        $signInRiskLevels       = Join-ConditionalAccessValues (@($policy.Conditions.SignInRiskLevels | ForEach-Object { Convert-ConditionalAccessRiskLevelToLabel $_ }))
        $userRiskLevels         = Join-ConditionalAccessValues (@($policy.Conditions.UserRiskLevels | ForEach-Object { Convert-ConditionalAccessRiskLevelToLabel $_ }))
        $deviceFilterSummary    = if ($deviceFilter -and $deviceFilter.Mode -and $deviceFilter.Rule) {
            "{0}: {1}" -f $deviceFilter.Mode, $deviceFilter.Rule
        }
        elseif ($deviceFilter -and $deviceFilter.Mode) {
            $deviceFilter.Mode
        }
        else {
            '—'
        }

        [PSCustomObject]@{
            Name                = $policy.DisplayName
            State               = $policy.State
            IncludeUsers        = $includeUsers
            ExcludeUsers        = $excludeUsers
            IncludeGroups       = $includeGroups
            ExcludeGroups       = $excludeGroups
            IncludeRoles        = $includeRoles
            ExcludeRoles        = $excludeRoles
            IncludeApplications = $includeApplications
            ExcludeApplications = $excludeApplications
            UserActions         = $includeUserActions
            GrantControls       = $grantControlsLabel
            GrantOperator       = $grantOperatorLabel
            RequiresMFA         = ($grantControls.BuiltInControls -contains "mfa")
            ClientAppTypes      = $clientAppTypes
            IncludePlatforms    = $includePlatforms
            ExcludePlatforms    = $excludePlatforms
            IncludeLocations    = $includeLocations
            ExcludeLocations    = $excludeLocations
            SignInRiskLevels    = $signInRiskLevels
            UserRiskLevels      = $userRiskLevels
            DeviceFilter        = $deviceFilterSummary
        }
    }

    $caPolicyData | Export-Csv "$outputDir\Entra_CA_Policies.csv" -NoTypeInformation -Encoding UTF8
    Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Collecting Conditional Access policies..." -CurrentOperation "Saved: Entra_CA_Policies.csv" -PercentComplete ([int]($step / $totalSteps * 100))
}
catch {
    Add-AuditIssue -Severity 'Warning' -Section 'Entra' -Collector 'Get-MgIdentityConditionalAccessPolicy' -Description ($_.Exception.Message ?? "$_") -Action 'Check Policy.Read.All permissions or re-run Setup-365AuditApp.ps1'
    Write-Warning "Unable to retrieve Conditional Access policies: $_"
}


# ================================
# ===   Named / Trusted Locations
# ================================
$step++
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Collecting named locations..." -PercentComplete ([int]($step / $totalSteps * 100))

try {
    Initialize-ConditionalAccessNamedLocationCache
    $namedLocations = if ($script:ConditionalAccessNamedLocations -and $script:ConditionalAccessNamedLocations.Count -gt 0) {
        $script:ConditionalAccessNamedLocations
    }
    else {
        @(Get-MgIdentityConditionalAccessNamedLocation -All)
    }

    $locationData = foreach ($loc in $namedLocations) {
        $ipRanges = if ($loc.AdditionalProperties.ContainsKey('ipRanges')) {
            ($loc.AdditionalProperties['ipRanges'] | ForEach-Object { $_['cidrAddress'] }) -join ", "
        }
        else { "-" }

        [PSCustomObject]@{
            Name      = $loc.DisplayName
            IsTrusted = $loc.AdditionalProperties['isTrusted']
            Type      = $loc.ODataType
            IPRanges  = $ipRanges
        }
    }

    $locationData | Export-Csv "$outputDir\Entra_TrustedLocations.csv" -NoTypeInformation -Encoding UTF8
    Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Collecting named locations..." -CurrentOperation "Saved: Entra_TrustedLocations.csv" -PercentComplete ([int]($step / $totalSteps * 100))
}
catch {
    Add-AuditIssue -Severity 'Warning' -Section 'Entra' -Collector 'Get-MgIdentityConditionalAccessNamedLocation' -Description ($_.Exception.Message ?? "$_") -Action 'Check Policy.Read.All permissions or re-run Setup-365AuditApp.ps1'
    Write-Warning "Unable to retrieve named locations: $_"
}


# ================================
# ===   Identity Secure Score   ===
# ================================
$step++
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Collecting Identity Secure Score..." -PercentComplete ([int]($step / $totalSteps * 100))

try {
    $scoreResponse = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/security/secureScores?$top=1' -OutputType PSObject -ErrorAction Stop
    $latestScore   = $scoreResponse.value | Select-Object -First 1

    if ($latestScore) {
        $currentScore = [math]::Round($latestScore.currentScore, 2)
        $maxScore     = [math]::Round($latestScore.maxScore, 2)
        $percentage   = if ($maxScore -gt 0) { [math]::Round(($currentScore / $maxScore) * 100, 1) } else { 0 }

        [PSCustomObject]@{
            Date         = $latestScore.createdDateTime
            CurrentScore = $currentScore
            MaxScore     = $maxScore
            Percentage   = $percentage
        } | Export-Csv "$outputDir\Entra_SecureScore.csv" -NoTypeInformation -Encoding UTF8

        # Fetch human-readable titles from control profiles (paginated).
        # Only store non-empty titles; some controls have null titles in the API.
        $profileTitles = @{}
        $profileUri = 'https://graph.microsoft.com/v1.0/security/secureScoreControlProfiles?$select=controlName,title&$top=250'
        while ($profileUri) {
            $profilePage = Invoke-MgGraphRequest -Method GET -Uri $profileUri -OutputType PSObject -ErrorAction Stop
            foreach ($p in $profilePage.value) {
                if ($p.controlName -and $p.title) { $profileTitles[$p.controlName] = $p.title }
            }
            $profileUri = $profilePage.'@odata.nextLink'
        }

        # Fallback: convert raw API key to a readable label by stripping vendor
        # prefixes (mdo_, AATP_, AAD_, etc.) then splitting underscores and
        # camelCase, and title-casing the result.
        $titleInfo = [System.Globalization.CultureInfo]::InvariantCulture.TextInfo
        $knownPrefixes = '^(mdo|aatp|aad|azure|intune|teams|dlp|mcas|mdca|defender|compliancepolicy)_'
        filter ConvertTo-ReadableControlName {
            $clean  = $_ -ireplace $knownPrefixes, ''
            $words  = ($clean -split '_') | ForEach-Object {
                        $_ -creplace '(?<=[a-z])(?=[A-Z])', ' '
                      }
            $titleInfo.ToTitleCase(($words -join ' ').ToLower())
        }

        $controlRows = foreach ($ctrl in $latestScore.controlScores) {
            if (-not $ctrl.controlName) { continue }
            $title = if ($profileTitles.ContainsKey($ctrl.controlName)) {
                         $profileTitles[$ctrl.controlName]
                     } else {
                         $ctrl.controlName | ConvertTo-ReadableControlName
                     }
            [PSCustomObject]@{
                ControlName  = $title
                Score        = $ctrl.score
                Description  = $ctrl.description
            }
        }
        $controlRows | Sort-Object Score -Descending |
            Export-Csv "$outputDir\Entra_SecureScoreControls.csv" -NoTypeInformation -Encoding UTF8

        Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Collecting Identity Secure Score..." -CurrentOperation "Saved: Entra_SecureScore.csv ($currentScore / $maxScore = $percentage%)" -PercentComplete ([int]($step / $totalSteps * 100))
    }
}
catch {
    Add-AuditIssue -Severity 'Warning' -Section 'Entra' -Collector 'Invoke-MgGraphRequest' -Description ($_.Exception.Message ?? "$_") -Action 'Check SecurityEvents.Read.All permissions or re-run Setup-365AuditApp.ps1'
    if ($_.Exception.Message -match '403|Forbidden|valid permissions|valid roles') {
        Write-Warning "Secure Score: permission denied (SecurityEvents.Read.All not yet granted). Re-run Setup-365AuditApp.ps1 to add the missing permission."
    }
    else {
        Write-Warning "Unable to retrieve Secure Score: $($_.Exception.Message)"
    }
    Write-Verbose "Secure Score full error: $_"
}


# ================================
# ===   Security Defaults       ===
# ================================
$step++
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Collecting Security Defaults configuration..." -PercentComplete ([int]($step / $totalSteps * 100))

try {
    $secDefaults = Get-MgPolicyIdentitySecurityDefaultEnforcementPolicy

    [PSCustomObject]@{
        SecurityDefaultsEnabled = $secDefaults.IsEnabled
    } | Export-Csv "$outputDir\Entra_SecurityDefaults.csv" -NoTypeInformation -Encoding UTF8
    Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Collecting Security Defaults configuration..." -CurrentOperation "Saved: Entra_SecurityDefaults.csv" -PercentComplete ([int]($step / $totalSteps * 100))
}
catch {
    Add-AuditIssue -Severity 'Warning' -Section 'Entra' -Collector 'Get-MgPolicyIdentitySecurityDefaultEnforcementPolicy' -Description ($_.Exception.Message ?? "$_") -Action 'Check Policy.Read.All permissions or re-run Setup-365AuditApp.ps1'
    Write-Warning "Unable to retrieve Security Defaults: $_"
}


# ================================
# ===   Enterprise Applications ===
# ================================
$step++
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Collecting enterprise app consent grants..." -PercentComplete ([int]($step / $totalSteps * 100))

try {
    # Microsoft first-party tenant IDs — filter these out to show only third-party apps
    $msTenantIds = @(
        'f8cdef31-a31e-4b4a-93e4-5f571e91255a'  # Microsoft Services
        '72f988bf-86f1-41af-91ab-2d7cd011db47'  # Microsoft
        '47d73278-e43c-4cc2-a606-c500b66883ef'  # Microsoft Partner Network
    )

    $thirdPartyApps = Get-MgServicePrincipal -All -ErrorAction Stop |
        Where-Object {
            $_.ServicePrincipalType -eq 'Application' -and
            $_.Tags -contains 'WindowsAzureActiveDirectoryIntegratedApp' -and
            $_.AppOwnerOrganizationId -notin $msTenantIds
        }

    # Cache service principal details (AppRoles + Scopes) to resolve permission names without repeated API calls.
    # Do NOT use -Property/$select — Graph API may return an empty appRoles collection when $select is present.
    $_spPermCache = @{}
    function Get-CachedSP {
        param([string]$SpId)
        if (-not $_spPermCache.ContainsKey($SpId)) {
            try {
                $_spPermCache[$SpId] = Get-MgServicePrincipal -ServicePrincipalId $SpId -ErrorAction SilentlyContinue
            } catch { $_spPermCache[$SpId] = $null }
        }
        return $_spPermCache[$SpId]
    }

    $appData     = [System.Collections.Generic.List[object]]::new()
    $appPermData = [System.Collections.Generic.List[object]]::new()

    foreach ($app in $thirdPartyApps) {
        $roleAssignments = @(Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $app.Id -ErrorAction SilentlyContinue)
        $appData.Add([PSCustomObject]@{
            DisplayName     = $app.DisplayName
            AppId           = $app.AppId
            PublisherName   = $app.PublisherName
            PublisherDomain = $app.VerifiedPublisher.DisplayName ?? $app.AppOwnerOrganizationId
            Enabled         = $app.AccountEnabled
            AdminConsented  = $roleAssignments.Count -gt 0
            ConsentedRoles  = $roleAssignments.Count
        })

        # Application permissions (app role assignments)
        foreach ($ra in $roleAssignments) {
            $_resSP   = Get-CachedSP -SpId $ra.ResourceId
            $_roleName = ($_resSP?.AppRoles | Where-Object { "$($_.Id)" -eq "$($ra.AppRoleId)" } | Select-Object -First 1)?.Value ?? "$($ra.AppRoleId)"
            $appPermData.Add([PSCustomObject]@{
                AppDisplayName  = $app.DisplayName
                PermissionType  = 'Application'
                ResourceApp     = $_resSP?.DisplayName ?? $ra.ResourceDisplayName ?? $ra.ResourceId
                PermissionName  = $_roleName
            })
        }

        # Delegated permissions (OAuth2 grants)
        try {
            $grants = @(Get-MgServicePrincipalOauth2PermissionGrant -ServicePrincipalId $app.Id -ErrorAction SilentlyContinue)
            foreach ($grant in $grants) {
                foreach ($scope in ($grant.Scope -split ' ' | Where-Object { $_ })) {
                    $_resSP2 = Get-CachedSP -SpId $grant.ResourceId
                    $appPermData.Add([PSCustomObject]@{
                        AppDisplayName  = $app.DisplayName
                        PermissionType  = 'Delegated'
                        ResourceApp     = $_resSP2?.DisplayName ?? $grant.ResourceId
                        PermissionName  = $scope
                    })
                }
            }
        } catch {}
    }

    if ($appData.Count -gt 0) {
        $appData | Export-Csv "$outputDir\Entra_EnterpriseApps.csv" -NoTypeInformation -Encoding UTF8
        Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Collecting enterprise app consent grants..." -CurrentOperation "Saved: Entra_EnterpriseApps.csv ($($appData.Count) third-party apps)" -PercentComplete ([int]($step / $totalSteps * 100))
    }
    if ($appPermData.Count -gt 0) {
        $appPermData | Export-Csv "$outputDir\Entra_EnterpriseAppPermissions.csv" -NoTypeInformation -Encoding UTF8
    }
}
catch {
    Add-AuditIssue -Severity 'Warning' -Section 'Entra' -Collector 'Get-MgServicePrincipal' -Description ($_.Exception.Message ?? "$_") -Action 'Check Application.Read.All permissions or re-run Setup-365AuditApp.ps1'
    Write-Warning "Unable to retrieve enterprise applications: $_"
}


# ================================
# ===   Identity Protection     ===
# ================================
$step++
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Collecting Identity Protection risk data..." -PercentComplete ([int]($step / $totalSteps * 100))

# Identity Protection requires Azure AD Premium P2.
# Pre-check SKUs before attempting API calls to avoid noisy warnings on unlicensed tenants.
$_p2Skus = @(
    'AAD_PREMIUM_P2', 'EMS_E5', 'EMSPREMIUM', 'SPE_E5', 'SPE_E5_USGOV_GCCHIGH',
    'M365EDU_A5_FACULTY', 'M365EDU_A5_STUDENT', 'IDENTITY_THREAT_PROTECTION',
    'IDENTITY_THREAT_PROTECTION_FOR_SMB'
)
$_hasP2 = ($subscribedSkus.SkuPartNumber | Where-Object { $_ -in $_p2Skus }).Count -gt 0

if (-not $_hasP2) {
    Write-Verbose "No Azure AD Premium P2 licence detected — skipping Identity Protection collection."
}
else {
    # Use Invoke-MgGraphRequest directly — no SDK cmdlets exist for these endpoints
    try {
        $riskyUsersResp = Invoke-MgGraphRequest -Method GET -Uri '/v1.0/identityProtection/riskyUsers?$top=500' -OutputType PSObject -ErrorAction Stop
        $riskyUsers = @($riskyUsersResp.value)
        if ($riskyUsers.Count -gt 0) {
            $riskyUsers | ForEach-Object {
                [PSCustomObject]@{
                    UserPrincipalName = $_.userPrincipalName
                    DisplayName       = $_.userDisplayName
                    RiskLevel         = $_.riskLevel
                    RiskState         = $_.riskState
                    RiskDetail        = $_.riskDetail
                    RiskLastUpdated   = $_.riskLastUpdatedDateTime
                }
            } | Export-Csv "$outputDir\Entra_RiskyUsers.csv" -NoTypeInformation -Encoding UTF8
            Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Collecting Identity Protection risk data..." -CurrentOperation "Saved: Entra_RiskyUsers.csv ($($riskyUsers.Count) risky users)" -PercentComplete ([int]($step / $totalSteps * 100))
        }
        else {
            Write-Verbose "No risky users detected — skipping Entra_RiskyUsers.csv"
        }
    }
    catch {
        Add-AuditIssue -Severity 'Warning' -Section 'Entra' -Collector 'Invoke-MgGraphRequest' -Description ($_.Exception.Message ?? "$_") -Action 'Check IdentityRiskEvent.Read.All permissions or re-run Setup-365AuditApp.ps1'
        Write-Warning "Unable to retrieve risky users: $_"
    }

    try {
        $riskySignInsResp = Invoke-MgGraphRequest -Method GET -Uri '/v1.0/identityProtection/riskySignIns?$top=500' -OutputType PSObject -ErrorAction Stop
        $riskySignIns = @($riskySignInsResp.value)
        if ($riskySignIns.Count -gt 0) {
            $riskySignIns | ForEach-Object {
                [PSCustomObject]@{
                    UserPrincipalName = $_.userPrincipalName
                    UserDisplayName   = $_.userDisplayName
                    RiskLevel         = $_.riskLevelDuringSignIn
                    RiskState         = $_.riskState
                    RiskEventTypes    = ($_.riskEventTypes -join ', ')
                    IPAddress         = $_.ipAddress
                    City              = $_.location.city
                    CountryOrRegion   = $_.location.countryOrRegion
                    AppDisplayName    = $_.appDisplayName
                    CreatedDateTime   = $_.createdDateTime
                }
            } | Export-Csv "$outputDir\Entra_RiskySignIns.csv" -NoTypeInformation -Encoding UTF8
            Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Collecting Identity Protection risk data..." -CurrentOperation "Saved: Entra_RiskySignIns.csv ($($riskySignIns.Count) risky sign-ins)" -PercentComplete ([int]($step / $totalSteps * 100))
        }
        else {
            Write-Verbose "No risky sign-ins detected — skipping Entra_RiskySignIns.csv"
        }
    }
    catch {
        Add-AuditIssue -Severity 'Warning' -Section 'Entra' -Collector 'Invoke-MgGraphRequest' -Description ($_.Exception.Message ?? "$_") -Action 'Check IdentityRiskEvent.Read.All permissions or re-run Setup-365AuditApp.ps1'
        Write-Warning "Unable to retrieve risky sign-ins: $_"
    }
}


# ================================
# ===   15. Authentication Methods Policy  ===
# ================================
$step++
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Collecting Authentication Methods Policy..." -PercentComplete ([int]($step / $totalSteps * 100))
try {
    $authMethodsPolicyResp = Invoke-MgGraphRequest -Method GET `
        -Uri 'https://graph.microsoft.com/v1.0/policies/authenticationMethodsPolicy' `
        -OutputType PSObject -ErrorAction Stop

    $_campaignTargets = @($authMethodsPolicyResp.registrationEnforcement.authenticationMethodsRegistrationCampaign.includeTargets | ForEach-Object { $_.targetType })

    $_friendlyNames = @{
        'microsoftAuthenticator' = 'Microsoft Authenticator'
        'fido2'                  = 'FIDO2 Security Key'
        'sms'                    = 'SMS'
        'voice'                  = 'Voice Call'
        'email'                  = 'Email OTP'
        'temporaryAccessPass'    = 'Temporary Access Pass'
        'softwareOath'           = 'Software OATH Token'
        'x509Certificate'        = 'Certificate-Based Auth'
        'windowsHelloForBusiness'= 'Windows Hello for Business'
        'hardwareOath'           = 'Hardware OATH Token'
    }

    $authMethodRows = foreach ($method in @($authMethodsPolicyResp.authenticationMethodConfigurations)) {
        $friendlyName = if ($_friendlyNames.ContainsKey($method.id)) { $_friendlyNames[$method.id] } else { $method.id }
        [PSCustomObject]@{
            MethodType             = $friendlyName
            MethodId               = $method.id
            State                  = $method.state
            IsRegistrationRequired = ($method.id -in $_campaignTargets)
        }
    }
    $authMethodRows | Export-Csv "$outputDir\Entra_AuthMethodsPolicy.csv" -NoTypeInformation -Encoding UTF8
    Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Collecting Authentication Methods Policy..." -CurrentOperation "Saved: Entra_AuthMethodsPolicy.csv ($($authMethodRows.Count) methods)" -PercentComplete ([int]($step / $totalSteps * 100))
}
catch {
    Add-AuditIssue -Severity 'Warning' -Section 'Entra' -Collector 'Invoke-MgGraphRequest' -Description ($_.Exception.Message ?? "$_") -Action 'Check Policy.Read.All permissions or re-run Setup-365AuditApp.ps1'
    Write-Warning "Unable to retrieve Authentication Methods Policy: $_"
}


# ================================
# ===   16. External Collaboration Settings  ===
# ================================
$step++
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Collecting External Collaboration Settings..." -PercentComplete ([int]($step / $totalSteps * 100))
try {
    $authzPolicy = Invoke-MgGraphRequest -Method GET `
        -Uri 'https://graph.microsoft.com/v1.0/policies/authorizationPolicy' `
        -OutputType PSObject -ErrorAction Stop

    $_guestRoleMap = @{
        '10dae51f-b6af-4016-8d66-8c2a99b929b7' = 'Restricted Guest User (most secure)'
        '2af84b1e-32c8-42b7-82bc-daa82404023b' = 'Guest User (standard)'
        'a0b1b346-4d3e-4e8b-98f8-753987be4970' = 'Member (least secure)'
    }
    $_guestRoleId   = $authzPolicy.guestUserRoleId
    $_guestRoleName = if ($_guestRoleMap.ContainsKey($_guestRoleId)) { $_guestRoleMap[$_guestRoleId] } else { $_guestRoleId }

    $_extIdPolicy = $null
    try {
        $_extIdPolicy = Invoke-MgGraphRequest -Method GET `
            -Uri 'https://graph.microsoft.com/v1.0/policies/externalIdentitiesPolicy' `
            -OutputType PSObject -ErrorAction Stop
    }
    catch {
        Write-Verbose "externalIdentitiesPolicy not available (may not be configured): $_"
    }

    [PSCustomObject]@{
        AllowInvitesFrom                          = $authzPolicy.allowInvitesFrom
        GuestUserRoleId                           = $_guestRoleId
        GuestUserRoleName                         = $_guestRoleName
        AllowedToSignUpEmailBasedSubscriptions    = if ($null -ne $_extIdPolicy) { $_extIdPolicy.allowedToSignUpEmailBasedSubscriptions } else { '' }
        AllowEmailVerifiedUsersToJoinOrganization = if ($null -ne $_extIdPolicy) { $_extIdPolicy.allowEmailVerifiedUsersToJoinOrganization } else { '' }
    } | Export-Csv "$outputDir\Entra_ExternalCollab.csv" -NoTypeInformation -Encoding UTF8
    Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Collecting External Collaboration Settings..." -CurrentOperation "Saved: Entra_ExternalCollab.csv" -PercentComplete ([int]($step / $totalSteps * 100))
}
catch {
    Add-AuditIssue -Severity 'Warning' -Section 'Entra' -Collector 'Invoke-MgGraphRequest' -Description ($_.Exception.Message ?? "$_") -Action 'Check Policy.Read.All permissions or re-run Setup-365AuditApp.ps1'
    Write-Warning "Unable to retrieve External Collaboration Settings: $_"
}


# ================================
# ===   17. App Registrations   ===
# ================================
$step++
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Collecting App Registrations..." -PercentComplete ([int]($step / $totalSteps * 100))
try {
    $_appRegs    = [System.Collections.Generic.List[object]]::new()
    $_appNextUri = 'https://graph.microsoft.com/v1.0/applications?$select=displayName,appId,createdDateTime,passwordCredentials,keyCredentials,requiredResourceAccess&$top=999'
    while ($_appNextUri) {
        $_appPage    = Invoke-MgGraphRequest -Method GET -Uri $_appNextUri -OutputType PSObject -ErrorAction Stop
        $_appNextUri = $_appPage.'@odata.nextLink'
        foreach ($app in @($_appPage.value)) { $_appRegs.Add($app) }
    }

    $appRegRows = foreach ($app in $_appRegs) {
        $creds = @()
        foreach ($cred in @($app.passwordCredentials)) {
            if ($null -ne $cred) { $creds += [PSCustomObject]@{ Type = 'Password'; Name = $cred.displayName; Expiry = $cred.endDateTime } }
        }
        foreach ($cred in @($app.keyCredentials)) {
            if ($null -ne $cred) { $creds += [PSCustomObject]@{ Type = 'Certificate'; Name = $cred.displayName; Expiry = $cred.endDateTime } }
        }
        if ($creds.Count -eq 0) {
            [PSCustomObject]@{
                DisplayName      = $app.displayName
                AppId            = $app.appId
                CreatedDateTime  = $app.createdDateTime
                CredentialType   = 'None'
                CredentialName   = ''
                CredentialExpiry = ''
                DaysUntilExpiry  = ''
            }
        }
        else {
            foreach ($cred in $creds) {
                $days = if ($null -ne $cred.Expiry) {
                    [int]([datetime]$cred.Expiry - [datetime]::UtcNow).TotalDays
                }
                else { $null }
                [PSCustomObject]@{
                    DisplayName      = $app.displayName
                    AppId            = $app.appId
                    CreatedDateTime  = $app.createdDateTime
                    CredentialType   = $cred.Type
                    CredentialName   = $cred.Name
                    CredentialExpiry = if ($null -ne $cred.Expiry) { $cred.Expiry } else { '' }
                    DaysUntilExpiry  = if ($null -ne $days) { $days } else { '' }
                }
            }
        }
    }
    $appRegRows | Export-Csv "$outputDir\Entra_AppRegistrations.csv" -NoTypeInformation -Encoding UTF8
    Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Collecting App Registrations..." -CurrentOperation "Saved: Entra_AppRegistrations.csv ($($_appRegs.Count) registrations)" -PercentComplete ([int]($step / $totalSteps * 100))

    # Collect App Registration permissions (requiredResourceAccess → resolved permission names)
    # Uses the same $_spPermCache built during enterprise apps if that step ran first; falls back gracefully.
    if (-not (Get-Variable -Name '_spPermCache' -Scope Script -ErrorAction SilentlyContinue)) {
        $_spPermCache = @{}
        function Get-CachedSP {
            param([string]$SpId)
            if (-not $_spPermCache.ContainsKey($SpId)) {
                try {
                    # ConsistencyLevel eventual required for appId eq filter on /servicePrincipals
                    # No -Property: $select causes Graph API to silently return empty appRoles
                    $_spPermCache[$SpId] = Get-MgServicePrincipal -Filter "appId eq '$SpId'" `
                        -ConsistencyLevel eventual -ErrorAction SilentlyContinue | Select-Object -First 1
                } catch { $_spPermCache[$SpId] = $null }
            }
            return $_spPermCache[$SpId]
        }
        $_arSpLookupByAppId = $true
    } else {
        $_arSpLookupByAppId = $false
    }

    $_arPermRows = [System.Collections.Generic.List[object]]::new()
    foreach ($app in $_appRegs) {
        foreach ($resource in @($app.requiredResourceAccess)) {
            if (-not $resource) { continue }
            $_resAppId = $resource.resourceAppId
            # Lookup key differs: enterprise apps cache uses SP object ID; app reg cache uses appId
            if ($_arSpLookupByAppId) {
                $_resSP = Get-CachedSP -SpId $_resAppId
            } else {
                # Enterprise apps cache is keyed by SP object ID; look up by appId into a separate key
                if (-not $_spPermCache.ContainsKey($_resAppId)) {
                    try {
                        $_spPermCache[$_resAppId] = Get-MgServicePrincipal -Filter "appId eq '$_resAppId'" `
                            -ConsistencyLevel eventual -ErrorAction SilentlyContinue | Select-Object -First 1
                    } catch { $_spPermCache[$_resAppId] = $null }
                }
                $_resSP = $_spPermCache[$_resAppId]
            }
            foreach ($access in @($resource.resourceAccess)) {
                if (-not $access) { continue }
                $_permName = if ($access.type -eq 'Role') {
                    ($_resSP?.AppRoles | Where-Object { "$($_.Id)" -eq "$($access.id)" } | Select-Object -First 1)?.Value ?? "$($access.id)"
                } else {
                    ($_resSP?.Oauth2PermissionScopes | Where-Object { "$($_.Id)" -eq "$($access.id)" } | Select-Object -First 1)?.Value ?? "$($access.id)"
                }
                $_arPermRows.Add([PSCustomObject]@{
                    AppDisplayName  = $app.displayName
                    PermissionType  = if ($access.type -eq 'Role') { 'Application' } else { 'Delegated' }
                    ResourceApp     = $_resSP?.DisplayName ?? $_resAppId
                    PermissionName  = $_permName
                })
            }
        }
    }
    if ($_arPermRows.Count -gt 0) {
        $_arPermRows | Export-Csv "$outputDir\Entra_AppRegistrationPermissions.csv" -NoTypeInformation -Encoding UTF8
    }
}
catch {
    Add-AuditIssue -Severity 'Warning' -Section 'Entra' -Collector 'Invoke-MgGraphRequest' -Description ($_.Exception.Message ?? "$_") -Action 'Check Application.Read.All permissions or re-run Setup-365AuditApp.ps1'
    Write-Warning "Unable to retrieve App Registrations: $_"
}


# ================================
# ===   18. PIM Role Assignments (P2 required)  ===
# ================================
$step++
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Collecting PIM Role Assignments..." -PercentComplete ([int]($step / $totalSteps * 100))
if (-not $_hasP2) {
    Write-Verbose "No Azure AD Premium P2 licence detected — skipping PIM role assignment collection."
}
else {
    try {
        $_pimRows    = [System.Collections.Generic.List[object]]::new()
        $_pimNextUri = 'https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignmentScheduleInstances?$expand=principal,roleDefinition&$top=999'
        while ($_pimNextUri) {
            $_pimPage    = Invoke-MgGraphRequest -Method GET -Uri $_pimNextUri -OutputType PSObject -ErrorAction Stop
            $_pimNextUri = $_pimPage.'@odata.nextLink'
            foreach ($inst in @($_pimPage.value)) { $_pimRows.Add($inst) }
        }

        $_pimRows | ForEach-Object {
            [PSCustomObject]@{
                RoleName              = $_.roleDefinition.displayName
                PrincipalDisplayName  = $_.principal.displayName
                PrincipalUPN          = if ($_.principal.userPrincipalName) { $_.principal.userPrincipalName } else { '' }
                AssignmentType        = $_.assignmentType
                MemberType            = $_.memberType
                StartDateTime         = if ($_.startDateTime) { $_.startDateTime } else { '' }
                EndDateTime           = if ($_.endDateTime)   { $_.endDateTime   } else { '' }
            }
        } | Export-Csv "$outputDir\Entra_PIMAssignments.csv" -NoTypeInformation -Encoding UTF8
        Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Collecting PIM Role Assignments..." -CurrentOperation "Saved: Entra_PIMAssignments.csv ($($_pimRows.Count) assignments)" -PercentComplete ([int]($step / $totalSteps * 100))
    }
    catch {
        Add-AuditIssue -Severity 'Warning' -Section 'Entra' -Collector 'Invoke-MgGraphRequest' -Description ($_.Exception.Message ?? "$_") -Action 'Check RoleManagement.Read.All permissions or re-run Setup-365AuditApp.ps1'
        Write-Warning "Unable to retrieve PIM role assignments: $_"
    }
}


# ================================
# ===   19. Org-level User / App Settings ===
# ================================
$step++
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Collecting organisation-level user and app settings..." -PercentComplete ([int]($step / $totalSteps * 100))
try {
    # Authorization policy: user default role permissions + guest invite settings
    $_orgAuthzPolicy = Invoke-MgGraphRequest -Method GET `
        -Uri 'https://graph.microsoft.com/v1.0/policies/authorizationPolicy' `
        -OutputType PSObject -ErrorAction Stop

    $_defaultPerms = $_orgAuthzPolicy.defaultUserRolePermissions

    # Admin consent request workflow
    $_adminConsentPolicy = $null
    try {
        $_adminConsentPolicy = Invoke-MgGraphRequest -Method GET `
            -Uri 'https://graph.microsoft.com/v1.0/policies/adminConsentRequestPolicy' `
            -OutputType PSObject -ErrorAction Stop
    }
    catch { Write-Verbose "Admin consent request policy not available: $_" }

    [PSCustomObject]@{
        AllowedToCreateApps             = $null -ne $_defaultPerms.allowedToCreateApps ? "$($_defaultPerms.allowedToCreateApps)" : ''
        AllowedToCreateTenants          = $null -ne $_defaultPerms.allowedToCreateTenants ? "$($_defaultPerms.allowedToCreateTenants)" : ''
        AllowedToCreateSecurityGroups   = $null -ne $_defaultPerms.allowedToCreateSecurityGroups ? "$($_defaultPerms.allowedToCreateSecurityGroups)" : ''
        AllowedToReadBitlockerKeys      = $null -ne $_defaultPerms.allowedToReadBitlockerKeysForOwnedDevice ? "$($_defaultPerms.allowedToReadBitlockerKeysForOwnedDevice)" : ''
        DefaultUserRolePermissionsJson  = ($null -ne $_defaultPerms) ? ($_defaultPerms | ConvertTo-Json -Compress -Depth 3) : ''
        AdminConsentWorkflowEnabled     = if ($_adminConsentPolicy) { "$($_adminConsentPolicy.isEnabled)" } else { '' }
    } | Export-Csv "$outputDir\Entra_OrgSettings.csv" -NoTypeInformation -Encoding UTF8
    Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Collecting organisation-level user and app settings..." -CurrentOperation "Saved: Entra_OrgSettings.csv" -PercentComplete ([int]($step / $totalSteps * 100))
}
catch {
    Add-AuditIssue -Severity 'Warning' -Section 'Entra' -Collector 'Invoke-MgGraphRequest' -Description ($_.Exception.Message ?? "$_") -Action 'Check Policy.Read.All permissions or re-run Setup-365AuditApp.ps1'
    Write-Warning "Unable to retrieve organisation settings: $_"
}


# ================================
# ===   Done                    ===
# ================================
Write-Progress -Id 1 -Activity $activity -Completed
Write-Host "`nEntra Audit complete. Results saved to: $outputDir`n" -ForegroundColor Green
