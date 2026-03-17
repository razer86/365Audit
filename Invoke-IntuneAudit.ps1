<#
.SYNOPSIS
    Performs a security-focused audit of Microsoft Intune / Endpoint Manager.

.DESCRIPTION
    Connects to Microsoft Graph and collects Intune device management data including:
    - Licence check (skips gracefully if no qualifying Intune SKU is found)
    - Managed device inventory with compliance state and ownership
    - Per-device compliance policy states
    - Compliance policies with platform, assignments, grace period, and full settings
    - Configuration profiles with platform, type, and assignment scope
    - Managed application install summary
    - Autopilot device identities
    - Enrollment restrictions

    Output CSVs:
    - Intune_LicenceCheck.csv
    - Intune_Devices.csv
    - Intune_DeviceComplianceStates.csv
    - Intune_CompliancePolicies.csv
    - Intune_CompliancePolicySettings.csv
    - Intune_ConfigProfiles.csv
    - Intune_ConfigProfileSettings.csv
    - Intune_Apps.csv
    - Intune_AutopilotDevices.csv
    - Intune_EnrollmentRestrictions.csv

.NOTES
    Author      : Raymond Slater
    Version     : 1.3.0
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

$ScriptVersion = "1.3.0"
Write-Verbose "Invoke-IntuneAudit.ps1 loaded (v$ScriptVersion)"

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

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
    $outputDir = Join-Path $context.OutputPath "Intune"
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
}
catch {
    Write-Error "Failed to initialise audit output directory: $_"
    exit 1
}

# === Connect to Microsoft Graph ===
try {
    Connect-MgGraphSecure
}
catch {
    Write-Error "Microsoft Graph connection failed: $_"
    exit 1
}

Write-Host "`nStarting Intune Audit for $($context.OrgName)..." -ForegroundColor Cyan

$step       = 0
$totalSteps = 8
$activity   = "Intune Audit — $($context.OrgName)"

# Known SKU part numbers that include Intune
$_intuneSkus = @(
    'INTUNE_A',
    'EMS_S_1',
    'EMS_S_3',
    'EMS_S_5',
    'ENTERPRISEPREMIUM',
    'ENTERPRISEPACK',
    'SPB',
    'BUSINESS_PREMIUM',
    'M365_F1',
    'M365_F3'
)

# Metadata keys to skip when iterating AdditionalProperties on compliance policies
$_odataSkipKeys = @(
    '@odata.type', 'id', 'createdDateTime', 'lastModifiedDateTime',
    'displayName', 'description', 'version', 'roleScopeTagIds', 'scheduledActionsForRule'
)


function Get-IntuneGraphErrorMessage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [System.Management.Automation.ErrorRecord]$ErrorRecord,

        [Parameter(Mandatory)]
        [string]$Operation
    )

    $message = $ErrorRecord.Exception.Message

    if ($message -match 'Application is not authorized to perform this operation\. Application must have one of the following scopes:\s*(?<scopes>.+?)\s*-\s*Operation ID') {
        $requiredPermissions = ($matches.scopes -split ',\s*' | Where-Object { $_ }) -join ', '
        return "$Operation requires Intune application permission(s): $requiredPermissions. Grant admin consent to the 365Audit app registration and rerun."
    }

    if ($message -match "Invalid filter clause: Could not find a property named 'isAssigned'") {
        return "$Operation failed because the Microsoft Graph mobile app endpoint does not support filtering on 'isAssigned'."
    }

    return "$Operation failed: $message"
}


# ================================
# ===   Step 1 — Licence Check  ===
# ================================
$step++
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Checking Intune licence..." -PercentComplete ([int]($step / $totalSteps * 100))

$_subscribedSkus = @(Get-MgSubscribedSku -All -ErrorAction SilentlyContinue)
$_intuneLicSkus  = @($_subscribedSkus | Where-Object { $_.SkuPartNumber -in $_intuneSkus })
$_hasIntune      = $_intuneLicSkus.Count -gt 0

[PSCustomObject]@{
    HasIntune   = $_hasIntune
    LicencedSKUs = if ($_hasIntune) { ($_intuneLicSkus.SkuPartNumber -join ', ') } else { '' }
} | Export-Csv -Path (Join-Path $outputDir 'Intune_LicenceCheck.csv') -NoTypeInformation -Encoding UTF8

if (-not $_hasIntune) {
    Write-Warning "No Intune-capable licence found for $($context.OrgName). Skipping Intune data collection."
    Write-Progress -Id 1 -Activity $activity -Completed
    return
}

Write-Host "  Intune licence confirmed: $($_intuneLicSkus.SkuPartNumber -join ', ')" -ForegroundColor Green


# =======================================
# ===   Step 2 — Managed Devices      ===
# =======================================
$step++
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Retrieving managed devices..." -PercentComplete ([int]($step / $totalSteps * 100))

$_allDevices = @()
try {
    $_allDevices = @(Get-MgDeviceManagementManagedDevice -All -ErrorAction Stop)
    $_deviceRows = foreach ($_dev in $_allDevices) {
        [PSCustomObject]@{
            DeviceName         = $_dev.DeviceName
            OS                 = $_dev.OperatingSystem
            OSVersion          = $_dev.OsVersion
            DeviceType         = $_dev.DeviceType
            OwnerType          = $_dev.ManagedDeviceOwnerType
            EnrolledDateTime   = $_dev.EnrolledDateTime
            LastSyncDateTime   = $_dev.LastSyncDateTime
            ComplianceState    = $_dev.ComplianceState
            AssignedUser       = $_dev.UserPrincipalName
            Manufacturer       = $_dev.Manufacturer
            Model              = $_dev.Model
            SerialNumber       = $_dev.SerialNumber
            ManagementAgent    = $_dev.ManagementAgent
            AzureADRegistered  = $_dev.AzureADRegistered
            JoinType           = $_dev.JoinType
        }
    }
    $_deviceRows | Export-Csv -Path (Join-Path $outputDir 'Intune_Devices.csv') -NoTypeInformation -Encoding UTF8
    Write-Host "  Devices: $($_allDevices.Count) managed device(s) found." -ForegroundColor Green
}
catch {
    Write-Warning (Get-IntuneGraphErrorMessage -ErrorRecord $_ -Operation 'Managed devices')
    @() | Export-Csv -Path (Join-Path $outputDir 'Intune_Devices.csv') -NoTypeInformation -Encoding UTF8
}


# ================================================
# ===   Step 3 — Per-Device Compliance States  ===
# ================================================
$step++
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Retrieving per-device compliance policy states..." -PercentComplete ([int]($step / $totalSteps * 100))

$_complianceStateRows = [System.Collections.Generic.List[object]]::new()
foreach ($_dev in $_allDevices) {
    try {
        $_states = @(Get-MgDeviceManagementManagedDeviceCompliancePolicyState `
            -ManagedDeviceId $_dev.Id -All -ErrorAction Stop)
        foreach ($_state in $_states) {
            $_complianceStateRows.Add([PSCustomObject]@{
                DeviceName            = $_dev.DeviceName
                PolicyName            = $_state.DisplayName
                State                 = $_state.State
                LastReportedDateTime  = $_state.LastReportedDateTime
            })
        }
    }
    catch {
        Write-Verbose "Could not retrieve compliance states for device '$($_dev.DeviceName)': $_"
    }
}
$_complianceStateRows | Export-Csv -Path (Join-Path $outputDir 'Intune_DeviceComplianceStates.csv') -NoTypeInformation -Encoding UTF8
Write-Host "  Compliance states: $($_complianceStateRows.Count) policy state record(s)." -ForegroundColor Green


# ==================================================
# ===   Step 4 — Compliance Policies + Settings  ===
# ==================================================
$step++
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Retrieving compliance policies..." -PercentComplete ([int]($step / $totalSteps * 100))

$_policyRows        = [System.Collections.Generic.List[object]]::new()
$_policySettingRows = [System.Collections.Generic.List[object]]::new()

try {
    $_policies = @(Get-MgDeviceManagementDeviceCompliancePolicy -All -ErrorAction Stop)
    foreach ($_pol in $_policies) {
        # Assignments
        $_assignments = @()
        try {
            $_assignments = @(Get-MgDeviceManagementDeviceCompliancePolicyAssignment `
                -DeviceCompliancePolicyId $_pol.Id -All -ErrorAction SilentlyContinue)
        } catch {}
        $_assignedTo = if ($_assignments.Count -eq 0) { 'None' }
                       elseif ($_assignments | Where-Object { $_.Target.AdditionalProperties.'@odata.type' -like '*allDevices*' }) { 'All Devices' }
                       elseif ($_assignments | Where-Object { $_.Target.AdditionalProperties.'@odata.type' -like '*allLicensedUsers*' }) { 'All Users' }
                       else { "$($_assignments.Count) group(s)" }

        # Grace period (hours) from scheduled action rules
        $_gracePeriod = 0
        try {
            $_scheduledActions = @(Get-MgDeviceManagementDeviceCompliancePolicyScheduledActionsForRule `
                -DeviceCompliancePolicyId $_pol.Id -All -ErrorAction SilentlyContinue)
            foreach ($_sa in $_scheduledActions) {
                foreach ($_cfg in @($_sa.ScheduledActionConfigurations)) {
                    if ($_cfg.ActionType -eq 'block' -and $_cfg.GracePeriodHours -gt $_gracePeriod) {
                        $_gracePeriod = $_cfg.GracePeriodHours
                    }
                }
            }
        } catch {}

        # Derive platform from OData type
        $_odataType = $_pol.AdditionalProperties.'@odata.type' ?? ''
        $_platform  = switch -Wildcard ($_odataType) {
            '*windows10*'  { 'Windows 10/11' }
            '*ios*'        { 'iOS' }
            '*android*'    { 'Android' }
            '*macOS*'      { 'macOS' }
            default        { $_odataType -replace '#microsoft.graph.', '' }
        }

        $_policyRows.Add([PSCustomObject]@{
            PolicyName       = $_pol.DisplayName
            Platform         = $_platform
            AssignedTo       = $_assignedTo
            GracePeriodHours = $_gracePeriod
        })

        # Settings — iterate AdditionalProperties, skip metadata keys
        foreach ($_kv in $_pol.AdditionalProperties.GetEnumerator()) {
            if ($_kv.Key -in $_odataSkipKeys) { continue }
            $_policySettingRows.Add([PSCustomObject]@{
                PolicyName   = $_pol.DisplayName
                Platform     = $_platform
                SettingName  = $_kv.Key
                SettingValue = $_kv.Value
            })
        }
    }

    $_policyRows        | Export-Csv -Path (Join-Path $outputDir 'Intune_CompliancePolicies.csv')       -NoTypeInformation -Encoding UTF8
    $_policySettingRows | Export-Csv -Path (Join-Path $outputDir 'Intune_CompliancePolicySettings.csv') -NoTypeInformation -Encoding UTF8
    Write-Host "  Compliance policies: $($_policies.Count) policy(ies), $($_policySettingRows.Count) setting record(s)." -ForegroundColor Green
}
catch {
    Write-Warning (Get-IntuneGraphErrorMessage -ErrorRecord $_ -Operation 'Compliance policies')
    @() | Export-Csv -Path (Join-Path $outputDir 'Intune_CompliancePolicies.csv')       -NoTypeInformation -Encoding UTF8
    @() | Export-Csv -Path (Join-Path $outputDir 'Intune_CompliancePolicySettings.csv') -NoTypeInformation -Encoding UTF8
}


# =================================================
# ===   Step 5 — Configuration Profiles          ===
# =================================================
$step++
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Retrieving configuration profiles..." -PercentComplete ([int]($step / $totalSteps * 100))

$_profileSettingRows = [System.Collections.Generic.List[object]]::new()

try {
    $_configProfiles = @(Get-MgDeviceManagementDeviceConfiguration -All -ErrorAction Stop)
    $_profileRows = foreach ($_prof in $_configProfiles) {
        $_profAssignments = @()
        try {
            $_profAssignments = @(Get-MgDeviceManagementDeviceConfigurationAssignment `
                -DeviceConfigurationId $_prof.Id -All -ErrorAction SilentlyContinue)
        } catch {}
        $_profAssignedTo = if ($_profAssignments.Count -eq 0) { 'None' }
                           elseif ($_profAssignments | Where-Object { $_.Target.AdditionalProperties.'@odata.type' -like '*allDevices*' }) { 'All Devices' }
                           elseif ($_profAssignments | Where-Object { $_.Target.AdditionalProperties.'@odata.type' -like '*allLicensedUsers*' }) { 'All Users' }
                           else { "$($_profAssignments.Count) group(s)" }

        $_profOdata    = $_prof.AdditionalProperties.'@odata.type' ?? ''
        $_profPlatform = switch -Wildcard ($_profOdata) {
            '*windows10*'  { 'Windows 10/11' }
            '*ios*'        { 'iOS' }
            '*android*'    { 'Android' }
            '*macOS*'      { 'macOS' }
            default        { $_profOdata -replace '#microsoft.graph.', '' }
        }

        # Settings — iterate AdditionalProperties, skip metadata keys
        foreach ($_kv in $_prof.AdditionalProperties.GetEnumerator()) {
            if ($_kv.Key -in $_odataSkipKeys) { continue }
            $_profileSettingRows.Add([PSCustomObject]@{
                ProfileName  = $_prof.DisplayName
                Platform     = $_profPlatform
                ProfileType  = $_profOdata #-replace '#microsoft.graph.', ''
                SettingName  = $_kv.Key
                SettingValue = $($_kv.Value)
            })
        }

        [PSCustomObject]@{
            ProfileName          = $_prof.DisplayName
            Platform             = $_profPlatform
            ProfileType          = $_profOdata -replace '#microsoft.graph.', ''
            LastModifiedDateTime = $_prof.LastModifiedDateTime
            AssignedTo           = $_profAssignedTo
        }
    }
    $_profileRows        | Export-Csv -Path (Join-Path $outputDir 'Intune_ConfigProfiles.csv')        -NoTypeInformation -Encoding UTF8
    $_profileSettingRows | Export-Csv -Path (Join-Path $outputDir 'Intune_ConfigProfileSettings.csv') -NoTypeInformation -Encoding UTF8
    Write-Host "  Configuration profiles: $($_configProfiles.Count) profile(s), $($_profileSettingRows.Count) setting record(s)." -ForegroundColor Green
}
catch {
    Write-Warning (Get-IntuneGraphErrorMessage -ErrorRecord $_ -Operation 'Configuration profiles')
    @() | Export-Csv -Path (Join-Path $outputDir 'Intune_ConfigProfiles.csv')        -NoTypeInformation -Encoding UTF8
    @() | Export-Csv -Path (Join-Path $outputDir 'Intune_ConfigProfileSettings.csv') -NoTypeInformation -Encoding UTF8
}


# =============================================
# ===   Step 6 — Apps + Install Summary     ===
# =============================================
$step++
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Retrieving assigned apps..." -PercentComplete ([int]($step / $totalSteps * 100))

try {
    $_apps = @(Get-MgDeviceAppManagementMobileApp -All -ErrorAction Stop)
    $_appRows = foreach ($_app in $_apps) {
        $_summary = $null
        try {
            $_summary = Get-MgDeviceAppManagementMobileAppInstallSummary -MobileAppId $_app.Id -ErrorAction SilentlyContinue
        } catch {}

        [PSCustomObject]@{
            AppName              = $_app.DisplayName
            AppType              = $_app.AdditionalProperties.'@odata.type' -replace '#microsoft.graph.', ''
            Publisher            = $_app.Publisher
            InstalledDeviceCount = if ($_summary) { $_summary.InstalledDeviceCount } else { 0 }
            FailedDeviceCount    = if ($_summary) { $_summary.FailedDeviceCount    } else { 0 }
            PendingInstallCount  = if ($_summary) { $_summary.PendingInstallCount  } else { 0 }
        }
    }
    $_appRows | Export-Csv -Path (Join-Path $outputDir 'Intune_Apps.csv') -NoTypeInformation -Encoding UTF8
    Write-Host "  Apps: $($_apps.Count) app(s)." -ForegroundColor Green
}
catch {
    Write-Warning (Get-IntuneGraphErrorMessage -ErrorRecord $_ -Operation 'Apps')
    @() | Export-Csv -Path (Join-Path $outputDir 'Intune_Apps.csv') -NoTypeInformation -Encoding UTF8
}


# =============================================
# ===   Step 7 — Autopilot Devices          ===
# =============================================
$step++
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Retrieving Autopilot devices..." -PercentComplete ([int]($step / $totalSteps * 100))

try {
    $_autopilot = @(Get-MgDeviceManagementWindowsAutopilotDeviceIdentity -All -ErrorAction Stop)
    $_autopilotRows = foreach ($_ap in $_autopilot) {
        [PSCustomObject]@{
            SerialNumber          = $_ap.SerialNumber
            Model                 = $_ap.Model
            Manufacturer          = $_ap.Manufacturer
            GroupTag              = $_ap.GroupTag
            AssignedUser          = $_ap.UserPrincipalName
            EnrollmentState       = $_ap.EnrollmentState
            LastContactedDateTime = $_ap.LastContactedDateTime
        }
    }
    $_autopilotRows | Export-Csv -Path (Join-Path $outputDir 'Intune_AutopilotDevices.csv') -NoTypeInformation -Encoding UTF8
    Write-Host "  Autopilot: $($_autopilot.Count) device(s) registered." -ForegroundColor Green
}
catch {
    Write-Warning (Get-IntuneGraphErrorMessage -ErrorRecord $_ -Operation 'Autopilot devices')
    @() | Export-Csv -Path (Join-Path $outputDir 'Intune_AutopilotDevices.csv') -NoTypeInformation -Encoding UTF8
}


# ===================================================
# ===   Step 8 — Enrollment Restrictions           ===
# ===================================================
$step++
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Retrieving enrollment restrictions..." -PercentComplete ([int]($step / $totalSteps * 100))

try {
    $_restrictions = @(Get-MgDeviceManagementDeviceEnrollmentConfiguration -All -ErrorAction Stop)
    $_restrictionRows = foreach ($_res in $_restrictions) {
        $_resAssignments = @()
        try {
            $_resAssignments = @(Get-MgDeviceManagementDeviceEnrollmentConfigurationAssignment `
                -DeviceEnrollmentConfigurationId $_res.Id -All -ErrorAction SilentlyContinue)
        } catch {}
        $_resAssignedTo = if ($_resAssignments.Count -eq 0) { 'None' }
                          elseif ($_resAssignments | Where-Object { $_.Target.AdditionalProperties.'@odata.type' -like '*allDevices*' }) { 'All Devices' }
                          elseif ($_resAssignments | Where-Object { $_.Target.AdditionalProperties.'@odata.type' -like '*allLicensedUsers*' }) { 'All Users' }
                          else { "$($_resAssignments.Count) group(s)" }

        $_ap = $_res.AdditionalProperties
        $_blockPersonal  = if ($null -ne $_ap.platformRestrictions) {
            $platforms = @('android', 'androidForWork', 'ios', 'mac', 'windows')
            $blocked = $false
            foreach ($p in $platforms) {
                if ($_ap.platformRestrictions.$p.personalDeviceEnrollmentBlocked -eq $true) { $blocked = $true; break }
            }
            $blocked
        } else { $false }

        $_maxDevices = if ($null -ne $_ap.limit) { $_ap.limit } else { 'N/A' }

        [PSCustomObject]@{
            ConfigName           = $_res.DisplayName
            Platform             = $_res.AdditionalProperties.'@odata.type' -replace '#microsoft.graph.', ''
            BlockPersonalDevices = $_blockPersonal
            MaxDevicesPerUser    = $_maxDevices
            Priority             = $_res.Priority
            AssignedTo           = $_resAssignedTo
        }
    }
    $_restrictionRows | Export-Csv -Path (Join-Path $outputDir 'Intune_EnrollmentRestrictions.csv') -NoTypeInformation -Encoding UTF8
    Write-Host "  Enrollment restrictions: $($_restrictions.Count) configuration(s)." -ForegroundColor Green
}
catch {
    Write-Warning (Get-IntuneGraphErrorMessage -ErrorRecord $_ -Operation 'Enrollment restrictions')
    @() | Export-Csv -Path (Join-Path $outputDir 'Intune_EnrollmentRestrictions.csv') -NoTypeInformation -Encoding UTF8
}


Write-Progress -Id 1 -Activity $activity -Completed
Write-Host "`nIntune Audit complete. Output saved to: $outputDir" -ForegroundColor Green
