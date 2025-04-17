param (
    [string]$AuditFolder,
    [switch]$DevMode = $false
)

if (-not $DevMode -and $MyInvocation.InvocationName -eq $MyInvocation.MyCommand.Name) {
    Write-Error "This script must be run from the 365Audit launcher. Use -DevMode for development." -ErrorAction Stop
}

﻿<#
.SYNOPSIS
    Performs a security-focused audit of Entra users into a consolidated CSV.

.DESCRIPTION
    This script connects to Microsoft Graph and collects a wide range of identity-related data, including:
    - User UPN, first name, last name
    - Assigned licenses with friendly names
    - MFA status and method types
    - Password policy and last change date
    - Administrative role assignments
    - Guest user list
    - SSPR (Self-Service Password Reset) configuration

    Output CSVs:
    - Entra_Users.csv      # Consolidated user info, license, MFA, password, and sign-in data
    - Entra_AdminRoles.csv        # All admin role assignments
    - Entra_GlobalAdmins.csv      # Subset of Global Admins
    - Entra_GuestUsers.csv        # List of all guest users
    - Entra_SSPR.csv              # Status of self-service password reset

.NOTES
    Author      : Raymond Slater
    Version     : 1.0.2
    Change Log  :
        1.0.0 - Initial release
        1.0.1 - Refactor output directory initialization
        1.0.2 - Combine user info, license, MFA, and sign-in into a single export

.LINK
    https://github.com/razer86/365Audit
#>

# Force UTF-8 output for emoji if supported
if ($PSVersionTable.PSVersion.Major -ge 6) {
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
}

# === Convert Sku Names to Friendly Name ===
function Get-FriendlySkuName {
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
        "SMB_APPS"                    = "Business Apps (Free)"
    } 

    if ($skuMap.ContainsKey($Sku)) {
        return $skuMap[$Sku]
    } else {
        return $Sku
    }
}

# === Ensure helper functions available ===
if (-not (Get-Command Connect-MgGraphSecure -ErrorAction SilentlyContinue)) {
    Write-Error "Connect-MgGraphSecure is not loaded. Please ensure the launcher was used."
    exit 1
}
if (-not (Get-Command Initialize-AuditOutput -ErrorAction SilentlyContinue)) {
    Write-Error "Initialize-AuditOutput is not loaded. Please ensure the launcher was used."
    exit 1
}

# === Retrieve Output Folder ===
try {
    $context = Initialize-AuditOutput
    $outputDir = $context.OutputPath
}
catch {
    Write-Error "❌ Failed to locate audit output directory: $_"
    exit 1
}

# === Connect to Graph ===
try {
    Connect-MgGraphSecure
} catch {
    Write-Error "❌  Microsoft Graph connection failed: $_"
    exit 1
}

Write-Host "`n🛡  Starting Entra Audit for $($context.OrgName)...`n"

# === Detect sign-in capability ===
$validSkus = @("AAD_PREMIUM", "AAD_PREMIUM_P2", "ENTERPRISEPREMIUM", "EMS", "EMS_PREMIUM", "IDENTITY_GOVERNANCE")
$hasSignInLicense = ($subscribedSkus.SkuPartNumber | Where-Object { $_ -in $validSkus }).Count -gt 0

# === Pull sign-in log (if licensed) ===
$signIns = @{}
if ($hasSignInLicense) {
    try {
        $rawSignIns = Get-MgAuditLogSignIn -All
        foreach ($entry in $rawSignIns) {
            if (-not $signIns.ContainsKey($entry.UserPrincipalName)) {
                $signIns[$entry.UserPrincipalName] = $entry.CreatedDateTime
            }
        }
    } catch {
        Write-Warning "⚠ Failed to retrieve sign-ins: $_"
    }
} else {
    Write-Warning "⚠  Skipping LastSignInDateTime reporting — this tenant may not have the required license (e.g. AAD P1/P2, M365 E5, or equivalent)."
}
# ================================
# ===   License Summary
# ================================

Write-Host "`n⏳  Collecting license summary..."
$subscriptions = Get-MgSubscribedSku -All

$tenantLicenses = $subscriptions | Select-Object SkuPartNumber, SkuId, ConsumedUnits, PrepaidUnits, SubscriptionIds, CapabilityStatus, @{
    Name = "PurchaseChannel"
    Expression = { if ($_.AppliesTo -eq "User") { "Direct" } else { "Partner" } }
}

$licenseDetails = foreach ($sku in $tenantLicenses) {
    [PSCustomObject]@{
        SkuPartNumber     = $sku.SkuPartNumber
        SkuFriendlyName   = Get-FriendlySkuName $sku.SkuPartNumber
        SkuId             = $sku.SkuId
        EnabledUnits      = $sku.PrepaidUnits.Enabled
        SuspendedUnits    = $sku.PrepaidUnits.Suspended
        WarningUnits      = $sku.PrepaidUnits.Warning
        ConsumedUnits     = $sku.ConsumedUnits
        CapabilityStatus  = $sku.CapabilityStatus
        SubscriptionIds   = ($sku.SubscriptionIds -join ", ")
        PurchaseChannel   = $sku.PurchaseChannel
    }
}

$licenseDetails | Export-Csv -Path "$outputdir\Entra_Licenses.csv" -NoTypeInformation -Encoding UTF8
if (Test-Path "$outputdir\Entra_Licenses.csv") {
    Write-Host "`t📂 License Summary exported to Entra_Licenses.csv" -ForegroundColor Green
} else {
    Write-Error "`t❌ Failed to save License Summary: $_"
}


# === Self-Service Password Reset (SSPR) ===
Write-Host "`n⏳  Collecting Self-Service Password Reset (SSPR) configuration..."

try {
    $authPolicy = Get-MgPolicyAuthenticationMethodPolicy
    $ssprState = $authPolicy.RegistrationEnforcement.AuthenticationMethodsRegistrationCampaign.State

    # Map 'default' to something more user-friendly
    switch ($ssprState) {
        "enabled"  { $friendlySspr = "Enabled" }
        "disabled" { $friendlySspr = "Disabled" }
        "default"  { $friendlySspr = "Not Enforced (Default)" }
        default    { $friendlySspr = "Unknown" }
    }

    [PSCustomObject]@{
        SSPREnabled = $friendlySspr
    } | Export-Csv "$outputDir\Entra_SSPR.csv" -NoTypeInformation

    if (Test-Path "$outputDir\Entra_SSPR.csv") {
        Write-Host "`t📂 Self-Service Password Reset (SSPR) configuration exported to Entra_SSPR.csv" -ForegroundColor Green
    } else {
        Write-Error "`t❌ Failed to save Self-Service Password Reset (SSPR) configuration: $_"
    }

}
catch {
    Write-Warning "`t⚠ Unable to retrieve SSPR configuration: $_"
}


# ================================
# ===    Pull base user info   ===
# ================================
Write-Host "`n⏳  Collecting user summary..."
$users = Get-MgUser -All -Property DisplayName, GivenName, Surname, UserPrincipalName, Id, PasswordPolicies, LastPasswordChangeDateTime

# === Build SKU lookup ===
$subscribedSkus = Get-MgSubscribedSku
$skuLookup = @{}
$subscribedSkus | ForEach-Object {
    $skuLookup[$_.SkuId] = $_.SkuPartNumber
}

# === Build MFA map with method names and count (excluding password-only) ===
$mfaMap = @{}
foreach ($user in $users) {
    try {
        $methods = Get-MgUserAuthenticationMethod -UserId $user.Id
        $types = $methods | ForEach-Object { $_.AdditionalProperties['@odata.type'] }

        # Filter out password-only method
        $filteredTypes = $types | Where-Object { $_ -ne "#microsoft.graph.passwordAuthenticationMethod" }

        # Map to friendly names
        $friendlyTypes = $filteredTypes | ForEach-Object {
            switch ($_ -replace "#microsoft.graph.", "") {
                "phoneAuthenticationMethod"                   { "Phone (SMS/Call)" }
                "microsoftAuthenticatorAuthenticationMethod"  { "Authenticator App" }
                "fido2AuthenticationMethod"                   { "FIDO2 Key" }
                "windowsHelloForBusinessAuthenticationMethod" { "Windows Hello" }
                "emailAuthenticationMethod"                   { "Email" }
                "softwareOathAuthenticationMethod"            { "Software OTP"}
                default { $_ }
            }
        }

        $uniqueTypes = $friendlyTypes | Sort-Object -Unique

        $mfaMap[$user.UserPrincipalName] = @{
            Types = $uniqueTypes
            Count = $uniqueTypes.Count
        }
    }
    catch {
        Write-Warning "`t⚠ Unable to get MFA methods for $($user.UserPrincipalName): $_"
        $mfaMap[$user.UserPrincipalName] = @{
            Types = @()
            Count = 0
        }
    }
}


# === Build license assignments using LicenseDetail + friendly name fallback ===
$licenseMap = @{}
$allUsers = Get-MgUser -All -Property Id, UserPrincipalName

foreach ($user in $allUsers) {
    $licenses = @()
    try {
        $licenseDetails = Get-MgUserLicenseDetail -UserId $user.Id
        foreach ($detail in $licenseDetails) {
            $skuPartNumber = $detail.SkuPartNumber
            $licenses += Get-FriendlySkuName $skuPartNumber
        }
    }
    catch {
        if ($debugMode) {
            Write-Warning "[DEBUG] Failed to retrieve license details for $($user.UserPrincipalName): $_"
        }
    }
    if ($user.UserPrincipalName) {
        $licenseMap[$user.UserPrincipalName] = $licenses -join ", "
    }
}

# === Generate combined user report ===
$userReport = @()
foreach ($user in $users) {
    $upn = $user.UserPrincipalName
    $first = $user.GivenName
    $last = $user.Surname
    $mfaEnabled = $mfaMap.ContainsKey($upn) -and ($mfaMap[$upn].Count -gt 0)
    $mfaTypes = if ($mfaMap[$upn].Count -gt 0) { $mfaMap[$upn].Types -join ", " } else { "None" }
    $mfaCount = $mfaMap[$upn].Count
    $license = if ($licenseMap.ContainsKey($upn)) { $licenseMap[$upn] } else { "None" }
    $pwdExpiry = if ($user.PasswordPolicies -notmatch "DisablePasswordExpiration") { "Enabled" } else { "Disabled" }
    $pwdLastSet = $user.LastPasswordChangeDateTime
    $lastSignIn = if ($hasSignInLicense -and $signIns.ContainsKey($upn)) {
        $signIns[$upn]
    } else {
        "Unavailable"
    }

    $userReport += [PSCustomObject]@{
        UPN               = $upn
        FirstName         = $first
        LastName          = $last
        AssignedLicense   = $license
        MFAEnabled        = $mfaEnabled
        MFAMethods        = $mfaTypes
        MFACount          = $mfaCount
        DisablePasswordExpiration    = $pwdExpiry
        LastPasswordChange= $pwdLastSet
    }

    if ($hasSignInLicense -and $signIns.ContainsKey($upn)) {
        $userReport['LastSignIn'] = $signIns[$upn]
    }
}

$userReport | Export-Csv "$outputDir\Entra_Users.csv" -NoTypeInformation
if (Test-Path "$outputDir\Entra_Users.csv") {
    Write-Host "`t✅ User summary exported to Entra_Users.csv" -ForegroundColor Green
} else {
    Write-Error "`t❌ Failed to save User summary: $_"
}

# === Admin Role Assignments ===
Write-Host "`n⏳  Collecting Admin Role assignments..."

$roles = Get-MgDirectoryRole
$adminReport = @()
$globalAdmins = @()

foreach ($role in $roles) {
    try {
        $members = Get-MgDirectoryRoleMember -DirectoryRoleId $role.Id

        foreach ($member in $members) {
            $entry = [PSCustomObject]@{
                RoleName              = $role.DisplayName
                MemberDisplayName     = $member.AdditionalProperties.displayName
                MemberUserPrincipalName = $member.AdditionalProperties.userPrincipalName
            }

            $adminReport += $entry

            # Comment
            if ($role.DisplayName -eq "Global Administrator") {
                $globalAdmins += $entry
            }
        }
    } catch {
        Write-Warning "⚠ Could not retrieve members for role: $($role.DisplayName)"
    }
}

# === Export full role assignment list ===
$adminReport | Export-Csv "$outputDir\Entra_AdminRoles.csv" -NoTypeInformation
if (Test-Path "$outputDir\Entra_AdminRoles.csv") {
    Write-Host "`t📂 Admin Role assignments exported to Entra_AdminRoles.csv" -ForegroundColor Green
} else {
    Write-Error "`t❌ Failed to save Admin Role assignments: $_"
}

$globalAdmins | Export-Csv "$outputDir\Entra_GlobalAdmins.csv" -NoTypeInformation
if (Test-Path "$outputDir\Entra_GlobalAdmins.csv") {
    Write-Host "`t📂 Global Admin summary exported to Entra_GlobalAdmins.csv" -ForegroundColor Green
} else {
    Write-Error "`t❌ Failed to save Global Admin summary: $_"
}

# === Check Global Admin count ===
switch ($globalAdmins.Count) {
    1 {
        Write-Host "`n⚠  Only ONE Global Administrator found in the tenant!" -ForegroundColor Yellow
        Write-Host "   Best practice is to have at least TWO to avoid lockout risks.`n" -ForegroundColor Yellow
    }
    default {
        Write-Host "✅ Global Administrators: $($globalAdmins.Count)"
    }
}

# === Guest Users ===
Write-Host "`n⏳  Collecting Guest user summary..."
$guestUsers = Get-MgUser -Filter "UserType eq 'Guest'" -All
$guestUsers | Select DisplayName, UserPrincipalName, CreatedDateTime | Export-Csv "$outputDir\Entra_GuestUsers.csv" -NoTypeInformation

if (Test-Path "$outputDir\Entra_GuestUsers.csv") {
    Write-Host "`t📂 Guest user summary exported to Entra_GuestUsers.csv" -ForegroundColor Green
} else {
    Write-Error "`t❌ Failed to save Guest user summary: $_"
}






# ================================
# Group Members and Owners
# ================================

Write-Host "`n⏳  Collecting Groups summary..."
$groups = Get-MgGroup -All

$groupData = foreach ($group in $groups) {
    $owners = (Get-MgGroupOwner -GroupId $group.Id -ErrorAction SilentlyContinue | ForEach-Object {
        $_.AdditionalProperties['userPrincipalName']
    }) -join "; "
    
    $members = (Get-MgGroupMember -GroupId $group.Id -ErrorAction SilentlyContinue | ForEach-Object {
        $_.AdditionalProperties['userPrincipalName']
    }) -join "; "
    

    [PSCustomObject]@{
        DisplayName         = $group.DisplayName
        GroupId             = $group.Id
        GroupType           = if ($group.GroupTypes -and $group.GroupTypes -contains "Unified") { "Microsoft 365" } else { "Security" }
        MembershipType      = if ($group.MembershipRule) { "Dynamic" } else { "Assigned" }
        Owners              = $owners
        Members             = $members
        MailEnabled         = $group.MailEnabled
        SecurityEnabled     = $group.SecurityEnabled
        IsAssignableToRole  = $group.IsAssignableToRole
        Visibility          = $group.Visibility
        OnPremSyncEnabled   = $group.OnPremisesSyncEnabled
    }
    
}
$groupData | Export-Csv -Path "$outputdir\Entra_Groups.csv" -NoTypeInformation -Encoding UTF8
if (Test-Path "$outputdir\Entra_Groups.csv" ) {
    Write-Host "`t📂 Groups summary exported to Entra_SSPR.csv" -ForegroundColor Green
} else {
    Write-Error "`t❌ Failed to save Self-Service Password Reset (SSPR) configuration: $_"
}



# ================================
# === Done                     ===
# ================================
Write-Host "`n✅ Entra Audit Complete. Results saved to: $outputDir`n"
