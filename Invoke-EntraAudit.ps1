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

.NOTES
    Author      : Raymond Slater
    Version     : 1.10.2
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

$ScriptVersion = "1.10.2"
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
    $outputDir = $context.OutputPath
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

Write-Host "`nStarting Entra Audit for $($context.OrgName)..." -ForegroundColor Cyan

$step       = 0
$totalSteps = 12
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
    [PSCustomObject]@{
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
        Owners             = $owners
        Members            = $members
        MailEnabled        = $group.MailEnabled
        SecurityEnabled    = $group.SecurityEnabled
        IsAssignableToRole = $group.IsAssignableToRole
        Visibility         = $group.Visibility
        OnPremSyncEnabled  = $group.OnPremisesSyncEnabled
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
    $caPolicyData = foreach ($policy in (Get-MgIdentityConditionalAccessPolicy -All)) {
        [PSCustomObject]@{
            Name           = $policy.DisplayName
            State          = $policy.State
            IncludeUsers   = ($policy.Conditions.Users.IncludeUsers -join ", ")
            ExcludeUsers   = ($policy.Conditions.Users.ExcludeUsers -join ", ")
            IncludeGroups  = ($policy.Conditions.Users.IncludeGroups -join ", ")
            GrantControls  = ($policy.GrantControls.BuiltInControls -join ", ")
            RequiresMFA    = ($policy.GrantControls.BuiltInControls -contains "mfa")
            ClientAppTypes = ($policy.Conditions.ClientAppTypes -join ", ")
        }
    }

    $caPolicyData | Export-Csv "$outputDir\Entra_CA_Policies.csv" -NoTypeInformation -Encoding UTF8
    Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Collecting Conditional Access policies..." -CurrentOperation "Saved: Entra_CA_Policies.csv" -PercentComplete ([int]($step / $totalSteps * 100))
}
catch {
    Write-Warning "Unable to retrieve Conditional Access policies: $_"
}


# ================================
# ===   Named / Trusted Locations
# ================================
$step++
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Collecting named locations..." -PercentComplete ([int]($step / $totalSteps * 100))

try {
    $locationData = foreach ($loc in (Get-MgIdentityConditionalAccessNamedLocation -All)) {
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
    Write-Warning "Unable to retrieve Security Defaults: $_"
}


# ================================
# ===   Done                    ===
# ================================
Write-Progress -Id 1 -Activity $activity -Completed
Write-Host "`nEntra Audit complete. Results saved to: $outputDir`n" -ForegroundColor Green
