function Get-EntraInventory {
    <#
    .SYNOPSIS
        Collects Entra ID inventory data and exports to CSV files.

    .DESCRIPTION
        Queries Microsoft Graph for identity-related data including users, licenses,
        admin roles, groups, conditional access policies, sign-in logs, enterprise apps,
        secure score, and more. Exports raw data to CSV files in the Raw/ output directory.

        Does NOT perform security evaluation — that is delegated to Maester.

    .OUTPUTS
        Hashtable with summary counts for the orchestrator.
    #>
    [CmdletBinding()]
    param()

    $ctx       = Get-AuditContext
    $outputDir = $ctx.RawOutputPath

    # ── Import Graph sub-modules for Entra queries ──────────────────────────
    Import-GraphModule -ModuleName 'Microsoft.Graph.Users'
    Import-GraphModule -ModuleName 'Microsoft.Graph.Groups'
    Import-GraphModule -ModuleName 'Microsoft.Graph.Reports'
    Import-GraphModule -ModuleName 'Microsoft.Graph.Identity.SignIns'
    Import-GraphModule -ModuleName 'Microsoft.Graph.Applications'

    Write-Host "`nStarting Entra Inventory for $($ctx.OrgName)..." -ForegroundColor Cyan

    $step       = 0
    $totalSteps = 20
    $activity   = "Entra Inventory — $($ctx.OrgName)"

    # ── Helper: CA label conversion functions (internal to this collector) ───
    # These are needed for flattening Conditional Access policies to CSV.
    # Kept inline since they're only used by the CA export section.

    $_caUserCache   = @{}
    $_caGroupCache  = @{}
    $_caSpCache     = @{}
    $_caRoleCache   = @{}
    $_caRoleCacheOk = $false
    $_caLocCache    = @{}
    $_caLocCacheOk  = $false
    $_caLocations   = @()

    function _CaTokenLabel([string]$Kind, [string]$Value) {
        switch ($Kind) {
            'User' { switch ($Value) { 'All' { 'All users' } 'None' { 'None' } 'GuestsOrExternalUsers' { 'All guest and external users' } default { $null } } }
            'Group' { switch ($Value) { 'All' { 'All groups' } 'None' { 'None' } default { $null } } }
            'Role' { switch ($Value) { 'All' { 'All roles' } 'None' { 'None' } default { $null } } }
            'Application' { switch ($Value) { 'All' { 'All cloud apps' } 'Office365' { 'Office 365' } 'MicrosoftAdminPortals' { 'Microsoft Admin Portals' } 'None' { 'None' } default { $null } } }
            'Location' { switch ($Value) { 'All' { 'All locations' } 'AllTrusted' { 'All trusted locations' } 'None' { 'None' } default { $null } } }
        }
    }

    function _CaControlLabel([string]$Control) {
        switch ($Control) { 'mfa' { 'Require MFA' } 'block' { 'Block access' } 'compliantDevice' { 'Require compliant device' } 'domainJoinedDevice' { 'Require Entra hybrid joined device' } 'approvedApplication' { 'Require approved client app' } 'compliantApplication' { 'Require app protection policy' } 'passwordChange' { 'Require password change' } default { $Control } }
    }

    function _CaClientAppLabel([string]$Type) {
        switch ($Type) { 'all' { 'All client apps' } 'browser' { 'Browser' } 'mobileAppsAndDesktopClients' { 'Mobile apps and desktop clients' } 'exchangeActiveSync' { 'Exchange ActiveSync' } 'easSupported' { 'Exchange ActiveSync clients' } 'other' { 'Other clients' } default { $Type } }
    }

    function _CaPlatformLabel([string]$P) {
        switch ($P) { 'all' { 'All platforms' } 'android' { 'Android' } 'iOS' { 'iOS' } 'windows' { 'Windows' } 'macOS' { 'macOS' } 'linux' { 'Linux' } default { $P } }
    }

    function _CaRiskLabel([string]$R) {
        switch ($R) { 'low' { 'Low' } 'medium' { 'Medium' } 'high' { 'High' } 'hidden' { 'Hidden' } 'none' { 'None' } default { $R } }
    }

    function _CaUserActionLabel([string]$Action) {
        switch ($Action) { 'urn:user:registersecurityinfo' { 'Register security information' } default { $Action } }
    }

    function _CaGrantOperatorLabel([string]$Operator) {
        switch ($Operator) { 'AND' { 'All selected controls required' } 'OR' { 'One of the selected controls required' } default { $Operator } }
    }

    function _CaResolveId([string]$Kind, [string]$Id) {
        if ([string]::IsNullOrWhiteSpace($Id)) { return $null }
        $wk = _CaTokenLabel $Kind $Id; if ($wk) { return $wk }
        switch ($Kind) {
            'User'        { if ($_caUserCache.ContainsKey($Id)) { return $_caUserCache[$Id] }; try { $u = Get-MgUser -UserId $Id -Property DisplayName,UserPrincipalName -EA Stop; $r = if ($u.DisplayName -and $u.UserPrincipalName -and $u.DisplayName -ne $u.UserPrincipalName) { "$($u.DisplayName) ($($u.UserPrincipalName))" } elseif ($u.UserPrincipalName) { $u.UserPrincipalName } else { $Id }; Set-Variable -Scope 1 -Name _caUserCache -Value (($_caUserCache.Clone()).GetEnumerator() | ForEach-Object { $_ }) } catch { $r = $Id }; $_caUserCache[$Id] = $r; return $r }
            'Group'       { if ($_caGroupCache.ContainsKey($Id)) { return $_caGroupCache[$Id] }; try { $g = Get-MgGroup -GroupId $Id -Property DisplayName -EA Stop; $r = if ($g.DisplayName) { $g.DisplayName } else { $Id } } catch { $r = $Id }; $_caGroupCache[$Id] = $r; return $r }
            'Role'        { if (-not $_caRoleCacheOk) { try { Get-MgDirectoryRoleTemplate -All -EA Stop | ForEach-Object { if ($_.Id) { $_caRoleCache[$_.Id] = $_.DisplayName } }; Get-MgDirectoryRole -All -EA Stop | ForEach-Object { if ($_.Id) { $_caRoleCache[$_.Id] = $_.DisplayName }; if ($_.RoleTemplateId) { $_caRoleCache[$_.RoleTemplateId] = $_.DisplayName } } } catch {}; Set-Variable -Scope 0 -Name _caRoleCacheOk -Value $true }; if ($_caRoleCache.ContainsKey($Id)) { return $_caRoleCache[$Id] }; return $Id }
            'Application' { if ($_caSpCache.ContainsKey($Id)) { return $_caSpCache[$Id] }; try { $sp = Get-MgServicePrincipal -ServicePrincipalId $Id -Property DisplayName,AppId -EA Stop; $r = if ($sp.DisplayName -and $sp.AppId) { "$($sp.DisplayName) ($($sp.AppId))" } elseif ($sp.DisplayName) { $sp.DisplayName } else { $Id } } catch { $r = $Id }; $_caSpCache[$Id] = $r; return $r }
            'Location'    { if (-not $_caLocCacheOk) { try { $_caLocations = @(Get-MgIdentityConditionalAccessNamedLocation -All -EA Stop); $_caLocations | ForEach-Object { if ($_.Id) { $_caLocCache[$_.Id] = $_.DisplayName } } } catch {}; Set-Variable -Scope 0 -Name _caLocCacheOk -Value $true }; if ($_caLocCache.ContainsKey($Id)) { return $_caLocCache[$Id] }; return $Id }
        }
    }

    function _CaResolveList([object[]]$Values, [string]$Kind) {
        @($Values | ForEach-Object { if ($_) { _CaResolveId $Kind ([string]$_) } } | Where-Object { $_ } | Select-Object -Unique)
    }

    function _JoinValues([object[]]$V, [string]$D = '—') {
        $items = @($V | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | Select-Object -Unique)
        if ($items.Count -eq 0) { return $D }; return ($items -join '; ')
    }

    function _FormatUserLabel([string]$DN, [string]$UPN, [string]$FB) {
        if ($DN -and $UPN -and $DN -ne $UPN) { return "$DN ($UPN)" }
        if ($UPN) { return $UPN }; if ($DN) { return $DN }; return $FB
    }

    # ════════════════════════════════════════════════════════════════════════
    # DATA COLLECTION — each section fetches data and exports to CSV
    # ════════════════════════════════════════════════════════════════════════

    # ── Subscribed SKUs (used by multiple sections) ─────────────────────────
    $subscribedSkus = Get-MgSubscribedSku -All
    $premiumSignInSkus = @("AAD_PREMIUM", "AAD_PREMIUM_P2", "ENTERPRISEPREMIUM", "ENTERPRISEPACK",
                           "EMS", "EMS_PREMIUM", "SPB", "O365_BUSINESS_PREMIUM", "M365_F3", "IDENTITY_GOVERNANCE")
    $retentionDays = if (($subscribedSkus.SkuPartNumber | Where-Object { $_ -in $premiumSignInSkus }).Count -gt 0) { 30 } else { 7 }

    # P2 licence detection (used by Identity Protection + PIM sections)
    $_p2Skus = @(
        'AAD_PREMIUM_P2', 'EMS_E5', 'EMSPREMIUM', 'SPE_E5', 'SPE_E5_USGOV_GCCHIGH',
        'M365EDU_A5_FACULTY', 'M365EDU_A5_STUDENT', 'IDENTITY_THREAT_PROTECTION',
        'IDENTITY_THREAT_PROTECTION_FOR_SMB'
    )
    $_hasP2 = ($subscribedSkus.SkuPartNumber | Where-Object { $_ -in $_p2Skus }).Count -gt 0

    # ── 1. Sign-in Logs ─────────────────────────────────────────────────────
    $step++; Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Sign-in logs..." -PercentComplete ([int]($step / $totalSteps * 100))

    $signIns   = @{}
    $signInMap = @{}
    try {
        $rawSignIns = Get-MgAuditLogSignIn -All -ErrorAction Stop
        foreach ($entry in $rawSignIns) {
            $upn = $entry.UserPrincipalName; if (-not $upn) { continue }
            if (-not $signIns.ContainsKey($upn)) { $signIns[$upn] = $entry.CreatedDateTime }
            if (-not $signInMap.ContainsKey($upn)) { $signInMap[$upn] = [System.Collections.Generic.List[object]]::new() }
            if ($signInMap[$upn].Count -lt 10) {
                $loc = $entry.Location
                $signInMap[$upn].Add([PSCustomObject]@{
                    UPN = $upn; Timestamp = $entry.CreatedDateTime.ToUniversalTime().ToString("yyyy-MM-dd HH:mm") + " UTC"
                    App = $entry.AppDisplayName; IPAddress = $entry.IpAddress; City = $loc.City; Country = $loc.CountryOrRegion
                    Success = ($entry.Status.ErrorCode -eq 0); FailureReason = if ($entry.Status.ErrorCode -ne 0) { $entry.Status.FailureReason } else { "" }
                })
            }
        }
        ($signInMap.Values | ForEach-Object { $_ }) | Export-Csv (Join-Path $outputDir 'Entra_SignIns.csv') -NoTypeInformation -Encoding UTF8
    }
    catch {
        Add-AuditIssue -Severity 'Warning' -Section 'Entra' -Collector 'Sign-in Logs' -Description $_.Exception.Message
        Write-Warning "Failed to retrieve sign-in logs: $_"
    }

    # ── 2. Directory Audit Events ───────────────────────────────────────────
    $step++; Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Audit events..." -PercentComplete ([int]($step / $totalSteps * 100))

    $auditFrom  = (Get-Date).AddDays(-$retentionDays).ToString("yyyy-MM-ddTHH:mm:ssZ")
    $dateFilter = "activityDateTime ge $auditFrom"

    function _AuditInitiator($Entry) {
        if ($Entry.InitiatedBy.User.UserPrincipalName) { return $Entry.InitiatedBy.User.UserPrincipalName }
        if ($Entry.InitiatedBy.App.DisplayName)        { return "$($Entry.InitiatedBy.App.DisplayName) [app]" }
        return "System"
    }

    # Account Creations
    try {
        $rawCreations = Get-MgAuditLogDirectoryAudit -Filter "$dateFilter and activityDisplayName eq 'Add user' and result eq 'success'" -All -ErrorAction Stop
        $acctCreations = foreach ($entry in $rawCreations) {
            $target = $entry.TargetResources | Select-Object -First 1
            $nameProp = $target.ModifiedProperties | Where-Object { $_.DisplayName -eq 'displayName' } | Select-Object -First 1
            [PSCustomObject]@{
                Timestamp = $entry.ActivityDateTime.ToUniversalTime().ToString("yyyy-MM-dd HH:mm") + " UTC"
                InitiatedBy = _AuditInitiator $entry; TargetUPN = $target.UserPrincipalName
                TargetName = if ($nameProp -and $nameProp.NewValue) { $nameProp.NewValue } elseif ($target.DisplayName) { $target.DisplayName } else { '' }
            }
        }
        $acctCreations | Export-Csv (Join-Path $outputDir 'Entra_AccountCreations.csv') -NoTypeInformation -Encoding UTF8
    }
    catch { Add-AuditIssue -Severity 'Warning' -Section 'Entra' -Collector 'Account Creations' -Description $_.Exception.Message; Write-Warning "Failed to retrieve account creations: $_" }

    # Account Deletions
    try {
        $rawDeletions = Get-MgAuditLogDirectoryAudit -Filter "$dateFilter and activityDisplayName eq 'Delete user' and result eq 'success'" -All -ErrorAction Stop
        $acctDeletions = foreach ($entry in $rawDeletions) {
            $target = $entry.TargetResources | Select-Object -First 1
            [PSCustomObject]@{
                Timestamp = $entry.ActivityDateTime.ToUniversalTime().ToString("yyyy-MM-dd HH:mm") + " UTC"
                InitiatedBy = _AuditInitiator $entry; TargetUPN = $target.UserPrincipalName
                TargetName = if ($target.DisplayName) { $target.DisplayName } else { '' }
            }
        }
        $acctDeletions | Export-Csv (Join-Path $outputDir 'Entra_AccountDeletions.csv') -NoTypeInformation -Encoding UTF8
    }
    catch { Add-AuditIssue -Severity 'Warning' -Section 'Entra' -Collector 'Account Deletions' -Description $_.Exception.Message; Write-Warning "Failed to retrieve account deletions: $_" }

    # Notable Audit Events (role changes + security info changes)
    $securityActivityNames = @("Reset user password", "User registered security info", "User deleted security info",
        "User changed default security info", "Admin registered security info for a user", "Admin deleted security info for a user", "Admin updated security info for a user")
    try {
        $roleEvents     = @(Get-MgAuditLogDirectoryAudit -Filter "$dateFilter and category eq 'RoleManagement'" -All -ErrorAction Stop)
        $rawUserMgmt    = @(Get-MgAuditLogDirectoryAudit -Filter "$dateFilter and category eq 'UserManagement'" -All -ErrorAction Stop)
        $securityEvents = @($rawUserMgmt | Where-Object { $_.ActivityDisplayName -in $securityActivityNames })
        $auditEvents = foreach ($entry in ($roleEvents + $securityEvents)) {
            $targetUser = $entry.TargetResources | Where-Object { $_.Type -eq 'User' } | Select-Object -First 1
            $targetRole = $entry.TargetResources | Where-Object { $_.Type -match '(?i)role' } | Select-Object -First 1
            if (-not $targetUser) { $targetUser = $entry.TargetResources | Select-Object -First 1 }
            [PSCustomObject]@{
                Timestamp = $entry.ActivityDateTime.ToUniversalTime().ToString("yyyy-MM-dd HH:mm") + " UTC"
                Category = $entry.Category; Activity = $entry.ActivityDisplayName; InitiatedBy = _AuditInitiator $entry
                TargetUPN = _FormatUserLabel $targetUser.DisplayName $targetUser.UserPrincipalName ''
                TargetName = $targetUser.DisplayName; TargetRole = if ($targetRole) { $targetRole.DisplayName } else { '' }; Result = $entry.Result
            }
        }
        $auditEvents | Sort-Object Timestamp -Descending | Export-Csv (Join-Path $outputDir 'Entra_AuditEvents.csv') -NoTypeInformation -Encoding UTF8
    }
    catch { Add-AuditIssue -Severity 'Warning' -Section 'Entra' -Collector 'Audit Events' -Description $_.Exception.Message; Write-Warning "Failed to retrieve audit events: $_" }

    # ── 3. Licence Summary ──────────────────────────────────────────────────
    $step++; Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Licences..." -PercentComplete ([int]($step / $totalSteps * 100))

    $licenseDetails = foreach ($sku in $subscribedSkus) {
        [PSCustomObject]@{
            SkuPartNumber = $sku.SkuPartNumber; SkuFriendlyName = Get-FriendlySkuName $sku.SkuPartNumber
            SkuId = $sku.SkuId; EnabledUnits = $sku.PrepaidUnits.Enabled; SuspendedUnits = $sku.PrepaidUnits.Suspended
            WarningUnits = $sku.PrepaidUnits.Warning; ConsumedUnits = $sku.ConsumedUnits
            CapabilityStatus = $sku.CapabilityStatus; SubscriptionIds = ($sku.SubscriptionIds -join ", ")
        }
    }
    $licenseDetails | Export-Csv (Join-Path $outputDir 'Entra_Licenses.csv') -NoTypeInformation -Encoding UTF8

    # ── 4. SSPR Configuration ───────────────────────────────────────────────
    $step++; Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — SSPR configuration..." -PercentComplete ([int]($step / $totalSteps * 100))

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
            Export-Csv (Join-Path $outputDir 'Entra_SSPR.csv') -NoTypeInformation -Encoding UTF8
    }
    catch {
        Add-AuditIssue -Severity 'Warning' -Section 'Entra' -Collector 'Get-MgPolicyAuthenticationMethodPolicy' -Description $_.Exception.Message
        Write-Warning "Unable to retrieve SSPR configuration: $_"
    }

    # ── 5. User Summary ─────────────────────────────────────────────────────
    $step++; Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — User summary..." -PercentComplete ([int]($step / $totalSteps * 100))

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

    # Populate CA user cache with member users
    foreach ($user in $users) {
        if ($user.Id) {
            $_caUserCache[$user.Id] = _FormatUserLabel $user.DisplayName $user.UserPrincipalName $user.Id
        }
    }

    $licensedUsers   = @($userReport | Where-Object { $_.AssignedLicense -ne "None" })
    $unlicensedUsers = @($userReport | Where-Object { $_.AssignedLicense -eq "None" })

    $licensedUsers   | Export-Csv (Join-Path $outputDir 'Entra_Users.csv')            -NoTypeInformation -Encoding UTF8
    $unlicensedUsers | Export-Csv (Join-Path $outputDir 'Entra_Users_Unlicensed.csv') -NoTypeInformation -Encoding UTF8

    # ── 6. Admin Role Assignments ───────────────────────────────────────────
    $step++; Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Admin role assignments..." -PercentComplete ([int]($step / $totalSteps * 100))

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

    $adminReport  | Export-Csv (Join-Path $outputDir 'Entra_AdminRoles.csv')   -NoTypeInformation -Encoding UTF8
    $globalAdmins | Export-Csv (Join-Path $outputDir 'Entra_GlobalAdmins.csv') -NoTypeInformation -Encoding UTF8

    if ($globalAdmins.Count -eq 1) {
        Write-Warning "Only ONE Global Administrator found. Best practice is at least two to avoid lockout."
    }

    # ── 7. Guest Users ──────────────────────────────────────────────────────
    $step++; Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Guest users..." -PercentComplete ([int]($step / $totalSteps * 100))

    $guestData = foreach ($guest in (Get-MgUser -Filter "UserType eq 'Guest'" -All -Property Id,DisplayName,UserPrincipalName,CreatedDateTime,SignInActivity -ErrorAction SilentlyContinue)) {
        if ($guest.Id) {
            $_caUserCache[$guest.Id] = _FormatUserLabel $guest.DisplayName $guest.UserPrincipalName $guest.Id
        }

        [PSCustomObject]@{
            UserId            = $guest.Id
            DisplayName       = $guest.DisplayName
            UserPrincipalName = $guest.UserPrincipalName
            CreatedDateTime   = $guest.CreatedDateTime
            LastSignIn        = if ($guest.SignInActivity.LastSignInDateTime) { $guest.SignInActivity.LastSignInDateTime.ToString("yyyy-MM-dd HH:mm") + " UTC" } else { $null }
        }
    }
    $guestData | Export-Csv (Join-Path $outputDir 'Entra_GuestUsers.csv') -NoTypeInformation -Encoding UTF8

    # ── 8. Groups ───────────────────────────────────────────────────────────
    $step++; Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Groups..." -PercentComplete ([int]($step / $totalSteps * 100))

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

    # Populate CA group cache
    foreach ($group in $groupData) {
        if ($group.GroupId -and $group.DisplayName) {
            $_caGroupCache[$group.GroupId] = $group.DisplayName
        }
    }

    $groupData | Export-Csv (Join-Path $outputDir 'Entra_Groups.csv') -NoTypeInformation -Encoding UTF8

    # ── 9. Conditional Access Policies ──────────────────────────────────────
    $step++; Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Conditional Access policies..." -PercentComplete ([int]($step / $totalSteps * 100))

    try {
        # Pre-populate named location cache for CA resolution
        if (-not $_caLocCacheOk) {
            try {
                $_caLocations = @(Get-MgIdentityConditionalAccessNamedLocation -All -EA Stop)
                $_caLocations | ForEach-Object { if ($_.Id) { $_caLocCache[$_.Id] = $_.DisplayName } }
                $_caLocCacheOk = $true
            } catch {}
        }

        $caPolicyData = foreach ($policy in (Get-MgIdentityConditionalAccessPolicy -All)) {
            $userConditions         = $policy.Conditions.Users
            $applicationConditions  = $policy.Conditions.Applications
            $platformConditions     = $policy.Conditions.Platforms
            $locationConditions     = $policy.Conditions.Locations
            $deviceConditions       = $policy.Conditions.Devices
            $deviceFilter           = if ($deviceConditions) { $deviceConditions.DeviceFilter } else { $null }
            $grantControls          = $policy.GrantControls

            $includeUsers           = _JoinValues (_CaResolveList $userConditions.IncludeUsers 'User')
            $excludeUsers           = _JoinValues (_CaResolveList $userConditions.ExcludeUsers 'User')
            $includeGroups          = _JoinValues (_CaResolveList $userConditions.IncludeGroups 'Group')
            $excludeGroups          = _JoinValues (_CaResolveList $userConditions.ExcludeGroups 'Group')
            $includeRoles           = _JoinValues (_CaResolveList $userConditions.IncludeRoles 'Role')
            $excludeRoles           = _JoinValues (_CaResolveList $userConditions.ExcludeRoles 'Role')
            $includeApplications    = _JoinValues (_CaResolveList $applicationConditions.IncludeApplications 'Application')
            $excludeApplications    = _JoinValues (_CaResolveList $applicationConditions.ExcludeApplications 'Application')
            $includeUserActions     = _JoinValues (@($applicationConditions.IncludeUserActions | ForEach-Object { _CaUserActionLabel $_ }))
            $grantControlsLabel     = _JoinValues (@($grantControls.BuiltInControls | ForEach-Object { _CaControlLabel $_ }))
            $grantOperatorLabel     = if ($grantControls.Operator) { _CaGrantOperatorLabel $grantControls.Operator } else { '—' }
            $clientAppTypes         = _JoinValues (@($policy.Conditions.ClientAppTypes | ForEach-Object { _CaClientAppLabel $_ })) -D 'All client apps'
            $includePlatforms       = _JoinValues (@($platformConditions.IncludePlatforms | ForEach-Object { _CaPlatformLabel $_ }))
            $excludePlatforms       = _JoinValues (@($platformConditions.ExcludePlatforms | ForEach-Object { _CaPlatformLabel $_ }))
            $includeLocations       = _JoinValues (_CaResolveList $locationConditions.IncludeLocations 'Location')
            $excludeLocations       = _JoinValues (_CaResolveList $locationConditions.ExcludeLocations 'Location')
            $signInRiskLevels       = _JoinValues (@($policy.Conditions.SignInRiskLevels | ForEach-Object { _CaRiskLabel $_ }))
            $userRiskLevels         = _JoinValues (@($policy.Conditions.UserRiskLevels | ForEach-Object { _CaRiskLabel $_ }))
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

        $caPolicyData | Export-Csv (Join-Path $outputDir 'Entra_CA_Policies.csv') -NoTypeInformation -Encoding UTF8
    }
    catch {
        Add-AuditIssue -Severity 'Warning' -Section 'Entra' -Collector 'Get-MgIdentityConditionalAccessPolicy' -Description $_.Exception.Message
        Write-Warning "Unable to retrieve Conditional Access policies: $_"
    }

    # ── 10. Named / Trusted Locations ───────────────────────────────────────
    $step++; Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Named locations..." -PercentComplete ([int]($step / $totalSteps * 100))

    try {
        # Re-use cached locations if available from CA section
        $namedLocations = if ($_caLocations -and $_caLocations.Count -gt 0) {
            $_caLocations
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

        $locationData | Export-Csv (Join-Path $outputDir 'Entra_TrustedLocations.csv') -NoTypeInformation -Encoding UTF8
    }
    catch {
        Add-AuditIssue -Severity 'Warning' -Section 'Entra' -Collector 'Get-MgIdentityConditionalAccessNamedLocation' -Description $_.Exception.Message
        Write-Warning "Unable to retrieve named locations: $_"
    }

    # ── 11. Identity Secure Score ───────────────────────────────────────────
    $step++; Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Secure Score..." -PercentComplete ([int]($step / $totalSteps * 100))

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
            } | Export-Csv (Join-Path $outputDir 'Entra_SecureScore.csv') -NoTypeInformation -Encoding UTF8

            # Fetch human-readable titles from control profiles (paginated)
            $profileTitles = @{}
            $profileUri = 'https://graph.microsoft.com/v1.0/security/secureScoreControlProfiles?$select=controlName,title&$top=250'
            while ($profileUri) {
                $profilePage = Invoke-MgGraphRequest -Method GET -Uri $profileUri -OutputType PSObject -ErrorAction Stop
                foreach ($p in $profilePage.value) {
                    if ($p.controlName -and $p.title) { $profileTitles[$p.controlName] = $p.title }
                }
                $profileUri = $profilePage.'@odata.nextLink'
            }

            # Fallback: convert raw API key to a readable label
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
                Export-Csv (Join-Path $outputDir 'Entra_SecureScoreControls.csv') -NoTypeInformation -Encoding UTF8
        }
    }
    catch {
        Add-AuditIssue -Severity 'Warning' -Section 'Entra' -Collector 'Invoke-MgGraphRequest' -Description $_.Exception.Message
        if ($_.Exception.Message -match '403|Forbidden|valid permissions|valid roles') {
            Write-Warning "Secure Score: permission denied (SecurityEvents.Read.All not yet granted). Re-run Setup-365AuditApp.ps1 to add the missing permission."
        }
        else {
            Write-Warning "Unable to retrieve Secure Score: $($_.Exception.Message)"
        }
        Write-Verbose "Secure Score full error: $_"
    }

    # ── 12. Security Defaults ───────────────────────────────────────────────
    $step++; Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Security Defaults..." -PercentComplete ([int]($step / $totalSteps * 100))

    try {
        $secDefaults = Get-MgPolicyIdentitySecurityDefaultEnforcementPolicy

        [PSCustomObject]@{
            SecurityDefaultsEnabled = $secDefaults.IsEnabled
        } | Export-Csv (Join-Path $outputDir 'Entra_SecurityDefaults.csv') -NoTypeInformation -Encoding UTF8
    }
    catch {
        Add-AuditIssue -Severity 'Warning' -Section 'Entra' -Collector 'Get-MgPolicyIdentitySecurityDefaultEnforcementPolicy' -Description $_.Exception.Message
        Write-Warning "Unable to retrieve Security Defaults: $_"
    }

    # ── 13. Enterprise Applications + Permissions ───────────────────────────
    $step++; Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Enterprise apps..." -PercentComplete ([int]($step / $totalSteps * 100))

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

        # Cache service principal details (AppRoles + Scopes) to resolve permission names
        $_spPermCache = @{}
        function _GetCachedSP {
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
                $_resSP   = _GetCachedSP -SpId $ra.ResourceId
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
                        $_resSP2 = _GetCachedSP -SpId $grant.ResourceId
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
            $appData | Export-Csv (Join-Path $outputDir 'Entra_EnterpriseApps.csv') -NoTypeInformation -Encoding UTF8
        }
        if ($appPermData.Count -gt 0) {
            $appPermData | Export-Csv (Join-Path $outputDir 'Entra_EnterpriseAppPermissions.csv') -NoTypeInformation -Encoding UTF8
        }
    }
    catch {
        Add-AuditIssue -Severity 'Warning' -Section 'Entra' -Collector 'Get-MgServicePrincipal' -Description $_.Exception.Message
        Write-Warning "Unable to retrieve enterprise applications: $_"
    }

    # ── 14. Identity Protection (P2 required) ──────────────────────────────
    $step++; Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Identity Protection..." -PercentComplete ([int]($step / $totalSteps * 100))

    if (-not $_hasP2) {
        Write-Verbose "No Azure AD Premium P2 licence detected — skipping Identity Protection collection."
    }
    else {
        # Risky Users
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
                } | Export-Csv (Join-Path $outputDir 'Entra_RiskyUsers.csv') -NoTypeInformation -Encoding UTF8
            }
            else {
                Write-Verbose "No risky users detected — skipping Entra_RiskyUsers.csv"
            }
        }
        catch {
            Add-AuditIssue -Severity 'Warning' -Section 'Entra' -Collector 'Invoke-MgGraphRequest' -Description $_.Exception.Message
            Write-Warning "Unable to retrieve risky users: $_"
        }

        # Risky Sign-Ins
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
                } | Export-Csv (Join-Path $outputDir 'Entra_RiskySignIns.csv') -NoTypeInformation -Encoding UTF8
            }
            else {
                Write-Verbose "No risky sign-ins detected — skipping Entra_RiskySignIns.csv"
            }
        }
        catch {
            Add-AuditIssue -Severity 'Warning' -Section 'Entra' -Collector 'Invoke-MgGraphRequest' -Description $_.Exception.Message
            Write-Warning "Unable to retrieve risky sign-ins: $_"
        }
    }

    # ── 15. Authentication Methods Policy ───────────────────────────────────
    $step++; Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Authentication Methods Policy..." -PercentComplete ([int]($step / $totalSteps * 100))

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
        $authMethodRows | Export-Csv (Join-Path $outputDir 'Entra_AuthMethodsPolicy.csv') -NoTypeInformation -Encoding UTF8
    }
    catch {
        Add-AuditIssue -Severity 'Warning' -Section 'Entra' -Collector 'Invoke-MgGraphRequest' -Description $_.Exception.Message
        Write-Warning "Unable to retrieve Authentication Methods Policy: $_"
    }

    # ── 16. External Collaboration Settings ─────────────────────────────────
    $step++; Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — External Collaboration Settings..." -PercentComplete ([int]($step / $totalSteps * 100))

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
        } | Export-Csv (Join-Path $outputDir 'Entra_ExternalCollab.csv') -NoTypeInformation -Encoding UTF8
    }
    catch {
        Add-AuditIssue -Severity 'Warning' -Section 'Entra' -Collector 'Invoke-MgGraphRequest' -Description $_.Exception.Message
        Write-Warning "Unable to retrieve External Collaboration Settings: $_"
    }

    # ── 17. App Registrations + Permissions ─────────────────────────────────
    $step++; Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — App Registrations..." -PercentComplete ([int]($step / $totalSteps * 100))

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
        $appRegRows | Export-Csv (Join-Path $outputDir 'Entra_AppRegistrations.csv') -NoTypeInformation -Encoding UTF8

        # Collect App Registration permissions (requiredResourceAccess -> resolved permission names)
        # Re-use $_spPermCache from enterprise apps section if available
        if (-not (Get-Variable -Name '_spPermCache' -ErrorAction SilentlyContinue) -or $null -eq $_spPermCache) {
            $_spPermCache = @{}
        }
        function _GetCachedSPByAppId {
            param([string]$AppId)
            if (-not $_spPermCache.ContainsKey($AppId)) {
                try {
                    $_spPermCache[$AppId] = Get-MgServicePrincipal -Filter "appId eq '$AppId'" `
                        -ConsistencyLevel eventual -ErrorAction SilentlyContinue | Select-Object -First 1
                } catch { $_spPermCache[$AppId] = $null }
            }
            return $_spPermCache[$AppId]
        }

        $_arPermRows = [System.Collections.Generic.List[object]]::new()
        foreach ($app in $_appRegs) {
            foreach ($resource in @($app.requiredResourceAccess)) {
                if (-not $resource) { continue }
                $_resAppId = $resource.resourceAppId
                $_resSP = _GetCachedSPByAppId -AppId $_resAppId
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
            $_arPermRows | Export-Csv (Join-Path $outputDir 'Entra_AppRegistrationPermissions.csv') -NoTypeInformation -Encoding UTF8
        }
    }
    catch {
        Add-AuditIssue -Severity 'Warning' -Section 'Entra' -Collector 'Invoke-MgGraphRequest' -Description $_.Exception.Message
        Write-Warning "Unable to retrieve App Registrations: $_"
    }

    # ── 18. PIM Role Assignments (P2 required) ──────────────────────────────
    $step++; Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — PIM Role Assignments..." -PercentComplete ([int]($step / $totalSteps * 100))

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
            } | Export-Csv (Join-Path $outputDir 'Entra_PIMAssignments.csv') -NoTypeInformation -Encoding UTF8
        }
        catch {
            Add-AuditIssue -Severity 'Warning' -Section 'Entra' -Collector 'Invoke-MgGraphRequest' -Description $_.Exception.Message
            Write-Warning "Unable to retrieve PIM role assignments: $_"
        }
    }

    # ── 19. Org-level User / App Settings ───────────────────────────────────
    $step++; Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Organisation settings..." -PercentComplete ([int]($step / $totalSteps * 100))

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
        } | Export-Csv (Join-Path $outputDir 'Entra_OrgSettings.csv') -NoTypeInformation -Encoding UTF8
    }
    catch {
        Add-AuditIssue -Severity 'Warning' -Section 'Entra' -Collector 'Invoke-MgGraphRequest' -Description $_.Exception.Message
        Write-Warning "Unable to retrieve organisation settings: $_"
    }

    # ── 20. Microsoft 365 Lighthouse Status ─────────────────────────────────
    $step++; Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Lighthouse enrollment..." -PercentComplete ([int]($step / $totalSteps * 100))

    try {
        $_lighthouseAppId = '2828a423-d5cc-4818-8285-aa945d95017a'
        $_lighthouseSp    = Get-MgServicePrincipal -Filter "appId eq '$_lighthouseAppId'" -Property Id,DisplayName,AppId -ErrorAction Stop | Select-Object -First 1

        [PSCustomObject]@{
            Enrolled           = if ($_lighthouseSp) { 'True' } else { 'False' }
            ServicePrincipalId = if ($_lighthouseSp) { $_lighthouseSp.Id }          else { '' }
            DisplayName        = if ($_lighthouseSp) { $_lighthouseSp.DisplayName } else { '' }
        } | Export-Csv (Join-Path $outputDir 'Entra_LighthouseStatus.csv') -NoTypeInformation -Encoding UTF8
    }
    catch {
        Add-AuditIssue -Severity 'Info' -Section 'Entra' -Collector 'Get-MgServicePrincipal (Lighthouse)' -Description $_.Exception.Message
        Write-Warning "Unable to check Lighthouse enrollment: $_"
    }

    # ════════════════════════════════════════════════════════════════════════
    # DONE
    # ════════════════════════════════════════════════════════════════════════
    Write-Progress -Id 1 -Activity $activity -Completed
    Write-Host "Entra Inventory complete. Results saved to: $outputDir" -ForegroundColor Green

    return @{
        UserCount       = @($userReport).Count
        LicensedCount   = $licensedUsers.Count
        UnlicensedCount = $unlicensedUsers.Count
        LicenseCount    = $subscribedSkus.Count
        AdminRoleCount  = $adminReport.Count
        GlobalAdminCount = $globalAdmins.Count
        GuestCount      = @($guestData).Count
        GroupCount       = @($groupData).Count
        CAPolicyCount   = @($caPolicyData).Count
        SignInRetention  = $retentionDays
        HasP2           = $_hasP2
    }
}
