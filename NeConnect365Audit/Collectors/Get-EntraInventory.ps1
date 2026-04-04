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

    # ── Continue with remaining sections... ─────────────────────────────────
    # The remaining data collection sections follow the same pattern:
    # fetch from Graph → transform to PSCustomObject → Export-Csv
    # TODO: Port remaining sections from Invoke-EntraAudit.ps1:
    #   4. SSPR Configuration
    #   5. Authentication Methods Policy
    #   6. User Summary (licensed + unlicensed)
    #   7. Admin Roles + Global Admins
    #   8. Guest Users
    #   9. Groups
    #  10. Conditional Access Policies
    #  11. Trusted/Named Locations
    #  12. Security Defaults
    #  13. Secure Score + Controls
    #  14. Enterprise Apps + Permissions
    #  15. Risky Users + Risky Sign-Ins
    #  16. Organisation Settings
    #  17. Partner Relationships
    #  18. Lighthouse Status

    Write-Progress -Id 1 -Activity $activity -Completed
    Write-Host "Entra Inventory complete. Results saved to: $outputDir" -ForegroundColor Green

    return @{
        UserCount    = $signInMap.Count
        LicenseCount = $subscribedSkus.Count
        SignInRetention = $retentionDays
    }
}
