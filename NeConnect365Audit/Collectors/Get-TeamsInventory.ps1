function Get-TeamsInventory {
    <#
    .SYNOPSIS
        Collects Microsoft Teams configuration data and exports to CSV files.

    .DESCRIPTION
        Queries Microsoft Teams PowerShell for tenant configuration data including:
        - Federation and external access settings
        - Teams client configuration (guest access, cloud storage)
        - Global meeting policy (lobby, anonymous join, recording)
        - Guest meeting and calling configuration
        - Messaging policies
        - App permission and setup policies
        - Channel policies

        Graph and Teams connections must already be established by the orchestrator.

    .OUTPUTS
        Hashtable with summary counts for the orchestrator.
    #>
    [CmdletBinding()]
    param()

    $ctx       = Get-AuditContext
    $outputDir = $ctx.RawOutputPath

    Write-Host "`nStarting Microsoft Teams Inventory for $($ctx.OrgName)..." -ForegroundColor Cyan

    $step       = 0
    $totalSteps = 9
    $activity   = "Teams Inventory — $($ctx.OrgName)"

    # ── Tracking counters for summary ──────────────────────────────────────
    $_meetingPolicyCount    = 0
    $_messagingPolicyCount  = 0
    $_appPermPolicyCount    = 0
    $_appSetupPolicyCount   = 0
    $_channelPolicyCount    = 0

    # =========================================
    # ===   Step 1 — Federation Config      ===
    # =========================================
    $step++
    Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Federation and external access settings..." -PercentComplete ([int]($step / $totalSteps * 100))
    try {
        $_fedConfig = Get-CsTenantFederationConfiguration -ErrorAction Stop

        # AllowedDomains is a collection object; extract domain names
        $_allowedDomains = @()
        $_blockedDomains = @()
        try {
            if ($_fedConfig.AllowedDomains -and $_fedConfig.AllowedDomains.AllowedDomain) {
                $_allowedDomains = @($_fedConfig.AllowedDomains.AllowedDomain | ForEach-Object { $_.Domain })
            }
        } catch {}
        try {
            if ($_fedConfig.BlockedDomains) {
                $_blockedDomains = @($_fedConfig.BlockedDomains | ForEach-Object { $_.Domain })
            }
        } catch {}

        [PSCustomObject]@{
            AllowFederatedUsers              = $null -ne $_fedConfig.AllowFederatedUsers ? "$($_fedConfig.AllowFederatedUsers)" : ''
            AllowPublicUsers                 = $null -ne $_fedConfig.AllowPublicUsers ? "$($_fedConfig.AllowPublicUsers)" : ''
            AllowTeamsConsumer               = $null -ne $_fedConfig.AllowTeamsConsumer ? "$($_fedConfig.AllowTeamsConsumer)" : ''
            AllowTeamsConsumerInbound        = $null -ne $_fedConfig.AllowTeamsConsumerInbound ? "$($_fedConfig.AllowTeamsConsumerInbound)" : ''
            TreatDiscoveredPartnersAsUnverified = $null -ne $_fedConfig.TreatDiscoveredPartnersAsUnverified ? "$($_fedConfig.TreatDiscoveredPartnersAsUnverified)" : ''
            AllowedDomainsCount              = $_allowedDomains.Count
            AllowedDomainsList               = ($_allowedDomains -join '; ')
            BlockedDomainsCount              = $_blockedDomains.Count
            BlockedDomainsList               = ($_blockedDomains -join '; ')
        } | Export-Csv -Path (Join-Path $outputDir 'Teams_FederationConfig.csv') -NoTypeInformation -Encoding UTF8
        Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Federation and external access settings..." -CurrentOperation "Saved: Teams_FederationConfig.csv" -PercentComplete ([int]($step / $totalSteps * 100))
    }
    catch {
        Add-AuditIssue -Severity 'Warning' -Section 'Teams' -Collector 'Get-CsTenantFederationConfiguration' -Description ($_.Exception.Message ?? "$_") -Action 'Check permissions or re-run Setup-365AuditApp.ps1'
        Write-Warning "Step 1 - Federation config failed: $_"
    }


    # =========================================
    # ===   Step 2 — Client Configuration   ===
    # =========================================
    $step++
    Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Teams client configuration..." -PercentComplete ([int]($step / $totalSteps * 100))
    try {
        $_clientCfg = Get-CsTeamsClientConfiguration -ErrorAction Stop

        [PSCustomObject]@{
            AllowBox                         = $null -ne $_clientCfg.AllowBox ? "$($_clientCfg.AllowBox)" : ''
            AllowDropBox                     = $null -ne $_clientCfg.AllowDropBox ? "$($_clientCfg.AllowDropBox)" : ''
            AllowEgnyte                      = $null -ne $_clientCfg.AllowEgnyte ? "$($_clientCfg.AllowEgnyte)" : ''
            AllowGoogleDrive                 = $null -ne $_clientCfg.AllowGoogleDrive ? "$($_clientCfg.AllowGoogleDrive)" : ''
            AllowShareFile                   = $null -ne $_clientCfg.AllowShareFile ? "$($_clientCfg.AllowShareFile)" : ''
            AllowGuestUser                   = $null -ne $_clientCfg.AllowGuestUser ? "$($_clientCfg.AllowGuestUser)" : ''
            AllowScopedPeopleSearchandAccess = $null -ne $_clientCfg.AllowScopedPeopleSearchandAccess ? "$($_clientCfg.AllowScopedPeopleSearchandAccess)" : ''
            AllowSkypeBusinessInterop        = $null -ne $_clientCfg.AllowSkypeBusinessInterop ? "$($_clientCfg.AllowSkypeBusinessInterop)" : ''
        } | Export-Csv -Path (Join-Path $outputDir 'Teams_ClientConfig.csv') -NoTypeInformation -Encoding UTF8
        Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Teams client configuration..." -CurrentOperation "Saved: Teams_ClientConfig.csv" -PercentComplete ([int]($step / $totalSteps * 100))
    }
    catch {
        Add-AuditIssue -Severity 'Warning' -Section 'Teams' -Collector 'Get-CsTeamsClientConfiguration' -Description ($_.Exception.Message ?? "$_") -Action 'Check permissions or re-run Setup-365AuditApp.ps1'
        Write-Warning "Step 2 - Client configuration failed: $_"
    }


    # =========================================
    # ===   Step 3 — Meeting Policies       ===
    # =========================================
    $step++
    Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Teams meeting policies..." -PercentComplete ([int]($step / $totalSteps * 100))
    try {
        $_meetingPolicies = @(Get-CsTeamsMeetingPolicy -ErrorAction Stop)
        $_meetingPolicyCount = $_meetingPolicies.Count

        $_meetingPolicies | ForEach-Object {
            [PSCustomObject]@{
                Identity                                  = $_.Identity
                AllowAnonymousUsersToStartMeeting         = $null -ne $_.AllowAnonymousUsersToStartMeeting ? "$($_.AllowAnonymousUsersToStartMeeting)" : ''
                AutoAdmittedUsers                         = "$($_.AutoAdmittedUsers)"
                AllowCloudRecording                       = $null -ne $_.AllowCloudRecording ? "$($_.AllowCloudRecording)" : ''
                AllowRecordingStorageOutsideRegion        = $null -ne $_.AllowRecordingStorageOutsideRegion ? "$($_.AllowRecordingStorageOutsideRegion)" : ''
                AllowExternalParticipantGiveRequestControl = $null -ne $_.AllowExternalParticipantGiveRequestControl ? "$($_.AllowExternalParticipantGiveRequestControl)" : ''
                AllowExternalNonTrustedMeetingChat        = $null -ne $_.AllowExternalNonTrustedMeetingChat ? "$($_.AllowExternalNonTrustedMeetingChat)" : ''
                AllowIPVideo                              = $null -ne $_.AllowIPVideo ? "$($_.AllowIPVideo)" : ''
                MeetingChatEnabledType                    = "$($_.MeetingChatEnabledType)"
                AllowMeetNow                              = $null -ne $_.AllowMeetNow ? "$($_.AllowMeetNow)" : ''
                AllowTranscription                        = $null -ne $_.AllowTranscription ? "$($_.AllowTranscription)" : ''
                AllowBreakoutRooms                        = $null -ne $_.AllowBreakoutRooms ? "$($_.AllowBreakoutRooms)" : ''
            }
        } | Export-Csv -Path (Join-Path $outputDir 'Teams_MeetingPolicies.csv') -NoTypeInformation -Encoding UTF8
        Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Teams meeting policies..." -CurrentOperation "Saved: Teams_MeetingPolicies.csv ($($_meetingPolicies.Count) policies)" -PercentComplete ([int]($step / $totalSteps * 100))
    }
    catch {
        Add-AuditIssue -Severity 'Warning' -Section 'Teams' -Collector 'Get-CsTeamsMeetingPolicy' -Description ($_.Exception.Message ?? "$_") -Action 'Check permissions or re-run Setup-365AuditApp.ps1'
        Write-Warning "Step 3 - Meeting policies failed: $_"
    }


    # =========================================
    # ===   Step 4 — Guest Meeting Config   ===
    # =========================================
    $step++
    Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Guest meeting configuration..." -PercentComplete ([int]($step / $totalSteps * 100))
    try {
        $_guestMtg = Get-CsTeamsGuestMeetingConfiguration -ErrorAction Stop

        [PSCustomObject]@{
            AllowIPVideo            = $null -ne $_guestMtg.AllowIPVideo ? "$($_guestMtg.AllowIPVideo)" : ''
            ScreenSharingMode       = "$($_guestMtg.ScreenSharingMode)"
            AllowMeetNow            = $null -ne $_guestMtg.AllowMeetNow ? "$($_guestMtg.AllowMeetNow)" : ''
            LiveCaptionsEnabledType = "$($_guestMtg.LiveCaptionsEnabledType)"
        } | Export-Csv -Path (Join-Path $outputDir 'Teams_GuestMeetingConfig.csv') -NoTypeInformation -Encoding UTF8
        Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Guest meeting configuration..." -CurrentOperation "Saved: Teams_GuestMeetingConfig.csv" -PercentComplete ([int]($step / $totalSteps * 100))
    }
    catch {
        Add-AuditIssue -Severity 'Warning' -Section 'Teams' -Collector 'Get-CsTeamsGuestMeetingConfiguration' -Description ($_.Exception.Message ?? "$_") -Action 'Check permissions or re-run Setup-365AuditApp.ps1'
        Write-Warning "Step 4 - Guest meeting config failed: $_"
    }


    # =========================================
    # ===   Step 5 — Guest Calling Config   ===
    # =========================================
    $step++
    Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Guest calling configuration..." -PercentComplete ([int]($step / $totalSteps * 100))
    try {
        $_guestCall = Get-CsTeamsGuestCallingConfiguration -ErrorAction Stop

        [PSCustomObject]@{
            AllowPrivateCalling = $null -ne $_guestCall.AllowPrivateCalling ? "$($_guestCall.AllowPrivateCalling)" : ''
        } | Export-Csv -Path (Join-Path $outputDir 'Teams_GuestCallingConfig.csv') -NoTypeInformation -Encoding UTF8
        Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Guest calling configuration..." -CurrentOperation "Saved: Teams_GuestCallingConfig.csv" -PercentComplete ([int]($step / $totalSteps * 100))
    }
    catch {
        Add-AuditIssue -Severity 'Warning' -Section 'Teams' -Collector 'Get-CsTeamsGuestCallingConfiguration' -Description ($_.Exception.Message ?? "$_") -Action 'Check permissions or re-run Setup-365AuditApp.ps1'
        Write-Warning "Step 5 - Guest calling config failed: $_"
    }


    # =========================================
    # ===   Step 6 — Messaging Policies     ===
    # =========================================
    $step++
    Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Teams messaging policies..." -PercentComplete ([int]($step / $totalSteps * 100))
    try {
        $_msgPolicies = @(Get-CsTeamsMessagingPolicy -ErrorAction Stop)
        $_messagingPolicyCount = $_msgPolicies.Count

        $_msgPolicies | ForEach-Object {
            [PSCustomObject]@{
                Identity               = $_.Identity
                AllowUserEditMessage   = $null -ne $_.AllowUserEditMessage ? "$($_.AllowUserEditMessage)" : ''
                AllowUserDeleteMessage = $null -ne $_.AllowUserDeleteMessage ? "$($_.AllowUserDeleteMessage)" : ''
                AllowUserDeleteChat    = $null -ne $_.AllowUserDeleteChat ? "$($_.AllowUserDeleteChat)" : ''
                AllowGiphy             = $null -ne $_.AllowGiphy ? "$($_.AllowGiphy)" : ''
                GiphyRatingType        = "$($_.GiphyRatingType)"
                AllowMemes             = $null -ne $_.AllowMemes ? "$($_.AllowMemes)" : ''
                AllowImmersiveReader   = $null -ne $_.AllowImmersiveReader ? "$($_.AllowImmersiveReader)" : ''
                AllowUserChat          = $null -ne $_.AllowUserChat ? "$($_.AllowUserChat)" : ''
            }
        } | Export-Csv -Path (Join-Path $outputDir 'Teams_MessagingPolicies.csv') -NoTypeInformation -Encoding UTF8
        Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Teams messaging policies..." -CurrentOperation "Saved: Teams_MessagingPolicies.csv ($($_msgPolicies.Count) policies)" -PercentComplete ([int]($step / $totalSteps * 100))
    }
    catch {
        Add-AuditIssue -Severity 'Warning' -Section 'Teams' -Collector 'Get-CsTeamsMessagingPolicy' -Description ($_.Exception.Message ?? "$_") -Action 'Check permissions or re-run Setup-365AuditApp.ps1'
        Write-Warning "Step 6 - Messaging policies failed: $_"
    }


    # =========================================
    # ===   Step 7 — App Permission Policies ===
    # =========================================
    $step++
    Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Teams app permission policies..." -PercentComplete ([int]($step / $totalSteps * 100))
    try {
        $_appPermPolicies = @(Get-CsTeamsAppPermissionPolicy -ErrorAction Stop)
        $_appPermPolicyCount = $_appPermPolicies.Count

        $_appPermPolicies | ForEach-Object {
            [PSCustomObject]@{
                Identity           = $_.Identity
                DefaultCatalogApps = "$($_.DefaultCatalogApps)"
                GlobalCatalogApps  = "$($_.GlobalCatalogApps)"
                PrivateCatalogApps = "$($_.PrivateCatalogApps)"
            }
        } | Export-Csv -Path (Join-Path $outputDir 'Teams_AppPermissionPolicies.csv') -NoTypeInformation -Encoding UTF8
        Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Teams app permission policies..." -CurrentOperation "Saved: Teams_AppPermissionPolicies.csv ($($_appPermPolicies.Count) policies)" -PercentComplete ([int]($step / $totalSteps * 100))
    }
    catch {
        Add-AuditIssue -Severity 'Warning' -Section 'Teams' -Collector 'Get-CsTeamsAppPermissionPolicy' -Description ($_.Exception.Message ?? "$_") -Action 'Check permissions or re-run Setup-365AuditApp.ps1'
        Write-Warning "Step 7 - App permission policies failed: $_"
    }


    # =========================================
    # ===   Step 8 — App Setup Policies     ===
    # =========================================
    $step++
    Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Teams app setup policies..." -PercentComplete ([int]($step / $totalSteps * 100))
    try {
        $_appSetupPolicies = @(Get-CsTeamsAppSetupPolicy -ErrorAction Stop)
        $_appSetupPolicyCount = $_appSetupPolicies.Count

        $_appSetupPolicies | ForEach-Object {
            [PSCustomObject]@{
                Identity         = $_.Identity
                AllowSideloading = $null -ne $_.AllowSideloading ? "$($_.AllowSideloading)" : ''
                AllowUserPinning = $null -ne $_.AllowUserPinning ? "$($_.AllowUserPinning)" : ''
            }
        } | Export-Csv -Path (Join-Path $outputDir 'Teams_AppSetupPolicies.csv') -NoTypeInformation -Encoding UTF8
        Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Teams app setup policies..." -CurrentOperation "Saved: Teams_AppSetupPolicies.csv ($($_appSetupPolicies.Count) policies)" -PercentComplete ([int]($step / $totalSteps * 100))
    }
    catch {
        Add-AuditIssue -Severity 'Warning' -Section 'Teams' -Collector 'Get-CsTeamsAppSetupPolicy' -Description ($_.Exception.Message ?? "$_") -Action 'Check permissions or re-run Setup-365AuditApp.ps1'
        Write-Warning "Step 8 - App setup policies failed: $_"
    }


    # =========================================
    # ===   Step 9 — Channel Policies       ===
    # =========================================
    $step++
    Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Teams channel policies..." -PercentComplete ([int]($step / $totalSteps * 100))
    try {
        $_channelPolicies = @(Get-CsTeamsChannelPolicy -ErrorAction Stop)
        $_channelPolicyCount = $_channelPolicies.Count

        $_channelPolicies | ForEach-Object {
            [PSCustomObject]@{
                Identity                  = $_.Identity
                AllowOrgWideTeamCreation  = $null -ne $_.AllowOrgWideTeamCreation ? "$($_.AllowOrgWideTeamCreation)" : ''
                AllowPrivateTeamDiscovery = $null -ne $_.AllowPrivateTeamDiscovery ? "$($_.AllowPrivateTeamDiscovery)" : ''
                AllowSharedChannels       = $null -ne $_.AllowSharedChannels ? "$($_.AllowSharedChannels)" : ''
                AllowPrivateChannels      = $null -ne $_.AllowPrivateChannels ? "$($_.AllowPrivateChannels)" : ''
            }
        } | Export-Csv -Path (Join-Path $outputDir 'Teams_ChannelPolicies.csv') -NoTypeInformation -Encoding UTF8
        Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Teams channel policies..." -CurrentOperation "Saved: Teams_ChannelPolicies.csv ($($_channelPolicies.Count) policies)" -PercentComplete 100
    }
    catch {
        if ($_.Exception -is [System.Management.Automation.CommandNotFoundException] -or $_ -match 'not recognized|not found') {
            Write-Verbose "Get-CsTeamsChannelPolicy not available in this Teams module version — skipping channel policies."
        } else {
            Add-AuditIssue -Severity 'Warning' -Section 'Teams' -Collector 'Get-CsTeamsChannelPolicy' -Description ($_.Exception.Message ?? "$_") -Action 'Check permissions or re-run Setup-365AuditApp.ps1'
            Write-Warning "Step 9 - Channel policies failed: $_"
        }
    }


    # ================================
    # ===   Done                    ===
    # ================================
    Write-Progress -Id 1 -Activity $activity -Completed
    Write-Host "Teams Inventory complete. Results saved to: $outputDir" -ForegroundColor Green

    return @{
        MeetingPolicyCount    = $_meetingPolicyCount
        MessagingPolicyCount  = $_messagingPolicyCount
        AppPermPolicyCount    = $_appPermPolicyCount
        AppSetupPolicyCount   = $_appSetupPolicyCount
        ChannelPolicyCount    = $_channelPolicyCount
    }
}
