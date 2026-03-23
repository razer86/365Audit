<#
.SYNOPSIS
    Performs a security-focused audit of Microsoft Teams tenant configuration.

.DESCRIPTION
    Connects to Microsoft Teams PowerShell and Microsoft Graph to collect Teams
    configuration data including:
    - Federation and external access settings
    - Teams client configuration (guest access, cloud storage)
    - Global meeting policy (lobby, anonymous join, recording)
    - Guest meeting and calling configuration
    - Messaging policies
    - App permission and setup policies
    - Channel policies
    - Org-wide Teams app settings (via Graph)

    Output CSVs:
    - Teams_FederationConfig.csv
    - Teams_ClientConfig.csv
    - Teams_MeetingPolicies.csv
    - Teams_GuestMeetingConfig.csv
    - Teams_GuestCallingConfig.csv
    - Teams_MessagingPolicies.csv
    - Teams_AppPermissionPolicies.csv
    - Teams_AppSetupPolicies.csv
    - Teams_ChannelPolicies.csv

.NOTES
    Author      : Raymond Slater
    Version     : 1.1.2
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

$ScriptVersion = "1.1.2"
Write-Verbose "Invoke-TeamsAudit.ps1 loaded (v$ScriptVersion)"

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

# === Retrieve shared output folder ===
try {
    $context   = Initialize-AuditOutput
    $outputDir = $context.RawOutputPath
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
}
catch {
    Write-Error "Failed to initialise audit output folder: $_"
    exit 1
}

$step       = 0
$totalSteps = 9
$activity   = "Teams Audit — $($context.OrgName)"

# === Connect to Microsoft Graph ===
Write-Host "`nConnecting to Microsoft Graph..." -ForegroundColor Cyan
Connect-MgGraphSecure

# === Connect to Microsoft Teams ===
Write-Host "Connecting to Microsoft Teams..." -ForegroundColor Cyan
Connect-TeamsSecure

Write-Host "`nStarting Microsoft Teams Audit for $($context.OrgName)..." -ForegroundColor Cyan


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
    } | Export-Csv -Path (Join-Path $outputDir "Teams_FederationConfig.csv") -NoTypeInformation -Encoding UTF8
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
    } | Export-Csv -Path (Join-Path $outputDir "Teams_ClientConfig.csv") -NoTypeInformation -Encoding UTF8
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
    } | Export-Csv -Path (Join-Path $outputDir "Teams_MeetingPolicies.csv") -NoTypeInformation -Encoding UTF8
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
    } | Export-Csv -Path (Join-Path $outputDir "Teams_GuestMeetingConfig.csv") -NoTypeInformation -Encoding UTF8
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
    } | Export-Csv -Path (Join-Path $outputDir "Teams_GuestCallingConfig.csv") -NoTypeInformation -Encoding UTF8
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
    } | Export-Csv -Path (Join-Path $outputDir "Teams_MessagingPolicies.csv") -NoTypeInformation -Encoding UTF8
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

    $_appPermPolicies | ForEach-Object {
        [PSCustomObject]@{
            Identity           = $_.Identity
            DefaultCatalogApps = "$($_.DefaultCatalogApps)"
            GlobalCatalogApps  = "$($_.GlobalCatalogApps)"
            PrivateCatalogApps = "$($_.PrivateCatalogApps)"
        }
    } | Export-Csv -Path (Join-Path $outputDir "Teams_AppPermissionPolicies.csv") -NoTypeInformation -Encoding UTF8
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

    $_appSetupPolicies | ForEach-Object {
        [PSCustomObject]@{
            Identity         = $_.Identity
            AllowSideloading = $null -ne $_.AllowSideloading ? "$($_.AllowSideloading)" : ''
            AllowUserPinning = $null -ne $_.AllowUserPinning ? "$($_.AllowUserPinning)" : ''
        }
    } | Export-Csv -Path (Join-Path $outputDir "Teams_AppSetupPolicies.csv") -NoTypeInformation -Encoding UTF8
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

    $_channelPolicies | ForEach-Object {
        [PSCustomObject]@{
            Identity                  = $_.Identity
            AllowOrgWideTeamCreation  = $null -ne $_.AllowOrgWideTeamCreation ? "$($_.AllowOrgWideTeamCreation)" : ''
            AllowPrivateTeamDiscovery = $null -ne $_.AllowPrivateTeamDiscovery ? "$($_.AllowPrivateTeamDiscovery)" : ''
            AllowSharedChannels       = $null -ne $_.AllowSharedChannels ? "$($_.AllowSharedChannels)" : ''
            AllowPrivateChannels      = $null -ne $_.AllowPrivateChannels ? "$($_.AllowPrivateChannels)" : ''
        }
    } | Export-Csv -Path (Join-Path $outputDir "Teams_ChannelPolicies.csv") -NoTypeInformation -Encoding UTF8
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
Write-Host "`nMicrosoft Teams Audit complete. Results saved to: $outputDir`n" -ForegroundColor Green
