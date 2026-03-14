<#
.SYNOPSIS
    Performs an Exchange Online audit of mailboxes, permissions, rules, and security policies.

.DESCRIPTION
    Connects to Exchange Online and generates the following reports:
    - Mailbox inventory (user & shared) with archive status and sizes
    - Mailbox permissions (Full Access, Send As, Send on Behalf)
    - Distribution lists and dynamic rules
    - Inbox rules with internal/external forwarding detection
    - Mail flow (transport) rules
    - External forwarding global settings
    - Anti-phishing policies
    - Anti-spam and malware filter policies
    - DKIM signing configuration
    - Mailbox audit settings
    - Resource mailboxes (room & equipment)

    Output CSVs are written to the shared audit output folder.

.NOTES
    Author      : Raymond Slater
    Version     : 1.9.0
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

$ScriptVersion = "1.9.0"
Write-Verbose "Invoke-ExchangeAudit.ps1 loaded (v$ScriptVersion)"

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# === Retrieve shared output folder ===
try {
    $context   = Initialize-AuditOutput
    $outputDir = $context.OutputPath
}
catch {
    Write-Error "Failed to initialise audit output directory: $_"
    exit 1
}

# === Ensure ExchangeOnlineManagement module is available ===
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Write-Host "Installing ExchangeOnlineManagement module..." -ForegroundColor Yellow
    Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force
}
Import-Module ExchangeOnlineManagement -ErrorAction Stop

# === Connect to Exchange Online ===
Connect-ExchangeOnlineSecure

Write-Host "`nStarting Exchange Audit for $($context.OrgName)..." -ForegroundColor Cyan

$step       = 0
$totalSteps = 15
$activity   = "Exchange Audit — $($context.OrgName)"


# === 1. Mailbox Inventory ===
$step++
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Gathering mailbox inventory..." -PercentComplete ([int]($step / $totalSteps * 100))

$mailboxes        = Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox, SharedMailbox
$mailboxInventory = foreach ($mbx in $mailboxes) {
    $stats = Get-MailboxStatistics -Identity $mbx.PrimarySmtpAddress
    # TotalItemSize is a deserialized ByteQuantifiedSize in EXO v3 — parse bytes from the string representation
    $sizeStr   = $stats.TotalItemSize.ToString()
    $sizeBytes = if ($sizeStr -match '\((\d[\d,]+)\s+bytes\)') { [long]($Matches[1] -replace ',') } else { 0 }
    $usedMB    = [math]::Round($sizeBytes / 1MB, 2)

    # Parse ProhibitSendReceiveQuota for the mailbox storage limit
    $quotaStr   = $mbx.ProhibitSendReceiveQuota.ToString()
    $quotaBytes = if ($quotaStr -match '\((\d[\d,]+)\s+bytes\)') { [long]($Matches[1] -replace ',') } else { 0 }
    $limitMB    = if ($quotaBytes -gt 0) { [math]::Round($quotaBytes / 1MB, 2) } else { $null }
    $freeMB     = if ($limitMB)          { [math]::Round($limitMB - $usedMB, 2) } else { $null }

    # Archive size (only if archive is active)
    $archiveSizeMB = $null
    if ($mbx.ArchiveStatus -eq "Active") {
        try {
            $archStats     = Get-MailboxStatistics -Identity $mbx.PrimarySmtpAddress -Archive -ErrorAction Stop
            $archSizeStr   = $archStats.TotalItemSize.ToString()
            $archBytes     = if ($archSizeStr -match '\((\d[\d,]+)\s+bytes\)') { [long]($Matches[1] -replace ',') } else { 0 }
            $archiveSizeMB = [math]::Round($archBytes / 1MB, 2)
        }
        catch { }
    }

    [PSCustomObject]@{
        DisplayName           = $mbx.DisplayName
        UserPrincipalName     = $mbx.UserPrincipalName
        RecipientType         = $mbx.RecipientTypeDetails
        ArchiveEnabled        = $mbx.ArchiveStatus -eq "Active"
        TotalSizeMB           = $usedMB
        FreeMB                = $freeMB
        LimitMB               = $limitMB
        ArchiveSizeMB         = $archiveSizeMB
        ItemCount             = $stats.ItemCount
        LitigationHoldEnabled = $mbx.LitigationHoldEnabled
    }
}
$mailboxInventory | Export-Csv "$outputDir\Exchange_Mailboxes.csv" -NoTypeInformation -Encoding UTF8
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Gathering mailbox inventory..." -CurrentOperation "Saved: Exchange_Mailboxes.csv" -PercentComplete ([int]($step / $totalSteps * 100))


# === 2. Mailbox Permissions ===
$step++
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Collecting mailbox permissions..." -PercentComplete ([int]($step / $totalSteps * 100))

# EXO requires -Identity per mailbox; calling without it is unsupported in REST mode
$fullAccessPerms = foreach ($mbx in $mailboxes) {
    try {
        Get-MailboxPermission -Identity $mbx.PrimarySmtpAddress |
            Where-Object { -not $_.IsInherited -and $_.User -notlike 'NT AUTHORITY\*' } |
            ForEach-Object {
                [PSCustomObject]@{
                    MailboxUPN   = $mbx.UserPrincipalName
                    Identity     = $_.Identity
                    User         = $_.User.ToString()
                    AccessRights = ($_.AccessRights -join ", ")
                }
            }
    }
    catch {
        Write-Warning "Failed to get permissions for: $($mbx.DisplayName)"
    }
}
$fullAccessPerms | Export-Csv "$outputDir\Exchange_Permissions_FullAccess.csv" -NoTypeInformation -Encoding UTF8
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Collecting mailbox permissions..." -CurrentOperation "Saved: Exchange_Permissions_FullAccess.csv" -PercentComplete ([int]($step / $totalSteps * 100))

$sendAsPerms = foreach ($mbx in $mailboxes) {
    try {
        Get-RecipientPermission -Identity $mbx.PrimarySmtpAddress -ErrorAction Stop |
            Where-Object { $_.AccessRights -contains "SendAs" -and $_.Trustee -notlike 'NT AUTHORITY\*' } |
            ForEach-Object {
                [PSCustomObject]@{
                    MailboxUPN   = $mbx.UserPrincipalName
                    Identity     = $_.Identity
                    Trustee      = $_.Trustee.ToString()
                    AccessRights = ($_.AccessRights -join ", ")
                }
            }
    }
    catch {
        Write-Warning "Failed to get SendAs permissions for: $($mbx.DisplayName)"
    }
}
$sendAsPerms | Export-Csv "$outputDir\Exchange_Permissions_SendAs.csv" -NoTypeInformation -Encoding UTF8
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Collecting mailbox permissions..." -CurrentOperation "Saved: Exchange_Permissions_SendAs.csv" -PercentComplete ([int]($step / $totalSteps * 100))

$sendOnBehalf = foreach ($mbx in $mailboxes) {
    foreach ($delegate in $mbx.GrantSendOnBehalfTo) {
        [PSCustomObject]@{
            MailboxUPN = $mbx.UserPrincipalName
            Mailbox    = $mbx.DisplayName
            Delegate   = $delegate.Name
        }
    }
}
$sendOnBehalf | Export-Csv "$outputDir\Exchange_Permissions_SendOnBehalf.csv" -NoTypeInformation -Encoding UTF8
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Collecting mailbox permissions..." -CurrentOperation "Saved: Exchange_Permissions_FullAccess.csv, Exchange_Permissions_SendAs.csv, Exchange_Permissions_SendOnBehalf.csv" -PercentComplete ([int]($step / $totalSteps * 100))


# === 3. Distribution Lists ===
$step++
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Gathering distribution list details..." -PercentComplete ([int]($step / $totalSteps * 100))

$dlReport = foreach ($dl in (Get-DistributionGroup -ResultSize Unlimited)) {
    $members   = Get-DistributionGroupMember -Identity $dl.Identity -ErrorAction SilentlyContinue
    $isDynamic = $dl.RecipientTypeDetails -eq 'DynamicDistributionGroup'
    [PSCustomObject]@{
        DisplayName    = $dl.DisplayName
        EmailAddress   = $dl.PrimarySmtpAddress
        GroupType      = if ($isDynamic) { "Dynamic" } else { "Static" }
        MemberCount    = $members.Count
        Members        = ($members.DisplayName -join "; ")
        MembershipRule = if ($isDynamic) { $dl.RecipientFilter } else { "N/A" }
    }
}
$dlReport | Export-Csv "$outputDir\Exchange_DistributionLists.csv" -NoTypeInformation -Encoding UTF8
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Gathering distribution list details..." -CurrentOperation "Saved: Exchange_DistributionLists.csv" -PercentComplete ([int]($step / $totalSteps * 100))


# === 4. Inbox Rules with Forwarding ===
$step++
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Checking inbox rules for forwarding..." -PercentComplete ([int]($step / $totalSteps * 100))

$inboxRules = foreach ($mbx in $mailboxes) {
    try {
        foreach ($rule in (Get-InboxRule -Mailbox $mbx.UserPrincipalName)) {
            $fwdTo      = $rule.ForwardTo             | ForEach-Object { $_.Name }
            $fwdCc      = $rule.ForwardAsAttachmentTo | ForEach-Object { $_.Name }
            $redirectTo = $rule.RedirectTo            | ForEach-Object { $_.Name }

            if ($fwdTo -or $redirectTo -or $fwdCc) {
                [PSCustomObject]@{
                    Mailbox            = $mbx.DisplayName
                    RuleName           = $rule.Name
                    ForwardTo          = ($fwdTo -join "; ")
                    RedirectTo         = ($redirectTo -join "; ")
                    ForwardCc          = ($fwdCc -join "; ")
                    ExternalForwarding = [bool]($rule.ForwardTo -or $rule.RedirectTo)
                }
            }
        }
    }
    catch {
        Write-Warning "Failed to get inbox rules for $($mbx.DisplayName)"
    }
}
$inboxRules | Export-Csv "$outputDir\Exchange_InboxForwardingRules.csv" -NoTypeInformation -Encoding UTF8
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Checking inbox rules for forwarding..." -CurrentOperation "Saved: Exchange_InboxForwardingRules.csv" -PercentComplete ([int]($step / $totalSteps * 100))


# === 5. Mail Flow (Transport) Rules ===
$step++
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Exporting transport rules..." -PercentComplete ([int]($step / $totalSteps * 100))

Get-TransportRule |
    Select-Object -Property Name, Priority, State, Mode, FromAddressContainsWords, SentTo, RedirectMessageTo, BlindCopyTo, ApplyHtmlDisclaimerLocation, ApplyHtmlDisclaimerText |
    Export-Csv "$outputDir\Exchange_TransportRules.csv" -NoTypeInformation -Encoding UTF8
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Exporting transport rules..." -CurrentOperation "Saved: Exchange_TransportRules.csv" -PercentComplete ([int]($step / $totalSteps * 100))


# === 6. External Forwarding Global Settings ===
$step++
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Checking external forwarding settings..." -PercentComplete ([int]($step / $totalSteps * 100))

Get-RemoteDomain |
    Select-Object -Property DomainName, AutoForwardEnabled |
    Export-Csv "$outputDir\Exchange_RemoteDomainForwarding.csv" -NoTypeInformation -Encoding UTF8
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Checking external forwarding settings..." -CurrentOperation "Saved: Exchange_RemoteDomainForwarding.csv" -PercentComplete ([int]($step / $totalSteps * 100))


# === 7. Anti-Phishing Policies ===
$step++
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Exporting anti-phish policies..." -PercentComplete ([int]($step / $totalSteps * 100))

Get-AntiPhishPolicy |
    Select-Object -Property Name, EnableTargetedUserProtection, EnableMailboxIntelligence, EnableSpoofIntelligence, EnableATPForSpoof |
    Export-Csv "$outputDir\Exchange_AntiPhishPolicies.csv" -NoTypeInformation -Encoding UTF8
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Exporting anti-phish policies..." -CurrentOperation "Saved: Exchange_AntiPhishPolicies.csv" -PercentComplete ([int]($step / $totalSteps * 100))


# === 8. Anti-Spam / Malware Policies ===
$step++
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Exporting spam and malware filter policies..." -PercentComplete ([int]($step / $totalSteps * 100))

Get-HostedContentFilterPolicy |
    Select-Object -Property Name, SpamAction, HighConfidenceSpamAction, BulkSpamAction |
    Export-Csv "$outputDir\Exchange_SpamPolicies.csv" -NoTypeInformation -Encoding UTF8

Get-MalwareFilterPolicy |
    Select-Object -Property Name, Action, EnableExternalSenderAdminNotification |
    Export-Csv "$outputDir\Exchange_MalwarePolicies.csv" -NoTypeInformation -Encoding UTF8
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Exporting spam and malware filter policies..." -CurrentOperation "Saved: Exchange_SpamPolicies.csv, Exchange_MalwarePolicies.csv" -PercentComplete ([int]($step / $totalSteps * 100))


# === 9. DKIM Status ===
$step++
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Checking DKIM signing configuration..." -PercentComplete ([int]($step / $totalSteps * 100))

$dkimStatus = foreach ($domain in (Get-AcceptedDomain)) {
    try {
        $dkimConfig = Get-DkimSigningConfig -Identity $domain.DomainName -ErrorAction Stop -WarningAction SilentlyContinue
        [PSCustomObject]@{
            Domain         = $domain.DomainName
            DKIMEnabled    = $dkimConfig.Enabled
            Selector1CNAME = $dkimConfig.Selector1CNAME
            Selector2CNAME = $dkimConfig.Selector2CNAME
        }
    }
    catch {
        [PSCustomObject]@{
            Domain         = $domain.DomainName
            DKIMEnabled    = "Not Configured"
            Selector1CNAME = "N/A"
            Selector2CNAME = "N/A"
        }
    }
}
$dkimStatus | Export-Csv "$outputDir\Exchange_DKIM_Status.csv" -NoTypeInformation -Encoding UTF8
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Checking DKIM signing configuration..." -CurrentOperation "Saved: Exchange_DKIM_Status.csv" -PercentComplete ([int]($step / $totalSteps * 100))


# === 10. Mailbox Audit Settings ===
$step++
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Checking mailbox audit settings..." -PercentComplete ([int]($step / $totalSteps * 100))

Get-Mailbox -ResultSize Unlimited |
    Select-Object -Property DisplayName, UserPrincipalName, AuditEnabled |
    Export-Csv "$outputDir\Exchange_MailboxAuditStatus.csv" -NoTypeInformation -Encoding UTF8

$tenantAuditConfig = Get-AdminAuditLogConfig
[PSCustomObject]@{
    UnifiedAuditLogIngestionEnabled = $tenantAuditConfig.UnifiedAuditLogIngestionEnabled
    AdminAuditLogEnabled            = $tenantAuditConfig.AdminAuditLogEnabled
    AuditLogAgeLimit                = $tenantAuditConfig.AdminAuditLogAgeLimit
} | Export-Csv "$outputDir\Exchange_AuditConfig.csv" -NoTypeInformation -Encoding UTF8
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Checking mailbox audit settings..." -CurrentOperation "Saved: Exchange_MailboxAuditStatus.csv, Exchange_AuditConfig.csv" -PercentComplete ([int]($step / $totalSteps * 100))


# === 11. Resource Mailboxes ===
$step++
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Exporting room and equipment mailboxes..." -PercentComplete ([int]($step / $totalSteps * 100))

$resourceReport = foreach ($res in (Get-Mailbox -RecipientTypeDetails RoomMailbox, EquipmentMailbox -ResultSize Unlimited)) {
    $calSettings = Get-MailboxCalendarConfiguration -Identity $res.Identity -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
    [PSCustomObject]@{
        DisplayName       = $res.DisplayName
        ResourceType      = $res.RecipientTypeDetails
        Email             = $res.PrimarySmtpAddress
        BookingWindowDays = $calSettings.BookingWindowInDays
        AllowConflicts    = $calSettings.AllowConflicts
        BookingDelegates  = ($calSettings.ResourceDelegates -join "; ")
    }
}
$resourceReport | Export-Csv "$outputDir\Exchange_ResourceMailboxes.csv" -NoTypeInformation -Encoding UTF8
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Exporting room and equipment mailboxes..." -CurrentOperation "Saved: Exchange_ResourceMailboxes.csv" -PercentComplete ([int]($step / $totalSteps * 100))


# === 12. Outbound Spam Auto-Forward Policy ===
$step++
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Checking outbound spam auto-forwarding policy..." -PercentComplete ([int]($step / $totalSteps * 100))

Get-HostedOutboundSpamFilterPolicy |
    Select-Object -Property Name, AutoForwardingMode |
    Export-Csv "$outputDir\Exchange_OutboundSpamAutoForward.csv" -NoTypeInformation -Encoding UTF8
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Checking outbound spam auto-forwarding policy..." -CurrentOperation "Saved: Exchange_OutboundSpamAutoForward.csv" -PercentComplete ([int]($step / $totalSteps * 100))


# === 13. Shared Mailbox Sign-In Status ===
$step++
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Checking shared mailbox sign-in status..." -PercentComplete ([int]($step / $totalSteps * 100))

Get-User -RecipientTypeDetails SharedMailbox -ResultSize Unlimited |
    Select-Object -Property DisplayName, UserPrincipalName, AccountDisabled |
    Export-Csv "$outputDir\Exchange_SharedMailboxSignIn.csv" -NoTypeInformation -Encoding UTF8
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Checking shared mailbox sign-in status..." -CurrentOperation "Saved: Exchange_SharedMailboxSignIn.csv" -PercentComplete ([int]($step / $totalSteps * 100))


# === 14. Safe Attachments (requires Defender for Office 365 P1) ===
$step++
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Checking Safe Attachments policies..." -PercentComplete ([int]($step / $totalSteps * 100))

try {
    Get-SafeAttachmentPolicy -ErrorAction Stop |
        Select-Object -Property Name, Enable, Action, ActionOnError |
        Export-Csv "$outputDir\Exchange_SafeAttachments.csv" -NoTypeInformation -Encoding UTF8
    Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Checking Safe Attachments policies..." -CurrentOperation "Saved: Exchange_SafeAttachments.csv" -PercentComplete ([int]($step / $totalSteps * 100))
}
catch {
    Write-Warning "  Safe Attachments not available — requires Defender for Office 365 P1 or higher"
}


# === 15. Safe Links (requires Defender for Office 365 P1) ===
$step++
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Checking Safe Links policies..." -PercentComplete ([int]($step / $totalSteps * 100))

try {
    Get-SafeLinksPolicy -ErrorAction Stop |
        Select-Object -Property Name, EnableSafeLinksForEmail, EnableSafeLinksForTeams, DisableUrlRewrite, TrackClicks |
        Export-Csv "$outputDir\Exchange_SafeLinks.csv" -NoTypeInformation -Encoding UTF8
    Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Checking Safe Links policies..." -CurrentOperation "Saved: Exchange_SafeLinks.csv" -PercentComplete ([int]($step / $totalSteps * 100))
}
catch {
    Write-Warning "  Safe Links not available — requires Defender for Office 365 P1 or higher"
}


# ================================
# ===   Done                    ===
# ================================
Write-Progress -Id 1 -Activity $activity -Completed
Write-Host "`nExchange Audit complete. Results saved to: $outputDir`n" -ForegroundColor Green
