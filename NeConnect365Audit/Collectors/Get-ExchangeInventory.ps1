function Get-ExchangeInventory {
    <#
    .SYNOPSIS
        Collects Exchange Online inventory data and exports to CSV files.

    .DESCRIPTION
        Queries Exchange Online for mailbox, permission, rule, and security policy data
        including:
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
        - Mail connectors (inbound and outbound)
        - Accepted domains
        - Legacy/basic auth policies
        - Exchange org config
        - External sender tagging
        - Connection filter policies
        - OWA mailbox policies

        Expects Exchange Online to already be connected by the orchestrator.
        Exports raw data to CSV files in the Raw/ output directory.

    .OUTPUTS
        Hashtable with summary counts for the orchestrator.
    #>
    [CmdletBinding()]
    param()

    $ctx       = Get-AuditContext
    $outputDir = $ctx.RawOutputPath

    Write-Host "`nStarting Exchange Inventory for $($ctx.OrgName)..." -ForegroundColor Cyan

    $step       = 0
    $totalSteps = 22
    $activity   = "Exchange Inventory — $($ctx.OrgName)"


    # === 1. Mailbox Inventory ===
    $step++
    Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Gathering mailbox inventory..." -PercentComplete ([int]($step / $totalSteps * 100))

    $mailboxes        = Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox, SharedMailbox
    $mailboxInventory = foreach ($mbx in $mailboxes) {
        try {
            # ExchangeGuid is always unique; PrimarySmtpAddress can be ambiguous for linked mailboxes
            $stats = Get-MailboxStatistics -Identity $mbx.ExchangeGuid.ToString() -ErrorAction Stop
        }
        catch {
            Write-Warning "Could not retrieve mailbox statistics for $($mbx.PrimarySmtpAddress): $_"
            $stats = $null
        }

        # TotalItemSize is a deserialized ByteQuantifiedSize in EXO v3 — parse bytes from the string representation
        $sizeStr   = if ($stats) { $stats.TotalItemSize.ToString() } else { '' }
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
                $archStats     = Get-MailboxStatistics -Identity $mbx.ExchangeGuid.ToString() -Archive -ErrorAction Stop
                $archSizeStr   = $archStats.TotalItemSize.ToString()
                $archBytes     = if ($archSizeStr -match '\((\d[\d,]+)\s+bytes\)') { [long]($Matches[1] -replace ',') } else { 0 }
                $archiveSizeMB = [math]::Round($archBytes / 1MB, 2)
            }
            catch { } # Archive mailbox may not exist or may be disabled — skip silently
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
            ItemCount             = if ($stats) { $stats.ItemCount } else { $null }
            LitigationHoldEnabled = $mbx.LitigationHoldEnabled
        }
    }
    $mailboxInventory | Export-Csv (Join-Path $outputDir 'Exchange_Mailboxes.csv') -NoTypeInformation -Encoding UTF8
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
    $fullAccessPerms | Export-Csv (Join-Path $outputDir 'Exchange_Permissions_FullAccess.csv') -NoTypeInformation -Encoding UTF8
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
    $sendAsPerms | Export-Csv (Join-Path $outputDir 'Exchange_Permissions_SendAs.csv') -NoTypeInformation -Encoding UTF8
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
    $sendOnBehalf | Export-Csv (Join-Path $outputDir 'Exchange_Permissions_SendOnBehalf.csv') -NoTypeInformation -Encoding UTF8
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
    $dlReport | Export-Csv (Join-Path $outputDir 'Exchange_DistributionLists.csv') -NoTypeInformation -Encoding UTF8
    Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Gathering distribution list details..." -CurrentOperation "Saved: Exchange_DistributionLists.csv" -PercentComplete ([int]($step / $totalSteps * 100))


    # === 4. Inbox Rules with Forwarding ===
    $step++
    Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Checking inbox rules for forwarding..." -PercentComplete ([int]($step / $totalSteps * 100))

    $brokenInboxRules = [System.Collections.Generic.List[PSObject]]::new()

    $inboxRules = foreach ($mbx in $mailboxes) {
        try {
            $ruleWarnings = @()
            $rules = Get-InboxRule -Mailbox $mbx.UserPrincipalName `
                -WarningAction SilentlyContinue -WarningVariable ruleWarnings

            # Warnings of the form: The Inbox rule "Name" contains errors.
            foreach ($w in $ruleWarnings) {
                if ("$w" -match 'Inbox rule\s+"(.+?)"\s+contains errors') {
                    $brokenInboxRules.Add([PSCustomObject]@{
                        Mailbox  = $mbx.DisplayName
                        RuleName = $Matches[1]
                        Status   = 'Broken — rule contains configuration errors, edit or re-create it'
                    })
                }
            }

            foreach ($rule in $rules) {
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
    $inboxRules       | Export-Csv (Join-Path $outputDir 'Exchange_InboxForwardingRules.csv') -NoTypeInformation -Encoding UTF8
    $brokenInboxRules | Export-Csv (Join-Path $outputDir 'Exchange_BrokenInboxRules.csv')     -NoTypeInformation -Encoding UTF8
    Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Checking inbox rules for forwarding..." -CurrentOperation "Saved: Exchange_InboxForwardingRules.csv" -PercentComplete ([int]($step / $totalSteps * 100))


    # === 5. Mail Flow (Transport) Rules ===
    $step++
    Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Exporting transport rules..." -PercentComplete ([int]($step / $totalSteps * 100))

    Get-TransportRule |
        Select-Object -Property Name, Priority, State, Mode, FromAddressContainsWords, SentTo, RedirectMessageTo, BlindCopyTo, ApplyHtmlDisclaimerLocation, ApplyHtmlDisclaimerText |
        Export-Csv (Join-Path $outputDir 'Exchange_TransportRules.csv') -NoTypeInformation -Encoding UTF8
    Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Exporting transport rules..." -CurrentOperation "Saved: Exchange_TransportRules.csv" -PercentComplete ([int]($step / $totalSteps * 100))


    # === 6. External Forwarding Global Settings ===
    $step++
    Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Checking external forwarding settings..." -PercentComplete ([int]($step / $totalSteps * 100))

    Get-RemoteDomain |
        Select-Object -Property DomainName, AutoForwardEnabled |
        Export-Csv (Join-Path $outputDir 'Exchange_RemoteDomainForwarding.csv') -NoTypeInformation -Encoding UTF8
    Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Checking external forwarding settings..." -CurrentOperation "Saved: Exchange_RemoteDomainForwarding.csv" -PercentComplete ([int]($step / $totalSteps * 100))


    # === 7. Anti-Phishing Policies ===
    $step++
    Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Exporting anti-phish policies..." -PercentComplete ([int]($step / $totalSteps * 100))

    Get-AntiPhishPolicy |
        Select-Object -Property Name, EnableTargetedUserProtection, EnableMailboxIntelligence, EnableSpoofIntelligence, EnableATPForSpoof |
        Export-Csv (Join-Path $outputDir 'Exchange_AntiPhishPolicies.csv') -NoTypeInformation -Encoding UTF8
    Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Exporting anti-phish policies..." -CurrentOperation "Saved: Exchange_AntiPhishPolicies.csv" -PercentComplete ([int]($step / $totalSteps * 100))


    # === 8. Anti-Spam / Malware Policies ===
    $step++
    Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Exporting spam and malware filter policies..." -PercentComplete ([int]($step / $totalSteps * 100))

    Get-HostedContentFilterPolicy |
        Select-Object -Property Name, SpamAction, HighConfidenceSpamAction, BulkSpamAction, HighConfidencePhishAction, ZapEnabled, PhishZapEnabled, SpamZapEnabled |
        Export-Csv (Join-Path $outputDir 'Exchange_SpamPolicies.csv') -NoTypeInformation -Encoding UTF8

    Get-MalwareFilterPolicy |
        Select-Object -Property Name, Action, EnableExternalSenderAdminNotification |
        Export-Csv (Join-Path $outputDir 'Exchange_MalwarePolicies.csv') -NoTypeInformation -Encoding UTF8
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
    $dkimStatus | Export-Csv (Join-Path $outputDir 'Exchange_DKIM_Status.csv') -NoTypeInformation -Encoding UTF8
    Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Checking DKIM signing configuration..." -CurrentOperation "Saved: Exchange_DKIM_Status.csv" -PercentComplete ([int]($step / $totalSteps * 100))


    # === 10. Mailbox Audit Settings ===
    $step++
    Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Checking mailbox audit settings..." -PercentComplete ([int]($step / $totalSteps * 100))

    Get-Mailbox -ResultSize Unlimited |
        Select-Object -Property DisplayName, UserPrincipalName, AuditEnabled |
        Export-Csv (Join-Path $outputDir 'Exchange_MailboxAuditStatus.csv') -NoTypeInformation -Encoding UTF8

    $tenantAuditConfig = Get-AdminAuditLogConfig
    [PSCustomObject]@{
        UnifiedAuditLogIngestionEnabled = $tenantAuditConfig.UnifiedAuditLogIngestionEnabled
        AdminAuditLogEnabled            = $tenantAuditConfig.AdminAuditLogEnabled
        AuditLogAgeLimit                = $tenantAuditConfig.AdminAuditLogAgeLimit
    } | Export-Csv (Join-Path $outputDir 'Exchange_AuditConfig.csv') -NoTypeInformation -Encoding UTF8
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
    $resourceReport | Export-Csv (Join-Path $outputDir 'Exchange_ResourceMailboxes.csv') -NoTypeInformation -Encoding UTF8
    Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Exporting room and equipment mailboxes..." -CurrentOperation "Saved: Exchange_ResourceMailboxes.csv" -PercentComplete ([int]($step / $totalSteps * 100))


    # === 12. Outbound Spam Auto-Forward Policy ===
    $step++
    Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Checking outbound spam auto-forwarding policy..." -PercentComplete ([int]($step / $totalSteps * 100))

    Get-HostedOutboundSpamFilterPolicy |
        Select-Object -Property Name, AutoForwardingMode |
        Export-Csv (Join-Path $outputDir 'Exchange_OutboundSpamAutoForward.csv') -NoTypeInformation -Encoding UTF8
    Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Checking outbound spam auto-forwarding policy..." -CurrentOperation "Saved: Exchange_OutboundSpamAutoForward.csv" -PercentComplete ([int]($step / $totalSteps * 100))


    # === 13. Shared Mailbox Sign-In Status ===
    $step++
    Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Checking shared mailbox sign-in status..." -PercentComplete ([int]($step / $totalSteps * 100))

    Get-User -RecipientTypeDetails SharedMailbox -ResultSize Unlimited |
        Select-Object -Property DisplayName, UserPrincipalName, AccountDisabled |
        Export-Csv (Join-Path $outputDir 'Exchange_SharedMailboxSignIn.csv') -NoTypeInformation -Encoding UTF8
    Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Checking shared mailbox sign-in status..." -CurrentOperation "Saved: Exchange_SharedMailboxSignIn.csv" -PercentComplete ([int]($step / $totalSteps * 100))


    # === 14. Safe Attachments (requires Defender for Office 365 P1) ===
    $step++
    Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Checking Safe Attachments policies..." -PercentComplete ([int]($step / $totalSteps * 100))

    try {
        Get-SafeAttachmentPolicy -ErrorAction Stop |
            Select-Object -Property Name, Enable, Action, ActionOnError |
            Export-Csv (Join-Path $outputDir 'Exchange_SafeAttachments.csv') -NoTypeInformation -Encoding UTF8
        Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Checking Safe Attachments policies..." -CurrentOperation "Saved: Exchange_SafeAttachments.csv" -PercentComplete ([int]($step / $totalSteps * 100))
    }
    catch {
        Add-AuditIssue -Severity 'Warning' -Section 'Exchange' -Collector 'Get-SafeAttachmentPolicy' -Description ($_.Exception.Message ?? "$_") -Action 'Check permissions or re-run Setup-365AuditApp.ps1'
        Write-Warning "  Safe Attachments not available — requires Defender for Office 365 P1 or higher"
    }


    # === 15. Safe Links (requires Defender for Office 365 P1) ===
    $step++
    Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Checking Safe Links policies..." -PercentComplete ([int]($step / $totalSteps * 100))

    try {
        Get-SafeLinksPolicy -ErrorAction Stop |
            Select-Object -Property Name, EnableSafeLinksForEmail, EnableSafeLinksForTeams, DisableUrlRewrite, TrackClicks |
            Export-Csv (Join-Path $outputDir 'Exchange_SafeLinks.csv') -NoTypeInformation -Encoding UTF8
        Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Checking Safe Links policies..." -CurrentOperation "Saved: Exchange_SafeLinks.csv" -PercentComplete ([int]($step / $totalSteps * 100))
    }
    catch {
        Add-AuditIssue -Severity 'Warning' -Section 'Exchange' -Collector 'Get-SafeLinksPolicy' -Description ($_.Exception.Message ?? "$_") -Action 'Check permissions or re-run Setup-365AuditApp.ps1'
        Write-Warning "  Safe Links not available — requires Defender for Office 365 P1 or higher"
    }


    # === 16. Mail Connectors ===
    $step++
    Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Collecting mail connectors..." -PercentComplete ([int]($step / $totalSteps * 100))

    $connectorRows = @()
    $connectorRows += @(Get-InboundConnector  -ErrorAction SilentlyContinue) | ForEach-Object {
        [PSCustomObject]@{
            Direction           = 'Inbound'
            Name                = $_.Name
            Enabled             = $_.Enabled
            ConnectorType       = $_.ConnectorType
            ConnectorSource     = $_.ConnectorSource
            SenderDomains       = ($_.SenderDomains -join ', ')
            TlsSenderCertName   = $_.TlsSenderCertificateName
            Comment             = $_.Comment
        }
    }
    $connectorRows += @(Get-OutboundConnector -ErrorAction SilentlyContinue) | ForEach-Object {
        [PSCustomObject]@{
            Direction           = 'Outbound'
            Name                = $_.Name
            Enabled             = $_.Enabled
            ConnectorType       = $_.ConnectorType
            ConnectorSource     = $_.ConnectorSource
            SenderDomains       = ''
            TlsSenderCertName   = ''
            Comment             = $_.Comment
        }
    }

    if ($connectorRows.Count -gt 0) {
        $connectorRows | Export-Csv (Join-Path $outputDir 'Exchange_MailConnectors.csv') -NoTypeInformation -Encoding UTF8
        Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Collecting mail connectors..." -CurrentOperation "Saved: Exchange_MailConnectors.csv ($($connectorRows.Count) connectors)" -PercentComplete ([int]($step / $totalSteps * 100))
    }


    # === 17. Accepted Domains ===
    $step++
    Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Collecting Accepted Domains..." -PercentComplete ([int]($step / $totalSteps * 100))
    try {
        $acceptedDomains = Get-AcceptedDomain -ErrorAction Stop
        $acceptedDomains | Select-Object DomainName, DomainType, Default |
            Export-Csv (Join-Path $outputDir 'Exchange_AcceptedDomains.csv') -NoTypeInformation -Encoding UTF8
        Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Collecting Accepted Domains..." -CurrentOperation "Saved: Exchange_AcceptedDomains.csv ($($acceptedDomains.Count) domains)" -PercentComplete ([int]($step / $totalSteps * 100))
    }
    catch {
        Add-AuditIssue -Severity 'Warning' -Section 'Exchange' -Collector 'Get-AcceptedDomain' -Description ($_.Exception.Message ?? "$_") -Action 'Check permissions or re-run Setup-365AuditApp.ps1'
        Write-Warning "Unable to retrieve accepted domains: $_"
    }


    # === 18. Legacy Auth / Basic Auth Policy ===
    $step++
    Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Collecting Legacy Auth Policy..." -PercentComplete ([int]($step / $totalSteps * 100))
    try {
        $_orgCfg         = Get-OrganizationConfig -ErrorAction Stop
        $_defaultAuthPol = $_orgCfg.DefaultAuthenticationPolicy

        $_authPolicies = Get-AuthenticationPolicy -ErrorAction Stop
        $_authPolicies | ForEach-Object {
            [PSCustomObject]@{
                PolicyName                          = $_.Name
                AllowBasicAuthActiveSync            = $_.AllowBasicAuthActiveSync
                AllowBasicAuthImap                  = $_.AllowBasicAuthImap
                AllowBasicAuthPop                   = $_.AllowBasicAuthPop
                AllowBasicAuthSmtp                  = $_.AllowBasicAuthSmtp
                AllowBasicAuthWebServices           = $_.AllowBasicAuthWebServices
                AllowBasicAuthRpc                   = $_.AllowBasicAuthRpc
                AllowBasicAuthPowerShell            = $_.AllowBasicAuthPowerShell
                AllowBasicAuthOfflineAddressBook    = $_.AllowBasicAuthOfflineAddressBook
                AllowBasicAuthReportingWebServices  = $_.AllowBasicAuthReportingWebServices
                IsDefault                           = ($_.Name -eq $_defaultAuthPol)
                ModernAuthEnabled                   = $_orgCfg.OAuth2ClientProfileEnabled
            }
        } | Export-Csv (Join-Path $outputDir 'Exchange_LegacyAuth.csv') -NoTypeInformation -Encoding UTF8
        Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Collecting Legacy Auth Policy..." -CurrentOperation "Saved: Exchange_LegacyAuth.csv ($($_authPolicies.Count) policies)" -PercentComplete ([int]($step / $totalSteps * 100))
    }
    catch {
        Add-AuditIssue -Severity 'Warning' -Section 'Exchange' -Collector 'Get-AuthenticationPolicy' -Description ($_.Exception.Message ?? "$_") -Action 'Check permissions or re-run Setup-365AuditApp.ps1'
        Write-Warning "Unable to retrieve authentication policies: $_"
    }


    # === 19. Exchange Org Config (SMTP AUTH, Customer Lockbox, Modern Auth) ===
    $step++
    Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Collecting Exchange organisation config..." -PercentComplete ([int]($step / $totalSteps * 100))
    try {
        # Reuse $_orgCfg if already in scope from step 18; otherwise fetch fresh
        if (-not $_orgCfg) { $_orgCfg = Get-OrganizationConfig -ErrorAction Stop }
        [PSCustomObject]@{
            SmtpClientAuthDisabled = $_orgCfg.SmtpClientAuthenticationDisabled
            CustomerLockboxEnabled = $_orgCfg.CustomerLockboxEnabled
            ModernAuthEnabled      = $_orgCfg.OAuth2ClientProfileEnabled
            AuditDisabled          = $_orgCfg.AuditDisabled
            MessageCopyForSent     = $_orgCfg.MessageCopyForSentAsEnabled
        } | Export-Csv (Join-Path $outputDir 'Exchange_OrgConfig.csv') -NoTypeInformation -Encoding UTF8
        Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Collecting Exchange organisation config..." -CurrentOperation "Saved: Exchange_OrgConfig.csv" -PercentComplete ([int]($step / $totalSteps * 100))
    }
    catch {
        Add-AuditIssue -Severity 'Warning' -Section 'Exchange' -Collector 'Get-OrganizationConfig' -Description ($_.Exception.Message ?? "$_") -Action 'Check permissions or re-run Setup-365AuditApp.ps1'
        Write-Warning "Unable to retrieve Exchange organisation config: $_"
    }


    # === 20. External Sender Tagging ===
    $step++
    Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Checking external sender identification..." -PercentComplete ([int]($step / $totalSteps * 100))
    try {
        $extSenderCfg = Get-ExternalInOutlook -ErrorAction Stop
        [PSCustomObject]@{
            Enabled         = $extSenderCfg.Enabled
            AllUsers        = $extSenderCfg.AllUsers
            AllowList       = ($extSenderCfg.AllowList -join '; ')
        } | Export-Csv (Join-Path $outputDir 'Exchange_ExternalSenderTagging.csv') -NoTypeInformation -Encoding UTF8
        Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Checking external sender identification..." -CurrentOperation "Saved: Exchange_ExternalSenderTagging.csv" -PercentComplete ([int]($step / $totalSteps * 100))
    }
    catch {
        if ($_ -match 'server side error|operation could not be completed') {
            Write-Verbose "Get-ExternalInOutlook: server-side error — feature may not be available for this tenant. Skipping."
        } else {
            Add-AuditIssue -Severity 'Warning' -Section 'Exchange' -Collector 'Get-ExternalInOutlook' -Description ($_.Exception.Message ?? "$_") -Action 'Check permissions or re-run Setup-365AuditApp.ps1'
            Write-Warning "Unable to retrieve external sender tagging settings: $_"
        }
    }


    # === 21. Connection Filter Policy (IP Allow List) ===
    $step++
    Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Checking connection filter policies..." -PercentComplete ([int]($step / $totalSteps * 100))
    try {
        $connFilters = @(Get-HostedConnectionFilterPolicy -ErrorAction Stop)
        $connFilters | ForEach-Object {
            [PSCustomObject]@{
                PolicyName    = $_.Name
                IsDefault     = $_.IsDefault
                IPAllowList   = ($_.IPAllowList -join '; ')
                IPBlockList   = ($_.IPBlockList -join '; ')
                EnableSafeList = $_.EnableSafeList
            }
        } | Export-Csv (Join-Path $outputDir 'Exchange_ConnectionFilter.csv') -NoTypeInformation -Encoding UTF8
        Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Checking connection filter policies..." -CurrentOperation "Saved: Exchange_ConnectionFilter.csv ($($connFilters.Count) policies)" -PercentComplete ([int]($step / $totalSteps * 100))
    }
    catch {
        Add-AuditIssue -Severity 'Warning' -Section 'Exchange' -Collector 'Get-HostedConnectionFilterPolicy' -Description ($_.Exception.Message ?? "$_") -Action 'Check permissions or re-run Setup-365AuditApp.ps1'
        Write-Warning "Unable to retrieve connection filter policies: $_"
    }


    # === 22. OWA Mailbox Policy ===
    $step++
    Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Checking OWA mailbox policies..." -PercentComplete ([int]($step / $totalSteps * 100))
    try {
        $owaPolicies = @(Get-OwaMailboxPolicy -ErrorAction Stop)
        $owaPolicies | ForEach-Object {
            [PSCustomObject]@{
                PolicyName                        = $_.Name
                AdditionalStorageProvidersAvailable = $_.AdditionalStorageProvidersAvailable
                ClassicAttachmentsEnabled         = $_.ClassicAttachmentsEnabled
                OneDriveAttachmentsEnabled        = $_.OneDriveAttachmentsEnabled
                ThirdPartyAttachmentsEnabled      = $_.ThirdPartyAttachmentsEnabled
                PersonalAccountCalendarsEnabled   = $_.PersonalAccountCalendarsEnabled
            }
        } | Export-Csv (Join-Path $outputDir 'Exchange_OwaPolicy.csv') -NoTypeInformation -Encoding UTF8
        Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Checking OWA mailbox policies..." -CurrentOperation "Saved: Exchange_OwaPolicy.csv ($($owaPolicies.Count) policies)" -PercentComplete ([int]($step / $totalSteps * 100))
    }
    catch {
        Add-AuditIssue -Severity 'Warning' -Section 'Exchange' -Collector 'Get-OwaMailboxPolicy' -Description ($_.Exception.Message ?? "$_") -Action 'Check permissions or re-run Setup-365AuditApp.ps1'
        Write-Warning "Unable to retrieve OWA mailbox policies: $_"
    }


    # ── Done ───────────────────────────────────────────────────────────────────
    Write-Progress -Id 1 -Activity $activity -Completed
    Write-Host "Exchange Inventory complete. Results saved to: $outputDir" -ForegroundColor Green

    return @{
        MailboxCount       = @($mailboxes).Count
        SharedMailboxCount = @($mailboxes | Where-Object { $_.RecipientTypeDetails -eq 'SharedMailbox' }).Count
        UserMailboxCount   = @($mailboxes | Where-Object { $_.RecipientTypeDetails -eq 'UserMailbox' }).Count
        DistributionLists  = @($dlReport).Count
        ForwardingRules    = @($inboxRules).Count
        BrokenInboxRules   = $brokenInboxRules.Count
    }
}
