Describe "NeConnect MSP Checks" -Tag "NeConnect", "MSP", "Custom" {

    It "NC.001: Technical contacts should use MSP domains" -Tag "Severity:Medium", "Entra" {
        $ctx = Get-AuditContext -NoThrow
        $mspDomains = if ($ctx) { $ctx.MspDomains } else { @() }

        if (-not $mspDomains -or $mspDomains.Count -eq 0) {
            Set-ItResult -Skipped -Because "No MSP domains configured (-MspDomains parameter)"
            return
        }

        $org = Get-MgOrganization | Select-Object -First 1
        $techContacts = @($org.TechnicalNotificationMails | Where-Object { $_ })

        if ($techContacts.Count -eq 0) {
            Set-ItResult -Skipped -Because "No technical notification email addresses configured"
            return
        }

        $nonMsp = @($techContacts | Where-Object {
            $domain = ($_ -split '@')[1]
            $domain -and $domain -notin $mspDomains
        })

        $nonMsp.Count | Should -Be 0 -Because "all technical contacts should use MSP domains ($($mspDomains -join ', ')). Found non-MSP: $($nonMsp -join ', ')"
    }

    It "NC.002: Global Administrator count should be between 2 and 4" -Tag "Severity:High", "Entra" {
        $gaRole = Get-MgDirectoryRole -Filter "displayName eq 'Global Administrator'" -ErrorAction SilentlyContinue
        if (-not $gaRole) {
            Set-ItResult -Skipped -Because "Global Administrator role not found"
            return
        }

        $members = @(Get-MgDirectoryRoleMember -DirectoryRoleId $gaRole.Id -ErrorAction Stop)
        $members.Count | Should -BeGreaterOrEqual 2 -Because "at least 2 Global Administrators should exist for break-glass resilience"
        $members.Count | Should -BeLessOrEqual 4 -Because "Microsoft recommends no more than 4 Global Administrators"
    }

    It "NC.003: No mailbox forwarding rules to external domains" -Tag "Severity:High", "Exchange" {
        $_exoConnected = Get-ConnectionInformation -ErrorAction SilentlyContinue |
            Where-Object { $_.State -eq 'Connected' }
        if (-not $_exoConnected) {
            Set-ItResult -Skipped -Because "Not connected to Exchange Online"
            return
        }

        $forwardingMailboxes = @(Get-Mailbox -ResultSize Unlimited -ErrorAction Stop |
            Where-Object { $_.ForwardingSmtpAddress -and $_.ForwardingSmtpAddress -notmatch '\.onmicrosoft\.com$' })

        $forwardingMailboxes.Count | Should -Be 0 -Because "no mailboxes should have external forwarding configured. Found: $($forwardingMailboxes.PrimarySmtpAddress -join ', ')"
    }

    It "NC.004: Shared mailboxes should have sign-in blocked" -Tag "Severity:Medium", "Exchange" {
        $_exoConnected = Get-ConnectionInformation -ErrorAction SilentlyContinue |
            Where-Object { $_.State -eq 'Connected' }
        if (-not $_exoConnected) {
            Set-ItResult -Skipped -Because "Not connected to Exchange Online"
            return
        }

        $sharedMailboxes = @(Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited -ErrorAction Stop)
        if ($sharedMailboxes.Count -eq 0) {
            Set-ItResult -Skipped -Because "No shared mailboxes found"
            return
        }

        $enabledShared = @($sharedMailboxes | ForEach-Object {
            $user = Get-MgUser -UserId $_.ExternalDirectoryObjectId -Property AccountEnabled -ErrorAction SilentlyContinue
            if ($user -and $user.AccountEnabled) { $_.PrimarySmtpAddress }
        } | Where-Object { $_ })

        $enabledShared.Count | Should -Be 0 -Because "shared mailboxes should have sign-in blocked. Found sign-in enabled: $($enabledShared -join ', ')"
    }

    It "NC.005: Microsoft 365 Lighthouse should be enrolled" -Tag "Severity:Low", "Entra" {
        $_lighthouseAppId = '2828a423-d5cc-4818-8285-aa945d95017a'
        $sp = Get-MgServicePrincipal -Filter "appId eq '$_lighthouseAppId'" -Property Id,DisplayName -ErrorAction SilentlyContinue | Select-Object -First 1

        $sp | Should -Not -BeNullOrEmpty -Because "the tenant should be enrolled in Microsoft 365 Lighthouse for MSP management"
    }

    It "NC.006: External sender tagging should be enabled" -Tag "Severity:Medium", "Exchange" {
        $_exoConnected = Get-ConnectionInformation -ErrorAction SilentlyContinue |
            Where-Object { $_.State -eq 'Connected' }
        if (-not $_exoConnected) {
            Set-ItResult -Skipped -Because "Not connected to Exchange Online"
            return
        }

        $config = Get-ExternalInOutlook -ErrorAction Stop
        $config.Enabled | Should -Be $true -Because "external sender tagging should be enabled to help users identify emails from outside the organisation"
    }

    It "NC.007: Audit log search should be enabled" -Tag "Severity:High", "Exchange" {
        $_exoConnected = Get-ConnectionInformation -ErrorAction SilentlyContinue |
            Where-Object { $_.State -eq 'Connected' }
        if (-not $_exoConnected) {
            Set-ItResult -Skipped -Because "Not connected to Exchange Online"
            return
        }

        $adminAudit = Get-AdminAuditLogConfig -ErrorAction Stop
        $adminAudit.UnifiedAuditLogIngestionEnabled | Should -Be $true -Because "unified audit log ingestion must be enabled for security monitoring and incident response"
    }
}
