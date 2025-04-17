param (
    [string]$AuditFolder,
    [switch]$DevMode = $false
)

if (-not $DevMode -and $MyInvocation.InvocationName -eq $MyInvocation.MyCommand.Name) {
    Write-Error "This script must be run from the 365Audit launcher. Use -DevMode for development." -ErrorAction Stop
}


if (-not $DevMode -and $MyInvocation.InvocationName -eq $MyInvocation.MyCommand.Name) {
    Write-Error "This script must be run from the 365Audit launcher. Use -DevMode for development." -ErrorAction Stop
}

<#
.SYNOPSIS
    Performs an Exchange Online audit of user and shared mailboxes, permissions, inbox rules, and distribution groups.

.DESCRIPTION
    This script connects to Exchange Online via EXO V2 module and generates the following reports:
    - Mailbox inventory (user & shared)
    - Archive status and mailbox sizes
    - Mailbox permissions (Full Access, Send As, Send on Behalf)
    - Distribution Lists and dynamic rules
    - Inbox rules including internal/external forwarding

    Results are saved in CSV format to a timestamped folder named after the organization.

.NOTES
    Author      : Raymond Slater
    Version     : 1.0.1
    Change Log  :
        1.0.0 - Initial release
        1.0.1 - Refactor Output directory initialisation
        1.0.2 - Helper function refactor
        
.LINK
    https://github.com/razer86/365Audit
#>

# === Retrieve Output Folder ===
try {
    $context = Initialize-AuditOutput
    $outputDir = $context.OutputPath
}
catch {
    Write-Error "❌ Failed to locate audit output directory: $_"
    exit 1
}

# === 1. Ensure Exchange Online module is available ===
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Write-Host "📦 Installing ExchangeOnlineManagement module..." -ForegroundColor Yellow
    Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force
}
Import-Module ExchangeOnlineManagement

# === 2. Connect to Exchange Online if not already ===
if (-not (Get-ConnectionInformation)) {
    Write-Host "🔐 Connecting to Exchange Online..."
    Connect-ExchangeOnline -ShowBanner:$false
}

# === 3. Create output folder ===
$org = Get-OrganizationConfig
$companyName = $org.Name -replace '[^a-zA-Z0-9]', '_'
$timestamp = Get-Date -Format "yyyyMMdd-HHmm"
$outputDir = "${companyName}_ExchangeAudit_$timestamp"
New-Item -ItemType Directory -Path $outputDir -Force | Out-Null

Write-Host "`n📬 Starting Exchange Audit for $companyName..." -ForegroundColor Cyan

# === 4. Mailbox Inventory ===
Write-Host "✔ Gathering mailbox inventory..."

$mailboxes = Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox, SharedMailbox
$mailboxInventory = foreach ($mbx in $mailboxes) {
    $stats = Get-MailboxStatistics -Identity $mbx.Identity
    [PSCustomObject]@{
        DisplayName   = $mbx.DisplayName
        UserPrincipalName = $mbx.UserPrincipalName
        RecipientType = $mbx.RecipientTypeDetails
        ArchiveEnabled = $mbx.ArchiveStatus -eq "Active"
        TotalSizeMB   = [math]::Round(($stats.TotalItemSize.Value.ToMB()), 2)
        ItemCount     = $stats.ItemCount
    }
}
$mailboxInventory | Export-Csv "$outputDir\Exchange_Mailboxes.csv" -NoTypeInformation

# === 5. Mailbox Permissions ===
Write-Host "✔ Collecting mailbox permissions..."

# Full Access
$fullAccess = Get-MailboxPermission -ResultSize Unlimited | Where-Object { -not $_.IsInherited }
$fullAccess | Select Identity, User, AccessRights | Export-Csv "$outputDir\Exchange_Permissions_FullAccess.csv" -NoTypeInformation

# Send As
$sendAs = Get-RecipientPermission -ResultSize Unlimited | Where-Object { $_.AccessRights -contains "SendAs" }
$sendAs | Select Identity, Trustee, AccessRights | Export-Csv "$outputDir\Exchange_Permissions_SendAs.csv" -NoTypeInformation

# Send on Behalf
$sendOnBehalf = @()
foreach ($mbx in $mailboxes) {
    if ($mbx.GrantSendOnBehalfTo.Count -gt 0) {
        foreach ($delegate in $mbx.GrantSendOnBehalfTo) {
            $sendOnBehalf += [PSCustomObject]@{
                Mailbox = $mbx.DisplayName
                Delegate = $delegate.Name
            }
        }
    }
}
$sendOnBehalf | Export-Csv "$outputDir\Exchange_Permissions_SendOnBehalf.csv" -NoTypeInformation

# === 6. Distribution Lists ===
Write-Host "✔ Gathering distribution list details..."

$dlGroups = Get-DistributionGroup -ResultSize Unlimited
$dlReport = foreach ($dl in $dlGroups) {
    $members = Get-DistributionGroupMember -Identity $dl.Identity -ErrorAction SilentlyContinue
    $isDynamic = $dl.RecipientTypeDetails -eq 'DynamicDistributionGroup'
    [PSCustomObject]@{
        DisplayName     = $dl.DisplayName
        EmailAddress    = $dl.PrimarySmtpAddress
        GroupType       = if ($isDynamic) { "Dynamic" } else { "Static" }
        MemberCount     = $members.Count
        Members         = ($members.DisplayName -join "; ")
        MembershipRule  = if ($isDynamic) { $dl.RecipientFilter } else { "N/A" }
    }
}
$dlReport | Export-Csv "$outputDir\Exchange_DistributionLists.csv" -NoTypeInformation

# === 7. Inbox Rules & Forwarding ===
Write-Host "✔ Checking inbox rules for forwarding..."

$inboxRules = @()
foreach ($mbx in $mailboxes) {
    try {
        $rules = Get-InboxRule -Mailbox $mbx.UserPrincipalName
        foreach ($rule in $rules) {
            $fwdTo = $rule.ForwardTo | ForEach-Object { $_.Name }
            $fwdCc = $rule.ForwardAsAttachmentTo | ForEach-Object { $_.Name }
            $redirectTo = $rule.RedirectTo | ForEach-Object { $_.Name }

            if ($fwdTo -or $redirectTo -or $fwdCc) {
                $inboxRules += [PSCustomObject]@{
                    Mailbox       = $mbx.DisplayName
                    RuleName      = $rule.Name
                    ForwardTo     = ($fwdTo -join "; ")
                    RedirectTo    = ($redirectTo -join "; ")
                    ForwardCc     = ($fwdCc -join "; ")
                    ExternalForwarding = if ($rule.ForwardTo -or $rule.RedirectTo) {
                        $true  # Could later enhance to check domain
                    } else {
                        $false
                    }
                }
            }
        }
    }
    catch {
        Write-Warning "Failed to get inbox rules for $($mbx.DisplayName)"
    }
}
$inboxRules | Export-Csv "$outputDir\Exchange_InboxForwardingRules.csv" -NoTypeInformation

# === 8. Mail Flow Rules (Transport Rules) ===
Write-Host "✔ Exporting transport (mail flow) rules..."

$transportRules = Get-TransportRule
$transportRules | Select Name, Priority, State, Mode, FromAddressContainsWords, SentTo, RedirectMessageTo, BlindCopyTo, ApplyHtmlDisclaimerLocation, ApplyHtmlDisclaimerText | Export-Csv "$outputDir\Exchange_TransportRules.csv" -NoTypeInformation


# === 9. External Forwarding Global Settings ===
Write-Host "✔ Checking external forwarding settings..."

$remoteDomains = Get-RemoteDomain
$forwardingAllowed = $remoteDomains | Select DomainName, AutoForwardEnabled
$forwardingAllowed | Export-Csv "$outputDir\Exchange_RemoteDomainForwarding.csv" -NoTypeInformation


# === 10. Anti-Phishing Policies ===
Write-Host "✔ Exporting anti-phish policies..."

$phishPolicies = Get-AntiPhishPolicy
$phishPolicies | Select Name, EnableTargetedUserProtection, EnableMailboxIntelligence, EnableSpoofIntelligence, EnableATPForSpoof | Export-Csv "$outputDir\Exchange_AntiPhishPolicies.csv" -NoTypeInformation


# === 11. Anti-Spam / Malware Policies ===
Write-Host "✔ Exporting spam and malware filter policies..."

$spamPolicies = Get-HostedContentFilterPolicy
$spamPolicies | Select Name, SpamAction, HighConfidenceSpamAction, BulkSpamAction | Export-Csv "$outputDir\Exchange_SpamPolicies.csv" -NoTypeInformation

$malwarePolicies = Get-MalwareFilterPolicy
$malwarePolicies | Select Name, Action, EnableExternalSenderAdminNotification | Export-Csv "$outputDir\Exchange_MalwarePolicies.csv" -NoTypeInformation


# === 12. DKIM / DMARC / SPF Status ===
Write-Host "✔ Checking DKIM / DMARC / SPF records..."

$acceptedDomains = Get-AcceptedDomain
$dkimStatus = @()

foreach ($domain in $acceptedDomains) {
    try {
        $dkimConfig = Get-DkimSigningConfig -Identity $domain.DomainName -ErrorAction Stop
        $dkimStatus += [PSCustomObject]@{
            Domain         = $domain.DomainName
            DKIMEnabled    = $dkimConfig.Enabled
            Selector1CNAME = $dkimConfig.Selector1CNAME
            Selector2CNAME = $dkimConfig.Selector2CNAME
        }
    }
    catch {
        $dkimStatus += [PSCustomObject]@{
            Domain         = $domain.DomainName
            DKIMEnabled    = "Not Configured"
            Selector1CNAME = "N/A"
            Selector2CNAME = "N/A"
        }
    }
}

$dkimStatus | Export-Csv "$outputDir\Exchange_DKIM_Status.csv" -NoTypeInformation

# Optional: use Resolve-DnsName for SPF/DMARC if needed


# === 13. Mailbox Audit Settings ===
Write-Host "✔ Checking mailbox audit settings..."

$mailboxAuditing = Get-Mailbox -ResultSize Unlimited | Select DisplayName, UserPrincipalName, AuditEnabled
$mailboxAuditing | Export-Csv "$outputDir\Exchange_MailboxAuditStatus.csv" -NoTypeInformation

$tenantAuditConfig = Get-AdminAuditLogConfig
[PSCustomObject]@{
    UnifiedAuditLogIngestionEnabled = $tenantAuditConfig.UnifiedAuditLogIngestionEnabled
    AdminAuditLogEnabled            = $tenantAuditConfig.AdminAuditLogEnabled
    AuditLogAgeLimit                = $tenantAuditConfig.AdminAuditLogAgeLimit
} | Export-Csv "$outputDir\Exchange_AuditConfig.csv" -NoTypeInformation


# === 14. Receive Connectors & Anonymous Relay ===
Write-Host "✔ Checking for anonymous relay connectors..."

$receiveConnectors = Get-ReceiveConnector -ErrorAction SilentlyContinue
$anonymousRelay = @()

foreach ($rc in $receiveConnectors) {
    if ($rc.PermissionGroups -contains "AnonymousUsers") {
        $anonymousRelay += [PSCustomObject]@{
            Server     = $rc.Server
            Name       = $rc.Name
            AuthGroups = $rc.PermissionGroups -join ", "
        }
    }
}
$anonymousRelay | Export-Csv "$outputDir\Exchange_AnonymousRelayConnectors.csv" -NoTypeInformation


# === 15. Resource Mailboxes (Room & Equipment) ===
Write-Host "✔ Exporting room and equipment mailbox info..."

$resources = Get-Mailbox -RecipientTypeDetails RoomMailbox, EquipmentMailbox -ResultSize Unlimited
$resourceReport = foreach ($res in $resources) {
    $calendarSettings = Get-MailboxCalendarConfiguration -Identity $res.Identity -ErrorAction SilentlyContinue
    [PSCustomObject]@{
        DisplayName        = $res.DisplayName
        ResourceType       = $res.RecipientTypeDetails
        Email              = $res.PrimarySmtpAddress
        BookingWindowDays  = $calendarSettings.BookingWindowInDays
        AllowConflicts     = $calendarSettings.AllowConflicts
        BookingDelegates   = ($calendarSettings.ResourceDelegates -join "; ")
    }
}
$resourceReport | Export-Csv "$outputDir\Exchange_ResourceMailboxes.csv" -NoTypeInformation

































Write-Host "`n✅ Exchange audit complete. Results saved to: $outputDir"