# Invoke-ExchangeAudit.ps1

Connects to Exchange Online and audits mailboxes, permissions, and mail flow.

Covers user and shared mailboxes, delegated permissions, distribution lists, inbox and transport rules, external forwarding, anti-spam/phish/malware policies, Safe Attachments, Safe Links, DKIM, mailbox audit configuration, and resource mailboxes. Defender for Office 365 features (Safe Attachments, Safe Links) are skipped gracefully when not licensed.

## Required Permissions

Granted automatically by `Setup-365AuditApp.ps1`:

| Permission | Type | Purpose |
|---|---|---|
| `Exchange.ManageAsApp` | Application | All Exchange Online operations via app-only auth |
| Exchange Administrator | Directory Role | Required alongside `Exchange.ManageAsApp` |

## Output Files

| File | Description |
|---|---|
| `Exchange_Mailboxes.csv` | User and shared mailboxes with size, item count, archive status, and litigation hold |
| `Exchange_Permissions_FullAccess.csv` | Non-inherited Full Access mailbox permissions |
| `Exchange_Permissions_SendAs.csv` | Send As delegated permissions |
| `Exchange_Permissions_SendOnBehalf.csv` | Send on Behalf delegations |
| `Exchange_DistributionLists.csv` | Distribution groups with member count, type, and filter rules |
| `Exchange_InboxForwardingRules.csv` | Inbox rules that forward or redirect mail |
| `Exchange_BrokenInboxRules.csv` | Inbox rules in a broken/non-functional state |
| `Exchange_TransportRules.csv` | Mail flow (transport) rule summaries |
| `Exchange_RemoteDomainForwarding.csv` | Auto-forward enabled flag per remote domain |
| `Exchange_OutboundSpamAutoForward.csv` | Auto-forward mode per outbound spam filter policy |
| `Exchange_SharedMailboxSignIn.csv` | Shared mailboxes with interactive sign-in enabled |
| `Exchange_AntiPhishPolicies.csv` | Anti-phishing policy configuration |
| `Exchange_SpamPolicies.csv` | Hosted content filter (anti-spam) policies |
| `Exchange_MalwarePolicies.csv` | Malware filter policies |
| `Exchange_SafeAttachments.csv` | Safe Attachments policies (requires Defender for Office 365 P1; absent if unlicensed) |
| `Exchange_SafeLinks.csv` | Safe Links policies (requires Defender for Office 365 P1; absent if unlicensed) |
| `Exchange_DKIM_Status.csv` | DKIM signing configuration and CNAME selectors per domain |
| `Exchange_MailboxAuditStatus.csv` | Per-mailbox audit enabled flag |
| `Exchange_AuditConfig.csv` | Tenant unified audit log and admin audit log settings |
| `Exchange_AnonymousRelayConnectors.csv` | Receive connectors permitting anonymous relay |
| `Exchange_ResourceMailboxes.csv` | Room and equipment mailboxes with booking settings |
