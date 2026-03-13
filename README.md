# 365Audit

Monthly Microsoft 365 audit toolkit for MSP maintenance reporting. Run per customer to generate a report covering identity, messaging, and storage health.

---

## First-Time Setup

Before running the toolkit for any customer, run `Setup-365AuditApp.ps1` once in that tenant as a **Global Administrator**:

```powershell
.\Setup-365AuditApp.ps1
```

The script will:
1. Create an app registration with the required Microsoft Graph and Exchange permissions, and grant admin consent
2. Create a second app registration for SharePoint interactive authentication (PnP)
3. Print all credentials to the terminal — **store these in Hudu** immediately

**Credentials to store per customer:**

| Field | Used For |
|---|---|
| App ID (Client ID) | Entra and Exchange silent authentication |
| Tenant ID | All modules |
| Client Secret | Entra and Exchange silent authentication |
| Secret Expiry | Reminder to rotate before expiry |
| PnP App ID | SharePoint interactive authentication |

**Example launch command (printed by Setup script):**

```powershell
.\Start-365Audit.ps1 -AppId '<AppId>' -AppSecret '<secret>' -TenantId '<TenantId>' -PnPAppId '<PnPAppId>'
```

> `-PnPAppId` is optional. Omitting it skips the SharePoint Online audit.

---

## Requirements

### PowerShell Version

- **7.4 or later** — required for the SharePoint module (PnP.PowerShell v3)
- **7.2 or later** — minimum for Entra, Exchange, and Mail Security modules only

Download: https://github.com/PowerShell/PowerShell/releases

### Required Modules

| Module | Required By | Install |
|---|---|---|
| `Microsoft.Graph` | All modules | `Install-Module Microsoft.Graph -Scope CurrentUser` |
| `ExchangeOnlineManagement` | Exchange, Mail Security | `Install-Module ExchangeOnlineManagement -Scope CurrentUser` |
| `PnP.PowerShell` | SharePoint | `Install-Module PnP.PowerShell -Scope CurrentUser` |

Modules are checked at runtime and installed automatically if missing.

---

## Usage

Open a PowerShell 7.4+ terminal, navigate to the toolkit directory, and run with credentials from `Setup-365AuditApp.ps1`:

```powershell
.\Start-365Audit.ps1 -AppId '<AppId>' -AppSecret '<AppSecret>' -TenantId '<TenantId>' -PnPAppId '<PnPAppId>'
```

On launch the toolkit will:
1. Check local script versions against the GitHub version manifest and warn if updates are available
2. Display a notice if `-PnPAppId` was not provided (SharePoint audit will be skipped)
3. Present the module selection menu
4. Connect silently to Microsoft Graph and Exchange Online on first module run (no browser prompt)
5. Create a per-customer output folder: `<CompanyName>_<yyyyMMdd>/`

> SharePoint uses an interactive browser sign-in once per session. After the first sign-in, the MSAL token is reused across all site connections — no repeated prompts.

Select one or more modules by number (comma-separated, e.g. `1,2,3`).

---

## Menu

| Option | Module | Description |
|---|---|---|
| 1 | Microsoft Entra Audit | Identity, MFA, roles, Conditional Access |
| 2 | Exchange Online Audit | Mailboxes, permissions, mail flow |
| 3 | SharePoint Online Audit | Sites, permissions, storage, OneDrive *(requires -PnPAppId)* |
| 4 | Mail Security Audit | DKIM, DMARC, SPF, anti-spam/phish policies |
| 9 | Run All (1, 2, 3, 4 + summary) | Full audit and summary in one pass |
| 0 | Exit | — |

> Each module option (1–4 and 9) automatically regenerates the HTML summary report on completion.

---

## Module Reference

### Invoke-EntraAudit.ps1

Connects to Microsoft Graph and audits Entra ID (Azure Active Directory).

**Output files:**

| File | Description |
|---|---|
| `Entra_Users.csv` | UPN, name, licence, MFA status and methods, password policy, last sign-in |
| `Entra_Licenses.csv` | Subscriptions: total, consumed, suspended, and warning seat counts |
| `Entra_SSPR.csv` | Self-Service Password Reset enforcement status |
| `Entra_AdminRoles.csv` | All directory role assignments |
| `Entra_GlobalAdmins.csv` | Subset of Global Administrator accounts |
| `Entra_GuestUsers.csv` | Guest accounts with creation date and last sign-in |
| `Entra_Groups.csv` | All groups: type, membership rule, owners, and members |
| `Entra_CA_Policies.csv` | Conditional Access policy names, states, targets, client app types, and grant controls |
| `Entra_TrustedLocations.csv` | Named locations with trusted flag and IP ranges |
| `Entra_SecurityDefaults.csv` | Whether Security Defaults are enabled |

---

### Invoke-ExchangeAudit.ps1

Connects to Exchange Online and audits mailboxes, permissions, and mail flow.

**Output files:**

| File | Description |
|---|---|
| `Exchange_Mailboxes.csv` | User and shared mailboxes with size, item count, archive status, and litigation hold |
| `Exchange_Permissions_FullAccess.csv` | Non-inherited Full Access mailbox permissions |
| `Exchange_Permissions_SendAs.csv` | Send As delegated permissions |
| `Exchange_Permissions_SendOnBehalf.csv` | Send on Behalf delegations |
| `Exchange_DistributionLists.csv` | Distribution groups with member count, type, and filter rules |
| `Exchange_InboxForwardingRules.csv` | Inbox rules that forward or redirect mail |
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

---

### Invoke-SharePointAudit.ps1

Connects to SharePoint Online via PnP.PowerShell and audits sites and OneDrive.

> Requires PowerShell 7.4+ and the PnP.PowerShell module. Must be launched with `-PnPAppId`.

**Output files:**

| File | Description |
|---|---|
| `SharePoint_TenantStorage.csv` | Total, used, and available storage across the tenant |
| `SharePoint_Sites.csv` | Site collections with URL, template, storage, and owner |
| `SharePoint_SPGroups.csv` | SharePoint groups with owners and members per site |
| `SharePoint_SitePermissions.csv` | Role assignments per site |
| `SharePoint_ExternalSharing_Tenant.csv` | Tenant-wide external sharing policy, default link type, and anonymous link expiry |
| `SharePoint_ExternalSharing_SiteOverrides.csv` | Sites that override the tenant sharing setting |
| `SharePoint_OneDriveUsage.csv` | Per-user OneDrive storage consumption |
| `SharePoint_AccessControlPolicies.csv` | Idle session timeout, IP restrictions, Conditional Access, unmanaged device sync restriction, and legacy Mac sync settings |
| `SharePoint_OneDrive_Unlicensed.csv` | OneDrive accounts belonging to users without an active licence |

---

### Invoke-MailSecurityAudit.ps1

Connects to Exchange Online and audits mail security configuration.

Flat data is exported as CSV for the HTML summary. Nested policy objects are exported as JSON for detailed review.

**CSV output:**

| File | Description |
|---|---|
| `MailSec_DKIM.csv` | DKIM signing status and selector CNAMEs per domain |
| `MailSec_DMARC.csv` | DMARC TXT records per accepted domain |
| `MailSec_SPF.csv` | SPF TXT records per accepted domain |

**JSON output (supplementary):**

| File | Description |
|---|---|
| `MailSec_AntiSpam.json` | Hosted content filter policy objects |
| `MailSec_AntiSpamRules.json` | Hosted content filter rules |
| `MailSec_AntiPhish.json` | Anti-phishing policy objects |
| `MailSec_AntiPhishRules.json` | Anti-phishing rules |
| `MailSec_SpoofIntelligence.json` | Spoof intelligence insights (requires Defender for Office 365) |
| `MailSec_InboundConnectors.json` | Inbound connector configuration |
| `MailSec_OutboundConnectors.json` | Outbound connector configuration |
| `MailSec_TransportRules.json` | Mail flow rule summaries |

---

### Generate-AuditSummary.ps1

Reads CSV files from the current audit run and compiles them into a single HTML report (`M365_AuditSummary.html`), which opens automatically in the default browser.

**Action Items**

The top of the report shows a prioritised list of findings requiring attention:

| Badge | Category | Example findings |
|---|---|---|
| Critical | Entra / Auth | Legacy authentication not blocked by Conditional Access |
| Warning | Entra / Accounts | Licensed users with no sign-in for 90+ days |
| Warning | Entra / Guests | Guest accounts inactive for 90+ days |
| Warning | Exchange | Shared mailboxes with interactive sign-in enabled |
| Critical | Exchange | Outbound spam policy allows unrestricted auto-forwarding |
| Warning | Exchange | No Safe Attachments or Safe Links policy enabled |
| Warning | SharePoint | Default sharing link allows anonymous (anyone) access |
| Warning | SharePoint | OneDrive sync not restricted to managed devices |

**Report sections:**

- **Microsoft Entra** — MFA coverage, stale licensed accounts, licence table, SSPR status, Security Defaults, global admin count, role summary, guest accounts and stale guest count, CA policies, legacy auth check
- **Exchange Online** — Mailbox count, external forwarding rule alerts, shared mailbox sign-in status, outbound spam auto-forward policy, Safe Attachments and Safe Links status
- **SharePoint / OneDrive** — Tenant storage gauge, site collection table with expandable groups panel, external sharing policy and site overrides, access control policies, OneDrive usage and unlicensed accounts
- **Mail Security** — DKIM, DMARC, and SPF coverage per domain

---

## Output Structure

All module output lands in a single folder created at the start of each session:

```
<repo root>/
└── <CompanyName>_<yyyyMMdd>/
    ├── OrgInfo.json
    ├── Entra_Users.csv
    ├── Entra_Licenses.csv
    ├── ... (all module CSVs and JSON files)
    └── M365_AuditSummary.html
```

The folder name is derived from the Entra organisation display name (alphanumeric only) and the current date. Running the toolkit multiple times on the same day reuses the same folder.

---

## Version Check

On each launch, the toolkit downloads `version.json` from GitHub and compares it against the `$ScriptVersion` declared in each local script. Outdated scripts are listed by name with the installed and latest versions.

The check is non-blocking — a network failure produces a warning and the toolkit continues normally.

---

## Development

To run a module directly without the launcher (bypasses the guard clause):

```powershell
.\Invoke-EntraAudit.ps1 -DevMode
```

All modules accept the `-DevMode` switch for standalone testing.

---

## File Structure

```
365Audit/
├── Start-365Audit.ps1           # Launcher and menu
├── Setup-365AuditApp.ps1        # One-time app registration setup
├── Invoke-EntraAudit.ps1        # Entra ID module
├── Invoke-ExchangeAudit.ps1     # Exchange Online module
├── Invoke-SharePointAudit.ps1   # SharePoint / OneDrive module
├── Invoke-MailSecurityAudit.ps1 # Mail security module
├── Generate-AuditSummary.ps1    # HTML report generator
├── version.json                 # GitHub version manifest
└── common/
    └── Audit-Common.ps1         # Shared helpers and version check
```
