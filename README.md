# NeConnect Microsoft 365 Audit Toolkit

Monthly Microsoft 365 audit toolkit for MSP maintenance reporting. Run per customer to generate a report covering identity, messaging, and storage health.

---

## First-Time Setup

Before running the toolkit for any customer, run `Setup-365AuditApp.ps1` once in that tenant as a **Global Administrator**.

The script will:
1. Create an app registration with all required Microsoft Graph, Exchange Online, and SharePoint permissions and grant admin consent
2. Generate a self-signed certificate, upload the public key to the app registration
3. Print all credentials to the terminal
4. **Automatically push credentials to Hudu** if `HUDU_API_KEY` is set in your environment

### With Hudu integration (recommended)

If your Hudu environment variables are configured (see [Hudu Integration](#hudu-integration)), pass the company slug or ID and credentials are stored automatically:

```powershell
.\Setup-365AuditApp.ps1 -HuduCompanyId '<company-slug>'
.\Setup-365AuditApp.ps1 -HuduCompanyName 'Contoso Ltd'
```

### Without Hudu integration

```powershell
.\Setup-365AuditApp.ps1
```

Credentials are printed to the terminal — store them in Hudu manually under the **NeConnect Audit Toolkit** asset layout using these fields:

| Field | Used For |
|---|---|
| Application ID | All modules (silent app-only authentication) |
| Tenant ID | All modules |
| Cert Base64 | Certificate decoded at runtime — no file path needed |
| Cert Password | Decrypts the certificate |
| Cert Expiry | Reminder to rotate before expiry |
| Powershell Launch Command | Pre-built launch commands (populated automatically) |

---

## Rotating the Certificate

Certificates are valid for 2 years by default (`-CertExpiryYears 1–5`).

### Interactive rotation (requires Global Admin browser login)

```powershell
.\Setup-365AuditApp.ps1 -HuduCompanyId '<slug>'           # rotate only if expiring within 30 days
.\Setup-365AuditApp.ps1 -HuduCompanyId '<slug>' -Force    # rotate regardless of expiry
```

### Non-interactive rotation (no browser — automated-friendly)

Once the initial setup has been run interactively (which grants `Application.ReadWrite.OwnedBy` and registers the service principal as owner of the app registration), future renewals can be done without a browser login by supplying the existing credentials:

```powershell
.\Setup-365AuditApp.ps1 -AppId '<AppId>' -TenantId '<TenantId>' `
    -CertBase64 '<paste>' -CertPassword (Read-Host -AsSecureString 'Cert Password') -Force
```

Or, if using Hudu, pass only the company ID — credentials are fetched automatically and the cert is renewed if expiring within 30 days:

```powershell
.\Setup-365AuditApp.ps1 -HuduCompanyId '<slug>'
```

The updated credentials are pushed to Hudu automatically. The audit HTML summary report also includes a **Toolkit / Certificate** action item when fewer than 30 days remain, prompting the reviewing tech to schedule a renewal.

---

## Hudu Integration

The toolkit integrates with Hudu to store and retrieve credentials automatically. Each tech needs their own Hudu API key; the base URL can be shared.

### Environment Variables

| Variable | Description |
|---|---|
| `HUDU_API_KEY` | Your personal Hudu API key (Profile → API Keys) |
| `HUDU_BASE_URL` | Hudu instance URL — defaults to `https://neconnect.huducloud.com` if unset |

Add to your PowerShell profile (`$PROFILE`) for persistent configuration:

```powershell
$env:HUDU_API_KEY  = 'your-api-key-here'
$env:HUDU_BASE_URL = 'https://hudu.yourcompany.com'   # omit if using neconnect.huducloud.com
```

For scheduled/automated runs, set these as system environment variables.

### Asset Layout

Credentials are stored in the **NeConnect Audit Toolkit** asset layout (one asset per customer company). The asset is created automatically by `Setup-365AuditApp.ps1` and named `NeConnect Audit Toolkit - <Company Name>`.

---

## Running the Audit

Open a PowerShell 7.4+ terminal, navigate to the toolkit directory, and run using one of the following methods:

### With Hudu API key (recommended)

Credentials are fetched automatically from Hudu — no copy/pasting required:

```powershell
.\Start-365Audit.ps1 -HuduCompanyId '<company-slug>'
.\Start-365Audit.ps1 -HuduCompanyName 'Contoso Ltd'    # exact name match
```

### Without Hudu API key

Copy the App ID, Tenant ID, and certificate details from the Hudu asset:

```powershell
# Prompts for Base64 and password interactively
.\Start-365Audit.ps1 -AppId '<AppId>' -TenantId '<TenantId>'

# Supply all credentials on the command line
.\Start-365Audit.ps1 -AppId '<AppId>' -TenantId '<TenantId>' `
    -CertBase64 '<paste>' -CertPassword (Read-Host -AsSecureString 'Cert Password')
```

### Non-interactive / automated (skip the menu)

Supply `-Modules` to bypass the menu entirely. The HTML report is generated but not opened automatically, making this suitable for scheduled tasks and bulk runs:

```powershell
.\Start-365Audit.ps1 -HuduCompanyId '<slug>' -Modules 1,2,3,4    # all modules
.\Start-365Audit.ps1 -HuduCompanyId '<slug>' -Modules 9          # same as above (option 9)
.\Start-365Audit.ps1 -HuduCompanyId '<slug>' -Modules 1,2        # Entra + Exchange only
```

On launch the toolkit will:
1. Fetch credentials from Hudu (if using `-HuduCompanyId` / `-HuduCompanyName`)
2. Check system clock drift against Microsoft's servers — certificate auth fails if drift exceeds 5 minutes; an **expired** certificate causes an immediate hard stop
3. Check local script versions against the GitHub version manifest and warn if updates are available
4. Decode the certificate from base64 to a temp `.pfx` in `$env:TEMP` (deleted on exit)
5. Present the module selection menu (skipped when `-Modules` is supplied)

Select one or more modules by number (comma-separated, e.g. `1,2,3`). All modules connect silently — no browser prompts.

---

## Automated Bulk Runs

`Start-UnattendedAudit.ps1` processes multiple customers in sequence without any interaction. For each customer it:

1. Calls `Setup-365AuditApp.ps1 -HuduCompanyId` to check the certificate — if expiring within 30 days, renews it automatically (no browser required) and pushes the new credentials back to Hudu
2. Calls `Start-365Audit.ps1 -HuduCompanyId -Modules 1,2,3,4` with credentials freshly fetched from Hudu
3. Generates the HTML summary report (not opened automatically)

### Setup

1. Copy `Start-UnattendedAudit.ps1.example` → `Start-UnattendedAudit.ps1` 
2. Copy `UnattendedCustomers.json.example` → `UnattendedCustomers.json` 
3. Edit `UnattendedCustomers.json` — add one entry per customer:
   ```json
   {
       "customers": [
           { "HuduCompanySlug": "a1b2c3d4e5f6", "Modules": [1, 2, 3, 4] },
           { "HuduCompanySlug": "f6e5d4c3b2a1", "Modules": [1, 2] }
       ]
   }
   ```
   The slug is the 12-character hex string from the Hudu company URL: `https://hudu.example.com/c/<slug>`
4. Set your Hudu API key in the environment:
   ```powershell
   $env:HUDU_API_KEY = 'your-api-key'
   ```
5. Run:
   ```powershell
   .\Start-UnattendedAudit.ps1
   ```

### Options

```powershell
.\Start-UnattendedAudit.ps1 -Customers 'contoso','fabrikam'    # run only these slugs from the JSON
.\Start-UnattendedAudit.ps1 -Modules 1,2                       # override modules for all customers this run
.\Start-UnattendedAudit.ps1 -SkipCertCheck                     # skip cert expiry check
```

> **Note:** Non-interactive cert renewal requires the app registration to have `Application.ReadWrite.OwnedBy` granted and the service principal registered as an owner of the app. This is set up automatically during the initial interactive `Setup-365AuditApp.ps1` run for each customer.

---

## Menu

| Option | Module | Description |
|---|---|---|
| 1 | Microsoft Entra Audit | Identity, MFA, roles, Conditional Access, Secure Score |
| 2 | Exchange Online Audit | Mailboxes, permissions, mail flow |
| 3 | SharePoint Online Audit | Sites, permissions, storage, OneDrive |
| 4 | Mail Security Audit | DKIM, DMARC, SPF, anti-spam/phish policies |
| 9 | Run All (1, 2, 3, 4) | Full audit, then generates the HTML summary once |
| 0 | Exit | — |

---

## Requirements

### PowerShell Version

- **7.4 or later** — required for SharePoint (`Invoke-SharePointAudit.ps1` uses PnP.PowerShell v3) and for `Setup-365AuditApp.ps1`
- **7.2 or later** — minimum for Entra, Exchange, and Mail Security modules only

Download: https://github.com/PowerShell/PowerShell/releases

### Required Modules

| Module | Required By | Install |
|---|---|---|
| `Microsoft.Graph.*` | All modules | `Install-Module Microsoft.Graph -Scope CurrentUser` |
| `ExchangeOnlineManagement` | Exchange, Mail Security | `Install-Module ExchangeOnlineManagement -Scope CurrentUser` |
| `PnP.PowerShell` (v3+) | SharePoint | `Install-Module PnP.PowerShell -Scope CurrentUser` |

Modules are checked at runtime and installed automatically if missing.

### Linux / macOS Additional Dependencies

All scripts have cross-platform support, but Linux and macOS require two additional system packages:

| Package | Purpose | Install (Debian/Ubuntu) | Install (macOS) |
|---|---|---|---|
| `openssl` | Certificate generation | `apt install openssl` | `brew install openssl` |
| `bind-utils` / `dnsutils` | DNS TXT lookups (`dig`) | `apt install dnsutils` | included with macOS |

Tested with **OpenSSL 3.x**. The `-legacy` flag is used automatically when exporting `.pfx` files for .NET compatibility.

---

## Module Reference

### Invoke-EntraAudit.ps1

Connects to Microsoft Graph and audits Entra ID (Azure Active Directory).

**Output files:**

| File | Description |
|---|---|
| `Entra_Users.csv` | UPN, name, licence, MFA status and methods, password policy, last sign-in |
| `Entra_Users_Unlicensed.csv` | Member accounts with no active licence |
| `Entra_Licenses.csv` | Subscriptions: total, consumed, suspended, and warning seat counts |
| `Entra_SSPR.csv` | Self-Service Password Reset enforcement status |
| `Entra_AdminRoles.csv` | All directory role assignments |
| `Entra_GlobalAdmins.csv` | Subset of Global Administrator accounts |
| `Entra_GuestUsers.csv` | Guest accounts with creation date and last sign-in |
| `Entra_Groups.csv` | All groups: type, membership rule, owners, and members |
| `Entra_CA_Policies.csv` | Conditional Access policy names, states, targets, client app types, and grant controls |
| `Entra_TrustedLocations.csv` | Named locations with trusted flag and IP ranges |
| `Entra_SecurityDefaults.csv` | Whether Security Defaults are enabled |
| `Entra_SecureScore.csv` | Identity Secure Score: current, max, and percentage |
| `Entra_SecureScoreControls.csv` | Per-control score and description (human-readable names) |
| `Entra_SignIns.csv` | Last 10 interactive sign-ins per user |
| `Entra_AccountCreations.csv` | Account creation events within the audit retention window |
| `Entra_AccountDeletions.csv` | Account deletion events within the audit retention window |
| `Entra_AuditEvents.csv` | Notable events: role changes and MFA/security info modifications |

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

---

### Invoke-SharePointAudit.ps1

Connects to SharePoint Online using PnP.PowerShell with certificate-based app-only authentication and audits sites and OneDrive.

> Requires PowerShell 7.4+ and the PnP.PowerShell v3+ module.

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
| Warning | Exchange / Rules | Inbox rules forwarding or redirecting mail |
| Warning | Exchange / Rules | Inbox rules in a broken/non-functional state |
| Warning | Exchange | No Safe Attachments or Safe Links policy enabled |
| Warning | SharePoint | Default sharing link allows anonymous (anyone) access |
| Warning | SharePoint | OneDrive sync not restricted to managed devices |

**Report sections:**

- **Microsoft Entra** — MFA coverage, stale licensed accounts, licence table, SSPR status, Security Defaults, global admin count, role summary, guest accounts and stale guest count, CA policies, legacy auth check, Identity Secure Score with control breakdown (To Action / Implemented)
- **Exchange Online** — Mailbox count and storage, delegated permissions, external forwarding rule alerts, broken inbox rules, shared mailbox sign-in status, outbound spam auto-forward policy, Safe Attachments and Safe Links status
- **SharePoint / OneDrive** — Tenant storage gauge, site collection table with expandable groups panel, external sharing policy and site overrides, access control policies, OneDrive usage and unlicensed accounts
- **Mail Security** — DKIM, DMARC, and SPF coverage per domain

---

## Output Structure

All module output lands in a folder created at the start of each session, one level above the repository root (to avoid git tracking audit data):

```
365Audit/            ← repository
<parent folder>/
└── <CompanyName>_<yyyyMMdd>/
    ├── OrgInfo.json
    ├── Entra_Users.csv
    ├── Entra_SecureScore.csv
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
├── Start-365Audit.ps1                    # Interactive launcher and module menu
├── Start-UnattendedAudit.ps1.example     # Bulk runner template (copy to Start-UnattendedAudit.ps1, excluded from git)
├── UnattendedCustomers.json.example      # Customer list template (copy to UnattendedCustomers.json, excluded from git)
├── Setup-365AuditApp.ps1                 # One-time app registration, certificate setup, and renewal
├── Invoke-EntraAudit.ps1        # Entra ID module
├── Invoke-ExchangeAudit.ps1     # Exchange Online module
├── Invoke-SharePointAudit.ps1   # SharePoint / OneDrive module
├── Invoke-MailSecurityAudit.ps1 # Mail security module
├── Generate-AuditSummary.ps1    # HTML report generator
├── version.json                 # GitHub version manifest
├── CHANGELOG.md                 # Full version history for all scripts
└── common/
    └── Audit-Common.ps1         # Shared helpers (Graph/EXO auth, output folder, version check)
```
