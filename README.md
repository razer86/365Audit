# NeConnect Microsoft 365 Audit Toolkit

Monthly Microsoft 365 audit toolkit for MSP maintenance reporting. Run per customer to generate a report covering identity, messaging, and storage health.

---

## First-Time Setup

> **New Hudu instance?** Before setting up any customers, run `New-HuduAssetLayout.ps1` once to create the required asset layout in Hudu. See [Helpers — New-HuduAssetLayout.ps1](#new-huduassetlayoutps1).

Before running the toolkit for any customer, run `Setup-365AuditApp.ps1` once in that tenant as a **Global Administrator**.

The script will:
1. Create an app registration with all required Microsoft Graph, Exchange Online, and SharePoint permissions and grant admin consent
2. Generate a self-signed certificate, upload the public key to the app registration
3. Print all credentials to the terminal
4. **Push credentials to Hudu** if `HuduApiKey` is configured in `config.psd1` and a matching asset already exists for the company

### With Hudu integration (recommended)

If `HuduApiKey` is configured in `config.psd1` (see [Hudu Integration](#hudu-integration)), pass the company slug or ID and credentials are stored automatically:

```powershell
.\Setup-365AuditApp.ps1 -HuduCompanyId '<company-slug>'
.\Setup-365AuditApp.ps1 -HuduCompanyName 'Contoso Ltd'
```

### Without Hudu integration

```powershell
.\Setup-365AuditApp.ps1
```

Credentials are printed to the terminal — store them in Hudu manually under the asset layout configured as `HuduAssetName` in `config.psd1`, using these fields:

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

Certificates are valid for 2 years by default (`-CertExpiryYears 1–5`). All rotation paths are non-interactive once the initial setup has been completed — no browser login is required for renewals.

### With Hudu (recommended)

Credentials are fetched automatically from Hudu — no copy/pasting required:

```powershell
.\Setup-365AuditApp.ps1 -HuduCompanyId '<slug>'           # rotate only if expiring within 30 days
.\Setup-365AuditApp.ps1 -HuduCompanyId '<slug>' -Force    # rotate regardless of expiry
```

Updated credentials are written back to the existing Hudu asset automatically.

### Without Hudu

Supply the existing credentials explicitly:

```powershell
.\Setup-365AuditApp.ps1 -AppId '<AppId>' -TenantId '<TenantId>' `
    -CertBase64 '<paste>' -CertPassword (Read-Host -AsSecureString 'Cert Password') -Force
```

> **Note:** Non-interactive renewal requires the initial interactive setup (`Setup-365AuditApp.ps1` run as Global Admin without parameters) to have been completed at least once per tenant. This grants `Application.ReadWrite.OwnedBy` and registers the service principal as an owner of the app registration, which is required for the app to renew its own certificate without a browser login.

The audit HTML summary report includes a **Toolkit / Certificate** action item when fewer than 30 days remain, prompting the reviewing tech to schedule a renewal.

---

## Hudu Integration

The toolkit integrates with Hudu to store and retrieve credentials automatically. Each tech needs their own Hudu API key; the base URL is shared.

### config.psd1

Hudu credentials and toolkit settings are stored in `config.psd1` in the repository root. Copy the example file to get started:

```powershell
Copy-Item config.psd1.example config.psd1
```

Then edit `config.psd1`:

```powershell
@{
    HuduBaseUrl       = 'https://your-hudu-instance.com'
    HuduApiKey        = 'your-api-key-here'     # Profile → API Keys in Hudu
    HuduAssetLayoutId = 67                       # See Helpers\Get-HuduAssetLayouts.ps1
    HuduAssetName     = 'M365 Audit Toolkit'     # Prefix for asset names: "<HuduAssetName> - <Company>"
    MspDomains        = @('yourdomain.com')      # Used for Technical Contact checks in summary
    KnownPartners     = @('Your Company Name')   # Used for GDAP partner checks in summary
}
```

> `config.psd1` is excluded from git to protect credentials.

### Asset Layout

Credentials are stored in a dedicated Hudu asset layout — one asset per customer company. The asset layout must already exist in Hudu and be configured via `HuduAssetLayoutId` in `config.psd1`. When a matching asset is found for the company, `Setup-365AuditApp.ps1` updates it with the latest credentials. The asset is named `<HuduAssetName> - <Company Name>`, where `HuduAssetName` is set in `config.psd1`.

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

# Supply Base64 on the command line — password is still prompted as a SecureString
.\Start-365Audit.ps1 -AppId '<AppId>' -TenantId '<TenantId>' `
    -CertBase64 '<paste>' -CertPassword (Read-Host -AsSecureString 'Cert Password')
```

> **`-CertPassword` only accepts a `SecureString`** — plain text strings are intentionally not supported. `Read-Host -AsSecureString` is the standard way to supply it interactively. This ensures the password is never stored in plaintext in shell history, process memory, or log files.

### Non-interactive / automated (skip the menu)

Supply `-Modules` to bypass the menu entirely. The HTML report is generated but not opened automatically, making this suitable for scheduled tasks and bulk runs:

```powershell
.\Start-365Audit.ps1 -HuduCompanyId '<slug>' -Modules 1,2,3,4,5,6,7    # all modules
.\Start-365Audit.ps1 -HuduCompanyId '<slug>' -Modules A              # same as above (Run All)
.\Start-365Audit.ps1 -HuduCompanyId '<slug>' -Modules 1,2            # Entra + Exchange only
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
2. Calls `Start-365Audit.ps1 -HuduCompanyId -Modules <customer modules>` with credentials freshly fetched from Hudu — modules are taken from the customer's entry in `UnattendedCustomers.psd1`, or overridden globally via `-Modules`
3. Generates the HTML summary report (not opened automatically)

### Setup

1. Copy `UnattendedCustomers.psd1.example` → `UnattendedCustomers.psd1`
2. Edit `UnattendedCustomers.psd1` — add one entry per customer:

   ```powershell
   @{
       Customers = @(
           @{ HuduCompanySlug = 'a1b2c3d4e5f6'; HuduCompanyName = 'Contoso Ltd';  Modules = @('A') }
           @{ HuduCompanySlug = 'f6e5d4c3b2a1'; HuduCompanyName = 'Fabrikam Inc'; Modules = @(1, 2) }
       )
   }
   ```
   The slug is the 12-character hex string from the Hudu company URL: `https://hudu.example.com/c/<slug>`
3. Ensure `HuduApiKey` is set in `config.psd1`
4. Run:
   ```powershell
   .\Start-UnattendedAudit.ps1
   ```

### Options

```powershell
.\Start-UnattendedAudit.ps1 -Customers 'contoso','fabrikam'    # run only these slugs from the PSD1
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
| 5 | Intune / Endpoint Audit | Devices, compliance policies, configuration profiles, apps, enrollment |
| 6 | Teams Audit | Federation, client config, meeting/guest/messaging policies, app policies |
| 7 | ScubaGear CIS Baseline | CISA M365 Foundations Benchmark assessment (runs in Windows PowerShell 5.1) |
| A | Run All (1–7) | Full audit across all modules, then generates the HTML summary once |
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
| `Entra_EnterpriseApps.csv` | Third-party enterprise apps with admin-consent status and consented role count |
| `Entra_EnterpriseAppPermissions.csv` | Application and delegated permissions granted to each enterprise app |
| `Entra_AppRegistrations.csv` | App registrations with credential expiry dates and secret/cert counts |
| `Entra_AppRegistrationPermissions.csv` | Permissions declared in `requiredResourceAccess` for each app registration |
| `Entra_RiskyUsers.csv` | Users flagged by Identity Protection (requires Azure AD P2; absent if unlicensed) |
| `Entra_RiskySignIns.csv` | Risky sign-in events (requires Azure AD P2; absent if unlicensed) |
| `Entra_PIMAssignments.csv` | Privileged Identity Management role assignments (requires Azure AD P2; absent if unlicensed) |

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

### Invoke-IntuneAudit.ps1

Connects to Microsoft Graph and audits Intune / Endpoint Management. Gracefully skips with an informational note if no Intune-capable licence is detected.

**Output files:**

| File | Description |
|---|---|
| `Intune_Devices.csv` | Managed device inventory with OS, ownership, compliance state, and last sync |
| `Intune_DeviceComplianceStates.csv` | Per-device compliance policy state |
| `Intune_CompliancePolicies.csv` | Compliance policies with platform, assignment scope, grace period, and settings |
| `Intune_ConfigProfiles.csv` | Configuration profiles with platform, type, last modified, and assignments |
| `Intune_ConfigProfileSettings.csv` | Per-setting detail for each configuration profile |
| `Intune_Apps.csv` | Assigned app inventory with install/failed/pending counts and assignment details |
| `Intune_AutopilotDevices.csv` | Windows Autopilot device identities (skipped gracefully on 403) |
| `Intune_EnrollmentRestrictions.csv` | Enrollment restriction policies |

---

### Invoke-TeamsAudit.ps1

Connects to Microsoft Teams via the MicrosoftTeams module and audits federation, client, meeting, and policy configuration.

**Output files:**

| File | Description |
|---|---|
| `Teams_FederationConfig.csv` | External access (federation) settings — allowed/blocked domains, Teams and Skype federation flags |
| `Teams_ClientConfig.csv` | Teams client configuration — file sharing, cloud storage, external app, and communication settings |
| `Teams_MeetingPolicies.csv` | Meeting policies — recording, transcription, lobby, and external participant settings |
| `Teams_GuestMeetingConfig.csv` | Guest meeting configuration — IP audio/video, screen share, and meeting flags |
| `Teams_GuestCallingConfig.csv` | Guest calling settings |
| `Teams_MessagingPolicies.csv` | Messaging policies — Giphy, memes, URL previews, read receipts, and chat edit/delete settings |
| `Teams_AppPermissionPolicies.csv` | App permission policies — Microsoft, third-party, and tenant app access controls |
| `Teams_AppSetupPolicies.csv` | App setup policies — pinned apps and user pinning settings |
| `Teams_ChannelPolicies.csv` | Channel policies — private and shared channel creation permissions (skipped gracefully if not available in module version) |

---

### Invoke-ScubaGearAudit.ps1

Runs the [CISA ScubaGear M365 Foundations Benchmark](https://github.com/cisagov/ScubaGear) assessment against the tenant. ScubaGear and its dependencies must run in Windows PowerShell 5.1 — this module spawns a clean PS 5.1 subprocess automatically and cleans up the temporary certificate import on exit.

> Requires `powershell.exe` (Windows PowerShell 5.1) to be available on the machine. Power Platform is excluded (requires an interactive one-time registration).

**Output folder:** `Raw\ScubaGear_<timestamp>\`

| File | Description |
|---|---|
| `BaselineReports.html` | ScubaGear's own interactive HTML report with full control details |
| `ScubaResults_<uuid>.json` | Consolidated JSON results ingested by `Generate-AuditSummary.ps1` |
| `ScubaResults.csv` | Flat CSV of all controls with pass/fail/warning status |
| `ActionPlan.csv` | Failing Shall controls with blank remediation fields for MSP follow-up |
| `IndividualReports\` | Per-product HTML and JSON reports (AAD, Defender, EXO, SharePoint, Teams) |

`Generate-AuditSummary.ps1` detects the `ScubaGear_*` folder automatically and adds failing controls to the action items list plus a collapsible CIS Baseline section to the report.

---

### Generate-AuditSummary.ps1

Reads CSV files from the current audit run's `Raw` folder and compiles them into a single HTML report (`M365_AuditSummary.html`), which opens automatically in the default browser.

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

- **Microsoft Entra** — MFA coverage, licence table, SSPR status, Security Defaults, global admin count, role summary, guest accounts, CA policies, legacy auth check, Identity Secure Score with To Action / Implemented breakdown, external collaboration settings, app registrations with collapsible permissions, enterprise apps with collapsible permissions, risky users / sign-ins
- **Exchange Online** — Mailbox count and storage, delegated permissions, external forwarding rule alerts, broken inbox rules, shared mailbox sign-in status, outbound spam auto-forward policy, Safe Attachments and Safe Links status
- **SharePoint / OneDrive** — Tenant storage gauge, site collection table with expandable groups panel, external sharing policy and site overrides, access control policies, OneDrive per-user storage table and unlicensed accounts
- **Mail Security** — DKIM, DMARC, and SPF coverage per domain
- **Intune** — Licence status, device inventory with OS and compliance breakdown, stale devices, compliance policies, configuration profiles, app install summary
- **Teams** — Federation settings, client config, meeting policies, guest access settings, app policies
- **ScubaGear CIS Baseline** — Per-product pass/fail/warning counts with link to full ScubaGear HTML report (collapsed by default; only rendered when ScubaGear output is present)
- **Compliance Overview** — Distribution bar (passed / warnings / critical) with per-module breakdown and list of CIS controls with findings
- **Technical Issues** — Collection failures recorded by `Add-AuditIssue` catch blocks across all modules

---

## Output Structure

Each audit run creates a customer folder one level above the repository root (to avoid git tracking audit data). `OrgInfo.json` and `M365_AuditSummary.html` stay at the root of that folder, while all generated CSVs, JSON files, and the session transcript are stored under `Raw`:

```
365Audit/            ← repository
<parent folder>/
└── <CompanyName>_<yyyyMMdd>/
    ├── OrgInfo.json
    └── M365_AuditSummary.html
    └── Raw/
        ├── AuditLog.txt
        ├── Entra_Users.csv
        ├── Entra_SecureScore.csv
        ├── ... (all module CSVs and JSON files)
```

The folder name is derived from the Entra organisation display name (alphanumeric only) and the current date. Running the toolkit multiple times on the same day reuses the same folder.

---

## Version Check

On each launch, the toolkit downloads `version.json` from GitHub and compares it against the `$ScriptVersion` declared in each local script. Outdated scripts are listed by name with the installed and latest versions.

The check is non-blocking — a network failure produces a warning and the toolkit continues normally.

---

## Helpers

Standalone utility scripts in the `Helpers\` folder. None of these are required for normal audit runs — they assist with setup, maintenance, and troubleshooting.

### Get-HuduAssetLayouts.ps1

Connects to Hudu and lists all asset layouts with their numeric IDs. Use this to find the correct value for `HuduAssetLayoutId` in `config.psd1` — Hudu's UI only shows slugs, not IDs.

```powershell
.\Helpers\Get-HuduAssetLayouts.ps1
```

Reads `HuduBaseUrl` and `HuduApiKey` from `config.psd1`.

---

### Get-ModuleVersionStatus.ps1

Performs a single bulk PSGallery lookup for all modules required by the toolkit and displays a status table showing installed vs latest versions.

```powershell
.\Helpers\Get-ModuleVersionStatus.ps1
```

| Status | Meaning |
|---|---|
| `OK` | Installed and up to date |
| `UPDATE AVAILABLE` | Newer version exists in PSGallery |
| `NOT INSTALLED` | Not yet installed — will be installed automatically on first run |
| `MULTIPLE VERSIONS` | More than one version installed — run `Uninstall-AuditModules.ps1` to clean up |

---

### New-HuduAssetLayout.ps1

Creates the M365 Audit Toolkit asset layout in Hudu. Run this once on a new Hudu instance before running `Setup-365AuditApp.ps1` for the first time — the layout must exist before credentials can be pushed to it.

```powershell
# Preview what would be created
.\Helpers\New-HuduAssetLayout.ps1 -WhatIf

# Create the layout
.\Helpers\New-HuduAssetLayout.ps1
```

After creation, copy the printed layout ID into `config.psd1`:

```powershell
HuduAssetLayoutId = <id printed by the script>
```

> Requires **Hudu Administrator or Super Administrator** — a standard user API key will receive a 422 error.

Reads `HuduBaseUrl`, `HuduApiKey`, and `HuduAssetName` from `config.psd1`.

---

### Remove-AuditCustomer.ps1

Offboards a customer by removing their app registration from Entra ID and deleting the corresponding Hudu asset. Use when a customer leaves or when you need to fully reset a customer's 365Audit configuration.

```powershell
# Hudu lookup — resolves AppId/TenantId from the asset automatically
.\Helpers\Remove-AuditCustomer.ps1 -HuduCompanyId 'a1b2c3d4e5f6'
.\Helpers\Remove-AuditCustomer.ps1 -HuduCompanyName 'Contoso Ltd'

# Direct — removes Entra app only (no Hudu asset involved)
.\Helpers\Remove-AuditCustomer.ps1 -AppId '<AppId>' -TenantId '<TenantId>'

# Preview without making changes
.\Helpers\Remove-AuditCustomer.ps1 -HuduCompanyId 'a1b2c3d4e5f6' -WhatIf

# Permanently purge from Entra recycle bin (cannot be undone)
.\Helpers\Remove-AuditCustomer.ps1 -HuduCompanyId 'a1b2c3d4e5f6' -PermanentDelete
```

By default the app is **soft-deleted** and remains recoverable from the Entra recycle bin for 30 days. Use `-PermanentDelete` only when certain the customer will not be re-onboarded.

Reads `HuduBaseUrl`, `HuduApiKey`, `HuduAssetLayoutId`, and `HuduAssetName` from `config.psd1`.

---

### Uninstall-AuditModules.ps1

Removes all installed versions of every module required by the toolkit. Useful for testing a clean first-run install experience or resolving conflicting module versions.

```powershell
# Preview what would be removed
.\Helpers\Uninstall-AuditModules.ps1 -WhatIf

# Remove everything
.\Helpers\Uninstall-AuditModules.ps1
```

> Run in a **fresh PowerShell session** that has not loaded any 365Audit scripts — loaded modules cannot be uninstalled until the session is closed. If any modules were installed in `AllUsers` scope, run as Administrator.

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
├── Common/
│   └── Audit-Common.ps1                      # Shared helpers (Graph/EXO auth, output folder, version check)
├── Helpers/
│   ├── Get-HuduAssetLayouts.ps1              # Lists Hudu asset layouts to find the correct layout ID
│   ├── Get-ModuleVersionStatus.ps1           # Checks installed vs latest PSGallery versions for all modules
│   ├── New-HuduAssetLayout.ps1               # Creates the M365 Audit Toolkit asset layout in Hudu
│   ├── Remove-AuditCustomer.ps1              # Offboards a customer — removes Entra app and Hudu asset
│   ├── Sync-UnattendedCustomers.ps1          # Syncs UnattendedCustomers.psd1 from Hudu assets
│   └── Uninstall-AuditModules.ps1            # Removes all toolkit modules (clean reinstall / conflict resolution)
├── CHANGELOG.md                              # Full version history for all scripts
├── config.psd1.example                       # Config template (copy to config.psd1)
├── Generate-AuditSummary.ps1                 # HTML report generator
├── Invoke-EntraAudit.ps1                     # Entra ID module
├── Invoke-ExchangeAudit.ps1                  # Exchange Online module
├── Invoke-IntuneAudit.ps1                    # Intune / Endpoint module
├── Invoke-MailSecurityAudit.ps1              # Mail security module
├── Invoke-ScubaGearAudit.ps1                 # CISA ScubaGear CIS Baseline module
├── Invoke-SharePointAudit.ps1                # SharePoint / OneDrive module
├── Invoke-TeamsAudit.ps1                     # Teams module
├── Setup-365AuditApp.ps1                     # One-time app registration, certificate setup, and renewal
├── Start-365Audit.ps1                        # Interactive launcher and module menu
├── Start-UnattendedAudit.ps1                 # Automated bulk runner
├── UnattendedCustomers.psd1.example          # Customer list template (copy to UnattendedCustomers.psd1)
└── version.json                              # GitHub version manifest
```
