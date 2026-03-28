# Generate-AuditSummary.ps1

Reads CSV and JSON files from the current audit run's `Raw` folder and compiles them into a single styled HTML report (`M365_AuditSummary.html`).

Called automatically at the end of every audit run by `Start-365Audit.ps1`. When running via `-Modules` (non-interactive / automated), the report is generated but not opened in the browser.

## Parameters

| Parameter | Type | Description |
|---|---|---|
| `-AuditFolder` | String (Mandatory) | Path to the customer audit folder |
| `-NoOpen` | Switch | Generate the report without opening it in the browser |
| `-CertExpiryDays` | Int | When 0–30, inserts a Toolkit / Certificate warning action item |
| `-OutputPath` | String | Override the default output path for the generated HTML report |

## Action Items

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
| Warning | Toolkit / Certificate | Certificate expiring within 30 days |

## Report Sections

- **Microsoft Entra** — MFA coverage, licence table, SSPR status, Security Defaults, global admin count, role summary, guest accounts, CA policies, legacy auth check, Identity Secure Score with To Action / Implemented breakdown, external collaboration settings, app registrations with collapsible permissions, enterprise apps with collapsible permissions, risky users / sign-ins
- **Exchange Online** — Mailbox count and storage, delegated permissions, external forwarding rule alerts, broken inbox rules, shared mailbox sign-in status, outbound spam auto-forward policy, Safe Attachments and Safe Links status
- **SharePoint / OneDrive** — Tenant storage gauge, site collection table with expandable groups panel, external sharing policy and site overrides, access control policies, OneDrive per-user storage table and unlicensed accounts
- **Mail Security** — DKIM, DMARC, and SPF coverage per domain
- **Intune** — Licence status, device inventory with OS and compliance breakdown, stale devices, compliance policies, configuration profiles, app install summary
- **Teams** — Federation settings, client config, meeting policies, guest access settings, app policies
- **ScubaGear CIS Baseline** — Per-product pass/fail/warning counts with link to full ScubaGear HTML report (collapsed by default; only rendered when ScubaGear output is present)
- **Compliance Overview** — Distribution bar (passed / warnings / critical) with per-module breakdown and list of CIS controls with findings
- **Technical Issues** — Collection failures recorded by `Add-AuditIssue` catch blocks across all modules

