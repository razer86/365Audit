# NeConnect Microsoft 365 Audit Toolkit

Monthly Microsoft 365 audit toolkit for MSP maintenance reporting. Run per customer to generate a report covering identity, messaging, and storage health.

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

| Package | Purpose | Install (Debian/Ubuntu) | Install (macOS) |
|---|---|---|---|
| `openssl` | Certificate generation | `apt install openssl` | `brew install openssl` |
| `bind-utils` / `dnsutils` | DNS TXT lookups (`dig`) | `apt install dnsutils` | included with macOS |

Tested with **OpenSSL 3.x**. The `-legacy` flag is used automatically when exporting `.pfx` files for .NET compatibility.

---

## First-Time Setup

> **New Hudu instance?** Before setting up any customers, run `New-HuduAssetLayout.ps1` once to create the required asset layout in Hudu. See [Helpers — New-HuduAssetLayout.ps1](docs/Helpers.md#new-huduassetlayoutps1).

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

```powershell
.\Setup-365AuditApp.ps1 -HuduCompanyId '<slug>'           # rotate only if expiring within 30 days
.\Setup-365AuditApp.ps1 -HuduCompanyId '<slug>' -Force    # rotate regardless of expiry
```

Updated credentials are written back to the existing Hudu asset automatically.

### Without Hudu

```powershell
.\Setup-365AuditApp.ps1 -AppId '<AppId>' -TenantId '<TenantId>' `
    -CertBase64 '<paste>' -CertPassword (Read-Host -AsSecureString 'Cert Password') -Force
```

> **Note:** Non-interactive renewal requires the initial interactive setup to have been completed at least once per tenant. This grants `Application.ReadWrite.OwnedBy` and registers the service principal as an owner of the app registration.

The audit HTML summary report includes a **Toolkit / Certificate** action item when fewer than 30 days remain.

---

## Hudu Integration

The toolkit integrates with Hudu to store and retrieve credentials automatically. Each tech needs their own Hudu API key; the base URL is shared.

### config.psd1

Copy the example file to get started:

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

Credentials are stored in a dedicated Hudu asset layout — one asset per customer company. The layout must already exist in Hudu and be configured via `HuduAssetLayoutId` in `config.psd1`. When a matching asset is found for the company, `Setup-365AuditApp.ps1` updates it with the latest credentials.

---

## Running the Audit

Open a PowerShell 7.4+ terminal, navigate to the toolkit directory, and run using one of the following methods:

### With Hudu API key (recommended)

```powershell
.\Start-365Audit.ps1 -HuduCompanyId '<company-slug>'
.\Start-365Audit.ps1 -HuduCompanyName 'Contoso Ltd'    # exact name match
```

### Without Hudu API key

```powershell
# Prompts for Base64 and password interactively
.\Start-365Audit.ps1 -AppId '<AppId>' -TenantId '<TenantId>'

# Supply Base64 on the command line — password is still prompted as a SecureString
.\Start-365Audit.ps1 -AppId '<AppId>' -TenantId '<TenantId>' `
    -CertBase64 '<paste>' -CertPassword (Read-Host -AsSecureString 'Cert Password')
```

> **`-CertPassword` only accepts a `SecureString`** — plain text strings are intentionally not supported. This ensures the password is never stored in plaintext in shell history, process memory, or log files.

### Non-interactive / automated (skip the menu)

Supply `-Modules` to bypass the menu entirely. The HTML report is generated but not opened automatically:

```powershell
.\Start-365Audit.ps1 -HuduCompanyId '<slug>' -Modules 1,2,3,4   # specific modules
.\Start-365Audit.ps1 -HuduCompanyId '<slug>' -Modules A          # all modules
```

On launch the toolkit will:
1. Fetch credentials from Hudu (if using `-HuduCompanyId` / `-HuduCompanyName`)
2. Check system clock drift — certificate auth fails if drift exceeds 5 minutes; an **expired** certificate causes an immediate hard stop
3. Check local script versions against the GitHub version manifest and warn if updates are available
4. Decode the certificate from base64 to a temp `.pfx` in `$env:TEMP` (deleted on exit)
5. Present the module selection menu (skipped when `-Modules` is supplied)

---

## Module Menu

| Option | Module | Description | Docs |
|---|---|---|---|
| 1 | Invoke-EntraAudit.ps1 | Identity, MFA, roles, Conditional Access, Secure Score | [Details](docs/Invoke-EntraAudit.md) |
| 2 | Invoke-ExchangeAudit.ps1 | Mailboxes, permissions, mail flow | [Details](docs/Invoke-ExchangeAudit.md) |
| 3 | Invoke-SharePointAudit.ps1 | Sites, permissions, storage, OneDrive | [Details](docs/Invoke-SharePointAudit.md) |
| 4 | Invoke-MailSecurityAudit.ps1 | DKIM, DMARC, SPF, anti-spam/phish policies | [Details](docs/Invoke-MailSecurityAudit.md) |
| 5 | Invoke-IntuneAudit.ps1 | Devices, compliance policies, configuration profiles, apps | [Details](docs/Invoke-IntuneAudit.md) |
| 6 | Invoke-TeamsAudit.ps1 | Federation, client config, meeting/guest/messaging policies | [Details](docs/Invoke-TeamsAudit.md) |
| 7 | Invoke-ScubaGearAudit.ps1 | CISA M365 Foundations Benchmark (Windows PowerShell 5.1) | [Details](docs/Invoke-ScubaGearAudit.md) |
| A | Run All (1–7) | Full audit across all modules, then generates the HTML summary | — |
| 0 | Exit | — | — |

The HTML summary report is documented separately: [Generate-AuditSummary.ps1](docs/Generate-AuditSummary.md)

---

## Automated Bulk Runs

`Start-UnattendedAudit.ps1` processes multiple customers in sequence without any interaction. For each customer it:

1. Calls `Setup-365AuditApp.ps1 -HuduCompanyId` to check the certificate — if expiring within 30 days, renews automatically and pushes new credentials back to Hudu
2. Calls `Start-365Audit.ps1 -HuduCompanyId -Modules` with credentials freshly fetched from Hudu
3. Generates the HTML summary report (not opened automatically)

### Setup

1. Copy `UnattendedCustomers.json.example` → `UnattendedCustomers.json`
2. Edit `UnattendedCustomers.json` — add one entry per customer:

   ```json
   {
       "customers": [
           { "HuduCompanySlug": "a1b2c3d4e5f6", "Modules": [1, 2, 3, 4] },
           { "HuduCompanySlug": "f6e5d4c3b2a1", "Modules": [1, 2] }
       ]
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
.\Start-UnattendedAudit.ps1 -Customers 'a1b2c3d4e5f6'    # run only this customer
.\Start-UnattendedAudit.ps1 -Modules 1,2                  # override modules for all customers this run
.\Start-UnattendedAudit.ps1 -SkipCertCheck                # skip cert expiry check
```

> **Note:** Non-interactive cert renewal requires the app registration to have `Application.ReadWrite.OwnedBy` granted and the service principal registered as an owner of the app. This is set up automatically during the initial interactive `Setup-365AuditApp.ps1` run for each customer.

---

## Output Structure

By default, output is written one level above the repository root to prevent audit data from being tracked by git. Override this with `-OutputRoot`:

```powershell
.\Start-365Audit.ps1 -HuduCompanyId '<slug>' -OutputRoot 'D:\AuditReports'
```

`OutputRoot` can also be set permanently in `config.psd1` so it applies to every run without needing to pass it each time.

Each audit run creates a customer folder inside the output root:

```
365Audit/            ← repository
<parent folder>/
└── <CompanyName>_<yyyyMMdd>/
    ├── OrgInfo.json
    ├── M365_AuditSummary.html
    └── Raw/
        ├── AuditLog.txt
        ├── Entra_Users.csv
        └── ... (all module CSVs and JSON files)
```

The folder name is derived from the Entra organisation display name (alphanumeric only) and the current date. Running the toolkit multiple times on the same day reuses the same folder.

---

## Version Check

On each launch, the toolkit downloads `version.json` from GitHub and compares it against the `$ScriptVersion` declared in each local script. Outdated scripts are listed by name with the installed and latest versions. The check is non-blocking — a network failure produces a warning and the toolkit continues normally.

---

## Helpers

Standalone utility scripts in the `Helpers\` folder. None are required for normal audit runs.

| Script | Purpose |
|---|---|
| [Get-HuduAssetLayouts.ps1](docs/Helpers.md#get-huduassetlayoutsps1) | Find numeric asset layout IDs for `config.psd1` |
| [Get-ModuleVersionStatus.ps1](docs/Helpers.md#get-moduleversionstatusps1) | Diagnose module version conflicts |
| [New-HuduAssetLayout.ps1](docs/Helpers.md#new-huduassetlayoutps1) | One-time Hudu layout creation for new instances |
| [Publish-HuduAuditReport.ps1](docs/Helpers.md#publish-huduauditreportps1) | Push a completed report into Hudu |
| [Remove-AuditCustomer.ps1](docs/Helpers.md#remove-auditcustomerps1) | Offboard a customer — remove Entra app and archive Hudu asset |
| [Sync-UnattendedCustomers.ps1](docs/Helpers.md#sync-unattendedcustomersps1) | Populate `UnattendedCustomers.json` from Hudu automatically |
| [Uninstall-AuditModules.ps1](docs/Helpers.md#uninstall-auditmodulesps1) | Clean-remove all toolkit modules for a fresh reinstall |

See [docs/Helpers.md](docs/Helpers.md) for full usage, parameters, and examples.

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
├── docs/
│   ├── Generate-AuditSummary.md              # HTML report generator — sections, action items, parameters
│   ├── Invoke-EntraAudit.md                  # Entra module — permissions, output files
│   ├── Invoke-ExchangeAudit.md               # Exchange module — permissions, output files
│   ├── Invoke-IntuneAudit.md                 # Intune module — permissions, output files
│   ├── Invoke-MailSecurityAudit.md           # Mail Security module — permissions, output files
│   ├── Invoke-ScubaGearAudit.md              # ScubaGear module — requirements, output files
│   ├── Invoke-SharePointAudit.md             # SharePoint module — permissions, output files
│   └── Invoke-TeamsAudit.md                  # Teams module — permissions, output files
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
├── UnattendedCustomers.json.example          # Customer list template (copy to UnattendedCustomers.json)
└── version.json                              # GitHub version manifest
```
