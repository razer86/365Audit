# Helpers

Standalone utility scripts in the `Helpers\` folder. None are required for normal audit runs — they assist with initial setup, ongoing maintenance, troubleshooting, and customer lifecycle management.

---

## Get-HuduAssetLayouts.ps1

**Purpose:** Find the numeric asset layout ID needed for `config.psd1`.

Hudu's web UI only exposes slugs in page URLs — the numeric ID required by the API is not visible anywhere in the UI. Run this once after setting up a new Hudu instance, or whenever you need to confirm an existing layout ID.

```powershell
.\Helpers\Get-HuduAssetLayouts.ps1
```

Output example:

```
ID   Name
--   ----
42   Passwords
67   M365 Audit Toolkit
68   Monthly Audit Report
71   Network Devices
```

Copy the ID for the `M365 Audit Toolkit` layout into `config.psd1` as `HuduAssetLayoutId`.

**Reads from:** `HuduBaseUrl`, `HuduApiKey` in `config.psd1` (or pass via parameters).

---

## Get-ModuleVersionStatus.ps1

**Purpose:** Diagnose module version issues before opening a bug report.

Performs a single bulk PSGallery lookup for every module required by the toolkit and displays a status table comparing installed vs latest versions. Run this first when troubleshooting MSAL assembly conflicts, unexpected authentication failures, or any error that mentions a module name or version.

```powershell
.\Helpers\Get-ModuleVersionStatus.ps1
```

| Status | Meaning |
|---|---|
| `OK` | Installed and up to date |
| `UPDATE AVAILABLE` | Newer version exists in PSGallery — run `Update-Module <name>` |
| `NOT INSTALLED` | Not yet installed — will be installed automatically on first audit run |
| `MULTIPLE VERSIONS` | More than one version installed side-by-side — run `Uninstall-AuditModules.ps1` then reinstall |

Also checks the Windows PowerShell 5.1 module path for ScubaGear dependencies.

---

## New-HuduAssetLayout.ps1

**Purpose:** One-time setup — create the required asset layout in a new Hudu instance.

`Setup-365AuditApp.ps1` stores audit credentials in a Hudu asset under a specific layout. That layout must already exist before credentials can be pushed. Run this script once per Hudu instance — you do not need to run it again for each customer.

```powershell
# Preview what would be created
.\Helpers\New-HuduAssetLayout.ps1 -WhatIf

# Create the layout
.\Helpers\New-HuduAssetLayout.ps1

# Override the layout name or appearance
.\Helpers\New-HuduAssetLayout.ps1 -LayoutName 'M365 Audit Toolkit' -Color '#1E40AF' -Icon 'fas fa-shield-halved'
```

After creation the script prints the new layout ID. Copy it into `config.psd1`:

```powershell
HuduAssetLayoutId = <printed ID>
```

> **Requires Hudu Administrator or Super Administrator** — a standard user API key will receive a `422 Unprocessable Entity` error.

**Reads from:** `HuduBaseUrl`, `HuduApiKey`, `HuduAssetName` in `config.psd1`.

---

## Publish-HuduAuditReport.ps1

**Purpose:** Push a completed audit report into Hudu after a run.

After a successful audit this script:

1. Locates or creates a `Monthly Audit Report` asset for the company (uses a separate asset layout from the credentials asset)
2. Writes the content of `M365_HuduReport.html` into the asset's `report_summary` rich-text field — this is the Hudu-optimised summary intended to be read inline
3. Uploads `M365_AuditSummary.html` (the full interactive report) as an attachment to the asset
4. Compresses the entire output folder to a `.zip` and uploads that as a second attachment, preserving all raw CSV/JSON data

```powershell
.\Helpers\Publish-HuduAuditReport.ps1 `
    -OutputPath  'C:\AuditReports\ContosoPty_20260326' `
    -CompanySlug 'a1b2c3d4e5f6' `
    -HuduBaseUrl 'https://hudu.example.com' `
    -HuduApiKey  'your-api-key'
```

This is typically called automatically by `Start-UnattendedAudit.ps1` at the end of each customer's run. It can also be run manually to re-publish an existing report.

| Parameter | Description |
|---|---|
| `-OutputPath` | Path to the customer's audit output folder |
| `-CompanySlug` | 12-character hex Hudu company slug |
| `-HuduBaseUrl` | Hudu instance base URL |
| `-HuduApiKey` | Hudu API key |
| `-ReportLayoutId` | Asset layout ID for the `Monthly Audit Report` layout (default: `68`) |

---

## Remove-AuditCustomer.ps1

**Purpose:** Offboard a customer — remove the Entra app registration and archive the Hudu asset.

Used when a customer leaves or when you need to fully reset a customer's 365Audit configuration. When run with Hudu parameters, the script resolves the `AppId` and `TenantId` automatically from the existing Hudu asset — no need to look them up manually.

```powershell
# Hudu lookup — resolves AppId/TenantId automatically
.\Helpers\Remove-AuditCustomer.ps1 -HuduCompanyId 'a1b2c3d4e5f6'
.\Helpers\Remove-AuditCustomer.ps1 -HuduCompanyName 'Contoso Ltd'

# Direct — Entra app only, no Hudu interaction
.\Helpers\Remove-AuditCustomer.ps1 -AppId '<AppId>' -TenantId '<TenantId>'

# Preview without making any changes
.\Helpers\Remove-AuditCustomer.ps1 -HuduCompanyId 'a1b2c3d4e5f6' -WhatIf

# Permanently purge from Entra recycle bin immediately
.\Helpers\Remove-AuditCustomer.ps1 -HuduCompanyId 'a1b2c3d4e5f6' -PermanentDelete
```

**Entra app deletion** is a soft-delete by default — the app registration is moved to the Entra recycle bin and can be restored via the Azure portal for up to 30 days. Use `-PermanentDelete` only when certain the customer will not be re-onboarded.

**Hudu asset** is always archived rather than deleted. The asset is hidden from active views but the audit history and credentials are preserved and can be unarchived if needed.

Requires an interactive browser sign-in (Global Administrator) to connect to Graph and perform the deletion.

**Reads from:** `HuduBaseUrl`, `HuduApiKey`, `HuduAssetLayoutId`, `HuduAssetName` in `config.psd1`.

---

## Sync-UnattendedCustomers.ps1

**Purpose:** Automatically populate `UnattendedCustomers.json` from Hudu.

Instead of manually adding each customer to the JSON file, this script queries Hudu for every company that already has a `365Audit` asset (i.e. has been set up via `Setup-365AuditApp.ps1`) and merges them into `UnattendedCustomers.json`. Run it after onboarding a batch of new customers rather than editing the file by hand.

**Merge behaviour:**
- Companies already in the file are left untouched — their `Modules` config is preserved
- New companies are appended with the `DefaultModules` value (default: `A` — Run All)
- Companies in the file that no longer have a Hudu asset are flagged as warnings but are **not** removed — manual review required before deleting

```powershell
# Sync all customers, new entries default to Run All
.\Helpers\Sync-UnattendedCustomers.ps1

# New entries default to Entra + Exchange + Mail Security only
.\Helpers\Sync-UnattendedCustomers.ps1 -DefaultModules '1','2','4'

# Preview without writing
.\Helpers\Sync-UnattendedCustomers.ps1 -WhatIf
```

**Reads from:** `HuduBaseUrl`, `HuduApiKey`, `HuduAssetLayoutId` in `config.psd1`.

---

## Uninstall-AuditModules.ps1

**Purpose:** Clean-remove all toolkit modules for a fresh reinstall or to resolve version conflicts.

Removes every installed version of every module required by the toolkit — across both `CurrentUser` and `AllUsers` scopes where found. Typically used when `Get-ModuleVersionStatus.ps1` shows `MULTIPLE VERSIONS` or when a module update has introduced a dependency conflict that a simple `Update-Module` cannot resolve.

```powershell
# Preview what would be removed
.\Helpers\Uninstall-AuditModules.ps1 -WhatIf

# Remove everything
.\Helpers\Uninstall-AuditModules.ps1
```

> Run in a **fresh PowerShell session** that has not yet loaded any 365Audit scripts — loaded modules cannot be uninstalled while in use. If any modules were installed in `AllUsers` scope, run as Administrator.

After running, reinstall with:

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
Install-Module ExchangeOnlineManagement -Scope CurrentUser
Install-Module PnP.PowerShell -Scope CurrentUser
```
