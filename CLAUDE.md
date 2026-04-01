# 365Audit — Project Reference

## Architecture

PowerShell 7.4+ toolkit for MSP monthly M365 auditing. Connects to customer tenants via app-only cert auth, collects data across modules, generates HTML summary + Hudu report.

## Key Scripts & Entry Points

| Script | Purpose |
|---|---|
| `Start-365Audit.ps1` | Interactive launcher — resolves Hudu creds, decodes cert, presents module menu |
| `Start-UnattendedAudit.ps1` | Batch runner — loops `UnattendedCustomers.psd1`, cert check + audit + Hudu publish per tenant |
| `Setup-365AuditApp.ps1` | First-time app registration, cert generation, renewal. Has own Graph module loader (separate from Audit-Common) |
| `Generate-AuditSummary.ps1` | Builds HTML summary + Hudu report from Raw/ CSVs. Contains Hudu helper functions |
| `Common/Audit-Common.ps1` | Shared functions dot-sourced by all modules |

## Common/Audit-Common.ps1 — Key Functions

| Function | Line hint | Purpose |
|---|---|---|
| `Connect-MgGraphSecure` | ~345 | Connects to Graph. Reads `$AuditAppId`, `$AuditTenantId`, `$AuditCertFilePath`, `$AuditCertPassword` from parent scope via `Get-Variable` |
| `Initialize-GraphSdk` | ~281 | Preloads Graph SDK assemblies to prevent MSAL version conflicts |
| `Initialize-GraphDependencies` | — | Loads DLLs from Graph module's Dependencies folder |
| `Import-GraphModuleVersioned` | — | Imports a specific Graph sub-module matching the resolved version |
| `Resolve-GraphModuleVersion` | — | Finds the installed Graph module version to use |
| `Connect-ExchangeOnlineSecure` | ~422 | Connects to EXO (app-only or interactive) |
| `Initialize-AuditOutput` | — | Creates per-run output folder, writes OrgInfo.json |
| `Get-GraphOrganizationSafe` | ~318 | Wraps Get-MgOrganization with error handling |

## Hudu Credential Flow (Start-365Audit.ps1 ~215-296)

1. Resolve company from slug: `GET /api/v1/companies?slug=...`
2. Find toolkit asset: `GET /api/v1/assets?company_id=...&asset_layout_id=$_huduAssetLayoutId` (default 67)
3. Extract fields from `$huduAsset.fields`: `Application ID`, `Tenant ID`, `Cert Base64`, `Cert Password`
4. Decode cert to temp .pfx, set `$AuditAppId` / `$AuditTenantId` / `$AuditCertFilePath` / `$AuditCertPassword`
5. These variables are read by `Connect-MgGraphSecure` via `Get-Variable` from parent scope

Helper scripts that need Graph auth should: dot-source `Audit-Common.ps1`, set the `$Audit*` variables, then call `Connect-MgGraphSecure`.

## Hudu Report (Generate-AuditSummary.ps1 ~4709-5150)

Helper functions for the Hudu HTML report (inline-styled, no CSS classes, Hudu sanitizer-safe):

| Function | Purpose |
|---|---|
| `New-HuduKpiTile` | KPI card (label, value, sub-text, colour, delta marker) |
| `New-HuduSection` | Collapsible `<details>` section with numbered navy badge |
| `New-HuduStatGrid` | Flex row of stat tiles |
| `New-HuduTable` | HTML table with optional row limit |
| `New-HuduAiTable` | Left-border callout rows for action items |
| `New-HuduModuleAi` | Per-module action item panel, filtered by category prefix |

Section counter reset: `$script:_huduSectionCounter = 0` before first `New-HuduSection` call (~line 4840).

Header uses WAT gradient: `linear-gradient(135deg, #1e3a5f 0%, #2e5c6e 100%)`.
Dark mode: no forced text colours on headers/labels — uses `opacity` for muted text so Hudu's theme applies.

## Publish-HuduAuditReport.ps1 (~Helpers/)

Publishes report to Hudu asset. Key fields:
- `report_summary` — full Hudu HTML body
- `mfa_coverage`, `secure_score`, `tenant_storage`, `critical_items` — plain text KPI fields
- Delta computation: downloads prior month's zip, extracts AuditMetrics.json + ActionItems.json, computes diffs
- Tile markers: `<!-- TILE_DELTA_MFA -->` etc. replaced with coloured delta spans
- Delta section injected at `<!-- AUDIT_DELTA_INJECT -->`

## SKU Friendly Names

`Get-FriendlySkuName` in `Invoke-EntraAudit.ps1` loads `Resources/SkuFriendlyNames.csv` (Microsoft's official CSV, auto-downloaded if missing). Updated by `Helpers/Update-SkuFriendlyNames.ps1`. Lookup key: `String_Id` column (= SkuPartNumber).

## Version Management

All scripts declare `$ScriptVersion` and `.NOTES Version` — these must match `version.json`. CI enforces consistency via `.github/workflows/version-check.yml`.

**When bumping a version, update ALL THREE locations:**

1. `.NOTES Version` in the `<# ... #>` help block near the top of the script
2. `$ScriptVersion = "x.y.z"` variable assignment (usually within 20 lines of the help block)
3. `version.json` — the matching entry for that script path

**Bump strategy:**

- Patch (x.y.Z): bug fixes, typo corrections, minor tweaks — only when modifications have been tested and confirmed
- Minor (x.Y.0): new features, new parameters, new output fields, behaviour changes
- Major (X.0.0): breaking changes (rare — parameter renames, output format changes)

**Which scripts to bump:** Only bump scripts that were actually modified. Check `git diff --name-only` before committing. Common patterns:

- Hudu report styling changes → `Generate-AuditSummary.ps1`
- New Hudu asset fields → `Helpers/Publish-HuduAuditReport.ps1`
- Auth flow changes → `Setup-365AuditApp.ps1` or `Common/Audit-Common.ps1`
- Batch runner changes → `Start-UnattendedAudit.ps1`
- Data collection changes → the relevant `Invoke-*.ps1` module

**Helper scripts** in `Helpers/` also have `.NOTES Version` and `$ScriptVersion` but are NOT all tracked in `version.json` — only those listed in `version.json` need that third update.

**After bumping versions:** Update `README.md` if the change affects documented behaviour, parameters, setup steps, or the file structure. Update `CHANGELOG.md` with a summary of what changed under the new version number.

## Config

`config.psd1` (gitignored) — keys: `HuduBaseUrl`, `HuduApiKey`, `HuduAssetLayoutId` (67), `HuduReportLayoutId` (68), `HuduAssetName`, `HuduReportAssetName`, `MspDomains`, `KnownPartners`, `OutputRoot`, `CleanupLocalReports`.

## Conventions

### PowerShell Standards

- Use **approved verbs** (`Get-Verb` for the full list) — e.g. `Invoke-`, `New-`, `Update-`, `Test-`, not `Run-`, `Create-`, `Check-`
- `#Requires -Version 7.2` minimum on all scripts; `7.4` where PnP.PowerShell v3 is needed
- `[CmdletBinding()]` on all functions and scripts that accept parameters
- `$ErrorActionPreference = 'Stop'` at script level; use `-ErrorAction Stop` on individual calls when inside try/catch
- Use `Write-Host` for user-facing status, `Write-Verbose` for debug detail, `Write-Warning` for non-fatal issues, `Write-Error` for fatal stops
- Prefer `[PSCustomObject]@{}` over `New-Object PSObject` for output objects

### Security & Secrets

- **Never commit secrets** — API keys, cert passwords, tenant IDs, customer data stay in `config.psd1` / `UnattendedCustomers.psd1` (both gitignored)
- Only `.example` files are pushed (e.g. `config.psd1.example`, `UnattendedCustomers.psd1.example`)
- `$CertPassword` is always `[SecureString]` — never accept plain text passwords as parameters
- Temp cert `.pfx` files written to `$env:TEMP` are deleted in `finally` blocks
- No credentials in `Write-Host` / `Write-Verbose` output

### Code Organisation

- **Common functions** go in `Common/Audit-Common.ps1` — auth, Graph SDK init, output folder creation, shared utilities
- **Helper scripts** go in `Helpers/` — standalone utilities that are not part of the audit pipeline
- **Resource files** go in `Resources/` — data files like `SkuFriendlyNames.csv` (gitignored, auto-downloaded)
- Avoid duplicating logic — if two modules need the same function, move it to `Audit-Common.ps1`
- Each `Invoke-*.ps1` module is self-contained for its data collection; shared auth/output is in Common

### Hudu Report HTML

- All HTML in Hudu report uses inline styles (no `<style>` blocks, no CSS classes) — Hudu's sanitiser strips them
- RGBA colours for dark-mode transparency; no hardcoded text colours on content elements — use `opacity` for muted text
- WAT colour palette: `#1e3a5f` (accent navy), `#2e5c6e` (accent-light teal), `#059669` (good), `#d97706` (warn), `#dc2626` (bad)
- Sidecar JSON files: `AuditMetrics.json`, `ActionItems.json` written alongside reports for delta tracking

### Naming & Style

- Script-scoped variables: prefix with `$script:` or `$_` for internal/temp variables (e.g. `$_configPath`, `$_huduHtml`)
- Function naming: `Verb-Noun` with descriptive nouns (e.g. `Connect-MgGraphSecure`, `New-HuduKpiTile`)
- Parameter naming: PascalCase, descriptive (e.g. `$HuduCompanyId`, `$OutputRoot`)
- Comment headers: `# ── Section Name ──────────` with box-drawing characters for major sections
- No trailing whitespace; UTF-8 encoding on all output files
