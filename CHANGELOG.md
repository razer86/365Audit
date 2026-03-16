# Changelog

All notable changes to each script in the 365Audit toolkit are documented here.

---

## Start-365Audit.ps1

| Version | Notes |
|---------|-------|
| 2.9.0 | Console output is now saved to `AuditLog.txt` in the customer audit output folder; transcript is started early (before any output) and moved to the output folder in the `finally` block; skipped automatically when called from `Start-UnattendedAudit.ps1` (detected via `AUDIT_PARENT_TRANSCRIPT` env var) |
| 2.8.0 | Added `-Modules` parameter: when supplied, skips the menu entirely for non-interactive/automated runs; HTML summary is generated but not opened (passes `-NoOpen` to `Generate-AuditSummary.ps1`); expired certificate now causes an immediate hard stop in all modes (was warn-and-continue); near-expiry cert days remaining passed to summary as `-CertExpiryDays` so a Toolkit / Certificate action item appears in the report |
| 2.7.0 | Added `-HuduCompanyId` and `-HuduCompanyName` parameter sets — credentials (AppId, TenantId, CertBase64, CertPassword) are fetched automatically from the NeConnect Audit Toolkit Hudu asset; `-HuduBaseUrl` defaults to `$env:HUDU_BASE_URL` then `https://neconnect.huducloud.com`; added system clock drift check on startup (HEAD request to `login.microsoftonline.com`, warn >60s, stop >300s); disconnect existing Graph/EXO/SPO sessions at startup and in `finally` block to prevent stale connections when re-running in the same PS session; removed noisy "Loading local script" print for local module loads |
| 2.6.1 | Linux/macOS: temp dir falls back to `$env:TMPDIR` then `/tmp` when `$env:TEMP` is absent; `X509KeyStorageFlags` is platform-guarded — Windows keeps `EphemeralKeySet` (key stays in memory only), Linux/macOS use `Exportable\|PersistKeySet` (required by .NET on non-Windows) |
| 2.6.0 | Validate CertBase64 decodes cleanly before writing to disk (clear error if paste is truncated); check certificate expiry on startup and warn if ≤30 days remaining or already expired |
| 2.5.0 | `-CertBase64` is now optional; if omitted the script prompts `Read-Host 'Paste certificate Base64'` — same UX as `-CertPassword` |
| 2.4.0 | Removed `-CertFilePath` and `-CertSharePointUrl`; `-CertBase64` is now the only cert input method — paste from Hudu, no file path or SharePoint URL needed; temp .pfx written to `$env:TEMP` and deleted on exit |
| 2.3.0 | Add `-CertBase64` parameter set; add `-CertSharePointUrl` (browser open) |
| 2.2.0 | Add `-CertSharePointUrl` parameter set: opens SharePoint URL in browser, waits for tech to save .pfx to script folder, deletes on exit via try/finally; remove stale `-PnPAppId` warning |
| 2.1.0 | Converted to certificate-based auth; removed `-AppSecret`/`-PnPAppId` |
| 1.9.0 | `Generate-AuditSummary.ps1` removed from menu Script arrays; summary now runs once after all selected modules complete to avoid multiple report generations when selecting e.g. "1,2" |
| 1.8.0 | `AppId`, `AppSecret`, and `TenantId` are now mandatory parameters with HelpMessage guidance; PnPAppId warning displayed at startup when not provided; removed unnecessary `Start-Sleep` stalls |
| 1.7.0 | Added `Generate-AuditSummary.ps1` to option 9 script list so "Run All" automatically generates the summary report after all modules complete |
| 1.6.0 | Added `-PnPAppId` parameter for the dedicated PnP interactive auth app registered by `Setup-365AuditApp.ps1` (`Register-PnPEntraIDAppForInteractiveLogin`) |
| 1.5.0 | Reverted SharePoint to interactive auth; removed `-CertThumbprint` parameter and SharePoint skip block |
| 1.4.0 | Added `-CertThumbprint` parameter; SharePoint audit now requires a certificate (SharePoint admin APIs reject client-secret tokens) |
| 1.3.0 | SharePoint module now skipped with setup guidance when app credentials are not supplied at launch |
| 1.2.0 | Added optional `-AppId`/`-AppSecret`/`-TenantId` parameters for app-only SharePoint authentication (MSP cross-tenant support) |
| 1.1.0 | Removed duplicate helper functions (moved to `Audit-Common.ps1`); fixed menu option 9 script name; added version check on startup; added `Invoke-MailSecurityAudit.ps1` as option 4 |
| 1.0.1 | Standardised comments; pass folder to summary |
| 1.0.0 | Initial release |

---

## Setup-365AuditApp.ps1

| Version | Notes |
|---------|-------|
| 2.5.3 | Added `DelegatedAdminRelationship.Read.All` to required Graph permissions — enables GDAP/partner relationship collection in `Invoke-EntraAudit.ps1` |
| 2.5.2 | `Connect-GraphForSetup` now runs `Connect-MgGraph` in a `Start-ThreadJob` with a 120-second countdown timer; if the browser sign-in is not completed in time the job is cancelled and an error is thrown, allowing unattended callers to fail fast and continue to the next customer rather than hanging indefinitely |
| 2.5.1 | `Invoke-PermissionCheck` now handles `Authorization_RequestDenied` from `Update-MgApplication` gracefully — when running app-only and the app lacks write permission to update itself (bootstrapping case: `Application.ReadWrite.OwnedBy` not yet granted), disconnects the app-only session and reopens an interactive browser login, then retries the permission update under the Global Admin session; subsequent steps (`Grant-AdminConsent`, `Invoke-OwnerCheck`) also run under the elevated session |
| 2.5.0 | Refactored main body into three explicit modes (Mode 1: no params — interactive, no Hudu push; Mode 2: `-HuduCompanyId`/`-HuduCompanyName` — app-only for existing assets, interactive for first-time with Hudu push; Mode 3: explicit `-AppId`/`-TenantId` — app-only, no Hudu push); extracted `Invoke-PermissionCheck`, `Invoke-OwnerCheck`, and `New-EntraApp` as shared helper functions; `Invoke-PermissionCheck` and `Invoke-OwnerCheck` now called on every run in all three modes — no longer conditional on cert status; eliminated `$script:OwnerCheckOnly` flag; consolidated `Push-HuduAuditAsset` to a single call site per code path |
| 2.4.5 | Non-interactive block now checks and applies missing permissions on every run (same as interactive path) — ensures toolkit updates that add new permissions are applied automatically; fixed double-load Windows CNG crash: `$fetchedCert` from the Hudu expiry check is reused directly instead of reloading the same PFX bytes with `EphemeralKeySet`, which caused `ERROR_PATH_NOT_FOUND` on the second import |
| 2.4.4 | First-time run with `-HuduCompanyId`/`-HuduCompanyName` now works correctly — `Get-HuduAuditCredentials` returns `$null` (instead of throwing) when no Hudu asset or missing fields are found, causing the script to fall through to the interactive browser-login path which creates the app, generates the cert, and pushes credentials to Hudu; this makes `-HuduCompanyId` the standard invocation for both first-time setup and subsequent cert renewals |
| 2.4.3 | Hudu healthy-cert path no longer returns early before attempting the SP owner assignment; falls through to the non-interactive Graph block which tries `New-MgApplicationOwnerByRef` using app-only auth — if `Application.ReadWrite.OwnedBy` is insufficient to self-assign ownership, a clear warning directs the user to run once interactively as Global Admin to bootstrap; `$script:OwnerCheckOnly` flag skips cert generation on the health-check path |
| 2.4.2 | Service principal owner check now runs unconditionally on every re-run of an existing app — previously only ran when permissions were missing, causing existing apps to never get the SP owner assignment needed for non-interactive cert renewal |
| 2.4.1 | Generated `.pfx` files are now automatically deleted in the `finally` block after setup completes — cert base64 is already in Hudu and printed to screen, so the on-disk file is redundant |
| 2.4.0 | Non-interactive cert renewal: added `-AppId`/`-TenantId`/`-CertBase64`/`-CertPassword` params — when all four are supplied, connects to Graph app-only using the existing cert and renews without a browser login; added `Get-HuduAuditCredentials` helper — when only `-HuduCompanyId`/`-HuduCompanyName` are provided, fetches credentials from the Hudu asset, checks cert expiry, and triggers non-interactive renewal if ≤30 days (or `-Force`); added `Application.ReadWrite.OwnedBy` to Graph permissions; service principal now registered as owner of the app registration during both new-app creation and existing-app update paths (required for `OwnedBy` to allow self-renewal) |
| 2.3.0 | Added `-HuduCompanyId`, `-HuduCompanyName`, `-HuduBaseUrl`, `-HuduApiKey` parameters; after certificate generation, `Push-HuduAuditAsset` automatically creates or updates the NeConnect Audit Toolkit asset for the specified company using the Hudu REST API (`custom_fields` with snake_case field names); asset named `NeConnect Audit Toolkit - <Company Name>`; Powershell Launch Command field stores both the manual and Hudu-based invocation as HTML; Hudu push is non-fatal (warns and continues on API error); renamed `Ensure-ServicePrincipal` → `Resolve-ServicePrincipal` (approved verb) |
| 2.2.1 | Linux/macOS: certificate generation now uses `openssl` (`req` + `pkcs12 -legacy`) instead of `New-SelfSignedCertificate`/`Export-PfxCertificate` which are Windows-only; Windows path unchanged; `$rawData` replaces `$cert.RawData` reference for shared Graph upload call |
| 2.2.0 | `New-AuditCertificate` now returns `CertBase64` (base64-encoded .pfx bytes); `Show-Credentials` displays base64 instead of file path; example run command uses `-CertBase64`; both secrets (base64 + password) stored in Hudu — no file path or SharePoint URL needed at audit time |
| 2.1.0 | Add `Sites.FullControl.All` (SharePoint Online app permission) to main app; remove `Register-PnPInteractiveApp` and separate PnP app registration; SharePoint audit now uses the same certificate as Graph/Exchange |
| 2.0.0 | Replace client secret with certificate-based auth: removes `New-AuditSecret`/`Get-SecretStatus`; adds `New-AuditCertificate` (self-signed, CSP key provider, exports .pfx with random password, uploads public key to Entra app); replaces `-SecretExpiryMonths` with `-CertExpiryYears`; `Show-Credentials` now outputs cert path and password for storage in Hudu |
| 1.9.4 | Fix `Set-ExchangeAdminRole`: add `-All` to `Get-MgDirectoryRoleMember` to prevent pagination from missing existing members; catch already-exists error as fallback; only call `Set-ExchangeAdminRole` when Exchange perms are actually missing |
| 1.9.3 | Fix existing-app permission check: now compares per individual permission ID so new permissions (e.g. `SecurityEvents.Read.All`) are applied on re-run without requiring a full fresh setup |
| 1.9.2 | Added `SecurityEvents.Read.All` to Graph permissions for Identity Secure Score |
| 1.9.1 | Updated URLs and references from 'MSA Audit Toolkit' to '365Audit' for branding consistency |
| 1.9.0 | Removed SharePoint (`Sites.FullControl.All`) from main app registration; SharePoint auth is handled entirely by the PnP interactive app which uses delegated permissions scoped to the signed-in technician's rights |
| 1.8.0 | Replaced http://localhost approach with `Register-PnPEntraIDAppForInteractiveLogin` (the PnP-recommended method since Sep 2024); dedicated PnP interactive app is registered separately and its App ID is shown alongside main app credentials; bumped `#Requires` to 7.4; updated PnP module check to enforce `MinimumVersion 3.0.0` |
| 1.7.0 | Removed PnP Management Shell registration (app deprecated in PnP.PowerShell v2); added `http://localhost` as a public client redirect URI to the app registration so `Connect-PnPOnline -Interactive -ClientId $AuditAppId` works from any machine |
| 1.6.0 | Removed certificate management (SharePoint reverted to interactive auth); updated description to reflect dual purpose: app credentials for silent Entra/Exchange auth + PnP Management Shell consent for SharePoint interactive |
| 1.5.0 | Added `Register-PnPManagementShell`: creates a service principal for the PnP Management Shell public app and grants tenant-wide admin consent for `AllSites.FullControl` |
| 1.4.0 | Added `New-AuditCertificate`: generates a self-signed certificate, installs it in `Cert:\CurrentUser\My`, and uploads the public key to the Azure AD app |
| 1.3.0 | Added `Exchange.ManageAsApp` permission and Exchange Administrator Entra role assignment so Exchange Online can authenticate via client credentials |
| 1.2.0 | Added `Request-AdminConsent`: opens Azure portal to API permissions page and prints instructions after consent is granted |
| 1.1.0 | Added Microsoft Graph application permissions required for app-only auth; existing apps without Graph permissions are updated automatically |
| 1.0.0 | Initial release |

---

## Invoke-EntraAudit.ps1

| Version | Notes |
|---------|-------|
| 1.11.0 | Added GDAP/partner relationship collection (`Entra_PartnerRelationships.csv`) — active delegated admin relationships fetched via `Invoke-MgGraphRequest` to `/tenantRelationships/delegatedAdminRelationships`; gracefully skips with a clear warning if `DelegatedAdminRelationship.Read.All` is not yet granted; added third-party enterprise app consent collection (`Entra_EnterpriseApps.csv`) — all non-Microsoft service principals tagged as enterprise apps with admin-consented API permissions; `$totalSteps` updated from 12 to 14 |
| 1.10.3 | Guard against null `controlName` in both the profile title lookup and the `controlScores` loop — prevents "array index evaluated to null" error on tenants where the API returns controls with missing names |
| 1.10.2 | Add `ConvertTo-ReadableControlName` fallback for controls missing a profile title: strips vendor prefixes (`mdo_`, `AATP_`, `AAD_`, etc.), splits underscores and camelCase, title-cases the result; also skip null/empty titles from the profiles API |
| 1.10.1 | Fetch `secureScoreControlProfiles` to resolve human-readable control titles; `ControlName` in CSV now uses title (e.g. "Require MFA for admins") instead of API key (e.g. "AdminMFAV2"); falls back to `controlName` if profile not found |
| 1.10.0 | Add Identity Secure Score collection: `Entra_SecureScore.csv` (date, current, max, percentage) and `Entra_SecureScoreControls.csv` (per-control name, score, description); requires `SecurityEvents.Read.All` |
| 1.9.0 | `Write-Progress -Status` now includes "Step X/Y — " prefix |
| 1.8.0 | Replaced per-section `Write-Host` progress lines with `Write-Progress` for cleaner terminal output |
| 1.7.0 | Guest users: add `LastSignIn` via `SignInActivity` property; CA policies: add `ClientAppTypes` for legacy auth detection |
| 1.6.0 | Format all audit and sign-in timestamps as "yyyy-MM-dd HH:mm UTC" at collection time for consistent timezone display |
| 1.5.0 | Add directory audit log collection scoped to tenant retention window: account creations, deletions, and notable events (role changes and security info/MFA changes) |
| 1.4.0 | Remove AAD P1 gate on sign-in log retrieval; always collect sign-ins (free tenants get 7 days, premium gets 30); store up to 10 entries per user; export `Entra_SignIns.csv` |
| 1.3.0 | Filter member-only users (exclude `#EXT#` guests); add `AccountEnabled`/`AccountStatus` column; fix MFA hashtable Count bug; split output into `Entra_Users.csv` (licensed) and `Entra_Users_Unlicensed.csv`; add SPB and RMSBASIC to SKU friendly-name map |
| 1.2.0 | Graph SDK v2 cmdlet rename: `Get-MgConditionalAccessPolicy` → `Get-MgIdentityConditionalAccessPolicy` |
| 1.1.0 | Added CA policies, trusted locations, security defaults; fixed `LastSignIn` property assignment; fixed Groups success message; removed alias usage; added `CmdletBinding` |
| 1.0.2 | Combine user info, license, MFA, and sign-in into a single export |
| 1.0.1 | Refactor output directory initialisation |
| 1.0.0 | Initial release |

---

## Invoke-ExchangeAudit.ps1

| Version | Notes |
|---------|-------|
| 1.10.0 | Added mail connector collection (`Exchange_MailConnectors.csv`) — inbound and outbound connectors with direction, name, enabled status, type, source, sender domains, and TLS certificate name; `$totalSteps` updated from 15 to 16 |
| 1.9.2 | Use `ExchangeGuid` instead of `PrimarySmtpAddress` for `Get-MailboxStatistics` identity (unambiguous for linked/duplicate mailboxes); suppress `Get-InboxRule` warnings for broken rules via `-WarningVariable` capture; broken rules exported to `Exchange_BrokenInboxRules.csv` with mailbox, rule name, and status |
| 1.9.1 | Wrap primary `Get-MailboxStatistics` in try/catch; null-guard `TotalItemSize` and `ItemCount` so a single inaccessible mailbox no longer aborts the entire inventory |
| 1.9.0 | Suppress EXO object-not-found warning from `Get-DkimSigningConfig` (caught by try/catch; warning was still emitted before the exception); suppress `Get-MailboxCalendarConfiguration` Events-from-Email deprecation warning |
| 1.8.0 | Added Step X/Y counter to `Write-Progress` status strings |
| 1.7.0 | Replaced per-section `Write-Host` progress lines with `Write-Progress` for cleaner terminal output |
| 1.6.0 | Exchange Online now uses app-only auth (via `Connect-ExchangeOnlineSecure`) when `-AppId`/`-AppSecret`/`-TenantId` are provided at launch |
| 1.5.0 | Mailbox inventory adds `LitigationHoldEnabled`; new sections for outbound spam auto-forward policy, shared mailbox sign-in status, Safe Attachments, and Safe Links (gracefully skipped if Defender for Office 365 P1 not licensed) |
| 1.4.0 | Mailbox inventory now includes `LimitMB`, `FreeMB`, `ArchiveSizeMB`; all three permission CSVs now include `MailboxUPN` for consistent joining; SendAs now loops per mailbox (REST mode compatible) |
| 1.3.0 | EXO connection check now filters by `State eq Connected`; stale sessions no longer prevent reconnection; `Import-Module` uses `-ErrorAction Stop` |
| 1.2.0 | Fixed mailbox size parsing for EXO v3 deserialized `ByteQuantifiedSize`; rewrote `Get-MailboxPermission` to loop per-mailbox with `-Identity`; removed `Get-ReceiveConnector` (on-premises only, not available in EXO) |
| 1.1.0 | Removed duplicate guard clause; fixed `outputDir` override; removed alias usage; added `CmdletBinding` |
| 1.0.2 | Helper function refactor |
| 1.0.1 | Refactor output directory initialisation |
| 1.0.0 | Initial release |

---

## Invoke-SharePointAudit.ps1

| Version | Notes |
|---------|-------|
| 2.8.0 | Switched to certificate-based app-only auth: reads `$AuditAppId`/`$AuditCertFilePath`/`$AuditCertPassword` from launcher scope; uses `Connect-PnPOnline -CertificatePath/-CertificatePassword` (portable .pfx, no cert store required); removes interactive auth, PnP app ID pre-flight, and `Get-PnPAccessToken`; falls back to interactive when cert vars absent |
| 2.7.0 | Replaced `-ReturnConnection` MSAL caching strategy with explicit `-AccessToken` pass-through: authenticate interactively once to the admin URL, capture the SPO access token via `Get-PnPAccessToken`, then connect to each site with `-AccessToken` (no browser prompt per site); also removed `Disconnect-PnPOnline -Connection` which is not a valid parameter in PnP.PowerShell v3 |
| 2.6.0 | Added Step X/Y counter to `Write-Progress` status strings |
| 2.5.0 | Per-site connections use `-ReturnConnection` so each connection is a named object; admin cmdlets pin to `$adminConn`, per-site cmdlets pin to `$siteConn`; guarantees exactly one browser sign-in (MSAL reuses the cached token for all subsequent site connections silently) |
| 2.4.0 | Replaced per-section `Write-Host` progress lines with `Write-Progress` for cleaner terminal output |
| 2.3.0 | PnP Management Shell app deprecated in PnP.PowerShell v2; switched to using `$AuditPnPAppId` as the `ClientId` for `Connect-PnPOnline -Interactive`; bumped `#Requires` to 7.4; updated PnP module check to enforce `MinimumVersion 3.0.0` |
| 2.2.0 | Added pre-flight check: verifies PnP Management Shell is registered in the tenant before attempting interactive auth; prints setup guidance if missing |
| 2.1.0 | Reverted to interactive authentication; certificate-based app-only auth is not portable across technician machines; interactive sign-in is sufficient for manual monthly audit runs |
| 2.0.0 | Root cause identified: SharePoint admin APIs block tokens with `azpacr=0`; switched to PnP.PowerShell with certificate-based app-only auth which produces `azpacr=1` tokens |
| 1.9.0 | Replaced all module dependencies with direct SharePoint REST API calls; enum mapping functions for string compatibility |
| 1.8.0 | Reverted to PnP.PowerShell; connect to admin URL for tenant-wide operations; reconnect per-site for group/user queries |
| 1.7.0 | Replaced PnP.PowerShell entirely with `Microsoft.Online.SharePoint.PowerShell`; `Connect-SPOService -AccessToken`; all PnP cmdlets replaced with SPO equivalents |
| 1.6.0 | Replaced PnP.PowerShell with OAuth2 client-credentials token approach |
| 1.5.0 | Extend `ExternalSharing_Tenant` CSV with `DefaultSharingLinkType` and `RequireAnonymousLinksExpireInDays`; extend `AccessControlPolicies` CSV with sync restriction fields |
| 1.4.0 | Conditional auth: app-only (`ClientId`/`Secret`) when launcher provides `-AppId`/`-AppSecret`/`-TenantId`; falls back to interactive PnP Management Shell |
| 1.3.0 | PnP.PowerShell v2+ requires `-ClientId` with `-Interactive`; using PnP Management Shell public app |
| 1.2.0 | Fixed `Connect-PnPOnline` (removed invalid `-Scopes` param); derive SharePoint admin URL from tenant `.onmicrosoft.com` domain; replaced SPO Management Shell cmdlets with PnP equivalents |
| 1.1.0 | Removed duplicate guard clause; fixed `outputDir` override; replaced deprecated `Get-MsolUser` with Microsoft Graph; removed alias usage; added `CmdletBinding` |
| 1.0.2 | Helper function refactor |
| 1.0.1 | Refactor output directory initialisation |
| 1.0.0 | Initial release |

---

## Invoke-MailSecurityAudit.ps1

| Version | Notes |
|---------|-------|
| 1.6.1 | Linux/macOS: added `Resolve-TxtRecord` helper that uses `dig` when `Resolve-DnsName` is unavailable; DMARC uses `-join ''` to correctly reassemble fragmented TXT records per RFC 7489; SPF fallback label changed from "DNS query failed" to "Not Found" for consistency |
| 1.6.0 | Added Step X/Y counter to `Write-Progress` status strings |
| 1.5.0 | Replaced per-section `Write-Host` progress lines with `Write-Progress` for cleaner terminal output |
| 1.4.0 | Exchange Online now uses app-only auth (via `Connect-ExchangeOnlineSecure`) when `-AppId`/`-AppSecret`/`-TenantId` are provided at launch |
| 1.3.0 | EXO connection check now filters by `State eq Connected`; `Import-Module` uses `-ErrorAction Stop` |
| 1.2.1 | `Export-Json` now skips gracefully on null/empty results |
| 1.2.0 | Integrated with launcher and shared output folder; flat data (DKIM, DMARC, SPF) now exported as CSV; JSON retained for nested policy objects; added guard clause and `CmdletBinding` |
| 1.1.0 | Minor updates |
| 1.0.0 | Initial release (standalone, JSON-only output) |

---

## Generate-AuditSummary.ps1

| Version | Notes |
|---------|-------|
| 1.23.0 | Added action items for previous-MSP detection: guest accounts in privileged admin roles (critical, from `Entra_AdminRoles.csv`); active GDAP/partner relationships (critical, from `Entra_PartnerRelationships.csv`); third-party enterprise apps with admin consent (warning, from `Entra_EnterpriseApps.csv`); enabled custom mail connectors (warning, from `Exchange_MailConnectors.csv`) |
| 1.22.0 | Added critical action item for non-NeConnect Technical Contact domains — if any address in TechnicalNotificationMails is not from ntit.com.au, nqbe.com.au, capconnect.com.au, widebayit.com.au, or neconnect.com.au, a `Tenant / Technical Contact` critical item is raised identifying the address(es) and directing the tech to update the Technical Notification email in the M365 admin centre |
| 1.21.0 | Added `-NoOpen` switch — suppresses automatic browser launch after report generation (used by `-Modules` automation path in launcher); added `-CertExpiryDays` parameter — when 0–30, inserts a `Toolkit / Certificate` warning action item prompting the tech to run `Setup-365AuditApp.ps1 -Force` before next audit run |
| 1.20.0 | Broken inbox rules: new action item and HTML table in Exchange section (reads `Exchange_BrokenInboxRules.csv`; displays mailbox, rule name, and "Broken — not processing mail" status) |
| 1.19.0 | Secure Score control breakdown split into To Action (open) and Implemented (collapsed) sub-tables; To Action controls sorted alphabetically, Implemented sorted by score descending |
| 1.18.0 | Polish pass: center Org Info box; add Secure Score control breakdown table; Teams Rooms Basic SKU display name; CA policy click-to-expand; doc links on all action items; mailbox table sorted `UserMailbox` → `SharedMailbox`; transport rule expander adds Mode description and disclaimer location; Safe Attachments/Links click-to-expand; SharePoint storage bar overflow fix |
| 1.17.0 | Add Identity Secure Score gauge to Entra section (reads `Entra_SecureScore.csv`; colour-coded progress bar) |
| 1.16.0 | Cross-platform report launch: `xdg-open` on Linux, `open` on macOS |
| 1.15.0 | SharePoint section fixes: tenant storage falls back to summing per-site + OneDrive CSVs when `StorageQuotaUsed` is null; claim token parsing for group members; expanded template label map |
| 1.14.0 | Full SharePoint / OneDrive summary section: tenant storage gauge, site collection table with expandable groups panel, external sharing policy + site override table, access control policy settings, OneDrive usage count + unlicensed accounts |
| 1.13.0 | New action items: legacy auth, stale accounts, stale guests, shared mailbox sign-in, outbound spam auto-forward, Safe Attachments/Links, SharePoint default link type, sync restriction; new summary sections for each |
| 1.12.0 | Add Action Items panel: checks Entra MFA, CA enforcement, admin count, SSPR, Exchange forwarding/audit/DKIM/anti-phish, mail security DMARC/SPF, SharePoint external sharing/unlicensed OD |
| 1.11.0 | Remote domains table; audit config human-readable; system mailbox filter; expandable rows for policies/rules/distribution lists |
| 1.10.0 | Exchange section: add summaries for all remaining CSVs |
| 1.9.1 | Replace Used/Free/Limit MB columns with a usage progress bar and limit shown in GB |
| 1.9.0 | Exchange section: full mailbox table (Used/Free/Limit/Archive) with expandable delegated-permissions panel per mailbox |
| 1.8.0 | Add company summary block at top of report from `OrgInfo.json`: company name, address, phone, tech contact, verified domains, Azure AD sync status with staleness colour-coding |
| 1.7.0 | Add audit window duration label (derived from licence tier) to all audit event messages |
| 1.6.0 | Add account creations, account deletions, and notable audit event summaries (role changes, MFA/security info changes) |
| 1.5.0 | Add interactive sign-in history rows; click user row to expand/collapse last 10 sign-ins |
| 1.4.0 | Sort user table by UPN; add Conditional Access policy summary with per-policy state table and licence-based warning when no CA policies are present |
| 1.3.0 | MFA stat counts licensed users only; add `AccountStatus` column to user table; show GA name when count = 1; replace admin role count with full assignment table; add unlicensed member count; consume `Entra_Users_Unlicensed.csv` |
| 1.2.0 | Fixed broken `$OutputPath` and `$latestFolder` variable references; removed duplicate param block; added Mail Security section; added Security Defaults to Entra section |
| 1.1.0 | Added `Entra_RiskyUsers` and `Entra_AdminRoles` summaries |
| 1.0.1 | Updated Entra audit sources |
| 1.0.0 | Initial release |

---

## Start-UnattendedAudit.ps1

| Version | Notes |
|---------|-------|
| 2.2.0 | Console output logging delegated to `Start-365Audit.ps1` — each customer's full run log is saved as `AuditLog.txt` in that customer's audit output folder; no separate bulk transcript needed since `Start-365Audit.ps1` stops its transcript in `finally` before the next customer begins |
| 2.1.0 | Customer list extracted to `UnattendedCustomers.json` — techs edit the JSON file rather than the script; each entry has `HuduCompanySlug` and `Modules` (per-customer module selection); `-Modules` param now acts as a global override for all customers; `-Customers` param filters by slug; summary table includes per-customer modules column; script hard-errors with copy hint if JSON file is not found |
| 2.0.0 | Full rewrite: Hudu-based credential management — customer list uses Hudu company IDs/slugs, no credentials stored in the script; per-customer flow: (1) call `Setup-365AuditApp.ps1 -HuduCompanyId` to check/renew cert automatically, (2) call `Start-365Audit.ps1 -HuduCompanyId -Modules` with fresh credentials from Hudu; supports `-Customers` override, `-Modules` selection, `-SkipCertCheck`, and `-HuduApiKey`/`-HuduBaseUrl`; per-customer error isolation (one failure does not stop remaining customers); final summary table with status per customer |
| 1.0.0 | Initial version (hardcoded credentials per customer, deprecated) |

---

## common/Audit-Common.ps1

| Version | Notes |
|---------|-------|
| 1.14.0 | Added "Connected to Exchange Online." confirmation after both app-only and interactive connect paths; `Connect-ExchangeOnlineSecure` no longer silently skips reconnection for changed tenants — session pre-disconnect is handled by the launcher |
| 1.13.0 | Switch app-only auth from client secret to certificate for both Graph and Exchange Online; `Connect-MgGraphSecure` now uses `X509Certificate2` loaded from `$AuditCertFilePath`; `Connect-ExchangeOnlineSecure` now uses `-CertificateFilePath`/`-CertificatePassword`; removes all OAuth token acquisition code |
| 1.11.0 | `Initialize-AuditOutput`: move output folder from repo root to parent directory to avoid git conflicts on update |
| 1.10.0 | `Connect-ExchangeOnlineSecure`: add missing `-AppId` to `Connect-ExchangeOnline -AccessToken` call; EXO v3 requires `-AppId` alongside `-AccessToken` to recognise the connection as app-only context |
| 1.9.0 | Change "Already connected" `Write-Host` calls to `Write-Verbose` so they do not add scroll lines when modules reuse an existing session |
| 1.8.0 | Add `Connect-ExchangeOnlineSecure`: uses client-credentials OAuth token for Exchange Online when app credentials are present; falls back to interactive |
| 1.7.0 | `Connect-MgGraphSecure` auto-detects `$AuditAppId`/`$AuditAppSecret`/`$AuditTenantId` from the launcher scope; uses app-only (`ClientSecretCredential`) auth when all three are present, falls back to interactive delegated auth otherwise |
| 1.6.0 | Add `AuditLog.Read.All` to required Graph scopes (needed for sign-in logs) |
| 1.5.0 | Auto-loading sub-modules still fails because `RequiredModules` resolution re-triggers a DLL load even when the DLL is in the AppDomain. Fix: import sub-modules explicitly inside `Connect-MgGraphSecure` AFTER `Connect-MgGraph` |
| 1.4.0 | Reverted to install-only at file scope; never explicitly import sub-modules. `SilentlyContinue` does not catch terminating .NET `FileLoadException` |
| 1.3.0 | Moved Graph sub-module loading to file scope (wrong approach) |
| 1.2.0 | Import required Microsoft Graph sub-modules in `Connect-MgGraphSecure` |
| 1.1.0 | Added `CmdletBinding`, `Invoke-VersionCheck`, centralised `RemoteBaseUrl` |
| 1.0.0 | Initial creation and migration of shared helpers from launcher |
