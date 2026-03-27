# Changelog

All notable changes to each script in the 365Audit toolkit are documented here.

---

## Start-365Audit.ps1

| Version | Notes |
|---------|-------|
| 2.15.0 | After a successful audit run, writes the customer output folder path to `$env:TEMP\365Audit_LastOutput.txt` so `Start-UnattendedAudit.ps1` can retrieve it without needing to re-derive the path |
| 2.14.1 | `OutputRoot` validation: resolves to absolute path, checks drive/UNC qualifier is accessible, and attempts `New-Item` before Graph connects — typo fails immediately with the resolved path in the error message |
| 2.14.0 | Added `-OutputRoot` parameter; falls back to `OutputRoot` in `config.psd1`, then the default path two levels above the toolkit; value is propagated to `Initialize-AuditOutput` via `$script:AuditOutputRoot` before the module loop so all module scripts and the summary use the same root |
| 2.13.0 | Run All option key changed from `9` to `A` — frees up single-digit keys 8 and 9 for additional modules; `-Modules` parameter type changed from `[int[]]` to `[string[]]` with `ValidateSet('1'..'7', 'A')`; interactive menu parser updated to accept `A` alongside digits; menu `Sort-Object` updated to sort `'A'` after numeric keys; `$selectedIndexes` coercion converts numeric string inputs to `[int]` for hashtable lookup while preserving `'A'` as string |
| 2.12.0 | Added option 7 "ScubaGear CIS Baseline" (`Invoke-ScubaGearAudit.ps1`); option 9 "Run All" updated to include all seven modules (1–7) |
| 2.11.0 | Config loaded from `config.psd1` (version note — no new features, version bump only) |
| 2.10.0 | Config loaded from `config.psd1` — `HuduApiKey`, `HuduBaseUrl`, and `HuduAssetLayoutId` sourced from file instead of environment variables; Hudu asset fetch now paginates (do/while until empty page — was a single request that missed assets beyond the first page); run context (mode, selected modules, timestamp, script path) written to transcript after module selection |
| 2.9.3 | Audit transcript now saved to `Raw Files\AuditLog.txt`; works with the shared `Raw Files` output layout instead of the `Logs\` subfolder |
| 2.9.2 | Audit transcript now saved to `Logs\AuditLog.txt` inside the customer output folder instead of the root; `Logs\` subfolder is created in the `finally` block before the move |
| 2.9.1 | Added option 5 "Intune / Endpoint Audit" (`Invoke-IntuneAudit.ps1`); option 9 "Run All" updated to include all five modules (1–5); `ValidateSet` on `-Modules` extended to include 5 |
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
| 2.10.0 | Added `PrivilegedAccess.Read.AzureAD` Graph permission (required by ScubaGear for PIM policy checks); added `Set-GlobalReaderRole` helper and calls in both new-app and update-app paths — Global Reader Entra role is now assigned to the service principal automatically (required by ScubaGear non-interactive assessment) |
| 2.9.0 | Config loaded from `config.psd1` (version note — no new features, version bump only) |
| 2.6.0 | Config loaded from `config.psd1` — `HuduApiKey`, `HuduBaseUrl`, `HuduAssetLayoutId`, and `AuditAppName` sourced from file instead of environment variables; `HuduAssetLayoutId` replaces all hardcoded `asset_layout_id=67` references |
| 2.5.5 | Existing-asset consent remediation now distinguishes missing application permissions from missing admin consent, retries consent interactively when app-only assignment is denied, and removes the `Start-ThreadJob` timeout wrapper from the interactive `Connect-MgGraph` fallback to avoid `$using:` timer errors |
| 2.5.4 | Added four Intune Graph application permissions: `DeviceManagementManagedDevices.Read.All`, `DeviceManagementConfiguration.Read.All`, `DeviceManagementApps.Read.All`, `DeviceManagementServiceConfig.Read.All` |
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
| 1.17.2 | Enterprise Apps Application permissions: use `$ra.ResourceDisplayName` as fallback for `ResourceApp` before falling back to raw ResourceId GUID; SP permission resolution switched from `-Property`-based SDK calls (which silently return empty `AppRoles`) to `Get-MgServicePrincipal` without `-Property` + `-ConsistencyLevel eventual` for appId filter queries; ID comparisons use string interpolation to handle `Nullable<Guid>` type differences in SDK v2 |
| 1.17.1 | Added `Add-AuditIssue` calls to 17 outer catch blocks (sign-in logs, account creations/deletions, directory audit events, SSPR, CA policies, named locations, Secure Score, Security Defaults, enterprise apps, risky users/sign-ins, auth methods policy, external collaboration, app registrations, PIM assignments, org settings) — collection failures now logged to `AuditIssues.csv` |
| 1.14.0 | Lazy-load Entra-specific Graph sub-modules via `Import-GraphSubModules` after connect — loads `Users`, `Groups`, `Reports`, `Identity.SignIns`, and `Applications` on demand instead of at startup; adds `Microsoft.Graph.Applications` required for `Get-MgServicePrincipal` and `Get-MgServicePrincipalAppRoleAssignment` |
| 1.13.0 | Output CSVs now written to the shared `Raw Files\` folder inside the customer output directory instead of the `Entra\` subfolder |
| 1.12.0 | Output CSVs written to `Entra\` subfolder inside the customer output directory instead of the root |
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
| 1.15.2 | Removed `-AllUsers` parameter from `Get-ExternalInOutlook` (parameter removed in newer EXO module versions); added server-side error detection in catch block to silently skip with `Write-Verbose` when the API returns "server side error" or "operation could not be completed" (feature unavailable for some tenants) |
| 1.15.1 | Added `Add-AuditIssue` calls to 8 outer catch blocks (Safe Attachments, Safe Links, accepted domains, authentication policies, Exchange org config, external sender tagging, connection filter, OWA mailbox policies) |
| 1.13.0 | Module install now verifies the module is discoverable after `Install-Module` and includes the installed version in the confirmation message; `Install-Module` uses `-ErrorAction Stop` for consistent failure handling |
| 1.12.0 | Output CSVs now written to the shared `Raw Files\` folder inside the customer output directory instead of the `Exchange\` subfolder |
| 1.11.0 | Output CSVs written to `Exchange\` subfolder inside the customer output directory instead of the root |
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
| 2.12.1 | Version bump for consistency with other audit script patch series; no qualifying outer catch blocks (only catch is inside a `ForEach-Object -Parallel` per-site loop — excluded by Add-AuditIssue pattern rules) |
| 2.11.0 | Per-site group/user collection converted to `ForEach-Object -Parallel -ThrottleLimit 5` for concurrent execution; `SecureString` cert password extracted to plain text before the parallel block and reconstructed inside each runspace to avoid serialisation failure; `Connect-PnPOnline -ReturnConnection` now used with explicit `-Connection` on every PnP cmdlet to prevent thread-local connection loss; `Disconnect-PnPOnline` wrapped in `try/catch` in `finally` block; `Microsoft.Graph.Users` loaded on demand for `Get-MgUser` (unlicensed OneDrive detection); module install now verifies post-install and shows installed version |
| 2.10.0 | Output CSVs now written to the shared `Raw Files\` folder inside the customer output directory instead of the `SharePoint\` subfolder |
| 2.9.0 | Output CSVs written to `SharePoint\` subfolder inside the customer output directory instead of the root |
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
| 1.9.1 | Added `Add-AuditIssue` call to Spoof Intelligence catch block (`Get-SpoofIntelligenceInsight`); DKIM per-domain and DNS helper catches excluded (per-item/lookup patterns) |
| 1.9.0 | `Resolve-TxtRecord` now falls back from `dig` to `nslookup` on Linux/macOS when `dig` is unavailable, with a clear warning if neither tool is found; module install now verifies post-install and shows installed version |
| 1.8.0 | Output CSVs and JSON files now written to the shared `Raw Files\` folder inside the customer output directory instead of the `MailSecurity\` subfolder |
| 1.7.0 | Output CSVs and JSON files written to `MailSecurity\` subfolder inside the customer output directory instead of the root |
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

## Invoke-IntuneAudit.ps1

| Version | Notes |
|---------|-------|
| 1.8.1 | Added `Add-AuditIssue` calls to 9 catch blocks (managed devices, compliance policies, modern config policies, config profiles, apps, Autopilot devices, enrollment restrictions, Windows Update Rings, App Protection Policies) |
| 1.7.0 | Lazy-load Intune-specific Graph sub-modules via `Import-GraphSubModules` after connect — loads `DeviceManagement`, `Devices.CorporateManagement`, and `DeviceManagement.Enrollment` on demand; Graph 429 throttle retry added to `Invoke-GraphCollectionRequest` — exponential backoff up to 5 attempts, honouring `Retry-After` header when present |
| 1.6.0 | Intune exports refined after live validation: app install counts now populate correctly; configuration profile setting names/values are more human-readable (including Edge/Startup labels and duplicate-child suppression); Intune assignment details now resolve Entra group display names instead of raw GUIDs |
| 1.5.0 | Output CSVs now written to the shared `Raw Files\` folder inside the customer output directory instead of the `Intune\` subfolder |
| 1.4.0 | Added Intune export enrichment for summary drilldowns: device table source remains `Intune_Devices.csv`; compliance policy exports now include IDs, descriptions, types, assignment details, and normalized setting values; configuration collection now includes modern `deviceManagement/configurationPolicies` via Graph beta with settings and assignment detail; app exports now include IDs, descriptions, timestamps, and assignment details |
| 1.2.0 | Output CSVs written to `Intune\` subfolder inside the customer output directory instead of the root |
| 1.1.0 | Added configuration profile settings collection — iterates `AdditionalProperties` on each `Get-MgDeviceManagementDeviceConfiguration` object (same approach as compliance policy settings) and exports `Intune_ConfigProfileSettings.csv` (one row per setting per profile: `ProfileName`, `Platform`, `ProfileType`, `SettingName`, `SettingValue`); metadata keys skipped via `$_odataSkipKeys` |
| 1.0.0 | Initial release — licence check against known Intune-capable SKUs (skips gracefully if unlicensed); managed device inventory with OS/ownership/compliance/last sync; per-device compliance policy states; compliance policies with platform, assignment scope, grace period (hours), and full `AdditionalProperties` key-value settings; configuration profiles with platform, type, last modified, and assignments; assigned app install summary (installed/failed/pending counts); Windows Autopilot device identities (skips gracefully on 403); enrollment restrictions |

---

## Invoke-TeamsAudit.ps1

| Version | Notes |
|---------|-------|
| 1.1.2 | `Get-CsTeamsChannelPolicy` catch block now silently skips with `Write-Verbose` on `CommandNotFoundException` or "not recognized" errors — command was removed in newer MicrosoftTeams module versions; previously raised a spurious Warning audit issue |
| 1.1.1 | Added `Add-AuditIssue` calls to all 9 step catch blocks (federation config, client config, meeting policies, guest meeting config, guest calling config, messaging policies, app permission policies, app setup policies, channel policies) |
| 1.1.0 | Replaced `Write-StepProgress` helper and `Write-Host` step output with `Write-Progress -Id 1` matching the pattern used by all other audit modules; added `New-Item -Force` output directory creation; completion message matches other modules |
| 1.0.0 | Initial release — federation config, client config, meeting policies, guest meeting/calling config, messaging policies, app permission/setup policies, channel policies |

---

## Invoke-ScubaGearAudit.ps1

| Version | Notes |
|---------|-------|
| 1.3.3 | Added `5>$null 6>$null` stream redirections to `Import-Module ScubaGear` and `Initialize-SCuBA` inside the PS 5.1 subprocess to suppress DEBUG output noise (ScubaGear sets module-scoped `$DebugPreference = 'Continue'` which bypasses the script-level preference); moved `Write-Progress -Completed` to before `Start-Process` so the PS 7 progress bar is cleared before the subprocess writes to the shared console window, preventing garbled output |
| 1.3.2 | Remove `-LogIn $false` from `Invoke-SCuBA` — ScubaGear's `Invoke-Connection` gate skips `Connect-Tenant` entirely when `LogIn=$false`, silently ignoring `-AppID`/`-CertificateThumbprint` and causing "Authentication needed" on every Graph call; omitting the flag lets `LogIn` default to `$true` so ScubaGear calls `Connect-MgGraph` with the cert thumbprint as intended |
| 1.3.1 | Clear stale MSAL token cache before `Invoke-SCuBA`: disconnect any existing Graph/EXO sessions and delete `$env:USERPROFILE\.graph` inside the PS 5.1 subprocess; fixes "Authentication needed. Please call Connect-MgGraph" failures when a cached token for a different tenant exists on disk (pattern from Galvnyz/M365-Assess) |
| 1.3.0 | Replaced inherited PSModulePath filtering with an explicit clean WinPS 5.1 module path (`Documents\WindowsPowerShell\Modules`, `Program Files\WindowsPowerShell\Modules`, `System32\WindowsPowerShell\v1.0\Modules`); PS 7's inherited path omits the user WindowsPowerShell folder entirely, so filtering alone left `Import-Module` unable to find the freshly installed ScubaGear; user module folder is created if absent |
| 1.2.0 | PS 7 module paths stripped from `$env:PSModulePath` at the start of the WinPS 5.1 subprocess so ScubaGear is installed into and loaded from `Documents\WindowsPowerShell\Modules` (PS 5.1 path) — prevents PackageManagement/PowerShellGet version conflicts caused by PS 7-installed modules being visible to PS 5.1; added `-SkipModuleCheck` to `Invoke-SCuBA` to bypass ScubaGear's internal dependency version checks after installation |
| 1.1.0 | ScubaGear now runs in a Windows PowerShell 5.1 subprocess to avoid module-version conflicts with the PS 7 modules loaded by the 365Audit session; cert is imported into `Cert:\CurrentUser\My` in PS 7 (store is shared), then the temp script is spawned via `powershell.exe` with all values passed as named parameters; temp script is deleted in the `finally` block alongside cert cleanup; `powershell.exe` availability checked at startup |
| 1.0.0 | Initial release — auto-installs/updates ScubaGear from PSGallery; calls `Initialize-SCuBA` to ensure OPA binary is present; bridges the .pfx cert to `Cert:\CurrentUser\My` for ScubaGear's thumbprint-based auth, then removes it in a `finally` block; runs `Invoke-SCuBA` for AAD, Defender, EXO, SharePoint, Teams (Power Platform excluded — requires interactive one-time registration); writes output to `Raw Files\ScubaGear_<timestamp>\`; collection failures written via `Add-AuditIssue` |

---

## Generate-AuditSummary.ps1

| Version | Notes |
|---------|-------|
| 1.50.0 | Month-over-month delta foundation: `Add-ActionItem` gains a `CheckId` parameter (stable `MODULE-SUBCATEGORY-NNN` identity string) stored in `ActionItems.json` and used by `Publish-HuduAuditReport.ps1` for action item diffing; all 81 non-ScubaGear call sites updated with stable IDs; `AuditMetrics.json` written alongside `ActionItems.json` each run — captures MFA coverage %, Secure Score, device count, storage GB/%, licence assigned/available, and action item counts; `New-HuduKpiTile` gains `-DeltaMarkerId` parameter that emits `<!-- TILE_DELTA_* -->` HTML comment markers inside each KPI tile; `<!-- AUDIT_DELTA_INJECT -->` marker inserted between KPI row and Action Items in Hudu HTML for downstream injection of the delta section |
| 1.49.0 | Hudu report (`M365_HuduReport.html`) rewritten as a pure HTML fragment compatible with Hudu's rich-text field renderer: removed `<!DOCTYPE html>` wrapper and `<style>` block; all surface backgrounds converted to `rgba()` for light/dark theme compatibility; borders use `rgba(128,128,128,0.2)`; semantic state backgrounds use `rgba()` (good: green/0.1, bad: red/0.1); header banner accent updated to `#1849a9`; section accent default updated to `#1849a9`; action item Category and Finding cells now inherit page text colour (fixes grey-on-dark readability); added **Tenant Storage** KPI tile to both the main report strip and the Hudu report header row — reads `SharePoint_TenantStorage.csv`, falls back to summing per-site and OneDrive CSVs, colour-coded green/amber/red at 75%/90% thresholds; shows `—` when SharePoint module was not run |
| 1.43.0 | Exchange / Mailboxes action item: flags mailboxes that are over 75% full and have no In-Place Archive enabled, listing each affected mailbox with display name, UPN, and usage percentage; mailbox table row highlighting: near-full rows rendered with an amber left border and `#fff8f0` background; Archive column replaced with a styled "No Archive" label (orange, bold, with tooltip) for flagged rows |
| 1.42.0 | Compliance Overview: fixed Passed count always showing 0 when total issues exceeded the hardcoded 60-check total — total is now `max(150, critCount + warnCount + 30)` so Passed is always a meaningful non-zero value; sidebar navigation now uses `New-SbSub` helper to skip links for sections whose source CSV was not generated (prevents dead anchor links when a module was not run); Secure Score progress bar: added `margin-right:8px` and removed orphaned `&nbsp;` that caused the date to run into the bar; CSS: added `td ul, td ol { overflow-wrap: break-word }` to prevent bullet-point lists overflowing table cell boundaries; External Collaboration Settings: GUIDs translated to human-readable labels for `AllowInvitesFrom` and default guest role; displayed as a formatted table instead of raw values; App Registrations: collapsible permissions dropdown per app row loaded from `Entra_AppRegistrationPermissions.csv`; Enterprise Apps: collapsible permissions dropdown per app row loaded from `Entra_EnterpriseAppPermissions.csv`; removed separate "Consented Roles" count column; Stale Licensed Accounts: removed as a separate section — stale users now highlighted in-place in the User Accounts table with red bold Last Sign-In and a mouseover tooltip; Groups table: added clarification note distinguishing Role-Assignable from Dynamic membership groups; OneDrive: added collapsible per-user storage table below the usage count line; SharePoint storage stat chip fallback: when `StorageUsedMB = 0` from the API, sums per-site and OneDrive CSVs to compute an accurate used-storage figure; ScubaGear section collapsed by default; added CISA attribution note linking to the ScubaGear project page; CIS Foundations Benchmark references added to Org User Settings, Exchange Org Config, SharePoint External Sharing, Access Control Policies, Teams External Access, and Teams Client Config subsections |
| 1.41.0 | ScubaGear integration: detects `Raw Files\ScubaGear_*\ScubaResults_*.json`; adds failing Shall controls as critical action items and failing Should/Warning controls as warnings under `ScubaGear / <product>` categories; adds ScubaGear Baseline section with per-product pass/fail/warning counts and link to ScubaGear HTML report; sidebar entry added conditionally when ScubaGear output is present |
| 1.40.0 | All `<h4>` subsections within module sections are now collapsible via JS-generated `<details>/<summary>` elements (open by default); module section order changed to follow CIS M365 Foundations Benchmark chapter sequence: Entra (CIS 1) → Exchange (CIS 3/6) → Mail Security (CIS 2) → SharePoint (CIS 7) → Teams (CIS 8) → Intune (CIS 5); sidebar navigation order updated to match |
| 1.39.0 | Stat chips in all six module section headers are now clickable anchor links — each chip navigates to the relevant subsection (e.g. MFA chip → `#entra-users`, Forwarding chip → `#exchange-forwarding`); company info block moved from standalone card into the coloured top bar with domain, tenant ID, and report metadata; removed standalone `Microsoft 365 Audit Summary` heading; removed `$_companyCard` from content area |
| 1.38.0 | Action Items block is now collapsible — wrapped in a clickable header showing the critical/warning count summary; uses the existing `toggleModule` JS; header background and hover styling match the sidebar aesthetic |
| 1.37.0 | Added section header stat chips to all six module sections (Entra: users/MFA/GAs; Exchange: mailboxes/shared/forwarding; SharePoint: sites/storage/external; Mail Security: SPF/DMARC/DKIM coverage %; Intune: device/compliant/non-compliant; Teams: federation/guest/channels); added Compliance Overview section (distribution bar showing pass/warn/critical, per-module breakdown, CIS controls with findings); added Technical Issues section that renders AuditIssues.csv (written by Add-AuditIssue in catch blocks); added Compliance Overview and Technical Issues links to sidebar nav; added CSS for `.section-stats`, `.stat-chip`, `.cov-bar`, and `.issue-sev-*` styles |
| 1.36.0 | CIS benchmark alignment: added CIS Microsoft 365 Foundations Benchmark v6.0.1 reference IDs to all applicable existing action items; added new action items for previously unchecked controls — CIS 1.2.3 (users can create tenants), CIS 1.3.6 (Customer Lockbox disabled), CIS 2.1.3 (malware admin notification disabled), CIS 2.1.13 (connection filter safe list enabled), CIS 7.2.11 (default sharing link permission Edit), CIS 8.2.1 (external federation open to all domains), CIS 8.5.7 (external participants can control meetings), CIS 8.5.8 (external meeting chat enabled), CIS 8.5.9 (cloud recording on by default) |
| 1.30.0 | `MspDomains` and `KnownPartners` loaded from `config.psd1` — replaces hardcoded domain and partner lists; Technical Contact domain check and GDAP partner checks are guarded and skipped with a warning when the respective config values are absent; action item text uses generic wording instead of MSP-specific names |
| 1.29.0 | Summary styling streamlined; added severity colouring for Technical Contact when the corresponding action item is raised; Action Items now group `CRITICAL` before `WARNING`, sort by module order (Entra, Exchange, SharePoint, Mail Security, Intune), and use fixed column alignment; Conditional Access drilldowns expanded to show richer scope and condition detail when newer `Entra_CA_Policies.csv` exports are present |
| 1.28.0 | All CSV discovery now reads from the shared `Raw Files\` folder while report sections continue grouping by filename prefix; raw-file links in the HTML report now point to `Raw Files\<filename>` |
| 1.27.0 | Intune section enhanced: full managed-device inventory table; clickable compliance policy drilldowns with per-setting detail; clickable app drilldowns with installation summary; configuration profile/policy section now supports the richer Intune exports including modern configuration policies |
| 1.26.0 | All CSV path lookups updated to use per-module subfolders (`Entra\`, `Exchange\`, `SharePoint\`, `MailSecurity\`, `Intune\`); subfolder path variables defined near top of script |
| 1.25.0 | Config profiles table now includes an expandable inline settings panel (reads `Intune_ConfigProfileSettings.csv`; click "N setting(s)" to expand a per-setting name/value table beneath each profile row) |
| 1.24.0 | Added Intune / Endpoint Management HTML section and action item checks: no compliance policies (critical), non-compliant devices (critical), stale devices >30 days (warning), encryption not required by policy (warning), grace period >24 hours (warning), personal enrolment not blocked (warning), apps with install failures (warning); gracefully renders an info note when no Intune-capable licence is detected |
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
| 2.6.0 | After each customer's audit completes, automatically calls `Helpers\Publish-HuduAuditReport.ps1` to push the report to Hudu — reads the output folder path from `$env:TEMP\365Audit_LastOutput.txt` written by `Start-365Audit.ps1`; Hudu publish failures are caught and logged as warnings without stopping remaining customers; customer list migrated from `UnattendedCustomers.json` to `UnattendedCustomers.psd1` |
| 2.5.1 | `OutputRoot` validation: resolves to absolute path, checks drive/UNC qualifier is accessible, and attempts `New-Item` before the customer loop — bad path stops the entire batch before any customer runs |
| 2.5.0 | Added `-OutputRoot` parameter; falls back to `OutputRoot` in `config.psd1`; passed through to each `Start-365Audit.ps1` call so all customers in a bulk run write to the same root |
| 2.4.0 | `-Modules` parameter type changed from `[int[]]` to `[string[]]` with `ValidateSet('1'..'7', 'A')` to match `Start-365Audit.ps1`; per-customer module fallback cast changed from `[int[]]` to `@()` to allow both numeric and `'A'` values from `UnattendedCustomers.psd1`; `UnattendedCustomers.psd1.example` updated to use `@('A')` and document options 6 (Teams) and 7 (ScubaGear) |
| 2.3.0 | Config loaded from `config.psd1` — `HuduApiKey` and `HuduBaseUrl` sourced from file instead of environment variables |
| 2.2.0 | Console output logging delegated to `Start-365Audit.ps1` — each customer's full run log is saved as `AuditLog.txt` in that customer's audit output folder; no separate bulk transcript needed since `Start-365Audit.ps1` stops its transcript in `finally` before the next customer begins |
| 2.1.0 | Customer list extracted to `UnattendedCustomers.json` — techs edit the JSON file rather than the script; each entry has `HuduCompanySlug` and `Modules` (per-customer module selection); `-Modules` param now acts as a global override for all customers; `-Customers` param filters by slug; summary table includes per-customer modules column; script hard-errors with copy hint if JSON file is not found |
| 2.0.0 | Full rewrite: Hudu-based credential management — customer list uses Hudu company IDs/slugs, no credentials stored in the script; per-customer flow: (1) call `Setup-365AuditApp.ps1 -HuduCompanyId` to check/renew cert automatically, (2) call `Start-365Audit.ps1 -HuduCompanyId -Modules` with fresh credentials from Hudu; supports `-Customers` override, `-Modules` selection, `-SkipCertCheck`, and `-HuduApiKey`/`-HuduBaseUrl`; per-customer error isolation (one failure does not stop remaining customers); final summary table with status per customer |
| 1.0.0 | Initial version (hardcoded credentials per customer, deprecated) |

---

## Common/Audit-Common.ps1

| Version | Notes |
|---------|-------|
| 1.24.0 | `Initialize-AuditOutput` accepts optional `-OutputRoot` parameter; persists value in `$script:AuditOutputRoot` so subsequent calls from dot-sourced module scripts automatically use the same root without needing to re-pass the parameter; if `$script:AuditOutputRoot` is set (by `-OutputRoot` or by the launcher setting it directly), per-customer folders are created inside that root instead of the default two-levels-above-toolkit path |
| 1.23.0 | Added `Add-AuditIssue` function — call from catch blocks to write collection failures, permission errors, and module issues to `AuditIssues.csv` in the audit output folder; supports Severity (Critical/Warning/Info), Section, Collector, Description, and optional Action fields; appends to existing CSV or creates with header on first write |
| 1.22.0 | `Connect-TeamsSecure` added — connects to Microsoft Teams PowerShell using certificate-based app-only auth; detects `$AuditAppId`/`$AuditTenantId`/`$AuditCertFilePath`/`$AuditCertPassword` from launcher scope; loads X509Certificate2 directly from PFX without cert-store import; falls back to interactive browser auth when credentials are absent |
| 1.21.0 | Graph sub-module loading refactored to lazy-load via new `Import-GraphSubModules` helper — only installs and imports the modules required by the running audit section instead of all 9 at startup; core bootstrap reduced to `Authentication` and `Identity.DirectoryManagement`; `Resolve-GraphModuleVersion` simplified to use `Authentication` version only (all Graph sub-modules are versioned in lockstep); `Install-Module` for sub-modules uses `-WarningAction SilentlyContinue` to suppress spurious dependency-in-use warnings; all install blocks now verify the module is discoverable post-install and display the installed version |
| 1.20.0 | Delegated `Connect-MgGraphSecure` now passes `-NoWelcome` so the Microsoft Graph SDK banner does not clutter launcher or diagnostic-script output |
| 1.19.0 | `Initialize-AuditOutput` now creates and returns a shared `Raw Files\` subfolder (`RawOutputPath`) under each customer run folder so all modules and the launcher transcript can write to a single raw-output location |
| 1.16.0 | Removed `Microsoft.Graph.DeviceManagement.Enrolment` from the sub-module list — that name (British spelling) is a v1.x-only module; its presence caused PowerShellGet to pull `Microsoft.Graph.Authentication` v1.28.0 as a dependency, which conflicts with the v2.x installation; all required Intune cmdlets are available via `Microsoft.Graph.DeviceManagement` and `Microsoft.Graph.Devices.CorporateManagement` in v2.x; added `-SkipPublisherCheck` to `Install-Module` to prevent catalog signing false-positives on Microsoft's own modules |
| 1.15.0 | Added `Microsoft.Graph.DeviceManagement`, `Microsoft.Graph.DeviceManagement.Enrolment`, and `Microsoft.Graph.Devices.CorporateManagement` to the `$_graphSubModules` install bootstrap — required by `Invoke-IntuneAudit.ps1` |
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

---

## Helpers/Get-HuduAssetLayouts.ps1

| Version | Notes |
|---------|-------|
| 1.0.0 | Initial release — connects to Hudu and lists all asset layouts with their IDs; assists in identifying the correct `HuduAssetLayoutId` for `config.psd1` since the ID is not exposed in the Hudu UI |

---

## Helpers/Get-ModuleVersionStatus.ps1

| Version | Notes |
|---------|-------|
| 1.0.0 | Initial release — performs a single bulk PSGallery lookup for all 365Audit required modules and displays installed vs latest version with status (OK / UPDATE AVAILABLE / NOT INSTALLED / MULTIPLE VERSIONS) |

---

## Helpers/New-HuduAssetLayout.ps1

| Version | Notes |
|---------|-------|
| 1.0.0 | Initial release — creates the M365 Audit Toolkit asset layout in Hudu via the REST API; reads layout name from `HuduAssetName` in `config.psd1`; prints a field summary and requires confirmation before creating; prints the new layout ID for use in `HuduAssetLayoutId`; handles 401, 404, and 422 with actionable error messages; supports `-WhatIf` and `-Force`; requires Hudu Administrator or Super Administrator |

---

## Helpers/Remove-AuditCustomer.ps1

| Version | Notes |
|---------|-------|
| 1.0.0 | Initial release — offboards a customer by removing their app registration from Entra ID and deleting the corresponding Hudu asset; Hudu lookup resolves AppId/TenantId automatically from the asset; `-PermanentDelete` purges from the Entra recycle bin immediately (default is soft-delete, recoverable for 30 days); `-AppId`/`-TenantId` can be used directly when Hudu is not involved; full `SupportsShouldProcess` with `ConfirmImpact = 'High'` and `-WhatIf` support |

---

## Helpers/Publish-HuduAuditReport.ps1

| Version | Notes |
| ------- | ----- |
| 1.2.1 | `HuduBaseUrl`, `HuduApiKey`, and `ReportLayoutId` are now optional — values fall back to `HuduBaseUrl`, `HuduApiKey`, and `HuduReportLayoutId` in `config.psd1`; zip attachment deleted after upload (previously left in the report root folder) |
| 1.2.0 | Month-over-month delta: loads `AuditMetrics.json` and `ActionItems.json` from the current output folder; finds the prior month's asset (`M365 Audit - yyyy-MM` for previous month) in the same asset query; downloads the prior zip attachment via the Hudu uploads API; extracts `AuditMetrics.json` and `ActionItems.json` using `System.IO.Compression.ZipFile`; replaces `<!-- TILE_DELTA_* -->` markers with coloured `+/-N%` spans (green = improvement, red = regression, neutral grey for device count); builds "Changes Since Last Month" `<details>` section with a licence/storage metric change row, a Resolved table (green header), and a New table (red header); injects the section at `<!-- AUDIT_DELTA_INJECT -->`; all prior-data logic is non-fatal (wrapped in `try/catch`) — absent or failed prior zip results in markers being cleared and the report published without delta; prior month asset lookup is performed in the same API call as the current month asset lookup (no extra request) |
| 1.1.0 | Reads `ActionItems.json` from `$OutputPath` to confirm audit completed before publishing |
| 1.0.0 | Initial release — resolves Hudu company by slug; finds or creates `M365 Audit - yyyy-MM` asset under the `Monthly Audit Report` layout; populates `report_summary` with `M365_HuduReport.html`; uploads `M365_AuditSummary.html` and a zip of the full output folder as attachments |

---

## Helpers/Sync-UnattendedCustomers.ps1

| Version | Notes |
|---------|-------|
| 1.0.1 | `$DefaultModules` type changed from `[int[]]` to `[string[]]` with `ValidateSet`; default changed from `@(9)` to `@('A')`; module comment and `.PARAMETER` doc updated to include modules 6=Teams, 7=ScubaGear, A=All; PSD1 serialisation now quotes module values so `'A'` renders as a valid string in the output file |
| 1.0.0 | Initial release — queries Hudu for all assets matching `HuduAssetLayoutId`, resolves company slugs, and merges results into `UnattendedCustomers.psd1`; preserves existing entries, appends new companies with `Modules = @(9)`, warns about slugs no longer in Hudu; supports `-WhatIf` and `-DefaultModules` |

---

## Helpers/Uninstall-AuditModules.ps1

| Version | Notes |
|---------|-------|
| 1.0.0 | Initial release — uninstalls all versions of all 365Audit required modules; pre-flight check blocks execution if any modules are currently loaded in the session; supports `-WhatIf` |
