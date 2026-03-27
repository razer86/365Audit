# Invoke-EntraAudit.ps1

Connects to Microsoft Graph and audits Entra ID (Azure Active Directory).

Covers identity health, licence assignments, MFA status, admin roles, Conditional Access, guest accounts, groups, Identity Secure Score, and risky users/sign-ins. Azure AD P2 features (Risky Users, Risky Sign-Ins, PIM) are skipped gracefully when the tenant is not licensed for them.

## Required Permissions

Granted automatically by `Setup-365AuditApp.ps1`:

| Permission | Type | Purpose |
|---|---|---|
| `User.Read.All` | Application | User inventory, MFA status, last sign-in |
| `Directory.Read.All` | Application | Groups, roles, CA policies, trusted locations |
| `AuditLog.Read.All` | Application | Sign-in logs, account creation/deletion events |
| `Policy.Read.All` | Application | Conditional Access policies, security defaults |
| `SecurityEvents.Read.All` | Application | Identity Secure Score |
| `IdentityRiskyUser.Read.All` | Application | Risky users (P2) |
| `IdentityRiskEvent.Read.All` | Application | Risky sign-ins (P2) |
| `PrivilegedAccess.Read.AzureAD` | Application | PIM assignments (P2) |

## Output Files

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
