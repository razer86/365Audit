# Invoke-SharePointAudit.ps1

Connects to SharePoint Online using PnP.PowerShell with certificate-based app-only authentication and audits sites, permissions, external sharing, and OneDrive usage.

> Requires **PowerShell 7.4+** and the **PnP.PowerShell v3+** module.

## Required Permissions

Granted automatically by `Setup-365AuditApp.ps1`:

| Permission | Type | Purpose |
|---|---|---|
| `Sites.FullControl.All` | Application (SharePoint) | Read all site collections, permissions, and OneDrive |

## Output Files

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

## Notes

PnP.PowerShell uses MSAL token caching — the browser prompt only appears once per session, not once per site.
