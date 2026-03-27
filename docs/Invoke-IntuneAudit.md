# Invoke-IntuneAudit.ps1

Connects to Microsoft Graph and audits Intune / Endpoint Management. Skipped gracefully with an informational note if no Intune-capable licence is detected in the tenant.

## Required Permissions

Granted automatically by `Setup-365AuditApp.ps1`:

| Permission | Type | Purpose |
|---|---|---|
| `DeviceManagementManagedDevices.Read.All` | Application | Device inventory and compliance states |
| `DeviceManagementConfiguration.Read.All` | Application | Compliance and configuration policies |
| `DeviceManagementApps.Read.All` | Application | App inventory and install status |

## Output Files

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
