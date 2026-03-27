# Invoke-ScubaGearAudit.ps1

Runs the [CISA ScubaGear M365 Foundations Benchmark](https://github.com/cisagov/ScubaGear) assessment against the tenant.

ScubaGear and its dependencies must run in **Windows PowerShell 5.1** — this module spawns a clean PS 5.1 subprocess automatically and cleans up the temporary certificate import on exit.

> Requires `powershell.exe` (Windows PowerShell 5.1) to be available on the machine. Power Platform is excluded (requires an interactive one-time registration).

## Required Permissions

ScubaGear uses the same app registration as the rest of the toolkit. `Setup-365AuditApp.ps1` grants all required permissions.

## Output

Output is written to `Raw\ScubaGear_<timestamp>\` inside the customer's audit folder.

| File | Description |
|---|---|
| `BaselineReports.html` | ScubaGear's own interactive HTML report with full control details |
| `ScubaResults_<uuid>.json` | Consolidated JSON results ingested by `Generate-AuditSummary.ps1` |
| `ScubaResults.csv` | Flat CSV of all controls with pass/fail/warning status |
| `ActionPlan.csv` | Failing Shall controls with blank remediation fields for MSP follow-up |
| `IndividualReports\` | Per-product HTML and JSON reports (AAD, Defender, EXO, SharePoint, Teams) |

## Summary Report Integration

`Generate-AuditSummary.ps1` detects the `ScubaGear_*` folder automatically and:
- Adds failing controls to the action items list
- Renders a collapsible CIS Baseline section with per-product pass/fail/warning counts and a link to the full ScubaGear HTML report
