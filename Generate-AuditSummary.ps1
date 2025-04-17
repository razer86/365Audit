param (
    [string]$AuditFolder,
    [switch]$DevMode = $false
)

if (-not $DevMode -and $MyInvocation.InvocationName -eq $MyInvocation.MyCommand.Name) {
    Write-Error "This script must be run from the 365Audit launcher. Use -DevMode for development." -ErrorAction Stop
}

﻿<#
.SYNOPSIS
    Generates an HTML summary report from Microsoft 365 audit CSV output files.

.DESCRIPTION
    Compiles key findings from Entra, Exchange, and SharePoint audit modules
    into a styled HTML summary with collapsible sections and links to full CSV output files.

    CSVs Parsed:
    - Entra_Users.csv             : MFA, license, password, and user identity overview
    - Entra_GlobalAdmins.csv      : List of Global Administrators
    - Entra_AdminRoles.csv        : Users assigned to Entra roles
    - Entra_RiskyUsers.csv        : Users flagged as risky sign-ins
    - Exchange_Mailboxes.csv      : User and shared mailbox summary
    - Exchange_InboxForwardingRules.csv : Inbox rules that forward externally
    - SharePoint_Sites.csv        : SharePoint sites, size, type
    - SharePoint_ExternalSharing_SiteOverrides.csv : Per-site sharing overrides
    - OneDriveUsage.csv           : OneDrive usage per user
    - OneDrive_Unlicensed.csv     : OneDrive accounts missing licenses

.NOTES
    Author      : Raymond Slater
    Version     : 1.1.0
    Change Log  :
        1.0.0 - Initial release
        1.0.1 - Updated Entra audit sources
        1.1.0 - Added Entra_RiskyUsers and Entra_AdminRoles summaries

.LINK
    https://github.com/razer86/365Audit
#>

# =========================================
# ===   Parameters and Input Validation   ===
# =========================================
param (
    [Parameter(Mandatory = $true)]
    [string]$AuditFolder
)

if (-not (Test-Path $AuditFolder)) {
    Write-Error "❌ Provided audit folder does not exist: $AuditFolder"
    exit 1
}

# =========================================
# ===   HTML Section Rendering Utility   ===
# =========================================
function Add-Section {
    param (
        [string]$Title,
        [string[]]$CsvFiles,
        [string]$SummaryHtml
    )

    $section = "<details class='section'>"
    $section += "<summary>$Title</summary><div class='content'>"
    $section += $SummaryHtml

    if ($CsvFiles.Count -gt 0) {
        $section += "<h4>CSV Files:</h4><ul>"
        foreach ($file in $CsvFiles) {
            $name = [System.IO.Path]::GetFileName($file)
            $section += "<li><a href='$name' target='_blank'>$name</a></li>"
        }
        $section += "</ul>"
    }

    $section += "</div></details>"
    return $section
}

# =========================================
# ===   HTML Base Layout Initialization   ===
# =========================================
$reportPath = Join-Path $AuditFolder "M365_AuditSummary.html"
$html = @()

$html += @"
<!DOCTYPE html>
<html lang='en'>
<head>
<meta charset='UTF-8'>
<title>Microsoft 365 Audit Summary</title>
<style>
body { font-family: Segoe UI, sans-serif; background: #f7f7f7; color: #333; margin: 2rem; }
h1 { text-align: center; }
.section { margin-bottom: 1.5rem; border: 1px solid #ccc; border-radius: 6px; background: #fff; box-shadow: 0 2px 5px rgba(0,0,0,0.05); }
summary { font-size: 1.2rem; font-weight: bold; padding: 1rem; cursor: pointer; background: #eaeaea; border-bottom: 1px solid #ccc; }
.content { padding: 1rem; }
.status-ok { color: green; font-weight: bold; }
.status-warning { color: darkorange; font-weight: bold; }
.status-critical { color: red; font-weight: bold; }
</style>
</head>
<body>
<h1>Microsoft 365 Audit Summary</h1>
"@

# === Entra Summary ===
$entraFiles = Get-ChildItem "$OutputPath\Entra_*.csv" -ErrorAction SilentlyContinue
$entraSummary = @()

if ($entraFiles.Count -gt 0) {
    $entraUsersCsv = Join-Path $OutputPath "Entra_Users.csv"
    $globalAdminsCsv = Join-Path $OutputPath "Entra_GlobalAdmins.csv"
    $adminRolesCsv = Join-Path $OutputPath "Entra_AdminRoles.csv"
    $groupsCsv = Join-Path $OutputPath "Entra_Groups.csv"
    $entraGuestUsersCsv = Join-Path $OutputPath "Entra_GuestUsers.csv"
    $entraLicencesCsv = Join-Path $OutputPath "Entra_Licenses.csv"
    $ssprCsv = Join-Path $OutputPath "Entra_SSPR.csv"

    $ssprStatus = "Unknown"
    if (Test-Path $ssprCsv) {
        $ssprData = Import-Csv $ssprCsv
        if ($ssprData -and $ssprData.RegistrationEnforced) {
            $ssprEnabled = $ssprData[0].RegistrationEnforced -eq "True"
            $ssprStatus = if ($ssprEnabled) {
                "<span class='status-ok'>✅  Self-Service Password Reset is <b>Enabled</b></span>"
            } else {
                "<span class='status-critical'>❌ Self-Service Password Reset is <b>Disabled</b></span>"
            }
        }
    }

    if (Test-Path $entraUsersCsv) {
        $userSummary = Import-Csv $entraUsersCsv
        $mfaTotal = $userSummary.Count
        $mfaEnabled = ($userSummary | Where-Object { $_.MFAEnabled -eq 'True' }).Count
        $mfaPercent = if ($mfaTotal -gt 0) { [math]::Round(($mfaEnabled / $mfaTotal) * 100, 1) } else { 0 }

        $tooltip = "Microsoft recommends using a passwordless method such as FIDO2 for emergency access accounts. See: https://learn.microsoft.com/en-us/entra/identity/role-based-access-control/security-emergency-access"
        $mfaTooltip = "<span title='$tooltip'>✅  MFA Enabled for <b>$mfaPercent%</b> of users</span>"
        $mfaClass = if ($mfaPercent -eq 100) {
            "status-ok"
        } elseif ($mfaPercent -gt 0) {
            "status-warning"
        } else {
            "status-critical"
        }
        $entraSummary += "<p class='$mfaClass'>$mfaTooltip</p>"
        
        # Add SSPR status line
        if ($ssprStatus -ne "Unknown") {
            $entraSummary += "<p>$ssprStatus</p>"
        }

        $entraSummary += "<table border='1' cellpadding='6' cellspacing='0'><thead><tr>"
        $entraSummary += "<th>User Principal Name</th><th>First Name</th><th>Last Name</th><th>License</th><th>MFA Enabled</th><th>MFA Types</th><th>Password Never Expires</th><th>Last Password Change</th><th>Last Sign In</th>"
        $entraSummary += "</tr></thead><tbody>"

        foreach ($user in $userSummary) {
            $entraSummary += "<tr>"
            $entraSummary += "<td>$($user.UPN)</td><td>$($user.FirstName)</td><td>$($user.LastName)</td><td>$($user.AssignedLicense)</td>"
            if ($user.MFAEnabled -eq "False") {
                $entraSummary += "<td style='background-color:#ffcccc;'>$($user.MFAEnabled)</td>"
            } else {
                $entraSummary += "<td>$($user.MFAEnabled)</td>"
            }
            $entraSummary += "<td>$($user.MFAMethods)</td><td>$($user.DisablePasswordExpiration)</td><td>$($user.LastPasswordChange)</td><td>$($user.LastSignIn)</td>"
            $entraSummary += "</tr>"
        }
        $entraSummary += "</tbody></table>"
    }

    # === License Summary Table ===
    if (Test-Path $entraLicencesCsv) {
        $licenses = Import-Csv $entraLicencesCsv
        if ($licenses.Count -gt 0) {
            $licenseHtml = @()
            $licenseHtml += "<h4>🔑 License Summary</h4>"
            $licenseHtml += "<table border='1' cellpadding='6' cellspacing='0' style='border-collapse: collapse;'>"
            $licenseHtml += "<tr><th>License Name</th><th>Total</th><th>Suspended</th><th>Warning</th><th>Assigned</th><th>Purchase Channel</th></tr>"

            foreach ($lic in $licenses) {
                $licenseHtml += "<tr>
                    <td>$($lic.SkuFriendlyName)</td>
                    <td>$($lic.EnabledUnits)</td>
                    <td>$($lic.SuspendedUnits)</td>
                    <td>$($lic.WarningUnits)</td>
                    <td>$($lic.ConsumedUnits)</td>
                    <td>$($lic.PurchaseChannel)</td>
                </tr>"
            }

            $licenseHtml += "</table><br/>"
            $entraSummary += $licenseHtml
        }
    }


    if (Test-Path $globalAdminsCsv) {
        $globalAdmins = Import-Csv $globalAdminsCsv
        $gaCount = $globalAdmins.Count
        if ($gaCount -eq 1) {
            $entraSummary += "<p>⚠ <span class='status-warning'>Only 1 Global Administrator found</span></p>"
        } elseif ($gaCount -eq 0) {
            $entraSummary += "<p>❗ <span class='status-critical'>No Global Administrators found</span></p>"
        } else {
            $entraSummary += "<p>✅  $gaCount Global Administrators configured</p>"
        }
    }

    if (Test-Path $adminRolesCsv) {
        $adminRoles = Import-Csv $adminRolesCsv
        $roleCount = ($adminRoles | Select-Object -ExpandProperty RoleName -Unique).Count
        $userCount = ($adminRoles | Select-Object -ExpandProperty MemberUserPrincipalName -Unique).Count
        $entraSummary += "<p>👮 $userCount users assigned to $roleCount Entra roles</p>"
    }


    $html += Add-Section -Title "Microsoft Entra" -CsvFiles $entraFiles.FullName -SummaryHtml ($entraSummary -join "`n")
}

# Build Exchange Online section
$exchangeFiles = Get-ChildItem "$($latestFolder.FullName)\Exchange_*.csv"
if ($exchangeFiles.Count -gt 0) {
    $forwarding = Import-Csv "$($latestFolder.FullName)\Exchange_InboxForwardingRules.csv" -ErrorAction SilentlyContinue
    $mbx = Import-Csv "$($latestFolder.FullName)\Exchange_Mailboxes.csv" -ErrorAction SilentlyContinue

    $exchangeSummary = @()
    $exchangeSummary += "<p>📬 $($mbx.Count) mailboxes audited</p>"
    if ($forwarding.Count -gt 0) {
        $exchangeSummary += "<p>⚠ <span class='status-warning'>$($forwarding.Count) inbox rules forward externally</span></p>"
    } else {
        $exchangeSummary += "<p>✅  No external forwarding inbox rules detected</p>"
    }
    $html += Add-Section -Title "Exchange Online" -CsvFiles $exchangeFiles.FullName -SummaryHtml ($exchangeSummary -join "`)n")
}

# Build SharePoint / OneDrive section
$spFiles = Get-ChildItem "$($latestFolder.FullName)\SharePoint_*.csv" | Sort-Object Name
if ($spFiles.Count -gt 0) {
    $sites = Import-Csv "$($latestFolder.FullName)\SharePoint_Sites.csv" -ErrorAction SilentlyContinue
    $sharingOverrides = Import-Csv "$($latestFolder.FullName)\SharePoint_ExternalSharing_SiteOverrides.csv" -ErrorAction SilentlyContinue
    $unlicensedOD = Import-Csv "$($latestFolder.FullName)\OneDrive_Unlicensed.csv" -ErrorAction SilentlyContinue

    $spSummary = @()
    $spSummary += "<p>🌐 $($sites.Count) SharePoint sites audited</p>"
    if ($sharingOverrides.Count -gt 0) {
        $spSummary += "<p>❗ <span class='status-critical'>$($sharingOverrides.Count) sites override external sharing settings</span></p>"
    } else {
        $spSummary += "<p>✅  No site-level external sharing overrides</p>"
    }
    if ($unlicensedOD.Count -gt 0) {
        $spSummary += "<p>⚠ <span class='status-warning'>$($unlicensedOD.Count) unlicensed OneDrive users detected</span></p>"
    } else {
        $spSummary += "<p>✅  All OneDrive users are licensed</p>"
    }
    $html += Add-Section -Title "SharePoint / OneDrive" -CsvFiles ($spFiles.FullName + "$($latestFolder.FullName)\OneDriveUsage.csv","$($latestFolder.FullName)\OneDrive_Unlicensed.csv") -SummaryHtml ($spSummary -join "`)n")
}
# Finalize HTML and write to disk
$html += "</body></html>"
$html -join "`n" | Set-Content -Path $reportPath -Encoding UTF8

Write-Host "✅ Summary written to $reportPath"
Start-Process $reportPath
