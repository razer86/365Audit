<#
.SYNOPSIS
    Audits Microsoft 365 mail security policies.

.DESCRIPTION
    Connects to Exchange Online and exports mail security configuration including:
    - DKIM signing status per domain        -> MailSec_DKIM.csv
    - DMARC records per accepted domain     -> MailSec_DMARC.csv
    - SPF records per accepted domain       -> MailSec_SPF.csv
    - Anti-spam policies (full detail)      -> MailSec_AntiSpam.json
    - Anti-spam rules                       -> MailSec_AntiSpamRules.json
    - Anti-phishing policies (full detail)  -> MailSec_AntiPhish.json
    - Anti-phishing rules                   -> MailSec_AntiPhishRules.json
    - Spoof intelligence insights           -> MailSec_SpoofIntelligence.json
    - Inbound connectors                    -> MailSec_InboundConnectors.json
    - Outbound connectors                   -> MailSec_OutboundConnectors.json
    - Transport (mail flow) rules summary   -> MailSec_TransportRules.json

    CSV files are consumed by the HTML summary report.
    JSON files are provided as supplementary deep-detail exports.

.NOTES
    Author      : Raymond Slater
    Version     : 1.6.0
    Change Log  :
        1.0.0 - Initial release (standalone, JSON-only output)
        1.1.0 - Minor updates
        1.2.0 - Integrated with launcher and shared output folder;
                flat data (DKIM, DMARC, SPF) now exported as CSV;
                JSON retained for nested policy objects;
                added guard clause and CmdletBinding
        1.2.1 - Export-Json now skips gracefully on null/empty results
        1.3.0 - EXO connection check now filters by State eq Connected;
                Import-Module uses -ErrorAction Stop
        1.4.0 - Exchange Online now uses app-only auth (via Connect-ExchangeOnlineSecure)
                when -AppId/-AppSecret/-TenantId are provided at launch
        1.5.0 - Replaced per-section Write-Host progress lines with Write-Progress
                for cleaner terminal output
        1.6.0 - Added Step X/Y counter to Write-Progress status strings

.LINK
    https://github.com/razer86/365Audit
#>

#Requires -Version 7.2

param (
    [string]$AuditFolder,
    [switch]$DevMode = $false
)

if (-not $DevMode -and $MyInvocation.InvocationName -eq $MyInvocation.MyCommand.Name) {
    Write-Error "This script must be run from the 365Audit launcher. Use -DevMode for development." -ErrorAction Stop
}

$ScriptVersion = "1.6.0"
Write-Verbose "Invoke-MailSecurityAudit.ps1 loaded (v$ScriptVersion)"

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# === Retrieve shared output folder ===
try {
    $context   = Initialize-AuditOutput
    $outputDir = $context.OutputPath
}
catch {
    Write-Error "Failed to initialise audit output directory: $_"
    exit 1
}

# === Ensure ExchangeOnlineManagement module is available ===
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Write-Host "Installing ExchangeOnlineManagement module..." -ForegroundColor Yellow
    Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force
}
Import-Module ExchangeOnlineManagement -ErrorAction Stop

# === Connect to Exchange Online ===
Connect-ExchangeOnlineSecure

Write-Host "`nStarting Mail Security Audit for $($context.OrgName)..." -ForegroundColor Cyan

# Helper: write an object to a JSON file; skips gracefully if data is null or empty
function Export-Json {
    [CmdletBinding()]
    param (
        $Data,
        [Parameter(Mandatory)] [string]$Path
    )
    if ($null -eq $Data -or ($Data -is [System.Array] -and $Data.Count -eq 0)) {
        Write-Verbose "No data returned for $(Split-Path $Path -Leaf) — skipping."
        return
    }
    $Data | ConvertTo-Json -Depth 10 | Set-Content -Path $Path -Encoding UTF8
}

$acceptedDomains = Get-AcceptedDomain

$step       = 0
$totalSteps = 4
$activity   = "Mail Security Audit — $($context.OrgName)"


# ================================
# ===   DKIM Status  ->  CSV    ===
# ================================
$step++
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Checking DKIM signing configuration..." -PercentComplete ([int]($step / $totalSteps * 100))

$dkimStatus = foreach ($domain in $acceptedDomains) {
    try {
        $cfg = Get-DkimSigningConfig -Identity $domain.DomainName -ErrorAction Stop
        [PSCustomObject]@{
            Domain         = $domain.DomainName
            DKIMEnabled    = $cfg.Enabled
            Selector1CNAME = $cfg.Selector1CNAME
            Selector2CNAME = $cfg.Selector2CNAME
        }
    }
    catch {
        [PSCustomObject]@{
            Domain         = $domain.DomainName
            DKIMEnabled    = "Not Configured"
            Selector1CNAME = "N/A"
            Selector2CNAME = "N/A"
        }
    }
}
$dkimStatus | Export-Csv "$outputDir\MailSec_DKIM.csv" -NoTypeInformation -Encoding UTF8
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Checking DKIM signing configuration..." -CurrentOperation "Saved: MailSec_DKIM.csv" -PercentComplete ([int]($step / $totalSteps * 100))


# ================================
# ===   DMARC Records  ->  CSV  ===
# ================================
$step++
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Checking DMARC records..." -PercentComplete ([int]($step / $totalSteps * 100))

$dmarcResults = foreach ($domain in $acceptedDomains) {
    try {
        $txt = (Resolve-DnsName "_dmarc.$($domain.DomainName)" -Type TXT -ErrorAction Stop).Strings -join ''
        [PSCustomObject]@{ Domain = $domain.DomainName; DMARC = $txt }
    }
    catch {
        [PSCustomObject]@{ Domain = $domain.DomainName; DMARC = "Not Found" }
    }
}
$dmarcResults | Export-Csv "$outputDir\MailSec_DMARC.csv" -NoTypeInformation -Encoding UTF8
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Checking DMARC records..." -CurrentOperation "Saved: MailSec_DMARC.csv" -PercentComplete ([int]($step / $totalSteps * 100))


# ================================
# ===   SPF Records  ->  CSV    ===
# ================================
$step++
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Checking SPF records..." -PercentComplete ([int]($step / $totalSteps * 100))

$spfResults = foreach ($domain in $acceptedDomains) {
    try {
        $txtRecords = (Resolve-DnsName $domain.DomainName -Type TXT -ErrorAction Stop).Strings
        $spf        = $txtRecords | Where-Object { $_ -like "v=spf1*" } | Select-Object -First 1
        [PSCustomObject]@{ Domain = $domain.DomainName; SPF = $spf }
    }
    catch {
        [PSCustomObject]@{ Domain = $domain.DomainName; SPF = "DNS query failed" }
    }
}
$spfResults | Export-Csv "$outputDir\MailSec_SPF.csv" -NoTypeInformation -Encoding UTF8
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Checking SPF records..." -CurrentOperation "Saved: MailSec_SPF.csv" -PercentComplete ([int]($step / $totalSteps * 100))


# ================================
# ===   Policy Exports  ->  JSON
# ================================
$step++
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Exporting mail security policies (JSON)..." -PercentComplete ([int]($step / $totalSteps * 100))

Export-Json -Data (Get-HostedContentFilterPolicy) -Path "$outputDir\MailSec_AntiSpam.json"
Export-Json -Data (Get-HostedContentFilterRule)   -Path "$outputDir\MailSec_AntiSpamRules.json"
Export-Json -Data (Get-AntiPhishPolicy)            -Path "$outputDir\MailSec_AntiPhish.json"
Export-Json -Data (Get-AntiPhishRule)              -Path "$outputDir\MailSec_AntiPhishRules.json"
Export-Json -Data (Get-InboundConnector)           -Path "$outputDir\MailSec_InboundConnectors.json"
Export-Json -Data (Get-OutboundConnector)          -Path "$outputDir\MailSec_OutboundConnectors.json"

Export-Json -Data (Get-TransportRule |
    Select-Object -Property Name, Priority, State, Mode, Comments) `
    -Path "$outputDir\MailSec_TransportRules.json"

try {
    Export-Json -Data (Get-SpoofIntelligenceInsight) -Path "$outputDir\MailSec_SpoofIntelligence.json"
}
catch {
    Write-Warning "Spoof Intelligence not available or permission denied: $_"
}

Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Exporting mail security policies..." -CurrentOperation "Saved: MailSec_AntiSpam.json, MailSec_AntiPhish.json, MailSec_Connectors.json, MailSec_TransportRules.json" -PercentComplete ([int]($step / $totalSteps * 100))


# ================================
# ===   Done                    ===
# ================================
Write-Progress -Id 1 -Activity $activity -Completed
Write-Host "`nMail Security Audit complete. Results saved to: $outputDir`n" -ForegroundColor Green
