<#
.SYNOPSIS
    Gathers Microsoft 365 mail security policies (anti-spam, anti-phishing, spoofing, transport rules) and exports them to JSON.

.AUTHOR
    Raymond Slater
.VERSION
    1.1 - 2025-04-17
.LINK
    https://github.com/razer86/365Audit
#>

# === Connect to Exchange Online if not already connected ===
if (-not (Get-ConnectionInformation -ErrorAction SilentlyContinue)) {
    Connect-ExchangeOnline -ShowBanner:$false
}

# === Output folder ===
$OutputPath = "$PSScriptRoot\PolicyExport_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null

function Export-Json {
    param (
        [Parameter(Mandatory)]
        $Data,

        [Parameter(Mandatory)]
        [string]$Path
    )
    $Data | ConvertTo-Json -Depth 10 | Out-File -FilePath $Path -Encoding UTF8
}

# === Anti-Spam Policies ===
Export-Json -Data (Get-HostedContentFilterPolicy) -Path "$OutputPath\AntiSpamPolicies.json"
Export-Json -Data (Get-HostedContentFilterRule) -Path "$OutputPath\AntiSpamRules.json"

# === Anti-Phishing Policies ===
Export-Json -Data (Get-AntiPhishPolicy) -Path "$OutputPath\AntiPhishPolicies.json"
Export-Json -Data (Get-AntiPhishRule) -Path "$OutputPath\AntiPhishRules.json"

# === Spoof Intelligence (Defender permissions required) ===
try {
    $spoofedDomains = Get-SpoofIntelligenceInsight
    Export-Json -Data $spoofedDomains -Path "$OutputPath\SpoofIntelligence.json"
} catch {
    Write-Warning "⚠️ Spoof Intelligence not available or permission denied."
}

# === Inbound/Outbound Connectors ===
Export-Json -Data (Get-InboundConnector) -Path "$OutputPath\InboundConnectors.json"
Export-Json -Data (Get-OutboundConnector) -Path "$OutputPath\OutboundConnectors.json"

# === Transport Rules (Mail Flow) ===
Export-Json -Data (Get-TransportRule | Select Name,Priority,State,Mode,Comments) -Path "$OutputPath\TransportRules.json"

# === DKIM Status ===
Export-Json -Data (Get-DkimSigningConfig | Select Domain,Enabled,Selector1CNAME,Selector2CNAME) -Path "$OutputPath\DkimStatus.json"

# === DMARC Record Check ===
$domains = (Get-AcceptedDomain).DomainName
$dmarcResults = foreach ($domain in $domains) {
    try {
        $txt = (Resolve-DnsName "_dmarc.$domain" -Type TXT -ErrorAction Stop).Strings -join ''
        [PSCustomObject]@{ Domain = $domain; DMARC = $txt }
    } catch {
        [PSCustomObject]@{ Domain = $domain; DMARC = "Not Found" }
    }
}
Export-Json -Data $dmarcResults -Path "$OutputPath\DMARCRecords.json"

# === SPF Record Check ===
$spfResults = foreach ($domain in $domains) {
    try {
        $txtRecords = (Resolve-DnsName $domain -Type TXT -ErrorAction Stop).Strings
        $spf = $txtRecords | Where-Object { $_ -like "v=spf1*" } | Select-Object -First 1
        [PSCustomObject]@{ Domain = $domain; SPF = $spf }
    } catch {
        [PSCustomObject]@{ Domain = $domain; SPF = "DNS query failed" }
    }
}
Export-Json -Data $spfResults -Path "$OutputPath\SPFRecords.json"

Write-Host "`n✅ Mail security policies exported to:`n$OutputPath"
