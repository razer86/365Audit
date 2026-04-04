function Get-MailSecurityInventory {
    <#
    .SYNOPSIS
        Collects mail security configuration from Exchange Online and exports to CSV/JSON files.

    .DESCRIPTION
        Queries Exchange Online for mail security data including:
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

        Expects Exchange Online to already be connected by the orchestrator.

    .OUTPUTS
        Hashtable with summary counts for the orchestrator.
    #>
    [CmdletBinding()]
    param()

    $ctx       = Get-AuditContext
    $outputDir = $ctx.RawOutputPath

    Write-Host "`nStarting Mail Security Audit for $($ctx.OrgName)..." -ForegroundColor Cyan

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

    # Helper: cross-platform DNS TXT record resolver.
    # On Windows uses Resolve-DnsName. On Linux/macOS tries dig first, then nslookup.
    # Returns an array of TXT string values, or $null on failure.
    function Resolve-TxtRecord {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory)] [string]$Domain,
            [string]$RecordPrefix   # e.g. '_dmarc' — leave empty for the root domain
        )

        $queryDomain = if ($RecordPrefix) { "$RecordPrefix.$Domain" } else { $Domain }

        try {
            if ($IsLinux -or $IsMacOS) {
                if (Get-Command dig -ErrorAction SilentlyContinue) {
                    $digOutput = & dig +short $queryDomain TXT 2>$null
                    if ($LASTEXITCODE -ne 0) { throw "dig failed for $queryDomain" }
                    return @($digOutput -split "`n" |
                             Where-Object { $_ -ne '' } |
                             ForEach-Object { $_.Trim('"') })
                }
                elseif (Get-Command nslookup -ErrorAction SilentlyContinue) {
                    $nslookupOutput = & nslookup -type=TXT $queryDomain 2>$null
                    return @($nslookupOutput |
                             Where-Object { $_ -match '"' } |
                             ForEach-Object { [regex]::Matches($_, '"([^"]+)"') |
                                 ForEach-Object { $_.Groups[1].Value } })
                }
                else {
                    Write-Warning "No DNS lookup tool found for $queryDomain — install bind-utils (dig) or ensure nslookup is available."
                    return $null
                }
            }
            else {
                return (Resolve-DnsName $queryDomain -Type TXT -ErrorAction Stop).Strings
            }
        }
        catch {
            Write-Warning "DNS TXT query failed for ${queryDomain}: $_"
            return $null
        }
    }

    $acceptedDomains = Get-AcceptedDomain

    $step       = 0
    $totalSteps = 4
    $activity   = "Mail Security Audit — $($ctx.OrgName)"


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
    $dkimStatus | Export-Csv (Join-Path $outputDir 'MailSec_DKIM.csv') -NoTypeInformation -Encoding UTF8
    Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Checking DKIM signing configuration..." -CurrentOperation "Saved: MailSec_DKIM.csv" -PercentComplete ([int]($step / $totalSteps * 100))


    # ================================
    # ===   DMARC Records  ->  CSV  ===
    # ================================
    $step++
    Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Checking DMARC records..." -PercentComplete ([int]($step / $totalSteps * 100))

    $dmarcResults = foreach ($domain in $acceptedDomains) {
        # Join fragments — RFC 7489 permits the policy value to span multiple TXT strings
        $txt = (Resolve-TxtRecord -Domain $domain.DomainName -RecordPrefix '_dmarc') -join ''
        [PSCustomObject]@{ Domain = $domain.DomainName; DMARC = if ($txt) { $txt } else { 'Not Found' } }
    }
    $dmarcResults | Export-Csv (Join-Path $outputDir 'MailSec_DMARC.csv') -NoTypeInformation -Encoding UTF8
    Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Checking DMARC records..." -CurrentOperation "Saved: MailSec_DMARC.csv" -PercentComplete ([int]($step / $totalSteps * 100))


    # ================================
    # ===   SPF Records  ->  CSV    ===
    # ================================
    $step++
    Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Checking SPF records..." -PercentComplete ([int]($step / $totalSteps * 100))

    $spfResults = foreach ($domain in $acceptedDomains) {
        $spf = Resolve-TxtRecord -Domain $domain.DomainName |
               Where-Object { $_ -like 'v=spf1*' } |
               Select-Object -First 1
        [PSCustomObject]@{ Domain = $domain.DomainName; SPF = if ($spf) { $spf } else { 'Not Found' } }
    }
    $spfResults | Export-Csv (Join-Path $outputDir 'MailSec_SPF.csv') -NoTypeInformation -Encoding UTF8
    Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Checking SPF records..." -CurrentOperation "Saved: MailSec_SPF.csv" -PercentComplete ([int]($step / $totalSteps * 100))


    # ================================
    # ===   Policy Exports  ->  JSON
    # ================================
    $step++
    Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Exporting mail security policies (JSON)..." -PercentComplete ([int]($step / $totalSteps * 100))

    Export-Json -Data (Get-HostedContentFilterPolicy) -Path (Join-Path $outputDir 'MailSec_AntiSpam.json')
    Export-Json -Data (Get-HostedContentFilterRule)   -Path (Join-Path $outputDir 'MailSec_AntiSpamRules.json')
    Export-Json -Data (Get-AntiPhishPolicy)            -Path (Join-Path $outputDir 'MailSec_AntiPhish.json')
    Export-Json -Data (Get-AntiPhishRule)              -Path (Join-Path $outputDir 'MailSec_AntiPhishRules.json')
    Export-Json -Data (Get-InboundConnector)           -Path (Join-Path $outputDir 'MailSec_InboundConnectors.json')
    Export-Json -Data (Get-OutboundConnector)          -Path (Join-Path $outputDir 'MailSec_OutboundConnectors.json')

    Export-Json -Data (Get-TransportRule |
        Select-Object -Property Name, Priority, State, Mode, Comments) `
        -Path (Join-Path $outputDir 'MailSec_TransportRules.json')

    try {
        Export-Json -Data (Get-SpoofIntelligenceInsight) -Path (Join-Path $outputDir 'MailSec_SpoofIntelligence.json')
    }
    catch {
        Add-AuditIssue -Severity 'Warning' -Section 'Mail Security' -Collector 'Get-SpoofIntelligenceInsight' -Description ($_.Exception.Message ?? "$_") -Action 'Check permissions or re-run Setup-365AuditApp.ps1'
        Write-Warning "Spoof Intelligence not available or permission denied: $_"
    }

    Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Exporting mail security policies..." -CurrentOperation "Saved: MailSec_AntiSpam.json, MailSec_AntiPhish.json, MailSec_Connectors.json, MailSec_TransportRules.json" -PercentComplete ([int]($step / $totalSteps * 100))


    # ================================
    # ===   Done                    ===
    # ================================
    Write-Progress -Id 1 -Activity $activity -Completed
    Write-Host "Mail Security Audit complete. Results saved to: $outputDir" -ForegroundColor Green

    # ── Summary counts for the orchestrator ────────────────────────────────
    $_domainCount       = @($acceptedDomains).Count
    $_dkimEnabledCount  = @($dkimStatus | Where-Object { $_.DKIMEnabled -eq $true }).Count
    $_dmarcFoundCount   = @($dmarcResults | Where-Object { $_.DMARC -ne 'Not Found' }).Count
    $_spfFoundCount     = @($spfResults | Where-Object { $_.SPF -ne 'Not Found' }).Count

    return @{
        DomainCount  = $_domainCount
        DkimEnabled  = $_dkimEnabledCount
        DmarcFound   = $_dmarcFoundCount
        SpfFound     = $_spfFoundCount
    }
}
