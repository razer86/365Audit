# Invoke-MailSecurityAudit.ps1

Connects to Exchange Online and audits mail security configuration — DNS authentication records (DKIM, DMARC, SPF) and Exchange Online Protection policy objects.

Flat data is exported as CSV for the HTML summary. Nested policy objects are exported as JSON for detailed review.

## Required Permissions

Same as the Exchange module — granted automatically by `Setup-365AuditApp.ps1`:

| Permission | Type | Purpose |
|---|---|---|
| `Exchange.ManageAsApp` | Application | Read EOP policy configuration via app-only auth |
| Exchange Administrator | Directory Role | Required alongside `Exchange.ManageAsApp` |

## Output Files

### CSV

| File | Description |
|---|---|
| `MailSec_DKIM.csv` | DKIM signing status and selector CNAMEs per domain |
| `MailSec_DMARC.csv` | DMARC TXT records per accepted domain |
| `MailSec_SPF.csv` | SPF TXT records per accepted domain |

### JSON (supplementary)

| File | Description |
|---|---|
| `MailSec_AntiSpam.json` | Hosted content filter policy objects |
| `MailSec_AntiSpamRules.json` | Hosted content filter rules |
| `MailSec_AntiPhish.json` | Anti-phishing policy objects |
| `MailSec_AntiPhishRules.json` | Anti-phishing rules |
| `MailSec_SpoofIntelligence.json` | Spoof intelligence insights (requires Defender for Office 365) |
| `MailSec_InboundConnectors.json` | Inbound connector configuration |
| `MailSec_OutboundConnectors.json` | Outbound connector configuration |
| `MailSec_TransportRules.json` | Mail flow rule summaries |
