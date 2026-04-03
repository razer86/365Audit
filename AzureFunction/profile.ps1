# Azure Function App profile — runs once per worker process cold start.
# Authenticates with Managed Identity so Invoke-AzAuditBatch.ps1 can
# retrieve secrets from Key Vault without explicit credentials.

if ($env:IDENTITY_ENDPOINT) {
    Connect-AzAccount -Identity -ErrorAction Stop | Out-Null
    Write-Host "Authenticated with Managed Identity."
}
