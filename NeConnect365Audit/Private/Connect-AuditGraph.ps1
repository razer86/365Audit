function Connect-AuditGraph {
    <#
    .SYNOPSIS
        Connects to Microsoft Graph using credentials from the audit context.
    .DESCRIPTION
        Reads AppId, TenantId, CertFilePath, and CertPassword from the module-scoped
        audit context (set by Set-AuditContext). Uses certificate-based app-only auth.
        Imports Microsoft.Graph.Identity.DirectoryManagement for org queries.
    #>
    [CmdletBinding()]
    param()

    if (Get-MgContext) {
        Write-Verbose "Already connected to Microsoft Graph."
        return
    }

    $ctx = Get-AuditContext

    if (-not ($ctx.AppId -and $ctx.CertFilePath -and $ctx.TenantId)) {
        throw "Audit context is missing AppId, CertFilePath, or TenantId. Cannot connect to Graph."
    }

    Write-Host "Connecting to Microsoft Graph (app-only auth)..." -ForegroundColor Cyan
    try {
        $_cert = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new(
            $ctx.CertFilePath, $ctx.CertPassword)
        Connect-MgGraph -ClientId $ctx.AppId -TenantId $ctx.TenantId -Certificate $_cert -NoWelcome -ErrorAction Stop
    }
    catch {
        throw "Failed to connect to Microsoft Graph: $($_.Exception.Message)"
    }

    Write-Host "Connected to Microsoft Graph." -ForegroundColor Green

    # Import the core sub-module for org queries (Get-MgOrganization)
    Import-GraphModule -ModuleName 'Microsoft.Graph.Identity.DirectoryManagement'
}
