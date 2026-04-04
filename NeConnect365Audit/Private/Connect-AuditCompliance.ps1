function Connect-AuditCompliance {
    <#
    .SYNOPSIS
        Connects to Security & Compliance Center using credentials from the audit context.
    #>
    [CmdletBinding()]
    param()

    $ctx = Get-AuditContext

    if (-not ($ctx.AppId -and $ctx.CertFilePath)) {
        Write-Verbose "Audit context missing AppId or CertFilePath — skipping S&C connection."
        return
    }

    # Resolve the initial .onmicrosoft.com domain
    $org = Get-MgOrganization -ErrorAction Stop | Select-Object -First 1
    $_orgDomain = $org.VerifiedDomains |
        Where-Object { $_.IsInitial -eq $true } |
        Select-Object -ExpandProperty Name -First 1

    if (-not $_orgDomain) {
        Write-Warning "Could not resolve initial domain — Security & Compliance tests will be skipped."
        return
    }

    Write-Host "Connecting to Security & Compliance..." -ForegroundColor Cyan
    try {
        Connect-IPPSSession `
            -AppId               $ctx.AppId `
            -Organization        $_orgDomain `
            -CertificateFilePath $ctx.CertFilePath `
            -CertificatePassword $ctx.CertPassword `
            -ShowBanner:$false `
            -ErrorAction Stop
        Write-Host "Connected to Security & Compliance." -ForegroundColor Green
    }
    catch {
        Write-Warning "Could not connect to Security & Compliance: $($_.Exception.Message)"
    }
}
