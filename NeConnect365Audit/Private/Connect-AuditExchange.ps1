function Connect-AuditExchange {
    <#
    .SYNOPSIS
        Connects to Exchange Online using credentials from the audit context.
    #>
    [CmdletBinding()]
    param()

    # Check if already connected
    $_exoConnected = Get-ConnectionInformation -ErrorAction SilentlyContinue |
        Where-Object { $_.State -eq 'Connected' }
    if ($_exoConnected) {
        Write-Verbose "Already connected to Exchange Online."
        return
    }

    $ctx = Get-AuditContext

    if (-not ($ctx.AppId -and $ctx.CertFilePath)) {
        throw "Audit context is missing AppId or CertFilePath. Cannot connect to Exchange Online."
    }

    # Resolve the initial .onmicrosoft.com domain for EXO connection
    $org = Get-MgOrganization -ErrorAction Stop | Select-Object -First 1
    $_orgDomain = $org.VerifiedDomains |
        Where-Object { $_.IsInitial -eq $true } |
        Select-Object -ExpandProperty Name -First 1

    if (-not $_orgDomain) {
        throw "Could not resolve initial .onmicrosoft.com domain for Exchange Online connection."
    }

    Write-Host "Connecting to Exchange Online (app-only auth)..." -ForegroundColor Cyan
    Connect-ExchangeOnline `
        -AppId               $ctx.AppId `
        -Organization        $_orgDomain `
        -CertificateFilePath $ctx.CertFilePath `
        -CertificatePassword $ctx.CertPassword `
        -ShowBanner:$false `
        -ErrorAction Stop

    Write-Host "Connected to Exchange Online." -ForegroundColor Green
}
