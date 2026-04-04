function Connect-AuditTeams {
    <#
    .SYNOPSIS
        Connects to Microsoft Teams using credentials from the audit context.
    #>
    [CmdletBinding()]
    param()

    # Ensure MicrosoftTeams module is available
    if (-not (Get-Module -ListAvailable -Name 'MicrosoftTeams')) {
        Write-Host "  Required module 'MicrosoftTeams' not found — installing..." -ForegroundColor Yellow
        Install-Module MicrosoftTeams -Scope CurrentUser -Force -AllowClobber -SkipPublisherCheck -ErrorAction Stop
        $_installedMod = Get-Module -ListAvailable -Name 'MicrosoftTeams' | Sort-Object Version -Descending | Select-Object -First 1
        if (-not $_installedMod) {
            throw "Installation of 'MicrosoftTeams' failed."
        }
        Write-Host "  Installed 'MicrosoftTeams' v$($_installedMod.Version)." -ForegroundColor Green
    }

    Import-Module MicrosoftTeams -ErrorAction Stop -WarningAction SilentlyContinue

    # Check if already connected
    try {
        Get-CsTenant -ErrorAction Stop | Out-Null
        Write-Verbose "Already connected to Microsoft Teams."
        return
    }
    catch { <# not connected — continue #> }

    $ctx = Get-AuditContext

    if (-not ($ctx.AppId -and $ctx.TenantId -and $ctx.CertFilePath)) {
        throw "Audit context is missing AppId, TenantId, or CertFilePath. Cannot connect to Teams."
    }

    Write-Host "Connecting to Microsoft Teams (app-only auth)..." -ForegroundColor Cyan

    # Load cert from file path + SecureString (cross-platform compatible)
    $certObj = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new(
        $ctx.CertFilePath, $ctx.CertPassword)

    Connect-MicrosoftTeams `
        -ApplicationId $ctx.AppId `
        -TenantId      $ctx.TenantId `
        -Certificate   $certObj `
        -ErrorAction   Stop

    $certObj.Dispose()
    Write-Host "Connected to Microsoft Teams." -ForegroundColor Green
}
