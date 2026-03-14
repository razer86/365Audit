<#
.SYNOPSIS
    One-time tenant setup for the 365Audit toolkit.

.DESCRIPTION
    First run  : Creates the 'NeConnect MSA Audit Toolkit' app registration with all
                 required permissions (Graph, Exchange, SharePoint), grants admin consent,
                 generates a self-signed certificate (.pfx), and prints all credentials
                 (AppId, TenantId, CertBase64, CertPassword) for storage in Hudu.

    Subsequent : Checks the existing certificate's expiry. If expiring within 30 days
                 (or -Force is used), generates and uploads a new certificate.

    At audit runtime, techs provide -AppId, -TenantId, -CertBase64, and -CertPassword
    to Start-365Audit.ps1. The script decodes the Base64 cert to a temporary .pfx,
    uses it for all module connections, and deletes the temp file on exit.

.PARAMETER AppName
    Display name for the Azure app registration.
    Default: 'NeConnect MSA Audit Toolkit'

.PARAMETER CertExpiryYears
    Validity period for the generated certificate (1–5 years). Default: 2.

.PARAMETER Force
    Generate a new certificate even when the existing one is not near expiry.

.EXAMPLE
    .\Setup-365AuditApp.ps1
    Interactive setup in the customer's tenant.

.EXAMPLE
    .\Setup-365AuditApp.ps1 -Force
    Force certificate renewal even when the current certificate is healthy.

.NOTES
    Author      : Raymond Slater
    Version     : 2.2.1
    Change Log  : See CHANGELOG.md

.LINK
    https://github.com/razer86/365Audit
#>

#Requires -Version 7.4

[CmdletBinding(SupportsShouldProcess)]
param (
    [string]$AppName = 'NeConnect MSA Audit Toolkit',

    [ValidateRange(1, 5)]
    [int]$CertExpiryYears = 2,

    [switch]$Force
)

$ScriptVersion      = '2.2.1'
$ErrorActionPreference = 'Stop'
$ProgressPreference    = 'SilentlyContinue'

# Microsoft Graph service principal app ID (constant in all Azure tenants)
$script:GraphResourceAppId = '00000003-0000-0000-c000-000000000000'

# Required Microsoft Graph application permissions for app-only audit authentication
$script:GraphPermissions = @(
    'Organization.Read.All',
    'Directory.Read.All',
    'User.Read.All',
    'Reports.Read.All',
    'Policy.Read.All',
    'UserAuthenticationMethod.Read.All',
    'RoleManagement.Read.Directory',
    'Group.Read.All',
    'AuditLog.Read.All',
    'SecurityEvents.Read.All'
)

# Office 365 Exchange Online service principal app ID (constant in all Azure tenants)
$script:ExchangeResourceAppId = '00000002-0000-0ff1-ce00-000000000000'

# Required Exchange Online application permission for app-only PowerShell authentication
$script:ExchangePermissions = @('Exchange.ManageAsApp')

# SharePoint Online service principal app ID (constant in all Azure tenants)
$script:SharePointResourceAppId = '00000003-0000-0ff1-ce00-000000000000'

# Required SharePoint Online application permission for app-only access.
# Sites.FullControl.All is the minimum required for SharePoint tenant admin API calls
# (Get-PnPTenant, Get-PnPTenantSite) when using app-only auth — even for read operations.
$script:SharePointPermissions = @('Sites.FullControl.All')

# Days before certificate expiry to trigger a warning / offer rotation
$script:ExpiryWarnDays = 30

# Output directory for generated .pfx files (alongside this script)
$script:CertOutputDir = $PSScriptRoot


# ============================================================
# Helper: write formatted status messages
# ============================================================
function Write-Status {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory, Position = 0)] [string]$Message,
        [ValidateSet('Info', 'Success', 'Error', 'Warning')]
        [string]$Type = 'Info'
    )
    $map = @{
        Info    = @{ Prefix = '[INFO]';    Color = 'Cyan' }
        Success = @{ Prefix = '[SUCCESS]'; Color = 'Green' }
        Error   = @{ Prefix = '[ERROR]';   Color = 'Red' }
        Warning = @{ Prefix = '[WARNING]'; Color = 'Yellow' }
    }
    Write-Host "$($map[$Type].Prefix) $Message" -ForegroundColor $map[$Type].Color
}


# ============================================================
# Connect to Microsoft Graph with admin scopes
# Returns the tenant ID string
# ============================================================
function Connect-GraphForSetup {
    [CmdletBinding()]
    param()

    if (Get-MgContext) {
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    }

    $scopes = @(
        'Application.ReadWrite.All'
        'AppRoleAssignment.ReadWrite.All'
        'Directory.Read.All'
        'RoleManagement.ReadWrite.Directory'
    )

    Write-Status 'Connecting to Microsoft Graph (browser window will open)...'
    Connect-MgGraph -Scopes $scopes -ContextScope Process -NoWelcome -ErrorAction Stop

    $ctx = Get-MgContext
    Write-Status "Connected — Tenant: $($ctx.TenantId)" -Type Success
    return $ctx.TenantId
}


# ============================================================
# Resolve AppRole IDs by name from a service principal
# Returns objects with Id, Name, ResourceSpId
# ============================================================
function Resolve-AppRoleIds {
    [CmdletBinding()]
    param (
        [string]   $ResourceAppId,
        [string[]] $PermissionNames
    )

    $sp = Get-MgServicePrincipal -Filter "appId eq '$ResourceAppId'" -ErrorAction Stop
    if (-not $sp) {
        throw "Service principal for appId '$ResourceAppId' not found in tenant."
    }

    foreach ($name in $PermissionNames) {
        $role = $sp.AppRoles | Where-Object { $_.Value -eq $name -and $_.AllowedMemberTypes -contains 'Application' }
        if (-not $role) {
            throw "Application permission '$name' not found on service principal '$ResourceAppId'."
        }
        [PSCustomObject]@{
            Id           = $role.Id
            Name         = $name
            ResourceSpId = $sp.Id
        }
    }
}


# ============================================================
# Ensure a service principal exists for our app
# ============================================================
function Ensure-ServicePrincipal {
    [CmdletBinding()]
    param ([string]$AppId)

    $sp = Get-MgServicePrincipal -Filter "appId eq '$AppId'" -ErrorAction SilentlyContinue
    if (-not $sp) {
        Write-Status 'Creating service principal for app...'
        $sp = New-MgServicePrincipal -AppId $AppId -ErrorAction Stop
    }
    return $sp
}


# ============================================================
# Grant admin consent for a list of resolved permissions
# ============================================================
function Grant-AdminConsent {
    [CmdletBinding()]
    param (
        [string]   $OurSpId,
        [object[]] $Permissions
    )

    foreach ($perm in $Permissions) {
        $existing = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $OurSpId -ErrorAction SilentlyContinue |
            Where-Object { $_.AppRoleId -eq $perm.Id }

        if ($existing) {
            Write-Verbose "Permission '$($perm.Name)' already granted — skipping."
            continue
        }

        New-MgServicePrincipalAppRoleAssignment `
            -ServicePrincipalId $OurSpId `
            -PrincipalId        $OurSpId `
            -ResourceId         $perm.ResourceSpId `
            -AppRoleId          $perm.Id `
            -ErrorAction Stop | Out-Null

        Write-Verbose "Granted: $($perm.Name)"
    }
}


# ============================================================
# Assign Exchange Administrator Entra role to service principal
# Required for Exchange Online PowerShell app-only authentication
# ============================================================
function Set-ExchangeAdminRole {
    [CmdletBinding()]
    param ([string]$ServicePrincipalId)

    # Activate the role in the tenant if not yet activated (lazy-loaded roles)
    $role = Get-MgDirectoryRole -Filter "displayName eq 'Exchange Administrator'" -ErrorAction SilentlyContinue
    if (-not $role) {
        $template = Get-MgDirectoryRoleTemplate -ErrorAction Stop |
            Where-Object { $_.DisplayName -eq 'Exchange Administrator' }
        if (-not $template) { throw "Exchange Administrator role template not found in tenant." }
        $role = New-MgDirectoryRole -RoleTemplateId $template.Id -ErrorAction Stop
    }

    # Skip if already assigned (-All prevents pagination from missing members in large tenants)
    $existing = Get-MgDirectoryRoleMember -DirectoryRoleId $role.Id -All -ErrorAction SilentlyContinue |
        Where-Object { $_.Id -eq $ServicePrincipalId }
    if ($existing) {
        Write-Verbose 'Exchange Administrator role already assigned — skipping.'
        return
    }

    try {
        $body = @{ '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$ServicePrincipalId" }
        New-MgDirectoryRoleMemberByRef -DirectoryRoleId $role.Id -BodyParameter $body -ErrorAction Stop
        Write-Verbose "Exchange Administrator role assigned to service principal."
    }
    catch {
        if ($_.Exception.Message -match 'already exist') {
            Write-Verbose 'Exchange Administrator role already assigned — skipping.'
        }
        else { throw }
    }
}


# ============================================================
# Analyse existing key (certificate) credentials on the app
# ============================================================
function Get-CertificateStatus {
    [CmdletBinding()]
    param (
        [Microsoft.Graph.PowerShell.Models.MicrosoftGraphApplication]$App
    )

    $now    = Get-Date
    $active = @($App.KeyCredentials | Where-Object { $_.Usage -eq 'Verify' -and $_.EndDateTime -gt $now })
    $soon   = @($active | Where-Object { $_.EndDateTime -lt $now.AddDays($script:ExpiryWarnDays) })
    $next   = $active | Sort-Object EndDateTime | Select-Object -First 1

    [PSCustomObject]@{
        HasActive           = $active.Count -gt 0
        ExpiresWithin30Days = $soon.Count -gt 0
        Soonest             = $next
    }
}


# ============================================================
# Generate a self-signed certificate, export the .pfx, and
# upload the public key to the Entra app registration.
# Returns: [PSCustomObject] PfxPath, PlainPassword, ExpiryDate
# ============================================================
function New-AuditCertificate {
    [CmdletBinding()]
    param (
        [string] $AppObjectId,
        [string] $AppId,
        [int]    $ExpiryYears
    )

    $certName  = "365Audit-$AppId"
    $notAfter  = [System.DateTimeOffset]::UtcNow.AddYears($ExpiryYears)

    # Generate a random 32-char password for the .pfx
    $chars     = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnpqrstuvwxyz23456789!@#%&*'
    $plainPwd  = -join ((1..32) | ForEach-Object { $chars[(Get-Random -Maximum $chars.Length)] })
    $securePwd = ConvertTo-SecureString $plainPwd -AsPlainText -Force

    # Export .pfx alongside this script
    $pfxPath = Join-Path $script:CertOutputDir "$certName.pfx"
    $days    = $ExpiryYears * 365

    if ($IsLinux -or $IsMacOS) {
        # New-SelfSignedCertificate and Cert:\ are Windows-only.
        # Use openssl (must be installed: apt install openssl / brew install openssl).
        $tmpKey  = [System.IO.Path]::GetTempFileName() + '.key'
        $tmpCert = [System.IO.Path]::GetTempFileName() + '.crt'
        try {
            & openssl req -x509 -newkey rsa:2048 -keyout $tmpKey -out $tmpCert `
                -days $days -nodes -subj "/CN=$certName" 2>$null
            if ($LASTEXITCODE -ne 0) { throw "openssl certificate generation failed (exit $LASTEXITCODE)" }

            # -legacy required on OpenSSL 3.x to produce a .pfx readable by .NET
            & openssl pkcs12 -export -legacy -out $pfxPath `
                -inkey $tmpKey -in $tmpCert -passout "pass:$plainPwd" 2>$null
            if ($LASTEXITCODE -ne 0) { throw "openssl pfx export failed (exit $LASTEXITCODE)" }

            $rawData = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new(
                [System.IO.File]::ReadAllBytes($tmpCert)
            ).RawData
        }
        finally {
            Remove-Item $tmpKey, $tmpCert -ErrorAction SilentlyContinue
        }
    }
    else {
        # -KeySpec KeyExchange forces the legacy CSP provider (not CNG).
        # EXO Connect-ExchangeOnline -CertificateFilePath requires a CSP cert.
        $cert = New-SelfSignedCertificate `
            -Subject           "CN=$certName" `
            -CertStoreLocation 'Cert:\CurrentUser\My' `
            -NotAfter          $notAfter.LocalDateTime `
            -KeySpec           KeyExchange `
            -ErrorAction Stop

        $cert | Export-PfxCertificate -FilePath $pfxPath -Password $securePwd -ErrorAction Stop | Out-Null
        $rawData = [byte[]]$cert.RawData

        # Remove from local cert store — the .pfx is the portable copy
        Remove-Item "Cert:\CurrentUser\My\$($cert.Thumbprint)" -ErrorAction SilentlyContinue
    }

    # Upload the public key to the Entra app registration.
    # Graph SDK requires DateTimeOffset (not DateTime) for key credential dates.
    Update-MgApplication -ApplicationId $AppObjectId -KeyCredentials @(
        @{
            type          = 'AsymmetricX509Cert'
            usage         = 'Verify'
            key           = $rawData
            displayName   = $certName
            startDateTime = [System.DateTimeOffset]::UtcNow
            endDateTime   = $notAfter
        }
    ) -ErrorAction Stop

    # Encode the .pfx as base64 so it can be stored as a Hudu secret
    $certBase64 = [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($pfxPath))

    return [PSCustomObject]@{
        PfxPath       = $pfxPath
        PlainPassword = $plainPwd
        ExpiryDate    = $notAfter.LocalDateTime
        CertBase64    = $certBase64
    }
}


# ============================================================
# Open Azure portal to API permissions page for admin consent
# ============================================================
function Request-AdminConsent {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)] [string]$ApplicationId,
        [Parameter(Mandatory)] [string]$TenantName
    )

    $portalUrl = "https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/~/CallAnAPI/appId/$ApplicationId/isMSAApp~/false"

    $sep = '=' * 72
    Write-Host "`n$sep" -ForegroundColor Yellow
    Write-Host '  ACTION REQUIRED: Grant admin consent in the Azure portal' -ForegroundColor Yellow
    Write-Host $sep -ForegroundColor Yellow
    Write-Host ''
    Write-Host '  Waiting for Entra ID to finish replicating the application...' -ForegroundColor DarkYellow
    Write-Host '  This usually takes 5-10 seconds.' -ForegroundColor DarkYellow

    Start-Sleep -Seconds 10

    Write-Host ''
    Write-Host '  Opening Azure portal — API Permissions page for this app...' -ForegroundColor Cyan
    Write-Host "  In the portal, click: " -NoNewline -ForegroundColor White
    Write-Host "'Grant admin consent for $TenantName'" -ForegroundColor Yellow
    Write-Host "  then click 'Yes' to confirm." -ForegroundColor White
    Write-Host ''
    Write-Host "  If the browser doesn't open automatically, use this URL:" -ForegroundColor DarkCyan
    Write-Host "  $portalUrl" -ForegroundColor Cyan
    Write-Host "$sep`n" -ForegroundColor Yellow

    try {
        if ($IsLinux) {
            xdg-open $portalUrl
        } elseif ($IsMacOS) {
            open $portalUrl
        } else {
            Start-Process $portalUrl -ErrorAction Stop
        }
    }
    catch {
        Write-Warning "Unable to open browser automatically: $($_.Exception.Message)"
    }
}



# ============================================================
# Print credentials in a clearly formatted block
# ============================================================
function Write-CredentialSummary {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingPlainTextForPassword', 'CertPassword',
        Justification = 'Intentionally plain text — displayed once so the operator can store it in the password manager.')]
    [CmdletBinding()]
    param (
        [string]   $AppId,
        [string]   $TenantId,
        [string]   $CertBase64,
        [string]   $CertPassword,
        [datetime] $CertExpiry
    )

    $sep = '=' * 72
    Write-Host "`n$sep" -ForegroundColor Cyan
    Write-Host '  NeConnect MSA Audit Toolkit — Store these credentials in Hudu' -ForegroundColor Cyan
    Write-Host $sep -ForegroundColor Cyan
    Write-Host "  App ID (Client ID) : $AppId"
    Write-Host "  Tenant ID          : $TenantId"
    Write-Host "  Cert Base64        : $CertBase64" -ForegroundColor Yellow
    Write-Host "  Cert Password      : $CertPassword" -ForegroundColor Yellow
    Write-Host "  Cert Expires       : $($CertExpiry.ToString('yyyy-MM-dd'))"
    Write-Host ''
    Write-Host '  Run the audit with:' -ForegroundColor DarkCyan
    Write-Host "  .\Start-365Audit.ps1 -AppId '$AppId' -TenantId '$TenantId' -CertBase64 '<paste base64>' -CertPassword (Read-Host -AsSecureString 'Cert Password')" -ForegroundColor Cyan
    Write-Host "$sep`n" -ForegroundColor Cyan
}



# ============================================================
# Main
# ============================================================
try {
    Write-Host "`n365Audit App Setup v$ScriptVersion`n" -ForegroundColor Cyan

    # Ensure required Graph modules are installed
    foreach ($mod in @(
            'Microsoft.Graph.Authentication'
            'Microsoft.Graph.Applications'
            'Microsoft.Graph.Identity.DirectoryManagement')) {
        if (-not (Get-Module -ListAvailable -Name $mod)) {
            Write-Status "Installing $mod..." -Type Warning
            Install-Module $mod -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        }
        if (-not (Get-Module -Name $mod)) {
            Import-Module $mod -ErrorAction Stop
        }
    }

    $tenantId = Connect-GraphForSetup

    $orgName = (Get-MgOrganization -ErrorAction Stop | Select-Object -First 1).DisplayName
    Write-Status "Tenant: $orgName ($tenantId)"

    # ----------------------------------------------------------
    # Look for existing app
    # ----------------------------------------------------------
    Write-Status "Searching for existing app: '$AppName'..."
    $existingApp = Get-MgApplication -Filter "displayName eq '$AppName'" -ErrorAction Stop |
        Select-Object -First 1

    if ($existingApp) {
        Write-Status "App found — App ID: $($existingApp.AppId)" -Type Info

        # ----------------------------------------------------------
        # Ensure all required Graph and Exchange permissions are present.
        # Checks per individual permission so new ones added in future
        # versions are automatically applied on re-run.
        # ----------------------------------------------------------
        $graphPerms      = @(Resolve-AppRoleIds -ResourceAppId $script:GraphResourceAppId      -PermissionNames $script:GraphPermissions)
        $exchangePerms   = @(Resolve-AppRoleIds -ResourceAppId $script:ExchangeResourceAppId   -PermissionNames $script:ExchangePermissions)
        $sharePointPerms = @(Resolve-AppRoleIds -ResourceAppId $script:SharePointResourceAppId -PermissionNames $script:SharePointPermissions)

        $currentIds        = @($existingApp.RequiredResourceAccess.ResourceAccess | ForEach-Object { $_.Id })
        $missingGraph      = @($graphPerms      | Where-Object { $_.Id -notin $currentIds })
        $missingExchange   = @($exchangePerms   | Where-Object { $_.Id -notin $currentIds })
        $missingSharePoint = @($sharePointPerms | Where-Object { $_.Id -notin $currentIds })

        if ($missingGraph.Count -gt 0 -or $missingExchange.Count -gt 0 -or $missingSharePoint.Count -gt 0) {
            $missingNames = ($missingGraph + $missingExchange + $missingSharePoint | ForEach-Object { $_.Name }) -join ', '
            Write-Status "Adding missing permissions: $missingNames" -Type Warning

            $resourceAccess = @(
                @{
                    resourceAppId  = $script:GraphResourceAppId
                    resourceAccess = @($graphPerms      | ForEach-Object { @{ id = $_.Id.ToString(); type = 'Role' } })
                },
                @{
                    resourceAppId  = $script:ExchangeResourceAppId
                    resourceAccess = @($exchangePerms   | ForEach-Object { @{ id = $_.Id.ToString(); type = 'Role' } })
                },
                @{
                    resourceAppId  = $script:SharePointResourceAppId
                    resourceAccess = @($sharePointPerms | ForEach-Object { @{ id = $_.Id.ToString(); type = 'Role' } })
                }
            )

            Update-MgApplication -ApplicationId $existingApp.Id -RequiredResourceAccess $resourceAccess -ErrorAction Stop
            $ourSp = Ensure-ServicePrincipal -AppId $existingApp.AppId

            if ($missingGraph.Count -gt 0)      { Grant-AdminConsent -OurSpId $ourSp.Id -Permissions $missingGraph }
            if ($missingSharePoint.Count -gt 0) { Grant-AdminConsent -OurSpId $ourSp.Id -Permissions $missingSharePoint }
            if ($missingExchange.Count -gt 0) {
                Grant-AdminConsent -OurSpId $ourSp.Id -Permissions $missingExchange
                Write-Status 'Assigning Exchange Administrator role to service principal...'
                Set-ExchangeAdminRole -ServicePrincipalId $ourSp.Id
            }
            Write-Status 'Permissions updated and admin consent granted.' -Type Success
            Request-AdminConsent -ApplicationId $existingApp.AppId -TenantName $orgName
        }
        else {
            Write-Verbose 'All required permissions already present on app.'
        }

        $certStatus = Get-CertificateStatus -App $existingApp

        if ($certStatus.HasActive) {
            $expiry = $certStatus.Soonest.EndDateTime
            Write-Status "Active certificate expires: $($expiry.ToString('yyyy-MM-dd'))"

            if ($certStatus.ExpiresWithin30Days) {
                Write-Status "Certificate expiring within $script:ExpiryWarnDays days — generating new certificate." -Type Warning
                $newCert = New-AuditCertificate -AppObjectId $existingApp.Id -AppId $existingApp.AppId -ExpiryYears $CertExpiryYears
                Write-CredentialSummary -AppId $existingApp.AppId -TenantId $tenantId -CertBase64 $newCert.CertBase64 -CertPassword $newCert.PlainPassword -CertExpiry $newCert.ExpiryDate
            }
            elseif ($Force) {
                Write-Status '-Force specified — generating new certificate.' -Type Warning
                $newCert = New-AuditCertificate -AppObjectId $existingApp.Id -AppId $existingApp.AppId -ExpiryYears $CertExpiryYears
                Write-CredentialSummary -AppId $existingApp.AppId -TenantId $tenantId -CertBase64 $newCert.CertBase64 -CertPassword $newCert.PlainPassword -CertExpiry $newCert.ExpiryDate
            }
            else {
                Write-Status "Certificate is healthy — re-run the audit with your existing .pfx file." -Type Success
                $sep = '=' * 72
                Write-Host "`n$sep" -ForegroundColor Cyan
                Write-Host "  App ID (Client ID) : $($existingApp.AppId)"
                Write-Host "  Tenant ID          : $tenantId"
                Write-Host "  Cert Expires       : $($expiry.ToString('yyyy-MM-dd'))"
                Write-Host "  Use -Force to rotate the certificate regardless of expiry." -ForegroundColor DarkCyan
                Write-Host "$sep`n" -ForegroundColor Cyan
            }
        }
        else {
            Write-Status 'No active certificate found — generating new certificate.' -Type Warning
            $newCert = New-AuditCertificate -AppObjectId $existingApp.Id -AppId $existingApp.AppId -ExpiryYears $CertExpiryYears
            Write-CredentialSummary -AppId $existingApp.AppId -TenantId $tenantId -CertBase64 $newCert.CertBase64 -CertPassword $newCert.PlainPassword -CertExpiry $newCert.ExpiryDate
        }
    }
    else {
        # ----------------------------------------------------------
        # App does not exist — create it
        # ----------------------------------------------------------
        Write-Status "App not found — creating '$AppName'..." -Type Info

        Write-Status 'Resolving permission IDs...'
        $graphPerms      = @(Resolve-AppRoleIds -ResourceAppId $script:GraphResourceAppId      -PermissionNames $script:GraphPermissions)
        $exchangePerms   = @(Resolve-AppRoleIds -ResourceAppId $script:ExchangeResourceAppId   -PermissionNames $script:ExchangePermissions)
        $sharePointPerms = @(Resolve-AppRoleIds -ResourceAppId $script:SharePointResourceAppId -PermissionNames $script:SharePointPermissions)

        # Graph SDK v2 requires camelCase keys and explicit string GUIDs in hashtables
        $resourceAccess = @(
            @{
                resourceAppId  = $script:GraphResourceAppId
                resourceAccess = @($graphPerms      | ForEach-Object { @{ id = $_.Id.ToString(); type = 'Role' } })
            },
            @{
                resourceAppId  = $script:ExchangeResourceAppId
                resourceAccess = @($exchangePerms   | ForEach-Object { @{ id = $_.Id.ToString(); type = 'Role' } })
            },
            @{
                resourceAppId  = $script:SharePointResourceAppId
                resourceAccess = @($sharePointPerms | ForEach-Object { @{ id = $_.Id.ToString(); type = 'Role' } })
            }
        )

        if ($PSCmdlet.ShouldProcess($AppName, 'Create Entra app registration')) {
            $newApp = New-MgApplication `
                -DisplayName            $AppName `
                -RequiredResourceAccess $resourceAccess `
                -ErrorAction Stop

            Write-Status "App created — App ID: $($newApp.AppId)" -Type Success

            Write-Status 'Creating service principal (waiting for Entra replication)...'
            $ourSp = Ensure-ServicePrincipal -AppId $newApp.AppId
            Start-Sleep -Seconds 5   # Allow Entra ID to replicate before granting consent

            Write-Status 'Granting admin consent for all permissions...'
            Grant-AdminConsent -OurSpId $ourSp.Id -Permissions $graphPerms
            Grant-AdminConsent -OurSpId $ourSp.Id -Permissions $exchangePerms
            Grant-AdminConsent -OurSpId $ourSp.Id -Permissions $sharePointPerms

            Write-Status 'Assigning Exchange Administrator role to service principal...'
            Set-ExchangeAdminRole -ServicePrincipalId $ourSp.Id
            Write-Status 'Admin consent granted.' -Type Success

            Request-AdminConsent -ApplicationId $newApp.AppId -TenantName $orgName

            Write-Status "Generating $CertExpiryYears-year certificate..."
            $newCert = New-AuditCertificate -AppObjectId $newApp.Id -AppId $newApp.AppId -ExpiryYears $CertExpiryYears

            Write-CredentialSummary -AppId $newApp.AppId -TenantId $tenantId -CertBase64 $newCert.CertBase64 -CertPassword $newCert.PlainPassword -CertExpiry $newCert.ExpiryDate
        }
    }
}
catch {
    Write-Status "Setup failed: $($_.Exception.Message)" -Type Error
    if ($VerbosePreference -eq 'Continue') {
        Write-Host $_.ScriptStackTrace -ForegroundColor Red
    }
    exit 1
}
finally {
    if (Get-MgContext) {
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    }
}
