<#
.SYNOPSIS
    One-time tenant setup for the 365Audit toolkit.

.DESCRIPTION
    First run  : Creates the 'NeConnect MSA Audit Toolkit' app registration with all
                 required permissions (Graph, Exchange), grants admin consent, registers a
                 dedicated PnP interactive auth app for SharePoint, and prints both sets of
                 credentials for storage in Hudu.

    Subsequent : Checks the existing app's secret expiry. If expiring within 30 days
                 (or -Force is used), generates a new secret.

    The app credentials (AppId/AppSecret/TenantId) are optional at audit runtime.
    When provided via -AppId/-AppSecret/-TenantId, Entra and Exchange modules authenticate
    silently without interactive prompts. SharePoint always uses interactive sign-in.

.PARAMETER AppName
    Display name for the Azure app registration.
    Default: 'NeConnect MSA Audit Toolkit'

.PARAMETER SecretExpiryMonths
    Validity period for the generated client secret (1–24 months). Default: 24.

.PARAMETER Force
    Generate a new client secret even when the existing one is not near expiry.

.EXAMPLE
    .\Setup-365AuditApp.ps1
    Interactive setup in the customer's tenant.

.EXAMPLE
    .\Setup-365AuditApp.ps1 -Force
    Force a new secret even when the current one is healthy.

.NOTES
    Author      : Raymond Slater
    Version     : 1.9.1
    Change Log  :
        1.0.0 - Initial release
        1.1.0 - Added Microsoft Graph application permissions required for app-only auth;
                existing apps without Graph permissions are updated automatically
        1.2.0 - Added Request-AdminConsent: opens Azure portal to API permissions page
                and prints instructions after consent is granted
        1.3.0 - Added Exchange.ManageAsApp permission and Exchange Administrator Entra role
                assignment so Exchange Online can authenticate via client credentials
        1.4.0 - Added New-AuditCertificate: generates a self-signed certificate, installs
                it in Cert:\CurrentUser\My, and uploads the public key to the Azure AD app;
                SharePoint admin APIs require azpacr=1 (certificate auth) tokens and reject
                client-secret tokens; certificate thumbprint is now shown in credentials output
        1.5.0 - Added Register-PnPManagementShell: creates a service principal for the PnP
                Management Shell public app and grants tenant-wide admin consent for
                AllSites.FullControl so any technician can use interactive SharePoint auth
                without per-user consent prompts or AADSTS700016 errors
        1.6.0 - Removed certificate management (SharePoint reverted to interactive auth);
                updated description to reflect dual purpose: app credentials for silent
                Entra/Exchange auth + PnP Management Shell consent for SharePoint interactive
        1.7.0 - Removed PnP Management Shell registration (app deprecated in PnP.PowerShell v2);
                added http://localhost as a public client redirect URI to the app registration
                so Connect-PnPOnline -Interactive -ClientId $AuditAppId works from any machine
        1.8.0 - Replaced http://localhost approach with Register-PnPEntraIDAppForInteractiveLogin
                (the PnP-recommended method since Sep 2024); dedicated PnP interactive app is
                registered separately and its App ID is shown alongside main app credentials;
                bumped #Requires to 7.4 (PnP.PowerShell v3 requires PowerShell 7.4+);
                updated PnP module check to enforce MinimumVersion 3.0.0
        1.9.0 - Removed SharePoint (Sites.FullControl.All) from main app registration;
                SharePoint auth is handled entirely by the PnP interactive app which uses
                delegated permissions scoped to the signed-in technician's rights
        1.9.1 - Updated URLs and references from 'MSA Audit Toolkit' to '365Audit' for branding consistency

.LINK
    https://github.com/razer86/365Audit
#>

#Requires -Version 7.4

[CmdletBinding(SupportsShouldProcess)]
param (
    [string]$AppName = 'NeConnect MSA Audit Toolkit',

    [ValidateRange(1, 24)]
    [int]$SecretExpiryMonths = 24,

    [switch]$Force
)

$ScriptVersion      = '1.9.0'
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
    'AuditLog.Read.All'
)

# Office 365 Exchange Online service principal app ID (constant in all Azure tenants)
$script:ExchangeResourceAppId = '00000002-0000-0ff1-ce00-000000000000'

# Required Exchange Online application permission for app-only PowerShell authentication
$script:ExchangePermissions = @('Exchange.ManageAsApp')

# Days before expiry to trigger a warning / offer secret rotation
$script:ExpiryWarnDays = 30


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

    # Skip if already assigned
    $existing = Get-MgDirectoryRoleMember -DirectoryRoleId $role.Id -ErrorAction SilentlyContinue |
        Where-Object { $_.Id -eq $ServicePrincipalId }
    if ($existing) {
        Write-Verbose 'Exchange Administrator role already assigned — skipping.'
        return
    }

    $body = @{ '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$ServicePrincipalId" }
    New-MgDirectoryRoleMemberByRef -DirectoryRoleId $role.Id -BodyParameter $body -ErrorAction Stop
    Write-Verbose "Exchange Administrator role assigned to service principal."
}


# ============================================================
# Analyse existing password credentials
# ============================================================
function Get-SecretStatus {
    [CmdletBinding()]
    param (
        [Microsoft.Graph.PowerShell.Models.MicrosoftGraphApplication]$App
    )

    $now    = Get-Date
    $active = @($App.PasswordCredentials | Where-Object { $_.EndDateTime -gt $now })
    $soon   = @($active | Where-Object { $_.EndDateTime -lt $now.AddDays($script:ExpiryWarnDays) })
    $next   = $active | Sort-Object EndDateTime | Select-Object -First 1

    [PSCustomObject]@{
        HasActive    = $active.Count -gt 0
        ExpiresWithin30Days = $soon.Count -gt 0
        Soonest      = $next
    }
}


# ============================================================
# Add a new client secret to the app
# ============================================================
function New-AuditSecret {
    [CmdletBinding()]
    param (
        [string] $ApplicationObjectId,
        [int]    $ExpiryMonths
    )

    $now = Get-Date
    Add-MgApplicationPassword `
        -ApplicationId      $ApplicationObjectId `
        -PasswordCredential @{
            DisplayName   = "AuditToolkit-$(Get-Date -Format 'yyyy-MM')"
            StartDateTime = $now
            EndDateTime   = $now.AddMonths($ExpiryMonths)
        } `
        -ErrorAction Stop
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
# Register a dedicated PnP interactive auth app
# Uses the PnP-recommended Register-PnPEntraIDAppForInteractiveLogin cmdlet.
# A browser window will open for sign-in during registration.
# ============================================================
function Register-PnPInteractiveApp {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$Tenant,
        [Parameter(Mandatory)] [string]$AppName
    )

    # Check if the PnP interactive app already exists by display name
    $existingPnpApp = Get-MgApplication -Filter "displayName eq '$AppName'" -ErrorAction SilentlyContinue |
        Select-Object -First 1
    if ($existingPnpApp) {
        Write-Status "PnP interactive app already registered — App ID: $($existingPnpApp.AppId)" -Type Info
        return $existingPnpApp.AppId
    }

    if (-not (Get-Module -ListAvailable -Name PnP.PowerShell | Where-Object Version -ge '3.0.0')) {
        Write-Status 'Installing PnP.PowerShell v3+...' -Type Warning
        Install-Module PnP.PowerShell -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
    }
    Import-Module PnP.PowerShell -WarningAction SilentlyContinue

    Write-Status "Registering PnP interactive auth app '$AppName'..."
    Write-Host '  A browser window will open — sign in as a Global or Application Administrator.' -ForegroundColor DarkCyan

    Register-PnPEntraIDAppForInteractiveLogin `
        -ApplicationName $AppName `
        -Tenant          $Tenant `
        -ErrorAction Stop | Out-Null

    # The cmdlet writes to the console but does not return a usable object;
    # resolve the App ID via Graph using the known display name.
    $createdApp = Get-MgApplication -Filter "displayName eq '$AppName'" -ErrorAction SilentlyContinue |
        Select-Object -First 1
    if (-not $createdApp) {
        throw "PnP app registration appeared to succeed but app '$AppName' was not found in Entra ID."
    }

    Write-Status "PnP interactive app registered — App ID: $($createdApp.AppId)" -Type Success
    return $createdApp.AppId
}


# ============================================================
# Print credentials in a clearly formatted block
# ============================================================
function Show-Credentials {
    [CmdletBinding()]
    param (
        [string]   $AppId,
        [string]   $TenantId,
        [string]   $SecretText,
        [datetime] $SecretExpiry,
        [string]   $PnPAppId
    )

    $sep = '=' * 72
    Write-Host "`n$sep" -ForegroundColor Cyan
    Write-Host '  NeConnect MSA Audit Toolkit — Store these credentials in Hudu' -ForegroundColor Cyan
    Write-Host $sep -ForegroundColor Cyan
    Write-Host "  App ID (Client ID) : $AppId"
    Write-Host "  Tenant ID          : $TenantId"
    Write-Host "  Client Secret      : $SecretText" -ForegroundColor Yellow
    Write-Host "  Secret Expires     : $($SecretExpiry.ToString('yyyy-MM-dd'))"
    if ($PnPAppId) {
        Write-Host ''
        Write-Host "  PnP App ID         : $PnPAppId" -ForegroundColor Cyan
        Write-Host '  (Used for interactive SharePoint authentication)' -ForegroundColor DarkCyan
    }
    Write-Host ''
    Write-Host '  Run the audit with:' -ForegroundColor DarkCyan
    if ($PnPAppId) {
        Write-Host "  .\Start-365Audit.ps1 -AppId '$AppId' -AppSecret '$SecretText' -TenantId '$TenantId' -PnPAppId '$PnPAppId'" -ForegroundColor Cyan
    }
    else {
        Write-Host "  .\Start-365Audit.ps1 -AppId '$AppId' -AppSecret '$SecretText' -TenantId '$TenantId'" -ForegroundColor Cyan
    }
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
        # Ensure Microsoft Graph and Exchange permissions are present.
        # ----------------------------------------------------------
        $hasGraphPerms    = $existingApp.RequiredResourceAccess |
            Where-Object { $_.ResourceAppId -eq $script:GraphResourceAppId }
        $hasExchangePerms = $existingApp.RequiredResourceAccess |
            Where-Object { $_.ResourceAppId -eq $script:ExchangeResourceAppId }

        if (-not $hasGraphPerms -or -not $hasExchangePerms) {
            Write-Status 'Adding missing permissions to existing app...' -Type Warning

            $graphPerms    = @(Resolve-AppRoleIds -ResourceAppId $script:GraphResourceAppId    -PermissionNames $script:GraphPermissions)
            $exchangePerms = @(Resolve-AppRoleIds -ResourceAppId $script:ExchangeResourceAppId -PermissionNames $script:ExchangePermissions)

            $resourceAccess = @(
                @{
                    resourceAppId  = $script:GraphResourceAppId
                    resourceAccess = @($graphPerms    | ForEach-Object { @{ id = $_.Id.ToString(); type = 'Role' } })
                },
                @{
                    resourceAppId  = $script:ExchangeResourceAppId
                    resourceAccess = @($exchangePerms | ForEach-Object { @{ id = $_.Id.ToString(); type = 'Role' } })
                }
            )

            Update-MgApplication -ApplicationId $existingApp.Id -RequiredResourceAccess $resourceAccess -ErrorAction Stop
            $ourSp = Ensure-ServicePrincipal -AppId $existingApp.AppId

            if (-not $hasGraphPerms)    { Grant-AdminConsent -OurSpId $ourSp.Id -Permissions $graphPerms }
            if (-not $hasExchangePerms) { Grant-AdminConsent -OurSpId $ourSp.Id -Permissions $exchangePerms }

            Write-Status 'Assigning Exchange Administrator role to service principal...'
            Set-ExchangeAdminRole -ServicePrincipalId $ourSp.Id
            Write-Status 'Permissions updated and admin consent granted.' -Type Success
            Request-AdminConsent -ApplicationId $existingApp.AppId -TenantName $orgName
        }
        else {
            Write-Verbose 'All required permissions already present on app.'
        }

        # Register the dedicated PnP interactive auth app for SharePoint
        $pnpInteractiveAppId = Register-PnPInteractiveApp -Tenant $tenantId -AppName "$AppName (SharePoint)"

        $status = Get-SecretStatus -App $existingApp

        if ($status.HasActive) {
            $expiry = $status.Soonest.EndDateTime
            Write-Status "Active secret expires: $($expiry.ToString('yyyy-MM-dd'))"

            if ($status.ExpiresWithin30Days) {
                Write-Status "Secret expiring within $script:ExpiryWarnDays days — generating new secret." -Type Warning
                $newSecret = New-AuditSecret -ApplicationObjectId $existingApp.Id -ExpiryMonths $SecretExpiryMonths
                Show-Credentials -AppId $existingApp.AppId -TenantId $tenantId -SecretText $newSecret.SecretText -SecretExpiry $newSecret.EndDateTime -PnPAppId $pnpInteractiveAppId
            }
            elseif ($Force) {
                Write-Status '-Force specified — generating new secret.' -Type Warning
                $newSecret = New-AuditSecret -ApplicationObjectId $existingApp.Id -ExpiryMonths $SecretExpiryMonths
                Show-Credentials -AppId $existingApp.AppId -TenantId $tenantId -SecretText $newSecret.SecretText -SecretExpiry $newSecret.EndDateTime -PnPAppId $pnpInteractiveAppId
            }
            else {
                Write-Status "Secret is healthy — re-run the audit with your existing credentials." -Type Success
                $sep = '=' * 72
                Write-Host "`n$sep" -ForegroundColor Cyan
                Write-Host "  App ID (Client ID) : $($existingApp.AppId)"
                Write-Host "  Tenant ID          : $tenantId"
                if ($pnpInteractiveAppId) {
                    Write-Host ''
                    Write-Host "  PnP App ID         : $pnpInteractiveAppId" -ForegroundColor Cyan
                    Write-Host '  (Used for interactive SharePoint authentication)' -ForegroundColor DarkCyan
                    Write-Host ''
                    Write-Host "  .\Start-365Audit.ps1 -AppId '$($existingApp.AppId)' -AppSecret '<your secret>' -TenantId '$tenantId' -PnPAppId '$pnpInteractiveAppId'" -ForegroundColor Cyan
                }
                else {
                    Write-Host "  .\Start-365Audit.ps1 -AppId '$($existingApp.AppId)' -AppSecret '<your secret>' -TenantId '$tenantId'" -ForegroundColor Cyan
                }
                Write-Host "  Use -Force to rotate the secret regardless of expiry." -ForegroundColor DarkCyan
                Write-Host "$sep`n" -ForegroundColor Cyan
            }
        }
        else {
            Write-Status 'No active secrets found — generating new secret.' -Type Warning
            $newSecret = New-AuditSecret -ApplicationObjectId $existingApp.Id -ExpiryMonths $SecretExpiryMonths
            Show-Credentials -AppId $existingApp.AppId -TenantId $tenantId -SecretText $newSecret.SecretText -SecretExpiry $newSecret.EndDateTime -PnPAppId $pnpInteractiveAppId
        }
    }
    else {
        # ----------------------------------------------------------
        # App does not exist — create it
        # ----------------------------------------------------------
        Write-Status "App not found — creating '$AppName'..." -Type Info

        Write-Status 'Resolving permission IDs...'
        $graphPerms    = @(Resolve-AppRoleIds -ResourceAppId $script:GraphResourceAppId    -PermissionNames $script:GraphPermissions)
        $exchangePerms = @(Resolve-AppRoleIds -ResourceAppId $script:ExchangeResourceAppId -PermissionNames $script:ExchangePermissions)

        # Graph SDK v2 requires camelCase keys and explicit string GUIDs in hashtables
        $resourceAccess = @(
            @{
                resourceAppId  = $script:GraphResourceAppId
                resourceAccess = @($graphPerms    | ForEach-Object { @{ id = $_.Id.ToString(); type = 'Role' } })
            },
            @{
                resourceAppId  = $script:ExchangeResourceAppId
                resourceAccess = @($exchangePerms | ForEach-Object { @{ id = $_.Id.ToString(); type = 'Role' } })
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

            Write-Status 'Assigning Exchange Administrator role to service principal...'
            Set-ExchangeAdminRole -ServicePrincipalId $ourSp.Id
            Write-Status 'Admin consent granted.' -Type Success

            Request-AdminConsent -ApplicationId $newApp.AppId -TenantName $orgName

            # Register the dedicated PnP interactive auth app for SharePoint
            $pnpInteractiveAppId = Register-PnPInteractiveApp -Tenant $tenantId -AppName "$AppName (SharePoint)"

            Write-Status "Generating $SecretExpiryMonths-month client secret..."
            $newSecret = New-AuditSecret -ApplicationObjectId $newApp.Id -ExpiryMonths $SecretExpiryMonths

            Show-Credentials -AppId $newApp.AppId -TenantId $tenantId -SecretText $newSecret.SecretText -SecretExpiry $newSecret.EndDateTime -PnPAppId $pnpInteractiveAppId
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
