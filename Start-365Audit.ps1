<#
.SYNOPSIS
    Interactive script launcher for Microsoft 365 audit modules (Azure, Exchange, SharePoint, etc.).

.DESCRIPTION
    Allows the user to interactively select and run one or more Microsoft 365 audit modules.
    Modules are run from local disk if available, or downloaded from GitHub if missing.
    The launcher also manages audit output folders and passes the folder name to the summary module.

.NOTES
    Author      : Raymond Slater
    Version     : 1.0.1
    Change Log  :
        1.0.0 - Initial release
        1.0.1 - Updated to standardize comments and pass folder to summary

.LINK
    https://github.com/razer86/365Audit
#>

# Force UTF-8 output for emoji if supported
if ($PSVersionTable.PSVersion.Major -ge 6) {
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
}

# Load common audit helper functions
$commonPath = Join-Path $PSScriptRoot "common\Audit-Common.ps1"
if (Test-Path $commonPath) {
    . $commonPath
} else {
    Write-Error "Required helper script not found: $commonPath"
    exit 1
}

# === Config ===
$debugMode = $true
$localPath = "$PSScriptRoot"
$remoteBaseUrl = "https://raw.githubusercontent.com/razer86/365Audit/refs/heads/main"

function Connect-MgGraphSecure {
    # Required scopes for all modules
    $requiredScopes = @(
        "User.Read.All",                               # Required to get basic user profile info
        "Directory.Read.All",                          # Required for admin role assignments, user properties, group membership, etc.
        "Reports.Read.All",                            # Required for sign-in activity and MFA reports
        "Organization.Read.All",                       # Used to retrieve tenant display name and info
        "Policy.Read.All",                             # Used for password policy and security defaults
        "Policy.ReadWrite.ConditionalAccess",          # Used to read conditional access policies
        "UserAuthenticationMethod.Read.All",           # Required for MFA method details per user
        "RoleManagement.Read.Directory",               # Used for Partner Relationship checking (DAP/GDAP)
        "Group.Read.All"                               # Required for Microsoft 365 group properties, privacy, members, Teams, and dynamic membership
    )    

    # Skip if already connected
    if (Get-MgContext) {
        write-host "ℹ  Already connected to Microsoft Graph." -ForegroundColor Green
        return
    }

    try {
        write-host "ℹ  Connecting to Microsoft Graph..." -ForegroundColor Blue
        Connect-MgGraph -Scopes $requiredScopes -ErrorAction Stop
    }
    catch {
        write-error "❌ Failed to connect to Microsoft Graph: $_"
        throw
    }

    # Verify all scopes were granted
    $context = Get-MgContext
    $missingScopes = $requiredScopes | Where-Object { $_ -notin $context.Scopes }

    if ($missingScopes.Count -gt 0) {
        write-error "❌ Missing required Microsoft Graph scopes:"
        $missingScopes | ForEach-Object { Write-Warning " - $_" }

        throw "Please re-run and ensure consent is granted for the missing scopes."
    }

    write-host "✅ Connected with all required Microsoft Graph permissions." -ForegroundColor Green
}

# === Initialize-AuditOutput ===
function Initialize-AuditOutput {
    if ($script:AuditOutputContext) {
        return $script:AuditOutputContext
    }

    if (-not (Get-MgContext)) {
        Connect-MgGraphSecure
    }

    $org = Get-MgOrganization
    $companyName = $org.DisplayName -replace '[^a-zA-Z0-9]', '_'
    $folderName = "${companyName}_$(Get-Date -Format 'yyyyMMdd')"
    $outputDir = Join-Path $PSScriptRoot "..\$folderName"
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null

    $script:AuditOutputContext = @{
        OrgName    = $companyName
        FolderName = $folderName
        OutputPath = $outputDir
    }

    return $script:AuditOutputContext
}

# === Define Menu Items ===
$menu = @{
    1 = @{ Name = "Microsoft Entra Audit";          Script = "Invoke-EntraAudit.ps1" }
    2 = @{ Name = "Exchange Online Audit";          Script = "Invoke-ExchangeAudit.ps1" }
    3 = @{ Name = "SharePoint Online Audit";        Script = "Invoke-SharePointAudit.ps1" }
    #4 = @{ Name = "Microsoft Teams Audit";         Script = "Invoke-TeamsAudit.ps1" }
    8 = @{ Name = "Generate Audit Summary";         Script = "Generate-AuditSummary.ps1" }
    9 = @{ Name = "Run All Modules (1,2,3,8)";      Script = @("Invoke-AzureAudit.ps1", "Invoke-ExchangeAudit.ps1", "Invoke-SharePointAudit.ps1") }
    0 = @{ Name = "Exit";                           Script = $null }
}

# === Display Menu ===
Write-Host "`n╔════════════════════════════════════╗"
Write-Host "║    Microsoft 365 Audit Launcher    ║"
Write-Host "╚════════════════════════════════════╝"

foreach ($key in ($menu.Keys | Sort-Object {[int]$_})) {
    Write-Host "$key. $($menu[$key].Name)"
}

# === User Selection ===
$selection = Read-Host "`nSelect one or more modules (comma separated, e.g. 1,2)"
if ($selection -eq "0") {
    write-host "`nℹ  Exiting script, Goodbye!"
    return
}

# === Parse Selection ===
$selectedIndexes = $selection -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ -match '^\d+$' } | ForEach-Object { [int]$_ }

# === Execute Selected Modules ===
foreach ($index in $selectedIndexes) {
    if ($menu.ContainsKey($index)) {
        $module = $menu[$index]

        if (-not $module.Script) { continue }

        $scriptsToRun = @($module.Script)

        foreach ($scriptName in $scriptsToRun) {
            $localScriptPath = Join-Path $localPath $scriptName
            $remoteScriptUrl = "$remoteBaseUrl/$scriptName"
            write-host "`nℹ  Attempting to load: $localScriptPath"

            Write-Host "`n■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■"
            Write-Host "           Starting: $scriptName" -ForegroundColor Cyan
            Write-Host "■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■"

            if ($scriptName -eq "Generate-AuditSummary.ps1") {
                if (Test-Path $outputDir) {
                    & $localScriptPath -AuditFolder $outputDir
                }
                else {
                    write-host "❌  Could not find any audit output folders."
                }
                continue
            }

            if (Test-Path $localScriptPath) {
                write-host "`nℹ  Found local script: $localScriptPath"
                write-host "`nℹ  This will open a new session window or prompt for login if required.`n"
                Start-Sleep -Seconds 2
                . $localScriptPath
            }
            else {
                write-host "ℹ  Local script not found. Downloading from GitHub..."
                write-host "ℹ  Fetching from: $remoteScriptUrl"
                write-host "ℹ  This will run interactively and may prompt for Microsoft 365 login.`n"
                Start-Sleep -Seconds 2
                try {
                    irm $remoteScriptUrl | iex
                }
                catch {
                    write-host "❌  Failed to download or run $scriptName from GitHub: $_"
                }
            }

            write-host "✅  Completed: $scriptName"
        }
    }
    else {
        write-host "❌  Invalid selection: $index"
    }
}
