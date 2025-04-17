<#
.SYNOPSIS
    Shared helper functions for Microsoft 365 Audit modules.

.DESCRIPTION
    Provides common functionality such as secure Microsoft Graph connection
    and standardized output folder initialization across all audit modules.

.NOTES
    Author      : Raymond Slater
    Version     : 1.0.0
    Change Log  :
        1.0.0 - Initial creation and migration of shared helpers from launcher.

.LINK
    https://github.com/razer86/365Audit
#>

# Force UTF-8 output for emoji if supported
if ($PSVersionTable.PSVersion.Major -ge 6) {
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
}

# ===============================================
# ===   Connect to Microsoft Graph Securely   ===
# ===============================================
function Connect-MgGraphSecure {
    $requiredScopes = @(
        "User.Read.All",
        "Directory.Read.All",
        "Reports.Read.All",
        "Organization.Read.All",
        "Policy.Read.All",
        "Policy.ReadWrite.ConditionalAccess",
        "UserAuthenticationMethod.Read.All",
        "RoleManagement.Read.Directory",
        "Group.Read.All"
    )

    try {
        Write-Host "🔐 Connecting to Microsoft Graph..."
        Connect-MgGraph -Scopes $requiredScopes -ErrorAction Stop
    }
    catch {
        Write-Error "❌ Failed to connect to Microsoft Graph: $_"
        throw
    }

    $context = Get-MgContext
    $missingScopes = $requiredScopes | Where-Object { $_ -notin $context.Scopes }

    if ($missingScopes.Count -gt 0) {
        Write-Warning "⚠ Missing required Microsoft Graph scopes:"
        $missingScopes | ForEach-Object { Write-Warning " - $_" }
        throw "Please re-run and ensure consent is granted for the missing scopes."
    }

    Write-Host "✅ Connected with all required Microsoft Graph permissions." -ForegroundColor Green
}

# ==========================================
# ===   Initialize Audit Output Folder   ===
# ==========================================
function Initialize-AuditOutput {
    if ($script:AuditOutputContext) {
        return $script:AuditOutputContext
    }

    if (-not (Get-MgContext)) {
        Connect-MgGraphSecure
    }

    # =========================================
# ===   Fetch and Save Organization Info   ===
# =========================================

    # === Retrieve all organizations ===
    $orgList = Get-MgOrganization

    # === Multi-tenant check ===
    if ($orgList.Count -gt 1) {
        $primaryDomain = $orgList[0].VerifiedDomains |
            Where-Object { $_.IsInitial -eq $true -and $_.Name -like "*.onmicrosoft.com" } |
            Select-Object -ExpandProperty Name

        Write-Warning "⚠️ Multiple organizations detected ($($orgList.Count))."
        Write-Host "➡️ Using first org: $($orgList[0].DisplayName) ($primaryDomain)"
    }

    # === Use the first org ===
    $org = $orgList | Select-Object -First 1
    $branding = Get-MgOrganizationBranding -OrganizationId $org.Id -ErrorAction SilentlyContinue

    # === Build expanded object ===
    $orgExpanded = [PSCustomObject]@{
        Id                          = $org.Id
        DisplayName                 = $org.DisplayName
        VerifiedDomains             = $org.VerifiedDomains
        TechnicalNotificationMails = $org.TechnicalNotificationMails
        MarketingNotificationEmails= $org.MarketingNotificationEmails
        DefaultDomain               = $org.DefaultDomain
        CountryLetterCode           = $org.CountryLetterCode
        PreferredLanguageTag        = $org.PreferredLanguageTag
        ProvisionedPlans            = $org.ProvisionedPlans
        AssignedPlans               = $org.AssignedPlans
        Branding                    = $branding
        Raw                         = $org
    }

    $cleanDisplayName = $orgExpanded.DisplayName -replace '[^a-zA-Z0-9]', ''
    $folderName = "${cleanDisplayName}_$(Get-Date -Format 'yyyyMMdd')"
    $outputDir = Join-Path $PSScriptRoot "..\$folderName"
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null

    # Save as JSON
    $orgExpanded | ConvertTo-Json -Depth 10 | Set-Content -Path (Join-Path $outputDir "OrgInfo.json") -Encoding UTF8

    $script:AuditOutputContext = @{
        OrgName    = $orgExpanded.DisplayName
        FolderName = $folderName
        OutputPath = $outputDir
    }

    return $script:AuditOutputContext
}