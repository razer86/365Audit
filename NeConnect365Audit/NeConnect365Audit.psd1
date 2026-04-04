@{
    RootModule        = 'NeConnect365Audit.psm1'
    ModuleVersion     = '1.0.0'
    GUID              = 'b3f7a2d1-8e4c-4f6b-9a1d-5c3e7f8b2d4a'
    Author            = 'Raymond Slater'
    CompanyName       = 'NeConnect'
    Copyright         = '(c) 2026 NeConnect. All rights reserved.'
    Description       = 'Microsoft 365 tenant security audit toolkit for MSPs. Connects to customer tenants via app-only cert auth, collects data across Entra, Exchange, SharePoint, Intune, Teams, and Maester CIS baselines, generates HTML summary reports, and publishes to Hudu.'
    PowerShellVersion = '7.2'

    RequiredModules   = @(
        @{ ModuleName = 'Microsoft.Graph.Authentication'; ModuleVersion = '2.0.0' }
    )

    FunctionsToExport = @(
        'Invoke-TenantAudit'
        'Register-AuditApp'
        'Publish-AuditReport'
        'Sync-AuditCustomers'
        'Remove-AuditCustomer'
        'Update-SkuFriendlyNames'
    )

    CmdletsToExport   = @()
    VariablesToExport  = @()
    AliasesToExport    = @()

    PrivateData = @{
        PSData = @{
            Tags         = @('Microsoft365', 'M365', 'Audit', 'Security', 'MSP', 'Hudu', 'CIS', 'Maester')
            LicenseUri   = 'https://github.com/razer86/365Audit/blob/main/LICENSE'
            ProjectUri   = 'https://github.com/razer86/365Audit'
            ReleaseNotes = 'Initial module release — rebuilt from script-based toolkit.'
        }
    }
}
