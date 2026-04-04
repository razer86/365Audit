function Set-AuditContext {
    <#
    .SYNOPSIS
        Stores per-tenant audit credentials and config in module scope.
    .DESCRIPTION
        Called by Invoke-TenantAudit before each customer run. All connection
        and audit module functions read from this context via Get-AuditContext.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$AppId,

        [Parameter(Mandatory)]
        [string]$TenantId,

        [string]$CertFilePath,

        [SecureString]$CertPassword,

        [string]$OutputPath,

        [string]$RawOutputPath,

        [string]$OrgName,

        [string]$HuduBaseUrl,

        [string]$HuduApiKey,

        [string]$HuduCompanySlug,

        [int]$HuduAssetLayoutId = 67,

        [int]$HuduReportLayoutId = 68,

        [string]$HuduReportAssetName = 'M365 - Monthly Audit Report',

        [string[]]$MspDomains = @()
    )

    $script:AuditContext = @{
        AppId               = $AppId
        TenantId            = $TenantId
        CertFilePath        = $CertFilePath
        CertPassword        = $CertPassword
        OutputPath          = $OutputPath
        RawOutputPath       = $RawOutputPath
        OrgName             = $OrgName
        HuduBaseUrl         = $HuduBaseUrl
        HuduApiKey          = $HuduApiKey
        HuduCompanySlug     = $HuduCompanySlug
        HuduAssetLayoutId   = $HuduAssetLayoutId
        HuduReportLayoutId  = $HuduReportLayoutId
        HuduReportAssetName = $HuduReportAssetName
        MspDomains          = $MspDomains
    }
}
