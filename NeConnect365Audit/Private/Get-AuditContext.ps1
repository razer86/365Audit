function Get-AuditContext {
    <#
    .SYNOPSIS
        Returns the current module-scoped audit context.
    .DESCRIPTION
        Returns the hashtable set by Set-AuditContext containing per-tenant
        credentials, output paths, and Hudu configuration. Throws if no
        context has been set (indicates a function was called outside of
        an Invoke-TenantAudit run).
    #>
    [CmdletBinding()]
    param (
        [switch]$NoThrow
    )

    if (-not $script:AuditContext) {
        if ($NoThrow) { return $null }
        throw "No audit context set. Call Invoke-TenantAudit or Set-AuditContext first."
    }

    return $script:AuditContext
}
