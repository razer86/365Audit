function Initialize-AuditOutput {
    <#
    .SYNOPSIS
        Creates the per-customer output folder structure and updates the audit context.
    .DESCRIPTION
        Queries the tenant org name from Graph, creates the output folder
        (<OrgName>_<yyyyMMdd>) with a Raw/ subdirectory, and updates the
        module-scoped audit context with the paths.
    #>
    [CmdletBinding()]
    param(
        [string]$OutputRoot
    )

    $ctx = Get-AuditContext

    # If already initialised (cached), return the context
    if ($ctx.OutputPath -and (Test-Path $ctx.OutputPath)) {
        return $ctx
    }

    # Resolve org name from Graph
    Import-GraphModule -ModuleName 'Microsoft.Graph.Identity.DirectoryManagement'
    $org = Get-MgOrganization -ErrorAction Stop | Select-Object -First 1

    if (-not $org) {
        throw "Unable to query Microsoft Graph organization details."
    }

    $orgName = ($org.DisplayName -replace '[^\w\s-]', '' -replace '\s+', '').Trim()
    if (-not $orgName) { $orgName = 'UnknownOrg' }

    # Resolve output root
    if (-not $OutputRoot) {
        $OutputRoot = $env:TEMP ?? $env:TMPDIR ?? '/tmp'
    }

    $folderName = "${orgName}_$(Get-Date -Format 'yyyyMMdd')"
    $outputPath = Join-Path $OutputRoot $folderName
    $rawPath    = Join-Path $outputPath 'Raw'

    New-Item -ItemType Directory -Path $rawPath -Force -ErrorAction Stop | Out-Null

    # Update context with resolved paths
    $ctx.OrgName       = $org.DisplayName
    $ctx.OutputPath    = $outputPath
    $ctx.RawOutputPath = $rawPath

    return $ctx
}
