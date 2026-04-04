function Add-AuditIssue {
    <#
    .SYNOPSIS
        Logs a non-fatal audit collection issue to AuditIssues.csv.
    .DESCRIPTION
        Used by audit modules to record errors that don't stop the entire audit
        (e.g., a single API call failing). Issues are rendered in the Technical
        Issues section of the summary report.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateSet('Warning', 'Error')]
        [string]$Severity,

        [Parameter(Mandatory)]
        [string]$Section,

        [Parameter(Mandatory)]
        [string]$Collector,

        [Parameter(Mandatory)]
        [string]$Description,

        [string]$Action,

        [string]$DocUrl
    )

    $ctx = Get-AuditContext -NoThrow
    if (-not $ctx -or -not $ctx.RawOutputPath) { return }

    $issuesCsv = Join-Path $ctx.RawOutputPath 'AuditIssues.csv'

    [PSCustomObject]@{
        Timestamp   = Get-Date -Format 'o'
        Severity    = $Severity
        Section     = $Section
        Collector   = $Collector
        Description = $Description
        Action      = $Action
        DocUrl      = $DocUrl
    } | Export-Csv -Path $issuesCsv -NoTypeInformation -Encoding UTF8 -Append
}
