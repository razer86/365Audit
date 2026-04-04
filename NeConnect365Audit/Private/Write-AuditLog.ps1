function Write-AuditLog {
    <#
    .SYNOPSIS
        Writes a timestamped line to the batch audit log file.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [string]$LogFile
    )

    if (-not $LogFile) { return }

    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')  $Message"
    Add-Content -Path $LogFile -Value $line -Encoding UTF8
}
