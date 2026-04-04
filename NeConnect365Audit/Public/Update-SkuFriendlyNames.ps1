function Update-SkuFriendlyNames {
    <#
    .SYNOPSIS
        Downloads the latest Microsoft SKU friendly names CSV.
    .DESCRIPTION
        Fetches Microsoft's official licensing reference CSV from learn.microsoft.com
        and saves it to the module's Resources directory. The CSV maps SKU PartNumbers
        to human-readable product display names for use in audit reports.
    #>
    [CmdletBinding()]
    param()

    $csvUrl  = 'https://download.microsoft.com/download/e/3/e/e3e9faf2-f28b-490a-9ada-c6089a1fc5b0/Product%20names%20and%20service%20plan%20identifiers%20for%20licensing.csv'
    $csvPath = Join-Path $PSScriptRoot '..' 'Resources' 'SkuFriendlyNames.csv'
    $resDir  = Split-Path $csvPath -Parent

    if (-not (Test-Path $resDir)) {
        New-Item -ItemType Directory -Path $resDir -Force | Out-Null
    }

    $beforeCount = 0
    if (Test-Path $csvPath) {
        $beforeCount = @(Import-Csv $csvPath | Select-Object -ExpandProperty String_Id -Unique).Count
    }

    Write-Host "Downloading SKU friendly names from Microsoft..." -ForegroundColor Cyan
    try {
        Invoke-WebRequest -Uri $csvUrl -OutFile $csvPath -TimeoutSec 30 -ErrorAction Stop
    }
    catch {
        Write-Error "Failed to download SKU CSV: $($_.Exception.Message)"
        return
    }

    $afterCount = @(Import-Csv $csvPath | Select-Object -ExpandProperty String_Id -Unique).Count
    Write-Host "  Updated: $afterCount unique SKUs (was $beforeCount)." -ForegroundColor Green

    # Clear the in-memory cache so the next Get-FriendlySkuName call reloads
    $script:SkuFriendlyMap = $null
}
