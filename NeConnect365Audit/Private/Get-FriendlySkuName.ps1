function Get-FriendlySkuName {
    <#
    .SYNOPSIS
        Resolves a Microsoft SKU PartNumber to a human-readable display name.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$SkuPartNumber
    )

    # Lazy-load the SKU map on first call
    if (-not $script:SkuFriendlyMap) {
        $csvPath = Join-Path $PSScriptRoot '..' 'Resources' 'SkuFriendlyNames.csv'
        if (Test-Path $csvPath) {
            $script:SkuFriendlyMap = @{}
            Import-Csv $csvPath | ForEach-Object {
                $script:SkuFriendlyMap[$_.String_Id] = $_.Product_Display_Name
            }
        }
        else {
            $script:SkuFriendlyMap = @{}
            Write-Verbose "SkuFriendlyNames.csv not found at $csvPath"
        }
    }

    if ($script:SkuFriendlyMap.ContainsKey($SkuPartNumber)) {
        return $script:SkuFriendlyMap[$SkuPartNumber]
    }

    return $SkuPartNumber
}
