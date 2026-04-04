function Sync-AuditCustomers {
    <#
    .SYNOPSIS
        Syncs the customer list from Hudu by querying all audit toolkit assets.

    .DESCRIPTION
        Queries Hudu for all companies that have an audit toolkit asset (matching
        the specified layout ID), resolves their company slugs, and returns the
        list as PSCustomObject entries with HuduCompanySlug, HuduCompanyName,
        and Modules properties.

    .PARAMETER HuduBaseUrl
        Hudu instance base URL.

    .PARAMETER HuduApiKey
        Hudu API key.

    .PARAMETER HuduAssetLayoutId
        Asset layout ID for the audit toolkit credential assets. Default: 67.

    .PARAMETER DefaultModules
        Module list assigned to discovered customers. Default: @('A') (All).

    .PARAMETER PassThru
        Return the customer list as output. Without this, the function only
        writes status to the host.

    .EXAMPLE
        $customers = Sync-AuditCustomers -HuduBaseUrl $url -HuduApiKey $key -PassThru
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$HuduBaseUrl,

        [Parameter(Mandatory)]
        [string]$HuduApiKey,

        [int]$HuduAssetLayoutId = 67,

        [ValidateSet('1', '2', '3', '4', '5', '6', '7', 'A')]
        [string[]]$DefaultModules = @('A'),

        [switch]$PassThru
    )

    Write-Host "Querying Hudu for assets with layout ID $HuduAssetLayoutId..." -ForegroundColor Cyan

    # Fetch all assets for the layout
    $allAssets = Invoke-HuduRequest -Endpoint "api/v1/assets?asset_layout_id=$HuduAssetLayoutId" `
        -HuduBaseUrl $HuduBaseUrl -HuduApiKey $HuduApiKey -Paginate

    Write-Host "Found $($allAssets.Count) asset(s) across all companies." -ForegroundColor Cyan

    if ($allAssets.Count -eq 0) {
        Write-Warning "No audit toolkit assets found in Hudu."
        if ($PassThru) { return @() }
        return
    }

    # Resolve company slug for each unique company_id
    Write-Host "Resolving company slugs..." -ForegroundColor DarkCyan
    $companyMap = @{}
    foreach ($asset in $allAssets) {
        $cid = $asset.company_id
        if (-not $cid -or $companyMap.ContainsKey($cid)) { continue }
        try {
            $company = Invoke-HuduRequest -Endpoint "api/v1/companies/$cid" `
                -HuduBaseUrl $HuduBaseUrl -HuduApiKey $HuduApiKey
            $_companyData = if ($company.PSObject.Properties.Name -contains 'company') { $company.company } else { $company }
            $companyMap[$cid] = @{
                Slug = $_companyData.slug
                Name = $_companyData.name
            }
        }
        catch { Write-Warning "Could not resolve company ID $cid — skipping: $($_.Exception.Message)" }
    }

    Write-Host "Resolved $($companyMap.Count) company slug(s)." -ForegroundColor Green

    # Build customer list
    $customerList = foreach ($cid in $companyMap.Keys) {
        $info = $companyMap[$cid]
        if (-not $info.Slug) { continue }

        Write-Host "    + $($info.Name) ($($info.Slug))" -ForegroundColor DarkGray

        [PSCustomObject]@{
            HuduCompanySlug = $info.Slug
            HuduCompanyName = $info.Name
            Modules         = $DefaultModules
        }
    }

    $customerList = @($customerList | Sort-Object HuduCompanyName)
    Write-Host "  $($customerList.Count) customer(s) ready for audit." -ForegroundColor Green

    if ($PassThru) {
        return $customerList
    }
}
