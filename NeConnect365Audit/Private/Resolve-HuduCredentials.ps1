function Resolve-HuduCredentials {
    <#
    .SYNOPSIS
        Fetches app registration credentials from a Hudu asset for a customer.
    .DESCRIPTION
        Queries Hudu for the customer's audit toolkit asset, extracts the
        Application ID, Tenant ID, Cert Base64, and Cert Password fields,
        then decodes the certificate to a temp .pfx file.

        Returns a hashtable with AppId, TenantId, CertFilePath, CertPassword.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$CompanySlug,

        [Parameter(Mandatory)]
        [string]$HuduBaseUrl,

        [Parameter(Mandatory)]
        [string]$HuduApiKey,

        [int]$AssetLayoutId = 67
    )

    # Resolve company from slug
    $company = Invoke-HuduRequest -Endpoint "api/v1/companies?slug=$CompanySlug" `
        -HuduBaseUrl $HuduBaseUrl -HuduApiKey $HuduApiKey

    $_companyData = if ($company.PSObject.Properties.Name -contains 'companies') {
        $company.companies | Select-Object -First 1
    } else {
        $company | Select-Object -First 1
    }

    if (-not $_companyData) {
        throw "Hudu company not found for slug '$CompanySlug'."
    }

    $companyId   = $_companyData.id
    $companyName = $_companyData.name

    # Find the audit toolkit asset
    $assets = Invoke-HuduRequest -Endpoint "api/v1/assets?company_id=$companyId&asset_layout_id=$AssetLayoutId" `
        -HuduBaseUrl $HuduBaseUrl -HuduApiKey $HuduApiKey -Paginate

    $asset = $assets | Select-Object -First 1
    if (-not $asset) {
        throw "No audit toolkit asset (layout $AssetLayoutId) found for company '$companyName' (ID: $companyId)."
    }

    # Extract credential fields
    $fields = @{}
    foreach ($f in $asset.fields) {
        $fields[$f.label] = $f.value
    }

    $appId    = $fields['Application ID']
    $tenantId = $fields['Tenant ID']
    $certB64  = $fields['Cert Base64']
    $certPwd  = $fields['Cert Password']

    if (-not $appId -or -not $tenantId -or -not $certB64 -or -not $certPwd) {
        $_missing = @()
        if (-not $appId)    { $_missing += 'Application ID' }
        if (-not $tenantId) { $_missing += 'Tenant ID' }
        if (-not $certB64)  { $_missing += 'Cert Base64' }
        if (-not $certPwd)  { $_missing += 'Cert Password' }
        throw "Hudu asset '$($asset.name)' is missing fields: $($_missing -join ', ')"
    }

    # Decode certificate to temp .pfx
    $certBytes = [Convert]::FromBase64String($certB64)
    $_tempDir  = $env:TEMP ?? $env:TMPDIR ?? '/tmp'
    $certPath  = Join-Path $_tempDir "365Audit-$(New-Guid).pfx"
    [System.IO.File]::WriteAllBytes($certPath, $certBytes)

    $securePassword = ConvertTo-SecureString $certPwd -AsPlainText -Force

    return @{
        AppId        = $appId
        TenantId     = $tenantId
        CertFilePath = $certPath
        CertPassword = $securePassword
        CompanyName  = $companyName
        CompanyId    = $companyId
        AssetName    = $asset.name
    }
}
