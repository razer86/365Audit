function Invoke-HuduRequest {
    <#
    .SYNOPSIS
        Makes authenticated requests to the Hudu API with pagination support.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Endpoint,

        [ValidateSet('GET', 'POST', 'PUT', 'PATCH', 'DELETE')]
        [string]$Method = 'GET',

        [hashtable]$Body,

        [string]$HuduBaseUrl,

        [string]$HuduApiKey,

        [switch]$Paginate
    )

    # Resolve Hudu connection from params or context
    if (-not $HuduBaseUrl -or -not $HuduApiKey) {
        $ctx = Get-AuditContext -NoThrow
        if ($ctx) {
            if (-not $HuduBaseUrl) { $HuduBaseUrl = $ctx.HuduBaseUrl }
            if (-not $HuduApiKey)  { $HuduApiKey  = $ctx.HuduApiKey }
        }
    }

    if (-not $HuduBaseUrl -or -not $HuduApiKey) {
        throw "Hudu base URL and API key are required. Pass -HuduBaseUrl and -HuduApiKey or set via Set-AuditContext."
    }

    $headers = @{
        'x-api-key'    = $HuduApiKey
        'Content-Type' = 'application/json'
    }

    $uri = "$($HuduBaseUrl.TrimEnd('/'))/$($Endpoint.TrimStart('/'))"

    if (-not $Paginate) {
        $params = @{
            Uri     = $uri
            Method  = $Method
            Headers = $headers
            ErrorAction = 'Stop'
        }
        if ($Body -and $Method -ne 'GET') {
            $params['Body'] = ($Body | ConvertTo-Json -Depth 10)
        }
        return Invoke-RestMethod @params
    }

    # Paginated GET
    $allResults = @()
    $page = 1
    do {
        $separator = if ($uri -match '\?') { '&' } else { '?' }
        $pagedUri = "${uri}${separator}page=$page&page_size=100"
        $response = Invoke-RestMethod -Uri $pagedUri -Method GET -Headers $headers -ErrorAction Stop

        # Hudu API returns different shapes depending on endpoint
        $items = if ($response -is [array]) { $response }
                 elseif ($response.PSObject.Properties.Name -contains 'assets')    { $response.assets }
                 elseif ($response.PSObject.Properties.Name -contains 'companies') { $response.companies }
                 else { @($response) }

        if ($items.Count -eq 0) { break }
        $allResults += $items
        $page++
    } while ($items.Count -ge 100)

    return $allResults
}
