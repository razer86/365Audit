<#
.SYNOPSIS
    Downloads the latest Microsoft licensing reference CSV to Resources\SkuFriendlyNames.csv.

.DESCRIPTION
    Fetches the official product names and service plan identifiers CSV from Microsoft
    and saves it to Resources\SkuFriendlyNames.csv. The audit scripts load this CSV
    at runtime to resolve SkuPartNumber values to friendly display names.

    Source: https://learn.microsoft.com/en-us/entra/identity/users/licensing-service-plan-reference

.EXAMPLE
    .\Helpers\Update-SkuFriendlyNames.ps1

.NOTES
    Author  : Raymond Slater
    Version : 1.0.0
#>

#Requires -Version 7.2

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

$csvUrl  = 'https://download.microsoft.com/download/e/3/e/e3e9faf2-f28b-490a-9ada-c6089a1fc5b0/Product%20names%20and%20service%20plan%20identifiers%20for%20licensing.csv'
$csvPath = Join-Path (Split-Path $PSScriptRoot -Parent) 'Resources\SkuFriendlyNames.csv'

# Count existing entries for comparison
$existingCount = 0
if (Test-Path $csvPath) {
    $existingCount = @(Import-Csv $csvPath | Select-Object -Property String_Id -Unique).Count
}

Write-Host "Downloading Microsoft licensing reference CSV..." -ForegroundColor Cyan
Invoke-WebRequest -Uri $csvUrl -OutFile $csvPath -ErrorAction Stop

$newEntries = Import-Csv $csvPath
$uniqueSkus = @($newEntries | Select-Object -Property String_Id -Unique)

Write-Host "  Saved: $csvPath" -ForegroundColor Green
Write-Host "  SKUs:  $($uniqueSkus.Count) unique (was $existingCount)" -ForegroundColor Green
