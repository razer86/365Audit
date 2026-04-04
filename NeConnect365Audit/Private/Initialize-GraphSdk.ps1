function Resolve-GraphModuleVersion {
    [CmdletBinding()]
    param()

    if ($script:GraphModuleVersion) {
        return $script:GraphModuleVersion
    }

    $versions = @(Get-Module -ListAvailable -Name 'Microsoft.Graph.Authentication' |
        Select-Object -ExpandProperty Version -Unique |
        Sort-Object -Descending)

    if (-not $versions) {
        throw "Microsoft.Graph.Authentication is not installed. Install it with: Install-Module Microsoft.Graph.Authentication"
    }

    $script:GraphModuleVersion = $versions | Select-Object -First 1
    Write-Verbose "Resolved Microsoft.Graph module version: $script:GraphModuleVersion"
    return $script:GraphModuleVersion
}


function ConvertTo-NormalizedVersion {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [version]$Version
    )

    $revision = if ($Version.Revision -lt 0) { 0 } else { $Version.Revision }
    return [version]::new($Version.Major, $Version.Minor, $Version.Build, $revision)
}


function Get-GraphModuleInfo {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ModuleName,

        [Parameter(Mandatory)]
        [version]$ModuleVersion
    )

    $normalizedModuleVersion = ConvertTo-NormalizedVersion -Version $ModuleVersion
    return Get-Module -ListAvailable -Name $ModuleName |
        Where-Object { (ConvertTo-NormalizedVersion -Version $_.Version) -eq $normalizedModuleVersion } |
        Select-Object -First 1
}


function Get-GraphDependencyAssemblyVersion {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ModuleName,

        [Parameter(Mandatory)]
        [version]$ModuleVersion,

        [Parameter(Mandatory)]
        [string]$AssemblyName
    )

    $moduleInfo = Get-GraphModuleInfo -ModuleName $ModuleName -ModuleVersion $ModuleVersion
    if (-not $moduleInfo) { return $null }

    $assemblyPath = Get-ChildItem -Path $moduleInfo.ModuleBase -Recurse -Filter "$AssemblyName.dll" -ErrorAction SilentlyContinue |
        Sort-Object FullName |
        Select-Object -First 1 -ExpandProperty FullName

    if (-not $assemblyPath) { return $null }
    return [Reflection.AssemblyName]::GetAssemblyName($assemblyPath).Version
}


function Get-GraphDependencyDirectories {
    [CmdletBinding()]
    param()

    if ($script:GraphDependencyDirectories) {
        return $script:GraphDependencyDirectories
    }

    $targetVersion = Resolve-GraphModuleVersion
    $authModule = Get-GraphModuleInfo -ModuleName 'Microsoft.Graph.Authentication' -ModuleVersion $targetVersion

    if (-not $authModule) {
        throw "Unable to locate Microsoft.Graph.Authentication $targetVersion."
    }

    $candidateDirs = @(
        (Join-Path $authModule.ModuleBase 'Dependencies\Core'),
        (Join-Path $authModule.ModuleBase 'Dependencies\Desktop'),
        (Join-Path $authModule.ModuleBase 'Dependencies'),
        $authModule.ModuleBase
    ) | Where-Object { Test-Path $_ }

    $script:GraphDependencyDirectories = $candidateDirs
    return $script:GraphDependencyDirectories
}


function Initialize-GraphDependencies {
    [CmdletBinding()]
    param()

    $assembliesToPrime = @(
        'Microsoft.Graph.Core',
        'Azure.Core',
        'Azure.Identity',
        'Microsoft.Kiota.Abstractions',
        'Microsoft.Kiota.Authentication.Azure',
        'Microsoft.Kiota.Http.HttpClientLibrary',
        'Microsoft.Kiota.Serialization.Json',
        'Microsoft.Kiota.Serialization.Form',
        'Microsoft.Kiota.Serialization.Text'
    )

    $alreadyLoaded = [AppDomain]::CurrentDomain.GetAssemblies() |
        Group-Object { $_.GetName().Name } -AsHashTable -AsString

    foreach ($assemblyName in $assembliesToPrime) {
        if ($alreadyLoaded.ContainsKey($assemblyName)) { continue }

        foreach ($dir in (Get-GraphDependencyDirectories)) {
            $candidatePath = Join-Path $dir ($assemblyName + '.dll')
            if (-not (Test-Path $candidatePath)) { continue }

            try {
                [System.Runtime.Loader.AssemblyLoadContext]::Default.LoadFromAssemblyPath($candidatePath) | Out-Null
                break
            }
            catch {
                Write-Verbose "Could not preload assembly '$assemblyName' from '$candidatePath': $_"
            }
        }
    }
}


function Import-GraphModule {
    <#
    .SYNOPSIS
        Imports a specific Microsoft Graph sub-module at the resolved version.
    .DESCRIPTION
        Ensures version consistency across all Graph sub-modules. If the required
        version is not installed, it is installed from PSGallery automatically.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ModuleName
    )

    $targetVersion = Resolve-GraphModuleVersion
    $normalizedTargetVersion = ConvertTo-NormalizedVersion -Version $targetVersion
    $targetModuleInfo = Get-GraphModuleInfo -ModuleName $ModuleName -ModuleVersion $targetVersion

    if (-not $targetModuleInfo) {
        Write-Host "  Required module '$ModuleName' not found — installing v$targetVersion..." -ForegroundColor Yellow
        Install-Module $ModuleName -RequiredVersion $targetVersion -Scope CurrentUser -Force -AllowClobber -SkipPublisherCheck -WarningAction SilentlyContinue -ErrorAction Stop
        $targetModuleInfo = Get-GraphModuleInfo -ModuleName $ModuleName -ModuleVersion $targetVersion
        if (-not $targetModuleInfo) {
            throw "Installation of '$ModuleName' v$targetVersion failed — module still not found after install."
        }
        Write-Host "  Installed '$ModuleName' v$targetVersion." -ForegroundColor Green
    }

    $loadedModules = @(Get-Module -Name $ModuleName -All)
    $mismatchedLoad = @($loadedModules | Where-Object {
            (ConvertTo-NormalizedVersion -Version $_.Version) -ne $normalizedTargetVersion
        })

    if ($mismatchedLoad) {
        $loadedSummary = $mismatchedLoad | Sort-Object Version -Descending |
            ForEach-Object { "{0} ({1})" -f $_.Name, $_.Version }
        throw (("Microsoft Graph module version mismatch for '{0}': {1}. " +
            "Start a new PowerShell session so NeConnect365Audit can load Microsoft.Graph {2} consistently.") -f
            $ModuleName, ($loadedSummary -join ', '), $targetVersion)
    }

    if (-not ($loadedModules | Where-Object {
                (ConvertTo-NormalizedVersion -Version $_.Version) -eq $normalizedTargetVersion
            })) {
        Import-Module -Name $targetModuleInfo.Path -ErrorAction Stop
    }
}


function Initialize-GraphSdk {
    <#
    .SYNOPSIS
        Preloads Microsoft Graph SDK assemblies and validates version consistency.
    .DESCRIPTION
        Called at module import time to prevent MSAL assembly conflicts between
        Microsoft.Graph and Az.Accounts modules sharing the same PowerShell session.
    #>
    [CmdletBinding()]
    param()

    $targetVersion = Resolve-GraphModuleVersion
    Initialize-GraphDependencies

    $expectedCore = Get-GraphDependencyAssemblyVersion -ModuleName 'Microsoft.Graph.Authentication' `
        -ModuleVersion $targetVersion -AssemblyName 'Microsoft.Graph.Core'
    $expectedMsal = Get-GraphDependencyAssemblyVersion -ModuleName 'Microsoft.Graph.Authentication' `
        -ModuleVersion $targetVersion -AssemblyName 'Microsoft.Identity.Client'

    $loadedCore = [AppDomain]::CurrentDomain.GetAssemblies() |
        Where-Object { $_.GetName().Name -eq 'Microsoft.Graph.Core' } | Select-Object -First 1
    $loadedMsal = [AppDomain]::CurrentDomain.GetAssemblies() |
        Where-Object { $_.GetName().Name -eq 'Microsoft.Identity.Client' } | Select-Object -First 1

    if ($loadedCore -and $expectedCore -and $loadedCore.GetName().Version -ne $expectedCore) {
        throw (("Microsoft.Graph.Core {0} is already loaded, but NeConnect365Audit requires {1}. " +
            "Start a new PowerShell session.") -f $loadedCore.GetName().Version, $expectedCore)
    }

    if ($loadedMsal -and $expectedMsal -and $loadedMsal.GetName().Version -ne $expectedMsal) {
        throw (("Microsoft.Identity.Client {0} is already loaded, but NeConnect365Audit requires {1}. " +
            "Another module likely loaded an incompatible MSAL version. Start a new PowerShell session.") -f
            $loadedMsal.GetName().Version, $expectedMsal)
    }

    Import-GraphModule -ModuleName 'Microsoft.Graph.Authentication'
}
