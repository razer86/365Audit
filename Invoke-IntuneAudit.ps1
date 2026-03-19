<#
.SYNOPSIS
    Performs a security-focused audit of Microsoft Intune / Endpoint Manager.

.DESCRIPTION
    Connects to Microsoft Graph and collects Intune device management data including:
    - Licence check (skips gracefully if no qualifying Intune SKU is found)
    - Managed device inventory with compliance state and ownership
    - Per-device compliance policy states
    - Compliance policies with platform, assignments, grace period, and full settings
    - Configuration profiles with platform, type, and assignment scope
    - Managed application install summary
    - Autopilot device identities
    - Enrollment restrictions

    Output CSVs:
    - Intune_LicenceCheck.csv
    - Intune_Devices.csv
    - Intune_DeviceComplianceStates.csv
    - Intune_CompliancePolicies.csv
    - Intune_CompliancePolicySettings.csv
    - Intune_ConfigProfiles.csv
    - Intune_ConfigProfileSettings.csv
    - Intune_Apps.csv
    - Intune_AutopilotDevices.csv
    - Intune_EnrollmentRestrictions.csv

.NOTES
    Author      : Raymond Slater
    Version     : 1.7.0
    Change Log  : See CHANGELOG.md

.LINK
    https://github.com/razer86/365Audit
#>

#Requires -Version 7.2

param (
    [string]$AuditFolder,
    [switch]$DevMode = $false
)

if (-not $DevMode -and $MyInvocation.InvocationName -eq $MyInvocation.MyCommand.Name) {
    Write-Error "This script must be run from the 365Audit launcher. Use -DevMode for development." -ErrorAction Stop
}

$ScriptVersion = "1.7.0"
Write-Verbose "Invoke-IntuneAudit.ps1 loaded (v$ScriptVersion)"

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# === Ensure helper functions are loaded ===
if (-not (Get-Command Connect-MgGraphSecure -ErrorAction SilentlyContinue)) {
    Write-Error "Connect-MgGraphSecure is not loaded. Please run from the 365Audit launcher."
    exit 1
}
if (-not (Get-Command Initialize-AuditOutput -ErrorAction SilentlyContinue)) {
    Write-Error "Initialize-AuditOutput is not loaded. Please run from the 365Audit launcher."
    exit 1
}

# === Initialise output folder ===
try {
    $context   = Initialize-AuditOutput
    $outputDir = $context.RawOutputPath
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
}
catch {
    Write-Error "Failed to initialise audit output directory: $_"
    exit 1
}

# === Connect to Microsoft Graph and load Intune-specific sub-modules ===
try {
    Connect-MgGraphSecure
    Import-GraphSubModules @(
        'Microsoft.Graph.DeviceManagement',
        'Microsoft.Graph.Devices.CorporateManagement',
        'Microsoft.Graph.DeviceManagement.Enrollment'
    )
}
catch {
    Write-Error "Microsoft Graph connection failed: $_"
    exit 1
}

Write-Host "`nStarting Intune Audit for $($context.OrgName)..." -ForegroundColor Cyan

$step       = 0
$totalSteps = 8
$activity   = "Intune Audit — $($context.OrgName)"

# Known SKU part numbers that include Intune
$_intuneSkus = @(
    'INTUNE_A',
    'EMS_S_1',
    'EMS_S_3',
    'EMS_S_5',
    'ENTERPRISEPREMIUM',
    'ENTERPRISEPACK',
    'SPB',
    'BUSINESS_PREMIUM',
    'M365_F1',
    'M365_F3'
)

# Metadata keys to skip when iterating AdditionalProperties on compliance policies
$_odataSkipKeys = @(
    '@odata.type', 'id', 'createdDateTime', 'lastModifiedDateTime',
    'displayName', 'description', 'version', 'roleScopeTagIds', 'scheduledActionsForRule'
)


function Get-IntuneGraphErrorMessage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [System.Management.Automation.ErrorRecord]$ErrorRecord,

        [Parameter(Mandatory)]
        [string]$Operation
    )

    $message = $ErrorRecord.Exception.Message

    if ($message -match 'Application is not authorized to perform this operation\. Application must have one of the following scopes:\s*(?<scopes>.+?)\s*-\s*Operation ID') {
        $requiredPermissions = ($matches.scopes -split ',\s*' | Where-Object { $_ }) -join ', '
        return "$Operation requires Intune application permission(s): $requiredPermissions. Grant admin consent to the 365Audit app registration and rerun."
    }

    if ($message -match "Invalid filter clause: Could not find a property named 'isAssigned'") {
        return "$Operation failed because the Microsoft Graph mobile app endpoint does not support filtering on 'isAssigned'."
    }

    return "$Operation failed: $message"
}


function Convert-IntuneValueToString {
    [CmdletBinding()]
    param(
        $Value
    )

    if ($null -eq $Value) {
        return ''
    }

    if ($Value -is [string] -or $Value -is [ValueType]) {
        return [string]$Value
    }

    try {
        return ($Value | ConvertTo-Json -Depth 20 -Compress)
    }
    catch {
        return [string]$Value
    }
}


function Resolve-IntuneIdentifierText {
    [CmdletBinding()]
    param(
        $Identifier
    )

    if ($null -eq $Identifier) {
        return $null
    }

    $text = $null
    if ($Identifier -is [string]) {
        $text = $Identifier
    }
    elseif ($Identifier -is [System.Collections.IEnumerable]) {
        $parts = @(
            $Identifier |
            ForEach-Object {
                if ($null -ne $_) {
                    [string]$_
                }
            } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        )

        if ($parts.Count -eq 0) {
            return $null
        }

        $multiCharParts = @($parts | Where-Object { $_.Length -gt 1 })
        if ($parts.Count -gt 1 -and $multiCharParts.Count -eq 0) {
            $text = $parts -join ''
        }
        else {
            $text = $parts[0]
        }
    }
    else {
        $text = [string]$Identifier
    }

    $text = $text.Trim()
    if ($text -match '^(?:[A-Za-z0-9](?:[\s,;:_/-]+|$)){4,}$') {
        $text = (($text -split '[\s,;:_/-]+' | Where-Object { $_ }) -join '')
    }

    return $text
}


function Convert-IntuneIdentifierToLabel {
    [CmdletBinding()]
    param(
        $Identifier
    )

    $label = Resolve-IntuneIdentifierText -Identifier $Identifier
    if ([string]::IsNullOrWhiteSpace($label)) {
        return 'Setting'
    }

    $label = $label -replace '#microsoft.graph\.', ''
    $label = $label -replace '^.*~', ''
    $label = $label -creplace '(?<=[a-z])(?=[A-Z])', '_'
    $label = $label -creplace '(?<=[A-Z])(?=[A-Z][a-z])', '_'

    foreach ($prefix in @(
            'device_vendor_msft_policy_config_',
            'user_vendor_msft_policy_config_',
            'device_vendor_msft_policy_',
            'user_vendor_msft_policy_',
            'device_vendor_msft_laps_policies_',
            'user_vendor_msft_laps_policies_',
            'device_vendor_msft_',
            'user_vendor_msft_',
            'vendor_msft_',
            'device_',
            'user_'
        )) {
        if ($label.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)) {
            $label = $label.Substring($prefix.Length)
            break
        }
    }

    foreach ($noiseToken in @('policy_config_', 'policy_', 'config_', 'policies_', 'admx_')) {
        if ($label.StartsWith($noiseToken, [System.StringComparison]::OrdinalIgnoreCase)) {
            $label = $label.Substring($noiseToken.Length)
        }
    }

    $compoundReplacements = [ordered]@{
        'onedrivengscv2'                  = 'one drive'
        'onedrivengsc'                    = 'one drive'
        'automountteamsites'              = 'auto mount team sites'
        'filesondemandenabled'            = 'files on demand enabled'
        'silentaccountconfig'             = 'silent account config'
        'backupdirectory'                 = 'backup directory'
        'passwordagedays'                 = 'password age days'
        'passwordcomplexity'              = 'password complexity'
        'passwordlength'                  = 'password length'
        'postauthenticationactions'       = 'post authentication actions'
        'postauthenticationresetdelay'    = 'post authentication reset delay'
        'controleventlogbehavior'         = 'control event log behavior'
        'specifymaximumfilesize'          = 'specify maximum file size'
        'applicationlog'                  = 'application log'
        'securitylog'                     = 'security log'
        'systemlog'                       = 'system log'
        'channel_logmaxsize'              = 'channel log max size'
        'auditcredentialvalidation'       = 'audit credential validation'
        'auditaccountlockout'             = 'audit account lockout'
        'auditgroupmembership'            = 'audit group membership'
        'auditapplicationgroupmanagement' = 'audit application group management'
        'auditauthenticationpolicychange' = 'audit authentication policy change'
        'auditauthorizationpolicychange'  = 'audit authorization policy change'
        'auditlogoff'                     = 'audit logoff'
        'auditlogon'                      = 'audit logon'
        'policychange'                    = 'policy change'
        'accountlogonlogoff'              = 'account logon logoff'
        'accountlogon'                    = 'account logon'
        'accountmanagement'               = 'account management'
        'eventlogservice'                 = 'event log service'
        'logmaxsize'                      = 'log max size'
        'retention'                       = 'retention'
        'deliveryoptimizationmode'        = 'delivery optimization mode'
        'prereleasefeatures'              = 'pre release features'
        'automaticupdatemode'             = 'automatic update mode'
        'microsoftupdateserviceallowed'   = 'microsoft update service allowed'
        'driversexcluded'                 = 'drivers excluded'
        'networkproxyapplysettingsdevicewide' = 'network proxy apply settings device wide'
        'networkproxydisableautodetect'   = 'network proxy disable auto detect'
        'bluetoothallowedservices'        = 'bluetooth allowed services'
        'bluetoothblockadvertising'       = 'bluetooth block advertising'
        'bluetoothblockdiscoverablemode'  = 'bluetooth block discoverable mode'
        'bluetoothblockprepairing'        = 'bluetooth block pre pairing'
        'homepagelocation'                = 'homepage location'
        'newtabpagelocation'              = 'new tab page location'
        'downloadrestrictions'            = 'download restrictions'
        'showrecommendationsenabled'      = 'show recommendations enabled'
        'internetexplorerintegrationtestingallowed' = 'internet explorer integration testing allowed'
        'internetexplorerintegrationlocalfileallowed' = 'internet explorer integration local file allowed'
        'internetexplorerintegrationlocalmhtfileallowed' = 'internet explorer integration local MHT file allowed'
        'msawebsitessousingthisprofileallowed' = 'MSA websites SSO using this profile allowed'
        'aadwebsitessousingthisprofileenabled' = 'Entra ID websites SSO using this profile enabled'
        'userfeedbackallowed'             = 'user feedback allowed'
        'allowgamesmenu'                  = 'allow games menu'
        'outlookhubmenuenabled'           = 'Outlook hub menu enabled'
        'enhancesecuritymodeallowuserbypass' = 'enhance security mode allow user bypass'
        'familysafetysettingsenabled'     = 'family safety settings enabled'
        'sitesafetyservicesenabled'       = 'site safety services enabled'
        'clickonceenabled'                = 'ClickOnce enabled'
        'directinvokeenabled'             = 'direct invoke enabled'
        'autoimportatfirstrun'            = 'auto import at first run'
        'bingadssuppression'              = 'Bing ads suppression'
        'browsersignin'                   = 'browser sign in'
        'managedfavorites'                = 'managed favorites'
        'managedsearchengines'            = 'managed search engines'
        'nonremovableprofileenabled'      = 'non removable profile enabled'
        'cryptowalletenabled'             = 'crypto wallet enabled'
        'edgeedropenabled'                = 'Edge Drop enabled'
        'favoritesbarenabled'             = 'favorites bar enabled'
        'forcesync'                       = 'force sync'
        'newpdfreaderenabled'             = 'new PDF reader enabled'
        'edgeshoppingassistantenabled'    = 'Edge shopping assistant enabled'
        'hubssidebarenabled'              = 'hubs sidebar enabled'
        'showmicrosoftrewards'            = 'show Microsoft Rewards'
        'walletdonationenabled'           = 'wallet donation enabled'
        'blockexternalextensions'         = 'block external extensions'
        'gamermodeenabled'                = 'gamer mode enabled'
        'passwordmanagerenabled'          = 'password manager enabled'
        'printingenabled'                 = 'printing enabled'
        'edgeblockautofill'               = 'edge block autofill'
        'edgeblocked'                     = 'edge blocked'
        'edgecookiepolicy'                = 'edge cookie policy'
        'edgeblockdevelopertools'         = 'edge block developer tools'
        'edgeblocksendingdonottrackheader' = 'edge block sending do not track header'
        'edgeblockextensions'             = 'edge block extensions'
        'edgeblockinprivatebrowsing'      = 'edge block in private browsing'
        'edgeblockjavascript'             = 'edge block JavaScript'
        'edgeblockpasswordmanager'        = 'edge block password manager'
        'edgeblockaddressbardropdown'     = 'edge block address bar dropdown'
        'edgeblockcompatibilitylist'      = 'edge block compatibility list'
    }

    foreach ($replacement in $compoundReplacements.GetEnumerator()) {
        $label = $label -replace [regex]::Escape($replacement.Key), $replacement.Value
    }

    $tokens = @(
        $label -split '[_\-/\.\s]+' |
        Where-Object {
            $_ -and $_ -notin @('device', 'user', 'vendor', 'msft', 'policy', 'config', 'policies', 'setting', 'instance')
        }
    )

    if ($tokens.Count -eq 0) {
        return 'Setting'
    }

    $textInfo = [System.Globalization.CultureInfo]::InvariantCulture.TextInfo
    $finalTokens = [System.Collections.Generic.List[string]]::new()
    foreach ($token in $tokens) {
        $lower = $token.ToLowerInvariant()
        switch ($lower) {
            'aad' { [void]$finalTokens.Add('Entra ID'); continue }
            'admx' { [void]$finalTokens.Add('ADMX'); continue }
            'dns' { [void]$finalTokens.Add('DNS'); continue }
            'gpo' { [void]$finalTokens.Add('GPO'); continue }
            'id' { [void]$finalTokens.Add('ID'); continue }
            'laps' { [void]$finalTokens.Add('LAPS'); continue }
            default { [void]$finalTokens.Add($textInfo.ToTitleCase($lower)) }
        }
    }

    $label = [string]::Join(' ', $finalTokens).Trim()
    $label = $label -replace '\bOne Drive\b', 'OneDrive'
    $label = Normalize-IntuneDisplayLabel -Label $label
    return $label
}


function Normalize-IntuneDisplayLabel {
    [CmdletBinding()]
    param(
        [string]$Label
    )

    if ([string]::IsNullOrWhiteSpace($Label)) {
        return $null
    }

    $tokens = [System.Collections.Generic.List[string]]::new()
    foreach ($token in @($Label -split '\s+' | Where-Object { $_ })) {
        if ($tokens.Count -gt 0 -and $tokens[$tokens.Count - 1].Equals([string]$token, [System.StringComparison]::OrdinalIgnoreCase)) {
            continue
        }
        [void]$tokens.Add([string]$token)
    }

    return [string]::Join(' ', $tokens).Trim()
}


function Get-IntuneChildDisplayLabel {
    [CmdletBinding()]
    param(
        [string]$ParentLabel,
        [string]$ChildLabel
    )

    $normalizedParent = Normalize-IntuneDisplayLabel -Label $ParentLabel
    $normalizedChild = Normalize-IntuneDisplayLabel -Label $ChildLabel

    if ([string]::IsNullOrWhiteSpace($normalizedChild)) {
        return $null
    }

    if ([string]::IsNullOrWhiteSpace($normalizedParent)) {
        return $normalizedChild
    }

    if ($normalizedChild.Equals($normalizedParent, [System.StringComparison]::OrdinalIgnoreCase)) {
        return $null
    }

    if ($normalizedChild.StartsWith($normalizedParent + ' ', [System.StringComparison]::OrdinalIgnoreCase)) {
        $remainder = Normalize-IntuneDisplayLabel -Label $normalizedChild.Substring($normalizedParent.Length).Trim()
        if ([string]::IsNullOrWhiteSpace($remainder)) {
            return $null
        }

        $parentTokens = @($normalizedParent -split '\s+' | Where-Object { $_ })
        $remainderTokens = @($remainder -split '\s+' | Where-Object { $_ })
        if ($remainderTokens.Count -gt 0 -and $remainderTokens.Count -le $parentTokens.Count) {
            $parentTail = [string]::Join(' ', $parentTokens[($parentTokens.Count - $remainderTokens.Count)..($parentTokens.Count - 1)])
            if ($remainder.Equals($parentTail, [System.StringComparison]::OrdinalIgnoreCase)) {
                return $null
            }
        }

        return $remainder
    }

    return $normalizedChild
}


function Convert-IntuneConfigurationSettingValueToString {
    [CmdletBinding()]
    param(
        $SettingValue,

        [int]$Depth = 0
    )

    if ($Depth -gt 8 -or $null -eq $SettingValue) {
        return ''
    }

    if ($SettingValue -is [string] -or $SettingValue -is [ValueType]) {
        return [string]$SettingValue
    }

    $propertyNames = @($SettingValue.PSObject.Properties.Name)
    $currentLabel = $null
    if ($propertyNames -contains 'settingDefinitionId' -and $SettingValue.settingDefinitionId) {
        $currentLabel = Convert-IntuneIdentifierToLabel -Identifier $SettingValue.settingDefinitionId
    }

    if ($propertyNames -contains 'simpleSettingValue') {
        return Convert-IntuneConfigurationSettingValueToString -SettingValue $SettingValue.simpleSettingValue -Depth ($Depth + 1)
    }

    if ($propertyNames -contains 'choiceSettingValue') {
        $parts = [System.Collections.Generic.List[string]]::new()
        $childParts = [System.Collections.Generic.List[string]]::new()
        $choiceValue = [string]$SettingValue.choiceSettingValue.value
        if ($choiceValue) {
            $definitionId = [string]$SettingValue.settingDefinitionId
            if ($definitionId -and $choiceValue.StartsWith($definitionId + '_', [System.StringComparison]::OrdinalIgnoreCase)) {
                $choiceValue = $choiceValue.Substring($definitionId.Length + 1)
            }
        }

        foreach ($child in @($SettingValue.choiceSettingValue.children)) {
            $childLabel = Get-IntuneChildDisplayLabel -ParentLabel $currentLabel -ChildLabel (Convert-IntuneIdentifierToLabel -Identifier $child.settingDefinitionId)
            $childValue = Convert-IntuneConfigurationSettingValueToString -SettingValue $child -Depth ($Depth + 1)
            if ($childValue) {
                if ($childLabel) {
                    $childParts.Add(('{0}: {1}' -f $childLabel, $childValue))
                }
                else {
                    $childParts.Add($childValue)
                }
            }
        }

        if ($choiceValue) {
            if ($choiceValue -match '^\d+$') {
                if ($childParts.Count -eq 0) {
                    $parts.Add("Selected option: $choiceValue")
                }
            }
            else {
                $parts.Add((Convert-IntuneIdentifierToLabel -Identifier $choiceValue))
            }
        }

        foreach ($childPart in $childParts) {
            $parts.Add($childPart)
        }

        return ($parts -join '; ')
    }

    if ($propertyNames -contains 'groupSettingCollectionValue') {
        $groups = [System.Collections.Generic.List[string]]::new()
        foreach ($group in @($SettingValue.groupSettingCollectionValue)) {
            $groupParts = [System.Collections.Generic.List[string]]::new()
            foreach ($child in @($group.children)) {
                $childLabel = Get-IntuneChildDisplayLabel -ParentLabel $currentLabel -ChildLabel (Convert-IntuneIdentifierToLabel -Identifier $child.settingDefinitionId)
                $childValue = Convert-IntuneConfigurationSettingValueToString -SettingValue $child -Depth ($Depth + 1)
                if ($childValue) {
                    if ($childLabel) {
                        $groupParts.Add(('{0}: {1}' -f $childLabel, $childValue))
                    }
                    else {
                        $groupParts.Add($childValue)
                    }
                }
            }
            if ($groupParts.Count -gt 0) {
                $groups.Add(($groupParts -join '; '))
            }
        }
        return ($groups -join ' | ')
    }

    if ($propertyNames -contains 'children') {
        $childParts = [System.Collections.Generic.List[string]]::new()
        foreach ($child in @($SettingValue.children)) {
            $childLabel = Get-IntuneChildDisplayLabel -ParentLabel $currentLabel -ChildLabel (Convert-IntuneIdentifierToLabel -Identifier $child.settingDefinitionId)
            $childValue = Convert-IntuneConfigurationSettingValueToString -SettingValue $child -Depth ($Depth + 1)
            if ($childValue) {
                if ($childLabel) {
                    $childParts.Add(('{0}: {1}' -f $childLabel, $childValue))
                }
                else {
                    $childParts.Add($childValue)
                }
            }
        }
        return ($childParts -join '; ')
    }

    if ($propertyNames -contains 'value') {
        return Convert-IntuneValueToString -Value $SettingValue.value
    }

    return Convert-IntuneValueToString -Value $SettingValue
}


function Convert-IntunePlatformToken {
    [CmdletBinding()]
    param(
        [string]$Token
    )

    switch -Wildcard ($Token) {
        'windows*'        { return 'Windows 10/11' }
        'ios*'            { return 'iOS' }
        'androidForWork*' { return 'Android Enterprise' }
        'android*'        { return 'Android' }
        'mac*'            { return 'macOS' }
        default           { return $Token }
    }
}


function Convert-IntunePlatformListToString {
    [CmdletBinding()]
    param(
        $Platforms
    )

    $tokens = @($Platforms | Where-Object { $_ })
    if ($tokens.Count -eq 0) {
        return 'Unknown'
    }

    return ($tokens | ForEach-Object { Convert-IntunePlatformToken -Token ([string]$_) } | Select-Object -Unique) -join ', '
}


$script:IntuneGroupDisplayNameCache = @{}
function Resolve-IntuneGroupDisplayName {
    [CmdletBinding()]
    param(
        [string]$GroupId
    )

    if ([string]::IsNullOrWhiteSpace($GroupId)) {
        return $null
    }

    if ($script:IntuneGroupDisplayNameCache.ContainsKey($GroupId)) {
        return $script:IntuneGroupDisplayNameCache[$GroupId]
    }

    $displayName = $GroupId

    try {
        $group = Get-MgGroup -GroupId $GroupId -Property DisplayName -ErrorAction Stop
        if ($group.DisplayName) {
            $displayName = $group.DisplayName
        }
    }
    catch {
        Write-Verbose "Could not resolve Intune assignment group '$GroupId': $_"
    }

    $script:IntuneGroupDisplayNameCache[$GroupId] = $displayName
    return $displayName
}


function Get-IntuneAssignmentLabel {
    [CmdletBinding()]
    param(
        $Assignment
    )

    $target = if ($Assignment.PSObject.Properties.Name -contains 'Target') { $Assignment.Target } else { $Assignment.target }
    if ($null -eq $target) {
        return $null
    }

    $odataType = $null
    if ($target.PSObject.Properties.Name -contains 'AdditionalProperties' -and $target.AdditionalProperties) {
        $odataType = $target.AdditionalProperties.'@odata.type'
    }
    if (-not $odataType -and $target.PSObject.Properties.Name -contains '@odata.type') {
        $odataType = $target.'@odata.type'
    }

    $groupId = $null
    foreach ($name in @('GroupId', 'groupId')) {
        if ($target.PSObject.Properties.Name -contains $name -and $target.$name) {
            $groupId = $target.$name
            break
        }
    }
    if (-not $groupId -and $target.PSObject.Properties.Name -contains 'AdditionalProperties' -and $target.AdditionalProperties) {
        $groupId = $target.AdditionalProperties.groupId
    }

    switch -Wildcard ($odataType) {
        '*allDevicesAssignmentTarget'       { return 'All Devices' }
        '*allLicensedUsersAssignmentTarget' { return 'All Users' }
        '*exclusionGroupAssignmentTarget'   {
            if ($groupId) {
                return ("Exclude Group ({0})" -f (Resolve-IntuneGroupDisplayName -GroupId $groupId))
            }
            return 'Exclude Group'
        }
        '*groupAssignmentTarget'            {
            if ($groupId) {
                return ("Group ({0})" -f (Resolve-IntuneGroupDisplayName -GroupId $groupId))
            }
            return 'Group'
        }
        default {
            if ($odataType) {
                return ($odataType -replace '#microsoft.graph.', '')
            }
            return 'Target'
        }
    }
}


function Get-IntuneAssignmentSummary {
    [CmdletBinding()]
    param(
        [object[]]$Assignments
    )

    $labels = @($Assignments | ForEach-Object { Get-IntuneAssignmentLabel -Assignment $_ } | Where-Object { $_ } | Select-Object -Unique)
    if ($labels.Count -eq 0) {
        return 'None'
    }
    if ($labels -contains 'All Devices') {
        return 'All Devices'
    }
    if ($labels -contains 'All Users') {
        return 'All Users'
    }

    $groupLabels = @($labels | Where-Object { $_ -like '*Group*' })
    if ($groupLabels.Count -gt 0) {
        return "$($groupLabels.Count) group(s)"
    }

    return "$($labels.Count) target(s)"
}


function Invoke-GraphCollectionRequest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Uri
    )

    $items   = [System.Collections.Generic.List[object]]::new()
    $nextUri = $Uri

    while ($nextUri) {
        $response = $null
        $attempt  = 0
        while ($null -eq $response) {
            try {
                $response = Invoke-MgGraphRequest -Method GET -Uri $nextUri -OutputType PSObject -ErrorAction Stop
            }
            catch {
                $status = $null
                try { $status = [int]$_.Exception.Response.StatusCode } catch {}
                if ($status -eq 429 -and $attempt -lt 5) {
                    $retryAfter = $null
                    try { $retryAfter = [int]$_.Exception.Response.Headers['Retry-After'] } catch {}
                    if (-not $retryAfter -or $retryAfter -lt 1) { $retryAfter = [int][math]::Pow(2, $attempt + 2) }
                    Write-Warning "Graph throttled (429) — retrying in ${retryAfter}s (attempt $($attempt + 1)/5)..."
                    Start-Sleep -Seconds $retryAfter
                    $attempt++
                }
                else { throw }
            }
        }
        foreach ($item in @($response.value)) {
            $items.Add($item)
        }
        $nextUri = $response.'@odata.nextLink'
    }

    return @($items)
}


function Get-ConfigurationPolicyTypeName {
    [CmdletBinding()]
    param(
        $Policy
    )

    if ($Policy.templateReference.templateDisplayName) {
        return $Policy.templateReference.templateDisplayName
    }
    if ($Policy.templateReference.templateFamily) {
        return $Policy.templateReference.templateFamily
    }
    if ($Policy.technologies) {
        return (@($Policy.technologies) -join ', ')
    }

    return 'configurationPolicy'
}


function Get-ConfigurationPolicySettingName {
    [CmdletBinding()]
    param(
        $Setting
    )

    $names = @()

    foreach ($def in @($Setting.settingDefinitions)) {
        if ($def.displayName) { $names += $def.displayName }
        elseif ($def.name)    { $names += $def.name }
        elseif ($def.id)      { $names += $def.id }
    }

    if ($names.Count -eq 0 -and $Setting.settingInstance.settingDefinitionId) {
        $names += $Setting.settingInstance.settingDefinitionId
    }
    if ($names.Count -eq 0 -and $Setting.name) {
        $names += $Setting.name
    }
    if ($names.Count -eq 0 -and $Setting.id) {
        $names += $Setting.id
    }

    $names = @(
        $names |
        ForEach-Object { Resolve-IntuneIdentifierText -Identifier $_ } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        Select-Object -Unique
    )

    if ($names.Count -eq 0) {
        return 'Setting'
    }

    return (@($names) -join ', ')
}


function Invoke-IntuneExportReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ReportName,

        [string[]]$Select,

        [string]$Filter
    )

    $requestBody = [ordered]@{
        reportName       = $ReportName
        format           = 'csv'
        localizationType = 'ReplaceLocalizableValues'
    }

    if ($Select -and $Select.Count -gt 0) {
        $requestBody.select = @($Select)
    }
    if ($Filter) {
        $requestBody.filter = $Filter
    }

    $job = Invoke-MgGraphRequest -Method POST `
        -Uri 'https://graph.microsoft.com/beta/deviceManagement/reports/exportJobs' `
        -Body ($requestBody | ConvertTo-Json -Depth 10) `
        -ContentType 'application/json' `
        -OutputType PSObject `
        -ErrorAction Stop

    if (-not $job.id) {
        throw "The Intune export job for '$ReportName' did not return a job ID."
    }

    $jobStatus = $null
    $jobUri = "https://graph.microsoft.com/beta/deviceManagement/reports/exportJobs('{0}')" -f $job.id
    for ($attempt = 0; $attempt -lt 45; $attempt++) {
        Start-Sleep -Seconds 2
        $jobStatus = Invoke-MgGraphRequest -Method GET -Uri $jobUri -OutputType PSObject -ErrorAction Stop
        if ($jobStatus.status -in @('completed', 'complete')) {
            break
        }
        if ($jobStatus.status -in @('failed', 'error')) {
            throw "The Intune export job for '$ReportName' failed with status '$($jobStatus.status)'."
        }
    }

    if (-not $jobStatus -or -not $jobStatus.url) {
        throw "The Intune export job for '$ReportName' did not complete in time."
    }

    $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("365Audit-Report-" + [guid]::NewGuid().ToString())
    $zipPath = Join-Path $tempRoot 'report.zip'
    $extractDir = Join-Path $tempRoot 'extract'

    try {
        New-Item -ItemType Directory -Path $extractDir -Force | Out-Null
        Invoke-WebRequest -Uri $jobStatus.url -OutFile $zipPath -ErrorAction Stop | Out-Null
        Expand-Archive -LiteralPath $zipPath -DestinationPath $extractDir -Force

        $csvPath = Get-ChildItem -Path $extractDir -Recurse -Filter '*.csv' -ErrorAction Stop |
            Sort-Object FullName |
            Select-Object -First 1 -ExpandProperty FullName

        if (-not $csvPath) {
            throw "No CSV file was found in the Intune export for '$ReportName'."
        }

        return @(Import-Csv -Path $csvPath)
    }
    finally {
        if (Test-Path $tempRoot) {
            Remove-Item -Path $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}


function Get-IntuneMobileAppInstallSummary {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$MobileAppId,

        [switch]$HasAssignments
    )

    $summaryObject = $null

    try {
        $response = Invoke-MgGraphRequest -Method GET -Uri ("https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/{0}/installSummary" -f $MobileAppId) -OutputType PSObject -ErrorAction Stop
        $summaryObject = if ($response.PSObject.Properties.Name -contains 'value' -and $response.value) { $response.value } else { $response }
    }
    catch {
        Write-Verbose "Could not retrieve mobile app install summary for '$MobileAppId': $_"
    }

    $installedCount = 0
    $failedCount = 0
    $pendingCount = 0

    if ($summaryObject) {
        $installedCount = [int]($summaryObject.installedDeviceCount ?? 0)
        $failedCount    = [int]($summaryObject.failedDeviceCount ?? 0)
        $pendingCount   = [int]($summaryObject.pendingInstallDeviceCount ?? $summaryObject.pendingDeviceCount ?? 0)
    }

    if ($HasAssignments -and ($installedCount + $failedCount + $pendingCount -eq 0)) {
        try {
            $deviceStatuses = @(Invoke-GraphCollectionRequest -Uri ("https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/{0}/deviceStatuses?$top=999" -f $MobileAppId))
            if ($deviceStatuses.Count -gt 0) {
                foreach ($status in $deviceStatuses) {
                    $installState = [string]($status.installState ?? $status.mobileAppInstallStatusValue ?? '')
                    switch -Wildcard ($installState.ToLowerInvariant()) {
                        'installed*' { $installedCount++ }
                        'failed*'    { $failedCount++ }
                        'pending*'   { $pendingCount++ }
                    }
                }
            }
        }
        catch {
            Write-Verbose "Could not retrieve mobile app device statuses for '$MobileAppId': $_"
        }
    }

    return [PSCustomObject]@{
        InstalledDeviceCount = $installedCount
        FailedDeviceCount    = $failedCount
        PendingInstallCount  = $pendingCount
    }
}


# ================================
# ===   Step 1 — Licence Check  ===
# ================================
$step++
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Checking Intune licence..." -PercentComplete ([int]($step / $totalSteps * 100))

$_subscribedSkus = @(Get-MgSubscribedSku -All -ErrorAction SilentlyContinue)
$_intuneLicSkus  = @($_subscribedSkus | Where-Object { $_.SkuPartNumber -in $_intuneSkus })
$_hasIntune      = $_intuneLicSkus.Count -gt 0

[PSCustomObject]@{
    HasIntune   = $_hasIntune
    LicencedSKUs = if ($_hasIntune) { ($_intuneLicSkus.SkuPartNumber -join ', ') } else { '' }
} | Export-Csv -Path (Join-Path $outputDir 'Intune_LicenceCheck.csv') -NoTypeInformation -Encoding UTF8

if (-not $_hasIntune) {
    Write-Warning "No Intune-capable licence found for $($context.OrgName). Skipping Intune data collection."
    Write-Progress -Id 1 -Activity $activity -Completed
    # Note: cannot use `return` here — script is dot-sourced by the launcher and return only exits
    # the current statement. Setting a skip flag instead; all remaining steps check this flag.
    $_intuneSkip = $true
}
else {
    $_intuneSkip = $false
}

if (-not $_intuneSkip) {

Write-Host "  Intune licence confirmed: $($_intuneLicSkus.SkuPartNumber -join ', ')" -ForegroundColor Green


# =======================================
# ===   Step 2 — Managed Devices      ===
# =======================================
$step++
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Retrieving managed devices..." -PercentComplete ([int]($step / $totalSteps * 100))

$_allDevices = @()
try {
    $_allDevices = @(Get-MgDeviceManagementManagedDevice -All -ErrorAction Stop)
    $_deviceRows = foreach ($_dev in $_allDevices) {
        [PSCustomObject]@{
            DeviceName         = $_dev.DeviceName
            OS                 = $_dev.OperatingSystem
            OSVersion          = $_dev.OsVersion
            DeviceType         = $_dev.DeviceType
            OwnerType          = $_dev.ManagedDeviceOwnerType
            EnrolledDateTime   = $_dev.EnrolledDateTime
            LastSyncDateTime   = $_dev.LastSyncDateTime
            ComplianceState    = $_dev.ComplianceState
            AssignedUser       = $_dev.UserPrincipalName
            Manufacturer       = $_dev.Manufacturer
            Model              = $_dev.Model
            SerialNumber       = $_dev.SerialNumber
            ManagementAgent    = $_dev.ManagementAgent
            AzureADRegistered  = $_dev.AzureADRegistered
            JoinType           = $_dev.JoinType
        }
    }
    $_deviceRows | Export-Csv -Path (Join-Path $outputDir 'Intune_Devices.csv') -NoTypeInformation -Encoding UTF8
    Write-Host "  Devices: $($_allDevices.Count) managed device(s) found." -ForegroundColor Green
}
catch {
    Write-Warning (Get-IntuneGraphErrorMessage -ErrorRecord $_ -Operation 'Managed devices')
    @() | Export-Csv -Path (Join-Path $outputDir 'Intune_Devices.csv') -NoTypeInformation -Encoding UTF8
}


# ================================================
# ===   Step 3 — Per-Device Compliance States  ===
# ================================================
$step++
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Retrieving per-device compliance policy states..." -PercentComplete ([int]($step / $totalSteps * 100))

$_complianceStateRows = [System.Collections.Generic.List[object]]::new()
foreach ($_dev in $_allDevices) {
    try {
        $_states = @(Get-MgDeviceManagementManagedDeviceCompliancePolicyState `
            -ManagedDeviceId $_dev.Id -All -ErrorAction Stop)
        foreach ($_state in $_states) {
            $_complianceStateRows.Add([PSCustomObject]@{
                DeviceName            = $_dev.DeviceName
                PolicyName            = $_state.DisplayName
                State                 = $_state.State
                LastReportedDateTime  = $_state.LastReportedDateTime
            })
        }
    }
    catch {
        Write-Verbose "Could not retrieve compliance states for device '$($_dev.DeviceName)': $_"
    }
}
$_complianceStateRows | Export-Csv -Path (Join-Path $outputDir 'Intune_DeviceComplianceStates.csv') -NoTypeInformation -Encoding UTF8
Write-Host "  Compliance states: $($_complianceStateRows.Count) policy state record(s)." -ForegroundColor Green


# ==================================================
# ===   Step 4 — Compliance Policies + Settings  ===
# ==================================================
$step++
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Retrieving compliance policies..." -PercentComplete ([int]($step / $totalSteps * 100))

$_policyRows        = [System.Collections.Generic.List[object]]::new()
$_policySettingRows = [System.Collections.Generic.List[object]]::new()

try {
    $_policies = @(Get-MgDeviceManagementDeviceCompliancePolicy -All -ErrorAction Stop)
    foreach ($_pol in $_policies) {
        # Assignments
        $_assignments = @()
        try {
            $_assignments = @(Get-MgDeviceManagementDeviceCompliancePolicyAssignment `
                -DeviceCompliancePolicyId $_pol.Id -All -ErrorAction SilentlyContinue)
        } catch {}
        $_assignedTo = if ($_assignments.Count -eq 0) { 'None' }
                       elseif ($_assignments | Where-Object { $_.Target.AdditionalProperties.'@odata.type' -like '*allDevices*' }) { 'All Devices' }
                       elseif ($_assignments | Where-Object { $_.Target.AdditionalProperties.'@odata.type' -like '*allLicensedUsers*' }) { 'All Users' }
                       else { "$($_assignments.Count) group(s)" }

        # Grace period (hours) from scheduled action rules
        $_gracePeriod = 0
        try {
            $_scheduledActions = @(Get-MgDeviceManagementDeviceCompliancePolicyScheduledActionsForRule `
                -DeviceCompliancePolicyId $_pol.Id -All -ErrorAction SilentlyContinue)
            foreach ($_sa in $_scheduledActions) {
                foreach ($_cfg in @($_sa.ScheduledActionConfigurations)) {
                    if ($_cfg.ActionType -eq 'block' -and $_cfg.GracePeriodHours -gt $_gracePeriod) {
                        $_gracePeriod = $_cfg.GracePeriodHours
                    }
                }
            }
        } catch {}

        # Derive platform from OData type
        $_odataType = $_pol.AdditionalProperties.'@odata.type' ?? ''
        $_platform  = switch -Wildcard ($_odataType) {
            '*windows10*'  { 'Windows 10/11' }
            '*ios*'        { 'iOS' }
            '*android*'    { 'Android' }
            '*macOS*'      { 'macOS' }
            default        { $_odataType -replace '#microsoft.graph.', '' }
        }

        $_assignmentDetails = (@($_assignments | ForEach-Object { Get-IntuneAssignmentLabel -Assignment $_ } | Where-Object { $_ } | Select-Object -Unique) -join '; ')

        $_policyRows.Add([PSCustomObject]@{
            PolicyId              = $_pol.Id
            PolicyName            = $_pol.DisplayName
            Description           = $_pol.Description
            Platform              = $_platform
            PolicyType            = ($_odataType -replace '#microsoft.graph.', '')
            AssignedTo            = $_assignedTo
            AssignmentDetails     = $_assignmentDetails
            GracePeriodHours      = $_gracePeriod
            CreatedDateTime       = $_pol.CreatedDateTime
            LastModifiedDateTime  = $_pol.LastModifiedDateTime
        })

        # Settings — iterate AdditionalProperties, skip metadata keys
        foreach ($_kv in $_pol.AdditionalProperties.GetEnumerator()) {
            if ($_kv.Key -in $_odataSkipKeys) { continue }
            $_policySettingRows.Add([PSCustomObject]@{
                PolicyId      = $_pol.Id
                PolicyName    = $_pol.DisplayName
                Platform      = $_platform
                SettingName   = $_kv.Key
                SettingValue  = (Convert-IntuneValueToString -Value $_kv.Value)
            })
        }
    }

    $_policyRows        | Export-Csv -Path (Join-Path $outputDir 'Intune_CompliancePolicies.csv')       -NoTypeInformation -Encoding UTF8
    $_policySettingRows | Export-Csv -Path (Join-Path $outputDir 'Intune_CompliancePolicySettings.csv') -NoTypeInformation -Encoding UTF8
    Write-Host "  Compliance policies: $($_policies.Count) policy(ies), $($_policySettingRows.Count) setting record(s)." -ForegroundColor Green
}
catch {
    Write-Warning (Get-IntuneGraphErrorMessage -ErrorRecord $_ -Operation 'Compliance policies')
    @() | Export-Csv -Path (Join-Path $outputDir 'Intune_CompliancePolicies.csv')       -NoTypeInformation -Encoding UTF8
    @() | Export-Csv -Path (Join-Path $outputDir 'Intune_CompliancePolicySettings.csv') -NoTypeInformation -Encoding UTF8
}


# =================================================
# ===   Step 5 — Configuration Profiles          ===
# =================================================
$step++
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Retrieving configuration profiles..." -PercentComplete ([int]($step / $totalSteps * 100))

$_profileSettingRows = [System.Collections.Generic.List[object]]::new()
$_profileRows = [System.Collections.Generic.List[object]]::new()

try {
    $_configProfiles = @(Get-MgDeviceManagementDeviceConfiguration -All -ErrorAction Stop)
    foreach ($_prof in $_configProfiles) {
        $_profAssignments = @()
        try {
            $_profAssignments = @(Get-MgDeviceManagementDeviceConfigurationAssignment `
                -DeviceConfigurationId $_prof.Id -All -ErrorAction SilentlyContinue)
        } catch {}
        $_profAssignedTo = Get-IntuneAssignmentSummary -Assignments $_profAssignments

        $_profOdata    = $_prof.AdditionalProperties.'@odata.type' ?? ''
        $_profPlatform = switch -Wildcard ($_profOdata) {
            '*windows10*'  { 'Windows 10/11' }
            '*ios*'        { 'iOS' }
            '*android*'    { 'Android' }
            '*macOS*'      { 'macOS' }
            default        { $_profOdata -replace '#microsoft.graph.', '' }
        }

        # Settings — iterate AdditionalProperties, skip metadata keys
        foreach ($_kv in $_prof.AdditionalProperties.GetEnumerator()) {
            if ($_kv.Key -in $_odataSkipKeys) { continue }
            $_profileSettingRows.Add([PSCustomObject]@{
                ProfileId     = $_prof.Id
                ProfileName   = $_prof.DisplayName
                Platform      = $_profPlatform
                ProfileType   = ($_profOdata -replace '#microsoft.graph.', '')
                SettingName   = (Convert-IntuneIdentifierToLabel -Identifier $_kv.Key)
                SettingValue  = (Convert-IntuneValueToString -Value $_kv.Value)
            })
        }

        $_assignmentDetails = (@($_profAssignments | ForEach-Object { Get-IntuneAssignmentLabel -Assignment $_ } | Where-Object { $_ } | Select-Object -Unique) -join '; ')

        $_profileRows.Add([PSCustomObject]@{
            ProfileId             = $_prof.Id
            ProfileName           = $_prof.DisplayName
            Description           = $_prof.Description
            Platform              = $_profPlatform
            ProfileType           = ($_profOdata -replace '#microsoft.graph.', '')
            Source                = 'deviceConfiguration'
            LastModifiedDateTime  = $_prof.LastModifiedDateTime
            AssignedTo            = $_profAssignedTo
            AssignmentDetails     = $_assignmentDetails
        })
    }

    try {
        $_modernProfiles = @(Invoke-GraphCollectionRequest -Uri 'https://graph.microsoft.com/beta/deviceManagement/configurationPolicies?$top=200')
        foreach ($_modernProf in $_modernProfiles) {
            $_modernAssignments = @()
            $_modernSettings = @()

            try {
                $_modernAssignments = @(Invoke-GraphCollectionRequest -Uri ("https://graph.microsoft.com/beta/deviceManagement/configurationPolicies/{0}/assignments?$top=200" -f $_modernProf.id))
            } catch {}
            try {
                $_modernSettings = @(Invoke-GraphCollectionRequest -Uri ("https://graph.microsoft.com/beta/deviceManagement/configurationPolicies/{0}/settings?$top=1000" -f $_modernProf.id))
            } catch {}

            $_modernAssignedTo = Get-IntuneAssignmentSummary -Assignments $_modernAssignments
            $_modernPlatform = Convert-IntunePlatformListToString -Platforms $_modernProf.platforms
            $_modernType = Get-ConfigurationPolicyTypeName -Policy $_modernProf

            $_assignmentDetails = (@($_modernAssignments | ForEach-Object { Get-IntuneAssignmentLabel -Assignment $_ } | Where-Object { $_ } | Select-Object -Unique) -join '; ')

            $_profileRows.Add([PSCustomObject]@{
                ProfileId             = $_modernProf.id
                ProfileName           = $_modernProf.name
                Description           = $_modernProf.description
                Platform              = $_modernPlatform
                ProfileType           = $_modernType
                Source                = 'configurationPolicy'
                LastModifiedDateTime  = $_modernProf.lastModifiedDateTime
                AssignedTo            = $_modernAssignedTo
                AssignmentDetails     = $_assignmentDetails
            })

            foreach ($_setting in $_modernSettings) {
                $_rawSettingName = Get-ConfigurationPolicySettingName -Setting $_setting
                $_profileSettingRows.Add([PSCustomObject]@{
                    ProfileId     = $_modernProf.id
                    ProfileName   = $_modernProf.name
                    Platform      = $_modernPlatform
                    ProfileType   = $_modernType
                    SettingName   = (Convert-IntuneIdentifierToLabel -Identifier $_rawSettingName)
                    SettingValue  = (Convert-IntuneConfigurationSettingValueToString -SettingValue $_setting.settingInstance)
                })
            }
        }
    }
    catch {
        Write-Warning (Get-IntuneGraphErrorMessage -ErrorRecord $_ -Operation 'Modern configuration policies')
    }

    $_profileRows        | Export-Csv -Path (Join-Path $outputDir 'Intune_ConfigProfiles.csv')        -NoTypeInformation -Encoding UTF8
    $_profileSettingRows | Export-Csv -Path (Join-Path $outputDir 'Intune_ConfigProfileSettings.csv') -NoTypeInformation -Encoding UTF8
    Write-Host "  Configuration profiles: $($_profileRows.Count) profile(s), $($_profileSettingRows.Count) setting record(s)." -ForegroundColor Green
}
catch {
    Write-Warning (Get-IntuneGraphErrorMessage -ErrorRecord $_ -Operation 'Configuration profiles')
    @() | Export-Csv -Path (Join-Path $outputDir 'Intune_ConfigProfiles.csv')        -NoTypeInformation -Encoding UTF8
    @() | Export-Csv -Path (Join-Path $outputDir 'Intune_ConfigProfileSettings.csv') -NoTypeInformation -Encoding UTF8
}


# =============================================
# ===   Step 6 — Apps + Install Summary     ===
# =============================================
$step++
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Retrieving assigned apps..." -PercentComplete ([int]($step / $totalSteps * 100))

try {
    $_apps = @(Get-MgDeviceAppManagementMobileApp -All -ErrorAction Stop)
    $_appSummaryMap = @{}

    try {
        $_appSummaryRows = @(Invoke-IntuneExportReport -ReportName 'AppInstallStatusAggregate' -Select @(
                'ApplicationId',
                'InstalledDeviceCount',
                'FailedDeviceCount',
                'PendingInstallDeviceCount'
            ))
        Write-Verbose "Retrieved Intune app install aggregate report with $($_appSummaryRows.Count) row(s)."
        foreach ($_summaryRow in $_appSummaryRows) {
            $appId = [string]$_summaryRow.ApplicationId
            if ($appId) {
                $_appSummaryMap[$appId.ToLowerInvariant()] = $_summaryRow
            }
        }
    }
    catch {
        Write-Verbose "Could not retrieve aggregate Intune app install report: $_"
    }

    $_appRows = foreach ($_app in $_apps) {
        $_assignments = @()
        try {
            $_assignments = @(Invoke-GraphCollectionRequest -Uri ("https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/{0}/assignments?$top=200" -f $_app.Id))
        } catch {}
        $_summary = $null
        $_summaryKey = ([string]$_app.Id).ToLowerInvariant()
        if ($_appSummaryMap.ContainsKey($_summaryKey)) {
            $_reportSummary = $_appSummaryMap[$_summaryKey]
            $_summary = [PSCustomObject]@{
                InstalledDeviceCount = [int]($_reportSummary.InstalledDeviceCount ?? 0)
                FailedDeviceCount    = [int]($_reportSummary.FailedDeviceCount ?? 0)
                PendingInstallCount  = [int]($_reportSummary.PendingInstallDeviceCount ?? 0)
            }
        }
        else {
            Write-Verbose "No aggregate report row found for app '$($_app.DisplayName)' ($($_app.Id)); falling back to per-app endpoints."
            $_summary = Get-IntuneMobileAppInstallSummary -MobileAppId $_app.Id -HasAssignments:($_assignments.Count -gt 0)
        }

        $_assignmentDetails = (@($_assignments | ForEach-Object { Get-IntuneAssignmentLabel -Assignment $_ } | Where-Object { $_ } | Select-Object -Unique) -join '; ')

        [PSCustomObject]@{
            AppId                 = $_app.Id
            AppName              = $_app.DisplayName
            AppType              = ($_app.AdditionalProperties.'@odata.type' -replace '#microsoft.graph.', '')
            Publisher            = $_app.Publisher
            Description          = $_app.Description
            CreatedDateTime      = $_app.CreatedDateTime
            LastModifiedDateTime = $_app.LastModifiedDateTime
            AssignedTo           = (Get-IntuneAssignmentSummary -Assignments $_assignments)
            AssignmentDetails    = $_assignmentDetails
            InstalledDeviceCount = $_summary.InstalledDeviceCount
            FailedDeviceCount    = $_summary.FailedDeviceCount
            PendingInstallCount  = $_summary.PendingInstallCount
        }
    }
    $_appRows | Export-Csv -Path (Join-Path $outputDir 'Intune_Apps.csv') -NoTypeInformation -Encoding UTF8
    Write-Host "  Apps: $($_apps.Count) app(s)." -ForegroundColor Green
}
catch {
    Write-Warning (Get-IntuneGraphErrorMessage -ErrorRecord $_ -Operation 'Apps')
    @() | Export-Csv -Path (Join-Path $outputDir 'Intune_Apps.csv') -NoTypeInformation -Encoding UTF8
}


# =============================================
# ===   Step 7 — Autopilot Devices          ===
# =============================================
$step++
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Retrieving Autopilot devices..." -PercentComplete ([int]($step / $totalSteps * 100))

try {
    $_autopilot = @(Get-MgDeviceManagementWindowsAutopilotDeviceIdentity -All -ErrorAction Stop)
    $_autopilotRows = foreach ($_ap in $_autopilot) {
        [PSCustomObject]@{
            SerialNumber          = $_ap.SerialNumber
            Model                 = $_ap.Model
            Manufacturer          = $_ap.Manufacturer
            GroupTag              = $_ap.GroupTag
            AssignedUser          = $_ap.UserPrincipalName
            EnrollmentState       = $_ap.EnrollmentState
            LastContactedDateTime = $_ap.LastContactedDateTime
        }
    }
    $_autopilotRows | Export-Csv -Path (Join-Path $outputDir 'Intune_AutopilotDevices.csv') -NoTypeInformation -Encoding UTF8
    Write-Host "  Autopilot: $($_autopilot.Count) device(s) registered." -ForegroundColor Green
}
catch {
    Write-Warning (Get-IntuneGraphErrorMessage -ErrorRecord $_ -Operation 'Autopilot devices')
    @() | Export-Csv -Path (Join-Path $outputDir 'Intune_AutopilotDevices.csv') -NoTypeInformation -Encoding UTF8
}


# ===================================================
# ===   Step 8 — Enrollment Restrictions           ===
# ===================================================
$step++
Write-Progress -Id 1 -Activity $activity -Status "Step $step/$totalSteps — Retrieving enrollment restrictions..." -PercentComplete ([int]($step / $totalSteps * 100))

try {
    $_restrictions = @(Get-MgDeviceManagementDeviceEnrollmentConfiguration -All -ErrorAction Stop)
    $_restrictionRows = foreach ($_res in $_restrictions) {
        $_resAssignments = @()
        try {
            $_resAssignments = @(Get-MgDeviceManagementDeviceEnrollmentConfigurationAssignment `
                -DeviceEnrollmentConfigurationId $_res.Id -All -ErrorAction SilentlyContinue)
        } catch {}
        $_resAssignedTo = if ($_resAssignments.Count -eq 0) { 'None' }
                          elseif ($_resAssignments | Where-Object { $_.Target.AdditionalProperties.'@odata.type' -like '*allDevices*' }) { 'All Devices' }
                          elseif ($_resAssignments | Where-Object { $_.Target.AdditionalProperties.'@odata.type' -like '*allLicensedUsers*' }) { 'All Users' }
                          else { "$($_resAssignments.Count) group(s)" }

        $_ap = $_res.AdditionalProperties
        $_blockPersonal  = if ($null -ne $_ap.platformRestrictions) {
            $platforms = @('android', 'androidForWork', 'ios', 'mac', 'windows')
            $blocked = $false
            foreach ($p in $platforms) {
                if ($_ap.platformRestrictions.$p.personalDeviceEnrollmentBlocked -eq $true) { $blocked = $true; break }
            }
            $blocked
        } else { $false }

        $_maxDevices = if ($null -ne $_ap.limit) { $_ap.limit } else { 'N/A' }

        [PSCustomObject]@{
            ConfigName           = $_res.DisplayName
            Platform             = ($_res.AdditionalProperties.'@odata.type' -replace '#microsoft.graph.', '')
            BlockPersonalDevices = $_blockPersonal
            MaxDevicesPerUser    = $_maxDevices
            Priority             = $_res.Priority
            AssignedTo           = $_resAssignedTo
        }
    }
    $_restrictionRows | Export-Csv -Path (Join-Path $outputDir 'Intune_EnrollmentRestrictions.csv') -NoTypeInformation -Encoding UTF8
    Write-Host "  Enrollment restrictions: $($_restrictions.Count) configuration(s)." -ForegroundColor Green
}
catch {
    Write-Warning (Get-IntuneGraphErrorMessage -ErrorRecord $_ -Operation 'Enrollment restrictions')
    @() | Export-Csv -Path (Join-Path $outputDir 'Intune_EnrollmentRestrictions.csv') -NoTypeInformation -Encoding UTF8
}


} # end if (-not $_intuneSkip)

Write-Progress -Id 1 -Activity $activity -Completed
Write-Host "`nIntune Audit complete. Output saved to: $outputDir" -ForegroundColor Green
