# ConditionalAccessCatalog.psm1
# Conditional Access catalog for the EOA web worker (app-only REST).
# Lists and locations are cached per session. User-scope writes PATCH Graph
# and clear the policy cache. User search uses $filter/$search, never Get-User -All.

$script:CaCatalog = @{
    Policies         = $null
    Locations        = $null
    Users            = [System.Collections.Generic.Dictionary[string, object]]::new([StringComparer]::OrdinalIgnoreCase)
    Groups           = [System.Collections.Generic.Dictionary[string, object]]::new([StringComparer]::OrdinalIgnoreCase)
    Roles            = [System.Collections.Generic.Dictionary[string, object]]::new([StringComparer]::OrdinalIgnoreCase)
    RoleGraphLoaded  = $false
}

function Clear-CaCatalogCache {
    $script:CaCatalog.Policies = $null
    $script:CaCatalog.Locations = $null
    $script:CaCatalog.Users.Clear()
    $script:CaCatalog.Groups.Clear()
    $script:CaCatalog.Roles.Clear()
    $script:CaCatalog.RoleGraphLoaded = $false
}

function Get-CaGraphProp {
    param($Object, [string[]]$Names)
    if ($null -eq $Object) { return $null }
    foreach ($n in $Names) {
        if ($Object -is [hashtable] -or $Object -is [System.Collections.IDictionary]) {
            if ($Object.ContainsKey($n)) { return $Object[$n] }
            foreach ($k in $Object.Keys) {
                if ([string]$k -eq $n) { return $Object[$k] }
            }
        }
        else {
            $p = $Object.PSObject.Properties[$n]
            if ($p) { return $p.Value }
        }
    }
    return $null
}

function ConvertTo-CaStringArray {
    param($Value)
    if ($null -eq $Value) { return @() }
    return @(
        $Value |
            Where-Object { $null -ne $_ -and [string]$_ -ne '' } |
            ForEach-Object { [string]$_ }
    )
}

function Get-CaUserScopeSummary {
    param($Users)
    if (-not $Users) { return '(none)' }
    $include = ConvertTo-CaStringArray (Get-CaGraphProp $Users @('includeUsers', 'IncludeUsers'))
    $exclude = ConvertTo-CaStringArray (Get-CaGraphProp $Users @('excludeUsers', 'ExcludeUsers'))
    $incGroups = ConvertTo-CaStringArray (Get-CaGraphProp $Users @('includeGroups', 'IncludeGroups'))
    $excGroups = ConvertTo-CaStringArray (Get-CaGraphProp $Users @('excludeGroups', 'ExcludeGroups'))
    $incRoles = ConvertTo-CaStringArray (Get-CaGraphProp $Users @('includeRoles', 'IncludeRoles'))
    $excRoles = ConvertTo-CaStringArray (Get-CaGraphProp $Users @('excludeRoles', 'ExcludeRoles'))
    $parts = [System.Collections.Generic.List[string]]::new()
    if ($include -contains 'All') { [void]$parts.Add('Include: All users') }
    elseif ($include.Count -gt 0) { [void]$parts.Add("Include users: $($include.Count)") }
    if ($exclude.Count -gt 0) { [void]$parts.Add("Exclude users: $($exclude.Count)") }
    if ($incGroups.Count -gt 0) { [void]$parts.Add("Include groups: $($incGroups.Count)") }
    if ($excGroups.Count -gt 0) { [void]$parts.Add("Exclude groups: $($excGroups.Count)") }
    if ($incRoles.Count -gt 0) { [void]$parts.Add("Include roles: $($incRoles.Count)") }
    if ($excRoles.Count -gt 0) { [void]$parts.Add("Exclude roles: $($excRoles.Count)") }
    if ($parts.Count -eq 0) { return '(none)' }
    return ($parts -join ' | ')
}

function Get-CaLocationIdsFromPolicy {
    param($Policy)
    $conditions = Get-CaGraphProp $Policy @('conditions', 'Conditions')
    $locs = Get-CaGraphProp $conditions @('locations', 'Locations')
    return @{
        Include = ConvertTo-CaStringArray (Get-CaGraphProp $locs @('includeLocations', 'IncludeLocations'))
        Exclude = ConvertTo-CaStringArray (Get-CaGraphProp $locs @('excludeLocations', 'ExcludeLocations'))
    }
}

function Get-CaIpRangeCidrs {
    param($Ranges)
    $out = [System.Collections.Generic.List[string]]::new()
    foreach ($r in @($Ranges)) {
        if ($null -eq $r -or $r -eq '') { continue }
        if ($r -is [string]) {
            if ($r -match '[0-9a-fA-F:\.]+/\d+') { [void]$out.Add($r.Trim()) }
            continue
        }
        $cidr = [string](Get-CaGraphProp $r @('cidrAddress', 'CidrAddress', 'cidr', 'addressPrefix'))
        if (-not $cidr) {
            foreach ($p in @($r.PSObject.Properties)) {
                if ($p.Name -like '@*') { continue }
                $v = [string]$p.Value
                if ($v -match '[0-9a-fA-F:\.]+/\d+') { $cidr = $v; break }
            }
        }
        if ($cidr) { [void]$out.Add($cidr.Trim()) }
    }
    return @($out)
}

function Get-CaPolicyPrincipalIds {
    param($Policy)
    $users = Get-CaGraphProp (Get-CaGraphProp $Policy @('conditions', 'Conditions')) @('users', 'Users')
    return [ordered]@{
        IncludeUsers  = @(ConvertTo-CaStringArray (Get-CaGraphProp $users @('includeUsers', 'IncludeUsers')))
        ExcludeUsers  = @(ConvertTo-CaStringArray (Get-CaGraphProp $users @('excludeUsers', 'ExcludeUsers')))
        IncludeGroups = @(ConvertTo-CaStringArray (Get-CaGraphProp $users @('includeGroups', 'IncludeGroups')))
        ExcludeGroups = @(ConvertTo-CaStringArray (Get-CaGraphProp $users @('excludeGroups', 'ExcludeGroups')))
        IncludeRoles  = @(ConvertTo-CaStringArray (Get-CaGraphProp $users @('includeRoles', 'IncludeRoles')))
        ExcludeRoles  = @(ConvertTo-CaStringArray (Get-CaGraphProp $users @('excludeRoles', 'ExcludeRoles')))
    }
}

function Get-CaExclusionSummary {
    param($Users)
    $exclude = ConvertTo-CaStringArray (Get-CaGraphProp $Users @('excludeUsers', 'ExcludeUsers'))
    $excGroups = ConvertTo-CaStringArray (Get-CaGraphProp $Users @('excludeGroups', 'ExcludeGroups'))
    $excRoles = ConvertTo-CaStringArray (Get-CaGraphProp $Users @('excludeRoles', 'ExcludeRoles'))
    $parts = [System.Collections.Generic.List[string]]::new()
    if ($exclude.Count -gt 0) { [void]$parts.Add("Users: $($exclude.Count)") }
    if ($excGroups.Count -gt 0) { [void]$parts.Add("Groups: $($excGroups.Count)") }
    if ($excRoles.Count -gt 0) { [void]$parts.Add("Roles: $($excRoles.Count)") }
    return [ordered]@{
        excludeUserCount  = $exclude.Count
        excludeGroupCount = $excGroups.Count
        excludeRoleCount  = $excRoles.Count
        text              = $(if ($parts.Count) { $parts -join ' | ' } else { '—' })
    }
}

function ConvertTo-CaPolicySummary {
    param($Policy)
    $conditions = Get-CaGraphProp $Policy @('conditions', 'Conditions')
    $users = Get-CaGraphProp $conditions @('users', 'Users')
    $loc = Get-CaLocationIdsFromPolicy -Policy $Policy
    $exc = Get-CaExclusionSummary -Users $users
    return [ordered]@{
        id                 = [string](Get-CaGraphProp $Policy @('id', 'Id'))
        displayName        = [string](Get-CaGraphProp $Policy @('displayName', 'DisplayName'))
        state              = [string](Get-CaGraphProp $Policy @('state', 'State'))
        userScope          = Get-CaUserScopeSummary -Users $users
        exclusions         = $exc.text
        excludeUserCount   = $exc.excludeUserCount
        excludeGroupCount  = $exc.excludeGroupCount
        includeLocationIds = @($loc.Include)
        excludeLocationIds = @($loc.Exclude)
    }
}

function Get-CaIsoCountryCodes {
    return @(
        'AD','AE','AF','AG','AI','AL','AM','AO','AQ','AR','AS','AT','AU','AW','AX','AZ',
        'BA','BB','BD','BE','BF','BG','BH','BI','BJ','BL','BM','BN','BO','BQ','BR','BS','BT','BV','BW','BY','BZ',
        'CA','CC','CD','CF','CG','CH','CI','CK','CL','CM','CN','CO','CR','CU','CV','CW','CX','CY','CZ',
        'DE','DJ','DK','DM','DO','DZ','EC','EE','EG','EH','ER','ES','ET','FI','FJ','FK','FM','FO','FR',
        'GA','GB','GD','GE','GF','GG','GH','GI','GL','GM','GN','GP','GQ','GR','GS','GT','GU','GW','GY',
        'HK','HM','HN','HR','HT','HU','ID','IE','IL','IM','IN','IO','IQ','IR','IS','IT',
        'JE','JM','JO','JP','KE','KG','KH','KI','KM','KN','KP','KR','KW','KY','KZ',
        'LA','LB','LC','LI','LK','LR','LS','LT','LU','LV','LY',
        'MA','MC','MD','ME','MF','MG','MH','MK','ML','MM','MN','MO','MP','MQ','MR','MS','MT','MU','MV','MW','MX','MY','MZ',
        'NA','NC','NE','NF','NG','NI','NL','NO','NP','NR','NU','NZ','OM',
        'PA','PE','PF','PG','PH','PK','PL','PM','PN','PR','PS','PT','PW','PY','QA','RE','RO','RS','RU','RW',
        'SA','SB','SC','SD','SE','SG','SH','SI','SJ','SK','SL','SM','SN','SO','SR','SS','ST','SV','SX','SY','SZ',
        'TC','TD','TF','TG','TH','TJ','TK','TL','TM','TN','TO','TR','TT','TV','TW','TZ',
        'UA','UG','UM','US','UY','UZ','VA','VC','VE','VG','VI','VN','VU','WF','WS','YE','YT','ZA','ZM','ZW'
    )
}

function Test-CaCountryLocationIsBlocklist {
    param([string]$DisplayName, [string[]]$Codes)
    $name = [string]$DisplayName
    if ($name -match '(?i)allow|trust|permit|safe|whitelist') { return $false }
    return (@($Codes).Count -gt 0)
}

function Get-CaAllowedCountryCodes {
    param([string[]]$Blocked)
    $blockedSet = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($c in @($Blocked)) {
        if ($c) { [void]$blockedSet.Add($c.Trim().ToUpperInvariant()) }
    }
    return @(Get-CaIsoCountryCodes | Where-Object { -not $blockedSet.Contains($_) })
}

function ConvertTo-CaLocationSummary {
    param($Location)
    $odata = [string](Get-CaGraphProp $Location @('@odata.type'))
    $type = 'Location'
    $detail = ''
    $cidrs = @()
    $countries = @()
    $allowed = @()
    $countryMode = ''
    $countries = @(ConvertTo-CaStringArray (Get-CaGraphProp $Location @('countriesAndRegions', 'CountriesAndRegions')) | ForEach-Object { $_.Trim().ToUpperInvariant() })
    if ($odata -match 'countryNamedLocation' -or $countries.Count -gt 0) {
        $type = 'Country'
        $name = [string](Get-CaGraphProp $Location @('displayName', 'DisplayName'))
        if (Test-CaCountryLocationIsBlocklist -DisplayName $name -Codes $countries) {
            $type = 'Country blocklist'
            $countryMode = 'blocklist'
            $allowed = @(Get-CaAllowedCountryCodes -Blocked $countries)
            $detail = if ($allowed.Count) { "Allowed: $($allowed -join ', ')" } else { 'Allowed: (none)' }
        }
        else {
            $countryMode = 'listed'
            $detail = ($countries -join ', ')
        }
    }
    else {
        $cidrs = @(Get-CaIpRangeCidrs (Get-CaGraphProp $Location @('ipRanges', 'IpRanges')))
        if ($odata -match 'ipNamedLocation' -or $cidrs.Count -gt 0) {
            $type = 'IP range'
            $detail = if ($cidrs.Count) { ($cidrs -join ', ') } else { '0 range(s)' }
        }
    }
    $trusted = $false
    $trustVal = Get-CaGraphProp $Location @('isTrusted', 'IsTrusted')
    if ($trustVal -eq $true -or [string]$trustVal -eq 'True') { $trusted = $true }
    return [ordered]@{
        id               = [string](Get-CaGraphProp $Location @('id', 'Id'))
        displayName      = [string](Get-CaGraphProp $Location @('displayName', 'DisplayName'))
        type             = $type
        detail           = $detail
        ipRanges         = @($cidrs)
        countries        = @($countries)
        allowedCountries = @($allowed)
        countryMode      = $countryMode
        isTrusted        = $trusted
    }
}

function ConvertTo-CaTrustedLocationRef {
    param($LocMap)
    $names = [System.Collections.Generic.List[string]]::new()
    $cidrs = [System.Collections.Generic.List[string]]::new()
    if ($LocMap) {
        foreach ($loc in @($LocMap.Values)) {
            if (-not $loc.isTrusted) { continue }
            if ($loc.displayName) { [void]$names.Add([string]$loc.displayName) }
            foreach ($c in @($loc.ipRanges)) {
                if ($c) { [void]$cidrs.Add([string]$c) }
            }
        }
    }
    $label = if ($names.Count) { "All trusted IPs ($($names -join ', '))" } else { 'All trusted IPs' }
    $detail = if ($cidrs.Count) { ($cidrs -join ', ') } else { '' }
    return [ordered]@{
        id               = 'AllTrusted'
        label            = $label
        type             = 'IP range'
        ipRanges         = @($cidrs)
        detail           = $detail
        countries        = @()
        allowedCountries = @()
        countryMode      = ''
        isTrusted        = $true
    }
}

function ConvertTo-CaLocationRef {
    param([string]$Id, $LocMap)
    if ($Id -eq 'All') { return [ordered]@{ id = $Id; label = 'All'; type = ''; ipRanges = @() } }
    if ($Id -eq 'AllTrusted') { return (ConvertTo-CaTrustedLocationRef -LocMap $LocMap) }
    if ($LocMap -and $LocMap.ContainsKey($Id)) {
        $loc = $LocMap[$Id]
        return [ordered]@{
            id               = $Id
            label            = [string]$loc.displayName
            type             = [string]$loc.type
            ipRanges         = @($loc.ipRanges)
            detail           = [string]$loc.detail
            countries        = @($loc.countries)
            allowedCountries = @($loc.allowedCountries)
            countryMode      = [string]$loc.countryMode
        }
    }
    return [ordered]@{ id = $Id; label = $Id; type = ''; ipRanges = @() }
}

function Assert-CaGraphRest {
    if (-not (Get-Command Test-GraphRestBearerToken -ErrorAction SilentlyContinue) -or -not (Test-GraphRestBearerToken)) {
        throw 'Graph REST bearer token is not set. Complete Graph Auth first.'
    }
}

function Get-CaCatalogPoliciesRaw {
    param([switch]$Refresh)
    if (-not $Refresh -and $script:CaCatalog.Policies) { return @($script:CaCatalog.Policies) }
    Assert-CaGraphRest
    $script:CaCatalog.Policies = @(Invoke-GraphRestPaged -Uri 'https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies')
    return @($script:CaCatalog.Policies)
}

function Get-CaCatalogLocationsRaw {
    param([switch]$Refresh)
    if (-not $Refresh -and $script:CaCatalog.Locations) { return @($script:CaCatalog.Locations) }
    Assert-CaGraphRest
    $script:CaCatalog.Locations = @(Invoke-GraphRestPaged -Uri 'https://graph.microsoft.com/v1.0/identity/conditionalAccess/namedLocations')
    return @($script:CaCatalog.Locations)
}

function Test-CaGuid {
    param([string]$Value)
    if (-not $Value) { return $false }
    $g = [guid]::Empty
    return [guid]::TryParse($Value, [ref]$g)
}

function Get-CaWellKnownRoleMap {
    if ($script:CaWellKnownRoleMap) { return $script:CaWellKnownRoleMap }
    $map = [System.Collections.Generic.Dictionary[string, string]]::new([StringComparer]::OrdinalIgnoreCase)
    $lines = @'
62e90394-69f5-4237-9190-012177145e10|Global Administrator
9b895d77-7237-7278-86d2-56d3bd9720bd|Application Administrator
9b895d92-2cd3-44c7-9d02-a6ac2d5ea5c3|Application Administrator
c4e39bd9-1100-46d3-8c65-fb160da0071f|Authentication Administrator
0526716b-113d-4c15-b2c8-68e3c22b9f80|Authentication Policy Administrator
b0f54661-2d74-4c50-afa3-1ec803f12efe|Billing Administrator
158c047a-c907-4556-b7ef-446551a6b5f7|Cloud Application Administrator
b1be1c3e-b65d-4f19-8427-f6fa0d97feb9|Conditional Access Administrator
44367163-eba1-44c3-98af-f5787879f96a|Dynamics 365 Administrator
29232cdf-9323-42fd-ade2-1d097af3e4de|Exchange Administrator
f2ef992c-3afb-46b9-b7cf-a126ee74c451|Global Reader
729827e3-9c14-49f7-bb1b-9608f156bbb8|Helpdesk Administrator
8ac3fc64-6eca-42ea-9e69-59f4c7b60eb2|Hybrid Identity Administrator
3a2c62dc-778b-46ae-8129-6b6e3c87330d|Intune Administrator
3a2c62db-5318-420d-8d74-23affee5d9d5|Intune Administrator
966707d0-3269-4727-9be2-8c5a1d0f0f90|Password Administrator
966707d0-3269-4727-9be2-8c3a10f19b9d|Password Administrator
1169501e-1cda-4adb-99d3-e6c394ecb116|Power Platform Administrator
11648597-926c-4cf3-9c36-bcebb0ba8dcc|Power Platform Administrator
7be44c8a-adaf-4e2a-884d-2d2b0dec86ba|Privileged Authentication Administrator
7be44c8a-adaf-4e2a-84d6-ab2649e08a13|Privileged Authentication Administrator
e8611ab8-c189-46e8-94e1-8229d6d9673b|Privileged Role Administrator
e8611ab8-c189-46e8-94e1-60213ab1f814|Privileged Role Administrator
194ae4cb-b126-40b2-bd5b-6091b380977d|Security Administrator
5d6b6bb7-de71-4623-b4af-96380a352509|Security Reader
f28a1f50-f6e7-4571-818b-6a12f2af6b6c|SharePoint Administrator
69091246-20e8-4a56-aa4d-066075b2a7a8|Teams Administrator
fe930be7-5e62-47db-91af-98c3a49a38b1|User Administrator
7698a772-787b-4ac8-901f-60d6b08affd2|Cloud Device Administrator
17315797-102d-40b4-93e0-432062caca18|Compliance Administrator
f023fd81-a637-4b56-95fd-791ac0226033|Service Support Administrator
4d6ac14f-3453-41d0-bef9-a3e0c569773a|License Administrator
88d8e3e3-8f55-4a1e-953a-9b9898b8876b|Directory Readers
5f2222b1-57c3-48ba-8ad5-d4759f1fde6f|Security Operator
ac434307-12b9-4fa1-a708-88bf58caabc1|Global Secure Access Administrator
7495fdc4-34c4-4d15-a289-98788ce399fd|Azure Information Protection Administrator
'@
    foreach ($line in ($lines -split '\r?\n')) {
        if ($line -match '^([0-9a-fA-F-]{36})\|(.+)$') {
            $map[$Matches[1]] = $Matches[2].Trim()
        }
    }
    $script:CaWellKnownRoleMap = $map
    return $map
}

function Get-CaWellKnownRoleName {
    param([string]$Id)
    if (-not $Id) { return $null }
    $map = Get-CaWellKnownRoleMap
    if ($map.ContainsKey($Id)) { return [string]$map[$Id] }
    return $null
}

function Get-CaDirectoryLabel {
    param($Object, [string]$FallbackId)
    $name = [string](Get-CaGraphProp $Object @('displayName', 'DisplayName'))
    $upn = [string](Get-CaGraphProp $Object @('userPrincipalName', 'UserPrincipalName'))
    $appId = [string](Get-CaGraphProp $Object @('appId', 'AppId'))
    if ($upn -and (Test-CaGuid $upn)) { $upn = '' }
    if ($name -and $upn) { return "$name ($upn)" }
    if ($name) { return $name }
    if ($upn) { return $upn }
    if ($appId) { return "App $appId" }
    if ($FallbackId) { return $FallbackId }
    return ''
}

function Set-CaDirectoryCacheEntry {
    param([string]$Id, [string]$ODataType, [string]$Label)
    if (-not $Id -or -not $Label) { return }
    if ($ODataType -match 'group') { $script:CaCatalog.Groups[$Id] = $Label }
    elseif ($ODataType -match 'directoryRole|roleDefinition') { $script:CaCatalog.Roles[$Id] = $Label }
    else { $script:CaCatalog.Users[$Id] = $Label }
}

function Test-CaDirectoryIdCached {
    param([string]$Id)
    if (-not $Id) { return $true }
    return $script:CaCatalog.Users.ContainsKey($Id) -or $script:CaCatalog.Groups.ContainsKey($Id) -or $script:CaCatalog.Roles.ContainsKey($Id)
}

function Ensure-CaWellKnownRolesSeeded {
    $map = Get-CaWellKnownRoleMap
    foreach ($key in $map.Keys) {
        if (-not $script:CaCatalog.Roles.ContainsKey($key)) {
            $script:CaCatalog.Roles[$key] = $map[$key]
        }
    }
}

function Ensure-CaRoleDefinitionCache {
    Ensure-CaWellKnownRolesSeeded
    if ($script:CaCatalog.RoleGraphLoaded) { return }
    if (-not (Get-Command Test-GraphRestBearerToken -ErrorAction SilentlyContinue) -or -not (Test-GraphRestBearerToken)) { return }
    try {
        $roles = @(Invoke-GraphRestPaged -Uri 'https://graph.microsoft.com/v1.0/roleManagement/directory/roleDefinitions?$select=id,displayName,templateId')
        foreach ($r in $roles) {
            $name = [string](Get-CaGraphProp $r @('displayName', 'DisplayName'))
            if (-not $name) { continue }
            $rid = [string](Get-CaGraphProp $r @('id', 'Id'))
            $tid = [string](Get-CaGraphProp $r @('templateId', 'TemplateId'))
            if ($rid) { $script:CaCatalog.Roles[$rid] = $name }
            if ($tid) { $script:CaCatalog.Roles[$tid] = $name }
        }
        $script:CaCatalog.RoleGraphLoaded = $true
    }
    catch {
        $script:CaCatalog.RoleGraphLoaded = $true
    }
}

function Resolve-CaMissingDirectoryId {
    param([string]$Id)
    if (-not $Id -or (Test-CaDirectoryIdCached $Id)) { return }
    $wn = Get-CaWellKnownRoleName -Id $Id
    if ($wn) { $script:CaCatalog.Roles[$Id] = $wn; return }
    $lookups = @(
        @{ Uri = "https://graph.microsoft.com/v1.0/users/$Id" + '?$select=id,displayName,userPrincipalName'; Kind = 'user' }
        @{ Uri = "https://graph.microsoft.com/v1.0/groups/$Id" + '?$select=id,displayName'; Kind = 'group' }
        @{ Uri = "https://graph.microsoft.com/v1.0/servicePrincipals/$Id" + '?$select=id,displayName,appId'; Kind = 'user' }
        @{ Uri = "https://graph.microsoft.com/v1.0/directoryObjects/$Id"; Kind = '' }
        @{ Uri = "https://graph.microsoft.com/v1.0/roleManagement/directory/roleDefinitions/$Id" + '?$select=id,displayName,templateId'; Kind = 'role' }
    )
    foreach ($lookup in $lookups) {
        try {
            $obj = Invoke-GraphRestRequest -Uri $lookup.Uri
            $label = Get-CaDirectoryLabel -Object $obj -FallbackId $Id
            $odata = [string](Get-CaGraphProp $obj @('@odata.type'))
            if ($lookup.Kind -eq 'role' -or $odata -match 'roleDefinition|directoryRole') {
                $script:CaCatalog.Roles[$Id] = $(if ($label) { $label } else { $Id })
                $tid = [string](Get-CaGraphProp $obj @('templateId', 'TemplateId'))
                if ($tid -and $label) { $script:CaCatalog.Roles[$tid] = $label }
                return
            }
            if ($lookup.Kind -eq 'group' -or $odata -match 'group') {
                $script:CaCatalog.Groups[$Id] = $(if ($label) { $label } else { $Id })
                return
            }
            if ($label -and $label -ne $Id) {
                Set-CaDirectoryCacheEntry -Id $Id -ODataType $odata -Label $label
                return
            }
        }
        catch { }
    }
    $script:CaCatalog.Users[$Id] = "Unknown directory object"
}

function Invoke-CaGetByIds {
    param([string[]]$Ids)
    $need = @($Ids | Where-Object { $_ -and -not (Test-CaDirectoryIdCached $_) })
    if ($need.Count -eq 0) { return }
    $idList = [System.Collections.Generic.List[string]]::new()
    $typeList = [System.Collections.Generic.List[string]]::new()
    foreach ($id in $need) { [void]$idList.Add($id) }
    foreach ($t in @('user', 'group', 'servicePrincipal')) { [void]$typeList.Add($t) }
    $body = @{ ids = $idList; types = $typeList } | ConvertTo-Json -Compress
    try {
        $resp = Invoke-GraphRestRequest -Method POST -Uri 'https://graph.microsoft.com/v1.0/directoryObjects/getByIds' -Body $body
        foreach ($obj in @($resp.value)) {
            $id = [string](Get-CaGraphProp $obj @('id', 'Id'))
            if (-not $id) { continue }
            $odata = [string](Get-CaGraphProp $obj @('@odata.type'))
            $label = Get-CaDirectoryLabel -Object $obj -FallbackId $id
            Set-CaDirectoryCacheEntry -Id $id -ODataType $odata -Label $label
        }
    }
    catch { }
    foreach ($id in $need) {
        if (-not (Test-CaDirectoryIdCached $id)) {
            Resolve-CaMissingDirectoryId -Id $id
        }
    }
}

function Resolve-CaDirectoryNames {
    param([string[]]$Ids)
    $skip = @('All', 'GuestsOrExternalUsers', 'AllTrusted', 'None')
    $need = @($Ids | Where-Object { $_ -and ($skip -notcontains $_) } | Select-Object -Unique)
    Ensure-CaWellKnownRolesSeeded
    if (@($need | Where-Object { -not (Test-CaDirectoryIdCached $_) }).Count -gt 0) {
        Ensure-CaRoleDefinitionCache
    }
    $missing = @($need | Where-Object { -not (Test-CaDirectoryIdCached $_) })
    if ($missing.Count -eq 0) { return }
    Assert-CaGraphRest
    $batch = [System.Collections.Generic.List[string]]::new()
    foreach ($id in $missing) {
        [void]$batch.Add($id)
        if ($batch.Count -ge 100) {
            Invoke-CaGetByIds -Ids @($batch)
            $batch.Clear()
        }
    }
    if ($batch.Count -gt 0) { Invoke-CaGetByIds -Ids @($batch) }
}

function ConvertTo-CaNamedList {
    param([string[]]$Ids, [string]$Kind)
    $out = [System.Collections.Generic.List[object]]::new()
    foreach ($id in @($Ids)) {
        $label = $id
        if ($id -eq 'All') { $label = $(if ($Kind -eq 'role') { 'All roles' } else { 'All users' }) }
        elseif ($id -eq 'GuestsOrExternalUsers') { $label = 'Guests or external users' }
        elseif ($id -eq 'AllTrusted') { $label = 'All trusted IPs' }
        elseif ($Kind -eq 'role' -and $script:CaCatalog.Roles.ContainsKey($id)) { $label = [string]$script:CaCatalog.Roles[$id] }
        elseif ($script:CaCatalog.Users.ContainsKey($id)) { $label = [string]$script:CaCatalog.Users[$id] }
        elseif ($script:CaCatalog.Groups.ContainsKey($id)) { $label = [string]$script:CaCatalog.Groups[$id] }
        elseif ($script:CaCatalog.Roles.ContainsKey($id)) { $label = [string]$script:CaCatalog.Roles[$id] }
        else {
            $wn = Get-CaWellKnownRoleName -Id $id
            if ($wn) { $label = $wn }
        }
        [void]$out.Add([ordered]@{ id = $id; label = $label })
    }
    return @($out)
}

function Get-CaCatalogPolicyDetail {
    param([Parameter(Mandatory)][string]$PolicyId)
    $policies = Get-CaCatalogPoliciesRaw
    $policy = $policies | Where-Object { [string](Get-CaGraphProp $_ @('id', 'Id')) -eq $PolicyId } | Select-Object -First 1
    if (-not $policy) { throw "Policy not found: $PolicyId" }
    $users = Get-CaGraphProp (Get-CaGraphProp $policy @('conditions', 'Conditions')) @('users', 'Users')
    $includeUsers = ConvertTo-CaStringArray (Get-CaGraphProp $users @('includeUsers', 'IncludeUsers'))
    $excludeUsers = ConvertTo-CaStringArray (Get-CaGraphProp $users @('excludeUsers', 'ExcludeUsers'))
    $includeGroups = ConvertTo-CaStringArray (Get-CaGraphProp $users @('includeGroups', 'IncludeGroups'))
    $excludeGroups = ConvertTo-CaStringArray (Get-CaGraphProp $users @('excludeGroups', 'ExcludeGroups'))
    $includeRoles = ConvertTo-CaStringArray (Get-CaGraphProp $users @('includeRoles', 'IncludeRoles'))
    $excludeRoles = ConvertTo-CaStringArray (Get-CaGraphProp $users @('excludeRoles', 'ExcludeRoles'))
    Resolve-CaDirectoryNames -Ids (@($includeUsers + $excludeUsers + $includeGroups + $excludeGroups + $includeRoles + $excludeRoles))
    $locs = Get-CaCatalogLocationsRaw
    $locMap = @{}
    foreach ($l in $locs) {
        $sum = ConvertTo-CaLocationSummary -Location $l
        if ($sum.id) { $locMap[$sum.id] = $sum }
    }
    $locIds = Get-CaLocationIdsFromPolicy -Policy $policy
    $namedInc = @(foreach ($id in $locIds.Include) { ConvertTo-CaLocationRef -Id $id -LocMap $locMap })
    $namedExc = @(foreach ($id in $locIds.Exclude) { ConvertTo-CaLocationRef -Id $id -LocMap $locMap })
    $summary = ConvertTo-CaPolicySummary -Policy $policy
    return [ordered]@{
        id               = $summary.id
        displayName      = $summary.displayName
        state            = $summary.state
        userScope        = $summary.userScope
        exclusions       = $summary.exclusions
        includeUsers     = @(ConvertTo-CaNamedList -Ids $includeUsers -Kind user)
        excludeUsers     = @(ConvertTo-CaNamedList -Ids $excludeUsers -Kind user)
        includeGroups    = @(ConvertTo-CaNamedList -Ids $includeGroups -Kind group)
        excludeGroups    = @(ConvertTo-CaNamedList -Ids $excludeGroups -Kind group)
        includeRoles     = @(ConvertTo-CaNamedList -Ids $includeRoles -Kind role)
        excludeRoles     = @(ConvertTo-CaNamedList -Ids $excludeRoles -Kind role)
        includeLocations = $namedInc
        excludeLocations = $namedExc
        grantControls    = Get-CaGraphProp $policy @('grantControls', 'GrantControls')
        sessionControls  = Get-CaGraphProp $policy @('sessionControls', 'SessionControls')
        conditions       = Get-CaGraphProp $policy @('conditions', 'Conditions')
    }
}

function Get-CaCatalogCommandToken {
    param([string]$Command, [string]$Name)
    if ($Command -match "\|${Name}:([^|]+)") { return $Matches[1].Trim() }
    return $null
}

function Get-CaCatalogJsonPayload {
    param([string]$Command)
    if (-not $Command) { return $null }
    $i = $Command.IndexOf('|')
    if ($i -lt 0) { return $null }
    $rest = $Command.Substring($i + 1).Trim()
    if (-not $rest.StartsWith('{')) { return $null }
    return ($rest | ConvertFrom-Json)
}

function Write-CaCatalogAudit {
    param(
        [string]$AuditFolder,
        [string]$Action,
        [string]$Result = 'success',
        [string]$Detail,
        [string]$Upn
    )
    if (-not $AuditFolder) { return }
    if (Get-Command Write-RemediationAuditCsv -ErrorAction SilentlyContinue) {
        Write-RemediationAuditCsv -Folder $AuditFolder -Tenant '' -ClientNumber 0 -Upn $Upn -Action $Action -Result $Result -Detail $Detail
    }
}

function Merge-CaPolicyUserIds {
    param(
        [string[]]$Current,
        [string[]]$Ids,
        [Parameter(Mandatory)][ValidateSet('add', 'remove')][string]$Op,
        [Parameter(Mandatory)][ValidateSet('include', 'exclude')][string]$Scope,
        [switch]$HasOtherIncludes,
        [switch]$AllowAllFallback
    )
    $specials = @('All', 'GuestsOrExternalUsers', 'None')
    $wanted = @($Ids | Where-Object { $_ -and ($specials -notcontains $_) } | Select-Object -Unique)
    $cur = @($Current | Where-Object { $_ })
    if ($Op -eq 'add') {
        $list = @($cur | Where-Object { $specials -notcontains $_ })
        foreach ($id in $wanted) {
            if ($list -notcontains $id) { $list += $id }
        }
        if ($Scope -eq 'include' -and $list.Count -eq 0) {
            throw 'No users to include after merge.'
        }
        return @($list)
    }
    $remove = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($id in $wanted) { [void]$remove.Add($id) }
    $list = @($cur | Where-Object { -not $remove.Contains($_) })
    if ($Scope -eq 'include' -and $list.Count -eq 0 -and -not $HasOtherIncludes) {
        if ($AllowAllFallback) { return @('All') }
        throw 'Removing the last included user would leave no include target. Confirm fallback to All users, or keep a group/role include.'
    }
    return @($list)
}

function Get-CaPolicyUsersHashtable {
    param($Policy)
    $users = Get-CaGraphProp (Get-CaGraphProp $Policy @('conditions', 'Conditions')) @('users', 'Users')
    $h = [ordered]@{}
    foreach ($pair in @(
            @{ Key = 'includeUsers'; Names = @('includeUsers', 'IncludeUsers') }
            @{ Key = 'excludeUsers'; Names = @('excludeUsers', 'ExcludeUsers') }
            @{ Key = 'includeGroups'; Names = @('includeGroups', 'IncludeGroups') }
            @{ Key = 'excludeGroups'; Names = @('excludeGroups', 'ExcludeGroups') }
            @{ Key = 'includeRoles'; Names = @('includeRoles', 'IncludeRoles') }
            @{ Key = 'excludeRoles'; Names = @('excludeRoles', 'ExcludeRoles') }
        )) {
        $vals = @(ConvertTo-CaStringArray (Get-CaGraphProp $users $pair.Names))
        if ($vals.Count -eq 0) { continue }
        $list = [System.Collections.Generic.List[string]]::new()
        foreach ($v in $vals) { [void]$list.Add($v) }
        $h[$pair.Key] = $list
    }
    return $h
}

function Search-CaCatalogUsers {
    param([string]$Query, [int]$Top = 50)
    $q = ([string]$Query).Trim()
    if ($q.Length -lt 2) { throw 'Search query must be at least 2 characters.' }
    Assert-CaGraphRest
    if ($Top -lt 1) { $Top = 1 }
    if ($Top -gt 50) { $Top = 50 }
    $found = [System.Collections.Generic.Dictionary[string, object]]::new([StringComparer]::OrdinalIgnoreCase)
    function Add-CaSearchHit {
        param($obj)
        $id = [string](Get-CaGraphProp $obj @('id', 'Id'))
        if (-not $id -or $found.ContainsKey($id)) { return }
        $name = [string](Get-CaGraphProp $obj @('displayName', 'DisplayName'))
        $upn = [string](Get-CaGraphProp $obj @('userPrincipalName', 'UserPrincipalName'))
        $mail = [string](Get-CaGraphProp $obj @('mail', 'Mail'))
        $dept = [string](Get-CaGraphProp $obj @('department', 'Department'))
        $label = if ($name -and $upn) { "$name ($upn)" } elseif ($upn) { $upn } elseif ($name) { $name } else { $id }
        $found[$id] = [ordered]@{
            id                = $id
            displayName       = $name
            userPrincipalName = $upn
            mail              = $mail
            department        = $dept
            label             = $label
        }
        $script:CaCatalog.Users[$id] = $label
    }
    $select = '$select=id,displayName,userPrincipalName,mail,department'
    if ((Test-CaGuid $q) -or $q -match '@') {
        try {
            $enc = [Uri]::EscapeDataString($q)
            $u = Invoke-GraphRestRequest -Uri "https://graph.microsoft.com/v1.0/users/$enc`?$select"
            Add-CaSearchHit $u
        }
        catch { }
    }
    $esc = $q.Replace("'", "''")
    foreach ($filter in @(
            "contains(displayName,'$esc')"
            "startswith(userPrincipalName,'$esc')"
            "startswith(mail,'$esc')"
        )) {
        try {
            $resp = Invoke-GraphRestRequest -Uri "https://graph.microsoft.com/v1.0/users?`$filter=$filter&`$top=$Top&$select"
            foreach ($u in @($resp.value)) { Add-CaSearchHit $u }
        }
        catch { }
    }
    try {
        $inner = $q.Replace('\', '\\').Replace('"', '\"')
        $searchUri = "https://graph.microsoft.com/v1.0/users?`$search=`"displayName:$inner`" OR `"userPrincipalName:$inner`"&`$top=$Top&$select"
        $resp = Invoke-GraphRestRequest -Uri $searchUri -Headers @{ ConsistencyLevel = 'eventual' }
        foreach ($u in @($resp.value)) { Add-CaSearchHit $u }
    }
    catch { }
    return @($found.Values | Sort-Object { $_.displayName })
}

function Get-CaCatalogPolicyRawById {
    param([string]$PolicyId, [switch]$Refresh)
    $policies = Get-CaCatalogPoliciesRaw -Refresh:$Refresh
    $policy = $policies | Where-Object { [string](Get-CaGraphProp $_ @('id', 'Id')) -eq $PolicyId } | Select-Object -First 1
    if (-not $policy -and -not $Refresh) {
        return (Get-CaCatalogPolicyRawById -PolicyId $PolicyId -Refresh)
    }
    return $policy
}

function Update-CaCatalogPolicyUsers {
    param(
        [Parameter(Mandatory)][string]$PolicyId,
        [Parameter(Mandatory)][ValidateSet('include', 'exclude')][string]$Scope,
        [Parameter(Mandatory)][ValidateSet('add', 'remove')][string]$Op,
        [string[]]$Ids,
        [switch]$AllowAllFallback
    )
    Assert-CaGraphRest
    $wanted = @($Ids | Where-Object { $_ -and $_ -ne 'All' -and $_ -ne 'GuestsOrExternalUsers' -and $_ -ne 'None' })
    if ($wanted.Count -eq 0) { throw 'No user ids to change.' }
    $policy = Get-CaCatalogPolicyRawById -PolicyId $PolicyId
    if (-not $policy) { throw "Policy not found: $PolicyId" }
    $h = Get-CaPolicyUsersHashtable -Policy $policy
    $hasOther = $false
    if ($h.Contains('includeGroups') -and @($h.includeGroups).Count -gt 0) { $hasOther = $true }
    if ($h.Contains('includeRoles') -and @($h.includeRoles).Count -gt 0) { $hasOther = $true }
    $key = if ($Scope -eq 'include') { 'includeUsers' } else { 'excludeUsers' }
    $current = @()
    if ($h.Contains($key)) { $current = @($h[$key]) }
    $merged = Merge-CaPolicyUserIds -Current $current -Ids $wanted -Op $Op -Scope $Scope -HasOtherIncludes:$hasOther -AllowAllFallback:$AllowAllFallback
    if ($merged.Count -gt 0) {
        $list = [System.Collections.Generic.List[string]]::new()
        foreach ($v in $merged) { [void]$list.Add([string]$v) }
        $h[$key] = $list
    }
    elseif ($h.Contains($key)) {
        $h.Remove($key)
    }
    $body = @{ conditions = @{ users = $h } } | ConvertTo-Json -Compress -Depth 6
    Invoke-GraphRestRequest -Method PATCH -Uri "https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies/$PolicyId" -Body $body
    $script:CaCatalog.Policies = $null
    return (Get-CaCatalogPolicyDetail -PolicyId $PolicyId)
}

function Invoke-CaCatalogWorkerCommand {
    param(
        [Parameter(Mandatory)][string]$Command,
        [string]$AuditFolder
    )
    try {
        if ($Command -eq 'CA_LIST_POLICIES' -or $Command -like 'CA_LIST_POLICIES|*') {
            $refresh = $Command -match 'REFRESH:1'
            $raw = Get-CaCatalogPoliciesRaw -Refresh:$refresh
            $locs = Get-CaCatalogLocationsRaw -Refresh:$refresh
            $locMap = @{}
            foreach ($l in $locs) {
                $sum = ConvertTo-CaLocationSummary -Location $l
                if ($sum.id) { $locMap[$sum.id] = $sum }
            }
            $allPrincipalIds = [System.Collections.Generic.List[string]]::new()
            foreach ($p in $raw) {
                $ids = Get-CaPolicyPrincipalIds -Policy $p
                foreach ($id in @($ids.IncludeUsers + $ids.ExcludeUsers + $ids.IncludeGroups + $ids.ExcludeGroups + $ids.IncludeRoles + $ids.ExcludeRoles)) {
                    if ($id) { [void]$allPrincipalIds.Add($id) }
                }
            }
            Resolve-CaDirectoryNames -Ids @($allPrincipalIds)
            $policies = @(foreach ($p in $raw) {
                    $s = ConvertTo-CaPolicySummary -Policy $p
                    $ids = Get-CaPolicyPrincipalIds -Policy $p
                    $incRefs = @(foreach ($id in $s.includeLocationIds) { ConvertTo-CaLocationRef -Id $id -LocMap $locMap })
                    $excRefs = @(foreach ($id in $s.excludeLocationIds) { ConvertTo-CaLocationRef -Id $id -LocMap $locMap })
                    $s['includeUsers'] = @(ConvertTo-CaNamedList -Ids $ids.IncludeUsers -Kind user)
                    $s['excludeUsers'] = @(ConvertTo-CaNamedList -Ids $ids.ExcludeUsers -Kind user)
                    $s['includeGroups'] = @(ConvertTo-CaNamedList -Ids $ids.IncludeGroups -Kind group)
                    $s['excludeGroups'] = @(ConvertTo-CaNamedList -Ids $ids.ExcludeGroups -Kind group)
                    $s['includeRoles'] = @(ConvertTo-CaNamedList -Ids $ids.IncludeRoles -Kind role)
                    $s['excludeRoles'] = @(ConvertTo-CaNamedList -Ids $ids.ExcludeRoles -Kind role)
                    $s['includeLocations'] = @($incRefs)
                    $s['excludeLocations'] = @($excRefs)
                    $s['includeIpRanges'] = @($incRefs | ForEach-Object { @($_.ipRanges) } | Where-Object { $_ })
                    $s['excludeIpRanges'] = @($excRefs | ForEach-Object { @($_.ipRanges) } | Where-Object { $_ })
                    $s
                })
            $payload = @{ Policies = @($policies); Count = $policies.Count } | ConvertTo-Json -Compress -Depth 8
            return "CA_POLICIES:$payload"
        }
        if ($Command -eq 'CA_LIST_LOCATIONS' -or $Command -like 'CA_LIST_LOCATIONS|*') {
            $refresh = $Command -match 'REFRESH:1'
            $raw = Get-CaCatalogLocationsRaw -Refresh:$refresh
            $locations = @($raw | ForEach-Object { ConvertTo-CaLocationSummary -Location $_ })
            $payload = @{ Locations = @($locations); Count = $locations.Count } | ConvertTo-Json -Compress -Depth 6
            return "CA_LOCATIONS:$payload"
        }
        if ($Command -match '^CA_GET_POLICY') {
            $id = Get-CaCatalogCommandToken -Command $Command -Name 'POLICY_ID'
            if (-not $id) {
                $jp = Get-CaCatalogJsonPayload -Command $Command
                if ($jp) { $id = [string]$jp.policyId }
            }
            if (-not $id) { return 'CA_FAILED:Missing POLICY_ID' }
            $detail = Get-CaCatalogPolicyDetail -PolicyId $id
            $payload = @{ Policy = $detail } | ConvertTo-Json -Compress -Depth 8
            return "CA_POLICY:$payload"
        }
        if ($Command -match '^CA_SEARCH_USERS') {
            $jp = Get-CaCatalogJsonPayload -Command $Command
            $q = if ($jp) { [string]$jp.q } else { Get-CaCatalogCommandToken -Command $Command -Name 'Q' }
            $users = @(Search-CaCatalogUsers -Query $q)
            $payload = @{ Users = @($users); Count = $users.Count } | ConvertTo-Json -Compress -Depth 5
            return "CA_USERS:$payload"
        }
        if ($Command -match '^CA_PATCH_USERS') {
            $jp = Get-CaCatalogJsonPayload -Command $Command
            if (-not $jp) { return 'CA_FAILED:Missing JSON payload' }
            $policyId = [string]$jp.policyId
            $scope = [string]$jp.scope
            $op = [string]$jp.op
            $ids = @(ConvertTo-CaStringArray $jp.ids)
            $allowAll = $false
            if ($jp.PSObject.Properties['allowAllFallback'] -and $jp.allowAllFallback) { $allowAll = $true }
            if (-not $policyId) { return 'CA_FAILED:Missing policyId' }
            $detail = Update-CaCatalogPolicyUsers -PolicyId $policyId -Scope $scope -Op $op -Ids $ids -AllowAllFallback:$allowAll
            $audit = [ordered]@{
                Action = "ca-users-$scope-$op"
                Result = 'success'
                Detail = "policy=$policyId count=$($ids.Count)"
                UPN    = ($ids -join ';')
            }
            Write-CaCatalogAudit -AuditFolder $AuditFolder -Action $audit.Action -Result $audit.Result -Detail $audit.Detail -Upn $audit.UPN
            $payload = @{ Policy = $detail; Audit = $audit } | ConvertTo-Json -Compress -Depth 8
            return "CA_POLICY:$payload"
        }
        return 'CA_FAILED:Unknown CA command'
    }
    catch {
        $msg = $_.Exception.Message -replace '[\r\n]+', ' '
        return "CA_FAILED:$msg"
    }
}

Export-ModuleMember -Function Clear-CaCatalogCache, Get-CaUserScopeSummary, Get-CaExclusionSummary, ConvertTo-CaPolicySummary, ConvertTo-CaLocationSummary, Get-CaIpRangeCidrs, Get-CaAllowedCountryCodes, Test-CaCountryLocationIsBlocklist, Invoke-CaCatalogWorkerCommand, Get-CaCatalogPoliciesRaw, Get-CaCatalogLocationsRaw, Get-CaCatalogPolicyDetail, ConvertTo-CaNamedList, Get-CaWellKnownRoleName, Merge-CaPolicyUserIds, Search-CaCatalogUsers, Get-CaCatalogJsonPayload
