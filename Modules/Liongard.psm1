function Get-LiongardProperty {
    param([object]$Object, [string[]]$Names, $Default = $null)
    if (-not $Object) { return $Default }
    foreach ($name in $Names) {
        if ($Object.PSObject.Properties.Name -contains $name) {
            $val = $Object.$name
            if ($null -ne $val -and -not [string]::IsNullOrWhiteSpace([string]$val)) { return $val }
        }
    }
    return $Default
}

function Get-LiongardCredentialsFromSettings {
    param([object]$Settings = $null)

    if (-not $Settings) {
        Import-Module (Join-Path $PSScriptRoot 'Settings.psm1') -Force -ErrorAction SilentlyContinue
        $Settings = Get-AppSettings
    }

    return [pscustomobject]@{
        Instance     = [string]$Settings.LiongardInstance
        AccessKey    = [string]$Settings.LiongardAccessKey
        AccessSecret = [string]$Settings.LiongardAccessSecret
    }
}

function Test-LiongardCredentialsComplete {
    param([Parameter(Mandatory = $true)]$Credentials)
    foreach ($name in @('Instance', 'AccessKey', 'AccessSecret')) {
        if ([string]::IsNullOrWhiteSpace($Credentials.$name)) { return $false }
    }
    return $true
}

function Get-LiongardConfigurationStatus {
    $creds = Get-LiongardCredentialsFromSettings
    $configured = Test-LiongardCredentialsComplete -Credentials $creds
    return [pscustomobject]@{
        Configured = $configured
        Instance   = $creds.Instance
        Message    = if ($configured) { 'Liongard API credentials configured.' } else { 'Set LiongardInstance, LiongardAccessKey, and LiongardAccessSecret in EOA settings.' }
    }
}

function Get-LiongardApiBaseUri {
    param(
        [Parameter(Mandatory = $true)][string]$Instance,
        [ValidateSet('v1', 'v2')][string]$Version = 'v2'
    )

    $inst = $Instance.Trim().TrimEnd('/')
    if ($inst -match '^https?://') {
        $base = $inst.TrimEnd('/')
    } else {
        $base = "https://$inst.app.liongard.com"
    }
    return "$base/api/$Version/"
}

function Invoke-LiongardApi {
    [CmdletBinding()]
    param(
        [ValidateSet('v1', 'v2')][string]$Version = 'v2',
        [Parameter(Mandatory = $true)][ValidateSet('GET', 'POST', 'PATCH', 'PUT', 'DELETE')][string]$Method,
        [Parameter(Mandatory = $true)][string]$RelativePath,
        [object]$Body = $null,
        [hashtable]$Query = $null,
        [object]$Credentials = $null
    )

    if (-not $Credentials) { $Credentials = Get-LiongardCredentialsFromSettings }
    if (-not (Test-LiongardCredentialsComplete -Credentials $Credentials)) {
        throw 'Liongard is not fully configured.'
    }

    $baseUri = Get-LiongardApiBaseUri -Instance $Credentials.Instance -Version $Version
    $relative = $RelativePath.TrimStart('/')
    $uriBuilder = [UriBuilder]::new([Uri]::new([Uri]$baseUri, $relative))
    if ($Query -and $Query.Count -gt 0) {
        $parts = [System.Collections.Generic.List[string]]::new()
        foreach ($key in @($Query.Keys)) {
            $parts.Add("$key=$([Uri]::EscapeDataString([string]$Query[$key]))")
        }
        $uriBuilder.Query = ($parts -join '&')
    }
    $uri = $uriBuilder.Uri

    $authRaw = "$($Credentials.AccessKey):$($Credentials.AccessSecret)"
    $authBytes = [Text.Encoding]::UTF8.GetBytes($authRaw)
    $headers = @{
        'X-ROAR-API-KEY' = [Convert]::ToBase64String($authBytes)
        Accept           = 'application/json'
    }

    $params = @{
        Uri         = $uri
        Method      = $Method
        Headers     = $headers
        ErrorAction = 'Stop'
    }
    if ($null -ne $Body) {
        $params.Body = ($Body | ConvertTo-Json -Depth 10 -Compress)
        $params.ContentType = 'application/json'
    }

    return Invoke-RestMethod @params
}

function Get-LiongardEnvironmentNameScore {
    param(
        [string]$CompanyName,
        [string]$EnvironmentName
    )

    if ([string]::IsNullOrWhiteSpace($CompanyName) -or [string]::IsNullOrWhiteSpace($EnvironmentName)) { return 0 }

    $a = ($CompanyName -replace '[^a-zA-Z0-9]', '').ToLowerInvariant()
    $b = ($EnvironmentName -replace '[^a-zA-Z0-9]', '').ToLowerInvariant()
    if ([string]::IsNullOrWhiteSpace($a) -or [string]::IsNullOrWhiteSpace($b)) { return 0 }
    if ($a -eq $b) { return 100 }
    if ($b.Contains($a) -or $a.Contains($b)) { return 85 }

    $aWords = @($CompanyName -split '\s+' | ForEach-Object { $_.Trim() } | Where-Object { $_.Length -gt 2 })
    $matched = 0
    foreach ($word in $aWords) {
        if ($EnvironmentName -match [regex]::Escape($word)) { $matched++ }
    }
    if ($aWords.Count -gt 0) {
        return [int](($matched / $aWords.Count) * 70)
    }
    return 0
}

function Get-LiongardApiItems {
    param([object]$Result)
    if (-not $Result) { return @() }
    if ($Result.PSObject.Properties.Name -contains 'items') { return @($Result.items) }
    if ($Result.PSObject.Properties.Name -contains 'Environments') { return @($Result.Environments) }
    if ($Result.PSObject.Properties.Name -contains 'Systems') { return @($Result.Systems) }
    if ($Result -is [System.Collections.IEnumerable] -and $Result -isnot [string]) { return @($Result) }
    return @()
}

function Find-LiongardEnvironmentByCompanyName {
    param(
        [Parameter(Mandatory = $true)][string]$CompanyName,
        [int]$MinScore = 50
    )

    $best = $null
    $page = 1
    do {
        $result = Invoke-LiongardApi -Method GET -RelativePath 'environments' -Query @{ page = $page; pageSize = 100 }
        $items = @(Get-LiongardApiItems -Result $result)
        foreach ($env in $items) {
            $name = [string](Get-LiongardProperty -Object $env -Names @('Name', 'name'))
            $id = [int](Get-LiongardProperty -Object $env -Names @('ID', 'Id', 'id') -Default 0)
            if ($id -le 0) { continue }
            $score = Get-LiongardEnvironmentNameScore -CompanyName $CompanyName -EnvironmentName $name
            if ($score -ge $MinScore -and (-not $best -or $score -gt $best.Score)) {
                $best = [pscustomobject]@{ environmentId = $id; environmentName = $name; score = $score }
            }
        }
        $page++
        $hasMore = $items.Count -ge 100
    } while ($hasMore -and $page -le 20)

    return $best
}

function Get-LiongardSystemsForEnvironment {
    param([Parameter(Mandatory = $true)][int]$EnvironmentId)

    $systems = [System.Collections.Generic.List[object]]::new()
    $page = 1
    do {
        try {
            $result = Invoke-LiongardApi -Version v1 -Method GET -RelativePath 'systems' -Query @{
                page          = $page
                pageSize      = 100
                EnvironmentID = $EnvironmentId
            }
        } catch {
            break
        }

        $items = @(Get-LiongardApiItems -Result $result)
        foreach ($item in $items) { [void]$systems.Add($item) }
        $page++
        $hasMore = $items.Count -ge 100
    } while ($hasMore -and $page -le 20)

    return @($systems)
}

function Get-LiongardInspectorSystems {
    param(
        [Parameter(Mandatory = $true)][int]$EnvironmentId,
        [Parameter(Mandatory = $true)][string]$InspectorPattern
    )

    $all = Get-LiongardSystemsForEnvironment -EnvironmentId $EnvironmentId
    return @($all | Where-Object {
        $inspector = [string](Get-LiongardProperty -Object $_ -Names @('Inspector', 'InspectorName', 'InspectorType') -Default '')
        $name = [string](Get-LiongardProperty -Object $_ -Names @('Name', 'FriendlyName') -Default '')
        ($inspector -match $InspectorPattern) -or ($name -match $InspectorPattern)
    })
}

function Get-LiongardHuntressCapabilities {
    param([Parameter(Mandatory = $true)][int]$EnvironmentId)

    $systems = Get-LiongardInspectorSystems -EnvironmentId $EnvironmentId -InspectorPattern 'Huntress'
    $hasHuntress = $systems.Count -gt 0
    $systemId = 0
    if ($systems.Count -gt 0) {
        $systemId = [int](Get-LiongardProperty -Object $systems[0] -Names @('ID', 'Id') -Default 0)
    }

    $caps = @{
        hasHuntress      = $hasHuntress
        edr              = $hasHuntress
        itdr             = $false
        siem             = $false
        huntressSystemId = $systemId
        agentCount       = $null
    }

    if (-not $hasHuntress) { return $caps }

    try {
        $metricResult = Invoke-LiongardApi -Method POST -RelativePath 'metrics/evaluate' -Body @{
            EnvironmentIDs = @($EnvironmentId)
            pageSize       = 100
        }
        $rows = @(Get-LiongardApiItems -Result $metricResult)

        foreach ($row in $rows) {
            $name = [string](Get-LiongardProperty -Object $row -Names @('Name', 'MetricName', 'Path') -Default '').ToLowerInvariant()
            $value = [double](Get-LiongardProperty -Object $row -Names @('Value', 'Result') -Default 0)
            if ($name -match 'itdr|microsoft_365_users_count|identity threat') {
                if ($value -gt 0) { $caps.itdr = $true }
            }
            if ($name -match 'siem|logs_sources_count') {
                if ($value -gt 0) { $caps.siem = $true }
            }
            if ($name -match 'agent' -and $name -notmatch 'non') {
                if ($null -eq $caps.agentCount -or $value -gt $caps.agentCount) { $caps.agentCount = [int]$value }
            }
        }
    } catch {
        Write-Verbose "Liongard metrics evaluate failed: $($_.Exception.Message)"
    }

    return $caps
}

function Get-LiongardSentinelOneCapabilities {
    param(
        [Parameter(Mandatory = $true)][int]$EnvironmentId,
        [string]$TicketContent = '',
        [string]$SocSource = 'unknown'
    )

    $systems = Get-LiongardInspectorSystems -EnvironmentId $EnvironmentId -InspectorPattern 'SentinelOne'
    $hasS1 = $systems.Count -gt 0
    $system = if ($systems.Count -gt 0) { $systems[0] } else { $null }

    $siteId = ''
    $siteName = ''
    $instanceId = ''

    if ($system) {
        $siteName = [string](Get-LiongardProperty -Object $system -Names @('Name', 'FriendlyName') -Default '')
        $siteId = [string](Get-LiongardProperty -Object $system -Names @('ExternalID', 'SiteID', 'SiteId') -Default '')
        if ([string]::IsNullOrWhiteSpace($siteId)) {
            $desc = [string](Get-LiongardProperty -Object $system -Names @('Description') -Default '')
            if ($desc -match 'site[_\s-]?id[:\s]+([A-Za-z0-9-]+)') { $siteId = $Matches[1] }
        }
    }

    $consoleHint = $SocSource
    if ($consoleHint -eq 'unknown' -and $hasS1) { $consoleHint = 'connectwise' }

    return @{
        hasSentinelOne = $hasS1
        siteId         = $siteId
        siteName       = $siteName
        instanceId     = $instanceId
        consoleHint    = $consoleHint
        systemId       = if ($system) { [int](Get-LiongardProperty -Object $system -Names @('ID', 'Id') -Default 0) } else { 0 }
    }
}

function Resolve-LiongardClient {
    param(
        [Parameter(Mandatory = $true)][string]$CompanyName,
        [string]$TicketContent = '',
        [int]$EnvironmentId = 0
    )

    Import-Module (Join-Path $PSScriptRoot 'Settings.psm1') -Force -ErrorAction SilentlyContinue
    Import-Module (Join-Path $PSScriptRoot 'SecurityIntegrations.psm1') -Force -ErrorAction SilentlyContinue
    $socSource = Get-SocSourceFromTicket -TicketContent $TicketContent
    $securityStack = Get-SecurityStackFromTicket -TicketContent $TicketContent

    $match = $null
    if ($EnvironmentId -gt 0) {
        try {
            $envDetail = Invoke-LiongardApi -Method GET -RelativePath "environments/$EnvironmentId"
            $envName = [string](Get-LiongardProperty -Object $envDetail -Names @('Name', 'name') -Default $CompanyName)
            $match = [pscustomobject]@{ environmentId = $EnvironmentId; environmentName = $envName; score = 100 }
        } catch {
            $match = Find-LiongardEnvironmentByCompanyName -CompanyName $CompanyName
        }
    } else {
        $match = Find-LiongardEnvironmentByCompanyName -CompanyName $CompanyName
    }

    if (-not $match) {
        return @{
            matched          = $false
            companyName      = $CompanyName
            socSource        = $socSource
            securityStack    = $securityStack
            huntress         = @{ hasHuntress = $false; edr = $false; itdr = $false; siem = $false }
            sentinelOne      = @{ hasSentinelOne = $false; consoleHint = $socSource }
            recommendedPulls = (Get-RecommendedSecurityPulls -AlertTypes (Get-AlertTypeFromTicket -TicketContent $TicketContent) -SocSource $socSource -SecurityStack $securityStack)
        }
    }

    $huntress = Get-LiongardHuntressCapabilities -EnvironmentId $match.environmentId
    $sentinelOne = Get-LiongardSentinelOneCapabilities -EnvironmentId $match.environmentId -TicketContent $TicketContent -SocSource $socSource
    if ($socSource -ne 'unknown') { $sentinelOne.consoleHint = $socSource }

    $alertTypes = Get-AlertTypeFromTicket -TicketContent $TicketContent
    $recommended = Get-RecommendedSecurityPulls -AlertTypes $alertTypes -HuntressCaps $huntress -SentinelOneCaps $sentinelOne -SocSource $socSource -SecurityStack $securityStack

    return @{
        matched          = $true
        environmentId    = $match.environmentId
        environmentName  = $match.environmentName
        matchScore       = $match.score
        companyName      = $CompanyName
        socSource        = $socSource
        securityStack    = $securityStack
        alertTypes       = $alertTypes
        huntress         = $huntress
        sentinelOne      = $sentinelOne
        recommendedPulls = $recommended
    }
}

function Export-LiongardContext {
    param(
        [Parameter(Mandatory = $true)][int]$EnvironmentId,
        [Parameter(Mandatory = $true)][string]$ExportFolder,
        [string]$TicketNumber = '',
        [datetime]$StartDate = $null,
        [datetime]$EndDate = $null
    )

    Import-Module (Join-Path $PSScriptRoot 'SecurityIntegrations.psm1') -Force
    $files = [System.Collections.Generic.List[string]]::new()

    $systems = Get-LiongardSystemsForEnvironment -EnvironmentId $EnvironmentId
    if ($systems.Count -gt 0) {
        $rows = foreach ($s in $systems) {
            [pscustomobject]@{
                SystemId  = Get-LiongardProperty -Object $s -Names @('ID', 'Id')
                Name      = Get-LiongardProperty -Object $s -Names @('Name', 'FriendlyName')
                Inspector = Get-LiongardProperty -Object $s -Names @('Inspector', 'InspectorName')
                Status    = Get-LiongardProperty -Object $s -Names @('Status', 'State')
                LastSeen  = Get-LiongardProperty -Object $s -Names @('LastSeen', 'UpdatedOn')
            }
        }
        [void]$files.Add((Export-SecurityIntegrationCsv -Folder $ExportFolder -BaseName 'LiongardSystems' -Rows $rows -TicketNumber $TicketNumber))
    }

    if (-not $StartDate) { $StartDate = (Get-Date).AddDays(-7) }
    if (-not $EndDate) { $EndDate = Get-Date }

    try {
        $detections = Invoke-LiongardApi -Method POST -RelativePath 'detections' -Body @{
            EnvironmentIDs = @($EnvironmentId)
            StartDate      = $StartDate.ToString('o')
            EndDate        = $EndDate.ToString('o')
            pageSize       = 200
        }
        $items = @(Get-LiongardApiItems -Result $detections)
        if ($items.Count -gt 0) {
            [void]$files.Add((Export-SecurityIntegrationCsv -Folder $ExportFolder -BaseName 'LiongardDetections' -Rows $items -TicketNumber $TicketNumber))
        }
    } catch {
        Write-Verbose "Liongard detections export skipped: $($_.Exception.Message)"
    }

    return @($files)
}

Export-ModuleMember -Function Get-LiongardConfigurationStatus,Invoke-LiongardApi,Find-LiongardEnvironmentByCompanyName,Get-LiongardHuntressCapabilities,Get-LiongardSentinelOneCapabilities,Resolve-LiongardClient,Export-LiongardContext
