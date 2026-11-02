function Get-SentinelOneProfileNames {
    return @('connectwise', 'barracuda_xdr')
}

function Get-SentinelOneProfilesFromSettings {
    param([object]$Settings = $null)

    if (-not $Settings) {
        Import-Module (Join-Path $PSScriptRoot 'Settings.psm1') -Force -ErrorAction SilentlyContinue
        $Settings = Get-AppSettings
    }

    return @{
        connectwise = [pscustomobject]@{
            ProfileName = 'connectwise'
            ConsoleType = 'connectwise'
            InstanceId  = [string]$Settings.SentinelOneConnectWiseInstanceId
            ApiToken    = [string]$Settings.SentinelOneConnectWiseApiToken
            ReadOnly    = $false
        }
        barracuda_xdr = [pscustomobject]@{
            ProfileName = 'barracuda_xdr'
            ConsoleType = 'barracuda_xdr'
            InstanceId  = [string]$Settings.SentinelOneBarracudaInstanceId
            ApiToken    = [string]$Settings.SentinelOneBarracudaApiToken
            ReadOnly    = $true
        }
    }
}

function Test-SentinelOneProfileComplete {
    param([Parameter(Mandatory = $true)]$Profile)
    if ([string]::IsNullOrWhiteSpace($Profile.InstanceId)) { return $false }
    if ([string]::IsNullOrWhiteSpace($Profile.ApiToken)) { return $false }
    return $true
}

function Get-SentinelOneConfigurationStatus {
    $profiles = Get-SentinelOneProfilesFromSettings
    $status = @{}
    foreach ($name in Get-SentinelOneProfileNames) {
        $p = $profiles[$name]
        $status[$name] = @{
            configured = (Test-SentinelOneProfileComplete -Profile $p)
            readOnly   = [bool]$p.ReadOnly
            instanceId = $p.InstanceId
        }
    }
    return [pscustomobject]@{
        Profiles = $status
        Message  = 'Configure SentinelOneConnectWise* and/or SentinelOneBarracuda* tokens in EOA settings.'
    }
}

function Get-SentinelOneApiBaseUri {
    param([Parameter(Mandatory = $true)][string]$InstanceId)

    $inst = $InstanceId.Trim().TrimEnd('/')
    if ($inst -match '^https?://') { return ($inst.TrimEnd('/') + '/web/api/v2.1/') }
    if ($inst -match '\.sentinelone\.(net|com)') { return "https://$inst/web/api/v2.1/" }
    return "https://$inst.sentinelone.net/web/api/v2.1/"
}

function Invoke-SentinelOneApi {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$ProfileName,
        [Parameter(Mandatory = $true)][ValidateSet('GET', 'POST', 'PATCH', 'PUT', 'DELETE')][string]$Method,
        [Parameter(Mandatory = $true)][string]$RelativePath,
        [hashtable]$Query = $null,
        [object]$Body = $null,
        [object]$Profile = $null,
        [int]$MaxRetries = 3
    )

    if (-not $Profile) {
        $profiles = Get-SentinelOneProfilesFromSettings
        if (-not $profiles.ContainsKey($ProfileName)) { throw "Unknown SentinelOne profile: $ProfileName" }
        $Profile = $profiles[$ProfileName]
    }
    if (-not (Test-SentinelOneProfileComplete -Profile $Profile)) {
        throw "SentinelOne profile '$ProfileName' is not fully configured."
    }
    if ($Profile.ReadOnly -and $Method -ne 'GET') {
        throw "SentinelOne profile '$ProfileName' is read-only; $Method is not allowed."
    }

    $baseUri = Get-SentinelOneApiBaseUri -InstanceId $Profile.InstanceId
    $relative = $RelativePath.TrimStart('/')
    $uriBuilder = [UriBuilder]::new([Uri]::new([Uri]$baseUri, $relative))
    if ($Query -and $Query.Count -gt 0) {
        $parts = [System.Collections.Generic.List[string]]::new()
        foreach ($key in @($Query.Keys)) {
            $parts.Add("$key=$([Uri]::EscapeDataString([string]$Query[$key]))")
        }
        $uriBuilder.Query = ($parts -join '&')
    }

    $headers = @{
        Authorization = "ApiToken $($Profile.ApiToken)"
        Accept        = 'application/json'
    }

    $attempt = 0
    while ($true) {
        $attempt++
        try {
            $params = @{
                Uri         = $uriBuilder.Uri
                Method      = $Method
                Headers     = $headers
                ErrorAction = 'Stop'
            }
            if ($null -ne $Body) {
                $params.Body = ($Body | ConvertTo-Json -Depth 10 -Compress)
                $params.ContentType = 'application/json'
            }
            return Invoke-RestMethod @params
        } catch {
            $status = $null
            if ($_.Exception.Response) { $status = [int]$_.Exception.Response.StatusCode }
            if ($status -eq 429 -and $attempt -lt $MaxRetries) {
                Start-Sleep -Seconds ([Math]::Pow(2, $attempt))
                continue
            }
            throw
        }
    }
}

function Get-SentinelOnePagedData {
    param(
        [Parameter(Mandatory = $true)][string]$ProfileName,
        [Parameter(Mandatory = $true)][string]$RelativePath,
        [hashtable]$Query = $null,
        [string]$DataKey = 'data',
        [int]$MaxPages = 20
    )

    $items = [System.Collections.Generic.List[object]]::new()
    $cursor = $null
    $page = 0
    do {
        $q = @{}
        if ($Query) { $Query.GetEnumerator() | ForEach-Object { $q[$_.Key] = $_.Value } }
        $q.limit = 200
        if ($cursor) { $q.cursor = $cursor }

        $result = Invoke-SentinelOneApi -ProfileName $ProfileName -Method GET -RelativePath $RelativePath -Query $q
        $batch = @()
        if ($result.PSObject.Properties.Name -contains $DataKey) { $batch = @($result.$DataKey) }
        elseif ($result -is [System.Collections.IEnumerable] -and $result -isnot [string]) { $batch = @($result) }

        foreach ($item in $batch) { if ($item) { [void]$items.Add($item) } }

        $cursor = $null
        if ($result.PSObject.Properties.Name -contains 'pagination') {
            $cursor = [string]$result.pagination.nextCursor
        }
        $page++
    } while (-not [string]::IsNullOrWhiteSpace($cursor) -and $page -lt $MaxPages)

    return @($items)
}

function Get-S1Threats {
    param(
        [Parameter(Mandatory = $true)][string]$ProfileName,
        [string]$SiteId = '',
        [datetime]$CreatedAfter = $null,
        [datetime]$CreatedBefore = $null
    )

    $query = @{}
    if (-not [string]::IsNullOrWhiteSpace($SiteId)) { $query.siteIds = $SiteId }
    if ($CreatedAfter) { $query.createdAt__gte = $CreatedAfter.ToUniversalTime().ToString('o') }
    if ($CreatedBefore) { $query.createdAt__lte = $CreatedBefore.ToUniversalTime().ToString('o') }
    return @(Get-SentinelOnePagedData -ProfileName $ProfileName -RelativePath 'threats' -Query $query)
}

function Get-S1Agents {
    param(
        [Parameter(Mandatory = $true)][string]$ProfileName,
        [string]$SiteId = '',
        [string]$Hostname = ''
    )

    $query = @{}
    if (-not [string]::IsNullOrWhiteSpace($SiteId)) { $query.siteIds = $SiteId }
    $items = @(Get-SentinelOnePagedData -ProfileName $ProfileName -RelativePath 'agents' -Query $query)
    if (-not [string]::IsNullOrWhiteSpace($Hostname)) {
        $items = @($items | Where-Object {
            [string]$_.computerName -match [regex]::Escape($Hostname) -or
            [string]$_.hostname -match [regex]::Escape($Hostname)
        })
    }
    return @($items)
}

function Get-S1Activities {
    param(
        [Parameter(Mandatory = $true)][string]$ProfileName,
        [string]$SiteId = '',
        [datetime]$CreatedAfter = $null,
        [datetime]$CreatedBefore = $null
    )

    $query = @{}
    if (-not [string]::IsNullOrWhiteSpace($SiteId)) { $query.siteIds = $SiteId }
    if ($CreatedAfter) { $query.createdAt__gte = $CreatedAfter.ToUniversalTime().ToString('o') }
    if ($CreatedBefore) { $query.createdAt__lte = $CreatedBefore.ToUniversalTime().ToString('o') }
    return @(Get-SentinelOnePagedData -ProfileName $ProfileName -RelativePath 'activities' -Query $query)
}

function Resolve-SentinelOneSite {
    param(
        [Parameter(Mandatory = $true)][string]$CompanyName,
        [string]$TicketContent = '',
        [string]$ProfileName = '',
        [string]$SiteIdHint = '',
        [int]$LiongardEnvironmentId = 0
    )

    Import-Module (Join-Path $PSScriptRoot 'Settings.psm1') -Force -ErrorAction SilentlyContinue
    $socSource = Get-SocSourceFromTicket -TicketContent $TicketContent

    if ([string]::IsNullOrWhiteSpace($ProfileName)) {
        if ($socSource -eq 'barracuda_xdr') { $ProfileName = 'barracuda_xdr' }
        elseif ($socSource -eq 'connectwise') { $ProfileName = 'connectwise' }
        else { $ProfileName = 'connectwise' }
    }

    $profiles = Get-SentinelOneProfilesFromSettings
    $profile = $profiles[$ProfileName]
    $configured = Test-SentinelOneProfileComplete -Profile $profile

    $siteId = $SiteIdHint
    $siteName = ''
    $instanceId = $profile.InstanceId

    if ($LiongardEnvironmentId -gt 0) {
        Import-Module (Join-Path $PSScriptRoot 'Liongard.psm1') -Force -ErrorAction SilentlyContinue
        $s1Caps = Get-LiongardSentinelOneCapabilities -EnvironmentId $LiongardEnvironmentId -TicketContent $TicketContent -SocSource $socSource
        if ([string]::IsNullOrWhiteSpace($siteId)) { $siteId = [string]$s1Caps.siteId }
        $siteName = [string]$s1Caps.siteName
        if (-not [string]::IsNullOrWhiteSpace([string]$s1Caps.instanceId)) { $instanceId = [string]$s1Caps.instanceId }
    }

    return @{
        profileName   = $ProfileName
        configured    = $configured
        readOnly      = [bool]$profile.ReadOnly
        instanceId    = $instanceId
        siteId        = $siteId
        siteName      = $siteName
        socSource     = $socSource
        barracudaFallback = if ($ProfileName -eq 'barracuda_xdr' -and -not $configured) {
            'Barracuda S1 API unavailable — use Barracuda portal and ticket IOCs. No ConnectWise profile fallback.'
        } else { '' }
    }
}

function Get-SentinelOnePreviewCounts {
    param(
        [Parameter(Mandatory = $true)][string]$ProfileName,
        [string]$SiteId = '',
        [datetime]$CreatedAfter = $null,
        [datetime]$CreatedBefore = $null
    )

    return @{
        threats    = @(Get-S1Threats -ProfileName $ProfileName -SiteId $SiteId -CreatedAfter $CreatedAfter -CreatedBefore $CreatedBefore).Count
        agents     = @(Get-S1Agents -ProfileName $ProfileName -SiteId $SiteId).Count
        activities = @(Get-S1Activities -ProfileName $ProfileName -SiteId $SiteId -CreatedAfter $CreatedAfter -CreatedBefore $CreatedBefore).Count
    }
}

function Export-SentinelOneInvestigation {
    param(
        [Parameter(Mandatory = $true)][string]$ProfileName,
        [Parameter(Mandatory = $true)][string]$ExportFolder,
        [hashtable]$Selections = @{},
        [string]$SiteId = '',
        [string]$TicketNumber = '',
        [datetime]$CreatedAfter = $null,
        [datetime]$CreatedBefore = $null
    )

    Import-Module (Join-Path $PSScriptRoot 'SecurityIntegrations.psm1') -Force
    $files = [System.Collections.Generic.List[string]]::new()

    if ($Selections.threats) {
        $rows = Get-S1Threats -ProfileName $ProfileName -SiteId $SiteId -CreatedAfter $CreatedAfter -CreatedBefore $CreatedBefore
        if ($rows.Count -gt 0) {
            [void]$files.Add((Export-SecurityIntegrationCsv -Folder $ExportFolder -BaseName 'S1Threats' -Rows $rows -TicketNumber $TicketNumber))
        }
    }
    if ($Selections.agents) {
        $rows = Get-S1Agents -ProfileName $ProfileName -SiteId $SiteId
        if ($rows.Count -gt 0) {
            [void]$files.Add((Export-SecurityIntegrationCsv -Folder $ExportFolder -BaseName 'S1Agents' -Rows $rows -TicketNumber $TicketNumber))
        }
    }
    if ($Selections.activities) {
        $rows = Get-S1Activities -ProfileName $ProfileName -SiteId $SiteId -CreatedAfter $CreatedAfter -CreatedBefore $CreatedBefore
        if ($rows.Count -gt 0) {
            [void]$files.Add((Export-SecurityIntegrationJson -Folder $ExportFolder -BaseName 'S1Activities' -Object $rows -TicketNumber $TicketNumber))
        }
    }

    $summary = @{
        profileName = $ProfileName
        siteId      = $SiteId
        exportedAt  = (Get-Date).ToString('o')
        fileCount   = $files.Count
    }
    [void]$files.Add((Export-SecurityIntegrationJson -Folder $ExportFolder -BaseName 'S1Summary' -Object $summary -TicketNumber $TicketNumber))
    return @($files)
}

Export-ModuleMember -Function Get-SentinelOneConfigurationStatus,Invoke-SentinelOneApi,Resolve-SentinelOneSite,Get-SentinelOnePreviewCounts,Export-SentinelOneInvestigation,Get-S1Threats,Get-S1Agents,Get-S1Activities,Get-SentinelOneProfilesFromSettings
