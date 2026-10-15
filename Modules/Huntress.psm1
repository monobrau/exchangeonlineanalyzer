function Get-HuntressCredentialsFromSettings {
    param([object]$Settings = $null)

    if (-not $Settings) {
        Import-Module (Join-Path $PSScriptRoot 'Settings.psm1') -Force -ErrorAction SilentlyContinue
        $Settings = Get-AppSettings
    }

    return [pscustomobject]@{
        ApiKey    = [string]$Settings.HuntressApiKey
        ApiSecret = [string]$Settings.HuntressApiSecret
    }
}

function Test-HuntressCredentialsComplete {
    param([Parameter(Mandatory = $true)]$Credentials)
    foreach ($name in @('ApiKey', 'ApiSecret')) {
        if ([string]::IsNullOrWhiteSpace($Credentials.$name)) { return $false }
    }
    return $true
}

function Get-HuntressConfigurationStatus {
    $creds = Get-HuntressCredentialsFromSettings
    $configured = Test-HuntressCredentialsComplete -Credentials $creds
    return [pscustomobject]@{
        Configured = $configured
        Message    = if ($configured) { 'Huntress API credentials configured.' } else { 'Set HuntressApiKey and HuntressApiSecret in EOA settings.' }
    }
}

function Invoke-HuntressApi {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][ValidateSet('GET', 'POST', 'PATCH', 'PUT', 'DELETE')][string]$Method,
        [Parameter(Mandatory = $true)][string]$RelativePath,
        [hashtable]$Query = $null,
        [object]$Body = $null,
        [object]$Credentials = $null,
        [int]$MaxRetries = 3
    )

    if (-not $Credentials) { $Credentials = Get-HuntressCredentialsFromSettings }
    if (-not (Test-HuntressCredentialsComplete -Credentials $Credentials)) {
        throw 'Huntress is not fully configured.'
    }

    $relative = $RelativePath.TrimStart('/')
    $uriBuilder = [UriBuilder]::new('https://api.huntress.io/v1/' + $relative)
    if ($Query -and $Query.Count -gt 0) {
        $parts = [System.Collections.Generic.List[string]]::new()
        foreach ($key in @($Query.Keys)) {
            $parts.Add("$key=$([Uri]::EscapeDataString([string]$Query[$key]))")
        }
        $uriBuilder.Query = ($parts -join '&')
    }

    $authBytes = [Text.Encoding]::UTF8.GetBytes("$($Credentials.ApiKey):$($Credentials.ApiSecret)")
    $headers = @{
        Authorization = 'Basic ' + [Convert]::ToBase64String($authBytes)
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

function Get-HuntressPagedItems {
    param(
        [Parameter(Mandatory = $true)][string]$RelativePath,
        [hashtable]$Query = $null,
        [int]$MaxPages = 20
    )

    $items = [System.Collections.Generic.List[object]]::new()
    $pageToken = $null
    $page = 0
    do {
        $q = @{}
        if ($Query) { $Query.GetEnumerator() | ForEach-Object { $q[$_.Key] = $_.Value } }
        $q.limit = 200
        if ($pageToken) { $q.page_token = $pageToken }

        $result = Invoke-HuntressApi -Method GET -RelativePath $RelativePath -Query $q
        $batch = @()
        if ($result.PSObject.Properties.Name -contains 'signals') { $batch = @($result.signals) }
        elseif ($result.PSObject.Properties.Name -contains 'incident_reports') { $batch = @($result.incident_reports) }
        elseif ($result.PSObject.Properties.Name -contains 'agents') { $batch = @($result.agents) }
        elseif ($result.PSObject.Properties.Name -contains 'identities') { $batch = @($result.identities) }
        elseif ($result.PSObject.Properties.Name -contains 'escalations') { $batch = @($result.escalations) }
        elseif ($result.PSObject.Properties.Name -contains 'organizations') { $batch = @($result.organizations) }
        elseif ($result.PSObject.Properties.Name -contains 'remediations') { $batch = @($result.remediations) }
        else { $batch = @($result) }

        foreach ($item in $batch) { if ($item) { [void]$items.Add($item) } }

        $pageToken = $null
        if ($result.PSObject.Properties.Name -contains 'pagination') {
            $pageToken = [string]$result.pagination.next_page_token
        }
        $page++
    } while (-not [string]::IsNullOrWhiteSpace($pageToken) -and $page -lt $MaxPages)

    return @($items)
}

function Find-HuntressOrganizationByName {
    param(
        [Parameter(Mandatory = $true)][string]$CompanyName,
        [int]$MinScore = 50
    )

    $orgs = Get-HuntressPagedItems -RelativePath 'organizations'
    $best = $null
    foreach ($org in $orgs) {
        $name = [string]$org.name
        if ([string]::IsNullOrWhiteSpace($name)) { continue }
        $a = ($CompanyName -replace '[^a-zA-Z0-9]', '').ToLowerInvariant()
        $b = ($name -replace '[^a-zA-Z0-9]', '').ToLowerInvariant()
        $score = 0
        if ($a -eq $b) { $score = 100 }
        elseif ($b.Contains($a) -or $a.Contains($b)) { $score = 85 }
        else {
            $words = @($CompanyName -split '\s+' | Where-Object { $_.Length -gt 2 })
            $matched = ($words | Where-Object { $name -match [regex]::Escape($_) }).Count
            if ($words.Count -gt 0) { $score = [int](($matched / $words.Count) * 70) }
        }
        if ($score -ge $MinScore -and (-not $best -or $score -gt $best.score)) {
            $best = [pscustomobject]@{ organizationId = [int]$org.id; organizationName = $name; score = $score }
        }
    }
    return $best
}

function Get-HuntressSignals {
    param(
        [int]$OrganizationId = 0,
        [string[]]$SignalTypes = @(),
        [datetime]$UpdatedSince = $null
    )

    $query = @{}
    if ($OrganizationId -gt 0) { $query.organization_id = $OrganizationId }
    if ($UpdatedSince) { $query.updated_at_min = $UpdatedSince.ToUniversalTime().ToString('o') }

    $items = Get-HuntressPagedItems -RelativePath 'signals' -Query $query
    if ($SignalTypes -and $SignalTypes.Count -gt 0) {
        $types = @($SignalTypes | ForEach-Object { $_.ToLowerInvariant() })
        $items = @($items | Where-Object {
            $t = [string]$_.signal_type
            if ([string]::IsNullOrWhiteSpace($t)) { $t = [string]$_.type }
            $types -contains $t.ToLowerInvariant()
        })
    }
    return @($items)
}

function Get-HuntressIncidentReports {
    param(
        [int]$OrganizationId = 0,
        [datetime]$UpdatedSince = $null
    )

    $query = @{}
    if ($OrganizationId -gt 0) { $query.organization_id = $OrganizationId }
    if ($UpdatedSince) { $query.updated_at_min = $UpdatedSince.ToUniversalTime().ToString('o') }
    return @(Get-HuntressPagedItems -RelativePath 'incident_reports' -Query $query)
}

function Get-HuntressAgents {
    param([int]$OrganizationId = 0)

    $query = @{}
    if ($OrganizationId -gt 0) { $query.organization_id = $OrganizationId }
    return @(Get-HuntressPagedItems -RelativePath 'agents' -Query $query)
}

function Get-HuntressIdentities {
    param(
        [int]$OrganizationId = 0,
        [string]$UpnFilter = ''
    )

    $query = @{}
    if ($OrganizationId -gt 0) { $query.organization_id = $OrganizationId }
    $items = @(Get-HuntressPagedItems -RelativePath 'identities' -Query $query)
    if (-not [string]::IsNullOrWhiteSpace($UpnFilter)) {
        $items = @($items | Where-Object {
            [string]$_.username -match [regex]::Escape($UpnFilter) -or
            [string]$_.email -match [regex]::Escape($UpnFilter)
        })
    }
    return @($items)
}

function Get-HuntressEscalations {
    param([int]$OrganizationId = 0)

    $query = @{}
    if ($OrganizationId -gt 0) { $query.organization_id = $OrganizationId }
    return @(Get-HuntressPagedItems -RelativePath 'escalations' -Query $query)
}

function Get-HuntressPreviewCounts {
    param(
        [Parameter(Mandatory = $true)][int]$OrganizationId,
        [datetime]$UpdatedSince = $null
    )

    return @{
        signalsFootholds       = @(Get-HuntressSignals -OrganizationId $OrganizationId -SignalTypes @('Footholds') -UpdatedSince $UpdatedSince).Count
        signalsAntivirus       = @(Get-HuntressSignals -OrganizationId $OrganizationId -SignalTypes @('Antivirus') -UpdatedSince $UpdatedSince).Count
        signalsProcessInsights = @(Get-HuntressSignals -OrganizationId $OrganizationId -SignalTypes @('Process Insights') -UpdatedSince $UpdatedSince).Count
        signalsManagedItdr     = @(Get-HuntressSignals -OrganizationId $OrganizationId -SignalTypes @('Managed ITDR') -UpdatedSince $UpdatedSince).Count
        signalsSiem            = @(Get-HuntressSignals -OrganizationId $OrganizationId -SignalTypes @('SIEM') -UpdatedSince $UpdatedSince).Count
        incidents              = @(Get-HuntressIncidentReports -OrganizationId $OrganizationId -UpdatedSince $UpdatedSince).Count
        agents                 = @(Get-HuntressAgents -OrganizationId $OrganizationId).Count
        identities             = @(Get-HuntressIdentities -OrganizationId $OrganizationId).Count
        escalations            = @(Get-HuntressEscalations -OrganizationId $OrganizationId).Count
    }
}

function Export-HuntressInvestigation {
    param(
        [Parameter(Mandatory = $true)][int]$OrganizationId,
        [Parameter(Mandatory = $true)][string]$ExportFolder,
        [hashtable]$Selections = @{},
        [string]$TicketNumber = '',
        [datetime]$UpdatedSince = $null
    )

    Import-Module (Join-Path $PSScriptRoot 'SecurityIntegrations.psm1') -Force
    $files = [System.Collections.Generic.List[string]]::new()

    if ($Selections.signalsFootholds) {
        $rows = Get-HuntressSignals -OrganizationId $OrganizationId -SignalTypes @('Footholds') -UpdatedSince $UpdatedSince
        if ($rows.Count -gt 0) {
            [void]$files.Add((Export-SecurityIntegrationCsv -Folder $ExportFolder -BaseName 'HuntressSignals_Footholds' -Rows $rows -TicketNumber $TicketNumber))
        }
    }
    if ($Selections.signalsAntivirus) {
        $rows = Get-HuntressSignals -OrganizationId $OrganizationId -SignalTypes @('Antivirus') -UpdatedSince $UpdatedSince
        if ($rows.Count -gt 0) {
            [void]$files.Add((Export-SecurityIntegrationCsv -Folder $ExportFolder -BaseName 'HuntressSignals_Antivirus' -Rows $rows -TicketNumber $TicketNumber))
        }
    }
    if ($Selections.signalsProcessInsights) {
        $rows = Get-HuntressSignals -OrganizationId $OrganizationId -SignalTypes @('Process Insights') -UpdatedSince $UpdatedSince
        if ($rows.Count -gt 0) {
            [void]$files.Add((Export-SecurityIntegrationCsv -Folder $ExportFolder -BaseName 'HuntressSignals_ProcessInsights' -Rows $rows -TicketNumber $TicketNumber))
        }
    }
    if ($Selections.signalsManagedItdr) {
        $rows = Get-HuntressSignals -OrganizationId $OrganizationId -SignalTypes @('Managed ITDR') -UpdatedSince $UpdatedSince
        if ($rows.Count -gt 0) {
            [void]$files.Add((Export-SecurityIntegrationCsv -Folder $ExportFolder -BaseName 'HuntressSignals_ManagedITDR' -Rows $rows -TicketNumber $TicketNumber))
        }
    }
    if ($Selections.signalsSiem) {
        $rows = Get-HuntressSignals -OrganizationId $OrganizationId -SignalTypes @('SIEM') -UpdatedSince $UpdatedSince
        if ($rows.Count -gt 0) {
            [void]$files.Add((Export-SecurityIntegrationCsv -Folder $ExportFolder -BaseName 'HuntressSignals_SIEM' -Rows $rows -TicketNumber $TicketNumber))
        }
    }
    if ($Selections.incidents) {
        $rows = Get-HuntressIncidentReports -OrganizationId $OrganizationId -UpdatedSince $UpdatedSince
        if ($rows.Count -gt 0) {
            [void]$files.Add((Export-SecurityIntegrationCsv -Folder $ExportFolder -BaseName 'HuntressIncidents' -Rows $rows -TicketNumber $TicketNumber))
        }
    }
    if ($Selections.agents) {
        $rows = Get-HuntressAgents -OrganizationId $OrganizationId
        if ($rows.Count -gt 0) {
            [void]$files.Add((Export-SecurityIntegrationCsv -Folder $ExportFolder -BaseName 'HuntressAgents' -Rows $rows -TicketNumber $TicketNumber))
        }
    }
    if ($Selections.identities) {
        $rows = Get-HuntressIdentities -OrganizationId $OrganizationId
        if ($rows.Count -gt 0) {
            [void]$files.Add((Export-SecurityIntegrationCsv -Folder $ExportFolder -BaseName 'HuntressIdentities' -Rows $rows -TicketNumber $TicketNumber))
        }
    }
    if ($Selections.escalations) {
        $rows = Get-HuntressEscalations -OrganizationId $OrganizationId
        if ($rows.Count -gt 0) {
            [void]$files.Add((Export-SecurityIntegrationCsv -Folder $ExportFolder -BaseName 'HuntressEscalations' -Rows $rows -TicketNumber $TicketNumber))
        }
    }

    try {
        $org = Invoke-HuntressApi -Method GET -RelativePath "organizations/$OrganizationId"
        [void]$files.Add((Export-SecurityIntegrationJson -Folder $ExportFolder -BaseName 'HuntressOrgSummary' -Object $org -TicketNumber $TicketNumber))
    } catch {}

    return @($files)
}

Export-ModuleMember -Function Get-HuntressConfigurationStatus,Invoke-HuntressApi,Find-HuntressOrganizationByName,Get-HuntressPreviewCounts,Export-HuntressInvestigation,Get-HuntressSignals,Get-HuntressIncidentReports,Get-HuntressAgents,Get-HuntressIdentities,Get-HuntressEscalations,Get-HuntressPagedItems
