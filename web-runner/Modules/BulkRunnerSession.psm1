$script:BulkRunnerHiddenWorkers = $true

function Set-BulkRunnerWorkerVisibility {
    param([bool]$Hidden = $true)
    $script:BulkRunnerHiddenWorkers = $Hidden
}

function Get-BulkRunnerDefaultReportSelections {
    # Align with BEC investigation preset (Settings Get-BecExportPresetSelections).
    $bec = $null
    try {
        $settingsMod = Join-Path $PSScriptRoot '..\..\Modules\Settings.psm1'
        if (Test-Path -LiteralPath $settingsMod) {
            Import-Module $settingsMod -Force -ErrorAction SilentlyContinue
            if (Get-Command Get-BecExportPresetSelections -ErrorAction SilentlyContinue) {
                $bec = Get-BecExportPresetSelections
            }
        }
    } catch { }
    if (-not $bec) {
        $bec = @{
            IncludeMessageTrace                  = $true
            IncludeInboxRules                    = $true
            IncludeTransportRules                = $true
            IncludeMailFlowConnectors            = $false
            IncludeMailboxForwarding             = $true
            IncludeUnifiedAuditLogs              = $true
            IncludeAuditLogs                     = $true
            IncludeSignInLogs                    = $true
            IncludeMfaCoverage                   = $true
            IncludeConditionalAccessPolicies     = $true
            IncludeAppRegistrations              = $true
            IncludeIntuneDevices                 = $true
            IncludeSharePointActivity            = $false
            IncludeOneDriveActivity              = $false
            IncludeTeamsActivity                 = $false
            IncludeSharePointSharing             = $false
            IncludeSecurityAlerts                = $true
            IncludeSecurityIncidents             = $true
            IncludeDLPViolations                 = $false
            IncludeAnonymousSharePointSharing    = $false
            IncludeSharePointFileSharingLinks    = $false
            IncludeSharePointOneDriveFileActions = $false
        }
    }
    $bec['SignInLogsDaysBack'] = 7
    $bec['MessageTraceDaysBack'] = 7
    return $bec
}

function ConvertTo-BulkRunnerReportSelectionsHashtable {
    param($JsonObject)

    if (-not $JsonObject) { return Get-BulkRunnerDefaultReportSelections }

    return @{
        IncludeMessageTrace              = if ($null -ne $JsonObject.IncludeMessageTrace) { [bool]$JsonObject.IncludeMessageTrace } else { $false }
        IncludeInboxRules                = if ($null -ne $JsonObject.IncludeInboxRules) { [bool]$JsonObject.IncludeInboxRules } else { $false }
        IncludeTransportRules            = if ($null -ne $JsonObject.IncludeTransportRules) { [bool]$JsonObject.IncludeTransportRules } else { $false }
        IncludeMailFlowConnectors        = if ($null -ne $JsonObject.IncludeMailFlowConnectors) { [bool]$JsonObject.IncludeMailFlowConnectors } else { $false }
        IncludeMailboxForwarding         = if ($null -ne $JsonObject.IncludeMailboxForwarding) { [bool]$JsonObject.IncludeMailboxForwarding } else { $false }
        IncludeAuditLogs                 = if ($null -ne $JsonObject.IncludeAuditLogs) { [bool]$JsonObject.IncludeAuditLogs } else { $false }
        IncludeConditionalAccessPolicies = if ($null -ne $JsonObject.IncludeConditionalAccessPolicies) { [bool]$JsonObject.IncludeConditionalAccessPolicies } else { $false }
        IncludeAppRegistrations          = if ($null -ne $JsonObject.IncludeAppRegistrations) { [bool]$JsonObject.IncludeAppRegistrations } else { $false }
        IncludeSignInLogs                = (($JsonObject.IncludeSignInLogs -eq $true) -or ("$($JsonObject.IncludeSignInLogs)" -match '^(?i)true$|^(?i)yes$|^1$'))
        IncludeIntuneDevices             = if ($null -ne $JsonObject.IncludeIntuneDevices -and "$($JsonObject.IncludeIntuneDevices)" -ne '') { [bool]$JsonObject.IncludeIntuneDevices } else { $false }
        IncludeMfaCoverage               = if ($null -ne $JsonObject.IncludeMfaCoverage -and "$($JsonObject.IncludeMfaCoverage)" -ne '') { [bool]$JsonObject.IncludeMfaCoverage } else { $false }
        IncludeSharePointActivity        = if ($null -ne $JsonObject.IncludeSharePointActivity) { [bool]$JsonObject.IncludeSharePointActivity } else { $false }
        IncludeOneDriveActivity          = if ($null -ne $JsonObject.IncludeOneDriveActivity) { [bool]$JsonObject.IncludeOneDriveActivity } else { $false }
        IncludeTeamsActivity             = if ($null -ne $JsonObject.IncludeTeamsActivity) { [bool]$JsonObject.IncludeTeamsActivity } else { $false }
        IncludeSharePointSharing         = if ($null -ne $JsonObject.IncludeSharePointSharing) { [bool]$JsonObject.IncludeSharePointSharing } else { $false }
        IncludeSecurityAlerts            = if ($null -ne $JsonObject.IncludeSecurityAlerts) { [bool]$JsonObject.IncludeSecurityAlerts } else { $false }
        IncludeSecurityIncidents         = if ($null -ne $JsonObject.IncludeSecurityIncidents) { [bool]$JsonObject.IncludeSecurityIncidents } else { $false }
        IncludeUnifiedAuditLogs          = if ($null -ne $JsonObject.IncludeUnifiedAuditLogs) { [bool]$JsonObject.IncludeUnifiedAuditLogs } else { $false }
        IncludeDLPViolations             = if ($null -ne $JsonObject.IncludeDLPViolations) { [bool]$JsonObject.IncludeDLPViolations } else { $false }
        IncludeAnonymousSharePointSharing = if ($null -ne $JsonObject.IncludeAnonymousSharePointSharing) { [bool]$JsonObject.IncludeAnonymousSharePointSharing } else { $false }
        IncludeSharePointFileSharingLinks = if ($null -ne $JsonObject.IncludeSharePointFileSharingLinks) { [bool]$JsonObject.IncludeSharePointFileSharingLinks } else { $false }
        IncludeSharePointOneDriveFileActions = if ($null -ne $JsonObject.IncludeSharePointOneDriveFileActions) { [bool]$JsonObject.IncludeSharePointOneDriveFileActions } else { $false }
        SignInLogsDaysBack               = if ($null -ne $JsonObject.SignInLogsDaysBack) { [int]$JsonObject.SignInLogsDaysBack } else { 7 }
        MessageTraceDaysBack             = if ($null -ne $JsonObject.MessageTraceDaysBack) { [int]$JsonObject.MessageTraceDaysBack } else { 7 }
    }
}

function Merge-BulkRunnerReportSelections {
    param(
        [hashtable]$Base,
        [hashtable]$Override
    )

    if (-not $Base) { $Base = Get-BulkRunnerDefaultReportSelections }
    if (-not $Override -or $Override.Count -eq 0) { return $Base }

    $merged = @{}
    foreach ($k in $Base.Keys) { $merged[$k] = $Base[$k] }
    foreach ($k in $Override.Keys) { $merged[$k] = $Override[$k] }
    return $merged
}

function Get-BulkRunnerHistoryDirectory {
    $root = Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'ExchangeOnlineAnalyzer\BulkWebRunner\history'
    if (-not (Test-Path -LiteralPath $root)) {
        $null = New-Item -ItemType Directory -Path $root -Force -ErrorAction Stop
    }
    return $root
}

function Get-BulkRunnerArchiveDirectory {
    $root = Join-Path (Get-BulkRunnerHistoryDirectory) 'archive'
    if (-not (Test-Path -LiteralPath $root)) {
        $null = New-Item -ItemType Directory -Path $root -Force -ErrorAction Stop
    }
    return $root
}

function Get-BulkRunnerSessionManifestPath {
    param([Parameter(Mandatory = $true)][string]$SessionId)
    return Join-Path (Get-BulkRunnerHistoryDirectory) "$SessionId.json"
}

function Get-BulkRunnerArchivedSessionManifestPath {
    param([Parameter(Mandatory = $true)][string]$SessionId)
    return Join-Path (Get-BulkRunnerArchiveDirectory) "$SessionId.json"
}

function Resolve-BulkRunnerSessionManifestPath {
    param([Parameter(Mandatory = $true)][string]$SessionId)

    $active = Get-BulkRunnerSessionManifestPath -SessionId $SessionId
    if (Test-Path -LiteralPath $active) { return $active }

    $archived = Get-BulkRunnerArchivedSessionManifestPath -SessionId $SessionId
    if (Test-Path -LiteralPath $archived) { return $archived }

    return $null
}

function Update-BulkRunnerSessionManifestMetadata {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][hashtable]$Fields
    )

    $m = Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json
    $ordered = [ordered]@{}
    foreach ($p in $m.PSObject.Properties) {
        if ($Fields.ContainsKey($p.Name)) { continue }
        $ordered[$p.Name] = $p.Value
    }
    foreach ($k in $Fields.Keys) {
        $ordered[$k] = $Fields[$k]
    }
    $ordered | ConvertTo-Json -Depth 8 | Set-Content -Path $Path -Encoding UTF8
}

function ConvertTo-BulkRunnerSessionHistoryRow {
    param(
        [Parameter(Mandatory = $true)]
        $Manifest,

        [switch]$Archived
    )

    if (-not $Archived -and -not (Test-BulkRunnerSessionHistoryVisible -Manifest $Manifest)) {
        return $null
    }

    $tickets = @()
    if ($Manifest.ticketNumbers) {
        $tickets = @($Manifest.ticketNumbers | ForEach-Object { "$_".Trim() } | Where-Object { $_ })
    }
    if ($Manifest.tenants) {
        $tickets += @($Manifest.tenants | ForEach-Object {
            if ($_.uiState -and $_.uiState.ticket) { "$($_.uiState.ticket)".Trim() }
        } | Where-Object { $_ })
    }
    $tickets = @($tickets | Select-Object -Unique)

    $orgs = @($Manifest.tenants | ForEach-Object {
        if ($_.exoOrganizationName) { [string]$_.exoOrganizationName }
        elseif ($_.graphTenantName) { [string]$_.graphTenantName }
        elseif ($_.uiState -and $_.uiState.organizationHint) { [string]$_.uiState.organizationHint }
    } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)

    $clients = @()
    if ($Manifest.tenants) {
        $clients = @($Manifest.tenants | ForEach-Object {
            $org = $null
            if ($_.exoOrganizationName) { $org = [string]$_.exoOrganizationName }
            elseif ($_.graphTenantName) { $org = [string]$_.graphTenantName }
            elseif ($_.uiState -and $_.uiState.organizationHint) { $org = [string]$_.uiState.organizationHint }
            $ticket = $null
            if ($_.uiState -and $_.uiState.ticket) { $ticket = "$($_.uiState.ticket)".Trim() }
            [pscustomobject]@{
                clientNumber = [int]$_.clientNumber
                organization = $org
                ticket       = $ticket
            }
        } | Sort-Object clientNumber)
    }

    $clientsLabel = ($clients | ForEach-Object {
        $parts = @("Client $($_.clientNumber)")
        if ($_.organization) { [void]$parts.Add($_.organization) }
        if ($_.ticket) { [void]$parts.Add("#$($_.ticket)") }
        $parts -join ' · '
    }) -join '; '

    return [pscustomobject]@{
        sessionId     = [string]$Manifest.sessionId
        createdAt     = [string]$Manifest.createdAt
        updatedAt     = [string]$Manifest.updatedAt
        archivedAt    = [string]$Manifest.archivedAt
        tenantCount   = [int]($Manifest.tenantCount | ForEach-Object { if ($null -ne $_) { $_ } else { 0 } })
        ticketNumbers = $tickets
        organizations = $orgs
        clients       = $clients
        clientsLabel  = $clientsLabel
        status        = [string]$Manifest.status
        archived      = [bool]$Archived
        hasOutputs    = [bool](@($Manifest.tenants | Where-Object { $_.outputFolder }).Count -gt 0)
    }
}

function Sort-BulkRunnerSessionHistoryRows {
    param(
        [array]$Rows = @(),

        [string]$SortBy = 'updatedAt',
        [ValidateSet('asc', 'desc')]
        [string]$SortOrder = 'desc'
    )

    if ($null -eq $Rows -or $Rows.Count -lt 1) { return @() }

    $desc = ($SortOrder -eq 'desc')
    $key = $SortBy.ToLowerInvariant()

    $expr = switch ($key) {
        'createdat' { { try { [datetime]::Parse($_.createdAt) } catch { [datetime]::MinValue } } }
        'organization' { { ($_.organizations -join '; ').ToLowerInvariant() } }
        'clients' { { if ($_.clientsLabel) { $_.clientsLabel.ToLowerInvariant() } else { ($_.organizations -join '; ').ToLowerInvariant() } } }
        'ticket' { { ($_.ticketNumbers -join ', ').ToLowerInvariant() } }
        'tenants' { { $_.tenantCount } }
        'sessionid' { { $_.sessionId.ToLowerInvariant() } }
        'archivedat' { { try { [datetime]::Parse($_.archivedAt) } catch { [datetime]::MinValue } } }
        default { { try { [datetime]::Parse($_.updatedAt) } catch { [datetime]::MinValue } } }
    }

    return @($Rows | Sort-Object -Property @{ Expression = $expr; Descending = $desc })
}

function Test-BulkRunnerSessionWorthSaving {
    param(
        [Parameter(Mandatory = $true)]
        $Session
    )

    if ($Session.Tenants.Count -gt 0) { return $true }

    foreach ($key in @($Session.TenantUiStates.Keys)) {
        $ui = $Session.TenantUiStates[$key]
        if (-not $ui) { continue }
        if (-not [string]::IsNullOrWhiteSpace([string]$ui.ticket)) { return $true }
        if (-not [string]::IsNullOrWhiteSpace([string]$ui.organizationHint)) { return $true }
        if (-not [string]::IsNullOrWhiteSpace([string]$ui.userSearch)) { return $true }
        if ($ui.dateStart -or $ui.dateEnd) { return $true }
        if ($ui.validatedUsers -and @($ui.validatedUsers).Count -gt 0) { return $true }
    }

    return $false
}

function Test-BulkRunnerSessionHistoryVisible {
    param(
        [Parameter(Mandatory = $true)]
        $Manifest
    )

    $tenantCount = 0
    if ($null -ne $Manifest.tenantCount) { $tenantCount = [int]$Manifest.tenantCount }

    $tickets = @()
    if ($Manifest.ticketNumbers) {
        $tickets = @($Manifest.ticketNumbers | ForEach-Object { "$_".Trim() } | Where-Object { $_ })
    }
    if ($Manifest.tenants) {
        $tickets += @($Manifest.tenants | ForEach-Object {
            if ($_.uiState -and $_.uiState.ticket) { "$($_.uiState.ticket)".Trim() }
        } | Where-Object { $_ })
    }
    $tickets = @($tickets | Select-Object -Unique)

    $hasAuthOrOutput = $false
    $hasUiData = $false
    if ($Manifest.tenants) {
        foreach ($t in @($Manifest.tenants)) {
            if ($t.graphAuthenticated -or $t.exchangeAuthenticated -or $t.outputFolder) {
                $hasAuthOrOutput = $true
            }
            if ($t.uiState) {
                $ui = $t.uiState
                if (-not [string]::IsNullOrWhiteSpace([string]$ui.ticket)) { $hasUiData = $true }
                if (-not [string]::IsNullOrWhiteSpace([string]$ui.organizationHint)) { $hasUiData = $true }
                if ($ui.dateStart -or $ui.dateEnd) { $hasUiData = $true }
                if ($ui.userSearch -and -not [string]::IsNullOrWhiteSpace([string]$ui.userSearch)) { $hasUiData = $true }
                if ($ui.validatedUsers -and @($ui.validatedUsers).Count -gt 0) { $hasUiData = $true }
            }
        }
    }

    if ($tenantCount -eq 0 -and $tickets.Count -eq 0) { return $false }
    if ([string]$Manifest.status -eq 'replaced' -and $tenantCount -eq 0) { return $false }
    if ([string]$Manifest.status -eq 'replaced' -and $tickets.Count -eq 0 -and -not $hasAuthOrOutput -and -not $hasUiData) {
        return $false
    }

    return $true
}

function Sync-BulkRunnerSessionManifest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Session,

        [switch]$Force
    )

    if (-not $Force -and -not (Test-BulkRunnerSessionWorthSaving -Session $Session)) {
        return $null
    }

    $last = $null
    if ($Session.PSObject.Properties['LastManifestSaveAt']) { $last = $Session.LastManifestSaveAt }
    if (-not $Force -and $last -and (((Get-Date) - $last).TotalSeconds -lt 5)) {
        return $null
    }

    $path = Save-BulkRunnerSessionManifest -Session $Session
    $Session | Add-Member -NotePropertyName LastManifestSaveAt -NotePropertyValue (Get-Date) -Force
    return $path
}

function Set-BulkRunnerSessionReportSelections {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Session,

        [Parameter(Mandatory = $true)]
        [hashtable]$ReportSelections,

        [int]$DaysBack = -1
    )

    $defaults = Get-BulkRunnerDefaultReportSelections
    $merged = Merge-BulkRunnerReportSelections -Base $defaults -Override $ReportSelections
    $Session.ReportSelections = $merged
    $merged | ConvertTo-Json -Depth 6 | Set-Content -Path $Session.ReportSelectionsFile -Encoding UTF8
    if ($DaysBack -ge 1) { $Session.DaysBack = $DaysBack }
}

function Get-BulkRunnerTenantUiProperty {
    param(
        $Ui,
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    if (-not $Ui) { return $null }
    if ($Ui -is [System.Collections.IDictionary]) {
        if ($Ui.Contains($Name)) { return $Ui[$Name] }
        return $null
    }
    $prop = $Ui.PSObject.Properties[$Name]
    if ($prop) { return $prop.Value }
    return $null
}

function Test-BulkRunnerTenantUsesSessionReportDefaults {
    param(
        [Parameter(Mandatory = $true)]
        $Session,
        [Parameter(Mandatory = $true)]
        [int]$ClientNumber
    )

    $key = [string]$ClientNumber
    if (-not $Session.TenantUiStates.ContainsKey($key)) { return $true }
    $flag = Get-BulkRunnerTenantUiProperty -Ui $Session.TenantUiStates[$key] -Name 'useSessionReportDefaults'
    if ($null -eq $flag) { return $true }
    return ($flag -eq $true -or "$flag" -eq 'true')
}

function Get-BulkRunnerTenantReportSelectionsOverride {
    param(
        [Parameter(Mandatory = $true)]
        $Session,
        [Parameter(Mandatory = $true)]
        [int]$ClientNumber
    )

    $key = [string]$ClientNumber
    if (-not $Session.TenantUiStates.ContainsKey($key)) { return $null }
    $ui = $Session.TenantUiStates[$key]
    if ($null -eq $ui) { return $null }

    if (Test-BulkRunnerTenantUsesSessionReportDefaults -Session $Session -ClientNumber $ClientNumber) { return $null }

    $rs = Get-BulkRunnerTenantUiProperty -Ui $ui -Name 'reportSelections'
    if (-not $rs) { return $null }

    $override = @{}
    if ($rs -is [hashtable]) {
        foreach ($k in $rs.Keys) { $override[$k] = $rs[$k] }
    } else {
        $rs.PSObject.Properties | ForEach-Object { $override[$_.Name] = $_.Value }
    }
    if ($override.Count -eq 0) { return $null }
    return $override
}

function Get-BulkRunnerEffectiveReportSelections {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Session,

        [Parameter(Mandatory = $false)]
        [int]$ClientNumber = 0
    )

    $base = if ($Session.ReportSelections) { $Session.ReportSelections } else { Get-BulkRunnerDefaultReportSelections }
    if ($ClientNumber -lt 1) { return $base }
    $override = Get-BulkRunnerTenantReportSelectionsOverride -Session $Session -ClientNumber $ClientNumber
    if (-not $override) { return $base }
    return Merge-BulkRunnerReportSelections -Base $base -Override $override
}

function Write-BulkRunnerTenantReportSelectionsFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Session,

        [Parameter(Mandatory = $true)]
        [int]$ClientNumber
    )

    $effective = Get-BulkRunnerEffectiveReportSelections -Session $Session -ClientNumber $ClientNumber
    $path = Join-Path $Session.TempDir "ReportSelections_Client${ClientNumber}.json"
    $effective | ConvertTo-Json -Depth 6 | Set-Content -Path $path -Encoding UTF8
    return $path
}

function Expand-BulkRunnerGenerateReportsCommand {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Session,

        [Parameter(Mandatory = $true)]
        [int]$ClientNumber,

        [Parameter(Mandatory = $true)]
        [string]$Command
    )

    if ($Command -notmatch '^GENERATE_REPORTS') { return $Command }
    $rsFile = Write-BulkRunnerTenantReportSelectionsFile -Session $Session -ClientNumber $ClientNumber
    return "$Command|ReportSelectionsFile:$rsFile"
}

function Sync-BulkRunnerSessionManifestIfMissing {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Session
    )

    $path = Get-BulkRunnerSessionManifestPath -SessionId $Session.SessionId
    if (Test-Path -LiteralPath $path) {
        Sync-BulkRunnerSessionManifest -Session $Session | Out-Null
    } elseif (Test-BulkRunnerSessionWorthSaving -Session $Session) {
        Save-BulkRunnerSessionManifest -Session $Session | Out-Null
        $Session | Add-Member -NotePropertyName LastManifestSaveAt -NotePropertyValue (Get-Date) -Force
    }
}

function Get-BulkRunnerDefaultSettings {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProjectRoot
    )

    $investigator = 'Security Administrator'
    $company = 'Organization'
    $settingsPath = Join-Path $ProjectRoot 'settings.json'
    if (Test-Path -LiteralPath $settingsPath) {
        try {
            $settings = Get-Content -LiteralPath $settingsPath -Raw | ConvertFrom-Json
            if ($settings.InvestigatorName) { $investigator = [string]$settings.InvestigatorName }
            if ($settings.CompanyName) { $company = [string]$settings.CompanyName }
        } catch { }
    }
    return [pscustomobject]@{ InvestigatorName = $investigator; CompanyName = $company }
}

function New-BulkRunnerSession {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProjectRoot,

        [Parameter(Mandatory = $true)]
        [hashtable]$ReportSelections,

        [string]$InvestigatorName = '',
        [string]$CompanyName = '',
        [int]$DaysBack = 7
    )

    $workerTemplate = Join-Path $ProjectRoot 'Scripts\BulkExportWorker.ps1'
    if (-not (Test-Path -LiteralPath $workerTemplate)) {
        throw "Missing worker script: $workerTemplate"
    }

    $tempDir = Join-Path $env:TEMP "EOA_BulkWeb_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
    $null = New-Item -ItemType Directory -Path $tempDir -Force -ErrorAction Stop
    $commandDir = Join-Path $tempDir 'commands'
    $null = New-Item -ItemType Directory -Path $commandDir -Force -ErrorAction Stop

    $reportSelectionsFile = Join-Path $tempDir 'ReportSelections.json'
    $ReportSelections | ConvertTo-Json -Depth 6 | Set-Content -Path $reportSelectionsFile -Encoding UTF8

    $workerScriptFile = Join-Path $tempDir 'BulkTenantWorker.ps1'
    Copy-Item -LiteralPath $workerTemplate -Destination $workerScriptFile -Force

    return [pscustomobject]@{
        SessionId            = [System.IO.Path]::GetFileName($tempDir)
        ProjectRoot          = $ProjectRoot
        WorkerTemplatePath   = $workerTemplate
        TempDir              = $tempDir
        CommandDir           = $commandDir
        ReportSelectionsFile = $reportSelectionsFile
        ReportSelections     = $ReportSelections
        WorkerScriptFile     = $workerScriptFile
        InvestigatorName     = $InvestigatorName
        CompanyName          = $CompanyName
        DaysBack             = $DaysBack
        Tenants              = @{}
        TenantUiStates       = @{}
        NextClientNumber     = 1
        CreatedAt            = Get-Date
    }
}

function Get-BulkRunnerAuthDetailsFromResponse {
    param([string]$Response)

    $details = @{
        GraphTenantId       = $null
        GraphTenantName     = $null
        ExoTenantId         = $null
        ExoOrganizationName = $null
    }
    if ([string]::IsNullOrWhiteSpace($Response)) { return $details }

    if ($Response -like 'GRAPH_AUTH_SUCCESS*') {
        if ($Response -match 'GRAPH_AUTH_SUCCESS:([^|]+)') {
            $details.GraphTenantName = $Matches[1].Trim()
        }
        if ($Response -match 'TENANT_ID:([a-fA-F0-9\-]{36})') {
            $details.GraphTenantId = $Matches[1]
        }
    }
    if ($Response -like 'EXCHANGE_AUTH_SUCCESS*') {
        if ($Response -match 'TENANT_ID:([a-fA-F0-9\-]{36})') {
            $details.ExoTenantId = $Matches[1]
        }
        if ($Response -match '\|ORG:([^|]+)') {
            $details.ExoOrganizationName = $Matches[1].Trim()
        }
    }
    return $details
}

function Sync-BulkRunnerWorkerScript {
    param(
        [Parameter(Mandatory = $true)]
        $Session
    )

    $template = if ($Session.WorkerTemplatePath) {
        $Session.WorkerTemplatePath
    } else {
        Join-Path $Session.ProjectRoot 'Scripts\BulkExportWorker.ps1'
    }
    if (-not (Test-Path -LiteralPath $template)) {
        throw "Missing worker script: $template"
    }
    Copy-Item -LiteralPath $template -Destination $Session.WorkerScriptFile -Force
}

function Start-BulkRunnerTenantProcess {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Session,

        [Parameter(Mandatory = $true)]
        [int]$ClientNumber,

        [ValidateSet('Normal', 'Hidden', 'Minimized')]
        [string]$WindowStyle
    )

    $statusFile = Join-Path $Session.TempDir "Client${ClientNumber}_Status.txt"
    $resultFile = Join-Path $Session.TempDir "Client${ClientNumber}_Result.txt"

    $investigatorName = if ([string]::IsNullOrWhiteSpace($Session.InvestigatorName)) { 'Security Administrator' } else { $Session.InvestigatorName.Trim() }
    $companyName = if ([string]::IsNullOrWhiteSpace($Session.CompanyName)) { 'Organization' } else { $Session.CompanyName.Trim() }

    Sync-BulkRunnerWorkerScript -Session $Session

    $psExe = $null
    $pwshCmd = Get-Command pwsh.exe -ErrorAction SilentlyContinue
    if ($pwshCmd -and -not [string]::IsNullOrWhiteSpace($pwshCmd.Source)) {
        $psExe = $pwshCmd.Source
    }
    if (-not $psExe) {
        $psExe = (Get-Process -Id $PID -ErrorAction Stop).Path
    }
    $argList = @(
        '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $Session.WorkerScriptFile,
        '-ClientNumber', [string]$ClientNumber,
        '-ScriptRoot', $Session.ProjectRoot,
        '-InvestigatorName', $investigatorName,
        '-CompanyName', $companyName,
        '-DaysBack', [string]$Session.DaysBack,
        '-ReportSelectionsFile', $Session.ReportSelectionsFile,
        '-StatusFile', $statusFile,
        '-ResultFile', $resultFile,
        '-CommandDir', $Session.CommandDir
    )

    $procStyle = if (-not [string]::IsNullOrWhiteSpace($WindowStyle)) {
        $WindowStyle
    } elseif ($script:BulkRunnerHiddenWorkers) {
        'Hidden'
    } else {
        'Normal'
    }
    $proc = Start-Process -FilePath $psExe -ArgumentList $argList -PassThru -WindowStyle $procStyle `
        -WorkingDirectory $Session.ProjectRoot -ErrorAction Stop

    return [pscustomobject]@{
        ClientNumber          = $ClientNumber
        ProcessId             = $proc.Id
        StatusFile            = $statusFile
        ResultFile            = $resultFile
        GraphAuthenticated    = $false
        ExchangeAuthenticated = $false
        GraphTenantId         = $null
        GraphTenantName       = $null
        ExoTenantId           = $null
        ExoOrganizationName   = $null
        LastResponse          = $null
        OutputFolder          = $null
        ReportInProgress      = $false
        CommandSentAt         = $null
    }
}

function Add-BulkRunnerTenant {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Session
    )

    $clientNumber = $Session.NextClientNumber
    $tenant = Start-BulkRunnerTenantProcess -Session $Session -ClientNumber $clientNumber
    $Session.NextClientNumber = $clientNumber + 1
    $Session.Tenants[[string]$clientNumber] = $tenant
    return $tenant
}

function Restart-BulkRunnerTenant {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Session,

        [Parameter(Mandatory = $true)]
        [int]$ClientNumber,

        [switch]$ShowConsole
    )

    if (-not $Session.Tenants.ContainsKey([string]$ClientNumber)) {
        throw "Unknown tenant client number: $ClientNumber"
    }

    $existing = $Session.Tenants[[string]$ClientNumber]
    if ($existing.ProcessId) {
        Stop-Process -Id $existing.ProcessId -Force -ErrorAction SilentlyContinue
        Start-Sleep -Milliseconds 500
    }

    foreach ($pattern in @(
            "Client${ClientNumber}_Command.txt"
            "Client${ClientNumber}_Response.txt"
        )) {
        $f = Join-Path $Session.CommandDir $pattern
        if (Test-Path -LiteralPath $f) {
            Remove-Item -LiteralPath $f -Force -ErrorAction SilentlyContinue
        }
    }

    if ($ShowConsole) {
        $tenant = Start-BulkRunnerTenantProcess -Session $Session -ClientNumber $ClientNumber -WindowStyle Normal
    } else {
        $tenant = Start-BulkRunnerTenantProcess -Session $Session -ClientNumber $ClientNumber
    }
    $Session.Tenants[[string]$ClientNumber] = $tenant
    return $tenant
}

function Stop-BulkRunnerTenantProcess {
    param(
        [Parameter(Mandatory = $true)]
        $Tenant
    )
    if ($Tenant.ProcessId) {
        Stop-Process -Id $Tenant.ProcessId -Force -ErrorAction SilentlyContinue
        Start-Sleep -Milliseconds 300
    }
}

function Remove-BulkRunnerTenant {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Session,

        [Parameter(Mandatory = $true)]
        [int]$ClientNumber
    )

    $key = [string]$ClientNumber
    if (-not $Session.Tenants.ContainsKey($key)) {
        throw "Unknown tenant client number: $ClientNumber"
    }

    $tenant = $Session.Tenants[$key]
    Stop-BulkRunnerTenantProcess -Tenant $tenant

    foreach ($pattern in @(
            "Client${ClientNumber}_Command.txt"
            "Client${ClientNumber}_Response.txt"
        )) {
        $f = Join-Path $Session.CommandDir $pattern
        if (Test-Path -LiteralPath $f) { Remove-Item -LiteralPath $f -Force -ErrorAction SilentlyContinue }
    }

    [void]$Session.Tenants.Remove($key)
    if ($Session.TenantUiStates.ContainsKey($key)) {
        [void]$Session.TenantUiStates.Remove($key)
    }
}

function Stop-BulkRunnerSessionTenants {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Session
    )
    foreach ($key in @($Session.Tenants.Keys)) {
        Stop-BulkRunnerTenantProcess -Tenant $Session.Tenants[$key]
    }
}

function Set-BulkRunnerTenantUiState {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Session,

        [Parameter(Mandatory = $true)]
        [int]$ClientNumber,

        [Parameter(Mandatory = $true)]
        [hashtable]$UiState
    )

    $key = [string]$ClientNumber
    if (-not $Session.Tenants.ContainsKey($key)) {
        throw "Unknown tenant client number: $ClientNumber"
    }
    if ($UiState.ContainsKey('ticket') -and $null -ne $UiState.ticket) {
        $UiState.ticket = [string]$UiState.ticket.Trim()
    }
    if ($UiState.ContainsKey('organizationHint') -and $null -ne $UiState.organizationHint) {
        $UiState.organizationHint = [string]$UiState.organizationHint.Trim()
    }
    $Session.TenantUiStates[$key] = $UiState
}

function Save-BulkRunnerSessionManifest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Session,

        [string]$Status = 'active'
    )

    Sync-BulkRunnerSessionResponsesFromFiles -Session $Session

    $tenantRows = @($Session.Tenants.Values | Sort-Object ClientNumber | ForEach-Object {
        $key = [string]$_.ClientNumber
        $ui = $null
        if ($Session.TenantUiStates.ContainsKey($key)) { $ui = $Session.TenantUiStates[$key] }
        [ordered]@{
            clientNumber          = $_.ClientNumber
            processId             = $_.ProcessId
            graphAuthenticated    = $_.GraphAuthenticated
            exchangeAuthenticated = $_.ExchangeAuthenticated
            graphTenantId         = $_.GraphTenantId
            graphTenantName       = $_.GraphTenantName
            exoTenantId           = $_.ExoTenantId
            exoOrganizationName   = $_.ExoOrganizationName
            lastResponse          = $_.LastResponse
            outputFolder          = $_.OutputFolder
            reportInProgress      = $_.ReportInProgress
            uiState               = $ui
        }
    })

    $manifest = [ordered]@{
        sessionId        = $Session.SessionId
        createdAt        = if ($Session.CreatedAt) { $Session.CreatedAt.ToString('o') } else { (Get-Date).ToString('o') }
        updatedAt        = (Get-Date).ToString('o')
        tempDir          = $Session.TempDir
        investigatorName = $Session.InvestigatorName
        companyName      = $Session.CompanyName
        daysBack         = $Session.DaysBack
        reportSelections = $Session.ReportSelections
        status           = $Status
        tenantCount      = $tenantRows.Count
        tenants          = $tenantRows
        ticketNumbers    = @($tenantRows | ForEach-Object {
            if ($_.uiState -and $_.uiState.ticket) { "$($_.uiState.ticket)".Trim() }
        } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
    }

    $path = Get-BulkRunnerSessionManifestPath -SessionId $Session.SessionId
    $manifest | ConvertTo-Json -Depth 8 | Set-Content -Path $path -Encoding UTF8
    return $path
}

function Get-BulkRunnerSessionHistory {
    [CmdletBinding()]
    param(
        [int]$Limit = 50,
        [switch]$Archived,
        [string]$SortBy = 'updatedAt',
        [ValidateSet('asc', 'desc')]
        [string]$SortOrder = 'desc'
    )

    $dir = if ($Archived) { Get-BulkRunnerArchiveDirectory } else { Get-BulkRunnerHistoryDirectory }
    $files = Get-ChildItem -LiteralPath $dir -Filter 'EOA_BulkWeb_*.json' -File -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending

    $rows = [System.Collections.ArrayList]::new()
    foreach ($f in $files) {
        try {
            $m = Get-Content -LiteralPath $f.FullName -Raw | ConvertFrom-Json
            $row = ConvertTo-BulkRunnerSessionHistoryRow -Manifest $m -Archived:$Archived
            if ($row) { [void]$rows.Add($row) }
        } catch { }
    }

    $sorted = Sort-BulkRunnerSessionHistoryRows -Rows @($rows) -SortBy $SortBy -SortOrder $SortOrder
    if ($Limit -gt 0 -and $sorted.Count -gt $Limit) {
        return @($sorted | Select-Object -First $Limit)
    }
    return @($sorted)
}

function Archive-BulkRunnerSessionHistory {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$SessionId
    )

    $src = Get-BulkRunnerSessionManifestPath -SessionId $SessionId
    if (-not (Test-Path -LiteralPath $src)) {
        throw "Session not found in saved history: $SessionId"
    }

    $dest = Get-BulkRunnerArchivedSessionManifestPath -SessionId $SessionId
    if (Test-Path -LiteralPath $dest) {
        Remove-Item -LiteralPath $dest -Force -ErrorAction Stop
    }

    Move-Item -LiteralPath $src -Destination $dest -Force -ErrorAction Stop
    Update-BulkRunnerSessionManifestMetadata -Path $dest -Fields @{
        status     = 'archived'
        archivedAt = (Get-Date).ToString('o')
    }
    return $dest
}

function Unarchive-BulkRunnerSessionHistory {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$SessionId
    )

    $src = Get-BulkRunnerArchivedSessionManifestPath -SessionId $SessionId
    if (-not (Test-Path -LiteralPath $src)) {
        throw "Session not found in archive: $SessionId"
    }

    $dest = Get-BulkRunnerSessionManifestPath -SessionId $SessionId
    if (Test-Path -LiteralPath $dest) {
        throw "Session already exists in saved history: $SessionId"
    }

    Move-Item -LiteralPath $src -Destination $dest -Force -ErrorAction Stop
    Update-BulkRunnerSessionManifestMetadata -Path $dest -Fields @{
        status     = 'active'
        archivedAt = $null
    }
    return $dest
}

function Remove-BulkRunnerSessionHistory {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$SessionId
    )

    $path = Resolve-BulkRunnerSessionManifestPath -SessionId $SessionId
    if (-not $path) {
        throw "Session history not found: $SessionId"
    }

    Remove-Item -LiteralPath $path -Force -ErrorAction Stop
    return @{ removed = $SessionId; path = $path }
}

function Get-BulkRunnerSessionManifest {
    param([Parameter(Mandatory = $true)][string]$SessionId)
    $path = Resolve-BulkRunnerSessionManifestPath -SessionId $SessionId
    if (-not $path) {
        throw "Session history not found: $SessionId"
    }
    return (Get-Content -LiteralPath $path -Raw | ConvertFrom-Json)
}

function New-BulkRunnerSessionFromManifest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProjectRoot,

        [Parameter(Mandatory = $true)]
        $Manifest
    )

    $selections = @{}
    if ($Manifest.reportSelections) {
        $Manifest.reportSelections.PSObject.Properties | ForEach-Object { $selections[$_.Name] = $_.Value }
    }
    $days = 7
    if ($Manifest.daysBack) { $days = [int]$Manifest.daysBack }

    $investigatorName = if ($Manifest.investigatorName) { [string]$Manifest.investigatorName } else { '' }
    $companyName = if ($Manifest.companyName) { [string]$Manifest.companyName } else { '' }

    $reuseTempDir = $false
    $tempDir = $null
    if ($Manifest.tempDir -and (Test-Path -LiteralPath ([string]$Manifest.tempDir))) {
        $tempDir = [string]$Manifest.tempDir
        $reuseTempDir = $true
    }

    if ($reuseTempDir) {
        $commandDir = Join-Path $tempDir 'commands'
        if (-not (Test-Path -LiteralPath $commandDir)) {
            $null = New-Item -ItemType Directory -Path $commandDir -Force -ErrorAction Stop
        }
        $reportSelectionsFile = Join-Path $tempDir 'ReportSelections.json'
        $selections | ConvertTo-Json -Depth 6 | Set-Content -Path $reportSelectionsFile -Encoding UTF8
        $workerTemplate = Join-Path $ProjectRoot 'Scripts\BulkExportWorker.ps1'
        if (-not (Test-Path -LiteralPath $workerTemplate)) {
            throw "Missing worker script: $workerTemplate"
        }
        $workerScriptFile = Join-Path $tempDir 'BulkTenantWorker.ps1'
        Copy-Item -LiteralPath $workerTemplate -Destination $workerScriptFile -Force

        $createdAt = Get-Date
        if ($Manifest.createdAt) {
            try { $createdAt = [datetime]$Manifest.createdAt } catch { }
        }

        $session = [pscustomobject]@{
            SessionId            = if ($Manifest.sessionId) { [string]$Manifest.sessionId } else { [System.IO.Path]::GetFileName($tempDir) }
            ProjectRoot          = $ProjectRoot
            WorkerTemplatePath   = $workerTemplate
            TempDir              = $tempDir
            CommandDir           = $commandDir
            ReportSelectionsFile = $reportSelectionsFile
            ReportSelections     = $selections
            WorkerScriptFile     = $workerScriptFile
            InvestigatorName     = $investigatorName
            CompanyName          = $companyName
            DaysBack             = $days
            Tenants              = @{}
            TenantUiStates       = @{}
            NextClientNumber     = 1
            CreatedAt            = $createdAt
        }
    } else {
        $session = New-BulkRunnerSession -ProjectRoot $ProjectRoot `
            -ReportSelections $selections `
            -InvestigatorName $investigatorName `
            -CompanyName $companyName `
            -DaysBack $days
    }

    $snapshots = [System.Collections.ArrayList]::new()
    if ($Manifest.tenants) {
        foreach ($t in @($Manifest.tenants)) {
            $clientNumber = [int]$t.clientNumber
            if ($clientNumber -lt 1) { continue }

            $ui = $null
            if ($t.uiState) {
                $ui = @{}
                $t.uiState.PSObject.Properties | ForEach-Object { $ui[$_.Name] = $_.Value }
                $session.TenantUiStates[[string]$clientNumber] = $ui
            }

            $statusFile = Join-Path $session.TempDir "Client${clientNumber}_Status.txt"
            $resultFile = Join-Path $session.TempDir "Client${clientNumber}_Result.txt"
            $tenant = [pscustomobject]@{
                ClientNumber          = $clientNumber
                ProcessId             = if ($t.processId) { [int]$t.processId } else { $null }
                StatusFile            = $statusFile
                ResultFile            = $resultFile
                GraphAuthenticated    = ($t.graphAuthenticated -eq $true)
                ExchangeAuthenticated = ($t.exchangeAuthenticated -eq $true)
                GraphTenantId         = if ($t.graphTenantId) { [string]$t.graphTenantId } else { $null }
                GraphTenantName       = if ($t.graphTenantName) { [string]$t.graphTenantName } else { $null }
                ExoTenantId           = if ($t.exoTenantId) { [string]$t.exoTenantId } else { $null }
                ExoOrganizationName   = if ($t.exoOrganizationName) { [string]$t.exoOrganizationName } else { $null }
                LastResponse          = if ($t.lastResponse) { [string]$t.lastResponse } else { $null }
                OutputFolder          = if ($t.outputFolder) { [string]$t.outputFolder } else { $null }
                ReportInProgress      = ($t.reportInProgress -eq $true)
                CommandSentAt         = $null
            }

            if ($tenant.ProcessId -and -not (Test-BulkRunnerTenantWorkerAlive -Tenant $tenant)) {
                $tenant.ProcessId = $null
                if ($tenant.ReportInProgress) {
                    $tenant.ReportInProgress = $false
                }
            }

            $session.Tenants[[string]$clientNumber] = $tenant
            if ($clientNumber -ge $session.NextClientNumber) {
                $session.NextClientNumber = $clientNumber + 1
            }

            [void]$snapshots.Add([pscustomobject]@{
                    clientNumber          = $clientNumber
                    exoOrganizationName = [string]$t.exoOrganizationName
                    graphTenantName     = [string]$t.graphTenantName
                    outputFolder        = [string]$t.outputFolder
                    uiState             = $ui
                })
        }
    }

    if ($reuseTempDir) {
        try { Sync-BulkRunnerWorkerScript -Session $session } catch { }
    }

    return [pscustomobject]@{
        Session    = $session
        Snapshots  = @($snapshots)
        SourceId   = [string]$Manifest.sessionId
    }
}

function Update-BulkRunnerTenantFromResponse {
    param(
        [Parameter(Mandatory = $true)]
        $Tenant,

        [Parameter(Mandatory = $true)]
        [string]$Response
    )

    $Tenant.LastResponse = $Response
    $authDetails = Get-BulkRunnerAuthDetailsFromResponse -Response $Response
    if ($Response -like 'GRAPH_AUTH_SUCCESS*') {
        $Tenant.GraphAuthenticated = $true
        if ($authDetails.GraphTenantId) { $Tenant.GraphTenantId = $authDetails.GraphTenantId }
        if ($authDetails.GraphTenantName) { $Tenant.GraphTenantName = $authDetails.GraphTenantName }
    }
    if ($Response -eq 'EXCHANGE_AUTH_SUCCESS' -or $Response -like 'EXCHANGE_AUTH_SUCCESS*') {
        $Tenant.ExchangeAuthenticated = $true
        if ($authDetails.ExoTenantId) { $Tenant.ExoTenantId = $authDetails.ExoTenantId }
        if ($authDetails.ExoOrganizationName) { $Tenant.ExoOrganizationName = $authDetails.ExoOrganizationName }
    }
    if ($Response -like 'GENERATE_REPORTS_STARTED*') {
        $Tenant.ReportInProgress = Test-BulkRunnerTenantWorkerAlive -Tenant $Tenant
    }
    if ($Response -like 'GENERATE_REPORTS_SUCCESS:*') {
        $Tenant.OutputFolder = ($Response -replace '^GENERATE_REPORTS_SUCCESS:', '').Trim()
        $Tenant.ReportInProgress = $false
    }
    if ($Response -like 'GENERATE_REPORTS_NO_DATA:*') {
        $Tenant.OutputFolder = ($Response -replace '^GENERATE_REPORTS_NO_DATA:', '').Trim()
        $Tenant.ReportInProgress = $false
    }
    if ($Response -like 'GENERATE_REPORTS_FAILED:*') {
        $Tenant.ReportInProgress = $false
    }
    if ($Response -like 'CANCEL_AUTH_SUCCESS*') {
        $Tenant.GraphAuthenticated = $false
        $Tenant.ExchangeAuthenticated = $false
        $Tenant.GraphTenantId = $null
        $Tenant.GraphTenantName = $null
        $Tenant.ExoTenantId = $null
        $Tenant.ExoOrganizationName = $null
    }
    if ($Response -like 'GRAPH_DISCONNECT_SUCCESS*') {
        $Tenant.GraphAuthenticated = $false
        $Tenant.GraphTenantId = $null
        $Tenant.GraphTenantName = $null
    }
}

function Sync-BulkRunnerTenantResponseFromFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Session,

        [Parameter(Mandatory = $true)]
        [int]$ClientNumber
    )

    if (-not $Session.Tenants.ContainsKey([string]$ClientNumber)) {
        return $null
    }

    $tenant = $Session.Tenants[[string]$ClientNumber]
    $responseFile = Join-Path $Session.CommandDir "Client${ClientNumber}_Response.txt"
    if (-not (Test-Path $responseFile)) {
        return $null
    }

    $responseItem = Get-Item -LiteralPath $responseFile -ErrorAction SilentlyContinue
    if ($tenant.CommandSentAt -and $responseItem -and $responseItem.LastWriteTime -lt $tenant.CommandSentAt.AddMilliseconds(-50)) {
        return $null
    }

    $response = (Get-Content $responseFile -Raw -ErrorAction SilentlyContinue)
    if ([string]::IsNullOrWhiteSpace($response)) {
        return $null
    }

    $response = $response.Trim()
    Update-BulkRunnerTenantFromResponse -Tenant $tenant -Response $response
    return $response
}

function Sync-BulkRunnerSessionResponsesFromFiles {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Session
    )

    foreach ($key in @($Session.Tenants.Keys)) {
        Sync-BulkRunnerTenantResponseFromFile -Session $Session -ClientNumber ([int]$key) | Out-Null
    }
}

function Test-BulkRunnerTenantWorkerAlive {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Tenant
    )

    if (-not $Tenant.ProcessId) {
        return $false
    }

    return $null -ne (Get-Process -Id $Tenant.ProcessId -ErrorAction SilentlyContinue)
}

function Get-BulkRunnerTenantWorkerState {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Session,

        [Parameter(Mandatory = $true)]
        [int]$ClientNumber
    )

    if (-not $Session.Tenants.ContainsKey([string]$ClientNumber)) {
        throw "Unknown tenant client number: $ClientNumber"
    }

    $tenant = $Session.Tenants[[string]$ClientNumber]
    $alive = Test-BulkRunnerTenantWorkerAlive -Tenant $tenant
    return [pscustomobject]@{
        alive          = $alive
        processId      = $tenant.ProcessId
        requiresReauth = -not $alive
    }
}

function Ensure-BulkRunnerTenantWorker {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Session,

        [Parameter(Mandatory = $true)]
        [int]$ClientNumber,

        [switch]$RestartIfDead,

        [switch]$ShowConsole
    )

    if (-not $Session.Tenants.ContainsKey([string]$ClientNumber)) {
        throw "Unknown tenant client number: $ClientNumber"
    }

    $tenant = $Session.Tenants[[string]$ClientNumber]
    if (Test-BulkRunnerTenantWorkerAlive -Tenant $tenant) {
        return [pscustomobject]@{
            alive          = $true
            restarted      = $false
            processId      = $tenant.ProcessId
            requiresReauth = $false
        }
    }

    if (-not $RestartIfDead) {
        return [pscustomobject]@{
            alive          = $false
            restarted      = $false
            processId      = $tenant.ProcessId
            requiresReauth = $true
        }
    }

    $tenant = Restart-BulkRunnerTenant -Session $Session -ClientNumber $ClientNumber -ShowConsole:$ShowConsole
    return [pscustomobject]@{
        alive          = $true
        restarted      = $true
        processId      = $tenant.ProcessId
        requiresReauth = $true
    }
}

function Send-BulkRunnerCommand {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Session,

        [Parameter(Mandatory = $true)]
        [int]$ClientNumber,

        [Parameter(Mandatory = $true)]
        [string]$Command,

        [int]$TimeoutSeconds = 60,

        [switch]$NoWait
    )

    if (-not $Session.Tenants.ContainsKey([string]$ClientNumber)) {
        throw "Unknown tenant client number: $ClientNumber"
    }

    $commandFile = Join-Path $Session.CommandDir "Client${ClientNumber}_Command.txt"
    $responseFile = Join-Path $Session.CommandDir "Client${ClientNumber}_Response.txt"
    $tenant = $Session.Tenants[[string]$ClientNumber]

    if (-not (Test-BulkRunnerTenantWorkerAlive -Tenant $tenant)) {
        if ($tenant.ReportInProgress) {
            $tenant.ReportInProgress = $false
        }
        throw "PowerShell worker is not running (last PID $($tenant.ProcessId)). Restart the worker, re-authenticate, then retry."
    }

    if (Test-Path $responseFile) {
        Remove-Item $responseFile -Force -ErrorAction SilentlyContinue
        Start-Sleep -Milliseconds 100
    }

    $tenant.LastResponse = $null
    $tenant.CommandSentAt = Get-Date
    if ($Command -match '^GENERATE_REPORTS') {
        $tenant.ReportInProgress = $true
    }

    $utf8NoBom = New-Object System.Text.UTF8Encoding $false
    [System.IO.File]::WriteAllText($commandFile, $Command, $utf8NoBom)

    if ($NoWait -or $TimeoutSeconds -le 0) {
        return $null
    }

    $startTime = Get-Date
    while (((Get-Date) - $startTime).TotalSeconds -lt $TimeoutSeconds) {
        if (Test-Path $responseFile) {
            Start-Sleep -Milliseconds 200
            $response = (Get-Content $responseFile -Raw -ErrorAction SilentlyContinue)
            if ($response) {
                $response = $response.Trim()
                Update-BulkRunnerTenantFromResponse -Tenant $Session.Tenants[[string]$ClientNumber] -Response $response
                return $response
            }
        }
        Start-Sleep -Milliseconds 200
    }

    return $null
}

function Get-BulkRunnerTenantStatus {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Session,

        [Parameter(Mandatory = $true)]
        [int]$ClientNumber,

        [int]$TailLines = 200,
        [long]$SinceOffset = 0
    )

    if (-not $Session.Tenants.ContainsKey([string]$ClientNumber)) {
        throw "Unknown tenant client number: $ClientNumber"
    }

    $statusFile = $Session.Tenants[[string]$ClientNumber].StatusFile
    if (-not (Test-Path -LiteralPath $statusFile)) {
        return @{
            status   = ''
            offset   = 0
            length   = 0
            statusFile = $statusFile
        }
    }

    $fileInfo = Get-Item -LiteralPath $statusFile
    $length = [long]$fileInfo.Length
    $raw = [System.IO.File]::ReadAllText($statusFile)

    if ($SinceOffset -gt 0 -and $SinceOffset -lt $raw.Length) {
        $raw = $raw.Substring([int]$SinceOffset)
    } elseif ($SinceOffset -ge $raw.Length) {
        $raw = ''
    }

    if ($TailLines -gt 0 -and $raw.Length -gt 0) {
        $lines = $raw -split "`r?`n"
        if ($lines.Count -gt $TailLines) {
            $raw = ($lines[-$TailLines..-1] -join "`n")
        }
    }

    return @{
        status     = $raw
        offset     = $length
        length     = $length
        statusFile = $statusFile
    }
}

function Get-BulkRunnerAppRegistrations {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProjectRoot,

        [switch]$ForceRefreshFromGraph,
        [switch]$SkipGraphLookup
    )

    Import-Module (Join-Path $ProjectRoot 'Modules\GraphAppCredential.psm1') -Force -ErrorAction Stop
    $list = @()
    if (Get-Command Get-WCMTenantListWithNamesForAppRegCombo -ErrorAction SilentlyContinue) {
        if ($ForceRefreshFromGraph) {
            $list = Get-WCMTenantListWithNamesForAppRegCombo -ForceRefreshFromGraph
        } elseif ($SkipGraphLookup) {
            $list = Get-WCMTenantListWithNamesForAppRegCombo -SkipGraphLookup
        } else {
            $list = Get-WCMTenantListWithNamesForAppRegCombo
        }
    } elseif (Get-Command Get-WCMTenantListWithNames -ErrorAction SilentlyContinue) {
        if ($ForceRefreshFromGraph) {
            $list = Get-WCMTenantListWithNames -ForceRefreshFromGraph
        } elseif ($SkipGraphLookup) {
            $list = Get-WCMTenantListWithNames -SkipGraphLookup
        } else {
            $list = Get-WCMTenantListWithNames
        }
    }

    return @($list | ForEach-Object {
        [pscustomobject]@{
            displayText = $_.DisplayText
            tenantId    = $_.TenantId
        }
    })
}

function Get-BulkRunnerSessionSummary {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Session
    )

    $lastSync = $null
    if ($Session.PSObject.Properties['LastResponseSyncAt']) { $lastSync = $Session.LastResponseSyncAt }
    $shouldSync = (-not $lastSync) -or (((Get-Date) - $lastSync).TotalSeconds -ge 2)
    if ($shouldSync) {
        Sync-BulkRunnerSessionResponsesFromFiles -Session $Session
        $Session | Add-Member -NotePropertyName LastResponseSyncAt -NotePropertyValue (Get-Date) -Force
    }
    Sync-BulkRunnerSessionManifest -Session $Session | Out-Null

    $tenants = @($Session.Tenants.Values | Sort-Object ClientNumber | ForEach-Object {
        $key = [string]$_.ClientNumber
        $ui = $null
        if ($Session.TenantUiStates.ContainsKey($key)) { $ui = $Session.TenantUiStates[$key] }
        [pscustomobject]@{
            clientNumber          = $_.ClientNumber
            processId             = $_.ProcessId
            workerAlive           = (Test-BulkRunnerTenantWorkerAlive -Tenant $_)
            graphAuthenticated    = $_.GraphAuthenticated
            exchangeAuthenticated = $_.ExchangeAuthenticated
            graphTenantId         = $_.GraphTenantId
            graphTenantName       = $_.GraphTenantName
            exoTenantId           = $_.ExoTenantId
            exoOrganizationName  = $_.ExoOrganizationName
            lastResponse          = $_.LastResponse
            outputFolder          = $_.OutputFolder
            reportInProgress      = $_.ReportInProgress
            uiState               = $ui
        }
    })

    return [pscustomobject]@{
        active           = $true
        sessionId        = $Session.SessionId
        tempDir          = $Session.TempDir
        createdAt        = $Session.CreatedAt
        daysBack         = $Session.DaysBack
        reportSelections = $Session.ReportSelections
        tenantCount      = $tenants.Count
        tenants          = $tenants
    }
}

Export-ModuleMember -Function Set-BulkRunnerWorkerVisibility, New-BulkRunnerSession, Add-BulkRunnerTenant, Remove-BulkRunnerTenant, Restart-BulkRunnerTenant, Stop-BulkRunnerSessionTenants, Test-BulkRunnerTenantWorkerAlive, Get-BulkRunnerTenantWorkerState, Ensure-BulkRunnerTenantWorker, Send-BulkRunnerCommand, Get-BulkRunnerTenantStatus, Get-BulkRunnerAppRegistrations, Get-BulkRunnerSessionSummary, Get-BulkRunnerDefaultSettings, Get-BulkRunnerDefaultReportSelections, ConvertTo-BulkRunnerReportSelectionsHashtable, Merge-BulkRunnerReportSelections, Set-BulkRunnerSessionReportSelections, Get-BulkRunnerTenantReportSelectionsOverride, Test-BulkRunnerTenantUsesSessionReportDefaults, Get-BulkRunnerEffectiveReportSelections, Write-BulkRunnerTenantReportSelectionsFile, Expand-BulkRunnerGenerateReportsCommand, Sync-BulkRunnerSessionManifest, Sync-BulkRunnerSessionManifestIfMissing, Sync-BulkRunnerTenantResponseFromFile, Sync-BulkRunnerSessionResponsesFromFiles, Set-BulkRunnerTenantUiState, Save-BulkRunnerSessionManifest, Get-BulkRunnerSessionHistory, Archive-BulkRunnerSessionHistory, Unarchive-BulkRunnerSessionHistory, Remove-BulkRunnerSessionHistory, Get-BulkRunnerSessionManifest, New-BulkRunnerSessionFromManifest
