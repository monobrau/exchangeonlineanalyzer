#Requires -Version 5.1
<#
.SYNOPSIS
    Local web UI + API for bulk tenant export (Option B — Windows runner).

.DESCRIPTION
    Serves a browser UI on http://127.0.0.1:8765/ and orchestrates per-tenant
    BulkExportWorker.ps1 processes using the same file IPC as BulkTenantExporter.ps1.

    Interactive Graph/Exchange auth popups appear on this machine.

.EXAMPLE
    .\Start-BulkWebRunner.ps1
    .\Start-BulkWebRunner.ps1 -Port 8765 -NoBrowser
    .\Start-BulkWebRunner.ps1 -ShowWorkers
#>
param(
    [int]$Port = 8765,
    [switch]$NoBrowser,
    [switch]$ShowWorkers
)

$ErrorActionPreference = 'Stop'
$script:RunnerRoot = $PSScriptRoot
$script:ProjectRoot = Split-Path $PSScriptRoot -Parent
Import-Module (Join-Path $script:RunnerRoot 'Modules\BulkRunnerSession.psm1') -Force
Set-BulkRunnerWorkerVisibility -Hidden:(-not $ShowWorkers)
$script:BulkSession = $null

function Write-JsonResponse {
    param([System.Net.HttpListenerResponse]$Response, [object]$Body, [int]$StatusCode = 200)
    try {
        if (-not $Response.OutputStream.CanWrite) { return }
        $json = $Body | ConvertTo-Json -Depth 8 -Compress
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($json)
        $Response.StatusCode = $StatusCode
        $Response.ContentType = 'application/json; charset=utf-8'
        $Response.ContentLength64 = $bytes.Length
        $Response.OutputStream.Write($bytes, 0, $bytes.Length)
    } finally {
        try { $Response.OutputStream.Close() } catch { }
        try { $Response.Close() } catch { }
    }
}

function Write-TextResponse {
    param(
        [System.Net.HttpListenerResponse]$Response,
        [string]$Body,
        [string]$ContentType = 'text/plain; charset=utf-8',
        [int]$StatusCode = 200
    )
    try {
        if (-not $Response.OutputStream.CanWrite) { return }
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($Body)
        $Response.StatusCode = $StatusCode
        $Response.ContentType = $ContentType
        $Response.ContentLength64 = $bytes.Length
        $Response.OutputStream.Write($bytes, 0, $bytes.Length)
    } finally {
        try { $Response.OutputStream.Close() } catch { }
        try { $Response.Close() } catch { }
    }
}

function Write-ByteResponse {
    param(
        [System.Net.HttpListenerResponse]$Response,
        [byte[]]$Bytes,
        [string]$ContentType,
        [int]$StatusCode = 200
    )
    try {
        if (-not $Response.OutputStream.CanWrite) { return }
        $Response.StatusCode = $StatusCode
        $Response.ContentType = $ContentType
        $Response.ContentLength64 = $Bytes.Length
        $Response.OutputStream.Write($Bytes, 0, $Bytes.Length)
    } finally {
        try { $Response.OutputStream.Close() } catch { }
        try { $Response.Close() } catch { }
    }
}

function Read-RequestBody {
    param([System.Net.HttpListenerRequest]$Request)
    if (-not $Request.HasEntityBody) { return $null }
    $reader = New-Object System.IO.StreamReader($Request.InputStream, $Request.ContentEncoding)
    $text = $reader.ReadToEnd()
    $reader.Close()
    if ([string]::IsNullOrWhiteSpace($text)) { return $null }
    try { return $text | ConvertFrom-Json } catch { return $text }
}

function Get-StaticFilePath {
    param([string]$UrlPath)
    $rel = $UrlPath.TrimStart('/')
    if ([string]::IsNullOrWhiteSpace($rel)) { $rel = 'index.html' }
    $full = Join-Path (Join-Path $script:RunnerRoot 'www') ($rel -replace '/', [IO.Path]::DirectorySeparatorChar)
    $wwwRoot = (Resolve-Path (Join-Path $script:RunnerRoot 'www')).Path
    $resolved = [System.IO.Path]::GetFullPath($full)
    if (-not $resolved.StartsWith($wwwRoot, [StringComparison]::OrdinalIgnoreCase)) {
        return $null
    }
    if (Test-Path -LiteralPath $resolved -PathType Leaf) { return $resolved }
    return $null
}

function Handle-ApiRequest {
    param(
        [string]$Method,
        [string]$Path,
        [object]$Body,
        [Uri]$RequestUrl = $null
    )

    if ($Method -eq 'GET' -and $Path -eq '/api/health') {
        return @{
            ok       = $true
            version  = '0.4.0'
            projectRoot = $script:ProjectRoot
            hiddenWorkers = (-not $ShowWorkers)
            features = @{
                sessionHistory        = $true
                sessionHistoryActions = $true
                reportSelections      = $true
                exportPresets         = $true
                noWaitCommands        = $true
                hiddenWorkers         = $true
                wcmManagement         = $true
                workerLogTabs         = $true
                liongardIntegration   = $true
                huntressIntegration   = $true
                sentinelOneIntegration = $true
            }
        }
    }

    if ($Method -eq 'GET' -and $Path -eq '/api/export-presets') {
        Import-Module (Join-Path $script:ProjectRoot 'Modules\Settings.psm1') -Force -ErrorAction Stop
        $raw = Get-ExportPresets
        $presets = [System.Collections.ArrayList]::new()
        foreach ($name in @($raw.Keys)) {
            $sel = $raw[$name]
            $selections = $null
            if ($sel) {
                $selections = @{}
                $sel.GetEnumerator() | ForEach-Object { $selections[$_.Key] = $_.Value }
            }
            [void]$presets.Add([ordered]@{ name = [string]$name; selections = $selections })
        }
        return @{ presets = @($presets) }
    }

    if ($Method -eq 'POST' -and $Path -eq '/api/session/open-temp') {
        if (-not $script:BulkSession) { throw 'No session.' }
        $opened = Open-BulkRunnerFolderInExplorer -Path $script:BulkSession.TempDir -AllowTempRoot
        return @{ ok = $true; path = $opened }
    }

    if ($Method -eq 'POST' -and $Path -eq '/api/analyze-reports') {
        if (-not $Body -or [string]::IsNullOrWhiteSpace([string]$Body.path)) { throw 'Missing path' }
        Import-Module (Join-Path $script:ProjectRoot 'Modules\ReportAnalysis.psm1') -Force -ErrorAction Stop
        $result = Invoke-ReportFolderAnalysis -Path ([string]$Body.path) -WriteOutputFiles
        return @{
            ok     = $true
            path   = [string]$Body.path
            result = $result
        }
    }

    if ($Method -eq 'POST' -and $Path -eq '/api/ticket/extract-emails') {
        if (-not $Body -or [string]::IsNullOrWhiteSpace([string]$Body.ticketContent)) { throw 'Missing ticketContent' }
        Import-Module (Join-Path $script:ProjectRoot 'Modules\Settings.psm1') -Force -ErrorAction Stop
        $emails = @(Extract-EmailsFromTicket -TicketContent ([string]$Body.ticketContent))
        return @{ emails = $emails; count = $emails.Count }
    }

    if ($Method -eq 'POST' -and $Path -eq '/api/session') {
        if ($script:BulkSession) {
            Save-BulkRunnerSessionManifest -Session $script:BulkSession -Status 'replaced' | Out-Null
        }
        $selections = Get-BulkRunnerDefaultReportSelections
        if ($Body.reportSelections) {
            $Body.reportSelections.PSObject.Properties | ForEach-Object { $selections[$_.Name] = $_.Value }
        }
        $defaults = Get-BulkRunnerDefaultSettings -ProjectRoot $script:ProjectRoot
        $investigator = if ($Body.investigatorName) { [string]$Body.investigatorName } else { $defaults.InvestigatorName }
        $company = if ($Body.companyName) { [string]$Body.companyName } else { $defaults.CompanyName }
        $script:BulkSession = New-BulkRunnerSession -ProjectRoot $script:ProjectRoot `
            -ReportSelections $selections `
            -InvestigatorName $investigator `
            -CompanyName $company `
            -DaysBack ([int]($Body.daysBack | ForEach-Object { if ($_) { $_ } else { 7 } }))
        return (Get-BulkRunnerSessionSummary -Session $script:BulkSession)
    }

    if ($Method -eq 'POST' -and $Path -eq '/api/session/report-selections') {
        if (-not $script:BulkSession) { throw 'No session.' }
        $selections = @{}
        if ($Body.reportSelections) {
            $Body.reportSelections.PSObject.Properties | ForEach-Object { $selections[$_.Name] = $_.Value }
        } else {
            throw 'Missing reportSelections'
        }
        $days = -1
        if ($null -ne $Body.daysBack) { $days = [int]$Body.daysBack }
        Set-BulkRunnerSessionReportSelections -Session $script:BulkSession -ReportSelections $selections -DaysBack $days
        Sync-BulkRunnerSessionManifest -Session $script:BulkSession -Force | Out-Null
        return (Get-BulkRunnerSessionSummary -Session $script:BulkSession)
    }

    if ($Method -eq 'POST' -and $Path -eq '/api/session/restore') {
        if (-not $Body.sessionId) { throw 'Missing sessionId' }
        $force = $false
        if ($Body.force -eq $true -or "$($Body.force)" -eq 'true') { $force = $true }
        if ($script:BulkSession -and -not $force) {
            throw 'An active session is in progress. Pass force:true to start a new session from history (does not stop existing worker windows).'
        }
        if ($script:BulkSession) {
            Save-BulkRunnerSessionManifest -Session $script:BulkSession -Status 'replaced' | Out-Null
        }
        $manifest = Get-BulkRunnerSessionManifest -SessionId ([string]$Body.sessionId)
        $restored = New-BulkRunnerSessionFromManifest -ProjectRoot $script:ProjectRoot -Manifest $manifest
        $script:BulkSession = $restored.Session
        Sync-BulkRunnerSessionManifest -Session $script:BulkSession -Force | Out-Null
        $summary = Get-BulkRunnerSessionSummary -Session $script:BulkSession
        return @{
            restoredFrom   = $restored.SourceId
            tenantSnapshots = @($restored.Snapshots)
            session        = $summary
        }
    }

    if ($Method -eq 'GET' -and $Path -eq '/api/sessions/history') {
        $limit = 50
        $archived = $false
        $sortBy = 'updatedAt'
        $sortOrder = 'desc'
        if ($RequestUrl -and -not [string]::IsNullOrWhiteSpace($RequestUrl.Query)) {
            $q = $RequestUrl.Query
            if ($q -match '[?&]limit=(\d+)') { $limit = [int]$Matches[1] }
            if ($q -match '[?&]archived=(?:1|true)(?:&|$)') { $archived = $true }
            if ($q -match '[?&]sort=([a-zA-Z]+)') { $sortBy = [string]$Matches[1] }
            if ($q -match '[?&]order=(asc|desc)(?:&|$)') { $sortOrder = [string]$Matches[1] }
        }
        return @{
            sessions = @(Get-BulkRunnerSessionHistory -Limit $limit -Archived:$archived -SortBy $sortBy -SortOrder $sortOrder)
            archived = $archived
            sortBy   = $sortBy
            sortOrder = $sortOrder
        }
    }

    if ($Method -eq 'POST' -and $Path -match '^/api/sessions/history/([^/]+)/archive$') {
        $sessionId = [string]$Matches[1]
        $path = Archive-BulkRunnerSessionHistory -SessionId $sessionId
        return @{ ok = $true; sessionId = $sessionId; archived = $true; path = $path }
    }

    if ($Method -eq 'POST' -and $Path -match '^/api/sessions/history/([^/]+)/unarchive$') {
        $sessionId = [string]$Matches[1]
        $path = Unarchive-BulkRunnerSessionHistory -SessionId $sessionId
        return @{ ok = $true; sessionId = $sessionId; archived = $false; path = $path }
    }

    if ($Method -eq 'DELETE' -and $Path -match '^/api/sessions/history/([^/]+)$') {
        $sessionId = [string]$Matches[1]
        return (Remove-BulkRunnerSessionHistory -SessionId $sessionId)
    }

    if ($Method -eq 'GET' -and $Path -match '^/api/sessions/history/([^/]+)$') {
        $manifest = Get-BulkRunnerSessionManifest -SessionId $Matches[1]
        return @{ manifest = $manifest }
    }

    if ($Method -eq 'GET' -and $Path -eq '/api/session') {
        if (-not $script:BulkSession) {
            return @{
                active      = $false
                sessionId   = $null
                tenantCount = 0
                tenants     = @()
            }
        }
        Sync-BulkRunnerSessionManifestIfMissing -Session $script:BulkSession | Out-Null
        return (Get-BulkRunnerSessionSummary -Session $script:BulkSession)
    }

    if ($Method -eq 'GET' -and $Path -eq '/api/app-registrations') {
        $forceRefresh = $false
        if ($RequestUrl -and -not [string]::IsNullOrWhiteSpace($RequestUrl.Query)) {
            $forceRefresh = ($RequestUrl.Query -match '(?:^|[&?])(?:refresh|forceRefresh)=1(?:&|$)')
        }
        $regParams = @{
            ProjectRoot     = $script:ProjectRoot
            SkipGraphLookup = (-not $forceRefresh)
        }
        if ($forceRefresh) { $regParams['ForceRefreshFromGraph'] = $true }
        return @(Get-BulkRunnerAppRegistrations @regParams)
    }

    if ($Path -like '/api/wcm/*') {
        Import-Module (Join-Path $script:ProjectRoot 'Modules\GraphAppCredential.psm1') -Force -ErrorAction Stop

        if ($Method -eq 'GET' -and $Path -eq '/api/wcm/entries') {
            $rows = [System.Collections.ArrayList]::new()
            foreach ($pfx in @('EOA', 'ESR')) {
                foreach ($t in @(Get-WCMTenantListWithNames -Prefix $pfx -SkipGraphLookup)) {
                    [void]$rows.Add([ordered]@{
                        kind        = 'Tenant'
                        tenantId    = [string]$t.TenantId
                        displayText = [string]$t.DisplayText
                        wcmPrefix   = $pfx
                        orphanTarget = $null
                    })
                }
                foreach ($o in @(Get-WCMUnrecognizedGraphAppTargets -Prefix $pfx)) {
                    [void]$rows.Add([ordered]@{
                        kind         = 'Orphan'
                        tenantId     = $null
                        displayText  = "Unrecognized WCM target ($pfx): $o"
                        wcmPrefix    = $pfx
                        orphanTarget = [string]$o
                    })
                }
            }
            return @{ entries = @($rows | Sort-Object displayText) }
        }

        if ($Method -eq 'POST' -and $Path -eq '/api/wcm/create-graph-app') {
            $tid = if ($Body.tenantId) { [string]$Body.tenantId } else { '' }
            $outcome = Invoke-GraphAppCreateWithWcmSave -ProjectRoot $script:ProjectRoot -TenantId $tid
            return @{
                exitCode = $outcome.ExitCode
                result   = $outcome.Result
                logPath  = $outcome.LogPath
            }
        }

        if ($Method -eq 'POST' -and $Path -eq '/api/wcm/delete-graph-app') {
            if (-not $Body.tenantIds -or @($Body.tenantIds).Count -eq 0) { throw 'Missing tenantIds' }
            $scriptPath = Join-Path $script:ProjectRoot 'Remove-GraphInboxRulesApp.ps1'
            if (-not (Test-Path -LiteralPath $scriptPath)) { throw "Missing script: $scriptPath" }
            $psExe = (Get-Process -Id $PID -ErrorAction Stop).Path
            $results = [System.Collections.ArrayList]::new()
            foreach ($tid in @($Body.tenantIds)) {
                $t = [string]$tid
                if ([string]::IsNullOrWhiteSpace($t)) { continue }
                $proc = Start-Process -FilePath $psExe -ArgumentList @(
                    '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $scriptPath, '-TenantId', $t, '-Force'
                ) -Wait -PassThru -WorkingDirectory $script:ProjectRoot
                [void]$results.Add([ordered]@{ tenantId = $t; exitCode = $proc.ExitCode })
            }
            return @{ removed = @($results) }
        }

        if ($Method -eq 'POST' -and $Path -eq '/api/wcm/export') {
            if (-not $Body.path -or -not $Body.password) { throw 'Missing path or password' }
            $sec = ConvertTo-SecureString ([string]$Body.password) -AsPlainText -Force
            $prefix = if ($Body.prefix) { [string]$Body.prefix } else { 'EOA' }
            $count = Export-GraphAppCredentialsToFile -Path ([string]$Body.path) -Password $sec -Prefix $prefix
            return @{ ok = $true; path = [string]$Body.path; count = $count }
        }

        if ($Method -eq 'POST' -and $Path -eq '/api/wcm/import') {
            if (-not $Body.path -or -not $Body.password) { throw 'Missing path or password' }
            $sec = ConvertTo-SecureString ([string]$Body.password) -AsPlainText -Force
            $count = Import-GraphAppCredentialsFromFile -Path ([string]$Body.path) -Password $sec
            return @{ ok = $true; imported = $count }
        }

        if ($Method -eq 'POST' -and $Path -eq '/api/wcm/clear-local') {
            if (-not $Body.items -or @($Body.items).Count -eq 0) { throw 'Missing items' }
            $removed = 0
            foreach ($item in @($Body.items)) {
                $pfx = if ($item.wcmPrefix) { [string]$item.wcmPrefix } else { 'EOA' }
                if ([string]$item.kind -eq 'Orphan' -and $item.orphanTarget) {
                    Remove-WCMGraphCredentialTarget -TargetName ([string]$item.orphanTarget) | Out-Null
                    $removed++
                } elseif ($item.tenantId) {
                    Remove-GraphAppCredentialsLocalOnly -TenantId @([string]$item.tenantId) -Prefix $pfx
                    $removed++
                }
            }
            return @{ ok = $true; removed = $removed }
        }
    }

    if ($Method -eq 'GET' -and $Path -eq '/api/liongard/status') {
        Import-Module (Join-Path $script:ProjectRoot 'Modules\Liongard.psm1') -Force -ErrorAction Stop
        $status = Get-LiongardConfigurationStatus
        return @{
            configured = $status.Configured
            instance   = $status.Instance
            message    = $status.Message
        }
    }

    if ($Method -eq 'POST' -and $Path -eq '/api/liongard/resolve-client') {
        if (-not $Body -or [string]::IsNullOrWhiteSpace([string]$Body.companyName)) {
            throw 'Missing companyName in request body.'
        }
        Import-Module (Join-Path $script:ProjectRoot 'Modules\Liongard.psm1') -Force -ErrorAction Stop
        $envId = 0
        if ($null -ne $Body.environmentId) { $envId = [int]$Body.environmentId }
        return (Resolve-LiongardClient -CompanyName ([string]$Body.companyName).Trim() `
            -TicketContent ([string]$Body.ticketContent) `
            -EnvironmentId $envId)
    }

    if ($Method -eq 'POST' -and $Path -eq '/api/liongard/export-context') {
        if (-not $Body -or -not $Body.environmentId) { throw 'Missing environmentId.' }
        Import-Module (Join-Path $script:ProjectRoot 'Modules\Liongard.psm1') -Force -ErrorAction Stop
        Import-Module (Join-Path $script:ProjectRoot 'Modules\SecurityIntegrations.psm1') -Force -ErrorAction Stop
        $folder = Get-SecurityIntegrationExportFolder -CompanyName ([string]$Body.companyName) -OutputFolder ([string]$Body.outputFolder) -TicketNumber ([string]$Body.ticketNumber)
        $start = $null
        $end = $null
        if ($Body.startDate) { $start = [datetime]$Body.startDate }
        if ($Body.endDate) { $end = [datetime]$Body.endDate }
        $files = Export-LiongardContext -EnvironmentId ([int]$Body.environmentId) -ExportFolder $folder -TicketNumber ([string]$Body.ticketNumber) -StartDate $start -EndDate $end
        return @{ ok = $true; folder = $folder; files = @($files) }
    }

    if ($Method -eq 'GET' -and $Path -eq '/api/huntress/status') {
        Import-Module (Join-Path $script:ProjectRoot 'Modules\Huntress.psm1') -Force -ErrorAction Stop
        $status = Get-HuntressConfigurationStatus
        return @{ configured = $status.Configured; message = $status.Message }
    }

    if ($Method -eq 'GET' -and $Path -eq '/api/huntress/organizations') {
        Import-Module (Join-Path $script:ProjectRoot 'Modules\Huntress.psm1') -Force -ErrorAction Stop
        $orgs = @(Get-HuntressPagedItems -RelativePath 'organizations' -MaxPages 50)
        $list = foreach ($o in $orgs) {
            @{ id = [int]$o.id; name = [string]$o.name }
        }
        return @{ organizations = @($list); count = $list.Count }
    }

    if ($Method -eq 'POST' -and $Path -eq '/api/huntress/preview') {
        if (-not $Body -or -not $Body.organizationId) { throw 'Missing organizationId.' }
        Import-Module (Join-Path $script:ProjectRoot 'Modules\Huntress.psm1') -Force -ErrorAction Stop
        $since = $null
        if ($Body.updatedSince) { $since = [datetime]$Body.updatedSince }
        $counts = Get-HuntressPreviewCounts -OrganizationId ([int]$Body.organizationId) -UpdatedSince $since
        return @{ ok = $true; counts = $counts }
    }

    if ($Method -eq 'POST' -and $Path -eq '/api/huntress/export') {
        if (-not $Body -or -not $Body.organizationId) { throw 'Missing organizationId.' }
        Import-Module (Join-Path $script:ProjectRoot 'Modules\Huntress.psm1') -Force -ErrorAction Stop
        Import-Module (Join-Path $script:ProjectRoot 'Modules\SecurityIntegrations.psm1') -Force -ErrorAction Stop
        $folder = Get-SecurityIntegrationExportFolder -CompanyName ([string]$Body.companyName) -OutputFolder ([string]$Body.outputFolder) -TicketNumber ([string]$Body.ticketNumber)
        $selections = @{}
        if ($Body.selections) {
            $Body.selections.PSObject.Properties | ForEach-Object { $selections[$_.Name] = [bool]$_.Value }
        }
        $since = $null
        if ($Body.updatedSince) { $since = [datetime]$Body.updatedSince }
        $files = Export-HuntressInvestigation -OrganizationId ([int]$Body.organizationId) -ExportFolder $folder `
            -Selections $selections -TicketNumber ([string]$Body.ticketNumber) -UpdatedSince $since
        return @{ ok = $true; folder = $folder; files = @($files) }
    }

    if ($Method -eq 'GET' -and $Path -eq '/api/sentinelone/status') {
        Import-Module (Join-Path $script:ProjectRoot 'Modules\SentinelOne.psm1') -Force -ErrorAction Stop
        $status = Get-SentinelOneConfigurationStatus
        return @{
            profiles = $status.Profiles
            message  = $status.Message
        }
    }

    if ($Method -eq 'POST' -and $Path -eq '/api/sentinelone/resolve-site') {
        if (-not $Body -or [string]::IsNullOrWhiteSpace([string]$Body.companyName)) {
            throw 'Missing companyName.'
        }
        Import-Module (Join-Path $script:ProjectRoot 'Modules\SentinelOne.psm1') -Force -ErrorAction Stop
        $envId = 0
        if ($null -ne $Body.liongardEnvironmentId) { $envId = [int]$Body.liongardEnvironmentId }
        return (Resolve-SentinelOneSite -CompanyName ([string]$Body.companyName).Trim() `
            -TicketContent ([string]$Body.ticketContent) `
            -ProfileName ([string]$Body.profileName) `
            -SiteIdHint ([string]$Body.siteId) `
            -LiongardEnvironmentId $envId)
    }

    if ($Method -eq 'POST' -and $Path -eq '/api/sentinelone/preview') {
        if (-not $Body -or [string]::IsNullOrWhiteSpace([string]$Body.profileName)) {
            throw 'Missing profileName.'
        }
        Import-Module (Join-Path $script:ProjectRoot 'Modules\SentinelOne.psm1') -Force -ErrorAction Stop
        $after = $null
        $before = $null
        if ($Body.createdAfter) { $after = [datetime]$Body.createdAfter }
        if ($Body.createdBefore) { $before = [datetime]$Body.createdBefore }
        $counts = Get-SentinelOnePreviewCounts -ProfileName ([string]$Body.profileName) `
            -SiteId ([string]$Body.siteId) -CreatedAfter $after -CreatedBefore $before
        return @{ ok = $true; counts = $counts }
    }

    if ($Method -eq 'POST' -and $Path -eq '/api/sentinelone/export') {
        if (-not $Body -or [string]::IsNullOrWhiteSpace([string]$Body.profileName)) {
            throw 'Missing profileName.'
        }
        Import-Module (Join-Path $script:ProjectRoot 'Modules\SentinelOne.psm1') -Force -ErrorAction Stop
        Import-Module (Join-Path $script:ProjectRoot 'Modules\SecurityIntegrations.psm1') -Force -ErrorAction Stop
        $folder = Get-SecurityIntegrationExportFolder -CompanyName ([string]$Body.companyName) -OutputFolder ([string]$Body.outputFolder) -TicketNumber ([string]$Body.ticketNumber)
        $selections = @{}
        if ($Body.selections) {
            $Body.selections.PSObject.Properties | ForEach-Object { $selections[$_.Name] = [bool]$_.Value }
        }
        $after = $null
        $before = $null
        if ($Body.createdAfter) { $after = [datetime]$Body.createdAfter }
        if ($Body.createdBefore) { $before = [datetime]$Body.createdBefore }
        $files = Export-SentinelOneInvestigation -ProfileName ([string]$Body.profileName) -ExportFolder $folder `
            -Selections $selections -SiteId ([string]$Body.siteId) -TicketNumber ([string]$Body.ticketNumber) `
            -CreatedAfter $after -CreatedBefore $before
        return @{ ok = $true; folder = $folder; files = @($files) }
    }

    if ($Method -eq 'GET' -and $Path -eq '/api/manage/status') {
        Import-Module (Join-Path $script:ProjectRoot 'Modules\Settings.psm1') -Force -ErrorAction SilentlyContinue
        Import-Module (Join-Path $script:ProjectRoot 'Modules\ConnectWiseManage.psm1') -Force -ErrorAction Stop
        $status = Get-ConnectWiseManageConfigurationStatus
        return @{
            configured    = $status.Configured
            source        = $status.Source
            message       = $status.Message
            vscanComplete = $status.VScanComplete
            settingsPath  = $status.SettingsPath
        }
    }

    if ($Method -eq 'POST' -and $Path -eq '/api/manage/ticket') {
        if (-not $Body -or [string]::IsNullOrWhiteSpace([string]$Body.ticketId)) {
            throw 'Missing ticketId in request body.'
        }
        Import-Module (Join-Path $script:ProjectRoot 'Modules\Settings.psm1') -Force -ErrorAction SilentlyContinue
        Import-Module (Join-Path $script:ProjectRoot 'Modules\ConnectWiseManage.psm1') -Force -ErrorAction Stop
        $fetched = Get-ConnectWiseManageServiceTicketText -TicketId ([string]$Body.ticketId).Trim()
        Import-Module (Join-Path $script:ProjectRoot 'Modules\Settings.psm1') -Force -ErrorAction SilentlyContinue
        $securityStack = $null
        if (Get-Command Get-SecurityStackFromTicket -ErrorAction SilentlyContinue) {
            $securityStack = Get-SecurityStackFromTicket -TicketContent $fetched.TicketContent -Summary $fetched.Summary
        }
        return @{
            success       = $true
            ticketId      = $fetched.TicketId
            summary       = $fetched.Summary
            companyName   = $fetched.CompanyName
            ticketNumbers = @($fetched.TicketNumbers)
            ticketContent = $fetched.TicketContent
            contentLength = $fetched.FilteredLength
            securityStack = $securityStack
        }
    }

    if ($Method -eq 'POST' -and $Path -eq '/api/open-folder') {
        if (-not $Body -or -not $Body.path) { throw 'Missing path' }
        $opened = Open-BulkRunnerFolderInExplorer -Path ([string]$Body.path)
        return @{ ok = $true; path = $opened }
    }

    if ($Method -eq 'POST' -and $Path -eq '/api/tenants') {
        if (-not $script:BulkSession) { throw 'No session. POST /api/session first.' }
        $tenant = Add-BulkRunnerTenant -Session $script:BulkSession
        Sync-BulkRunnerSessionManifest -Session $script:BulkSession -Force | Out-Null
        return @{
            clientNumber = $tenant.ClientNumber
            processId    = $tenant.ProcessId
        }
    }

    if ($Method -eq 'DELETE' -and $Path -match '^/api/tenants/(\d+)$') {
        if (-not $script:BulkSession) { throw 'No session.' }
        $clientNumber = [int]$Matches[1]
        Remove-BulkRunnerTenant -Session $script:BulkSession -ClientNumber $clientNumber
        Sync-BulkRunnerSessionManifest -Session $script:BulkSession -Force | Out-Null
        return @{ removed = $clientNumber; tenantCount = $script:BulkSession.Tenants.Count }
    }

    if ($Path -match '^/api/tenants/(\d+)/(command|status|response|restart|ui-state|worker|ensure-worker)$') {
        if (-not $script:BulkSession) { throw 'No session.' }
        $clientNumber = [int]$Matches[1]
        $action = $Matches[2]

        if ($action -eq 'worker' -and $Method -eq 'GET') {
            return Get-BulkRunnerTenantWorkerState -Session $script:BulkSession -ClientNumber $clientNumber
        }

        if ($action -eq 'ensure-worker' -and $Method -eq 'POST') {
            $restartIfDead = $true
            if ($Body -and $null -ne $Body.restartIfDead -and ($Body.restartIfDead -eq $false -or "$($Body.restartIfDead)" -eq 'false')) {
                $restartIfDead = $false
            }
            $showConsole = $false
            if ($Body -and ($Body.showConsole -eq $true -or "$($Body.showConsole)" -eq 'true')) {
                $showConsole = $true
            }
            $result = Ensure-BulkRunnerTenantWorker -Session $script:BulkSession -ClientNumber $clientNumber -RestartIfDead:$restartIfDead -ShowConsole:$showConsole
            Sync-BulkRunnerSessionManifest -Session $script:BulkSession -Force | Out-Null
            return @{
                alive          = $result.alive
                restarted      = $result.restarted
                processId      = $result.processId
                requiresReauth = $result.requiresReauth
            }
        }

        if ($action -eq 'restart' -and $Method -eq 'POST') {
            $showConsole = $false
            if ($Body.showConsole -eq $true -or "$($Body.showConsole)" -eq 'true') { $showConsole = $true }
            $tenant = Restart-BulkRunnerTenant -Session $script:BulkSession -ClientNumber $clientNumber -ShowConsole:$showConsole
            Sync-BulkRunnerSessionManifest -Session $script:BulkSession -Force | Out-Null
            return @{
                clientNumber = $tenant.ClientNumber
                processId    = $tenant.ProcessId
                showConsole  = $showConsole
            }
        }

        if ($action -eq 'ui-state' -and $Method -eq 'POST') {
            $ui = @{}
            if ($Body) {
                $Body.PSObject.Properties | ForEach-Object { $ui[$_.Name] = $_.Value }
            }
            Set-BulkRunnerTenantUiState -Session $script:BulkSession -ClientNumber $clientNumber -UiState $ui
            Sync-BulkRunnerSessionManifest -Session $script:BulkSession -Force | Out-Null
            return @{ ok = $true }
        }

        if ($action -eq 'response' -and $Method -eq 'GET') {
            $response = Sync-BulkRunnerTenantResponseFromFile -Session $script:BulkSession -ClientNumber $clientNumber
            return @{ response = $response }
        }

        if ($action -eq 'status' -and $Method -eq 'GET') {
            $tailLines = 200
            $sinceOffset = 0
            if ($RequestUrl -and -not [string]::IsNullOrWhiteSpace($RequestUrl.Query)) {
                $q = $RequestUrl.Query
                if ($q -match '[?&]tailLines=(\d+)') { $tailLines = [int]$Matches[1] }
                if ($q -match '[?&]sinceOffset=(\d+)') { $sinceOffset = [long]$Matches[1] }
            }
            return (Get-BulkRunnerTenantStatus -Session $script:BulkSession -ClientNumber $clientNumber -TailLines $tailLines -SinceOffset $sinceOffset)
        }

        if ($action -eq 'command' -and $Method -eq 'POST') {
            $cmd = [string]$Body.command
            if ([string]::IsNullOrWhiteSpace($cmd)) { throw 'Missing command' }
            $noWait = $false
            if ($Body.noWait -eq $true -or "$($Body.noWait)" -eq 'true') { $noWait = $true }
            $wait = 60
            if ($null -ne $Body.waitSeconds) { $wait = [int]$Body.waitSeconds }
            if ($noWait) { $wait = 0 }
            $cmd = Expand-BulkRunnerGenerateReportsCommand -Session $script:BulkSession -ClientNumber $clientNumber -Command $cmd
            $cmdParams = @{
                Session        = $script:BulkSession
                ClientNumber   = $clientNumber
                Command        = $cmd
                TimeoutSeconds = $wait
            }
            if ($noWait) { $cmdParams.NoWait = $true }
            $response = Send-BulkRunnerCommand @cmdParams
            if (-not $noWait) {
                Sync-BulkRunnerSessionManifest -Session $script:BulkSession -Force | Out-Null
            } else {
                Sync-BulkRunnerSessionManifest -Session $script:BulkSession | Out-Null
            }
            return @{ response = $response }
        }
    }

    if ($Path -match '^/api/tenants/(\d+)/report-selections$') {
        if (-not $script:BulkSession) { throw 'No session.' }
        $clientNumber = [int]$Matches[1]
        if (-not $script:BulkSession.Tenants.ContainsKey([string]$clientNumber)) {
            throw "Unknown tenant client number: $clientNumber"
        }
        if ($Method -eq 'GET') {
            $override = Get-BulkRunnerTenantReportSelectionsOverride -Session $script:BulkSession -ClientNumber $clientNumber
            $useDefaults = Test-BulkRunnerTenantUsesSessionReportDefaults -Session $script:BulkSession -ClientNumber $clientNumber
            return @{
                useSessionDefaults = $useDefaults
                sessionDefaults    = $script:BulkSession.ReportSelections
                tenantOverride     = $override
                effective          = Get-BulkRunnerEffectiveReportSelections -Session $script:BulkSession -ClientNumber $clientNumber
            }
        }
    }

    throw "Not found: $Method $Path"
}

function Open-BulkRunnerFolderInExplorer {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [switch]$AllowTempRoot
    )

    $resolved = [System.IO.Path]::GetFullPath($Path)
    if (-not (Test-Path -LiteralPath $resolved)) {
        throw "Folder not found: $resolved"
    }

    $docsRoot = [System.IO.Path]::GetFullPath([Environment]::GetFolderPath('MyDocuments'))
    $tempRoot = [System.IO.Path]::GetFullPath($env:TEMP)
    $allowed = ($resolved.StartsWith($docsRoot, [StringComparison]::OrdinalIgnoreCase)) `
        -or ($resolved.StartsWith($tempRoot, [StringComparison]::OrdinalIgnoreCase))
    if (-not $allowed -and -not $AllowTempRoot) {
        throw 'Only folders under Documents or TEMP can be opened from the web UI.'
    }
    if (-not $allowed -and $AllowTempRoot -and -not $resolved.StartsWith($tempRoot, [StringComparison]::OrdinalIgnoreCase)) {
        throw 'Session temp folder must be under TEMP.'
    }

    Start-Process explorer.exe -ArgumentList "`"$resolved`""
    return $resolved
}

$prefix = "http://127.0.0.1:$Port/"
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add($prefix)
$listener.Start()

Write-Host "Bulk Web Runner listening on $prefix" -ForegroundColor Green
Write-Host "Project root: $script:ProjectRoot" -ForegroundColor Gray
Write-Host "Press Ctrl+C to stop." -ForegroundColor Gray

if (-not $NoBrowser) {
    Start-Process $prefix
}

try {
    while ($listener.IsListening) {
        $context = $listener.GetContext()
        $request = $context.Request
        $response = $context.Response
        $path = $request.Url.AbsolutePath

        try {
            if ($path.StartsWith('/api/')) {
                $body = Read-RequestBody -Request $request
                $result = Handle-ApiRequest -Method $request.HttpMethod -Path $path -Body $body -RequestUrl $request.Url
                Write-JsonResponse -Response $response -Body $result
            } else {
                $file = Get-StaticFilePath -UrlPath $path
                if (-not $file) {
                    Write-TextResponse -Response $response -Body 'Not found' -StatusCode 404
                } else {
                    $ext = [IO.Path]::GetExtension($file).ToLowerInvariant()
                    $ctype = switch ($ext) {
                        '.html' { 'text/html; charset=utf-8' }
                        '.js'   { 'application/javascript; charset=utf-8' }
                        '.css'  { 'text/css; charset=utf-8' }
                        default { 'application/octet-stream' }
                    }
                    $response.AddHeader('Cache-Control', 'no-cache, no-store, must-revalidate')
                    $bytes = [System.IO.File]::ReadAllBytes($file)
                    Write-ByteResponse -Response $response -Bytes $bytes -ContentType $ctype
                }
            }
        } catch {
            try {
                Write-JsonResponse -Response $response -Body @{ error = $_.Exception.Message } -StatusCode 400
            } catch {
                Write-Host "Response error: $($_.Exception.Message)" -ForegroundColor Yellow
                try { $response.Close() } catch { }
            }
        }
    }
} finally {
    $listener.Stop()
    $listener.Close()
}
