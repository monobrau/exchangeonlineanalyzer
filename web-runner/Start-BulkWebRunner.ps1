#Requires -Version 5.1
<#
.SYNOPSIS
    Local web UI + API for bulk tenant export (Option B — Windows runner).

.DESCRIPTION
    Serves a browser UI on http://127.0.0.1:8765/ and orchestrates per-tenant
    BulkExportWorker.ps1 processes using the same file IPC as BulkTenantExporter.ps1.

    Interactive Graph/Exchange auth popups appear on this machine.

    Starting again on the same port stops any existing EOA web runner first
    (localhost shutdown when healthy, then force-stop leftovers).

.EXAMPLE
    .\Start-BulkWebRunner.ps1
    .\Start-BulkWebRunner.ps1 -Port 8765 -NoBrowser
    .\Start-BulkWebRunner.ps1 -ShowWorkers
    .\Start-BulkWebRunner.ps1 -ListenLan -NoBrowser
#>
param(
    [int]$Port = 8765,
    [switch]$NoBrowser,
    [switch]$ShowWorkers,
    [switch]$ListenLan
)

$ErrorActionPreference = 'Stop'
if ([string]::IsNullOrWhiteSpace([string]$env:TEMP)) {
    $tempFallback = [string]$env:TMPDIR
    if ([string]::IsNullOrWhiteSpace($tempFallback)) { $tempFallback = [string]$env:TMP }
    if ([string]::IsNullOrWhiteSpace($tempFallback)) { $tempFallback = '/tmp' }
    $env:TEMP = $tempFallback
}
if ([string]::IsNullOrWhiteSpace([string]$env:TMP)) { $env:TMP = [string]$env:TEMP }
$script:RunnerRoot = $PSScriptRoot
$script:ProjectRoot = Split-Path $PSScriptRoot -Parent
. (Join-Path $script:RunnerRoot 'LabUsers.ps1')
$script:ListenLan = [bool]$ListenLan
$script:ListenPort = $Port
Import-Module (Join-Path $script:RunnerRoot 'Modules\BulkRunnerSession.psm1') -Force
Set-BulkRunnerWorkerVisibility -Hidden:(-not $ShowWorkers)
# Per-browser isolation: each browser gets its own cookie → BulkRunner session + GraphAppJob.
# The single-threaded GetContext loop loads the matching context into $script:BulkSession /
# $script:GraphAppJob for the duration of each request, then writes it back.
$script:BrowserContexts = @{}
$script:BulkSession = $null
$script:ShutdownRequested = $false
$script:GraphAppJob = $null
$script:RequestBrowserId = $null
$script:BrowserCookieName = 'eoa_browser'

function Get-BulkWebBrowserIdFromCookie {
    param([string]$CookieHeader)
    if ([string]::IsNullOrWhiteSpace($CookieHeader)) { return $null }
    $name = [regex]::Escape($script:BrowserCookieName)
    if ($CookieHeader -match "(?:^|;\s*)$name=([0-9a-fA-F]{8}(?:-[0-9a-fA-F]{4}){3}-[0-9a-fA-F]{12})(?:;|$)") {
        return $Matches[1].ToLowerInvariant()
    }
    return $null
}

function New-BulkWebBrowserId {
    return [guid]::NewGuid().ToString('D').ToLowerInvariant()
}

function Get-BulkWebBrowserCookieHeader {
    param([Parameter(Mandatory = $true)][string]$BrowserId)
    # HttpOnly so JS cannot steal it; SameSite=Lax keeps same-site API calls working.
    return "$($script:BrowserCookieName)=$BrowserId; Path=/; HttpOnly; SameSite=Lax"
}

function Ensure-BulkWebBrowserContext {
    param([Parameter(Mandatory = $true)][string]$BrowserId)
    if (-not $script:BrowserContexts.ContainsKey($BrowserId)) {
        $script:BrowserContexts[$BrowserId] = [ordered]@{
            BrowserId   = $BrowserId
            Session     = $null
            GraphAppJob = $null
            LastAccess  = (Get-Date).ToUniversalTime()
        }
    }
    return $script:BrowserContexts[$BrowserId]
}

function Test-BulkWebLiveSessionConflict {
    param(
        [string]$SessionId,
        [string]$TempDir
    )
    foreach ($key in @($script:BrowserContexts.Keys)) {
        if ($key -eq $script:RequestBrowserId) { continue }
        $other = $script:BrowserContexts[$key].Session
        if (-not $other) { continue }
        if (-not [string]::IsNullOrWhiteSpace($SessionId) -and [string]$other.SessionId -eq $SessionId) { return $true }
        if (-not [string]::IsNullOrWhiteSpace($TempDir) -and [string]$other.TempDir -eq $TempDir) { return $true }
    }
    return $false
}

function Stop-BulkWebAllBrowserSessions {
    foreach ($key in @($script:BrowserContexts.Keys)) {
        $ctx = $script:BrowserContexts[$key]
        if ($ctx -and $ctx.Session) {
            try { Stop-BulkRunnerSessionTenants -Session $ctx.Session } catch { }
        }
    }
    $script:BrowserContexts.Clear()
    $script:BulkSession = $null
    $script:GraphAppJob = $null
    $script:RequestBrowserId = $null
}

function Write-JsonResponse {
    param(
        [System.Net.HttpListenerResponse]$Response,
        [object]$Body,
        [int]$StatusCode = 200,
        [int]$Depth = 20,
        [string]$SetCookie = $null
    )
    try {
        if (-not $Response.OutputStream.CanWrite) { return }
        $json = $Body | ConvertTo-Json -Depth $Depth -Compress
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($json)
        $Response.StatusCode = $StatusCode
        $Response.ContentType = 'application/json; charset=utf-8'
        if (-not [string]::IsNullOrWhiteSpace($SetCookie)) {
            $Response.Headers.Add('Set-Cookie', $SetCookie)
        }
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
        [int]$StatusCode = 200,
        [string]$ContentDisposition = $null,
        [string]$SetCookie = $null
    )
    try {
        if (-not $Response.OutputStream.CanWrite) { return }
        $Response.StatusCode = $StatusCode
        $Response.ContentType = $ContentType
        if (-not [string]::IsNullOrWhiteSpace($ContentDisposition)) {
            $Response.AddHeader('Content-Disposition', $ContentDisposition)
        }
        if (-not [string]::IsNullOrWhiteSpace($SetCookie)) {
            $Response.Headers.Add('Set-Cookie', $SetCookie)
        }
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

function Resolve-TriageWebFolder {
    param([string]$Folder)
    if ([string]::IsNullOrWhiteSpace($Folder)) { throw 'Missing folder' }
    $resolved = [System.IO.Path]::GetFullPath($Folder)
    $docs = [Environment]::GetFolderPath('MyDocuments')
    if ([string]::IsNullOrWhiteSpace($docs)) {
        $home = [string]$env:HOME
        if ([string]::IsNullOrWhiteSpace($home)) { $home = '/var/lib/exchangeonlineanalyzer' }
        $docs = Join-Path $home 'Documents'
    }
    $root = [System.IO.Path]::GetFullPath((Join-Path (Join-Path $docs 'ExchangeOnlineAnalyzer') 'SecurityInvestigation'))
    if (-not $resolved.StartsWith($root, [StringComparison]::OrdinalIgnoreCase)) {
        throw 'Triage folder must be under Documents\ExchangeOnlineAnalyzer\SecurityInvestigation'
    }
    if ((Split-Path $resolved -Leaf) -notmatch '^Triage_') { throw 'Not a triage pack folder' }
    if (-not (Test-Path -LiteralPath $resolved)) { throw "Folder not found: $resolved" }
    return $resolved
}

function Get-TriageRequestQueryValue {
    param($RequestUrl, [string]$Name)
    if (-not $RequestUrl -or [string]::IsNullOrWhiteSpace($RequestUrl.Query)) { return '' }
    $pattern = '[?&]' + [regex]::Escape($Name) + '=([^&]*)'
    if ($RequestUrl.Query -match $pattern) { return [System.Net.WebUtility]::UrlDecode($Matches[1]) }
    return ''
}

function Get-GraphAppJobPublicView {
    param([string]$LogPath)
    $deviceCode = $null
    $deviceUrl = 'https://login.microsoft.com/device'
    $message = $null
    $deviceFile = if ($LogPath) { $LogPath -replace '-last\.log$','-device.json' } else { $null }
    if ($deviceFile -and (Test-Path -LiteralPath $deviceFile)) {
        try {
            $pending = Get-Content -LiteralPath $deviceFile -Raw -ErrorAction Stop | ConvertFrom-Json
            if ($pending.user_code) { $deviceCode = [string]$pending.user_code }
            if ($pending.verification_uri) { $deviceUrl = [string]$pending.verification_uri }
        } catch { }
    }
    if ($LogPath -and (Test-Path -LiteralPath $LogPath)) {
        try {
            $lines = @(Get-Content -LiteralPath $LogPath -Tail 80 -ErrorAction Stop)
        } catch {
            $lines = @()
        }
        foreach ($line in $lines) {
            $text = [string]$line
            if ($text -match '(?i)Secret\s*:') { continue }
            if ($text -match '(?i)enter the code\s+([A-Z0-9]{6,16})') {
                $deviceCode = $Matches[1].ToUpperInvariant()
            }
            if ($text -match 'https?://\S+') {
                $url = $Matches[0].TrimEnd('.,)')
                if ($url -match '(?i)device|login\.microsoft') { $deviceUrl = $url }
            }
            $trim = $text.Trim()
            if ($trim -and $trim -notmatch '^(?i)(transcript|windows powershell|\*+|start time|end time|machine:|username:)') {
                $message = $trim
            }
        }
    }
    return [pscustomobject]@{
        DeviceCode = $deviceCode
        DeviceUrl  = $deviceUrl
        Message    = $message
    }
}

function Handle-ApiRequest {
    param(
        [string]$Method,
        [string]$Path,
        [object]$Body,
        [Uri]$RequestUrl = $null,
        [switch]$IsLocal
    )

    if ($Method -eq 'POST' -and $Path -eq '/api/shutdown') {
        if (-not $IsLocal) { throw 'Shutdown is only allowed from this machine.' }
        $script:ShutdownRequested = $true
        Stop-BulkWebAllBrowserSessions
        return @{ ok = $true; shuttingDown = $true }
    }

    if ($Method -eq 'GET' -and $Path -eq '/api/lab-users') {
        return @{ ok = $true; users = @(Get-LabUsers) }
    }
    if ($Method -eq 'POST' -and $Path -eq '/api/lab-users') {
        if (-not $Body) { throw 'Missing person.' }
        $saved = Save-LabUser -Name ([string]$Body.name) -Email ([string]$Body.email) -MemberIdentifier ([string]$Body.memberIdentifier) -Id ([string]$Body.id)
        return @{ ok = $true; user = $saved; users = @(Get-LabUsers) }
    }
    if ($Method -eq 'POST' -and $Path -eq '/api/lab-users/delete') {
        if (-not $Body) { throw 'Missing person.' }
        Remove-LabUser -Id ([string]$Body.id)
        return @{ ok = $true; users = @(Get-LabUsers) }
    }

    if ($Method -eq 'GET' -and $Path -eq '/api/health') {
        return @{
            ok       = $true
            version  = '0.5.2'
            pid      = $PID
            projectRoot = $script:ProjectRoot
            hiddenWorkers = (-not $ShowWorkers)
            listenLan = $script:ListenLan
            port = $script:ListenPort
            activeBrowsers = $script:BrowserContexts.Count
            lanUrls = if ($script:ListenLan -and $script:LanUrls) { @($script:LanUrls) } else { @() }
            features = @{
                sessionHistory        = $true
                sessionHistoryActions = $true
                reportSelections      = $true
                exportPresets         = $true
                noWaitCommands        = $true
                multiUserSessions     = $true
                appRegPermissionCheck = $true
                reportZipDownload     = $true
                hiddenWorkers         = $true
                wcmManagement         = $true
                workerLogTabs         = $true
                liongardIntegration   = $true
                huntressIntegration   = $true
                sentinelOneIntegration = $true
                investigatePacks      = $true
                investigateAccessGraph = $true
                # Linux/lab and browser-fail paths use device-code (https://microsoft.com/devicelogin).
                deviceCodeAuth        = -not [System.Runtime.InteropServices.RuntimeInformation]::IsOSPlatform([System.Runtime.InteropServices.OSPlatform]::Windows)
                isWindowsHost         = [System.Runtime.InteropServices.RuntimeInformation]::IsOSPlatform([System.Runtime.InteropServices.OSPlatform]::Windows)
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

    if ($Method -eq 'POST' -and $Path -eq '/api/curate/facets') {
        if (-not $Body -or [string]::IsNullOrWhiteSpace([string]$Body.path)) { throw 'Missing path' }
        Import-Module (Join-Path $script:ProjectRoot 'Modules\LogCuration.psm1') -Force -ErrorAction Stop
        $top = 40
        if ($Body.topValues) { $top = [int]$Body.topValues }
        $wanTop = 12
        if ($Body.wanTop) { $wanTop = [int]$Body.wanTop }
        $result = Get-LogCurationFacetsWithWan -Path ([string]$Body.path) -TopValues $top -WanTop $wanTop
        return @{ ok = $true; result = $result }
    }

    if ($Method -eq 'POST' -and $Path -eq '/api/curate/wan-ips') {
        if (-not $Body -or [string]::IsNullOrWhiteSpace([string]$Body.path)) { throw 'Missing path' }
        Import-Module (Join-Path $script:ProjectRoot 'Modules\LogCuration.psm1') -Force -ErrorAction Stop
        $wanTop = 12
        if ($Body.top) { $wanTop = [int]$Body.top }
        $result = Get-LogCurationWanIpSuggestions -Path ([string]$Body.path) -Top $wanTop
        return @{ ok = $true; result = $result }
    }

    if ($Method -eq 'POST' -and $Path -eq '/api/curate/preview') {
        if (-not $Body -or [string]::IsNullOrWhiteSpace([string]$Body.path)) { throw 'Missing path' }
        Import-Module (Join-Path $script:ProjectRoot 'Modules\LogCuration.psm1') -Force -ErrorAction Stop
        $mode = if ($Body.mode) { [string]$Body.mode } else { 'exclude' }
        $rules = @()
        if ($Body.rules) { $rules = @($Body.rules) }
        $result = Invoke-LogCurationPreview -Path ([string]$Body.path) -Mode $mode -Rules $rules
        return @{ ok = $true; result = $result }
    }

    if ($Method -eq 'POST' -and $Path -eq '/api/curate/export') {
        if (-not $Body -or [string]::IsNullOrWhiteSpace([string]$Body.path)) { throw 'Missing path' }
        Import-Module (Join-Path $script:ProjectRoot 'Modules\LogCuration.psm1') -Force -ErrorAction Stop
        $mode = if ($Body.mode) { [string]$Body.mode } else { 'exclude' }
        $rules = @()
        if ($Body.rules) { $rules = @($Body.rules) }
        $result = Export-LogCurationSet -Path ([string]$Body.path) -Mode $mode -Rules $rules
        return @{ ok = $true; result = $result }
    }

    if ($Method -eq 'GET' -and $Path -eq '/api/investigate/packs') {
        Import-Module (Join-Path $script:ProjectRoot 'Modules\LogInvestigate.psm1') -Force -ErrorAction Stop
        $query = ''
        $sessionId = ''
        $limit = 80
        if ($RequestUrl -and -not [string]::IsNullOrWhiteSpace($RequestUrl.Query)) {
            $qs = $RequestUrl.Query
            if ($qs -match '[?&]q=([^&]*)') { $query = [System.Net.WebUtility]::UrlDecode($Matches[1]) }
            if ($qs -match '[?&]session=([^&]*)') { $sessionId = [System.Net.WebUtility]::UrlDecode($Matches[1]) }
            if ($qs -match '[?&]limit=(\d+)') { $limit = [int]$Matches[1] }
        }
        $hints = @(Get-BulkRunnerKnownReportFolderHints -Session $script:BulkSession)
        $result = Get-InvestigatePacks -Hints $hints -Query $query -SessionId $sessionId -Limit $limit
        return @{ ok = $true; result = $result }
    }

    if ($Path -like '/api/investigate/*') {
        Import-Module (Join-Path $script:ProjectRoot 'Modules\LogInvestigate.psm1') -Force -ErrorAction Stop
        if (-not $Body -or [string]::IsNullOrWhiteSpace([string]$Body.path)) { throw 'Missing path' }
        $invPath = [string]$Body.path
        $invFilters = @()
        if ($Body.filters) { $invFilters = @($Body.filters) }

        if ($Method -eq 'POST' -and $Path -eq '/api/investigate/sources') {
            $result = Get-InvestigateSources -Path $invPath
            return @{ ok = $true; result = $result }
        }
        if ($Method -eq 'POST' -and $Path -eq '/api/investigate/query') {
            $source = [string]$Body.source
            if ([string]::IsNullOrWhiteSpace($source)) { throw 'Missing source' }
            $limit = 50
            if ($Body.limit) { $limit = [int]$Body.limit }
            $offset = 0
            if ($Body.offset) { $offset = [int]$Body.offset }
            $sortBy = [string]$Body.sortBy
            $ascending = $true
            if ($null -ne $Body.ascending) { $ascending = [bool]$Body.ascending }
            $columns = [string]$Body.columns
            $result = Invoke-InvestigateQuery -Path $invPath -Source $source -Filters $invFilters -Columns $columns -SortBy $sortBy -Ascending $ascending -Limit $limit -Offset $offset
            return @{ ok = $true; result = $result }
        }
        if ($Method -eq 'POST' -and $Path -eq '/api/investigate/mail-graph') {
            $minCount = 1
            if ($Body.minCount) { $minCount = [int]$Body.minCount }
            $maxNodes = 80
            if ($Body.maxNodes) { $maxNodes = [int]$Body.maxNodes }
            $hideInternal = $false
            if ($null -ne $Body.hideInternal) { $hideInternal = [bool]$Body.hideInternal }
            $bucket = 'hour'
            if ($Body.bucket) { $bucket = [string]$Body.bucket }
            $graphParams = @{
                Path      = $invPath
                Filters   = $invFilters
                MinCount  = $minCount
                MaxNodes  = $maxNodes
                Bucket    = $bucket
            }
            if ($hideInternal) { $graphParams.HideInternal = $true }
            $result = Get-InvestigateMailGraph @graphParams
            return @{ ok = $true; result = $result }
        }
        if ($Method -eq 'POST' -and $Path -eq '/api/investigate/access-graph') {
            $minCount = 1
            if ($Body.minCount) { $minCount = [int]$Body.minCount }
            $maxNodes = 80
            if ($Body.maxNodes) { $maxNodes = [int]$Body.maxNodes }
            $includeSignIn = $true
            if ($null -ne $Body.includeSignIn) { $includeSignIn = [bool]$Body.includeSignIn }
            $includeUal = $true
            if ($null -ne $Body.includeUal) { $includeUal = [bool]$Body.includeUal }
            $hidePrivateIp = $true
            if ($null -ne $Body.hidePrivateIp) { $hidePrivateIp = [bool]$Body.hidePrivateIp }
            $bucket = 'hour'
            if ($Body.bucket) { $bucket = [string]$Body.bucket }
            $result = Get-InvestigateAccessGraph -Path $invPath -Filters $invFilters -MinCount $minCount -MaxNodes $maxNodes -IncludeSignIn $includeSignIn -IncludeUal $includeUal -HidePrivateIp $hidePrivateIp -Bucket $bucket
            return @{ ok = $true; result = $result }
        }
        if ($Method -eq 'POST' -and $Path -eq '/api/investigate/export') {
            $source = [string]$Body.source
            if ([string]::IsNullOrWhiteSpace($source)) { throw 'Missing source' }
            $sortBy = [string]$Body.sortBy
            $ascending = $true
            if ($null -ne $Body.ascending) { $ascending = [bool]$Body.ascending }
            $columns = [string]$Body.columns
            $result = Export-InvestigateQuery -Path $invPath -Source $source -Filters $invFilters -Columns $columns -SortBy $sortBy -Ascending $ascending
            return @{ ok = $true; result = $result }
        }
        throw "Unknown investigate route: $Path"
    }

    if ($Path -like '/api/triage/*') {
        Import-Module (Join-Path $script:ProjectRoot 'Modules\EmailTriage.psm1') -Force -ErrorAction Stop
        $notesRoot = Join-Path $script:ProjectRoot 'ClientNotes'
        $aiSettingsPath = Join-Path $script:ProjectRoot 'web-runner\ai-review.settings.json'
        if ($Method -eq 'GET' -and $Path -eq '/api/triage/ai-settings') {
            return @{ ok = $true; result = (Get-TriageAiSettings -Path $aiSettingsPath) }
        }
        if ($Method -eq 'POST' -and $Path -eq '/api/triage/ai-settings') {
            $saved = Set-TriageAiSettings -Path $aiSettingsPath -Provider ([string]$Body.provider) -Model ([string]$Body.model) -ApiKey ([string]$Body.apiKey)
            return @{ ok = $true; result = $saved }
        }
        if ($Method -eq 'POST' -and $Path -eq '/api/triage/bec-review') {
            if (-not $Body -or [string]::IsNullOrWhiteSpace([string]$Body.mailbox)) { throw 'Missing mailbox' }
            return @{ ok = $true; result = (Invoke-TriageBecReview -SettingsPath $aiSettingsPath -Facts $Body) }
        }
        if ($Method -eq 'POST' -and $Path -eq '/api/triage/bec-review-start') {
            if (-not $Body -or [string]::IsNullOrWhiteSpace([string]$Body.mailbox)) { throw 'Missing mailbox' }
            return @{ ok = $true; job = (Start-TriageBecReviewJob -SettingsPath $aiSettingsPath -Facts $Body) }
        }
        if ($Method -eq 'GET' -and $Path -eq '/api/triage/bec-review-job') {
            $jobId = Get-TriageRequestQueryValue -RequestUrl $RequestUrl -Name 'id'
            $job = Get-TriageReviewJobSnapshot -JobId $jobId
            if (-not $job) { throw 'Review job was not found.' }
            return @{ ok = $true; job = $job }
        }
        if ($Method -eq 'POST' -and $Path -eq '/api/triage/log-plan') {
            if (-not $Body) { throw 'Missing ticket' }
            $phase = [string]$Body.phase
            if ($phase -ne 'followup') { $phase = 'initial' }
            return @{ ok = $true; result = (Invoke-TriageLogPlan -SettingsPath $aiSettingsPath -Ticket ([string]$Body.ticket) -Phase $phase -Folder ([string]$Body.folder)) }
        }
        if ($Method -eq 'POST' -and $Path -eq '/api/triage/memberberry-note') {
            if (-not $Body) { throw 'Finish a review before writing the ticket note.' }
            $audience = [string]$Body.audience
            if ([string]::IsNullOrWhiteSpace($audience)) { $audience = 'all' }
            return @{ ok = $true; result = (Invoke-TriageMemberberryNote -SettingsPath $aiSettingsPath -Review ([string]$Body.review) -Ticket ([string]$Body.ticket) -TicketNumber ([string]$Body.ticketNumber) -Company ([string]$Body.company) -Contact ([string]$Body.contact) -Audience $audience -Actions $Body.actions -Revision ([string]$Body.revision) -CurrentDrafts ([string]$Body.currentDrafts) -TechnicianName ([string]$Body.technicianName)) }
        }
        if ($Method -eq 'POST' -and $Path -eq '/api/triage/memberberry-save') {
            if (-not $Body) { throw 'Enter the note to add before saving.' }
            $saved = Add-TriageMemberberryClientNote -Client ([string]$Body.client) -Note ([string]$Body.note)
            return @{ ok = $true; result = $saved }
        }
        if ($Method -eq 'POST' -and $Path -eq '/api/triage/memory-save') {
            if (-not $Body) { throw 'Enter the memory to add before saving.' }
            $saved = Add-TriageMemoryVaultNote -Text ([string]$Body.text) -Space 'soc'
            return @{ ok = $true; result = $saved }
        }
        if ($Method -eq 'GET' -and $Path -eq '/api/triage/packs') {
            return @{ ok = $true; result = @(Get-TriagePacks) }
        }
        if ($Method -eq 'GET' -and $Path -eq '/api/triage/notes') {
            $client = Get-TriageRequestQueryValue -RequestUrl $RequestUrl -Name 'client'
            if ([string]::IsNullOrWhiteSpace($client)) { throw 'Missing client' }
            return @{ ok = $true; result = (Get-TriageClientNotes -NotesRoot $notesRoot -Client $client) }
        }
        if ($Method -eq 'POST' -and $Path -eq '/api/triage/notes') {
            if (-not $Body -or [string]::IsNullOrWhiteSpace([string]$Body.client)) { throw 'Missing client' }
            $saved = Set-TriageClientNotes -NotesRoot $notesRoot -Client ([string]$Body.client) -Notes $Body
            return @{ ok = $true; result = $saved }
        }
        $folder = ''
        if ($Method -eq 'GET') { $folder = Get-TriageRequestQueryValue -RequestUrl $RequestUrl -Name 'folder' }
        elseif ($Body) { $folder = [string]$Body.folder }
        $folder = Resolve-TriageWebFolder -Folder $folder
        if ($Method -eq 'POST' -and $Path -eq '/api/triage/ai-review') {
            $client = ''
            if ($Body -and $Body.client) { $client = [string]$Body.client }
            $vips = @()
            $extra = @()
            if ($client) {
                $notes = Get-TriageClientNotes -NotesRoot $notesRoot -Client $client
                $vips = @($notes.vips)
                $extra = @($notes.extraFreemail)
            }
            $company = Split-Path (Split-Path $folder -Parent) -Leaf
            $view = Get-TriagePackView -Folder $folder -VipNames $vips -InternalDomains @() -ExtraFreemail $extra -Company $company -Limit 80
            $review = Invoke-TriageAiReview -SettingsPath $aiSettingsPath -View $view
            return @{ ok = $true; result = $review }
        }
        if ($Method -eq 'GET' -and $Path -eq '/api/triage/file') {
            $name = Get-TriageRequestQueryValue -RequestUrl $RequestUrl -Name 'name'
            $file = Get-TriagePackFileContent -Folder $folder -Name $name
            return @{ ok = $true; result = $file }
        }
        if ($Method -eq 'POST' -and $Path -eq '/api/triage/override') {
            if (-not $Body -or [string]::IsNullOrWhiteSpace([string]$Body.rowKey)) { throw 'Missing rowKey' }
            Set-TriageVerdictOverride -Folder $folder -RowKey ([string]$Body.rowKey) -Verdict ([string]$Body.verdict) | Out-Null
        }
        if ($Method -eq 'GET' -and $Path -eq '/api/triage/result' -or ($Method -eq 'POST' -and $Path -eq '/api/triage/override')) {
            $client = ''
            if ($Method -eq 'GET') { $client = Get-TriageRequestQueryValue -RequestUrl $RequestUrl -Name 'client' }
            elseif ($Body -and $Body.client) { $client = [string]$Body.client }
            $vips = @()
            $extra = @()
            if ($client) {
                $notes = Get-TriageClientNotes -NotesRoot $notesRoot -Client $client
                $vips = @($notes.vips)
                $extra = @($notes.extraFreemail)
            }
            $company = Split-Path (Split-Path $folder -Parent) -Leaf
            $limit = 500
            $limitToken = $null
            $verdict = ''
            if ($Method -eq 'GET') {
                $limitToken = Get-TriageRequestQueryValue -RequestUrl $RequestUrl -Name 'limit'
                $verdict = Get-TriageRequestQueryValue -RequestUrl $RequestUrl -Name 'verdict'
            }
            elseif ($Body) {
                if ($Body.limit) { $limitToken = [string]$Body.limit }
                if ($Body.verdictFilter) { $verdict = [string]$Body.verdictFilter }
            }
            if ($limitToken) { $limit = [int]$limitToken }
            $view = Get-TriagePackView -Folder $folder -VipNames $vips -InternalDomains @() -ExtraFreemail $extra -Company $company -Limit $limit -Verdict $verdict
            return @{ ok = $true; result = $view }
        }
        throw "Unknown triage route: $Path"
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
            try { Stop-BulkRunnerSessionTenants -Session $script:BulkSession } catch { }
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
        $liveTemp = if ($manifest.tempDir) { [string]$manifest.tempDir } else { '' }
        $liveId = if ($manifest.sessionId) { [string]$manifest.sessionId } else { [string]$Body.sessionId }
        if (Test-BulkWebLiveSessionConflict -SessionId $liveId -TempDir $liveTemp) {
            throw 'That session is open in another browser.'
        }
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
        $checkPermissions = $false
        if ($RequestUrl -and -not [string]::IsNullOrWhiteSpace($RequestUrl.Query)) {
            $forceRefresh = ($RequestUrl.Query -match '(?:^|[&?])(?:refresh|forceRefresh)=1(?:&|$)')
            $checkPermissions = ($RequestUrl.Query -match '(?:^|[&?])(?:checkPermissions|checkPerms)=1(?:&|$)')
        }
        $regParams = @{
            ProjectRoot     = $script:ProjectRoot
            SkipGraphLookup = (-not $forceRefresh)
        }
        if ($forceRefresh) { $regParams['ForceRefreshFromGraph'] = $true }
        if ($checkPermissions) { $regParams['CheckPermissions'] = $true }
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

        if ($Method -eq 'GET' -and $Path -eq '/api/wcm/graph-app-job') {
            if (-not $script:GraphAppJob) { return @{ running = $false } }
            $proc = $script:GraphAppJob.Process
            $running = $false
            try { $running = ($null -ne $proc -and -not $proc.HasExited) } catch { $running = $false }
            $jobView = Get-GraphAppJobPublicView -LogPath $script:GraphAppJob.LogPath
            if ($running) {
                return @{
                    running    = $true
                    processId  = $script:GraphAppJob.ProcessId
                    logPath    = $script:GraphAppJob.LogPath
                    updated    = [bool]$script:GraphAppJob.UpdateExisting
                    deviceCode = $jobView.DeviceCode
                    deviceUrl  = $jobView.DeviceUrl
                    message    = $jobView.Message
                }
            }
            $result = $null
            if ($script:GraphAppJob.ResultPath -and (Test-Path -LiteralPath $script:GraphAppJob.ResultPath)) {
                try { $result = Get-Content -LiteralPath $script:GraphAppJob.ResultPath -Raw -ErrorAction Stop | ConvertFrom-Json } catch { }
            }
            $exitCode = $null
            try { $exitCode = $proc.ExitCode } catch { }
            return @{
                running    = $false
                processId  = $script:GraphAppJob.ProcessId
                exitCode   = $exitCode
                result     = $result
                logPath    = $script:GraphAppJob.LogPath
                updated    = [bool]$script:GraphAppJob.UpdateExisting
                deviceCode = $jobView.DeviceCode
                deviceUrl  = $jobView.DeviceUrl
                message    = $jobView.Message
            }
        }

        if ($Method -eq 'POST' -and $Path -eq '/api/wcm/create-graph-app') {
            $tid = if ($Body.tenantId) { [string]$Body.tenantId } else { '' }
            if ($script:GraphAppJob -and $script:GraphAppJob.Process) {
                $busy = $false
                try { $busy = -not $script:GraphAppJob.Process.HasExited } catch { }
                if ($busy) {
                    $jobView = Get-GraphAppJobPublicView -LogPath $script:GraphAppJob.LogPath
                    return @{
                        started        = $true
                        alreadyRunning = $true
                        processId      = $script:GraphAppJob.ProcessId
                        deviceCode     = $jobView.DeviceCode
                        deviceUrl      = $jobView.DeviceUrl
                        message        = 'A sign-in is already waiting. Use the code at the top of this page.'
                    }
                }
            }
            $jobId = if ($script:RequestBrowserId) { [string]$script:RequestBrowserId } else { [guid]::NewGuid().ToString('D') }
            $outcome = Invoke-GraphAppCreateWithWcmSave -ProjectRoot $script:ProjectRoot -TenantId $tid -NoWait -JobId $jobId
            $script:GraphAppJob = $outcome
            $script:GraphAppJob | Add-Member -NotePropertyName UpdateExisting -NotePropertyValue $false -Force
            return @{
                started    = $true
                processId  = $outcome.ProcessId
                logPath    = $outcome.LogPath
                deviceCode = $outcome.UserCode
                deviceUrl  = $outcome.VerificationUri
            }
        }

        if ($Method -eq 'POST' -and $Path -eq '/api/wcm/update-graph-app-scopes') {
            $tid = if ($Body.tenantId) { [string]$Body.tenantId } else { '' }
            if ($script:GraphAppJob -and $script:GraphAppJob.Process) {
                $busy = $false
                try { $busy = -not $script:GraphAppJob.Process.HasExited } catch { }
                if ($busy) {
                    $jobView = Get-GraphAppJobPublicView -LogPath $script:GraphAppJob.LogPath
                    return @{
                        started        = $true
                        alreadyRunning = $true
                        processId      = $script:GraphAppJob.ProcessId
                        deviceCode     = $jobView.DeviceCode
                        deviceUrl      = $jobView.DeviceUrl
                        message        = 'A sign-in is already waiting. Use the code at the top of this page.'
                        updated        = $true
                    }
                }
            }
            $jobId = if ($script:RequestBrowserId) { [string]$script:RequestBrowserId } else { [guid]::NewGuid().ToString('D') }
            $outcome = Invoke-GraphAppCreateWithWcmSave -ProjectRoot $script:ProjectRoot -TenantId $tid -UpdateExisting -NoWait -JobId $jobId
            $script:GraphAppJob = $outcome
            $script:GraphAppJob | Add-Member -NotePropertyName UpdateExisting -NotePropertyValue $true -Force
            return @{
                started    = $true
                processId  = $outcome.ProcessId
                logPath    = $outcome.LogPath
                deviceCode = $outcome.UserCode
                deviceUrl  = $outcome.VerificationUri
                updated    = $true
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
            if (-not $Body.password) { throw 'Missing password' }
            $sec = ConvertTo-SecureString ([string]$Body.password) -AsPlainText -Force
            $prefix = if ($Body.prefix) { [string]$Body.prefix } else { 'EOA' }
            $path = if ($Body.path) { [string]$Body.path } else { '' }
            $wantBrowse = ($Body.browse -eq $true) -or ($Body.browse -eq 'true')
            # Default is browser download. Old clients that only sent a typed path still write on the runner.
            $wantDownload = ($Body.download -eq $true) -or ($Body.download -eq 'true') -or (
                -not $wantBrowse -and [string]::IsNullOrWhiteSpace($path)
            )
            $onWindows = [System.Runtime.InteropServices.RuntimeInformation]::IsOSPlatform([System.Runtime.InteropServices.OSPlatform]::Windows)
            $fileName = 'GraphAppCredentials.eoa-creds'

            if ($wantDownload) {
                $payload = New-GraphAppCredentialsExportPayload -Password $sec -Prefix $prefix
                return @{
                    ok       = $true
                    download = $true
                    count    = [int]$payload.Count
                    fileName = $fileName
                    content  = [string]$payload.Content
                }
            }

            if ($wantBrowse) {
                if (-not $onWindows) {
                    throw 'File browse dialog is only available on Windows. Use browser download instead.'
                }
                $initialDir = ''
                try {
                    $downloads = Join-Path ([Environment]::GetFolderPath('UserProfile')) 'Downloads'
                    if ($downloads -and (Test-Path -LiteralPath $downloads)) { $initialDir = $downloads }
                    elseif ($script:ProjectRoot -and (Test-Path -LiteralPath $script:ProjectRoot)) { $initialDir = $script:ProjectRoot }
                } catch { }
                $path = Show-GraphAppCredentialSaveFileDialog -DefaultFileName $fileName -InitialDirectory $initialDir
                if ([string]::IsNullOrWhiteSpace($path)) {
                    return @{ ok = $false; cancelled = $true }
                }
            }

            if ([string]::IsNullOrWhiteSpace($path)) {
                throw 'Missing export path. Use download=true for a browser save, or browse=true on Windows.'
            }
            $count = Export-GraphAppCredentialsToFile -Path $path -Password $sec -Prefix $prefix
            return @{ ok = $true; path = $path; count = $count; browsed = [bool]$wantBrowse; download = $false }
        }

        if ($Method -eq 'POST' -and $Path -eq '/api/wcm/import') {
            if (-not $Body.password) { throw 'Missing password' }
            $sec = ConvertTo-SecureString ([string]$Body.password) -AsPlainText -Force
            $importParams = @{ Password = $sec }
            if (-not [string]::IsNullOrWhiteSpace([string]$Body.fileText)) {
                $importParams.FileText = [string]$Body.fileText
            }
            elseif (-not [string]::IsNullOrWhiteSpace([string]$Body.path)) {
                $importParams.Path = [string]$Body.path
            }
            else {
                throw 'Choose an .eoa-creds file.'
            }
            $count = Import-GraphAppCredentialsFromFile @importParams
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
            success          = $true
            ticketId         = $fetched.TicketId
            summary          = $fetched.Summary
            companyName      = $fetched.CompanyName
            contactName      = $fetched.ContactName
            statusName       = $fetched.StatusName
            contactEmail     = $fetched.ContactEmail
            automaticEmailCc = $fetched.AutomaticEmailCc
            emailContact     = [bool]$fetched.EmailContact
            emailCc          = [bool]$fetched.EmailCc
            ticketNumbers    = @($fetched.TicketNumbers)
            ticketContent    = $fetched.TicketContent
            contentLength    = $fetched.FilteredLength
            securityStack    = $securityStack
        }
    }

    if ($Method -eq 'GET' -and $Path -eq '/api/manage/security-tickets') {
        Import-Module (Join-Path $script:ProjectRoot 'Modules\Settings.psm1') -Force -ErrorAction SilentlyContinue
        Import-Module (Join-Path $script:ProjectRoot 'Modules\ConnectWiseManage.psm1') -Force -ErrorAction Stop
        $catalog = Get-ConnectWiseManageHunterCatalog
        $tickets = @(Get-ConnectWiseManageSecurityQueue)
        return @{
            ok      = $true
            catalog = @{
                boardId          = $catalog.BoardId
                boardName        = $catalog.BoardName
                memberIdentifier = $catalog.MemberIdentifier
                memberName       = $catalog.MemberName
                statuses         = @($catalog.Statuses | ForEach-Object { @{ key = $_.Key; id = $_.Id; name = $_.Name } })
                workTypes        = @($catalog.WorkTypes | ForEach-Object { @{ key = $_.Key; id = $_.Id; name = $_.Name; billableOption = $_.BillableOption } })
            }
            tickets = $tickets
        }
    }

    if ($Method -eq 'POST' -and $Path -eq '/api/manage/hunter-post') {
        if (-not $Body) { throw 'Missing note request.' }
        Import-Module (Join-Path $script:ProjectRoot 'Modules\Settings.psm1') -Force -ErrorAction SilentlyContinue
        Import-Module (Join-Path $script:ProjectRoot 'Modules\ConnectWiseManage.psm1') -Force -ErrorAction Stop
        $posted = Send-ConnectWiseManageHunterPost -Request $Body
        return @{
            ok               = [bool]$posted.ok
            ticketId         = $posted.ticketId
            memberIdentifier = $posted.memberIdentifier
            statusName       = $posted.statusName
            statusUpdated    = [bool]$posted.statusUpdated
            notes            = $posted.notes
            timeEntryId      = $posted.timeEntryId
            completed        = @($posted.completed)
            errors           = @($posted.errors)
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
            # Always noWait: blocking wait freezes the single-threaded listener for every browser.
            $cmd = Expand-BulkRunnerGenerateReportsCommand -Session $script:BulkSession -ClientNumber $clientNumber -Command $cmd
            $response = Send-BulkRunnerCommand -Session $script:BulkSession -ClientNumber $clientNumber -Command $cmd -TimeoutSeconds 0 -NoWait
            Sync-BulkRunnerSessionManifest -Session $script:BulkSession | Out-Null
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

    if ($Method -eq 'GET' -and $Path -match '^/api/tenants/(\d+)/report-zips$') {
        if (-not $script:BulkSession) { throw 'No session.' }
        $clientNumber = [int]$Matches[1]
        $zips = @(Get-BulkRunnerTenantReportZipFiles -Session $script:BulkSession -ClientNumber $clientNumber)
        $folder = $null
        if ($script:BulkSession.Tenants.ContainsKey([string]$clientNumber)) {
            $folder = $script:BulkSession.Tenants[[string]$clientNumber].OutputFolder
        }
        return @{
            outputFolder = $folder
            zips         = @($zips)
            count        = $zips.Count
        }
    }

    if ($Method -eq 'GET' -and $Path -match '^/api/tenants/(\d+)/report-zip$') {
        if (-not $script:BulkSession) { throw 'No session.' }
        $clientNumber = [int]$Matches[1]
        $name = ''
        if ($RequestUrl -and -not [string]::IsNullOrWhiteSpace($RequestUrl.Query)) {
            if ($RequestUrl.Query -match '[?&]name=([^&]*)') {
                $name = [System.Net.WebUtility]::UrlDecode($Matches[1])
            }
        }
        $zipPath = Resolve-BulkRunnerTenantReportZipPath -Session $script:BulkSession -ClientNumber $clientNumber -Name $name
        $leaf = [System.IO.Path]::GetFileName($zipPath)
        return @{
            __fileDownload = $true
            path           = $zipPath
            contentType    = 'application/zip'
            fileName       = $leaf
        }
    }

    if ($Path -match '^/api/tenants/(\d+)/containment-pack$' -and $Method -eq 'POST') {
        if (-not $script:BulkSession) { throw 'No session.' }
        $clientNumber = [int]$Matches[1]
        if (-not $script:BulkSession.Tenants.ContainsKey([string]$clientNumber)) {
            throw "Unknown tenant client number: $clientNumber"
        }
        $result = Export-BulkRunnerContainmentPacks -Session $script:BulkSession -ClientNumber $clientNumber `
            -OutputFolder ([string]$Body.outputFolder) -CompanyName ([string]$Body.companyName) `
            -Packs @($Body.packs) -Actions @($Body.actions)
        Sync-BulkRunnerSessionManifest -Session $script:BulkSession -Force | Out-Null
        return @{
            ok       = $true
            folder   = $result.folder
            files    = @($result.files)
            auditCsv = $result.auditCsv
        }
    }

    if ($Path -match '^/api/tenants/(\d+)/copy-containment-zips$' -and $Method -eq 'POST') {
        if (-not $script:BulkSession) { throw 'No session.' }
        $from = [string]$Body.from
        $to = [string]$Body.to
        if ([string]::IsNullOrWhiteSpace($from) -or [string]::IsNullOrWhiteSpace($to)) { throw 'Missing from or to' }
        $copied = @(Copy-BulkRunnerContainmentZips -FromFolder $from -ToFolder $to)
        return @{ ok = $true; files = $copied }
    }

    throw "Not found: $Method $Path"
}

function Set-BulkRunnerExplorerForeground {
    # Explorer opens behind the browser because this process is not the foreground app, so
    # Windows blocks SetForegroundWindow. Tapping ALT gives this process input credit, which
    # satisfies the foreground-change rules. Best effort only - never fail the open.
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [int]$TimeoutMs = 4000
    )

    try {
        if (-not ('EoaWin.Focus' -as [type])) {
            Add-Type -Namespace EoaWin -Name Focus -MemberDefinition @'
[DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
[DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
[DllImport("user32.dll")] public static extern void SwitchToThisWindow(IntPtr hWnd, bool fAltTab);
[DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
[DllImport("user32.dll")] public static extern void keybd_event(byte vk, byte scan, uint flags, UIntPtr extra);
'@
        }

        $wanted = ([System.IO.Path]::GetFullPath($Path)).TrimEnd('\')
        $deadline = [DateTime]::UtcNow.AddMilliseconds($TimeoutMs)
        $hwnd = [IntPtr]::Zero

        while ($hwnd -eq [IntPtr]::Zero -and [DateTime]::UtcNow -lt $deadline) {
            $shell = $null
            try {
                $shell = New-Object -ComObject Shell.Application
                $windows = $shell.Windows()
                for ($i = 0; $i -lt $windows.Count; $i++) {
                    $w = $windows.Item($i)
                    if (-not $w) { continue }
                    try {
                        $wPath = $w.Document.Folder.Self.Path
                        if ($wPath -and ([System.IO.Path]::GetFullPath($wPath)).TrimEnd('\') -ieq $wanted) {
                            $hwnd = [IntPtr]$w.HWND
                            break
                        }
                    } catch {}
                }
            } catch {} finally {
                if ($shell) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($shell) }
            }
            if ($hwnd -eq [IntPtr]::Zero) { Start-Sleep -Milliseconds 250 }
        }

        if ($hwnd -eq [IntPtr]::Zero) { return $false }

        [void][EoaWin.Focus]::ShowWindowAsync($hwnd, 9)  # SW_RESTORE
        [EoaWin.Focus]::keybd_event(0xA4, 0, 0, [UIntPtr]::Zero)  # LMENU down
        [EoaWin.Focus]::keybd_event(0xA4, 0, 2, [UIntPtr]::Zero)  # LMENU up
        if (-not [EoaWin.Focus]::SetForegroundWindow($hwnd)) {
            [EoaWin.Focus]::SwitchToThisWindow($hwnd, $true)
        }
        return ([EoaWin.Focus]::GetForegroundWindow() -eq $hwnd)
    } catch {
        return $false
    }
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

    $docs = [Environment]::GetFolderPath('MyDocuments')
    $docsRoot = ''
    if (-not [string]::IsNullOrWhiteSpace($docs)) {
        $docsRoot = [System.IO.Path]::GetFullPath($docs)
    }
    $tempRoot = [System.IO.Path]::GetFullPath([string]$env:TEMP)
    $underDocs = (-not [string]::IsNullOrWhiteSpace($docsRoot)) -and $resolved.StartsWith($docsRoot, [StringComparison]::OrdinalIgnoreCase)
    $allowed = $underDocs -or ($resolved.StartsWith($tempRoot, [StringComparison]::OrdinalIgnoreCase))
    if (-not $allowed -and -not $AllowTempRoot) {
        throw 'Only folders under Documents or TEMP can be opened from the web UI.'
    }
    if (-not $allowed -and $AllowTempRoot -and -not $resolved.StartsWith($tempRoot, [StringComparison]::OrdinalIgnoreCase)) {
        throw 'Session temp folder must be under TEMP.'
    }

    $onWindows = [System.Runtime.InteropServices.RuntimeInformation]::IsOSPlatform([System.Runtime.InteropServices.OSPlatform]::Windows)
    if (-not $onWindows) { return $resolved }
    Start-Process explorer.exe -ArgumentList "`"$resolved`""
    [void](Set-BulkRunnerExplorerForeground -Path $resolved)
    return $resolved
}

function Test-BulkWebRunnerLanFirewallRule {
    param([Parameter(Mandatory = $true)][int]$Port)

    $ruleName = "EOA Bulk Web Runner (TCP $Port)"
    try {
        $rule = Get-NetFirewallRule -DisplayName $ruleName -ErrorAction Stop |
            Select-Object -First 1
        return ($null -ne $rule -and $rule.Enabled -eq 'True')
    } catch {
        return $false
    }
}

function Ensure-BulkWebRunnerLanFirewallRule {
    param([Parameter(Mandatory = $true)][int]$Port)

    $onWindows = [System.Runtime.InteropServices.RuntimeInformation]::IsOSPlatform([System.Runtime.InteropServices.OSPlatform]::Windows)
    if (-not $onWindows) { return $true }
    if (Test-BulkWebRunnerLanFirewallRule -Port $Port) { return $true }

    $ruleName = "EOA Bulk Web Runner (TCP $Port)"
    $netshCmd = "netsh advfirewall firewall add rule name=`"$ruleName`" dir=in action=allow protocol=TCP localport=$Port profile=domain,private,public enable=yes"
    $null = & netsh advfirewall firewall add rule name="$ruleName" dir=in action=allow protocol=TCP localport=$Port profile=domain,private,public enable=yes 2>&1
    if ($LASTEXITCODE -eq 0) { return $true }

    Write-Host "Windows Firewall is blocking inbound TCP $Port from other PCs." -ForegroundColor Yellow
    Write-Host "Run once in elevated PowerShell, then retry from the remote PC:" -ForegroundColor Yellow
    Write-Host "  $netshCmd" -ForegroundColor Cyan
    return $false
}

function Get-BulkWebRunnerLanAddresses {
    $addrs = New-Object System.Collections.Generic.List[string]
    try {
        Get-NetIPAddress -AddressFamily IPv4 -ErrorAction Stop |
            Where-Object {
                $_.IPAddress -and
                $_.IPAddress -notlike '127.*' -and
                $_.IPAddress -notlike '169.254.*' -and
                $_.InterfaceAlias -notmatch 'vEthernet|Hyper-V|WSL|Docker|Loopback'
            } |
            ForEach-Object { if ($addrs -notcontains $_.IPAddress) { [void]$addrs.Add($_.IPAddress) } }
    } catch {
        foreach ($addr in [System.Net.Dns]::GetHostAddresses([System.Net.Dns]::GetHostName())) {
            if ($addr.AddressFamily -ne [System.Net.Sockets.AddressFamily]::InterNetwork) { continue }
            $ip = $addr.IPAddressToString
            if ($ip -like '127.*' -or $ip -like '169.254.*') { continue }
            if ($addrs -notcontains $ip) { [void]$addrs.Add($ip) }
        }
    }
    return @($addrs)
}

function Test-BulkWebRunnerUrlAcl {
    param([Parameter(Mandatory = $true)][string]$UrlPrefix)
    try {
        $out = netsh http show urlacl url=$UrlPrefix 2>$null | Out-String
        return ($out -match [regex]::Escape($UrlPrefix))
    } catch {
        return $false
    }
}

function Get-BulkWebRunnerPortHolderPids {
    param([Parameter(Mandatory = $true)][int]$Port)
    $pids = New-Object System.Collections.Generic.List[int]
    try {
        $lines = netsh http show servicestate view=requestq verbose=yes 2>$null | Out-String
        # Match "ID: 12345, image: ...pwsh.exe" near a registered URL for this port.
        $portToken = ":$Port/"
        $chunks = $lines -split 'Request queue name:'
        foreach ($chunk in $chunks) {
            if ($chunk -notmatch [regex]::Escape($portToken) -and $chunk -notmatch [regex]::Escape(":$Port`:")) { continue }
            foreach ($m in [regex]::Matches($chunk, 'ID:\s*(\d+),\s*image:\s*([^\r\n]+)')) {
                $procId = [int]$m.Groups[1].Value
                $image = [string]$m.Groups[2].Value
                if ($image -match '(?i)pwsh|powershell') {
                    if (-not $pids.Contains($procId)) { [void]$pids.Add($procId) }
                }
            }
        }
    } catch { }
    return @($pids)
}

function Get-BulkWebRunnerHealth {
    param([Parameter(Mandatory = $true)][int]$Port)
    $stale503 = $false
    foreach ($hostName in @('127.0.0.1', 'localhost')) {
        try {
            return Invoke-RestMethod -Uri "http://${hostName}:$Port/api/health" -TimeoutSec 2 -ErrorAction Stop
        } catch {
            $msg = [string]$_.Exception.Message
            $code = 0
            try { $code = [int]$_.Exception.Response.StatusCode } catch { }
            if ($msg -match '503' -or $code -eq 503) { $stale503 = $true }
        }
    }
    if ($stale503) { return [pscustomobject]@{ ok = $false; _stale503 = $true } }
    return $null
}

function Test-IsOurBulkWebRunnerHealth {
    param($Health)
    if (-not $Health) { return $false }
    if ($Health.ok -and $Health.projectRoot -and [string]$Health.projectRoot -eq [string]$script:ProjectRoot) { return $true }
    if ($Health.ok -and $Health.features -and $Health.features.wcmManagement) { return $true }
    return $false
}

function Get-BulkWebRunnerScriptPids {
    param([Parameter(Mandatory = $true)][int]$Port)
    $found = New-Object System.Collections.Generic.List[int]
    $procs = @()
    try { $procs += @(Get-CimInstance Win32_Process -Filter "Name = 'pwsh.exe'" -ErrorAction SilentlyContinue) } catch { }
    try { $procs += @(Get-CimInstance Win32_Process -Filter "Name = 'powershell.exe'" -ErrorAction SilentlyContinue) } catch { }
    foreach ($proc in $procs) {
        if (-not $proc) { continue }
        $procId = [int]$proc.ProcessId
        if ($procId -eq $PID) { continue }
        $cl = [string]$proc.CommandLine
        if ($cl -notmatch 'Start-BulkWebRunner\.ps1') { continue }
        if ($cl -match '(?i)[-\/]Port\s+(\d+)') {
            if ([int]$Matches[1] -ne $Port) { continue }
        } elseif ($Port -ne 8765) {
            continue
        }
        if (-not $found.Contains($procId)) { [void]$found.Add($procId) }
    }
    return @($found)
}

function Stop-BulkWebRunnerExisting {
    param([Parameter(Mandatory = $true)][int]$Port)
    $health = Get-BulkWebRunnerHealth -Port $Port
    if ($health -and $health.ok -and -not (Test-IsOurBulkWebRunnerHealth -Health $health)) {
        throw "Port $Port is already in use by another application (not this web runner)."
    }

    if (Test-IsOurBulkWebRunnerHealth -Health $health) {
        Write-Host "Existing web runner on port $Port — requesting shutdown..." -ForegroundColor Yellow
        try {
            $null = Invoke-RestMethod -Uri "http://127.0.0.1:$Port/api/shutdown" -Method POST -TimeoutSec 8 -ErrorAction Stop
        } catch { }
        $until = [datetime]::UtcNow.AddSeconds(8)
        while ([datetime]::UtcNow -lt $until) {
            Start-Sleep -Milliseconds 400
            $again = Get-BulkWebRunnerHealth -Port $Port
            if (-not $again -or -not $again.ok) { break }
        }
    }

    $targets = New-Object System.Collections.Generic.List[int]
    foreach ($procId in @($(if ($health -and $health.pid) { [int]$health.pid } else { 0 })) + @(Get-BulkWebRunnerPortHolderPids -Port $Port) + @(Get-BulkWebRunnerScriptPids -Port $Port)) {
        if ($procId -le 0 -or $procId -eq $PID) { continue }
        if (-not $targets.Contains($procId)) { [void]$targets.Add($procId) }
    }
    foreach ($procId in $targets) {
        try {
            $proc = Get-Process -Id $procId -ErrorAction Stop
            Write-Host "Stopping existing web-runner process PID $procId ($($proc.ProcessName)) on port $Port..." -ForegroundColor Yellow
            Stop-Process -Id $procId -Force -ErrorAction Stop
        } catch { }
    }
    if ($targets.Count -gt 0) { Start-Sleep -Seconds 1 }

    $still = Get-BulkWebRunnerHealth -Port $Port
    if ($still -and $still.ok) {
        throw "Port $Port is still in use after stopping the previous web runner."
    }
    if ($still -and $still._stale503) {
        Write-Host "Port $Port still returns HTTP 503 from HTTP.sys. Close other Start-BulkWebRunner windows, then retry." -ForegroundColor Yellow
    }
}

$script:TriageReviewJobs = @{}
$script:TriageReviewCache = @{}

function Get-TriageFactsCacheKey {
    param($Facts)
    $json = ConvertTo-Json -InputObject $Facts -Depth 8 -Compress
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes([string]$json)
        $hash = $sha.ComputeHash($bytes)
        return ([System.BitConverter]::ToString($hash)).Replace('-', '')
    }
    finally { $sha.Dispose() }
}

function Get-TriageReviewJobSnapshot {
    param([string]$JobId)
    if ([string]::IsNullOrWhiteSpace($JobId)) { return $null }
    $job = $script:TriageReviewJobs[$JobId]
    if (-not $job) { return $null }
    $sync = $job.Sync
    if ($sync.done -and $job.Key -and $sync.result -and -not $script:TriageReviewCache.ContainsKey($job.Key)) {
        $script:TriageReviewCache[$job.Key] = $sync.result
    }
    if ($sync.done -and -not $job.Cleaned -and $job.PowerShell) {
        $job.Cleaned = $true
        try { $null = $job.PowerShell.EndInvoke($job.Handle) } catch { }
        try { $job.PowerShell.Dispose() } catch { }
        try { $job.Runspace.Close() } catch { }
        try { $job.Runspace.Dispose() } catch { }
    }
    return @{
        id     = $JobId
        text   = [string]$sync.text
        done   = [bool]$sync.done
        error  = [string]$sync.error
        cached = [bool]$sync.cached
        result = $sync.result
    }
}

function Start-TriageBecReviewJob {
    param(
        [Parameter(Mandatory)][string]$SettingsPath,
        [Parameter(Mandatory)]$Facts
    )
    $key = Get-TriageFactsCacheKey -Facts $Facts
    $id = [guid]::NewGuid().ToString('n')
    if ($script:TriageReviewCache.ContainsKey($key)) {
        $cached = $script:TriageReviewCache[$key]
        $sync = [hashtable]::Synchronized(@{
            text   = [string]$cached.text
            done   = $true
            error  = ''
            cached = $true
            result = $cached
        })
        $script:TriageReviewJobs[$id] = @{ Sync = $sync; Key = $key; Cleaned = $true }
        return (Get-TriageReviewJobSnapshot -JobId $id)
    }
    $factsJson = ConvertTo-Json -InputObject $Facts -Depth 8 -Compress
    $modulePath = Join-Path $script:ProjectRoot 'Modules\EmailTriage.psm1'
    $sync = [hashtable]::Synchronized(@{ text = ''; done = $false; error = ''; cached = $false; result = $null })
    $rs = [runspacefactory]::CreateRunspace()
    $rs.Open()
    $ps = [powershell]::Create()
    $ps.Runspace = $rs
    $null = $ps.AddScript({
        param($modulePath, $settingsPath, $factsJson, $sync)
        Import-Module $modulePath -Force
        $facts = $factsJson | ConvertFrom-Json
        try {
            $result = Invoke-TriageBecReview -SettingsPath $settingsPath -Facts $facts -OnDelta {
                param($piece)
                $sync.text = [string]$sync.text + [string]$piece
            }
            $sync.result = $result
            if ($result -and $result.text) { $sync.text = [string]$result.text }
        }
        catch {
            $sync.error = $_.Exception.Message
        }
        finally {
            $sync.done = $true
        }
    }).AddArgument($modulePath).AddArgument($SettingsPath).AddArgument($factsJson).AddArgument($sync)
    $handle = $ps.BeginInvoke()
    $script:TriageReviewJobs[$id] = @{
        Sync       = $sync
        Key        = $key
        PowerShell = $ps
        Runspace   = $rs
        Handle     = $handle
        Cleaned    = $false
    }
    if ($script:TriageReviewJobs.Count -gt 40) {
        foreach ($oldId in @($script:TriageReviewJobs.Keys)) {
            $old = $script:TriageReviewJobs[$oldId]
            if ($old.Sync.done -and $oldId -ne $id) { $script:TriageReviewJobs.Remove($oldId) }
        }
    }
    return (Get-TriageReviewJobSnapshot -JobId $id)
}

function Start-BulkWebRunnerHttpListener {
    param(
        [Parameter(Mandatory = $true)][int]$Port,
        [switch]$ListenLan
    )

    $plusPrefix = "http://+:$Port/"
    # Prefer +:PORT whenever that reservation exists. Localhost-only binds against a +: ACL
    # often "start" successfully then serve HTTP.sys 503 forever.
    if (-not $ListenLan -and (Test-BulkWebRunnerUrlAcl -UrlPrefix $plusPrefix)) {
        Write-Host "Detected URL reservation $plusPrefix — using it (prevents HTTP 503 on localhost-only bind)." -ForegroundColor Yellow
        $script:ListenLan = $true
        $ListenLan = $true
    }

    Stop-BulkWebRunnerExisting -Port $Port

    $onWindows = [System.Runtime.InteropServices.RuntimeInformation]::IsOSPlatform([System.Runtime.InteropServices.OSPlatform]::Windows)
    $prefixes = @("http://127.0.0.1:$Port/")
    if ($ListenLan) {
        if ($onWindows) { $prefixes = @($plusPrefix) }
        else { $prefixes = @("http://*:$Port/", $plusPrefix) }
    }
    $listener = $null
    $started = $false
    $lastError = $null
    foreach ($prefix in $prefixes) {
        $listener = New-Object System.Net.HttpListener
        $listener.Prefixes.Add($prefix)
        try {
            $listener.Start()
            $started = $true
            break
        }
        catch {
            $lastError = $_
            try { $listener.Close() } catch {}
            $listener = $null
        }
    }
    if ($started) { return $listener }
    $_ = $lastError
    try { throw $lastError } catch {
        $msg = [string]$_.Exception.Message
        $accessDenied = $msg -match '(?i)Access is denied|Access denied|5\b'

        if (-not $ListenLan -and $accessDenied) {
            try { $listener.Close() } catch {}
            if (Test-BulkWebRunnerUrlAcl -UrlPrefix $plusPrefix) {
                Write-Host "Localhost bind access denied; falling back to $plusPrefix." -ForegroundColor Yellow
                $script:ListenLan = $true
                return (Start-BulkWebRunnerHttpListener -Port $Port -ListenLan)
            }
            $user = "$env:USERDOMAIN\$env:USERNAME"
            throw @"
Bind failed for http://127.0.0.1:$Port/ (Access denied).

Start with -ListenLan (recommended):

  .\web-runner\Start-BulkWebRunner.ps1 -ListenLan

Or reserve localhost once in elevated PowerShell:

  netsh http add urlacl url=http://127.0.0.1:$Port/ user="$user" listen=yes

Original error: $msg
"@
        }

        if (-not $ListenLan) { throw }

        $user = "$env:USERDOMAIN\$env:USERNAME"
        $netsh = "netsh http add urlacl url=$plusPrefix user=`"$user`" listen=yes"
        throw @"
LAN bind failed ($plusPrefix). Run this once in an elevated PowerShell, then retry -ListenLan:

  $netsh

Original error: $msg
"@
    }
}

$script:LanUrls = @()
$localPrefix = "http://127.0.0.1:$Port/"
$listener = Start-BulkWebRunnerHttpListener -Port $Port -ListenLan:$ListenLan
# Listener may auto-enable LAN mode when localhost ACL is missing but http://+:PORT/ is reserved.
$ListenLan = [bool]$script:ListenLan
if ($ListenLan) {
    $script:LanUrls = @(Get-BulkWebRunnerLanAddresses | ForEach-Object { "http://${_}:$Port/" })
}

Write-Host "Bulk Web Runner listening on $localPrefix" -ForegroundColor Green
if ($ListenLan) {
    Write-Host "LAN mode enabled (all interfaces on port $Port)" -ForegroundColor Yellow
    if ($script:LanUrls.Count -gt 0) {
        Write-Host "Open from other PCs on your network:" -ForegroundColor Cyan
        foreach ($lanUrl in $script:LanUrls) {
            Write-Host "  $lanUrl" -ForegroundColor Cyan
        }
    } else {
        Write-Host "Could not detect LAN IPv4 addresses; use http://<this-pc-ip>:$Port/" -ForegroundColor Yellow
    }
    $null = Ensure-BulkWebRunnerLanFirewallRule -Port $Port
    Write-Host "Auth popups still run on this Windows machine only." -ForegroundColor Gray
}
Write-Host "Project root: $script:ProjectRoot" -ForegroundColor Gray
Write-Host "Press Ctrl+C to stop." -ForegroundColor Gray

if (-not $NoBrowser) {
    Start-Process $localPrefix
}

try {
    while ($listener.IsListening) {
        $context = $listener.GetContext()
        $request = $context.Request
        $response = $context.Response
        $path = $request.Url.AbsolutePath
        $setCookie = $null
        $browserCtx = $null
        $bindBrowser = $false

        try {
            if ($path.StartsWith('/api/')) {
                $body = Read-RequestBody -Request $request
                $bindBrowser = ($path -ne '/api/health')
                if ($bindBrowser) {
                    $cookieHeader = [string]$request.Headers['Cookie']
                    $browserId = Get-BulkWebBrowserIdFromCookie -CookieHeader $cookieHeader
                    if (-not $browserId) {
                        $browserId = New-BulkWebBrowserId
                        $setCookie = Get-BulkWebBrowserCookieHeader -BrowserId $browserId
                    }
                    $browserCtx = Ensure-BulkWebBrowserContext -BrowserId $browserId
                    $script:RequestBrowserId = $browserId
                    $script:BulkSession = $browserCtx.Session
                    $script:GraphAppJob = $browserCtx.GraphAppJob
                    $browserCtx.LastAccess = (Get-Date).ToUniversalTime()
                }

                try {
                    $result = Handle-ApiRequest -Method $request.HttpMethod -Path $path -Body $body -RequestUrl $request.Url -IsLocal:$request.IsLocal
                } finally {
                    if ($bindBrowser) {
                        # Persist this browser's session/job even if the handler threw.
                        $browserCtx.Session = $script:BulkSession
                        $browserCtx.GraphAppJob = $script:GraphAppJob
                        $browserCtx.LastAccess = (Get-Date).ToUniversalTime()
                        $script:RequestBrowserId = $null
                        $script:BulkSession = $null
                        $script:GraphAppJob = $null
                    }
                }

                if ($result -and $result.__fileDownload -eq $true) {
                    $filePath = [string]$result.path
                    if (-not $filePath -or -not (Test-Path -LiteralPath $filePath)) {
                        throw "Download file not found: $filePath"
                    }
                    $bytes = [System.IO.File]::ReadAllBytes($filePath)
                    $ctype = if ($result.contentType) { [string]$result.contentType } else { 'application/octet-stream' }
                    $fname = if ($result.fileName) { [string]$result.fileName } else { [System.IO.Path]::GetFileName($filePath) }
                    $disp = 'attachment; filename="{0}"' -f ($fname -replace '"', '')
                    Write-ByteResponse -Response $response -Bytes $bytes -ContentType $ctype -ContentDisposition $disp -SetCookie $setCookie
                } else {
                    $jsonDepth = if ($path -like '/api/curate/*' -or $path -like '/api/investigate/*' -or $path -like '/api/triage/*') { 12 } else { 8 }
                    Write-JsonResponse -Response $response -Body $result -Depth $jsonDepth -SetCookie $setCookie
                }
                if ($script:ShutdownRequested) {
                    Write-Host "Shutdown requested - stopping listener." -ForegroundColor Yellow
                    try { $listener.Stop() } catch { }
                    break
                }
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
                        '.svg'  { 'image/svg+xml' }
                        default { 'application/octet-stream' }
                    }
                    $response.AddHeader('Cache-Control', 'no-cache, no-store, must-revalidate')
                    $bytes = [System.IO.File]::ReadAllBytes($file)
                    Write-ByteResponse -Response $response -Bytes $bytes -ContentType $ctype
                }
            }
        } catch {
            try {
                Write-JsonResponse -Response $response -Body @{ error = $_.Exception.Message } -StatusCode 400 -SetCookie $setCookie
            } catch {
                Write-Host "Response error: $($_.Exception.Message)" -ForegroundColor Yellow
                try { $response.Close() } catch { }
            }
        }
    }
} finally {
    try { Stop-BulkWebAllBrowserSessions } catch { }
    $listener.Stop()
    $listener.Close()
}
