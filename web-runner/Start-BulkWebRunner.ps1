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
#>
param(
    [int]$Port = 8765,
    [switch]$NoBrowser
)

$ErrorActionPreference = 'Stop'
$script:RunnerRoot = $PSScriptRoot
$script:ProjectRoot = Split-Path $PSScriptRoot -Parent
Import-Module (Join-Path $script:RunnerRoot 'Modules\BulkRunnerSession.psm1') -Force
$script:BulkSession = $null

function Write-JsonResponse {
    param([System.Net.HttpListenerResponse]$Response, [object]$Body, [int]$StatusCode = 200)
    $Response.StatusCode = $StatusCode
    $Response.ContentType = 'application/json; charset=utf-8'
    $json = $Body | ConvertTo-Json -Depth 8 -Compress
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($json)
    $Response.ContentLength64 = $bytes.Length
    $Response.OutputStream.Write($bytes, 0, $bytes.Length)
    $Response.OutputStream.Close()
}

function Write-TextResponse {
    param(
        [System.Net.HttpListenerResponse]$Response,
        [string]$Body,
        [string]$ContentType = 'text/plain; charset=utf-8',
        [int]$StatusCode = 200
    )
    $Response.StatusCode = $StatusCode
    $Response.ContentType = $ContentType
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($Body)
    $Response.ContentLength64 = $bytes.Length
    $Response.OutputStream.Write($bytes, 0, $bytes.Length)
    $Response.OutputStream.Close()
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
        [object]$Body
    )

    if ($Method -eq 'GET' -and $Path -eq '/api/health') {
        return @{ ok = $true; version = '0.1.0'; projectRoot = $script:ProjectRoot }
    }

    if ($Method -eq 'POST' -and $Path -eq '/api/session') {
        $selections = @{}
        if ($Body.reportSelections) {
            $Body.reportSelections.PSObject.Properties | ForEach-Object { $selections[$_.Name] = $_.Value }
        } else {
            $selections = @{
                IncludeMessageTrace = $true
                IncludeInboxRules   = $true
                SignInLogsDaysBack  = 7
                MessageTraceDaysBack = 7
            }
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

    if ($Method -eq 'GET' -and $Path -eq '/api/session') {
        if (-not $script:BulkSession) { return $null }
        return (Get-BulkRunnerSessionSummary -Session $script:BulkSession)
    }

    if ($Method -eq 'GET' -and $Path -eq '/api/app-registrations') {
        return @(Get-BulkRunnerAppRegistrations -ProjectRoot $script:ProjectRoot)
    }

    if ($Method -eq 'POST' -and $Path -eq '/api/tenants') {
        if (-not $script:BulkSession) { throw 'No session. POST /api/session first.' }
        $tenant = Add-BulkRunnerTenant -Session $script:BulkSession
        return @{
            clientNumber = $tenant.ClientNumber
            processId    = $tenant.ProcessId
        }
    }

    if ($Path -match '^/api/tenants/(\d+)/(command|status)$') {
        if (-not $script:BulkSession) { throw 'No session.' }
        $clientNumber = [int]$Matches[1]
        $action = $Matches[2]

        if ($action -eq 'status' -and $Method -eq 'GET') {
            return @{ status = (Get-BulkRunnerTenantStatus -Session $script:BulkSession -ClientNumber $clientNumber) }
        }

        if ($action -eq 'command' -and $Method -eq 'POST') {
            $cmd = [string]$Body.command
            if ([string]::IsNullOrWhiteSpace($cmd)) { throw 'Missing command' }
            $wait = 60
            if ($Body.waitSeconds) { $wait = [int]$Body.waitSeconds }
            $response = Send-BulkRunnerCommand -Session $script:BulkSession -ClientNumber $clientNumber `
                -Command $cmd -TimeoutSeconds $wait
            return @{ response = $response }
        }
    }

    throw "Not found: $Method $Path"
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
                $result = Handle-ApiRequest -Method $request.HttpMethod -Path $path -Body $body
                if ($null -eq $result -and $path -eq '/api/session') {
                    Write-JsonResponse -Response $response -Body @{ sessionId = $null } -StatusCode 200
                } else {
                    Write-JsonResponse -Response $response -Body $result
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
                        default { 'application/octet-stream' }
                    }
                    $bytes = [System.IO.File]::ReadAllBytes($file)
                    $response.StatusCode = 200
                    $response.ContentType = $ctype
                    $response.ContentLength64 = $bytes.Length
                    $response.OutputStream.Write($bytes, 0, $bytes.Length)
                    $response.OutputStream.Close()
                }
            }
        } catch {
            Write-JsonResponse -Response $response -Body @{ error = $_.Exception.Message } -StatusCode 400
        }
    }
} finally {
    $listener.Stop()
    $listener.Close()
}
