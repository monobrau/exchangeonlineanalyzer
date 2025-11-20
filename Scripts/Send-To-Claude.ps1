param(
    [Parameter(Mandatory=$false)][string]$ApiKey,
    [Parameter(Mandatory=$false)][string]$OutputFolder,
    [Parameter(Mandatory=$false)][string]$Model = 'claude-3-5-sonnet-20241022',
    [Parameter(Mandatory=$false)][string[]]$ExtraFiles,
    [Parameter(Mandatory=$false)][string]$ResponseFile = 'Claude_Response.md',
    [Parameter(Mandatory=$false)][int]$MaxCsvRows = 0,
    [Parameter(Mandatory=$false)][switch]$VerboseOutput
)

function Get-LatestInvestigationFolder {
    $base = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'ExchangeOnlineAnalyzer\SecurityInvestigation'
    if (-not (Test-Path $base)) { return $null }
    $candidates = @()
    $tenants = Get-ChildItem -Path $base -Directory -ErrorAction SilentlyContinue
    foreach ($t in $tenants) {
        $runs = Get-ChildItem -Path $t.FullName -Directory -ErrorAction SilentlyContinue
        if ($runs) { $candidates += $runs }
    }
    $legacy = Get-ChildItem -Path $base -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -match '^\d{8}_\d{6}$' }
    if ($legacy) { $candidates += $legacy }
    if (-not $candidates -or $candidates.Count -eq 0) { return $null }
    return ($candidates | Sort-Object LastWriteTime -Descending | Select-Object -First 1).FullName
}

if (-not $OutputFolder -or -not (Test-Path $OutputFolder)) { $OutputFolder = Get-LatestInvestigationFolder }

$settingsPath = Join-Path $PSScriptRoot '..\Modules\Settings.psm1'
try { if (Test-Path $settingsPath) { Import-Module $settingsPath -Force -ErrorAction SilentlyContinue } } catch {}
if (-not $ApiKey) { try { $s = Get-AppSettings; if ($s -and $s.ClaudeApiKey) { $ApiKey = $s.ClaudeApiKey } } catch {} }
if (-not $ApiKey) { Write-Error 'Claude API key is required. Save it in Settings or pass -ApiKey.'; exit 1 }

if (-not $OutputFolder -or -not (Test-Path $OutputFolder)) { Write-Error 'Could not find Security Investigation output folder. Pass -OutputFolder or generate a report.'; exit 1 }
if ($VerboseOutput) { Write-Host ("Using output folder: {0}" -f $OutputFolder) -ForegroundColor DarkGray }

function Get-MimeType([string]$path) {
    switch -regex ($path) {
        '\.csv$' { 'text/csv'; break }
        '\.txt$' { 'text/plain'; break }
        default { 'application/octet-stream' }
    }
}

function Get-AttachPath { param([string]$path,[int]$maxRows)
    if ($maxRows -gt 0 -and $path -match '\.csv$') {
        $linesNeeded = $maxRows + 1 # header + rows
        try {
            $tmp = [System.IO.Path]::GetTempFileName()
            Get-Content -Path $path -TotalCount $linesNeeded | Set-Content -Path $tmp -Encoding utf8
            return $tmp
        } catch { return $path }
    }
    return $path
}

$defaults = @('_AI_Readme.txt','MessageTrace.csv','InboxRules.csv','TransportRules.csv','MailFlowConnectors.csv','GraphAuditLogs.csv','UserSecurityPosture.csv')
$files = @()
foreach ($f in $defaults) { $p = Join-Path $OutputFolder $f; if (Test-Path $p) { $files += $p } }
if ($ExtraFiles) {
    foreach ($ef in $ExtraFiles) { if ($ef -and (Test-Path $ef)) { $files += (Resolve-Path $ef).Path } }
}
$files = $files | Select-Object -Unique
if ($files.Count -eq 0) { Write-Error 'No files to send.'; exit 1 }

if ($VerboseOutput) {
    Write-Host ("Attaching {0} file(s):" -f $files.Count) -ForegroundColor DarkGray
    foreach ($p in $files) { Write-Host (" - {0} ({1} KB)" -f $p,[Math]::Round(((Get-Item $p).Length/1KB),0)) -ForegroundColor DarkGray }
}

# Build Anthropic messages format
$contentParts = @()
$intro = @"
Please analyze the attached security investigation datasets and produce:
- Executive summary (non-technical)
- Timeline of events with timestamps and sources
- Evidence-backed findings
- Minimal immediate actions
Do not speculate beyond provided evidence. Use CSV content explicitly in references.
"@
$contentParts += @{ type = 'text'; text = $intro }

foreach ($p in $files) {
    $attach = Get-AttachPath -path $p -maxRows $MaxCsvRows
    if ((Get-Item $attach).Length -gt 20000000) { Write-Warning "File >20MB, skipping: $attach"; continue }
    $name = [System.IO.Path]::GetFileName($p)
    $text = try { Get-Content -Path $attach -Raw -Encoding UTF8 } catch { '' }
    if (-not $text) { $text = "(empty file)" }
    $maxChars = 300000
    if ($text.Length -gt $maxChars) { $text = $text.Substring(0,$maxChars) + "`n...[truncated]" }
    $contentParts += @{ type='text'; text=("=== {0} ===`n{1}" -f $name, $text) }
}

$req = @{ model = $Model; max_tokens = 2048; messages = @(@{ role = 'user'; content = $contentParts }) }
$json = $req | ConvertTo-Json -Depth 8
$reqPath = Join-Path $OutputFolder 'Claude_Request.json'
$json | Out-File -FilePath $reqPath -Encoding utf8
if ($VerboseOutput) { Write-Host ("Saved request JSON: {0}" -f $reqPath) -ForegroundColor DarkGray }

$headers = @{ 'x-api-key' = $ApiKey; 'anthropic-version' = '2023-06-01'; 'Content-Type' = 'application/json' }
$url = 'https://api.anthropic.com/v1/messages'

$resp = $null; $err = $null
try {
    Write-Host ("Submitting to Claude: {0}" -f $url) -ForegroundColor DarkGray
    $resp = Invoke-RestMethod -Method POST -Uri $url -Headers $headers -Body $json -ErrorAction Stop
} catch { $err = $_ }
if ($err) {
    $errPath = Join-Path $OutputFolder 'Claude_Error.txt'
    $lines = @()
    $lines += "Endpoint: $url"
    try { $sc = $err.Exception.Response.StatusCode; $sv = $err.Exception.Response.StatusCode.value__; $lines += "Status: $sv ($sc)" } catch {}
    if ($err.Exception.Message) { $lines += "Exception: $($err.Exception.Message)" }
    if ($err.ErrorDetails -and $err.ErrorDetails.Message) { $lines += "ErrorDetails: $($err.ErrorDetails.Message)" }
    try {
        $body = $null
        if ($err.Exception.Response -is [System.Net.Http.HttpResponseMessage]) { $body = $err.Exception.Response.Content.ReadAsStringAsync().Result }
        elseif ($err.Exception.Response -and $err.Exception.Response.GetResponseStream) { $reader = New-Object System.IO.StreamReader($err.Exception.Response.GetResponseStream()); $body = $reader.ReadToEnd() }
        if ($body) { $lines += 'Body:'; $lines += $body }
    } catch {}
    ($lines -join "`r`n") | Out-File -FilePath $errPath -Encoding utf8
    Write-Error ("Claude request failed. See {0}" -f $errPath)
    exit 1
}

$text = $null
try {
    if ($resp.content -and $resp.content.Count -gt 0) {
        $parts = $resp.content | Where-Object { $_.type -eq 'text' }
        if ($parts -and $parts[0].text) { $text = $parts[0].text }
    }
} catch {}
if (-not $text) { $text = ($resp | ConvertTo-Json -Depth 8) }

$out = Join-Path $OutputFolder $ResponseFile
$text | Out-File -FilePath $out -Encoding utf8
Write-Host ("Saved Claude response to: {0}" -f $out) -ForegroundColor Green


