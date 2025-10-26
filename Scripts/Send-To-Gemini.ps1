param(
    [Parameter(Mandatory=$false)][string]$ApiKey,
    [Parameter(Mandatory=$false)][string]$OutputFolder,
    [Parameter(Mandatory=$false)][string]$Model = 'models/gemini-1.5-pro',
    [Parameter(Mandatory=$false)][string[]]$ExtraFiles,
    [Parameter(Mandatory=$false)][string]$ResponseFile = 'Gemini_Response.md',
    [Parameter(Mandatory=$false)][switch]$DebugOutput
)

function Get-LatestInvestigationFolder {
    $base = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'ExchangeOnlineAnalyzer\SecurityInvestigation'
    if (-not (Test-Path $base)) { return $null }
    $tenants = Get-ChildItem -Path $base -Directory | Sort-Object LastWriteTime -Descending
    foreach ($t in $tenants) {
        $runs = Get-ChildItem -Path $t.FullName -Directory | Sort-Object LastWriteTime -Descending
        if ($runs -and $runs.Count -gt 0) { return $runs[0].FullName }
    }
    return $null
}

if (-not $OutputFolder -or -not (Test-Path $OutputFolder)) { $OutputFolder = Get-LatestInvestigationFolder }
$settingsPath = Join-Path $PSScriptRoot '..\Modules\Settings.psm1'
try { if (Test-Path $settingsPath) { Import-Module $settingsPath -Force -ErrorAction SilentlyContinue } } catch {}
if (-not $ApiKey) { try { $s = Get-AppSettings; if ($s -and $s.GeminiApiKey) { $ApiKey = $s.GeminiApiKey } } catch {} }
if (-not $ApiKey) { Write-Error "Gemini API key is required. Provide -ApiKey or save it in Settings."; exit 1 }

if (-not $OutputFolder -or -not (Test-Path $OutputFolder)) { Write-Error "Could not find Security Investigation output folder. Run the report first or pass -OutputFolder."; exit 1 }
Write-Verbose ("Using output folder: {0}" -f $OutputFolder)

$modelName = $Model
if ($modelName -like 'models/*') { $modelName = $modelName.Substring(7) }
$endpoints = @(
    ("https://generativelanguage.googleapis.com/v1/models/{0}:generateContent?key={1}" -f $modelName, $ApiKey),
    ("https://generativelanguage.googleapis.com/v1beta/models/{0}:generateContent?key={1}" -f $modelName, $ApiKey),
    ("https://generativelanguage.googleapis.com/v1beta/{0}:generateContent?key={1}" -f $Model, $ApiKey)
)
Write-Verbose ("Model endpoints (fallback order):`n - {0}`n - {1}`n - {2}" -f $endpoints[0], $endpoints[1], $endpoints[2])

# Collect files to attach
$files = @(
    'LLM_Instructions.txt',
    'MessageTrace.csv',
    'InboxRules.csv',
    'TransportRules.csv',
    'InboundConnectors.csv',
    'OutboundConnectors.csv',
    'GraphAuditLogs.csv',
    'MFAStatus.csv',
    'UserSecurityGroups.csv'
)

$existing = @()
foreach ($f in $files) {
    $p = Join-Path $OutputFolder $f
    if (Test-Path $p) { $existing += $p }
}
if ($ExtraFiles) {
    foreach ($ef in $ExtraFiles) {
        if ([string]::IsNullOrWhiteSpace($ef)) { continue }
        try {
            $rp = (Resolve-Path $ef -ErrorAction Stop).Path
            if (Test-Path $rp) { $existing += $rp }
        } catch {
            Write-Warning ("Extra file not found: {0}" -f $ef)
        }
    }
    $existing = $existing | Select-Object -Unique
}
if ($existing.Count -eq 0) { Write-Error "No known report files found in $OutputFolder."; exit 1 }
Write-Verbose ("Attaching {0} file(s):" -f $existing.Count)
foreach ($p in $existing) { Write-Verbose (" - {0} ({1} KB)" -f $p, [Math]::Round(((Get-Item $p).Length/1KB),0)) }

function Get-MimeType([string]$path) {
    switch -regex ($path) {
        '\.csv$' { 'text/csv'; break }
        '\.txt$' { 'text/plain'; break }
        default { 'application/octet-stream' }
    }
}

# Build request parts. First include a concise instruction tying files together.
$intro = @"
Please analyze the attached security investigation datasets and produce:
- Executive summary (non-technical)
- Timeline of events with timestamps and sources
- Evidence-backed findings
- Minimal immediate actions
Do not speculate beyond provided evidence. Use CSV content explicitly in references.
"@

$parts = @()
$parts += @{ text = $intro }

foreach ($path in $existing) {
    $mime = Get-MimeType $path
    if ((Get-Item $path).Length -gt 20000000) {
        Write-Warning "File exceeds 20MB REST payload limit for inlineData: $path (skipping)."
        continue
    }
    $bytes = [System.IO.File]::ReadAllBytes($path)
    $b64 = [Convert]::ToBase64String($bytes)
    $parts += @{ inlineData = @{ mimeType = $mime; data = $b64 } }
}

$body = @{ contents = @(@{ role = 'user'; parts = $parts }) }
$json = $body | ConvertTo-Json -Depth 6
if ($DebugOutput) { $reqPath = Join-Path $OutputFolder 'Gemini_Request.json'; $json | Out-File -FilePath $reqPath -Encoding utf8; Write-Verbose ("Saved request JSON: {0}" -f $reqPath) }

$resp = $null
$lastErr = $null
foreach ($ep in $endpoints) {
    try {
        Write-Verbose ("Submitting to: {0}" -f $ep)
        $resp = Invoke-RestMethod -Method POST -Uri $ep -ContentType 'application/json' -Body $json -ErrorAction Stop
        break
    } catch {
        $lastErr = $_
        # If 404 Not Found, try next endpoint; otherwise stop early
        $status = $null
        try { $status = $_.Exception.Response.StatusCode.value__ } catch {}
        if ($status -ne 404) { break }
    }
}
if (-not $resp) {
    $errPath = Join-Path $OutputFolder 'Gemini_Error.txt'
    $details = $lastErr.Exception.Message
    try {
        if ($lastErr.Exception.Response -and $lastErr.Exception.Response.GetResponseStream) {
            $reader = New-Object System.IO.StreamReader($lastErr.Exception.Response.GetResponseStream())
            $body = $reader.ReadToEnd()
            if ($body) { $details += "`n`nResponse:`n" + $body }
        }
    } catch {}
    $details | Out-File -FilePath $errPath -Encoding utf8
    Write-Error ("Gemini request failed. See {0}" -f $errPath)
    exit 1
}

$text = $null
try {
    if ($resp.candidates -and $resp.candidates[0].content.parts[0].text) {
        $text = $resp.candidates[0].content.parts[0].text
    } elseif ($resp.candidates -and $resp.candidates[0].content.parts) {
        $texts = @()
        foreach ($p in $resp.candidates[0].content.parts) { if ($p.text) { $texts += $p.text } }
        $text = ($texts -join "`n`n")
    }
} catch {}

if ($DebugOutput) { $respJsonPath = Join-Path $OutputFolder 'Gemini_Response.raw.json'; ($resp | ConvertTo-Json -Depth 8) | Out-File -FilePath $respJsonPath -Encoding utf8; Write-Verbose ("Saved raw response JSON: {0}" -f $respJsonPath) }
if (-not $text) { $text = ($resp | ConvertTo-Json -Depth 8) }

$outFile = Join-Path $OutputFolder $ResponseFile
$text | Out-File -FilePath $outFile -Encoding utf8
Write-Host ("Saved Gemini response to: {0}" -f $outFile) -ForegroundColor Green


