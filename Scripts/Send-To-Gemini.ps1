param(
    [Parameter(Mandatory=$true)][string]$ApiKey,
    [Parameter(Mandatory=$false)][string]$OutputFolder,
    [Parameter(Mandatory=$false)][string]$Model = 'models/gemini-1.5-pro'
)

function Get-LatestInvestigationFolder {
    $base = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'ExchangeOnlineAnalyzer\SecurityInvestigation'
    if (-not (Test-Path $base)) { return $null }
    Get-ChildItem -Path $base -Directory | Sort-Object LastWriteTime -Descending | Select-Object -First 1 | ForEach-Object { $_.FullName }
}

if (-not $OutputFolder -or -not (Test-Path $OutputFolder)) {
    $OutputFolder = Get-LatestInvestigationFolder
}
if (-not $OutputFolder -or -not (Test-Path $OutputFolder)) {
    Write-Error "Could not find Security Investigation output folder. Run the report first or pass -OutputFolder."
    exit 1
}

$endpoint = ("https://generativelanguage.googleapis.com/v1beta/{0}:generateContent?key={1}" -f $Model, $ApiKey)

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
if ($existing.Count -eq 0) {
    Write-Error "No known report files found in $OutputFolder."
    exit 1
}

function Get-MimeType([string]$path) {
    switch -regex ($path) {
        '\.csv$' { 'text/csv'; break }
        '\.txt$' { 'text/plain'; break }
        default { 'application/octet-stream' }
    }
}

# Build request parts. First include a concise instruction tying files together.
$intro = @"Please analyze the attached security investigation datasets and produce:
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

try {
    $resp = Invoke-RestMethod -Method POST -Uri $endpoint -ContentType 'application/json' -Body $json -ErrorAction Stop
} catch {
    Write-Error ("Gemini request failed: {0}" -f $_.Exception.Message)
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

if (-not $text) { $text = ($resp | ConvertTo-Json -Depth 8) }

$outFile = Join-Path $OutputFolder 'Gemini_Response.md'
$text | Out-File -FilePath $outFile -Encoding utf8
Write-Host ("Saved Gemini response to: {0}" -f $outFile) -ForegroundColor Green


