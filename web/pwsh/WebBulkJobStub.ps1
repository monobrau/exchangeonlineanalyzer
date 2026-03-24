# Web API worker stub: reads JSON payload from the API and writes an artifact folder.
# Replace later with real BulkExportWorker.ps1 orchestration + Graph/EXO.
param(
    [Parameter(Mandatory)][string]$PayloadJsonPath,
    [Parameter(Mandatory)][string]$JobId,
    [Parameter(Mandatory)][string]$OutputDir
)
$ErrorActionPreference = 'Stop'

New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null

$raw = Get-Content -Raw -LiteralPath $PayloadJsonPath -Encoding UTF8
$payload = $raw | ConvertFrom-Json

$tenantIds = @()
if ($payload.tenant_ids) { $tenantIds = @($payload.tenant_ids) }

$summary = [ordered]@{
    jobId           = $JobId
    tenantCount     = $tenantIds.Count
    tenantIdsSample = ($tenantIds | Select-Object -First 5)
    options         = $payload.options
    message         = 'Stub worker: artifact only. Wire BulkExportWorker + app credentials next.'
    psVersion       = $PSVersionTable.PSVersion.ToString()
    at              = (Get-Date).ToUniversalTime().ToString('o')
}

$jsonPath = Join-Path $OutputDir 'summary.json'
($summary | ConvertTo-Json -Depth 10) | Set-Content -LiteralPath $jsonPath -Encoding UTF8

"OK wrote $jsonPath"
