# Web API worker: reads JSON payload from the API and writes an artifact folder.
# Next: optional fork/exec of Scripts/BulkExportWorker.ps1 or ExportUtils when running on Windows with Graph.

param(
    [Parameter(Mandatory)][string]$PayloadJsonPath,
    [Parameter(Mandatory)][string]$JobId,
    [Parameter(Mandatory)][string]$OutputDir
)
$ErrorActionPreference = 'Stop'
$script:WorkerVersion = '2'

New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null

$raw = Get-Content -Raw -LiteralPath $PayloadJsonPath -Encoding UTF8
$payload = $raw | ConvertFrom-Json

$tenantIds = @()
if ($payload.tenant_ids) { $tenantIds = @($payload.tenant_ids) }

$repoRoot = $env:EOA_REPO_ROOT
if (-not $repoRoot) { $repoRoot = $null }

$summary = [ordered]@{
    workerVersion   = $script:WorkerVersion
    jobId           = $JobId
    tenantCount     = $tenantIds.Count
    tenantIdsSample = ($tenantIds | Select-Object -First 5)
    options         = $payload.options
    repoRootEnv     = $repoRoot
    message         = 'Web worker stub: wrote summary.json only. Full bulk export still runs from BulkTenantExporter.ps1 / BulkExportWorker.ps1 on a workstation with Graph/EXO.'
    psVersion       = $PSVersionTable.PSVersion.ToString()
    os              = [System.Runtime.InteropServices.RuntimeInformation]::OSDescription
    at              = (Get-Date).ToUniversalTime().ToString('o')
}

$jsonPath = Join-Path $OutputDir 'summary.json'
($summary | ConvertTo-Json -Depth 10) | Set-Content -LiteralPath $jsonPath -Encoding UTF8

"OK wrote $jsonPath"
