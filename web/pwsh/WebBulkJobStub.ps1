# Web API worker: reads JSON payload, writes summary.json, ReportSelections.json (desktop-compatible shape).
# Full EXO + interactive BulkExportWorker.ps1 remains a desktop flow; this stub prepares artifacts for parity testing.

param(
    [Parameter(Mandatory)][string]$PayloadJsonPath,
    [Parameter(Mandatory)][string]$JobId,
    [Parameter(Mandatory)][string]$OutputDir
)
$ErrorActionPreference = 'Stop'
$script:WorkerVersion = '3'

New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null

$raw = Get-Content -Raw -LiteralPath $PayloadJsonPath -Encoding UTF8
$payload = $raw | ConvertFrom-Json

$tenantIds = @()
if ($payload.tenant_ids) { $tenantIds = @($payload.tenant_ids) }

$repoRoot = $env:EOA_REPO_ROOT
if (-not $repoRoot) { $repoRoot = $null }

$opt = $payload.options
if ($null -eq $opt) { $opt = [pscustomobject]@{} }

function Get-OptionBool {
    param([string]$SnakeName, [bool]$Default = $false)
    $p = $opt.PSObject.Properties[$SnakeName]
    if (-not $p) { return $Default }
    $v = $p.Value
    if ($null -eq $v) { return $Default }
    return [bool]$v
}

function Get-OptionInt {
    param([string]$SnakeName, [int]$Default = 10)
    $p = $opt.PSObject.Properties[$SnakeName]
    if (-not $p) { return $Default }
    $v = $p.Value
    if ($null -eq $v) { return $Default }
    try { return [int]$v } catch { return $Default }
}

# Same keys as BulkTenantExporter.ps1 -> ReportSelections.json (BulkExportWorker.ps1 / New-SecurityInvestigationReport).
$reportSelections = [ordered]@{
    IncludeMessageTrace                 = (Get-OptionBool 'include_message_trace')
    IncludeInboxRules                   = (Get-OptionBool 'include_inbox_rules')
    IncludeTransportRules               = (Get-OptionBool 'include_transport_rules')
    IncludeMailFlowConnectors           = (Get-OptionBool 'include_mail_flow_connectors')
    IncludeMailboxForwarding            = (Get-OptionBool 'include_mailbox_forwarding')
    IncludeAuditLogs                    = (Get-OptionBool 'include_audit_logs')
    IncludeConditionalAccessPolicies    = (Get-OptionBool 'include_conditional_access_policies')
    IncludeAppRegistrations             = (Get-OptionBool 'include_app_registrations')
    IncludeSignInLogs                   = (Get-OptionBool 'include_sign_in_logs')
    IncludeMfaCoverage                  = (Get-OptionBool 'include_mfa_coverage')
    IncludeSharePointActivity           = (Get-OptionBool 'include_share_point_activity')
    IncludeOneDriveActivity             = (Get-OptionBool 'include_one_drive_activity')
    IncludeTeamsActivity                = (Get-OptionBool 'include_teams_activity')
    IncludeSharePointSharing            = (Get-OptionBool 'include_share_point_sharing')
    IncludeSecurityAlerts               = (Get-OptionBool 'include_security_alerts')
    IncludeSecurityIncidents            = (Get-OptionBool 'include_security_incidents')
    IncludeAnonymousSharePointSharing   = (Get-OptionBool 'include_anonymous_share_point_sharing')
    IncludeSharePointFileSharingLinks   = (Get-OptionBool 'include_share_point_file_sharing_links')
    IncludeDLPViolations                = (Get-OptionBool 'include_dlp_violations')
    IncludeIntuneDevices                = (Get-OptionBool 'include_intune_devices')
    IncludeUnifiedAuditLogs             = (Get-OptionBool 'include_unified_audit_logs')
    IncludeSharePointOneDriveFileActions = (Get-OptionBool 'include_share_point_one_drive_file_actions')
    SignInLogsDaysBack                  = (Get-OptionInt 'sign_in_logs_days_back' 7)
    MessageTraceDaysBack                = (Get-OptionInt 'message_trace_days_back' (Get-OptionInt 'days_back' 10))
}

$rsPath = Join-Path $OutputDir 'ReportSelections.json'
($reportSelections | ConvertTo-Json -Depth 6) | Set-Content -LiteralPath $rsPath -Encoding UTF8

$summary = [ordered]@{
    workerVersion       = $script:WorkerVersion
    jobId               = $JobId
    tenantCount         = $tenantIds.Count
    tenantIdsSample     = ($tenantIds | Select-Object -First 5)
    options             = $payload.options
    reportSelectionsPath = 'ReportSelections.json'
    repoRootEnv         = $repoRoot
    message             = 'Web pwsh worker: wrote summary.json and ReportSelections.json (desktop-compatible). For full EXO+Graph export use BulkTenantExporter.ps1 with BulkExportWorker.ps1 on a workstation.'
    psVersion           = $PSVersionTable.PSVersion.ToString()
    os                  = [System.Runtime.InteropServices.RuntimeInformation]::OSDescription
    at                  = (Get-Date).ToUniversalTime().ToString('o')
}

$summaryFileName = if ($env:EOA_PWSH_SUMMARY_NAME -and $env:EOA_PWSH_SUMMARY_NAME.Trim()) {
    $env:EOA_PWSH_SUMMARY_NAME.Trim()
} else {
    'summary.json'
}
$jsonPath = Join-Path $OutputDir $summaryFileName
($summary | ConvertTo-Json -Depth 10) | Set-Content -LiteralPath $jsonPath -Encoding UTF8

"OK wrote $jsonPath and $rsPath"
