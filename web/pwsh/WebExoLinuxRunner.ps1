# Web API worker (Linux/macOS/Windows pwsh): Exchange Online + Graph via same pipeline as desktop
# New-SecurityInvestigationReport (Modules/ExportUtils.psm1). Secrets ONLY from environment — never from job JSON.
#
# Prerequisites (Linux): PowerShell 7+, ExchangeOnlineManagement 3+, Microsoft.Graph.* modules, OpenSSL/WSMan per Microsoft docs.
# Auth: EXO app-only (certificate) + Graph app-only (client secret) — same model as unattended automation.
#
# Required env:
#   EOA_EXO_APP_ID, EOA_EXO_ORGANIZATION (e.g. contoso.onmicrosoft.com), EOA_EXO_CERT_THUMBPRINT
#   EOA_GRAPH_CLIENT_ID, EOA_GRAPH_CLIENT_SECRET (Graph application permissions for collectors)
#   EOA_REPO_ROOT — repo root (parent of web/)
#
# Optional: EOA_EXO_SKIP_CONNECT — if 'true', skip EXO (Graph-only; EXO slices will be empty)
#Requires -Version 7.0

param(
    [Parameter(Mandatory)][string]$PayloadJsonPath,
    [Parameter(Mandatory)][string]$JobId,
    [Parameter(Mandatory)][string]$OutputDir
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:WorkerVersion = '1'

function Test-IsGuid {
    param([string]$s)
    if ([string]::IsNullOrWhiteSpace($s)) { return $false }
    return [guid]::TryParse($s, [ref]([guid]::Empty))
}

function Get-OptionBool {
    param([object]$Opt, [string]$SnakeName, [bool]$Default = $false)
    if ($null -eq $Opt) { return $Default }
    $p = $Opt.PSObject.Properties[$SnakeName]
    if (-not $p) { return $Default }
    $v = $p.Value
    if ($null -eq $v) { return $Default }
    return [bool]$v
}

function Get-OptionInt {
    param([object]$Opt, [string]$SnakeName, [int]$Default = 10)
    if ($null -eq $Opt) { return $Default }
    $p = $Opt.PSObject.Properties[$SnakeName]
    if (-not $p) { return $Default }
    $v = $p.Value
    if ($null -eq $v) { return $Default }
    try { return [int]$v } catch { return $Default }
}

function Get-ClientCredentialsToken {
    param(
        [string]$TenantId,
        [string]$ClientId,
        [string]$ClientSecret
    )
    $uri = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
    $body = @{
        client_id     = $ClientId
        client_secret = $ClientSecret
        scope         = 'https://graph.microsoft.com/.default'
        grant_type    = 'client_credentials'
    }
    return Invoke-RestMethod -Method Post -Uri $uri -Body $body -ContentType 'application/x-www-form-urlencoded'
}

New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null

$raw = Get-Content -Raw -LiteralPath $PayloadJsonPath -Encoding UTF8
$payload = $raw | ConvertFrom-Json

$tenantIds = @()
if ($payload.tenant_ids) { $tenantIds = @($payload.tenant_ids | ForEach-Object { "$_".Trim() }) | Where-Object { $_ }
}

foreach ($tid in $tenantIds) {
    if (-not (Test-IsGuid $tid)) {
        throw "Invalid tenant_id (expected GUID): $tid"
    }
}

if ($tenantIds.Count -lt 1) {
    throw 'job payload must include at least one tenant_id (Entra directory GUID).'
}

if ($tenantIds.Count -gt 1) {
    Write-Warning "Multiple tenant_ids in payload; EXO app-only session uses the first only: $($tenantIds[0])"
}

$tenantId = $tenantIds[0]

$opt = $payload.options
if ($null -eq $opt) { $opt = [pscustomobject]@{} }

$exoApp = ($env:EOA_EXO_APP_ID ?? '').Trim()
$exoOrg = ($env:EOA_EXO_ORGANIZATION ?? '').Trim()
$exoThumb = ($env:EOA_EXO_CERT_THUMBPRINT ?? '').Trim()
$exoSkip = ($env:EOA_EXO_SKIP_CONNECT ?? '').Trim() -eq 'true'

$gClient = ($env:EOA_GRAPH_CLIENT_ID ?? '').Trim()
$gSecret = ($env:EOA_GRAPH_CLIENT_SECRET ?? '').Trim()

if (-not $gClient -or -not $gSecret) {
    throw 'Set EOA_GRAPH_CLIENT_ID and EOA_GRAPH_CLIENT_SECRET for Graph app-only (same as Python Graph worker).'
}

if (-not $exoSkip) {
    if (-not $exoApp -or -not $exoOrg -or -not $exoThumb) {
        throw 'Set EOA_EXO_APP_ID, EOA_EXO_ORGANIZATION, EOA_EXO_CERT_THUMBPRINT for Exchange Online, or set EOA_EXO_SKIP_CONNECT=true for Graph-only.'
    }
}

$repoRoot = ($env:EOA_REPO_ROOT ?? '').Trim()
if (-not $repoRoot) {
    # This script lives at web/pwsh/ — repo root is parent of web/
    $webDir = Split-Path -Parent $PSScriptRoot
    $repoRoot = (Resolve-Path (Join-Path $webDir '..')).Path
}

$exportUtils = Join-Path $repoRoot 'Modules' 'ExportUtils.psm1'
$loggingMod = Join-Path $repoRoot 'Modules' 'Logging.psm1'
if (-not (Test-Path -LiteralPath $exportUtils)) {
    throw "ExportUtils module not found: $exportUtils (set EOA_REPO_ROOT)"
}

if (Test-Path -LiteralPath $loggingMod) {
    Import-Module -LiteralPath $loggingMod -Force -ErrorAction SilentlyContinue
}
Import-Module -LiteralPath $exportUtils -Force -ErrorAction Stop

try {
    $tokenResp = Get-ClientCredentialsToken -TenantId $tenantId -ClientId $gClient -ClientSecret $gSecret
    $at = $tokenResp.access_token
    if (-not $at) { throw 'Token response missing access_token' }
    $secGraph = ConvertTo-SecureString -String $at -AsPlainText -Force
    Connect-MgGraph -AccessToken $secGraph -NoWelcome -ErrorAction Stop | Out-Null
}
catch {
    throw "Graph Connect-MgGraph failed: $($_.Exception.Message)"
}

try {
    if (-not $exoSkip) {
        if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
            throw 'ExchangeOnlineManagement module is not installed. Install-Module ExchangeOnlineManagement -Scope CurrentUser'
        }
        Import-Module ExchangeOnlineManagement -ErrorAction Stop
        Connect-ExchangeOnline -AppId $exoApp -CertificateThumbprint $exoThumb -Organization $exoOrg -ShowBanner:$false -ErrorAction Stop
    }
}
catch {
    throw "Exchange Online Connect-ExchangeOnline failed: $($_.Exception.Message)"
}

$selectedUsers = @()
$su = $opt.PSObject.Properties['selected_users']
if ($su -and $su.Value) {
    $arr = @($su.Value)
    foreach ($u in $arr) {
        $s = "$u".Trim()
        if ($s -and $s -match '^[^<>"|:\\*?]{1,320}$') {
            $selectedUsers += $s
        }
    }
}

$statusFile = Join-Path $OutputDir 'exporter_status.log'
$daysBack = Get-OptionInt -Opt $opt -SnakeName 'days_back' -Default 10
$signDays = Get-OptionInt -Opt $opt -SnakeName 'sign_in_logs_days_back' -Default 7
$msgDays = Get-OptionInt -Opt $opt -SnakeName 'message_trace_days_back' -Default $daysBack

$reportParams = @{
    InvestigatorName                  = 'Web EXO runner'
    CompanyName                       = 'Organization'
    DaysBack                          = $daysBack
    OutputFolder                      = $OutputDir
    StatusFile                        = $statusFile
    NoParallel                        = $true
    MessageTraceDaysBack              = $msgDays
    SignInLogsDaysBack                = $signDays
    SelectedUsers                     = $selectedUsers
    IncludeMessageTrace               = (Get-OptionBool -Opt $opt -SnakeName 'include_message_trace' -Default $true)
    IncludeInboxRules                 = (Get-OptionBool -Opt $opt -SnakeName 'include_inbox_rules' -Default $true)
    IncludeTransportRules             = (Get-OptionBool -Opt $opt -SnakeName 'include_transport_rules' -Default $true)
    IncludeMailFlowConnectors         = (Get-OptionBool -Opt $opt -SnakeName 'include_mail_flow_connectors' -Default $true)
    IncludeMailboxForwarding          = (Get-OptionBool -Opt $opt -SnakeName 'include_mailbox_forwarding' -Default $true)
    IncludeAuditLogs                  = (Get-OptionBool -Opt $opt -SnakeName 'include_audit_logs' -Default $true)
    IncludeConditionalAccessPolicies  = (Get-OptionBool -Opt $opt -SnakeName 'include_conditional_access_policies' -Default $true)
    IncludeAppRegistrations           = (Get-OptionBool -Opt $opt -SnakeName 'include_app_registrations' -Default $true)
    IncludeSignInLogs                 = (Get-OptionBool -Opt $opt -SnakeName 'include_sign_in_logs' -Default $false)
    IncludeIntuneDevices              = (Get-OptionBool -Opt $opt -SnakeName 'include_intune_devices' -Default $false)
    IncludeMfaCoverage                = (Get-OptionBool -Opt $opt -SnakeName 'include_mfa_coverage' -Default $false)
    IncludeSharePointActivity         = (Get-OptionBool -Opt $opt -SnakeName 'include_share_point_activity' -Default $true)
    IncludeOneDriveActivity           = (Get-OptionBool -Opt $opt -SnakeName 'include_one_drive_activity' -Default $true)
    IncludeTeamsActivity              = (Get-OptionBool -Opt $opt -SnakeName 'include_teams_activity' -Default $true)
    IncludeSharePointSharing          = (Get-OptionBool -Opt $opt -SnakeName 'include_share_point_sharing' -Default $true)
    IncludeSecurityAlerts             = (Get-OptionBool -Opt $opt -SnakeName 'include_security_alerts' -Default $true)
    IncludeSecurityIncidents          = (Get-OptionBool -Opt $opt -SnakeName 'include_security_incidents' -Default $true)
    IncludeUnifiedAuditLogs           = (Get-OptionBool -Opt $opt -SnakeName 'include_unified_audit_logs' -Default $true)
}

$report = $null
try {
    $report = New-SecurityInvestigationReport @reportParams
}
finally {
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue } catch {}
    try { if (-not $exoSkip) { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } } catch {}
}

if (-not $report) {
    throw 'New-SecurityInvestigationReport returned no report object.'
}

# Desktop-compatible selection file (parity with WebBulkJobStub)
$reportSelections = [ordered]@{
    IncludeMessageTrace                 = $reportParams.IncludeMessageTrace
    IncludeInboxRules                   = $reportParams.IncludeInboxRules
    IncludeTransportRules               = $reportParams.IncludeTransportRules
    IncludeMailFlowConnectors           = $reportParams.IncludeMailFlowConnectors
    IncludeMailboxForwarding            = $reportParams.IncludeMailboxForwarding
    IncludeAuditLogs                    = $reportParams.IncludeAuditLogs
    IncludeConditionalAccessPolicies    = $reportParams.IncludeConditionalAccessPolicies
    IncludeAppRegistrations             = $reportParams.IncludeAppRegistrations
    IncludeSignInLogs                   = $reportParams.IncludeSignInLogs
    IncludeMfaCoverage                  = $reportParams.IncludeMfaCoverage
    IncludeSharePointActivity           = $reportParams.IncludeSharePointActivity
    IncludeOneDriveActivity             = $reportParams.IncludeOneDriveActivity
    IncludeTeamsActivity                = $reportParams.IncludeTeamsActivity
    IncludeSharePointSharing            = $reportParams.IncludeSharePointSharing
    IncludeSecurityAlerts               = $reportParams.IncludeSecurityAlerts
    IncludeSecurityIncidents            = $reportParams.IncludeSecurityIncidents
    IncludeAnonymousSharePointSharing   = (Get-OptionBool -Opt $opt -SnakeName 'include_anonymous_share_point_sharing' -Default $false)
    IncludeSharePointFileSharingLinks   = (Get-OptionBool -Opt $opt -SnakeName 'include_share_point_file_sharing_links' -Default $false)
    IncludeDLPViolations                = (Get-OptionBool -Opt $opt -SnakeName 'include_dlp_violations' -Default $false)
    IncludeIntuneDevices                = $reportParams.IncludeIntuneDevices
    IncludeUnifiedAuditLogs             = $reportParams.IncludeUnifiedAuditLogs
    IncludeSharePointOneDriveFileActions = (Get-OptionBool -Opt $opt -SnakeName 'include_share_point_one_drive_file_actions' -Default $false)
    SignInLogsDaysBack                  = $signDays
    MessageTraceDaysBack                = $msgDays
}
$rsPath = Join-Path $OutputDir 'ReportSelections.json'
($reportSelections | ConvertTo-Json -Depth 6) | Set-Content -LiteralPath $rsPath -Encoding UTF8

$summary = [ordered]@{
    workerVersion       = $script:WorkerVersion
    workerBackend       = 'pwsh-exo-linux'
    jobId               = $JobId
    tenantId            = $tenantId
    ok                  = $true
    outputFolder        = $OutputDir
    exchangeSkipped     = [bool]$exoSkip
    reportKeys          = @($report.Keys)
    message             = 'Ran New-SecurityInvestigationReport (ExportUtils). Artifacts under OutputFolder; see exporter_status.log.'
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
($summary | ConvertTo-Json -Depth 8) | Set-Content -LiteralPath $jsonPath -Encoding UTF8

"OK wrote $jsonPath and report under $OutputDir"
