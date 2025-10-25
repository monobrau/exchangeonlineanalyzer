# Test script for v8.1-beta integration
# Purpose: run headless checks for connectivity and data collection, and export CSVs

param(
    [int]$DaysBack = 7,
    [switch]$SkipConnect,
    [string]$OutputRoot
)

$ErrorActionPreference = 'Stop'

try {
    if (-not $OutputRoot -or [string]::IsNullOrWhiteSpace($OutputRoot)) {
        $OutputRoot = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'ExchangeOnlineAnalyzer\TestRuns'
    }
    if (-not (Test-Path $OutputRoot)) { New-Item -ItemType Directory -Path $OutputRoot -Force | Out-Null }
    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $outFolder = Join-Path $OutputRoot $timestamp
    New-Item -ItemType Directory -Path $outFolder -Force | Out-Null
    Write-Host ("Output: {0}" -f $outFolder) -ForegroundColor Cyan
} catch { Write-Error $_; exit 1 }

# Import local modules
Import-Module (Join-Path $PSScriptRoot '..\Modules\ExportUtils.psm1') -Force
Import-Module (Join-Path $PSScriptRoot '..\Modules\GraphOnline.psm1') -Force
Import-Module (Join-Path $PSScriptRoot '..\Modules\ExchangeOnline.psm1') -Force -ErrorAction SilentlyContinue

# Connect if needed
if (-not $SkipConnect) {
    try {
        Write-Host 'Connecting to Exchange Online...' -ForegroundColor Yellow
        Connect-ExchangeOnline -ErrorAction Stop
    } catch { Write-Warning ("EXO connect failed: {0}" -f $_.Exception.Message) }

    try {
        Write-Host 'Connecting to Microsoft Graph...' -ForegroundColor Yellow
        Connect-GraphService | Out-Null
    } catch { Write-Warning ("Graph connect failed: {0}" -f $_.Exception.Message) }
}

# Robust connection checks
$exoConnected = $false
try { Get-OrganizationConfig | Out-Null; $exoConnected = $true } catch { }
$mgConnected = $false
try { $ctx = Get-MgContext; if ($ctx -and $ctx.Account) { $mgConnected = $true } } catch { }
Write-Host ("EXO: {0} | Graph: {1}" -f ($exoConnected?'Connected':'Not Connected'), ($mgConnected?'Connected':'Not Connected')) -ForegroundColor Green

# Collections
$summary = [ordered]@{}

if ($exoConnected) {
    try {
        Write-Host 'Collecting message trace (10 days)...' -ForegroundColor Yellow
        $mt = Get-ExchangeMessageTrace -DaysBack 10
        $summary['MessageTrace'] = $mt.Count
        if ($mt.Count -gt 0) { $mt | Export-Csv -Path (Join-Path $outFolder 'MessageTrace.csv') -NoTypeInformation -Encoding UTF8 }
    } catch { Write-Warning ("Message trace error: {0}" -f $_.Exception.Message) }

    try {
        Write-Host 'Exporting inbox rules...' -ForegroundColor Yellow
        $rules = Get-ExchangeInboxRules
        $summary['InboxRules'] = $rules.Count
        if ($rules.Count -gt 0) { $rules | Export-Csv -Path (Join-Path $outFolder 'InboxRules.csv') -NoTypeInformation -Encoding UTF8 }
    } catch { Write-Warning ("Inbox rules error: {0}" -f $_.Exception.Message) }
}

if ($mgConnected) {
    try {
        Write-Host 'Collecting audit logs...' -ForegroundColor Yellow
        $aud = Get-GraphAuditLogs -DaysBack $DaysBack
        $summary['AuditLogs'] = $aud.Count
        if ($aud.Count -gt 0) { $aud | Export-Csv -Path (Join-Path $outFolder 'GraphAuditLogs.csv') -NoTypeInformation -Encoding UTF8 }
    } catch { Write-Warning ("Audit logs error: {0}" -f $_.Exception.Message) }

    try {
        Write-Host 'Evaluating MFA coverage...' -ForegroundColor Yellow
        $mfa = Get-MfaCoverageReport
        $summary['MFAUsers'] = if ($mfa -and $mfa.Users) { $mfa.Users.Count } else { 0 }
        if ($mfa -and $mfa.Users -and $mfa.Users.Count -gt 0) { $mfa.Users | Export-Csv -Path (Join-Path $outFolder 'MFAStatus.csv') -NoTypeInformation -Encoding UTF8 }
        $summary['SecurityDefaults'] = if ($mfa) { $mfa.SecurityDefaultsEnabled } else { $false }
        $summary['CARequiresMfa(Tenant)'] = if ($mfa) { $mfa.CAPoliciesRequireMfa } else { $false }
    } catch { Write-Warning ("MFA coverage error: {0}" -f $_.Exception.Message) }

    try {
        Write-Host 'Collecting user security groups/roles...' -ForegroundColor Yellow
        $usg = Get-UserSecurityGroupsReport
        $summary['UserGroups'] = $usg.Count
        if ($usg.Count -gt 0) { $usg | Export-Csv -Path (Join-Path $outFolder 'UserSecurityGroups.csv') -NoTypeInformation -Encoding UTF8 }
    } catch { Write-Warning ("User groups error: {0}" -f $_.Exception.Message) }
}

Write-Host '--- Summary ---' -ForegroundColor Cyan
$summary.GetEnumerator() | Sort-Object Name | ForEach-Object { Write-Host ("{0}: {1}" -f $_.Key, $_.Value) }
Write-Host ("Output folder: {0}" -f $outFolder) -ForegroundColor Cyan

Write-Host 'Done.' -ForegroundColor Green


