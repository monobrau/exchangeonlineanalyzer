<#
.SYNOPSIS
    Removes "River Run Security Investigator" app registration(s) from the specified tenant(s).
.DESCRIPTION
    Finds and deletes app registrations with display name "River Run Security Investigator"
    (created by New-GraphInboxRulesApp.ps1), their service principals, and the WCM credential.
    Requires: Application.ReadWrite.All, AppRoleAssignment.ReadWrite.All (admin).
.PARAMETER TenantId
    Specific tenant ID to remove the app from. When provided, Connect-MgGraph targets this tenant.
    When omitted, connects to default tenant (user's home tenant).
.PARAMETER Force
    Skip confirmation prompt.
.EXAMPLE
    .\Remove-GraphInboxRulesApp.ps1
.EXAMPLE
    .\Remove-GraphInboxRulesApp.ps1 -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -Force
#>

#Requires -Version 5.1

param(
    [Parameter(Mandatory=$false)]
    [string]$TenantId = $null,
    [switch]$Force = $false
)

$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$displayName = 'River Run Security Investigator'

$scopes = @('Application.ReadWrite.All', 'AppRoleAssignment.ReadWrite.All')
Write-Host "`n=== Remove Graph Inbox Rules App ===" -ForegroundColor Cyan
Write-Host "Connecting with Application.ReadWrite.All, AppRoleAssignment.ReadWrite.All..." -ForegroundColor Yellow

try {
    if ($TenantId) {
        Connect-MgGraph -Scopes $scopes -NoWelcome -TenantId $TenantId -ErrorAction Stop
        Write-Host "Targeting tenant: $TenantId" -ForegroundColor Gray
    } else {
        Connect-MgGraph -Scopes $scopes -NoWelcome -ErrorAction Stop
    }
} catch {
    Write-Error "Graph connect failed: $_"
}

$tenantId = (Get-MgContext).TenantId
Write-Host "Connected. Tenant: $tenantId" -ForegroundColor Green

$existingApps = @()
try {
    $found = Get-MgApplication -Filter "displayName eq '$displayName'" -ErrorAction SilentlyContinue
    if ($found) {
        $existingApps = @($found)
    }
} catch {}

if ($existingApps.Count -eq 0) {
    Write-Host "`nNo apps named '$displayName' found in this tenant." -ForegroundColor Gray
    exit 0
}

Write-Host "`nFound $($existingApps.Count) app(s) named '$displayName':" -ForegroundColor Yellow
foreach ($a in $existingApps) {
    Write-Host "  - AppId: $($a.AppId)" -ForegroundColor Gray
}

if (-not $Force) {
    Write-Host "`nRemove all of them? (y/n): " -ForegroundColor Yellow -NoNewline
    $reply = Read-Host
    if ($reply -ne 'y' -and $reply -ne 'Y') {
        Write-Host "Cancelled." -ForegroundColor Gray
        exit 0
    }
}

$removed = 0
foreach ($a in $existingApps) {
    Write-Host "Removing app $($a.AppId)..." -ForegroundColor Yellow
    try {
        $sps = Get-MgServicePrincipal -Filter "appId eq '$($a.AppId)'" -ErrorAction SilentlyContinue
        if ($sps) {
            foreach ($sp in @($sps)) {
                Remove-MgServicePrincipal -ServicePrincipalId $sp.Id -ErrorAction SilentlyContinue
            }
        }
        Remove-MgApplication -ApplicationId $a.Id -ErrorAction Stop
        Write-Host "  Removed." -ForegroundColor Green
        $removed++
    } catch {
        Write-Warning "  Failed: $($_.Exception.Message)"
    }
}

# Remove WCM credential for this tenant
try {
    Import-Module (Join-Path $projectRoot 'Modules\GraphAppCredential.psm1') -Force -ErrorAction SilentlyContinue
    Remove-GraphAppCredentialFromWCM -TenantId $tenantId -ErrorAction SilentlyContinue
    Write-Host "Removed stored credential from Windows Credential Manager." -ForegroundColor Green
} catch {}

Write-Host "`nDone. Removed $removed app(s)." -ForegroundColor Cyan
