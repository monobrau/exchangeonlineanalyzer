<#
.SYNOPSIS
    Creates an Entra app registration for Graph security investigation reports (app-only) and grants admin consent.
.DESCRIPTION
    Creates app "River Run Security Investigator" with permissions for inbox rules, audit logs, sign-in logs,
    Conditional Access, app registrations, reports, SharePoint, security alerts, and MFA. Saves to WCM when -SaveToWCM.
    Requires: Application.ReadWrite.All, AppRoleAssignment.ReadWrite.All (admin).
    If an app with this name already exists, prompts to replace (delete and recreate) or cancel.
.PARAMETER SaveToWCM
    Save TenantId, ClientId, ClientSecret to Windows Credential Manager for this tenant.
    Requires: Install-Module CredentialManager
.EXAMPLE
    .\New-GraphInboxRulesApp.ps1 -SaveToWCM
#>

#Requires -Version 5.1

param([switch]$SaveToWCM = $false)

$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path

$scopes = @('Application.ReadWrite.All', 'AppRoleAssignment.ReadWrite.All')
Write-Host "`n=== Create Graph Inbox Rules App ===" -ForegroundColor Cyan
Write-Host "Connecting with Application.ReadWrite.All, AppRoleAssignment.ReadWrite.All..." -ForegroundColor Yellow

try {
    Connect-MgGraph -Scopes $scopes -NoWelcome -ErrorAction Stop
} catch {
    Write-Error "Graph connect failed: $_"
}

$tenantId = (Get-MgContext).TenantId
Write-Host "Connected. Tenant: $tenantId" -ForegroundColor Green

# Microsoft Graph resource app ID
$graphAppId = '00000003-0000-0000-c000-000000000000'

# App role IDs for security investigation report (inbox rules, audit, sign-in, CA, apps, reports, sites, alerts, MFA)
# Mail.Read required for Get-MgUserMailFolderMessageRule (inbox rules); MailboxSettings.Read is for auto-reply etc.
$appRoleIds = @(
    @{ id = '810c84a8-4a9e-49e6-bf7d-12d183f40d01'; name = 'Mail.Read' }
    @{ id = '40f97065-369a-49f4-947c-6a255697ae91'; name = 'MailboxSettings.Read' }
    @{ id = 'df021288-bdef-4463-88db-98f22de89214'; name = 'User.Read.All' }
    @{ id = 'b0afded3-3588-46d8-8b3d-9842eff778da'; name = 'AuditLog.Read.All' }
    @{ id = '7ab1d382-f21e-4acd-a863-ba3e13f7da61'; name = 'Directory.Read.All' }
    @{ id = '246dd0d5-5bd0-4def-940b-0421030a5b68'; name = 'Policy.Read.All' }
    @{ id = '9a5d68dd-52b0-4cc2-bd40-abcf44ac3a30'; name = 'Application.Read.All' }
    @{ id = '230c1aed-a721-4c5d-9cb4-a90514e508ef'; name = 'Reports.Read.All' }
    @{ id = '332a536c-c7ef-4017-ab91-336970924f0d'; name = 'Sites.Read.All' }
    @{ id = 'bf394140-e372-4bf9-a898-299cfc7564e5'; name = 'SecurityEvents.Read.All' }
    @{ id = '38d9df27-64da-44fd-b7c5-a6fbac20248f'; name = 'UserAuthenticationMethod.Read.All' }
)
$requiredResourceAccess = @{
    resourceAccess = $appRoleIds | ForEach-Object { @{ id = $_.id; type = 'Role' } }
    resourceAppId = $graphAppId
}

$displayName = 'River Run Security Investigator'

# Check for existing app(s) with same display name
$existingApps = @()
try {
    $filter = "displayName eq 'River Run Security Investigator'"
    $found = Get-MgApplication -Filter $filter -ErrorAction SilentlyContinue
    if ($found) {
        $existingApps = @($found)
    }
} catch {}

if ($existingApps.Count -gt 0) {
    Write-Host "`nFound $($existingApps.Count) existing app(s) named '$displayName':" -ForegroundColor Yellow
    foreach ($a in $existingApps) {
        Write-Host "  - AppId: $($a.AppId)" -ForegroundColor Gray
    }
    Write-Host "`nReplace (delete existing and create new)? (y/n): " -ForegroundColor Yellow -NoNewline
    $reply = Read-Host
    if ($reply -ne 'y' -and $reply -ne 'Y') {
        Write-Host "Cancelled." -ForegroundColor Gray
        exit 0
    }
    foreach ($a in $existingApps) {
        Write-Host "Deleting app $($a.AppId)..." -ForegroundColor Yellow
        try {
            $sps = Get-MgServicePrincipal -Filter "appId eq '$($a.AppId)'" -ErrorAction SilentlyContinue
            if ($sps) {
                foreach ($sp in @($sps)) {
                    Remove-MgServicePrincipal -ServicePrincipalId $sp.Id -ErrorAction SilentlyContinue
                }
            }
            Remove-MgApplication -ApplicationId $a.Id -ErrorAction Stop
            Write-Host "  Deleted." -ForegroundColor Green
        } catch {
            Write-Warning "  Failed to delete: $($_.Exception.Message)"
        }
    }
    # Remove WCM credential for this tenant so we don't keep stale ClientId/secret
    try {
        Import-Module (Join-Path $projectRoot 'Modules\GraphAppCredential.psm1') -Force -ErrorAction SilentlyContinue
        Remove-GraphAppCredentialFromWCM -TenantId $tenantId -ErrorAction SilentlyContinue
    } catch {}
}

Write-Host "`nCreating app: $displayName" -ForegroundColor Yellow
$app = New-MgApplication -DisplayName $displayName -RequiredResourceAccess $requiredResourceAccess
Write-Host "  AppId: $($app.AppId)" -ForegroundColor Gray

Write-Host "`nCreating service principal..." -ForegroundColor Yellow
$graphSp = Get-MgServicePrincipal -Filter "appId eq '$graphAppId'"
$sp = New-MgServicePrincipal -AppId $app.AppId
Write-Host "  ServicePrincipalId: $($sp.Id)" -ForegroundColor Gray

Write-Host "`nGranting admin consent (app role assignments)..." -ForegroundColor Yellow
foreach ($role in $appRoleIds) {
    try {
        New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id -PrincipalId $sp.Id -AppRoleId $role.id -ResourceId $graphSp.Id -ErrorAction Stop | Out-Null
        Write-Host "  $($role.name) - granted" -ForegroundColor Green
    } catch { Write-Warning "  $($role.name) - $($_.Exception.Message)" }
}

Write-Host "`nCreating client secret..." -ForegroundColor Yellow
$cred = Add-MgApplicationPassword -ApplicationId $app.Id
Write-Host "  Secret created (expires: $($cred.endDateTime))" -ForegroundColor Gray

if ($SaveToWCM) {
    Write-Host "`nSaving to Windows Credential Manager..." -ForegroundColor Yellow
    try {
        Import-Module (Join-Path $projectRoot 'Modules\GraphAppCredential.psm1') -Force -ErrorAction Stop
        Save-GraphAppCredentialToWCM -TenantId $tenantId -ClientId $app.AppId -ClientSecret $cred.secretText
        Write-Host "  Saved. Reports will use these credentials when pulling for this tenant." -ForegroundColor Green
    } catch {
        Write-Warning "Could not save to WCM: $($_.Exception.Message). Install-Module CredentialManager -Scope CurrentUser"
    }
}

Write-Host "`n=== App Created ===" -ForegroundColor Cyan
Write-Host "TenantId:  $tenantId"
Write-Host "ClientId:  $($app.AppId)"
Write-Host "Secret:    $($cred.secretText)"
Write-Host "`nSave the secret now - it is shown only once." -ForegroundColor Yellow
Write-Host ""
