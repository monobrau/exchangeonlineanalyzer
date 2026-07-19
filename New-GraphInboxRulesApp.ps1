<#
.SYNOPSIS
    Creates an Entra app registration for Graph security investigation reports (app-only) and grants admin consent.
.DESCRIPTION
    Creates app "River Run Security Investigator" with permissions for inbox rules, audit logs, sign-in logs,
    Conditional Access, app registrations, organization read (tenant display name via GET /organization), reports,
    SharePoint, security alerts, and MFA. Saves to WCM when -SaveToWCM.
    Requires: Application.ReadWrite.All, AppRoleAssignment.ReadWrite.All (admin).
    If an app with this name already exists, prompts to replace (delete and recreate) or cancel.
.PARAMETER SaveToWCM
    Save TenantId, ClientId, ClientSecret to Windows Credential Manager for this tenant.
    Requires: Install-Module CredentialManager
.PARAMETER TenantId
    Optional. If set, passed to Connect-MgGraph for tenant-scoped sign-in.
.PARAMETER UseDeviceCode
    Use device code sign-in instead of the default interactive/WAM flow.
.EXAMPLE
    .\New-GraphInboxRulesApp.ps1 -SaveToWCM
.EXAMPLE
    .\New-GraphInboxRulesApp.ps1 -SaveToWCM -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
#>

#Requires -Version 5.1

param(
    [switch]$SaveToWCM = $false,
    [string]$TenantId = $null,
    [switch]$UseDeviceCode = $false
)

$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$resultPath = Join-Path $env:TEMP 'EOA-GraphAppCreate-result.json'
$script:graphAppCreateTenantId = $null
$script:graphAppCreateTenantDisplayName = $null
$script:graphAppCreateClientId = $null
$script:graphAppCreateScriptError = $null
$script:graphAppCreateWcmSaved = $false
$script:graphAppCreateWcmError = $null

function Write-GraphAppCreateResultFile {
    try {
        @{
            TenantId          = $script:graphAppCreateTenantId
            TenantDisplayName = $script:graphAppCreateTenantDisplayName
            ClientId          = $script:graphAppCreateClientId
            WcmSaved          = [bool]$script:graphAppCreateWcmSaved
            WcmError          = $script:graphAppCreateWcmError
            ScriptError       = $script:graphAppCreateScriptError
            Timestamp         = (Get-Date).ToString('o')
        } | ConvertTo-Json | Set-Content -Path $resultPath -Encoding UTF8 -Force
    } catch {
        Write-Warning "Could not write result file $resultPath : $($_.Exception.Message)"
    }
}

function Resolve-GraphAppCreateModuleVersion {
    <#
    .SYNOPSIS
        Highest version where required Graph submodules are all installed (avoids Auth 2.37 + Applications 2.36 mismatch).
    #>
    $required = @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Applications')
    $candidates = @(Get-Module -ListAvailable -Name Microsoft.Graph.Authentication | Sort-Object Version -Descending)
    if ($candidates.Count -eq 0) {
        throw "Microsoft Graph PowerShell SDK not installed. In pwsh run: Install-Module Microsoft.Graph -Scope CurrentUser -Force"
    }
    foreach ($c in $candidates) {
        $ver = $c.Version
        $ok = $true
        foreach ($name in $required) {
            if (-not (Get-Module -ListAvailable -Name $name | Where-Object { $_.Version -eq $ver })) {
                $ok = $false
                break
            }
        }
        if ($ok) { return $ver }
    }
    $latestAuth = $candidates[0].Version
    throw @"
Microsoft Graph submodules are out of sync (e.g. Authentication $latestAuth but Applications not installed at that version).
In pwsh run: Update-Module Microsoft.Graph* -Scope CurrentUser -Force
"@
}

function Import-GraphAppCreateModuleStack {
    <#
    .SYNOPSIS
        Loads only missing Graph submodules needed for app registration (same version). Skips Import-Module when already loaded.
    #>
    $previousAutoLoad = $PSModuleAutoloadingPreference
    $PSModuleAutoloadingPreference = 'None'

    try {
        $ver = Resolve-GraphAppCreateModuleVersion
        $stack = [System.Collections.Generic.List[string]]::new()
        [void]$stack.Add('Microsoft.Graph.Authentication')
        [void]$stack.Add('Microsoft.Graph.Applications')
        if (Get-Module -ListAvailable -Name Microsoft.Graph.Identity.DirectoryManagement | Where-Object { $_.Version -eq $ver }) {
            [void]$stack.Add('Microsoft.Graph.Identity.DirectoryManagement')
        }

        $loaded = @(Get-Module -Name $stack -ErrorAction SilentlyContinue)
        $wrongVersion = @($loaded | Where-Object { $_.Version -ne $ver })
        if ($wrongVersion.Count -gt 0) {
            Write-Host "Reloading Graph modules (version mismatch)..." -ForegroundColor Gray
            foreach ($m in $loaded) {
                Remove-Module -Name $m.Name -Force -ErrorAction SilentlyContinue
            }
        }

        $toImport = [System.Collections.Generic.List[string]]::new()
        foreach ($name in $stack) {
            $m = Get-Module -Name $name -ErrorAction SilentlyContinue
            if ($m -and $m.Version -eq $ver) { continue }
            [void]$toImport.Add($name)
        }

        if ($toImport.Count -eq 0) {
            Write-Host "Graph modules already loaded (version $ver)." -ForegroundColor Gray
            return
        }

        Write-Host "Loading Graph modules for app create (version $ver)..." -ForegroundColor Gray
        foreach ($name in $toImport) {
            Write-Host "  $name" -ForegroundColor DarkGray
            Import-Module $name -RequiredVersion $ver -ErrorAction Stop
        }
    }
    finally {
        $PSModuleAutoloadingPreference = $previousAutoLoad
    }
}

$script:graphAppCreateTranscriptPath = Join-Path $env:TEMP 'EOA-GraphAppCreate-last.log'
try {
    Start-Transcript -Path $script:graphAppCreateTranscriptPath -Force -ErrorAction SilentlyContinue | Out-Null
} catch { }

try {
Write-Host "`n=== Create Graph Inbox Rules App ===" -ForegroundColor Cyan
Write-Host "Complete browser sign-in and answer the prompts below. Do not close this window until you see 'App Created' or an error." -ForegroundColor Yellow
Import-GraphAppCreateModuleStack

$guidPattern = '^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$'
$resolvedTenantId = $null
if ($TenantId -and $TenantId.Trim() -notmatch $guidPattern) {
    throw "Invalid -TenantId: expected a directory GUID."
}
if ($TenantId -and $TenantId.Trim() -match $guidPattern) {
    $resolvedTenantId = $TenantId.Trim()
    Write-Host "Using -TenantId: $resolvedTenantId" -ForegroundColor Gray
}

$scopes = @('Application.ReadWrite.All', 'AppRoleAssignment.ReadWrite.All')

# Bypass WAM (mandatory since Graph SDK 2.34+) so MSAL uses the system browser with an
# account picker instead of silently reusing the last-used Windows broker account.
$env:MSAL_FORCE_WAM = '0'

# Clear inherited broker env vars from the parent process
foreach ($k in @('AZURE_IDENTITY_DISABLE_BROKER', 'MSAL_DISABLE_BROKER', 'MSAL_EXPERIMENTAL_DISABLE_BROKER')) {
    if (Test-Path "Env:\$k") { Remove-Item "Env:\$k" -ErrorAction SilentlyContinue }
}

# Isolated MSAL cache for this run (do not call Disconnect-MgGraph / GraphSession here - loading those
# assemblies before Connect-MgGraph causes "assembly already loaded" version conflicts in pwsh).
$authCacheDir = Join-Path $env:TEMP "EOA_GraphAppCreate_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
try {
    if (Test-Path $authCacheDir) { Remove-Item -Path $authCacheDir -Recurse -Force -ErrorAction SilentlyContinue }
    New-Item -ItemType Directory -Path $authCacheDir -Force -ErrorAction Stop | Out-Null
    $env:MSAL_CACHE_DIR = $authCacheDir
    $env:IDENTITY_SERVICE_CACHE_DIR = $authCacheDir
} catch {}
try {
    $graphCache = Join-Path $env:LOCALAPPDATA "Microsoft\Graph"
    if (Test-Path $graphCache) { Get-ChildItem -Path $graphCache -Recurse -File -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue }
} catch {}
try {
    $defaultMsalCache = Join-Path $env:LOCALAPPDATA ".IdentityService"
    if (Test-Path $defaultMsalCache) {
        Get-ChildItem -Path $defaultMsalCache -Recurse -File -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
    }
} catch {}

function Connect-GraphWithScopes {
    param([string[]]$Scopes, [string]$TenantId, [switch]$UseDeviceCode)

    $connectParams = @{
        Scopes       = $Scopes
        ContextScope = 'Process'
        NoWelcome    = $true
        ErrorAction  = 'Stop'
    }
    if ($UseDeviceCode) { $connectParams.UseDeviceCode = $true }
    if ($TenantId)      { $connectParams.TenantId = $TenantId }

    if ($UseDeviceCode) {
        Write-Host "Connecting (device code) with $($Scopes -join ', ')..." -ForegroundColor Yellow
    } else {
        Write-Host "Connecting (browser) with $($Scopes -join ', ')..." -ForegroundColor Yellow
    }
    Connect-MgGraph @connectParams
}

# --- Connect and verify tenant -----------------------------------------------
$maxAttempts = 3
for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
    try {
        if ($attempt -eq 1) {
            Connect-GraphWithScopes -Scopes $scopes -TenantId $resolvedTenantId -UseDeviceCode:$UseDeviceCode
        } else {
            Write-Host "`nRetrying with device code (choose the correct account in the browser)..." -ForegroundColor Cyan
            Connect-GraphWithScopes -Scopes $scopes -TenantId $resolvedTenantId -UseDeviceCode
        }
    } catch {
        throw "Graph connect failed: $($_.Exception.Message)"
    }

    $tenantId = (Get-MgContext).TenantId
    $script:graphAppCreateTenantId = $tenantId
    $tenantDisplayName = $tenantId
    $script:graphAppCreateTenantDisplayName = $tenantDisplayName
    try {
        $org = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/organization' -ErrorAction Stop
        if ($org.value -and $org.value[0].displayName) {
            $tenantDisplayName = $org.value[0].displayName
            $script:graphAppCreateTenantDisplayName = $tenantDisplayName
        }
    } catch {}
    Write-Host "Connected. Tenant: $tenantDisplayName ($tenantId)" -ForegroundColor Green

    Write-Host "`n>>> Is this the correct tenant? Type Y and press Enter (n = sign in to a different tenant): " -ForegroundColor Yellow -NoNewline
    $confirm = Read-Host
    if ($confirm -eq '' -or $confirm -ieq 'y') { break }

    Write-Host "Wrong tenant - disconnecting..." -ForegroundColor Yellow
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue } catch { }

    if ($attempt -eq $maxAttempts) {
        throw "Unable to connect to the desired tenant after $maxAttempts attempts."
    }
}

# Microsoft Graph resource app ID
$graphAppId = '00000003-0000-0000-c000-000000000000'

# App role IDs for security investigation report (inbox rules, audit, sign-in, CA, apps, reports, sites, alerts, MFA)
# Plus Application.ReadWrite.All, AppRoleAssignment.ReadWrite.All for Entra Secret Rotate (secret rotation + Add ATR)
# Mail.Read required for Get-MgUserMailFolderMessageRule (inbox rules); MailboxSettings.Read is for auto-reply etc.
# Organization.Read.All: explicit app-only read for GET /organization (tenant displayName); see Microsoft Graph permissions reference.
# SecurityAlert.Read.All / SecurityIncident.Read.All: Defender security alerts & incidents collectors (app-only).
$appRoleIds = @(
    @{ id = '810c84a8-4a9e-49e6-bf7d-12d183f40d01'; name = 'Mail.Read' }
    @{ id = '40f97065-369a-49f4-947c-6a255697ae91'; name = 'MailboxSettings.Read' }
    @{ id = 'df021288-bdef-4463-88db-98f22de89214'; name = 'User.Read.All' }
    @{ id = 'b0afded3-3588-46d8-8b3d-9842eff778da'; name = 'AuditLog.Read.All' }
    @{ id = '7ab1d382-f21e-4acd-a863-ba3e13f7da61'; name = 'Directory.Read.All' }
    @{ id = '498476ce-e0fe-48b0-b801-37ba7e2685c6'; name = 'Organization.Read.All' }
    @{ id = '246dd0d5-5bd0-4def-940b-0421030a5b68'; name = 'Policy.Read.All' }
    # Application permission required for CA Manager / app-only Conditional Access policy changes
    @{ id = '01c0a623-fc9b-48e9-b794-0756f8e8f067'; name = 'Policy.ReadWrite.ConditionalAccess' }
    @{ id = '9a5d68dd-52b0-4cc2-bd40-abcf44ac3a30'; name = 'Application.Read.All' }
    @{ id = '1bfefb4e-e0b5-418b-a88f-73c46d2cc8e9'; name = 'Application.ReadWrite.All' }
    @{ id = '06b708a9-e830-4db3-a914-8e69da51d44f'; name = 'AppRoleAssignment.ReadWrite.All' }
    @{ id = '230c1aed-a721-4c5d-9cb4-a90514e508ef'; name = 'Reports.Read.All' }
    @{ id = '332a536c-c7ef-4017-ab91-336970924f0d'; name = 'Sites.Read.All' }
    @{ id = 'bf394140-e372-4bf9-a898-299cfc7564e5'; name = 'SecurityEvents.Read.All' }
    @{ id = '472e4a4d-bb4a-4026-98d1-0b0d74cb74a5'; name = 'SecurityAlert.Read.All' }
    @{ id = '45cc0394-e837-488b-a098-1918f48d186c'; name = 'SecurityIncident.Read.All' }
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

$script:graphAppCreateClientId = $app.AppId
if ($SaveToWCM) {
    if (-not (Get-Module -ListAvailable -Name CredentialManager)) {
        Write-Host "Installing CredentialManager module (one-time)..." -ForegroundColor Gray
        try {
            Install-Module CredentialManager -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        } catch {
            Write-Warning "Could not install CredentialManager: $($_.Exception.Message). WCM save will use CredWrite/cmdkey fallback."
        }
    }
    Write-Host "`nSaving to Windows Credential Manager..." -ForegroundColor Yellow
    try {
        $tenantDisplayName = $null
        try {
            $org = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/organization' -ErrorAction Stop
            if ($org.value -and $org.value[0].displayName) {
                $tenantDisplayName = $org.value[0].displayName
                $script:graphAppCreateTenantDisplayName = $tenantDisplayName
            }
        } catch {}
        Import-Module (Join-Path $projectRoot 'Modules\GraphAppCredential.psm1') -Force -ErrorAction Stop
        Save-GraphAppCredentialToWCM -TenantId $tenantId -ClientId $app.AppId -ClientSecret $cred.secretText -TenantDisplayName $tenantDisplayName
        if (Get-GraphAppCredentialFromWCM -TenantId $tenantId) {
            $script:graphAppCreateWcmSaved = $true
            Write-Host "  Saved and verified in Credential Manager." -ForegroundColor Green
        }
        else {
            throw 'Save completed but credential could not be read back from Credential Manager.'
        }
    } catch {
        $script:graphAppCreateWcmError = $_.Exception.Message
        Write-Warning "Could not save to WCM: $($script:graphAppCreateWcmError)"
    }
}

Write-GraphAppCreateResultFile

if ($SaveToWCM -and -not $script:graphAppCreateWcmSaved) {
    Write-Host "`nERROR: App exists in Entra but was NOT stored in Windows Credential Manager." -ForegroundColor Red
    Write-Host "The App reg tenant dropdown will not list this tenant until WCM save succeeds." -ForegroundColor Red
    if ($script:graphAppCreateWcmError) { Write-Host "Detail: $($script:graphAppCreateWcmError)" -ForegroundColor Red }
    exit 2
}

Write-Host "`n=== App Created ===" -ForegroundColor Cyan
Write-Host "TenantId:  $tenantId"
Write-Host "ClientId:  $($app.AppId)"
Write-Host "Secret:    $($cred.secretText)"
Write-Host "`nSave the secret now - it is shown only once." -ForegroundColor Yellow
Write-Host ""

}
catch {
    $script:graphAppCreateScriptError = $_.Exception.Message
    Write-GraphAppCreateResultFile
    Write-Error $_
    exit 1
}
finally {
    try { Stop-Transcript -ErrorAction SilentlyContinue | Out-Null } catch { }
}
