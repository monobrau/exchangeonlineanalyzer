<#
.SYNOPSIS
    Creates an Entra app registration for Graph security investigation reports (app-only) and grants admin consent.
.DESCRIPTION
    Creates app "River Run Security Investigator" with permissions for inbox rules, audit logs, sign-in logs,
    Conditional Access, app registrations, organization read (tenant display name via GET /organization), reports,
    SharePoint, security alerts, MFA, and web-runner containment writes (User.RevokeSessions.All,
    User.EnableDisableAccount.All, UserAuthenticationMethod.ReadWrite.All, Device.ReadWrite.All,
    Application.ReadWrite.All, User-PasswordProfile.ReadWrite.All, DelegatedPermissionGrant.ReadWrite.All,
    RoleManagement.ReadWrite.Directory, GroupMember.ReadWrite.All, DeviceManagementManagedDevices.ReadWrite.All,
    DeviceManagementManagedDevices.PrivilegedOperations.All). Saves to WCM when -SaveToWCM.
    Use -UpdateExisting to add missing application permissions and admin-consent them on the
    current River Run Security Investigator app (or the WCM ClientId for this tenant) without
    rotating the client secret. Create mode still prompts: Update scopes / Replace / Cancel.
    Requires: Application.ReadWrite.All, AppRoleAssignment.ReadWrite.All, Organization.Read.All (admin).
.PARAMETER SaveToWCM
    Save TenantId, ClientId, ClientSecret to Windows Credential Manager for this tenant.
    Requires: Install-Module CredentialManager
.PARAMETER TenantId
    Optional. If set, passed to Connect-MgGraph for tenant-scoped sign-in.
.PARAMETER UseDeviceCode
    Use device code sign-in instead of the default interactive/WAM flow.
.PARAMETER UpdateExisting
    Patch requiredResourceAccess and grant missing app-role assignments. Does not create a
    new secret or rewrite WCM. After it finishes, run Graph Auth again for a new token.
.EXAMPLE
    .\New-GraphInboxRulesApp.ps1 -SaveToWCM
.EXAMPLE
    .\New-GraphInboxRulesApp.ps1 -SaveToWCM -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
#>

#Requires -Version 5.1

param(
    [switch]$SaveToWCM = $false,
    [string]$TenantId = $null,
    [switch]$UseDeviceCode = $false,
    [switch]$UpdateExisting = $false
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
$script:graphAppCreateUpdatedExisting = $false
$script:graphAppCreateRolesGranted = 0
$script:graphAppCreateRolesAlready = 0
$script:graphAppCreateRolesFailed = 0

function Write-GraphAppCreateResultFile {
    try {
        @{
            TenantId          = $script:graphAppCreateTenantId
            TenantDisplayName = $script:graphAppCreateTenantDisplayName
            ClientId          = $script:graphAppCreateClientId
            WcmSaved          = [bool]$script:graphAppCreateWcmSaved
            WcmError          = $script:graphAppCreateWcmError
            UpdatedExisting   = [bool]$script:graphAppCreateUpdatedExisting
            RolesGranted      = [int]$script:graphAppCreateRolesGranted
            RolesAlready      = [int]$script:graphAppCreateRolesAlready
            RolesFailed       = [int]$script:graphAppCreateRolesFailed
            ScriptError       = $script:graphAppCreateScriptError
            Timestamp         = (Get-Date).ToString('o')
        } | ConvertTo-Json | Set-Content -Path $resultPath -Encoding UTF8 -Force
    } catch {
        Write-Warning "Could not write result file $resultPath : $($_.Exception.Message)"
    }
}

function Ensure-TransportDataPlatformServicePrincipal {
    # Microsoft-managed app required before Graph messageTraces will authorize (see Graph message trace onboarding).
    $tdpAppId = '8bd644d1-64a1-4d4b-ae52-2e0cbf64e373'
    try {
        $existing = @(Get-MgServicePrincipal -Filter "appId eq '$tdpAppId'" -ErrorAction SilentlyContinue)
        if ($existing.Count -gt 0) {
            Write-Host "Transport Data Platform service principal already present (Graph message trace)." -ForegroundColor DarkGray
            return
        }
    } catch { }
    try {
        New-MgServicePrincipal -AppId $tdpAppId -ErrorAction Stop | Out-Null
        Write-Host "Provisioned Transport Data Platform service principal (required for Graph messageTraces). Propagation can take a few hours." -ForegroundColor Green
    } catch {
        $msg = $_.Exception.Message
        if ($msg -match 'already|ObjectConflict|duplicate|being provisioned') {
            Write-Host "Transport Data Platform service principal already present." -ForegroundColor DarkGray
        } else {
            Write-Warning "Could not provision Transport Data Platform SP ($tdpAppId): $msg"
        }
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
if ($UpdateExisting) {
    Write-Host "`n=== Update Graph App scopes ===" -ForegroundColor Cyan
    Write-Host "Complete browser sign-in. This adds missing application permissions on the existing app and does not rotate the secret." -ForegroundColor Yellow
} else {
    Write-Host "`n=== Create Graph Inbox Rules App ===" -ForegroundColor Cyan
    Write-Host "Complete browser sign-in and answer the prompts below. Do not close this window until you see 'App Created' or an error." -ForegroundColor Yellow
}
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

$scopes = @('Application.ReadWrite.All', 'AppRoleAssignment.ReadWrite.All', 'Organization.Read.All')

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
    $tenantDisplayName = $null
    try {
        $org = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/organization' -ErrorAction Stop
        if ($org.value -and $org.value[0].displayName) {
            $tenantDisplayName = [string]$org.value[0].displayName
        }
    } catch {
        Write-Warning "Could not read organization display name: $($_.Exception.Message)"
    }
    if ([string]::IsNullOrWhiteSpace($tenantDisplayName) -or $tenantDisplayName -eq $tenantId) {
        $tenantDisplayName = $tenantId
    }
    $script:graphAppCreateTenantDisplayName = $tenantDisplayName
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

Ensure-TransportDataPlatformServicePrincipal

# Microsoft Graph resource app ID
$graphAppId = '00000003-0000-0000-c000-000000000000'

# App role IDs for security investigation report (inbox rules, audit, sign-in, CA, apps, reports, sites, alerts, MFA)
# Plus Application.ReadWrite.All, AppRoleAssignment.ReadWrite.All for Entra Secret Rotate (secret rotation + Add ATR)
# Mail.Read required for Get-MgUserMailFolderMessageRule (inbox rules); MailboxSettings.Read is for auto-reply etc.
# Organization.Read.All: explicit app-only read for GET /organization (tenant displayName); see Microsoft Graph permissions reference.
# SecurityAlert.Read.All / SecurityIncident.Read.All: Defender security alerts & incidents collectors (app-only).
# ExchangeMessageTrace.Read.All: Graph /admin/exchange/tracing/messageTraces when EXO Get-MessageTraceV2 is not in the REST session.
# User.RevokeSessions.All / User.EnableDisableAccount.All: web-runner containment (revoke sessions, block/unblock).
# UserAuthenticationMethod.ReadWrite.All / Device.ReadWrite.All: containment MFA methods and Entra device delete (app-only).
# User-PasswordProfile.ReadWrite.All: containment password reset (app-only also needs User Administrator on the app).
# Existing River Run apps: Update scopes (-UpdateExisting or U at the prompt) or replace (Y).
$appRoleIds = @(
    @{ id = '810c84a8-4a9e-49e6-bf7d-12d183f40d01'; name = 'Mail.Read' }
    @{ id = '89b20d8a-76e2-4057-867b-9961f800b9a4'; name = 'ExchangeMessageTrace.Read.All' }
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
    @{ id = '50483e42-d915-4231-9639-7fdb7fd190e5'; name = 'UserAuthenticationMethod.ReadWrite.All' }
    @{ id = '1138cb37-bd11-4084-a2b7-9f71582aeddb'; name = 'Device.ReadWrite.All' }
    @{ id = '77f3a031-c388-4f99-b373-dc68676a979e'; name = 'User.RevokeSessions.All' }
    @{ id = '3011c876-62b7-4ada-afa2-506cbbecc68c'; name = 'User.EnableDisableAccount.All' }
    @{ id = 'cc117bb9-00cf-4eb8-b580-ea2a878fe8f7'; name = 'User-PasswordProfile.ReadWrite.All' }
    @{ id = '8e8e4742-1d95-4f68-9d56-6ee75648c72a'; name = 'DelegatedPermissionGrant.ReadWrite.All' }
    @{ id = '9e3f62cf-ca93-4989-b6ce-bf83c28f9fe8'; name = 'RoleManagement.ReadWrite.Directory' }
    @{ id = 'dbaae8cf-10b5-4b86-a4a1-f871c94c6695'; name = 'GroupMember.ReadWrite.All' }
    @{ id = '2f51be20-0bb4-4fed-bf7b-db946066c75e'; name = 'DeviceManagementManagedDevices.Read.All' }
    @{ id = '243333ab-4d21-40cb-a475-36241daa0842'; name = 'DeviceManagementManagedDevices.ReadWrite.All' }
    @{ id = '5b07b0dd-2377-4e44-a38d-703f09a0dc3c'; name = 'DeviceManagementManagedDevices.PrivilegedOperations.All' }
)
$requiredResourceAccess = @{
    resourceAccess = $appRoleIds | ForEach-Object { @{ id = $_.id; type = 'Role' } }
    resourceAppId = $graphAppId
}

$displayName = 'River Run Security Investigator'

function Get-ExistingInvestigatorApps {
    $apps = @()
    try {
        $found = Get-MgApplication -Filter "displayName eq '$displayName'" -ErrorAction SilentlyContinue
        if ($found) { $apps = @($found) }
    } catch {}
    $wcmClientId = $null
    try {
        Import-Module (Join-Path $projectRoot 'Modules\GraphAppCredential.psm1') -Force -ErrorAction Stop
        $wcm = Get-GraphAppCredentialFromWCM -TenantId $tenantId
        if ($wcm -and $wcm.ClientId) { $wcmClientId = [string]$wcm.ClientId }
    } catch {}
    if ($wcmClientId) {
        try {
            $byId = Get-MgApplication -Filter "appId eq '$wcmClientId'" -ErrorAction SilentlyContinue
            if ($byId) { return @($byId) }
        } catch {}
    }
    return @($apps)
}

function Update-ExistingInvestigatorAppRoles {
    param([Parameter(Mandatory = $true)]$TargetApps)

    foreach ($app in @($TargetApps)) {
        Write-Host "`nUpdating requiredResourceAccess on $($app.DisplayName) ($($app.AppId))..." -ForegroundColor Yellow
        $payloadAccess = @()
        foreach ($r in @($app.RequiredResourceAccess)) {
            $rid = [string]$r.ResourceAppId
            if (-not $rid -or $rid -eq $graphAppId) { continue }
            $payloadAccess += @{
                resourceAppId  = $rid
                resourceAccess = @($r.ResourceAccess | ForEach-Object { @{ id = [string]$_.Id; type = [string]$_.Type } })
            }
        }
        $payloadAccess += @{
            resourceAppId  = $graphAppId
            resourceAccess = @($appRoleIds | ForEach-Object { @{ id = $_.id; type = 'Role' } })
        }
        Invoke-MgGraphRequest -Method PATCH -Uri "https://graph.microsoft.com/v1.0/applications/$($app.Id)" -Body @{
            requiredResourceAccess = @($payloadAccess)
        } -ErrorAction Stop
        Write-Host "  Manifest updated." -ForegroundColor Green

        $graphSp = Get-MgServicePrincipal -Filter "appId eq '$graphAppId'"
        $sp = @(Get-MgServicePrincipal -Filter "appId eq '$($app.AppId)'" -ErrorAction SilentlyContinue) | Select-Object -First 1
        if (-not $sp) {
            Write-Host "Creating service principal..." -ForegroundColor Yellow
            $sp = New-MgServicePrincipal -AppId $app.AppId
        }

        $existingRole = @{}
        try {
            $uri = "https://graph.microsoft.com/v1.0/servicePrincipals/$($sp.Id)/appRoleAssignments?`$top=999"
            do {
                $page = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
                foreach ($a in @($page.value)) {
                    if ([string]$a.resourceId -eq [string]$graphSp.Id) {
                        $existingRole[[string]$a.appRoleId] = $true
                    }
                }
                $uri = $page.'@odata.nextLink'
            } while ($uri)
        } catch {}

        Write-Host "Granting missing admin consent (app role assignments)..." -ForegroundColor Yellow
        foreach ($role in $appRoleIds) {
            if ($existingRole.ContainsKey([string]$role.id)) {
                Write-Host "  $($role.name) - already consented" -ForegroundColor DarkGray
                $script:graphAppCreateRolesAlready++
                continue
            }
            try {
                New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id -PrincipalId $sp.Id -AppRoleId $role.id -ResourceId $graphSp.Id -ErrorAction Stop | Out-Null
                Write-Host "  $($role.name) - granted" -ForegroundColor Green
                $script:graphAppCreateRolesGranted++
            } catch {
                $msg = $_.Exception.Message
                if ($msg -match 'already|PermissionGrant') {
                    Write-Host "  $($role.name) - already consented" -ForegroundColor DarkGray
                    $script:graphAppCreateRolesAlready++
                } else {
                    Write-Warning "  $($role.name) - $msg"
                    $script:graphAppCreateRolesFailed++
                }
            }
        }
        $script:graphAppCreateClientId = $app.AppId
    }
    $script:graphAppCreateUpdatedExisting = $true
}

function Write-GraphAppUpdateComplete {
    Write-GraphAppCreateResultFile
    Write-Host "`n=== App scopes updated ===" -ForegroundColor Cyan
    Write-Host "TenantId:  $tenantId"
    Write-Host "ClientId:  $($script:graphAppCreateClientId)"
    Write-Host "Granted:   $($script:graphAppCreateRolesGranted)  already: $($script:graphAppCreateRolesAlready)  failed: $($script:graphAppCreateRolesFailed)"
    Write-Host "`nClient secret and WCM credentials were not changed." -ForegroundColor Green
    Write-Host "Run Graph Auth again on this tenant so the worker gets a token with the new roles." -ForegroundColor Yellow
    Write-Host "App-only password reset still needs the User Administrator directory role on this app." -ForegroundColor Yellow
    Write-Host ""
}

$existingApps = Get-ExistingInvestigatorApps

if ($UpdateExisting) {
    if ($existingApps.Count -eq 0) {
        throw "No River Run Security Investigator app (and no matching WCM ClientId) in tenant $tenantId. Use Create Graph App instead."
    }
    Write-Host "`nUpdating $($existingApps.Count) existing app(s):" -ForegroundColor Yellow
    foreach ($a in $existingApps) {
        Write-Host "  - AppId: $($a.AppId)" -ForegroundColor Gray
    }
    Update-ExistingInvestigatorAppRoles -TargetApps $existingApps
    Write-GraphAppUpdateComplete
    exit 0
}

if ($existingApps.Count -gt 0) {
    Write-Host "`nFound $($existingApps.Count) existing app(s) named '$displayName':" -ForegroundColor Yellow
    foreach ($a in $existingApps) {
        Write-Host "  - AppId: $($a.AppId)" -ForegroundColor Gray
    }
    Write-Host "`nU = update scopes on the existing app (keeps the client secret and WCM entry)" -ForegroundColor Yellow
    Write-Host "Y = replace (delete and create a new secret — re-export .eoa-creds)" -ForegroundColor Yellow
    Write-Host "N = cancel" -ForegroundColor Yellow
    Write-Host "Choice (U/Y/N): " -ForegroundColor Yellow -NoNewline
    $reply = Read-Host
    if ($reply -ieq 'u') {
        Update-ExistingInvestigatorAppRoles -TargetApps $existingApps
        Write-GraphAppUpdateComplete
        exit 0
    }
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
        if ([string]::IsNullOrWhiteSpace($tenantDisplayName) -or $tenantDisplayName -eq $tenantId) {
            try {
                $org = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/organization' -ErrorAction Stop
                if ($org.value -and $org.value[0].displayName) {
                    $tenantDisplayName = [string]$org.value[0].displayName
                    $script:graphAppCreateTenantDisplayName = $tenantDisplayName
                }
            } catch {
                Write-Warning "Could not read organization display name before WCM save: $($_.Exception.Message)"
            }
        }
        Import-Module (Join-Path $projectRoot 'Modules\GraphAppCredential.psm1') -Force -ErrorAction Stop
        $nameToStore = if ($tenantDisplayName -and $tenantDisplayName -ne $tenantId) { $tenantDisplayName } else { $null }
        Save-GraphAppCredentialToWCM -TenantId $tenantId -ClientId $app.AppId -ClientSecret $cred.secretText -TenantDisplayName $nameToStore
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
