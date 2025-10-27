# Test script for v8.1b integration
# Purpose: connect to EXO and Graph, then test Firefox container/profile deep links

param(
    [string]$ProfileName,
    [string]$ContainerName,
    [ValidateSet('SignIns','RestrictedEntities','ConditionalAccess')]
    [string[]]$Targets = @('SignIns','RestrictedEntities','ConditionalAccess')
)

$ErrorActionPreference = 'Stop'

# Import modules
Import-Module (Join-Path $PSScriptRoot '..\Modules\GraphOnline.psm1') -Force
Import-Module (Join-Path $PSScriptRoot '..\Modules\BrowserIntegration.psm1') -Force
Import-Module ExchangeOnlineManagement -ErrorAction SilentlyContinue | Out-Null

Write-Host 'Connecting to Exchange Online...' -ForegroundColor Yellow
try { Connect-ExchangeOnline -ErrorAction Stop } catch { Write-Warning ("EXO connect failed: {0}" -f $_.Exception.Message) }

Write-Host 'Connecting to Microsoft Graph...' -ForegroundColor Yellow
try { Connect-GraphService | Out-Null } catch { Write-Warning ("Graph connect failed: {0}" -f $_.Exception.Message) }

# Summarize connections
$exoConnected = $false
try { Get-OrganizationConfig | Out-Null; $exoConnected = $true } catch {}
$mgConnected = $false
try { $ctx = Get-MgContext; if ($ctx -and $ctx.Account) { $mgConnected = $true } } catch {}
$exoStr = if ($exoConnected) { 'Connected' } else { 'Not Connected' }
$mgStr  = if ($mgConnected)  { 'Connected' } else { 'Not Connected' }
Write-Host ("EXO: {0} | Graph: {1}" -f $exoStr, $mgStr) -ForegroundColor Green

# Resolve profile and container
$profiles = Get-FirefoxProfiles
if (-not $ProfileName) {
    $def = ($profiles | Where-Object { $_.Default -eq $true } | Select-Object -First 1)
    if ($def) { $ProfileName = $def.Name } elseif ($profiles.Count -gt 0) { $ProfileName = $profiles[0].Name }
}
if (-not $ProfileName) { Write-Warning 'No Firefox profiles found. Ensure Firefox is installed and profiles.ini exists.'; exit 1 }

$prof = ($profiles | Where-Object { $_.Name -eq $ProfileName } | Select-Object -First 1)
$ppath = if ($prof.Path -like '*:*') { $prof.Path } else { Join-Path (Join-Path $env:APPDATA 'Mozilla\Firefox') $prof.Path }
$containers = Get-FirefoxContainers -ProfilePath $ppath

if (-not $ContainerName -and $containers.Count -gt 0) {
    $tenant = Get-TenantIdentity
    $best = Select-BestContainer -Containers $containers -TenantIdentity $tenant
    if ($best) { $ContainerName = $best.name }
}

$containerStr = if ($ContainerName) { $ContainerName } else { '(none)' }
Write-Host ("Profile: {0} | Container: {1}" -f $ProfileName, $containerStr) -ForegroundColor Cyan

# Open targets
foreach ($t in $Targets) {
    Write-Host ("Opening: {0}" -f $t) -ForegroundColor Yellow
    try { Open-EntraDeepLink -ProfileName $ProfileName -ContainerName $ContainerName -Target $t } catch { Write-Warning $_.Exception.Message }
    Start-Sleep -Seconds 1
}

Write-Host 'Done.' -ForegroundColor Green
