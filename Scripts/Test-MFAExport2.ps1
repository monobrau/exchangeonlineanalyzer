param(
    [string]$OutputFolder,
    [int]$ThrottleLimit = 6,
    [switch]$Delegated,
    [switch]$Diag
)

$ErrorActionPreference = 'Stop'
try { $PSStyle.OutputRendering = 'Ansi' } catch {}
function Write-Info($m){ Write-Host $m -ForegroundColor Cyan }
function Write-Ok($m){ Write-Host $m -ForegroundColor Green }
function Write-Warn($m){ Write-Warning $m }
function Write-Err($m){ Write-Host $m -ForegroundColor Red }

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$root = Split-Path -Parent $here
Import-Module (Join-Path $root 'Scripts/lib/GraphAppAuth.psm1') -Force -ErrorAction Stop

# Auth selection: app (from Scripts/test.config.json) unless -Delegated
$cfg = Get-TestConfig
$useApp = (-not $Delegated) -and $cfg.TenantId -and $cfg.ClientId -and $cfg.ClientSecret
$token = $null
if ($useApp) {
    Write-Info 'Acquiring Graph application token...'
    $token = Get-GraphAppToken -TenantId $cfg.TenantId -ClientId $cfg.ClientId -ClientSecret $cfg.ClientSecret
    $headers = @{ Authorization = "Bearer $token" }
} else {
    Write-Info 'Connecting to Microsoft Graph (delegated interactive with device code fallback)...'
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    $scopes = @('User.Read.All','Directory.Read.All','Policy.Read.All','Reports.Read.All','UserAuthenticationMethod.Read.All')
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue } catch {}
    $connected = $false
    try { Connect-MgGraph -Scopes $scopes -ErrorAction Stop | Out-Null; $connected = $true } catch {}
    if (-not $connected) { try { Connect-MgGraph -Scopes $scopes -UseDeviceCode -ErrorAction Stop | Out-Null; $connected = $true } catch {} }
    if (-not $connected) { throw 'Failed to connect to Microsoft Graph.' }
}

# Output folder
if ([string]::IsNullOrWhiteSpace($OutputFolder)) {
    $tenant = 'Tenant'
    try {
        if ($useApp) {
            $org = Invoke-GraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/organization?$select=displayName' -Headers $headers
            if ($org.value -and $org.value[0].displayName) { $tenant = $org.value[0].displayName }
        } else {
            $org = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/organization?$select=displayName'
            if ($org.value -and $org.value[0].displayName) { $tenant = $org.value[0].displayName }
        }
    } catch {}
    $invalid = [System.IO.Path]::GetInvalidFileNameChars()
    $safe = ($tenant.ToCharArray() | ForEach-Object { if ($invalid -contains $_) { '-' } else { $_ } }) -join ''
    $safe = ($safe -replace '\s+', ' ').Trim(); if ($safe.Length -gt 80) { $safe = $safe.Substring(0,80) }
    $base = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'ExchangeOnlineAnalyzer\SecurityInvestigation'
    $tenantRoot = Join-Path $base $safe
    $ts = Get-Date -Format 'yyyyMMdd_HHmmss'
    $OutputFolder = Join-Path $tenantRoot ("MFA_Test_" + $ts)
}
if (-not (Test-Path $OutputFolder)) { New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null }
Write-Info ("Output: {0}" -f $OutputFolder)

# Tenant-wide inputs
Write-Info 'Fetching Security Defaults and Conditional Access policies...'
$secDefaults = $false
try {
    if ($useApp) { $sd = Invoke-GraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/policies/identitySecurityDefaultsEnforcementPolicy' -Headers $headers }
    else { $sd = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/policies/identitySecurityDefaultsEnforcementPolicy' }
    if ($sd -and $sd.isEnabled -ne $null) { $secDefaults = [bool]$sd.isEnabled }
} catch {}

$caPolicies = @()
try {
    if ($useApp) { $cap = Invoke-GraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies?$top=999' -Headers $headers }
    else { $cap = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies?$top=999' }
    if ($cap.value) { $caPolicies = $cap.value }
} catch {}
$mfaPolicies = @()
foreach ($p in $caPolicies) {
    if ($p.state -eq 'enabled' -and $p.grantControls) {
        if ($p.grantControls.builtInControls -contains 'mfa' -or $p.grantControls.authenticationStrength) { $mfaPolicies += $p }
    }
}
if ($Diag) { Write-Ok ("CA policies: {0} (MFA policies: {1})" -f $caPolicies.Count, $mfaPolicies.Count) }

# Users
Write-Info 'Enumerating users...'
$users = @()
try {
    $uri = 'https://graph.microsoft.com/v1.0/users?$select=id,displayName,userPrincipalName,accountEnabled&$top=999'
    do {
        if ($useApp) { $page = Invoke-GraphRequest -Method GET -Uri $uri -Headers $headers }
        else { $page = Invoke-MgGraphRequest -Method GET -Uri $uri }
        if ($page.value) { $users += ($page.value | Where-Object { $_.accountEnabled -ne $false }) }
        $uri = $page.'@odata.nextLink'
    } while ($uri)
} catch { Write-Err ("Failed to enumerate users: {0}" -f $_.Exception.Message); exit 1 }
Write-Ok ("Users to evaluate: {0}" -f $users.Count)

# Per-user evaluation helpers
function Test-UserPerUserMfa {
    param([string]$UserId)
    try {
        $endpoint = "https://graph.microsoft.com/v1.0/users/$UserId/authentication/methods"
        $m = if ($useApp) { Invoke-GraphRequest -Method GET -Uri $endpoint -Headers $headers } else { Invoke-MgGraphRequest -Method GET -Uri $endpoint }
        if ($m.value) { foreach ($mm in $m.value) { if ($mm.'@odata.type' -match 'microsoftAuthenticator|phoneAuthentication|softwareOath|fido2|temporaryAccessPass') { return $true } } }
    } catch {}
    return $false
}

function Get-UserGroupsRoles {
    param([string]$UserId)
    $g=@(); $rt=@()
    try {
        $muri = "https://graph.microsoft.com/v1.0/users/$UserId/memberOf?$select=id,displayName,roleTemplateId&$top=999"
        do {
            $mr = if ($useApp) { Invoke-GraphRequest -Method GET -Uri $muri -Headers $headers } else { Invoke-MgGraphRequest -Method GET -Uri $muri }
            if ($mr.value) { foreach ($m in $mr.value) { if ($m.'@odata.type' -eq '#microsoft.graph.group') { $g += $m.id } elseif ($m.'@odata.type' -eq '#microsoft.graph.directoryRole') { if ($m.roleTemplateId) { $rt += $m.roleTemplateId } } } }
            $muri = $mr.'@odata.nextLink'
        } while ($muri)
    } catch {}
    return @{ Groups=$g; RoleTemplates=$rt }
}

function Test-UserCaRequiresMfa {
    param([object]$User,[object]$UserCtx)
    foreach ($p in $mfaPolicies) {
        if ($p.state -ne 'enabled') { continue }
        $cond = $p.conditions; if (-not $cond) { continue }
        $uc = $cond.users
        $applies = $false
        if ($uc) {
            if ($uc.includeUsers -and ($uc.includeUsers -contains 'All' -or $uc.includeUsers -contains $User.id)) { $applies = $true }
            if (-not $applies -and $uc.includeGroups) { if (@($uc.includeGroups) -ne $null) { if ($uc.includeGroups | Where-Object { $UserCtx.Groups -contains $_ }) { $applies = $true } } }
            if (-not $applies -and $uc.includeRoles) { if (@($uc.includeRoles) -ne $null) { foreach ($rid in $uc.includeRoles) { if ($UserCtx.RoleTemplates -contains $rid) { $applies = $true; break } } } }
            if ($uc.excludeUsers -and ($uc.excludeUsers -contains $User.id)) { $applies = $false }
            if ($uc.excludeGroups -and (@($uc.excludeGroups) -ne $null)) { if ($uc.excludeGroups | Where-Object { $UserCtx.Groups -contains $_ }) { $applies = $false } }
            if ($uc.excludeRoles -and (@($uc.excludeRoles) -ne $null)) { foreach ($rid in $uc.excludeRoles) { if ($UserCtx.RoleTemplates -contains $rid) { $applies = $false; break } } }
        }
        if ($applies) { return $true }
    }
    return $false
}

# Evaluate users (parallel when app mode)
$rows = New-Object System.Collections.Generic.List[object]
if ($useApp -and $PSVersionTable.PSVersion.Major -ge 7) {
    $par = $users | ForEach-Object -Parallel {
        param($u)
        Import-Module (Join-Path (Split-Path -Parent $using:here) 'Scripts/lib/GraphAppAuth.psm1') -Force -ErrorAction Stop
        $headers = @{ Authorization = "Bearer $using:token" }
        function _req($uri){ Invoke-GraphRequest -Method GET -Uri $uri -Headers $headers }
        $per = $false
        try { $m = _req ("https://graph.microsoft.com/v1.0/users/$($u.id)/authentication/methods"); if ($m.value) { foreach ($mm in $m.value) { if ($mm.'@odata.type' -match 'microsoftAuthenticator|phoneAuthentication|softwareOath|fido2|temporaryAccessPass') { $per = $true; break } } } } catch {}
        $g=@(); $rt=@();
        try { $mu = "https://graph.microsoft.com/v1.0/users/$($u.id)/memberOf?$select=id,displayName,roleTemplateId&$top=999"; do { $mr = _req $mu; if ($mr.value) { foreach ($m in $mr.value) { if ($m.'@odata.type' -eq '#microsoft.graph.group') { $g += $m.id } elseif ($m.'@odata.type' -eq '#microsoft.graph.directoryRole') { if ($m.roleTemplateId) { $rt += $m.roleTemplateId } } } } $mu = $mr.'@odata.nextLink' } while ($mu) } catch {}
        $ca = $false
        foreach ($p in $using:mfaPolicies) { if ($p.state -ne 'enabled') { continue } $cond = $p.conditions; if (-not $cond) { continue }; $uc = $cond.users; $applies=$false; if ($uc) { if ($uc.includeUsers -and ($uc.includeUsers -contains 'All' -or $uc.includeUsers -contains $u.id)) { $applies=$true } if (-not $applies -and $uc.includeGroups) { if (@($uc.includeGroups) -ne $null) { if ($uc.includeGroups | Where-Object { $g -contains $_ }) { $applies=$true } } } if (-not $applies -and $uc.includeRoles) { if (@($uc.includeRoles) -ne $null) { foreach ($rid in $uc.includeRoles) { if ($rt -contains $rid) { $applies=$true; break } } } } if ($uc.excludeUsers -and ($uc.excludeUsers -contains $u.id)) { $applies=$false } if ($uc.excludeGroups -and (@($uc.excludeGroups) -ne $null)) { if ($uc.excludeGroups | Where-Object { $g -contains $_ }) { $applies=$false } } if ($uc.excludeRoles -and (@($uc.excludeRoles) -ne $null)) { foreach ($rid in $uc.excludeRoles) { if ($rt -contains $rid) { $applies=$false; break } } } } if ($applies) { $ca=$true; break } }
        [pscustomobject]@{ DisplayName=$u.displayName; UserPrincipalName=$u.userPrincipalName; PerUserMfaEnabled=$per; SecurityDefaults=[bool]$using:secDefaults; CARequiresMfa=$ca; MfaCovered=[bool]($per -or $using:secDefaults -or $ca) }
    } -ThrottleLimit $ThrottleLimit
    if ($par) { foreach ($o in $par) { if ($o) { [void]$rows.Add($o) } } }
} else {
    foreach ($u in $users) { $per = Test-UserPerUserMfa -UserId $u.id; $ctx = Get-UserGroupsRoles -UserId $u.id; $ca = Test-UserCaRequiresMfa -User $u -UserCtx $ctx; $rows.Add([pscustomobject]@{ DisplayName=$u.displayName; UserPrincipalName=$u.userPrincipalName; PerUserMfaEnabled=$per; SecurityDefaults=$secDefaults; CARequiresMfa=$ca; MfaCovered=[bool]($per -or $secDefaults -or $ca) }) | Out-Null }
}

# Export
$csv = Join-Path $OutputFolder 'MFAStatus.csv'
if ($rows.Count -gt 0) {
    $rows | Export-Csv -Path $csv -NoTypeInformation -Encoding UTF8
    $covered = ($rows | Where-Object { $_.MfaCovered }).Count
    Write-Ok ("Wrote {0} users; Covered={1} -> {2}" -f $rows.Count, $covered, $csv)
} else {
    'DisplayName,UserPrincipalName,PerUserMfaEnabled,SecurityDefaults,CARequiresMfa,MfaCovered' | Set-Content -Path $csv -Encoding utf8
    Write-Warn 'No rows computed; wrote header only.'
}

