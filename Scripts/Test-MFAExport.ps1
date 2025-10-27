param(
    [string]$OutputFolder,
    [int]$ThrottleLimit = 4,
    [switch]$Diag
)

$ErrorActionPreference = 'Stop'
try { $PSStyle.OutputRendering = 'Ansi' } catch {}

function Write-Info($msg) { Write-Host $msg -ForegroundColor Cyan }
function Write-Ok($msg) { Write-Host $msg -ForegroundColor Green }
function Write-Warn($msg) { Write-Warning $msg }
function Write-Err($msg) { Write-Host $msg -ForegroundColor Red }

# Resolve script root
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$root = Split-Path -Parent $here

# Import modules
Import-Module (Join-Path $root 'Modules/GraphOnline.psm1') -Force -ErrorAction Stop
Import-Module (Join-Path $root 'Modules/ExportUtils.psm1') -Force -ErrorAction Stop

# Connect to Graph (scopes include Policy.Read.All, Reports.Read.All, UserAuthenticationMethod.Read.All)
Write-Info 'Connecting to Microsoft Graph...'
$null = Connect-GraphService

# Build output folder (tenant scoped)
if ([string]::IsNullOrWhiteSpace($OutputFolder)) {
    $tenantName = 'Tenant'
    try {
        $ti = $null
        $bi = Join-Path $root 'Modules/BrowserIntegration.psm1'
        if (Test-Path $bi) { Import-Module $bi -Force -ErrorAction SilentlyContinue }
        try { $ti = Get-TenantIdentity } catch {}
        if ($ti) {
            if ($ti.TenantDisplayName) { $tenantName = $ti.TenantDisplayName }
            elseif ($ti.PrimaryDomain) { $tenantName = $ti.PrimaryDomain }
        }
    } catch {}
    $invalid = [System.IO.Path]::GetInvalidFileNameChars()
    $safeName = ($tenantName.ToCharArray() | ForEach-Object { if ($invalid -contains $_) { '-' } else { $_ } }) -join ''
    $safeName = ($safeName -replace '\s+', ' ').Trim()
    if ($safeName.Length -gt 80) { $safeName = $safeName.Substring(0,80) }
    $base = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'ExchangeOnlineAnalyzer\SecurityInvestigation'
    $tenantRoot = Join-Path $base $safeName
    $ts = Get-Date -Format 'yyyyMMdd_HHmmss'
    $OutputFolder = Join-Path $tenantRoot ("MFA_Test_" + $ts)
}

if (-not (Test-Path $OutputFolder)) { New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null }
Write-Info ("Output: {0}" -f $OutputFolder)

# Collect MFA coverage (same function as the full report)
Write-Info 'Evaluating MFA coverage (per-user, Security Defaults, and Conditional Access)...'
$mfa = $null
try { $mfa = Get-MfaCoverageReport -Parallel -ThrottleLimit $ThrottleLimit } catch { $mfa = Get-MfaCoverageReport }
if (-not $mfa) { Write-Err 'MFA coverage returned null.'; exit 1 }
$secDefaults = [bool]$mfa.SecurityDefaultsEnabled
$tenantCaReq = [bool]$mfa.CAPoliciesRequireMfa
$usersCov = @()
if ($mfa.Users) { $usersCov = $mfa.Users }
Write-Ok ("Coverage objects: {0}; SecurityDefaults: {1}; CARequiresMfa: {2}" -f ($usersCov.Count), $secDefaults, $tenantCaReq)

# Join to authoritative Graph user list to ensure DisplayName/UPN populate
Write-Info 'Fetching users from Graph for join...'
$users = @()
try {
    $uri = 'https://graph.microsoft.com/v1.0/users?$select=id,displayName,userPrincipalName,accountEnabled&$top=999'
    do {
        $page = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
        if ($page.value) { $users += ($page.value | Where-Object { $_.accountEnabled -ne $false }) }
        $uri = $page.'@odata.nextLink'
    } while ($uri)
} catch {
    Write-Warn ("Failed to enumerate users from Graph: {0}" -f $_.Exception.Message)
}

if ($Diag) {
    Write-Info 'Diagnostics: Checking Conditional Access policies requiring MFA...'
    try {
        $cap = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies?$top=999' -ErrorAction Stop
        $mfaPols = @()
        if ($cap.value) {
            foreach ($p in $cap.value) {
                if ($p.state -eq 'enabled' -and $p.grantControls) {
                    if ($p.grantControls.builtInControls -contains 'mfa' -or $p.grantControls.authenticationStrength) { $mfaPols += $p }
                }
            }
        }
        Write-Ok ("CA policies fetched: {0}; MFA-requiring: {1}" -f (@($cap.value).Count), $mfaPols.Count)
        if (-not $cap.value) { Write-Warn 'No CA policies returned. Ensure Policy.Read.All is consented.' }
    } catch { Write-Warn ("CA policy fetch failed: {0}" -f $_.Exception.Message) }

    Write-Info 'Diagnostics: Sampling authentication methods for first 3 users...'
    $sample = $users | Select-Object -First 3
    foreach ($su in $sample) {
        try {
            $m = Invoke-MgGraphRequest -Method GET -Uri ("https://graph.microsoft.com/v1.0/users/{0}/authentication/methods" -f $su.id) -ErrorAction Stop
            $has = $false
            if ($m.value) {
                foreach ($mm in $m.value) {
                    $otype = $mm.'@odata.type'
                    if ($otype -match 'microsoftAuthenticator|phoneAuthentication|softwareOath|fido2|temporaryAccessPass') { $has = $true; break }
                }
            }
            Write-Host (" - {0}: methods={1} anyMfaLike={2}" -f $su.userPrincipalName, (@($m.value).Count), $has) -ForegroundColor Yellow
        } catch { Write-Warn (" - {0}: methods query failed: {1}" -f $su.userPrincipalName, $_.Exception.Message) }
    }
}

$covById = @{}; $covByUpn = @{}
foreach ($c in $usersCov) {
    $cid = $null; $cupn = $null
    if ($c.PSObject.Properties['UserId']) { $cid = $c.UserId } elseif ($c.PSObject.Properties['id']) { $cid = $c.id }
    if ($c.PSObject.Properties['UserPrincipalName']) { $cupn = $c.UserPrincipalName } elseif ($c.PSObject.Properties['userPrincipalName']) { $cupn = $c.userPrincipalName }
    if ($cid) { $covById[$cid] = $c }
    if ($cupn) { $covByUpn[$cupn.ToLower()] = $c }
}

$mfaRows = New-Object System.Collections.Generic.List[object]
foreach ($u in $users) {
    $cov = $null
    if ($covById.ContainsKey($u.id)) { $cov = $covById[$u.id] }
    elseif ($covByUpn.ContainsKey($u.userPrincipalName.ToLower())) { $cov = $covByUpn[$u.userPrincipalName.ToLower()] }

    $perUser = $false; $caReq = $tenantCaReq
    if ($cov -ne $null) {
        if ($cov.PSObject.Properties['PerUserMfaEnabled']) { $perUser = [bool]$cov.PerUserMfaEnabled }
        if ($cov.PSObject.Properties['CARequiresMfa']) { $caReq = [bool]$cov.CARequiresMfa }
    }
    $covered = [bool]($perUser -or $secDefaults -or $caReq)

    $mfaRows.Add([pscustomobject]@{
        DisplayName       = $u.displayName
        UserPrincipalName = $u.userPrincipalName
        PerUserMfaEnabled = $perUser
        SecurityDefaults  = $secDefaults
        CARequiresMfa     = $caReq
        MfaCovered        = $covered
    }) | Out-Null
}

# Export
$mfaCsv = Join-Path $OutputFolder 'MFAStatus.csv'
try {
    if ($mfaRows.Count -gt 0) {
        $mfaRows | Export-Csv -Path $mfaCsv -NoTypeInformation -Encoding UTF8
        Write-Ok ("Wrote {0}" -f $mfaCsv)
    } else {
        'DisplayName,UserPrincipalName,PerUserMfaEnabled,SecurityDefaults,CARequiresMfa,MfaCovered' | Set-Content -Path $mfaCsv -Encoding utf8
        Write-Warn 'No MFA rows computed; wrote header only.'
    }
} catch {
    Write-Err ("Failed to write {0}: {1}" -f $mfaCsv, $_.Exception.Message)
    exit 1
}

# Summary
$coveredCount = ($mfaRows | Where-Object { $_.MfaCovered -eq $true }).Count
Write-Ok ("Users evaluated: {0}; Covered: {1}" -f $mfaRows.Count, $coveredCount)
