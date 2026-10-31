<#
.SYNOPSIS
    Scans every client tenant stored in Windows Credential Manager for the compromised ConnectWise Cloud Backup (SkyKick) app.
.DESCRIPTION
    Uses the same EOA/ESR Graph app credentials as the analyzer. For each tenant it looks up the enterprise app
    (service principal) and any local app registration with the known application ID, plus display names that
    start with SkyKick or ConnectWise Cloud Backup.

    When the app is present, it also samples service-principal sign-ins since 2026-10-14 and flags the Huntress
    attacker IP 144.172.64.128. Sign-in logs need AuditLog.Read.All and often Entra ID P1; a missing license is
    recorded on the row and does not fail the scan.

    This script does not disable or delete anything.
.PARAMETER AppId
    Application (client) ID Huntress told customers to find. Default is the SkyKick / ConnectWise Cloud Backup multi-tenant app.
.PARAMETER TenantId
    Limit the scan to these tenant IDs. Omit to scan every tenant that has WCM Graph app credentials.
.PARAMETER CsvPath
    Full results CSV. Defaults to Documents\SkyKick-CloudBackup-scan-<timestamp>.csv.
.PARAMETER SignInSince
    UTC start of the service-principal sign-in sample. Default 2026-10-14T00:00:00Z.
.PARAMETER AttackerIp
    Confirmed attacker address to flag in the sign-in sample.
.EXAMPLE
    .\Scan-SkyKickCloudBackupApp.ps1
.EXAMPLE
    .\Scan-SkyKickCloudBackupApp.ps1 -TenantId 7b5619ee-8c43-464d-b93a-4153c6a72c55
#>
[CmdletBinding()]
param(
    [string]$AppId = '506e08f9-6e9a-4da4-942b-a099a03644ae',

    [string[]]$TenantId,

    [string]$CsvPath,

    [string]$SignInSince = '2026-10-14T00:00:00Z',

    [string]$AttackerIp = '144.172.64.128',

    [string[]]$NamePrefixes = @('SkyKick', 'ConnectWise Cloud Backup'),

    [switch]$PassThru
)

$ErrorActionPreference = 'Stop'

$graphMod = $null
foreach ($dir in @($PSScriptRoot, (Split-Path -Parent $PSScriptRoot))) {
    $candidate = Join-Path $dir 'Modules\GraphAppCredential.psm1'
    if (Test-Path -LiteralPath $candidate) {
        $graphMod = $candidate
        break
    }
}
if (-not $graphMod) {
    throw "GraphAppCredential.psm1 not found. Run this from the exchangeonlineanalyzer repo."
}
Import-Module $graphMod -Force

function Get-ScanGraphErrorText {
    param($ErrorRecord)
    if (-not $ErrorRecord) { return 'Unknown error' }
    $msg = [string]$ErrorRecord.Exception.Message
    $raw = $null
    if ($ErrorRecord.ErrorDetails -and $ErrorRecord.ErrorDetails.Message) {
        $raw = [string]$ErrorRecord.ErrorDetails.Message
    }
    if ($raw) {
        try {
            $j = $raw | ConvertFrom-Json -ErrorAction Stop
            if ($j.error.message) { return ($msg + ' — ' + [string]$j.error.message) }
            if ($j.error.code) { return ($msg + ' — ' + [string]$j.error.code) }
        } catch {
            if ($raw.Length -lt 400) { return ($msg + ' — ' + $raw) }
        }
    }
    return $msg
}

function Invoke-ScanGraphGet {
    param(
        [Parameter(Mandatory = $true)][string]$Token,
        [Parameter(Mandatory = $true)][string]$Uri
    )
    $headers = @{ Authorization = "Bearer $Token"; Accept = 'application/json' }
    return Invoke-RestMethod -Uri $Uri -Headers $headers -Method Get -ErrorAction Stop
}

function Get-ScanGraphValues {
    param(
        [Parameter(Mandatory = $true)][string]$Token,
        [Parameter(Mandatory = $true)][string]$Uri,
        [int]$MaxPages = 4
    )
    $all = [System.Collections.Generic.List[object]]::new()
    $next = $Uri
    $pages = 0
    while ($next -and $pages -lt $MaxPages) {
        $resp = Invoke-ScanGraphGet -Token $Token -Uri $next
        foreach ($item in @($resp.value)) {
            if ($null -ne $item) { $all.Add($item) }
        }
        $next = $resp.'@odata.nextLink'
        $pages++
    }
    return @($all)
}

function Get-ScanTenantToken {
    param(
        [Parameter(Mandatory = $true)][string]$Id,
        [string]$PreferredPrefix = 'EOA'
    )
    $order = [System.Collections.Generic.List[string]]::new()
    if ($PreferredPrefix) { $order.Add($PreferredPrefix) }
    foreach ($p in @('EOA', 'ESR')) {
        if ($order -notcontains $p) { $order.Add($p) }
    }
    $lastError = $null
    foreach ($prefix in $order) {
        $tokenErr = $null
        $token = Get-GraphAppTokenFromWCM -TenantId $Id -Prefix $prefix -FailureVariable 'tokenErr' -ErrorAction SilentlyContinue
        if ($token) {
            return [pscustomobject]@{ Token = $token; Prefix = $prefix; Error = $null }
        }
        if ($tokenErr) { $lastError = [string]$tokenErr }
    }
    if (-not $lastError) { $lastError = 'Token request failed' }
    return [pscustomobject]@{ Token = $null; Prefix = $PreferredPrefix; Error = $lastError }
}

function New-ScanRow {
    param(
        [string]$TenantName,
        [string]$Id,
        [string]$WcmPrefix,
        [bool]$Found,
        [string]$Match,
        [string]$ObjectType,
        [string]$DisplayName,
        [string]$FoundAppId,
        [string]$ServicePrincipalId,
        $SignInEnabled,
        [string]$PublisherName,
        [string]$SignInsChecked,
        $SignInCount,
        [string]$AttackerIpSeen,
        [string]$SignInIps,
        [string]$ErrorText
    )
    [pscustomobject]@{
        TenantName         = $TenantName
        TenantId           = $Id
        WcmPrefix          = $WcmPrefix
        Found              = $Found
        Match              = $Match
        ObjectType         = $ObjectType
        DisplayName        = $DisplayName
        AppId              = $FoundAppId
        ServicePrincipalId = $ServicePrincipalId
        SignInEnabled      = $SignInEnabled
        PublisherName      = $PublisherName
        SignInsChecked     = $SignInsChecked
        SignInCount        = $SignInCount
        AttackerIpSeen     = $AttackerIpSeen
        SignInIps          = $SignInIps
        Error              = $ErrorText
    }
}

if ([string]::IsNullOrWhiteSpace($CsvPath)) {
    $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $CsvPath = Join-Path ([Environment]::GetFolderPath('MyDocuments')) "SkyKick-CloudBackup-scan-$stamp.csv"
}

$wanted = @{}
if ($TenantId) {
    foreach ($id in $TenantId) {
        $norm = ($id -replace '[\{\}]', '').Trim()
        if ($norm) { $wanted[$norm] = $true }
    }
}

$tenants = @(Get-WCMTenantListWithNamesForAppRegCombo -SkipGraphLookup -ErrorAction SilentlyContinue)
if ($wanted.Count -gt 0) {
    $tenants = @($tenants | Where-Object { $wanted.ContainsKey([string]$_.TenantId) })
}
if ($tenants.Count -eq 0) {
    throw "No WCM Graph app credentials matched. Store client app credentials in this Windows profile first."
}

Write-Host "Scanning $($tenants.Count) tenant(s) for app $AppId (SkyKick / ConnectWise Cloud Backup)."
Write-Host "CSV: $CsvPath"
Write-Host ""

$rows = [System.Collections.Generic.List[object]]::new()
$index = 0
foreach ($tenant in $tenants) {
    $index++
    $id = [string]$tenant.TenantId
    $name = [string]$tenant.DisplayName
    if ([string]::IsNullOrWhiteSpace($name)) { $name = $id }
    $preferred = if ($tenant.Source) { [string]$tenant.Source } else { 'EOA' }
    Write-Host ("[{0}/{1}] {2}" -f $index, $tenants.Count, $name)

    $auth = Get-ScanTenantToken -Id $id -PreferredPrefix $preferred
    if (-not $auth.Token) {
        $rows.Add((New-ScanRow -TenantName $name -Id $id -WcmPrefix $auth.Prefix -Found $false -ErrorText $auth.Error))
        Write-Host "  token failed: $($auth.Error)" -ForegroundColor DarkYellow
        continue
    }

    if ([string]::IsNullOrWhiteSpace([string]$tenant.DisplayName)) {
        try {
            $org = Invoke-ScanGraphGet -Token $auth.Token -Uri 'https://graph.microsoft.com/v1.0/organization?$select=displayName'
            $orgName = [string]@($org.value)[0].displayName
            if (-not [string]::IsNullOrWhiteSpace($orgName)) { $name = $orgName.Trim() }
        } catch { }
    }

    $hits = @{}
    $queryError = $null

    $encodedApp = [uri]::EscapeDataString("appId eq '$AppId'")
    $spById = "https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=$encodedApp&`$select=id,appId,displayName,accountEnabled,servicePrincipalType,publisherName,appOwnerOrganizationId"
    try {
        foreach ($sp in @(Get-ScanGraphValues -Token $auth.Token -Uri $spById -MaxPages 2)) {
            $key = 'sp:' + [string]$sp.id
            $hits[$key] = [pscustomobject]@{ Item = $sp; Match = 'KnownAppId'; ObjectType = 'EnterpriseApp' }
        }
    } catch {
        $queryError = Get-ScanGraphErrorText $_
    }

    foreach ($prefixName in $NamePrefixes) {
        $safe = $prefixName.Replace("'", "''")
        $encodedName = [uri]::EscapeDataString("startswith(displayName,'$safe')")
        $spByName = "https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=$encodedName&`$select=id,appId,displayName,accountEnabled,servicePrincipalType,publisherName,appOwnerOrganizationId&`$top=25"
        try {
            foreach ($sp in @(Get-ScanGraphValues -Token $auth.Token -Uri $spByName -MaxPages 2)) {
                $key = 'sp:' + [string]$sp.id
                if ($hits.ContainsKey($key)) { continue }
                $match = if ([string]$sp.appId -eq $AppId) { 'KnownAppId' } else { 'DisplayName' }
                $hits[$key] = [pscustomobject]@{ Item = $sp; Match = $match; ObjectType = 'EnterpriseApp' }
            }
        } catch {
            if (-not $queryError) { $queryError = Get-ScanGraphErrorText $_ }
        }
    }

    $appById = "https://graph.microsoft.com/v1.0/applications?`$filter=$encodedApp&`$select=id,appId,displayName,signInAudience"
    try {
        foreach ($app in @(Get-ScanGraphValues -Token $auth.Token -Uri $appById -MaxPages 1)) {
            $key = 'app:' + [string]$app.id
            $hits[$key] = [pscustomobject]@{ Item = $app; Match = 'KnownAppId'; ObjectType = 'AppRegistration' }
        }
    } catch {
        if (-not $queryError) { $queryError = Get-ScanGraphErrorText $_ }
    }

    if ($hits.Count -eq 0) {
        $rows.Add((New-ScanRow -TenantName $name -Id $id -WcmPrefix $auth.Prefix -Found $false -ErrorText $queryError))
        if ($queryError) {
            Write-Host "  lookup failed: $queryError" -ForegroundColor DarkYellow
        }
        continue
    }

    foreach ($hit in @($hits.Values)) {
        $item = $hit.Item
        $signInEnabled = $null
        if ($null -ne $item.accountEnabled) { $signInEnabled = [bool]$item.accountEnabled }
        $signInsChecked = ''
        $signInCount = $null
        $attackerSeen = ''
        $ipText = ''
        $signInError = $null

        if ($hit.ObjectType -eq 'EnterpriseApp' -and [string]$item.appId -eq $AppId) {
            # Service principal sign-ins are not on the v1.0 signIn type. Beta accepts signInEventTypes.
            $signFilter = "signInEventTypes/any(t: t eq 'servicePrincipal') and appId eq '$AppId' and createdDateTime ge $SignInSince"
            $signUri = 'https://graph.microsoft.com/beta/auditLogs/signIns?$filter=' + [uri]::EscapeDataString($signFilter) + '&$top=50&$select=createdDateTime,ipAddress,appDisplayName,status'
            try {
                $signIns = @(Get-ScanGraphValues -Token $auth.Token -Uri $signUri -MaxPages 3)
                $signInsChecked = 'yes'
                $signInCount = $signIns.Count
                $ips = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
                $seenAttacker = $false
                foreach ($signIn in $signIns) {
                    $ip = [string]$signIn.ipAddress
                    if ([string]::IsNullOrWhiteSpace($ip)) { continue }
                    [void]$ips.Add($ip.Trim())
                    if ($ip.Trim() -eq $AttackerIp) { $seenAttacker = $true }
                }
                $attackerSeen = if ($seenAttacker) { 'yes' } else { 'no' }
                $ipText = (@($ips) | Sort-Object) -join '; '
            } catch {
                $signInsChecked = 'no'
                $signInError = Get-ScanGraphErrorText $_
            }
        }

        $errorText = $signInError
        if ($queryError -and -not $errorText) { $errorText = $queryError }
        $rows.Add((New-ScanRow -TenantName $name -Id $id -WcmPrefix $auth.Prefix -Found $true -Match $hit.Match -ObjectType $hit.ObjectType -DisplayName ([string]$item.displayName) -FoundAppId ([string]$item.appId) -ServicePrincipalId ([string]$item.id) -SignInEnabled $signInEnabled -PublisherName ([string]$item.publisherName) -SignInsChecked $signInsChecked -SignInCount $signInCount -AttackerIpSeen $attackerSeen -SignInIps $ipText -ErrorText $errorText))

        $enabledText = if ($signInEnabled -eq $true) { 'SIGN-IN ENABLED' } elseif ($signInEnabled -eq $false) { 'sign-in disabled' } else { $hit.ObjectType }
        $color = if ($signInEnabled -eq $true -and $hit.Match -eq 'KnownAppId') { 'Red' } elseif ($hit.Match -eq 'KnownAppId') { 'Yellow' } else { 'Cyan' }
        $ipNote = ''
        if ($attackerSeen -eq 'yes') { $ipNote = " ATTACKER IP $AttackerIp" }
        elseif ($signInCount -gt 0) { $ipNote = " sign-ins=$signInCount" }
        Write-Host ("  HIT {0} | {1} | {2} | {3}{4}" -f $hit.Match, $item.displayName, $item.appId, $enabledText, $ipNote) -ForegroundColor $color
        if ($signInError) { Write-Host "  sign-in sample unavailable: $signInError" -ForegroundColor DarkYellow }
    }
}

$out = @($rows)
$dir = Split-Path -Parent $CsvPath
if ($dir -and -not (Test-Path -LiteralPath $dir)) {
    New-Item -ItemType Directory -Path $dir -Force | Out-Null
}
$out | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8

$known = @($out | Where-Object { $_.Found -and $_.Match -eq 'KnownAppId' -and $_.ObjectType -eq 'EnterpriseApp' })
$enabled = @($known | Where-Object { $_.SignInEnabled -eq $true })
$attacker = @($known | Where-Object { $_.AttackerIpSeen -eq 'yes' })
$failed = @($out | Where-Object { -not $_.Found -and $_.Error })

Write-Host ""
Write-Host "Tenants scanned: $($tenants.Count)"
Write-Host "Known Cloud Backup enterprise app present: $($known.Count)"
Write-Host "Still enabled for sign-in: $($enabled.Count)"
Write-Host "Attacker IP ${AttackerIp} in sampled sign-ins: $($attacker.Count)"
Write-Host "Tenants that could not be queried: $($failed.Count)"
Write-Host "Wrote $CsvPath"

if ($enabled.Count -gt 0) {
    Write-Host ""
    Write-Host "Disable these in Entra (Enterprise applications > Properties > Enabled for users to sign-in = No). Do not delete them." -ForegroundColor Red
    $enabled | Select-Object TenantName, TenantId, DisplayName, AppId, ServicePrincipalId, AttackerIpSeen | Format-Table -AutoSize | Out-Host
}

if ($PassThru) { return $out }
