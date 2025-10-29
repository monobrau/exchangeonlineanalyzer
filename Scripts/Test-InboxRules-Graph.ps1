param(
    [string]$OutputFolder,
    [int]$ThrottleLimit = 8,
    [int]$BatchSize = 20,
    [switch]$Diag,
    [switch]$Delegated
)

$ErrorActionPreference = 'Stop'
try { $PSStyle.OutputRendering = 'Ansi' } catch {}
function Write-Info($m){ Write-Host $m -ForegroundColor Cyan }
function Write-Ok($m){ Write-Host $m -ForegroundColor Green }
function Write-Warn2($m){ Write-Warning $m }
function Write-Err($m){ Write-Host $m -ForegroundColor Red }

# Resolve roots and import shared auth
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$root = Split-Path -Parent $here
Import-Module (Join-Path $root 'Scripts/lib/GraphAppAuth.psm1') -Force -ErrorAction Stop

# Load test configuration; if missing and not explicitly using app mode, fall back to delegated
$cfg = Get-TestConfig
if (-not $cfg.TenantId -or -not $cfg.ClientId -or -not $cfg.ClientSecret) { $Delegated = $true }
if ($cfg.ThrottleLimit) { $ThrottleLimit = [int]$cfg.ThrottleLimit }

# Acquire token (delegated or application)
if ($Delegated) {
    Write-Info 'Connecting to Microsoft Graph (delegated interactive)...'
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    # Scopes: list users and read mail rules
    $scopes = @('User.Read.All','Mail.Read','MailboxSettings.Read')
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue } catch {}
    $connected = $false
    try {
        Connect-MgGraph -Scopes $scopes -ErrorAction Stop | Out-Null
        $connected = $true
    } catch {
        $msg = $_.Exception.Message
        if ($msg -match 'BaseAbstractApplicationBuilder.*WithLogging|InteractiveBrowserCredential authentication failed') {
            Write-Warning 'Graph module conflict detected; attempting repair and device code fallback...'
            try { Import-Module (Join-Path (Split-Path -Parent $root) 'Modules/GraphOnline.psm1') -Force -ErrorAction SilentlyContinue } catch {}
            try { $null = Fix-GraphModuleConflicts -statusLabel $null } catch {}
            try { Connect-MgGraph -Scopes $scopes -ErrorAction Stop | Out-Null; $connected = $true } catch {}
            if (-not $connected) {
                try { Connect-MgGraph -Scopes $scopes -UseDeviceCode -ErrorAction Stop | Out-Null; $connected = $true } catch {}
            }
        } else { throw }
    }
    if (-not $connected) { throw 'Failed to connect to Microsoft Graph with delegated auth.' }
    try { $token = (Get-MgContext).AccessToken } catch { throw "Failed to obtain delegated access token" }
} else {
    Write-Info 'Acquiring Graph application token...'
    $token = Get-GraphAppToken -TenantId $cfg.TenantId -ClientId $cfg.ClientId -ClientSecret $cfg.ClientSecret
}

$headers = @{ Authorization = "Bearer $token" }

# Resolve output folder (tenant-scoped)
if ([string]::IsNullOrWhiteSpace($OutputFolder)) {
    $tenantName = 'Tenant'
    try {
        $org = Invoke-GraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/organization?$select=displayName' -Headers $headers
        if ($org.value -and $org.value[0].displayName) { $tenantName = $org.value[0].displayName }
    } catch {}
    $invalid = [System.IO.Path]::GetInvalidFileNameChars()
    $safe = ($tenantName.ToCharArray() | ForEach-Object { if ($invalid -contains $_) { '-' } else { $_ } }) -join ''
    $safe = ($safe -replace '\s+', ' ').Trim()
    if ($safe.Length -gt 80) { $safe = $safe.Substring(0,80) }
    $base = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'ExchangeOnlineAnalyzer\SecurityInvestigation'
    $tenantRoot = Join-Path $base $safe
    $ts = Get-Date -Format 'yyyyMMdd_HHmmss'
    $OutputFolder = Join-Path $tenantRoot ("InboxRules_Test_" + $ts)
}
if (-not (Test-Path $OutputFolder)) { New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null }
Write-Info ("Output: {0}" -f $OutputFolder)

# Enumerate enabled users
Write-Info 'Enumerating users...'
$users = @()
try {
    $uri = 'https://graph.microsoft.com/v1.0/users?$select=id,displayName,userPrincipalName,accountEnabled&$top=999'
    do {
        $page = Invoke-GraphRequest -Method GET -Uri $uri -Headers $headers
        if ($page.value) { $users += ($page.value | Where-Object { $_.accountEnabled -ne $false }) }
        $uri = $page.'@odata.nextLink'
    } while ($uri)
} catch {
    Write-Err ("Failed to enumerate users: {0}" -f $_.Exception.Message)
    exit 1
}
Write-Ok ("Users to query: {0}" -f $users.Count)
if ($users.Count -eq 0) { Write-Warn2 'No users found.'; exit 0 }

# Build $batch requests: GET users/{id}/mailFolders/inbox/messageRules
Write-Info 'Building batch requests...'
$reqs = New-Object System.Collections.Generic.List[hashtable]
foreach ($u in $users) {
    $reqs.Add(@{ id = $u.id; method = 'GET'; url = ("users/{0}/mailFolders/inbox/messageRules" -f $u.id) }) | Out-Null
}

# Partition into chunks of (BatchSize * ThrottleLimit) and run
Write-Info ("Fetching rules with Graph $batch (BatchSize={0}, Throttle={1})..." -f $BatchSize, $ThrottleLimit)
$responses = New-Object System.Collections.Generic.List[object]
$haveToken = -not [string]::IsNullOrWhiteSpace($token)

if ($haveToken) {
    if ($PSVersionTable.PSVersion.Major -ge 7) {
        $sliceSize = $BatchSize * $ThrottleLimit
        for ($i=0; $i -lt $reqs.Count; $i += $sliceSize) {
            $slice = $reqs[$i..([Math]::Min($i+$sliceSize-1, $reqs.Count-1))]
            $chunks = @(); for ($j=0; $j -lt $slice.Count; $j += $BatchSize) { $chunks += ,$slice[$j..([Math]::Min($j+$BatchSize-1, $slice.Count-1))] }
            $parOut = $chunks | ForEach-Object -Parallel {
                param($chunk)
                Import-Module (Join-Path (Split-Path -Parent $using:here) 'Scripts/lib/GraphAppAuth.psm1') -Force -ErrorAction Stop
                $resp = Invoke-GraphBatch -AccessToken $using:token -Requests $chunk -ChunkSize $using:BatchSize
                if ($resp) { $resp }
            } -ThrottleLimit $ThrottleLimit
            if ($parOut) {
                foreach ($r in $parOut) {
                    if ($r -is [System.Array]) { foreach ($e in $r) { if ($e) { [void]$responses.Add($e) } } }
                    elseif ($r) { [void]$responses.Add($r) }
                }
            }
        }
    } else {
        $resp = Invoke-GraphBatch -AccessToken $token -Requests $reqs -ChunkSize $BatchSize
        if ($resp) { foreach ($e in $resp) { if ($e) { [void]$responses.Add($e) } } }
    }
} else {
    Write-Warning 'Access token unavailable from context; using Graph SDK request path (sequential $batch).'
    try { Import-Module Microsoft.Graph -ErrorAction SilentlyContinue } catch {}
    $endpoint = 'https://graph.microsoft.com/v1.0/$batch'
    for ($j=0; $j -lt $reqs.Count; $j += $BatchSize) {
        $chunk = $reqs[$j..([Math]::Min($j+$BatchSize-1, $reqs.Count-1))]
        $payload = @{ requests = $chunk } | ConvertTo-Json -Depth 8
        $resp = Invoke-MgGraphRequest -Method POST -Uri $endpoint -Body $payload -ErrorAction Stop
        if ($resp.responses) { foreach ($e in $resp.responses) { if ($e) { [void]$responses.Add($e) } } }
    }
}

# Map user id -> UPN for shaping
$idToUpn = @{}; $idToDn = @{}
foreach ($u in $users) { $idToUpn[$u.id] = $u.userPrincipalName; $idToDn[$u.id] = $u.displayName }

# Shape outputs
Write-Info 'Shaping results...'
$rows = New-Object System.Collections.Generic.List[object]
foreach ($r in $responses) {
    try {
        $uid = $r.id
        $upn = $idToUpn[$uid]
        $dn  = $idToDn[$uid]
        if ($r.status -eq 200 -and $r.body -and $r.body.value) {
            foreach ($rule in $r.body.value) {
                $fromContains   = if ($rule.conditions.senderContains) { ($rule.conditions.senderContains -join ';') } else { '' }
                $subjectContains= if ($rule.conditions.subjectContains) { ($rule.conditions.subjectContains -join ';') } else { '' }
                $sentTo         = if ($rule.conditions.sentToAddresses) { ($rule.conditions.sentToAddresses -join ';') } else { '' }
                $redirectTo     = ''
                $forwardTo      = ''
                if ($rule.actions) {
                    if ($rule.actions.redirectToRecipients) { $redirectTo = ($rule.actions.redirectToRecipients -join ';') }
                    if ($rule.actions.forwardToRecipients)  { $forwardTo  = ($rule.actions.forwardToRecipients  -join ';') }
                }
                $externalTargets = $null
                if ($redirectTo -or $forwardTo) {
                    $targets = @($redirectTo -split ';', $forwardTo -split ';') | Where-Object { $_ -and $_ -match '@' }
                    if ($targets.Count -gt 0) {
                        $own = $upn.Split('@')[-1]
                        $externalTargets = ($targets | Where-Object { $_ -notlike "*${own}" }) -join ';'
                    }
                }
                $rows.Add([pscustomobject]@{
                    MailboxOwner        = $upn
                    DisplayName         = $dn
                    Name                = $rule.displayName
                    Enabled             = [bool]$rule.isEnabled
                    Priority            = $rule.sequence
                    FromAddressContains = $fromContains
                    SubjectContains     = $subjectContains
                    SentTo              = $sentTo
                    RedirectTo          = $redirectTo
                    ForwardTo           = $forwardTo
                    DeleteMessage       = [bool]$rule.actions.delete
                    StopProcessing      = [bool]$rule.stopProcessingRules
                    ExternalTargets     = $externalTargets
                }) | Out-Null
            }
        } elseif ($r.status -eq 404 -or $r.status -eq 403) {
            if ($Diag) { Write-Warn2 ("No inbox or access for {0}: {1}" -f $upn, $r.status) }
        } else {
            if ($Diag) { Write-Warn2 ("Unexpected response for {0}: status {1}" -f $upn, $r.status) }
        }
    } catch {
        if ($Diag) { Write-Warn2 ("Shape failure: {0}" -f $_.Exception.Message) }
    }
}

# Export
$csv = Join-Path $OutputFolder 'InboxRules.csv'
if ($rows.Count -gt 0) {
    $rows | Export-Csv -Path $csv -NoTypeInformation -Encoding UTF8
    Write-Ok ("Wrote {0} rules for {1} mailboxes -> {2}" -f $rows.Count, ($rows | Select-Object -ExpandProperty MailboxOwner -Unique).Count, $csv)
} else {
    'MailboxOwner,DisplayName,Name,Enabled,Priority,FromAddressContains,SubjectContains,SentTo,RedirectTo,ForwardTo,DeleteMessage,StopProcessing,ExternalTargets' | Set-Content -Path $csv -Encoding utf8
    Write-Warn2 'No rules found; wrote header only.'
}


