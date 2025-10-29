param(
    [string]$OutputFolder,
    [int]$MaxMailboxes = 0,
    [switch]$Diag,
    [switch]$Parallel,
    [int]$ThrottleLimit = 6
)

$ErrorActionPreference = 'Stop'
try { $PSStyle.OutputRendering = 'Ansi' } catch {}
function Write-Info($m){ Write-Host $m -ForegroundColor Cyan }
function Write-Ok($m){ Write-Host $m -ForegroundColor Green }
function Write-Warn2($m){ Write-Warning $m }
function Write-Err($m){ Write-Host $m -ForegroundColor Red }

# Connect Exchange Online (interactive delegated)
Write-Info 'Connecting to Exchange Online (interactive)...'
Import-Module ExchangeOnlineManagement -ErrorAction Stop
Connect-ExchangeOnline -ShowBanner:$false | Out-Null

# Tenant display name for folder scoping (best-effort via Get-OrganizationConfig)
$tenantName = 'Tenant'
try { $org = Get-OrganizationConfig -ErrorAction Stop; if ($org.DisplayName) { $tenantName = $org.DisplayName } elseif ($org.Name) { $tenantName = $org.Name } } catch {}

if ([string]::IsNullOrWhiteSpace($OutputFolder)) {
    $invalid = [System.IO.Path]::GetInvalidFileNameChars()
    $safe = ($tenantName.ToCharArray() | ForEach-Object { if ($invalid -contains $_) { '-' } else { $_ } }) -join ''
    $safe = ($safe -replace '\s+', ' ').Trim()
    if ($safe.Length -gt 80) { $safe = $safe.Substring(0,80) }
    $base = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'ExchangeOnlineAnalyzer\SecurityInvestigation'
    $tenantRoot = Join-Path $base $safe
    $ts = Get-Date -Format 'yyyyMMdd_HHmmss'
    $OutputFolder = Join-Path $tenantRoot ("InboxRules_EXO_" + $ts)
}
if (-not (Test-Path $OutputFolder)) { New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null }
Write-Info ("Output: {0}" -f $OutputFolder)

# Enumerate mailboxes
Write-Info 'Enumerating mailboxes...'
$mailboxes = @()
try { $mailboxes = Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox,SharedMailbox -ErrorAction Stop } catch { $mailboxes = Get-Mailbox -ResultSize 2000 -ErrorAction Stop }
if ($MaxMailboxes -gt 0) { $mailboxes = $mailboxes | Select-Object -First $MaxMailboxes }
Write-Ok ("Mailboxes to query: {0}" -f $mailboxes.Count)
if ($mailboxes.Count -eq 0) { Write-Warn2 'No mailboxes found.'; exit 0 }

Write-Info ($(if ($Parallel) { "Collecting inbox rules (experimental parallel mode, Throttle=$ThrottleLimit)..." } else { 'Collecting inbox rules (sequential for reliability)...' }))
$rows = New-Object System.Collections.Generic.List[object]

function _shapeRule($upn,$r){
    $extTargets = @()
    if ($r.RedirectTo) { $extTargets += $r.RedirectTo }
    if ($r.ForwardTo)  { $extTargets += $r.ForwardTo }
    $external = $null
    if ($extTargets.Count -gt 0) {
        $own = ($upn -split '@')[-1]
        $external = ($extTargets | Where-Object { $_ -and ($_ -match '@') -and ($_ -notlike "*${own}") }) -join ';'
    }
    [pscustomobject]@{
        MailboxOwner        = $upn
        Name                = $r.Name
        Enabled             = $r.Enabled
        Priority            = $r.Priority
        FromAddressContains = ($r.FromAddressContainsWords -join ';')
        SubjectContains     = ($r.SubjectContainsWords -join ';')
        SentTo              = ($r.SentTo -join ';')
        RedirectTo          = ($r.RedirectTo -join ';')
        ForwardTo           = ($r.ForwardTo -join ';')
        ForwardAsAttachment = ($r.ForwardAsAttachmentTo -join ';')
        DeleteMessage       = $r.DeleteMessage
        StopProcessing      = $r.StopProcessingRules
        ExternalTargets     = $external
    }
}

if ($Parallel -and $PSVersionTable.PSVersion.Major -ge 7) {
    try {
        $par = $mailboxes | ForEach-Object -Parallel {
            param($mbx,$diag)
            $upn = if ($mbx.UserPrincipalName) { $mbx.UserPrincipalName } else { $mbx.PrimarySmtpAddress }
            $out = @()
            try {
                $rules = Get-InboxRule -Mailbox $upn -ErrorAction Stop
                foreach ($r in $rules) { $out += (_shapeRule $upn $r) }
            } catch {
                if ($diag) { Write-Warning ("Get-InboxRule failed for {0}: {1}" -f $upn, $_.Exception.Message) }
            }
            $out
        } -ThrottleLimit $ThrottleLimit -ArgumentList $Diag
        if ($par) { foreach ($o in $par) { if ($o -is [System.Array]) { foreach ($e in $o) { if ($e) { [void]$rows.Add($e) } } } elseif ($o) { [void]$rows.Add($o) } } }
    } catch {
        Write-Warn2 ("Parallel path failed, falling back to sequential: {0}" -f $_.Exception.Message)
        $Parallel = $false
    }
}

if (-not $Parallel) {
    foreach ($mbx in $mailboxes) {
        $upn = if ($mbx.UserPrincipalName) { $mbx.UserPrincipalName } else { $mbx.PrimarySmtpAddress }
        try {
            $rules = Get-InboxRule -Mailbox $upn -ErrorAction Stop
            foreach ($r in $rules) { [void]$rows.Add((_shapeRule $upn $r)) }
        } catch {
            if ($Diag) { Write-Warn2 ("Get-InboxRule failed for {0}: {1}" -f $upn, $_.Exception.Message) }
        }
    }
}

# Export
$csv = Join-Path $OutputFolder 'InboxRules.csv'
if ($rows.Count -gt 0) {
    $rows | Export-Csv -Path $csv -NoTypeInformation -Encoding UTF8
    Write-Ok ("Wrote {0} rules for {1} mailboxes -> {2}" -f $rows.Count, ($rows | Select-Object -ExpandProperty MailboxOwner -Unique).Count, $csv)
} else {
    'MailboxOwner,Name,Enabled,Priority,FromAddressContains,SubjectContains,SentTo,RedirectTo,ForwardTo,ForwardAsAttachment,DeleteMessage,StopProcessing,ExternalTargets' | Set-Content -Path $csv -Encoding utf8
    Write-Warn2 'No rules found; wrote header only.'
}

# Disconnect EXO session
try { Disconnect-ExchangeOnline -Confirm:$false } catch {}


