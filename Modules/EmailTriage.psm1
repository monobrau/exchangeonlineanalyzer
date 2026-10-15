#Requires -Version 5.1
# Email security triage. Collectors use Get-* / Resolve-DnsName / ORCA / DomainHealthChecker.
# actions.ps1 stays commented. The Actions tab applies a fixed set of Exchange cmdlets, or copies the script.

Set-StrictMode -Version 2.0

$script:TriageFreemail = @(
    'gmail.com', 'googlemail.com', 'yahoo.com', 'ymail.com', 'rocketmail.com',
    'hotmail.com', 'outlook.com', 'live.com', 'msn.com', 'aol.com',
    'icloud.com', 'me.com', 'mac.com', 'proton.me', 'protonmail.com', 'pm.me',
    'zoho.com', 'mail.com', 'gmx.com', 'gmx.net', 'yandex.com', 'yandex.ru',
    'fastmail.com', 'tutanota.com', 'tuta.com', 'hushmail.com',
    'comcast.net', 'verizon.net', 'att.net', 'sbcglobal.net', 'cox.net',
    'charter.net', 'spectrum.net', 'bellsouth.net', 'earthlink.net',
    'optonline.net', 'frontier.com', 'centurylink.net', 'qq.com', '163.com'
)

$script:TriageThreatCat = @{
    'PHSH' = $true; 'HPHSH' = $true; 'HPHISH' = $true; 'MALW' = $true; 'SPOOF' = $true
    'UIMP' = $true; 'DIMP' = $true; 'BIMP' = $true; 'GIMP' = $true; 'HSPM' = $true
    'SPM' = $true; 'INTOS' = $true; 'AMP' = $true; 'FTBP' = $true; 'SAP' = $true
}

$script:TriageMemberberryDoc = @{}

function Get-TriageProp {
    param($Obj, [string[]]$Names)
    if ($null -eq $Obj) { return $null }
    foreach ($n in $Names) {
        $value = $null
        $found = $false
        if ($Obj -is [System.Collections.IDictionary]) {
            if ($Obj.Contains($n)) { $value = $Obj[$n]; $found = $true }
        }
        else {
            $p = $Obj.PSObject.Properties[$n]
            if ($p) { $value = $p.Value; $found = $true }
        }
        if (-not $found -or $null -eq $value) { continue }
        # [string] on List[object] is safe; @() on List[object] is not. Return collections as-is.
        if ($value -is [string]) {
            if (-not [string]::IsNullOrWhiteSpace($value)) { return $value }
            continue
        }
        if ($value -is [System.Collections.IEnumerable]) {
            foreach ($item in $value) {
                if ($null -ne $item -and -not [string]::IsNullOrWhiteSpace([string]$item)) { return $value }
            }
            continue
        }
        if (-not [string]::IsNullOrWhiteSpace([string]$value)) { return $value }
    }
    return $null
}

function ConvertTo-TriageArray {
    param($Value)
    # @() throws "Argument types do not match" on List[object] in Windows PowerShell 5.1.
    if ($null -eq $Value) { return @() }
    if ($Value -is [System.Collections.IEnumerable] -and $Value -isnot [string]) {
        $copy = New-Object System.Collections.Generic.List[object]
        foreach ($item in $Value) { [void]$copy.Add($item) }
        if ($copy.Count -eq 0) { return @() }
        return $copy.ToArray()
    }
    return ,$Value
}

function Get-TriageDomain {
    param([string]$Address)
    if ([string]::IsNullOrWhiteSpace($Address)) { return '' }
    $a = $Address.Trim().ToLowerInvariant()
    $at = $a.LastIndexOf('@')
    if ($at -ge 0 -and $at -lt ($a.Length - 1)) { return $a.Substring($at + 1) }
    return $a.TrimStart('.')
}

function Get-TriageFreemailDomains {
    param([string[]]$Extra = @())
    $set = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($d in $script:TriageFreemail) { [void]$set.Add($d) }
    foreach ($d in @($Extra)) {
        $n = Get-TriageDomain $d
        if ($n) { [void]$set.Add($n) }
    }
    return @($set | Sort-Object)
}

function Get-TriageLayer {
    param(
        [string]$Type = '',
        [string]$PolicyType = '',
        [string]$PolicyName = ''
    )
    $blob = ('{0} {1} {2}' -f $Type, $PolicyType, $PolicyName)
    $vendor = $blob -match '(?i)inky|proofpoint|barracuda'
    $rule = ($Type -match '(?i)TransportRule') -or ($PolicyType -match '(?i)TransportRule|Transport')
    if ($vendor -and $rule) { return 'third-party filter' }
    if ($rule) { return 'transport rule' }
    if ($PolicyType -match '(?i)HostedContentFilter') { return 'HCP' }
    if (($Type -match '(?i)Phish|Malware') -or ($PolicyType -match '(?i)AntiPhish|Anti-Phish')) { return 'anti-phish' }
    if ($Type -match '(?i)Spam|Bulk') { return 'EOP anti-spam' }
    return 'Unknown'
}

function Get-TriageHeaderText {
    param($Result)
    if ($null -eq $Result) { return '' }
    $text = ''
    if ($Result -is [string]) { $text = $Result }
    else {
        $fromProp = Get-TriageProp $Result @('Header')
        if ($null -ne $fromProp) { $text = [string]$fromProp }
    }
    $text = $text.Trim()
    if (-not $text) { return '' }
    # Out-String on Get-QuarantineMessageHeader wraps the header column and cuts it off.
    if ($text -match '(?m)^\s*Identity\s+Organization\s+Header\s*$') { return '' }
    return $text
}

function ConvertFrom-TriageMessageHeader {
    param([string]$Header = '')
    $empty = [pscustomobject]@{
        SCL = $null; SFV = $null; CAT = $null; CIP = $null; CTRY = $null; IPV = $null; SFTY = $null
        Spf = $null; Dkim = $null; Dmarc = $null; CompAuth = $null; CompAuthReason = $null; Bcl = $null
    }
    if ([string]::IsNullOrWhiteSpace($Header)) { return $empty }
    $text = $Header -replace "`r`n", "`n" -replace "`r", "`n"
    $text = [regex]::Replace($text, "`n[ \t]+", ' ')

    $ff = ''
    $ms = ''
    $authParts = New-Object System.Collections.Generic.List[string]
    foreach ($line in ($text -split "`n")) {
        if ($line -match '^(?i)X-Forefront-Antispam-Report:\s*(.*)$') { $ff = $Matches[1] }
        elseif ($line -match '^(?i)X-Microsoft-Antispam:\s*(.*)$') { $ms = $Matches[1] }
        elseif ($line -match '^(?i)Authentication-Results:\s*(.*)$') { $authParts.Add($Matches[1]) }
    }
    $auth = ($authParts -join ' ')

    $fields = @{}
    foreach ($piece in ($ff -split ';')) {
        $piece = $piece.Trim()
        if (-not $piece) { continue }
        $colon = $piece.IndexOf(':')
        if ($colon -lt 1) { continue }
        $k = $piece.Substring(0, $colon).Trim().ToUpperInvariant()
        $v = $piece.Substring($colon + 1).Trim()
        if (-not $fields.ContainsKey($k)) { $fields[$k] = $v }
    }
    function Take-Auth([string]$Blob, [string]$Name) {
        $re = '(?i)\b' + [regex]::Escape($Name) + '=([a-z]+)'
        $m = [regex]::Match($Blob, $re)
        if ($m.Success) { return $m.Groups[1].Value.ToLowerInvariant() }
        return $null
    }
    $reason = $null
    $rm = [regex]::Match($auth, '(?i)\breason=(\d{3})\b')
    if ($rm.Success) { $reason = $rm.Groups[1].Value }
    $bcl = $null
    $bm = [regex]::Match($ms, '(?i)\bBCL[:=](\d+)')
    if ($bm.Success) { $bcl = $bm.Groups[1].Value }

    return [pscustomobject]@{
        SCL             = $(if ($fields.ContainsKey('SCL')) { $fields['SCL'] } else { $null })
        SFV             = $(if ($fields.ContainsKey('SFV')) { $fields['SFV'] } else { $null })
        CAT             = $(if ($fields.ContainsKey('CAT')) { $fields['CAT'] } else { $null })
        CIP             = $(if ($fields.ContainsKey('CIP')) { $fields['CIP'] } else { $null })
        CTRY            = $(if ($fields.ContainsKey('CTRY')) { $fields['CTRY'] } else { $null })
        IPV             = $(if ($fields.ContainsKey('IPV')) { $fields['IPV'] } else { $null })
        SFTY            = $(if ($fields.ContainsKey('SFTY')) { $fields['SFTY'] } else { $null })
        Spf             = (Take-Auth $auth 'spf')
        Dkim            = (Take-Auth $auth 'dkim')
        Dmarc           = (Take-Auth $auth 'dmarc')
        CompAuth        = (Take-Auth $auth 'compauth')
        CompAuthReason  = $reason
        Bcl             = $bcl
    }
}

function Get-TriageImpersonationMatch {
    param(
        [string]$Subject = '',
        [string]$SenderAddress = '',
        [string[]]$VipNames = @(),
        [string[]]$InternalDomains = @()
    )
    $sub = $Subject.ToLowerInvariant()
    $hit = $null
    foreach ($name in @($VipNames)) {
        $n = ([string]$name).Trim()
        if ($n.Length -lt 3) { continue }
        if ($sub.Contains($n.ToLowerInvariant())) { $hit = $n; break }
    }
    if (-not $hit) { return $null }
    $dom = Get-TriageDomain $SenderAddress
    $internal = $false
    foreach ($d in @($InternalDomains)) {
        if ($dom -and $dom -eq (Get-TriageDomain $d)) { $internal = $true; break }
    }
    $cat = 'External'
    $evidence = "Exec name '$hit' in subject from an external sender"
    if ($internal) {
        $cat = 'SpoofedInternal'
        $evidence = "Exec name '$hit' in subject from an internal-domain sender"
    }
    return [pscustomobject]@{ Category = $cat; Name = $hit; Evidence = $evidence }
}

function Get-TriageClassification {
    param(
        [string]$Subject = '',
        [string]$SenderAddress = '',
        [string[]]$VipNames = @(),
        [string[]]$InternalDomains = @(),
        [string[]]$FreemailDomains = @(),
        [int]$SenderRecipientCount = 1,
        [int]$SprayThreshold = 5,
        [int]$RotationDomainCount = 1,
        [string]$RotationLocalPart = '',
        $ParsedHeader = $null
    )
    if (-not $FreemailDomains -or @($FreemailDomains).Count -eq 0) {
        $FreemailDomains = @(Get-TriageFreemailDomains)
    }
    $evidence = New-Object System.Collections.Generic.List[string]
    $malicious = $false
    $legit = $false
    $noise = $false
    $dom = Get-TriageDomain $SenderAddress
    $freemail = $false
    foreach ($d in @($FreemailDomains)) {
        if ($dom -and $dom -eq $d.ToLowerInvariant()) { $freemail = $true; break }
    }
    $imp = Get-TriageImpersonationMatch -Subject $Subject -SenderAddress $SenderAddress -VipNames $VipNames -InternalDomains $InternalDomains
    if ($imp -and ($freemail -or $imp.Category -eq 'SpoofedInternal')) {
        $malicious = $true
        $evidence.Add([string]$imp.Evidence)
    }
    if ($Subject -match '(?i)password expires|voicemail|docusign|\bssa\b|review changes') {
        $malicious = $true
        $evidence.Add('Credential-lure wording in the subject')
    }
    if ($Subject -match '(?i)accounts payable|accounts-payable|wire transfer|bank details|past due|remittance|\bach\b|updated banking') {
        $malicious = $true
        $evidence.Add('Invoice or payment-fraud wording in the subject')
    }
    if ($RotationDomainCount -ge 3 -and $RotationLocalPart) {
        $malicious = $true
        $evidence.Add("Domain rotation: local part '$RotationLocalPart' across $RotationDomainCount domains")
    }
    if ($ParsedHeader) {
        $cat = [string](Get-TriageProp $ParsedHeader @('CAT'))
        if ($cat) {
            $evidence.Add("Forefront CAT: $cat")
            if ($script:TriageThreatCat.ContainsKey($cat.ToUpperInvariant())) { $malicious = $true }
        }
        $sfv = [string](Get-TriageProp $ParsedHeader @('SFV'))
        if ($sfv) {
            $evidence.Add("Forefront SFV: $sfv")
            if ($sfv -match '^(?i)(SPM|SKB|BLK|SKS)$') { $malicious = $true }
        }
        $sfty = [string](Get-TriageProp $ParsedHeader @('SFTY'))
        if ($sfty) { $malicious = $true; $evidence.Add("Forefront SFTY: $sfty") }
        foreach ($pair in @(
                @('SCL', 'SCL'), @('CIP', 'Connecting IP'), @('CTRY', 'Source country'), @('IPV', 'IP reputation'), @('Bcl', 'BCL')
            )) {
            $val = Get-TriageProp $ParsedHeader @($pair[0])
            if ($val) { $evidence.Add(('{0}: {1}' -f $pair[1], $val)) }
        }
        $sclNum = 0
        $sclRaw = Get-TriageProp $ParsedHeader @('SCL')
        if ($sclRaw -and [int]::TryParse([string]$sclRaw, [ref]$sclNum) -and $sclNum -ge 5) { $malicious = $true }
        $bclNum = 0
        $bclRaw = Get-TriageProp $ParsedHeader @('Bcl')
        if ($bclRaw -and [int]::TryParse([string]$bclRaw, [ref]$bclNum) -and $bclNum -ge 7) { $malicious = $true }
        foreach ($authName in @('Spf', 'Dkim', 'Dmarc', 'CompAuth')) {
            $av = Get-TriageProp $ParsedHeader @($authName)
            if (-not $av) { continue }
            $label = $authName.ToUpperInvariant()
            if ($authName -eq 'CompAuth') { $label = 'compauth' }
            $reason = Get-TriageProp $ParsedHeader @('CompAuthReason')
            if ($authName -eq 'CompAuth' -and $reason) { $evidence.Add("$label=$av reason $reason") }
            else { $evidence.Add("$label=$av") }
            if ([string]$av -eq 'fail') { $malicious = $true }
        }
    }
    $nonFreemail = $dom -and -not $freemail
    $reply = ($Subject -match '(?i)^\s*re:\s+\S+' -and $nonFreemail)
    if ($SenderRecipientCount -ge $SprayThreshold -and $SprayThreshold -gt 1 -and -not $reply) {
        $evidence.Add("Spray: one sender to $SenderRecipientCount recipients")
        $malicious = $true
    }
    if ($reply -and -not $malicious) {
        $legit = $true
        $evidence.Add('Reply subject from a non-freemail sender')
    }
    if ($Subject -match '(?i)merchant cash|\bmca\b|business loan|giveaway|cold sales|limited time offer|partnership opportunity') {
        $noise = $true
        $evidence.Add('Cold-sales or noise wording in the subject')
    }
    $verdict = 'Unknown'
    if ($malicious) { $verdict = 'Malicious' }
    elseif ($legit) { $verdict = 'Legit' }
    elseif ($noise) { $verdict = 'Noise' }
    return [pscustomobject]@{ Verdict = $verdict; Evidence = @($evidence.ToArray()) }
}

function Merge-TriageTraceLegs {
    param(
        $Rows,
        [int]$WindowSeconds = 120
    )
    $fold = @{ 'expanded' = $true; 'resolved' = $true }
    $filter = @{ 'quarantined' = $true; 'filteredasspam' = $true }
    $groups = @{}
    $i = 0
    foreach ($row in (ConvertTo-TriageArray $Rows)) {
        if (-not $row) { continue }
        $i++
        $mid = [string](Get-TriageProp $row @('MessageId', 'MessageID'))
        if (-not $mid) { $mid = "row-$i" }
        $key = $mid.ToLowerInvariant()
        if (-not $groups.ContainsKey($key)) { $groups[$key] = New-Object System.Collections.Generic.List[object] }
        $received = Get-TriageProp $row @('Received', 'ReceivedTime')
        $dt = $null
        if ($received) { $dt = [datetime]$received }
        $groups[$key].Add([pscustomobject]@{
                MessageId = $mid
                Recipient = [string](Get-TriageProp $row @('RecipientAddress', 'Recipient'))
                Sender    = [string](Get-TriageProp $row @('SenderAddress', 'Sender'))
                Subject   = [string](Get-TriageProp $row @('Subject'))
                Status    = [string](Get-TriageProp $row @('Status'))
                Received  = $dt
                TraceId   = [string](Get-TriageProp $row @('MessageTraceId'))
            })
    }
    $out = New-Object System.Collections.Generic.List[object]
    foreach ($key in $groups.Keys) {
        $items = @($groups[$key] | Sort-Object Received)
        $terminal = @($items | Where-Object { -not $fold.ContainsKey($_.Status.ToLowerInvariant()) })
        $folded = @($items | Where-Object { $fold.ContainsKey($_.Status.ToLowerInvariant()) })
        if ($terminal.Count -eq 0) {
            foreach ($f in $folded) {
                $out.Add([pscustomobject]@{
                        MessageId = $f.MessageId; RecipientAddress = $f.Recipient; SenderAddress = $f.Sender
                        Subject = $f.Subject; Received = $f.Received; Statuses = @($f.Status)
                        Outcome = 'Folded'; ReachedMailbox = $false; FinalStatus = $f.Status; MessageTraceId = $f.TraceId
                    })
            }
            continue
        }
        $byRcpt = @{}
        foreach ($t in $terminal) {
            $rk = $t.Recipient.ToLowerInvariant()
            if (-not $byRcpt.ContainsKey($rk)) { $byRcpt[$rk] = New-Object System.Collections.Generic.List[object] }
            $byRcpt[$rk].Add($t)
        }
        foreach ($rk in $byRcpt.Keys) {
            $ev = @($byRcpt[$rk] | Sort-Object Received)
            $reached = $false
            $deliveredAt = $null
            foreach ($e in $ev) {
                $st = $e.Status.ToLowerInvariant()
                if ($st -eq 'delivered') {
                    $deliveredAt = $e.Received
                    $reached = $true
                }
                elseif ($filter.ContainsKey($st) -and $null -ne $deliveredAt -and $null -ne $e.Received) {
                    $delta = ($e.Received - $deliveredAt).TotalSeconds
                    if ($delta -ge 0 -and $delta -le $WindowSeconds) { $reached = $false }
                }
            }
            $statuses = New-Object System.Collections.Generic.List[string]
            foreach ($f in $folded) { $statuses.Add($f.Status) }
            foreach ($e in $ev) { $statuses.Add($e.Status) }
            $last = $ev[-1]
            $outcome = 'Other'
            if ($reached) { $outcome = 'ReachedMailbox' }
            elseif (@($ev | Where-Object { $filter.ContainsKey($_.Status.ToLowerInvariant()) }).Count -gt 0) { $outcome = 'Filtered' }
            $out.Add([pscustomobject]@{
                    MessageId = $last.MessageId; RecipientAddress = $last.Recipient; SenderAddress = $last.Sender
                    Subject = $last.Subject; Received = $ev[0].Received; Statuses = @($statuses.ToArray())
                    Outcome = $outcome; ReachedMailbox = $reached; FinalStatus = $last.Status; MessageTraceId = $last.TraceId
                })
        }
    }
    return @($out.ToArray())
}

function ConvertTo-TriageUtc {
    param(
        $Time,
        [TimeZoneInfo]$TimeZone = $null
    )
    if ($null -eq $Time -or [string]::IsNullOrWhiteSpace([string]$Time)) { return $null }
    $dt = [datetime]$Time
    if ($dt.Kind -eq [DateTimeKind]::Utc) { return $dt }
    if (-not $TimeZone) {
        if ($dt.Kind -eq [DateTimeKind]::Local) { return $dt.ToUniversalTime() }
        $TimeZone = [TimeZoneInfo]::Local
    }
    $unspec = [datetime]::SpecifyKind($dt, [DateTimeKind]::Unspecified)
    return [datetime]::SpecifyKind([TimeZoneInfo]::ConvertTimeToUtc($unspec, $TimeZone), [DateTimeKind]::Utc)
}

function Get-TriageLocalZoneLabel {
    param([datetime]$Time = (Get-Date))
    try {
        $tz = [TimeZoneInfo]::Local
        $name = if ($tz.IsDaylightSavingTime($Time)) { $tz.DaylightName } else { $tz.StandardName }
        if ($name) { return $name }
        if ($tz.Id) { return $tz.Id }
    } catch {}
    return 'local time'
}

function ConvertTo-TriagePresentedTime {
    param([string]$Value)
    $trimmed = ([string]$Value).Trim()
    if (-not $trimmed) { return '' }
    $dto = [datetimeoffset]::MinValue
    $styles = [System.Globalization.DateTimeStyles]::RoundtripKind
    $culture = [System.Globalization.CultureInfo]::InvariantCulture
    if ($trimmed -match '(?i)(?:Z|[+-]\d{2}:?\d{2})$') {
        if ([datetimeoffset]::TryParse($trimmed, $culture, $styles, [ref]$dto)) {
            return $dto.ToLocalTime().ToString('yyyy-MM-dd h:mm tt')
        }
    }
    $dt = [datetime]::MinValue
    if ($trimmed -match '^\d{4}-\d{2}-\d{2}[T ]\d{2}:\d{2}') {
        $iso = $trimmed
        if ($iso -notmatch '(?i)Z$') { $iso = (($iso -replace ' ', 'T') + 'Z') }
        if ([datetimeoffset]::TryParse($iso, $culture, $styles, [ref]$dto)) {
            return $dto.ToLocalTime().ToString('yyyy-MM-dd h:mm tt')
        }
    }
    if ([datetime]::TryParse($trimmed, $culture, [System.Globalization.DateTimeStyles]::AssumeUniversal, [ref]$dt)) {
        return $dt.ToLocalTime().ToString('yyyy-MM-dd h:mm tt')
    }
    return $trimmed
}

function Get-TriageDnsHost {
    param(
        [string[]]$Ns = @(),
        [string]$Soa = ''
    )
    $blob = ((@($Ns) + @($Soa)) -join ' ').ToLowerInvariant()
    if ($blob -match 'cloudflare') { return 'Cloudflare' }
    if ($blob -match 'awsdns') { return 'Amazon Route 53' }
    if ($blob -match 'domaincontrol') { return 'GoDaddy' }
    if ($blob -match 'azure-dns') { return 'Azure DNS' }
    if ($blob -match 'googledomains|google\.com') { return 'Google' }
    if ($blob -match 'registrar-servers') { return 'Namecheap' }
    $hostName = ''
    if (@($Ns).Count -gt 0) { $hostName = [string]$Ns[0] }
    elseif ($Soa) { $hostName = $Soa }
    $hostName = $hostName.Trim().TrimEnd('.').ToLowerInvariant()
    if (-not $hostName) { return '' }
    $parts = $hostName.Split('.')
    if ($parts.Count -ge 2) { return ($parts[-2] + '.' + $parts[-1]) }
    return $hostName
}

function Test-TriageMxPointsAtEop {
    param([string[]]$Exchanges = @())
    foreach ($ex in @($Exchanges)) {
        if ([string]$ex -match '(?i)(^|\.)mail\.protection\.outlook\.com$') { return $true }
    }
    return $false
}

function Test-TriageWholeFreemailDomain {
    param([string]$Value, [string[]]$FreemailDomains)
    $v = ([string]$Value).Trim().ToLowerInvariant()
    if (-not $v) { return $false }
    if ($v -match '^\*@(.+)$') { $v = $Matches[1] }
    if ($v.Contains('@')) { return $false }
    foreach ($d in @($FreemailDomains)) {
        if ($v -eq $d.ToLowerInvariant()) { return $true }
    }
    return $false
}

function Test-TriageConfigFindings {
    param(
        $Config,
        [string[]]$FreemailDomains = @()
    )
    if (-not $FreemailDomains -or @($FreemailDomains).Count -eq 0) { $FreemailDomains = @(Get-TriageFreemailDomains) }
    $findings = New-Object System.Collections.Generic.List[object]
    function Add-Finding([string]$Check, [string]$Severity, [string]$Detail) {
        $findings.Add([pscustomobject]@{ Check = $Check; Severity = $Severity; Detail = $Detail })
    }
    foreach ($rule in @($(if ($Config) { Get-TriageProp $Config @('TransportRules') }))) {
        if (-not $rule) { continue }
        $name = [string](Get-TriageProp $rule @('Name', 'Identity'))
        $vals = @()
        $raw = Get-TriageProp $rule @('ExceptIfSenderDomainIs', 'Exceptions')
        if ($raw) { $vals = ConvertTo-TriageArray $raw }
        foreach ($v in $vals) {
            if (Test-TriageWholeFreemailDomain -Value ([string]$v) -FreemailDomains $FreemailDomains) {
                Add-Finding 'TransportRuleConsumerException' 'High' "Transport rule '$name' exempts the whole domain $v."
            }
        }
    }
    foreach ($item in @($(if ($Config) { Get-TriageProp $Config @('TablItems') }))) {
        if (-not $item) { continue }
        $action = [string](Get-TriageProp $item @('Action'))
        if ($action -notmatch '(?i)^Allow$') { continue }
        $value = [string](Get-TriageProp $item @('Value'))
        if (Test-TriageWholeFreemailDomain -Value $value -FreemailDomains $FreemailDomains) {
            Add-Finding 'TablFreemailAllow' 'High' "Tenant allow/block list allows the freemail or ISP domain $value."
        }
    }
    $mxPoints = $false
    $mxSeen = $false
    foreach ($mx in @($(if ($Config) { Get-TriageProp $Config @('Mx') }))) {
        if (-not $mx) { continue }
        $mxSeen = $true
        $ex = @()
        $rawEx = Get-TriageProp $mx @('Exchanges')
        if ($rawEx) { $ex = ConvertTo-TriageArray $rawEx }
        if (Test-TriageMxPointsAtEop -Exchanges $ex) { $mxPoints = $true }
    }
    $mxAway = $mxSeen -and -not $mxPoints
    if ($mxAway) {
        foreach ($conn in @($(if ($Config) { Get-TriageProp $Config @('InboundConnectors') }))) {
            if (-not $conn) { continue }
            $cname = [string](Get-TriageProp $conn @('Name', 'Identity'))
            if ($cname -notmatch '(?i)inky|proofpoint|barracuda') { continue }
            $ips = @()
            $rawIps = Get-TriageProp $conn @('EFSkipIPs')
            if ($rawIps) { $ips = @($rawIps | Where-Object { $_ }) }
            $skipLast = $false
            $skipValue = $null
            if ($conn -is [System.Collections.IDictionary] -and $conn.Contains('EFSkipLastIP')) { $skipValue = $conn['EFSkipLastIP'] }
            else {
                $skipProp = $conn.PSObject.Properties['EFSkipLastIP']
                if ($skipProp) { $skipValue = $skipProp.Value }
            }
            if ($skipValue -eq $true) { $skipLast = $true }
            if ($ips.Count -eq 0 -and -not $skipLast) {
                Add-Finding 'InboundConnectorEnhancedFiltering' 'High' "Inbound connector '$cname' has no EFSkipIPs or Enhanced Filtering, and MX does not point at Exchange Online Protection."
            }
        }
    }
    $policies = @($(if ($Config) { Get-TriageProp $Config @('AntiPhishPolicies') }))
    $sawPolicy = $false
    $anyUsers = $false
    foreach ($pol in $policies) {
        if (-not $pol) { continue }
        $sawPolicy = $true
        $users = @()
        $rawUsers = Get-TriageProp $pol @('TargetedUsersToProtect')
        if ($rawUsers) { $users = @($rawUsers | Where-Object { $_ }) }
        if ($users.Count -gt 0) { $anyUsers = $true }
    }
    if ($sawPolicy -and -not $anyUsers) {
        Add-Finding 'AntiPhishNoProtectedUsers' 'Medium' 'Anti-phish policies do not list any protected users. ORCA covers the protection toggle; this check is only the empty list.'
    }
    foreach ($health in @($(if ($Config) { Get-TriageProp $Config @('DomainHealth') }))) {
        if (-not $health) { continue }
        $name = [string](Get-TriageProp $health @('Name'))
        $rec = [string](Get-TriageProp $health @('DmarcRecord'))
        if ([string]::IsNullOrWhiteSpace($rec) -or $rec -match '(?i)no dmarc') {
            Add-Finding 'DmarcMissing' 'High' "DomainHealthChecker found no DMARC record for $name."
            continue
        }
        if ($rec -match '(?i)(^|;)\s*p=none(\s|;|$)') {
            Add-Finding 'DmarcPolicyNone' 'High' "DMARC for $name is p=none ($rec)."
        }
        if ($rec -notmatch '(?i)\brua=') {
            Add-Finding 'DmarcMissingRua' 'Medium' "DMARC for $name has no rua tag ($rec)."
        }
    }
    foreach ($dns in @($(if ($Config) { Get-TriageProp $Config @('Dns') }))) {
        if (-not $dns) { continue }
        $ns = @()
        $rawNs = Get-TriageProp $dns @('Ns')
        if ($rawNs) { $ns = ConvertTo-TriageArray $rawNs }
        $soa = [string](Get-TriageProp $dns @('Soa'))
        $hostLabel = [string](Get-TriageProp $dns @('Host'))
        if (-not $hostLabel) { $hostLabel = Get-TriageDnsHost -Ns $ns -Soa $soa }
        $domain = [string](Get-TriageProp $dns @('Domain'))
        if ($hostLabel) { Add-Finding 'DnsHost' 'Info' "DNS host for $domain is $hostLabel." }
    }
    return @($findings.ToArray())
}

function Get-TriageOrcaSummary {
    param($Orca)
    $results = @()
    if ($Orca) {
        $raw = Get-TriageProp $Orca @('Results')
        if ($raw) { $results = ConvertTo-TriageArray $raw }
    }
    $mdo = $false
    $fail = 0
    $pass = 0
    $info = 0
    $buckets = @{}
    function Map-Result([string]$Value) {
        switch ($Value) {
            '1' { return 'Pass' }
            '2' { return 'Informational' }
            '3' { return 'Fail' }
            'Pass' { return 'Pass' }
            'Informational' { return 'Informational' }
            'Fail' { return 'Fail' }
            default { return $Value }
        }
    }
    function Map-Chi([string]$Value) {
        switch ($Value) {
            '5' { return 'Low' }
            '10' { return 'Medium' }
            '15' { return 'High' }
            '20' { return 'VeryHigh' }
            '100' { return 'Critical' }
            '0' { return 'NotRated' }
            default { return $(if ($Value) { $Value } else { 'NotRated' }) }
        }
    }
    foreach ($check in $results) {
        if (-not $check) { continue }
        $svc = [string](Get-TriageProp $check @('Services'))
        if ($svc -eq '2' -or $svc -eq 'MDO') { $mdo = $true }
        $result = Map-Result ([string](Get-TriageProp $check @('Result')))
        if ($result -eq 'Fail') { $fail++ }
        elseif ($result -eq 'Pass') { $pass++ }
        elseif ($result -eq 'Informational') { $info++ }
        if ($result -ne 'Fail' -and $result -ne 'Informational') { continue }
        $sev = Map-Chi ([string](Get-TriageProp $check @('ChiValue')))
        if (-not $buckets.ContainsKey($sev)) { $buckets[$sev] = New-Object System.Collections.Generic.List[object] }
        $buckets[$sev].Add([pscustomobject]@{
                Name               = [string](Get-TriageProp $check @('Name'))
                Area               = [string](Get-TriageProp $check @('Area'))
                Result             = $result
                Severity           = $sev
                FailRecommendation = [string](Get-TriageProp $check @('FailRecommendation'))
                Services           = $svc
            })
    }
    $order = @('Critical', 'VeryHigh', 'High', 'Medium', 'Low', 'NotRated')
    $groups = New-Object System.Collections.Generic.List[object]
    foreach ($sev in $order) {
        if ($buckets.ContainsKey($sev)) {
            $groups.Add([pscustomobject]@{ Severity = $sev; Items = @($buckets[$sev].ToArray()) })
        }
    }
    $label = 'Exchange Online Protection (Defender for Office 365 was not detected)'
    if ($mdo) { $label = 'Defender for Office 365' }
    return [pscustomobject]@{
        DefenderForOffice365 = $mdo
        ServiceLabel         = $label
        Groups               = @($groups.ToArray())
        FailCount            = $fail
        PassCount            = $pass
        InformationalCount   = $info
    }
}

function Get-TriageBlockTarget {
    param(
        [string]$SenderAddress,
        [string[]]$FreemailDomains = @()
    )
    if (-not $FreemailDomains -or @($FreemailDomains).Count -eq 0) { $FreemailDomains = @(Get-TriageFreemailDomains) }
    $dom = Get-TriageDomain $SenderAddress
    $free = $false
    foreach ($d in @($FreemailDomains)) { if ($dom -eq $d.ToLowerInvariant()) { $free = $true } }
    if ($free) { return [pscustomobject]@{ Value = $SenderAddress.Trim(); Scope = 'Address' } }
    return [pscustomobject]@{ Value = $dom; Scope = 'Domain' }
}

function New-TriageActionScript {
    param(
        [string[]]$ReleaseIdentities = @(),
        $Blocks = @(),
        [string[]]$ProtectUsers = @(),
        [string[]]$SoftDeleteMessageIds = @()
    )
    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add('# Email Security Triage actions')
    $lines.Add('# WARNING: Every command in this file is commented out. Review each line before removing the leading #.')
    $lines.Add('# The web UI does not run these commands.')
    $lines.Add('# Soft-delete of delivered mail needs ExchangeOnlineManagement 3.9 or later.')
    $lines.Add('# Explorer alternative: Microsoft Defender portal > Explorer > search the message id > Soft delete.')
    $lines.Add('')
    foreach ($id in @($ReleaseIdentities)) {
        if ($id) { $lines.Add("# Release-QuarantineMessage -Identity '$id'") }
    }
    foreach ($b in @($Blocks)) {
        if (-not $b) { continue }
        $val = [string](Get-TriageProp $b @('Value'))
        if (-not $val) { continue }
        $lines.Add("# New-TenantAllowBlockListItems -ListType Sender -Block -Entries '$val' -NoExpiration")
    }
    foreach ($u in @($ProtectUsers)) {
        if (-not $u) { continue }
        $lines.Add("# Set-AntiPhishPolicy -Identity 'Office365 AntiPhish Default' -EnableTargetedUserProtection `$true -TargetedUsersToProtect @{Add='$u'}")
    }
    if (@($SoftDeleteMessageIds).Count -gt 0) {
        $lines.Add('# Connect-IPPSSession -EnableSearchOnlySession')
        foreach ($mid in @($SoftDeleteMessageIds)) {
            if (-not $mid) { continue }
            $safe = ($mid -replace "'", '')
            $lines.Add("# New-ComplianceSearch -Name 'TriageSoftDelete' -ExchangeLocation All -ContentMatchQuery 'MessageId:$safe'")
        }
        $lines.Add("# Start-ComplianceSearch -Identity 'TriageSoftDelete'")
        $lines.Add("# New-ComplianceSearchAction -SearchName 'TriageSoftDelete' -Purge -PurgeType SoftDelete")
    }
    return ($lines -join "`r`n")
}

function New-TriageTicketNote {
    param($Rows, $Findings, [string]$Company = 'the tenant')
    $rows = @($Rows)
    $counts = @{ Malicious = 0; Legit = 0; Noise = 0; Unknown = 0 }
    $reached = 0
    foreach ($r in $rows) {
        $v = [string](Get-TriageProp $r @('Verdict'))
        if ($counts.ContainsKey($v)) { $counts[$v]++ }
        $flag = Get-TriageProp $r @('ReachedMailbox')
        if ($flag -eq $true) { $reached++ }
    }
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("Email security triage - $Company")
    [void]$sb.AppendLine("Malicious: $($counts.Malicious)")
    [void]$sb.AppendLine("Legit: $($counts.Legit)")
    [void]$sb.AppendLine("Noise: $($counts.Noise)")
    [void]$sb.AppendLine("Unknown: $($counts.Unknown)")
    [void]$sb.AppendLine("Reached a mailbox: $reached")
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('Findings:')
    foreach ($f in @($Findings)) {
        if (-not $f) { continue }
        [void]$sb.AppendLine(('- [{0}] {1}' -f (Get-TriageProp $f @('Severity')), (Get-TriageProp $f @('Detail'))))
    }
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('Messages:')
    foreach ($r in $rows) {
        if (-not $r) { continue }
        $ev = Get-TriageProp $r @('Evidence')
        $evText = ''
        if ($ev) { $evText = (@($ev) -join '; ') }
        [void]$sb.AppendLine(('- {0} | {1} | {2} | {3}' -f (Get-TriageProp $r @('Verdict')), (Get-TriageProp $r @('SenderAddress')), (Get-TriageProp $r @('Subject')), $evText))
    }
    return $sb.ToString().TrimEnd()
}

function New-TriageClientEmail {
    param($Rows, [string]$Company = 'your organization', [string[]]$Contacts = @())
    $mal = @($Rows | Where-Object { $_ -and [string](Get-TriageProp $_ @('Verdict')) -eq 'Malicious' })
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('Hello,')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine("We reviewed recent email that Microsoft held or that needed a closer look for $Company.")
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('What we found')
    if ($mal.Count -eq 0) {
        [void]$sb.AppendLine('- We did not see a clear malicious message in this pass. We can still release anything your team recognizes.')
    }
    else {
        [void]$sb.AppendLine("- $($mal.Count) message(s) look like impersonation, a password lure, or a payment scam.")
    }
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('What your staff should do')
    [void]$sb.AppendLine('- If an email says a password is expiring or a document is waiting, open the service directly. Do not use the link in the email.')
    [void]$sb.AppendLine('- If someone asks to change bank or payment details, confirm with a phone call you already have on file.')
    [void]$sb.AppendLine('- Watch for a coworker''s name in the subject when the sender is a free mailbox such as Gmail.')
    if (@($Contacts).Count -gt 0) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('Questions can go to ' + (@($Contacts) -join ', ') + '.')
    }
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('We can walk through anything that looks familiar and should be released.')
    return $sb.ToString().TrimEnd()
}

function Get-TriageCommandToken {
    param([string]$Command, [string]$Name)
    if ($Command -match "\|${Name}:([^|]*)") { return $Matches[1].Trim() }
    return $null
}

function Test-TriageConfigModuleInstalled {
    param([string[]]$InstalledNames)
    if ($PSBoundParameters.ContainsKey('InstalledNames')) {
        $present = @($InstalledNames)
    }
    else {
        $present = @(Get-Module -ListAvailable -Name ORCA, DomainHealthChecker | ForEach-Object { $_.Name } | Select-Object -Unique)
    }
    $need = @('ORCA', 'DomainHealthChecker')
    $missing = @($need | Where-Object { $present -notcontains $_ })
    if ($missing.Count -eq 0) { return $null }
    return ('TRIAGE_FAILED:Missing {0}. Install-Module ORCA,DomainHealthChecker -Scope CurrentUser' -f ($missing -join ', '))
}

function Get-TriageRowKey {
    param([string]$Source, [string]$MessageId, [string]$Recipient, [string]$Identity)
    $raw = '{0}|{1}|{2}|{3}' -f $Source, $MessageId, $Recipient, $Identity
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [Text.Encoding]::UTF8.GetBytes($raw.ToLowerInvariant())
        $hash = [BitConverter]::ToString($sha.ComputeHash($bytes)).Replace('-', '').ToLowerInvariant()
        return $hash.Substring(0, 16)
    }
    finally { $sha.Dispose() }
}

function Read-TriageCsv {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return @() }
    $rows = @(Import-Csv -LiteralPath $Path)
    return $rows
}

function Read-TriageJson {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return $null }
    $raw = Get-Content -LiteralPath $Path -Raw -ErrorAction SilentlyContinue
    if ([string]::IsNullOrWhiteSpace($raw)) { return $null }
    return ($raw | ConvertFrom-Json)
}

function Get-TriageCampaignIndex {
    param($Messages)
    $bySender = @{}
    $byLocal = @{}
    # @() on a generic List throws "Argument types do not match" in Windows PowerShell.
    if ($null -eq $Messages) { $Messages = @() }
    elseif ($Messages -isnot [System.Collections.IList]) { $Messages = @($Messages) }
    foreach ($m in $Messages) {
        if (-not $m) { continue }
        $sender = [string](Get-TriageProp $m @('SenderAddress', 'Sender'))
        $rcpt = [string](Get-TriageProp $m @('RecipientAddress', 'Recipient'))
        if (-not $sender) { continue }
        $sk = $sender.ToLowerInvariant()
        if (-not $bySender.ContainsKey($sk)) { $bySender[$sk] = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase) }
        if ($rcpt) { [void]$bySender[$sk].Add($rcpt) }
        $local = $sk
        $at = $sk.LastIndexOf('@')
        $dom = ''
        if ($at -gt 0) { $local = $sk.Substring(0, $at); $dom = $sk.Substring($at + 1) }
        if ($local -and $dom) {
            if (-not $byLocal.ContainsKey($local)) { $byLocal[$local] = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase) }
            [void]$byLocal[$local].Add($dom)
        }
    }
    return @{ BySender = $bySender; ByLocal = $byLocal }
}

function ConvertTo-TriageRows {
    param(
        $Quarantine,
        $Legs,
        $Headers,
        [string[]]$VipNames = @(),
        [string[]]$InternalDomains = @(),
        [string[]]$FreemailDomains = @(),
        $Overrides = @()
    )
    if (-not $FreemailDomains -or @($FreemailDomains).Count -eq 0) { $FreemailDomains = @(Get-TriageFreemailDomains) }
    $headerMap = @{}
    foreach ($h in (ConvertTo-TriageArray $Headers)) {
        if (-not $h) { continue }
        $id = [string](Get-TriageProp $h @('Identity'))
        if ($id) { $headerMap[$id] = $h }
    }
    $overrideMap = @{}
    foreach ($o in (ConvertTo-TriageArray $Overrides)) {
        if (-not $o) { continue }
        $k = [string](Get-TriageProp $o @('rowKey', 'RowKey'))
        $v = [string](Get-TriageProp $o @('verdict', 'Verdict'))
        if ($k -and $v) { $overrideMap[$k] = $v }
    }
    # Quarantine is the review list. The trace only answers two questions: did a quarantined
    # message also reach a mailbox, and did a delivered message look like a lure or an exec impersonation.
    # Spray and domain-rotation counts are not applied here; on a full tenant trace they mark noreply@ as malicious.
    $byMessage = @{}
    $pending = New-Object System.Collections.Generic.List[object]
    foreach ($q in (ConvertTo-TriageArray $Quarantine)) {
        if (-not $q) { continue }
        $item = [pscustomobject]@{
            Source = 'Quarantine'
            SenderAddress = [string](Get-TriageProp $q @('SenderAddress'))
            RecipientAddress = [string](Get-TriageProp $q @('RecipientAddress'))
            Subject = [string](Get-TriageProp $q @('Subject'))
            MessageId = [string](Get-TriageProp $q @('MessageId'))
            Identity = [string](Get-TriageProp $q @('Identity'))
            Received = Get-TriageProp $q @('ReceivedTime', 'Received')
            Layer = Get-TriageLayer -Type ([string](Get-TriageProp $q @('Type'))) -PolicyType ([string](Get-TriageProp $q @('PolicyType'))) -PolicyName ([string](Get-TriageProp $q @('PolicyName')))
            ReachedMailbox = $false
            Outcome = 'Quarantine'
            TraceNote = $false
        }
        $pending.Add($item)
        $mid = $item.MessageId.ToLowerInvariant()
        if ($mid) {
            if (-not $byMessage.ContainsKey($mid)) { $byMessage[$mid] = New-Object System.Collections.Generic.List[object] }
            $byMessage[$mid].Add($item)
        }
    }
    foreach ($leg in (ConvertTo-TriageArray $Legs)) {
        if (-not $leg) { continue }
        $mid = [string]$leg.MessageId
        $midKey = $mid.ToLowerInvariant()
        if ($midKey -and $byMessage.ContainsKey($midKey)) {
            # A later quarantine event folds ReachedMailbox back to false. For this review that
            # delivery still matters: the message was in a mailbox before it was held.
            $delivered = ($leg.ReachedMailbox -eq $true)
            if (-not $delivered) {
                foreach ($st in @(Get-TriageProp $leg @('Statuses'))) {
                    if ([string]$st -eq 'Delivered') { $delivered = $true }
                }
            }
            if ($delivered) {
                foreach ($qItem in $byMessage[$midKey]) {
                    $qItem.ReachedMailbox = $true
                    $qItem.Outcome = 'ReachedMailbox'
                    $qItem.TraceNote = $true
                }
            }
            continue
        }
        $reached = $leg.ReachedMailbox -eq $true
        if (-not $reached) { continue }
        $subject = [string]$leg.Subject
        $sender = [string]$leg.SenderAddress
        $imp = Get-TriageImpersonationMatch -Subject $subject -SenderAddress $sender -VipNames $VipNames -InternalDomains $InternalDomains
        $lure = $subject -match '(?i)password expires|voicemail|docusign|\bssa\b|review changes|wire transfer|bank details|updated banking'
        if (-not $imp -and -not $lure) { continue }
        $pending.Add([pscustomobject]@{
                Source = 'Trace'
                SenderAddress = $sender
                RecipientAddress = [string]$leg.RecipientAddress
                Subject = $subject
                MessageId = $mid
                Identity = ''
                Received = $leg.Received
                Layer = ''
                ReachedMailbox = $true
                Outcome = [string]$leg.Outcome
                TraceNote = $false
            })
    }
    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($item in $pending) {
        $parsed = $null
        $raw = ''
        if ($item.Identity -and $headerMap.ContainsKey($item.Identity)) {
            $raw = [string](Get-TriageProp $headerMap[$item.Identity] @('Raw'))
            $parsed = Get-TriageProp $headerMap[$item.Identity] @('Parsed')
            if (-not $parsed -and $raw) { $parsed = ConvertFrom-TriageMessageHeader -Header $raw }
        }
        $cls = Get-TriageClassification -Subject $item.Subject -SenderAddress $item.SenderAddress -VipNames $VipNames -InternalDomains $InternalDomains -FreemailDomains $FreemailDomains -SenderRecipientCount 1 -RotationDomainCount 1 -ParsedHeader $parsed
        $evidence = New-Object System.Collections.Generic.List[string]
        foreach ($e in @($cls.Evidence)) { if ($e) { $evidence.Add([string]$e) } }
        if ($item.TraceNote) { $evidence.Add('Message trace shows this message also reached a mailbox.') }
        $imp = Get-TriageImpersonationMatch -Subject $item.Subject -SenderAddress $item.SenderAddress -VipNames $VipNames -InternalDomains $InternalDomains
        $impCat = ''
        if ($imp) { $impCat = [string]$imp.Category }
        $key = Get-TriageRowKey -Source $item.Source -MessageId $item.MessageId -Recipient $item.RecipientAddress -Identity $item.Identity
        $rowVerdict = $cls.Verdict
        if ($overrideMap.ContainsKey($key)) { $rowVerdict = $overrideMap[$key] }
        $rows.Add([pscustomobject]@{
                RowKey            = $key
                Source            = $item.Source
                ReceivedUtc       = $item.Received
                SenderAddress     = $item.SenderAddress
                RecipientAddress  = $item.RecipientAddress
                Subject           = $item.Subject
                MessageId         = $item.MessageId
                Identity          = $item.Identity
                Layer             = $item.Layer
                SuggestedVerdict  = $cls.Verdict
                Verdict           = $rowVerdict
                Evidence          = @($evidence.ToArray())
                ReachedMailbox    = [bool]$item.ReachedMailbox
                Outcome           = $item.Outcome
                Impersonation     = $impCat
                RawHeader         = $raw
            })
    }
    return @($rows.ToArray())
}

function New-TriageReportHtml {
    param($Rows, $Findings, [string]$Company = '')
    function Enc([string]$t) {
        if ($null -eq $t) { return '' }
        return [System.Net.WebUtility]::HtmlEncode($t)
    }
    $body = New-Object System.Text.StringBuilder
    [void]$body.AppendLine('<!DOCTYPE html><html><head><meta charset="utf-8"><title>Email triage</title>')
    [void]$body.AppendLine('<style>body{font-family:Segoe UI,sans-serif}table{border-collapse:collapse;width:100%}td,th{border-bottom:1px solid #ddd;text-align:left;padding:4px 6px;font-size:13px}</style></head><body>')
    [void]$body.AppendLine('<h1>Email security triage ' + (Enc $Company) + '</h1>')
    [void]$body.AppendLine('<table><thead><tr><th>Verdict</th><th>Sender</th><th>Subject</th><th>Evidence</th></tr></thead><tbody>')
    foreach ($r in @($Rows)) {
        if (-not $r) { continue }
        $ev = @((Get-TriageProp $r @('Evidence'))) -join '; '
        [void]$body.AppendLine('<tr><td>' + (Enc ([string](Get-TriageProp $r @('Verdict')))) + '</td><td>' + (Enc ([string](Get-TriageProp $r @('SenderAddress')))) + '</td><td>' + (Enc ([string](Get-TriageProp $r @('Subject')))) + '</td><td>' + (Enc $ev) + '</td></tr>')
    }
    [void]$body.AppendLine('</tbody></table><h2>Findings</h2><ul>')
    foreach ($f in @($Findings)) {
        if (-not $f) { continue }
        [void]$body.AppendLine('<li>' + (Enc ([string](Get-TriageProp $f @('Severity')))) + ': ' + (Enc ([string](Get-TriageProp $f @('Detail')))) + '</li>')
    }
    [void]$body.AppendLine('</ul></body></html>')
    return $body.ToString()
}

function Get-TriageNotesPath {
    param([string]$NotesRoot, [string]$Client)
    if (-not $NotesRoot -or -not $Client) { return $null }
    if ($Client -notmatch '^[A-Za-z0-9_.-]{1,80}$') { throw 'Invalid client id' }
    if (-not (Test-Path -LiteralPath $NotesRoot)) { New-Item -ItemType Directory -Path $NotesRoot -Force | Out-Null }
    return (Join-Path $NotesRoot ($Client + '.json'))
}

function Get-TriageClientNotes {
    param([string]$NotesRoot, [string]$Client)
    $path = Get-TriageNotesPath -NotesRoot $NotesRoot -Client $Client
    $notes = [ordered]@{
        client = $Client; vips = @(); filterVendor = ''; contacts = @(); dnsHost = ''; extraFreemail = @()
    }
    if ($path -and (Test-Path -LiteralPath $path)) {
        $saved = Get-Content -LiteralPath $path -Raw | ConvertFrom-Json
        foreach ($p in @('vips', 'filterVendor', 'contacts', 'dnsHost', 'extraFreemail')) {
            $prop = $saved.PSObject.Properties[$p]
            if ($prop) { $notes[$p] = $prop.Value }
        }
    }
    return [pscustomobject]$notes
}

function Set-TriageClientNotes {
    param([string]$NotesRoot, [string]$Client, $Notes)
    $path = Get-TriageNotesPath -NotesRoot $NotesRoot -Client $Client
    $current = Get-TriageClientNotes -NotesRoot $NotesRoot -Client $Client
    foreach ($p in @('vips', 'filterVendor', 'contacts', 'dnsHost', 'extraFreemail')) {
        $prop = $Notes.PSObject.Properties[$p]
        if ($prop) {
            if ($p -eq 'filterVendor' -or $p -eq 'dnsHost') { $current.$p = [string]$prop.Value }
            else { $current.$p = @($prop.Value) }
        }
    }
    ($current | ConvertTo-Json -Depth 5) | Set-Content -LiteralPath $path -Encoding UTF8
    return $current
}

function Get-TriagePackRoot {
    param([string]$CompanyName)
    $safe = $CompanyName
    if ([string]::IsNullOrWhiteSpace($safe)) { $safe = 'UnknownCompany' }
    $safe = ($safe -replace '[^\w\-. ]', '_').Trim()
    $docs = [Environment]::GetFolderPath('MyDocuments')
    $root = Join-Path (Join-Path (Join-Path $docs 'ExchangeOnlineAnalyzer') 'SecurityInvestigation') $safe
    return $root
}

function New-TriagePackFolder {
    param([string]$CompanyName, [string]$ExistingFolder)
    if ($ExistingFolder) {
        $resolved = [IO.Path]::GetFullPath($ExistingFolder)
        $parent = Get-TriagePackRoot -CompanyName $CompanyName
        if (-not $resolved.StartsWith([IO.Path]::GetFullPath($parent), [StringComparison]::OrdinalIgnoreCase)) {
            throw 'Triage folder is outside the company pack root'
        }
        if ((Split-Path $resolved -Leaf) -notmatch '^Triage_') { throw 'Not a triage pack folder' }
        if (-not (Test-Path -LiteralPath $resolved)) { New-Item -ItemType Directory -Path $resolved -Force | Out-Null }
        return $resolved
    }
    $root = Get-TriagePackRoot -CompanyName $CompanyName
    if (-not (Test-Path -LiteralPath $root)) { New-Item -ItemType Directory -Path $root -Force | Out-Null }
    $folder = Join-Path $root ('Triage_' + (Get-Date -Format 'yyyyMMdd-HHmmss'))
    New-Item -ItemType Directory -Path $folder -Force | Out-Null
    return $folder
}

function Update-TriagePackOutputs {
    param(
        [Parameter(Mandatory)][string]$Folder,
        [string[]]$VipNames = @(),
        [string[]]$InternalDomains = @(),
        [string[]]$ExtraFreemail = @(),
        [string]$Company = ''
    )
    $freemail = @(Get-TriageFreemailDomains -Extra $ExtraFreemail)
    if (@($InternalDomains).Count -eq 0) {
        $accepted = Read-TriageJson -Path (Join-Path $Folder 'AcceptedDomains.json')
        if ($accepted) {
            $rawDomains = Get-TriageProp $accepted @('InternalDomains')
            if ($rawDomains) { $InternalDomains = @($rawDomains) }
        }
    }
    $q = Read-TriageCsv -Path (Join-Path $Folder 'Quarantine.csv')
    $trace = Read-TriageCsv -Path (Join-Path $Folder 'Trace.csv')
    $legs = @(Merge-TriageTraceLegs -Rows $trace -WindowSeconds 120)
    $headers = @()
    $headerDoc = Read-TriageJson -Path (Join-Path $Folder 'QuarantineHeaders.json')
    if ($headerDoc) { $headers = @($headerDoc) }
    $overrides = @()
    $ov = Read-TriageJson -Path (Join-Path $Folder 'overrides.json')
    if ($ov -and $ov.PSObject.Properties['items']) { $overrides = @($ov.items) }
    $config = Read-TriageJson -Path (Join-Path $Folder 'Config.json')
    $health = Read-TriageJson -Path (Join-Path $Folder 'DomainHealth.json')
    if ($config -and $health -and -not ($config.PSObject.Properties['DomainHealth'])) {
        $config | Add-Member -NotePropertyName DomainHealth -NotePropertyValue @($health) -Force
    }
    elseif ($config -and $health) { $config.DomainHealth = @($health) }
    $findings = @(Test-TriageConfigFindings -Config $config -FreemailDomains $freemail)
    $rows = @(ConvertTo-TriageRows -Quarantine $q -Legs $legs -Headers $headers -VipNames $VipNames -InternalDomains $InternalDomains -FreemailDomains $freemail -Overrides $overrides)
    $orca = Get-TriageOrcaSummary -Orca (Read-TriageJson -Path (Join-Path $Folder 'ORCA.json'))
    $ticket = New-TriageTicketNote -Rows $rows -Findings $findings -Company $Company
    $email = New-TriageClientEmail -Rows $rows -Company $Company
    $releases = @($rows | Where-Object { $_.Source -eq 'Quarantine' -and $_.Verdict -eq 'Legit' -and $_.Identity } | ForEach-Object { $_.Identity })
    $blocks = @($rows | Where-Object { $_.Verdict -eq 'Malicious' -and $_.SenderAddress } | ForEach-Object { Get-TriageBlockTarget -SenderAddress $_.SenderAddress -FreemailDomains $freemail })
    $protect = @($VipNames)
    $soft = @($rows | Where-Object { $_.Verdict -eq 'Malicious' -and $_.ReachedMailbox -and $_.MessageId } | ForEach-Object { $_.MessageId })
    $scriptText = New-TriageActionScript -ReleaseIdentities $releases -Blocks $blocks -ProtectUsers $protect -SoftDeleteMessageIds $soft
    $report = New-TriageReportHtml -Rows $rows -Findings $findings -Company $Company
    @{ Findings = @($findings); GeneratedUtc = (Get-Date).ToUniversalTime().ToString('o') } | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $Folder 'Findings.json') -Encoding UTF8
    $rows | Select-Object RowKey, Source, ReceivedUtc, SenderAddress, RecipientAddress, Subject, MessageId, Identity, Layer, SuggestedVerdict, Verdict, @{n='Evidence';e={@($_.Evidence) -join '; '}}, ReachedMailbox, Outcome, Impersonation |
        Export-Csv -LiteralPath (Join-Path $Folder 'Triage.csv') -NoTypeInformation -Encoding UTF8
    Set-Content -LiteralPath (Join-Path $Folder 'actions.ps1') -Value $scriptText -Encoding UTF8
    Set-Content -LiteralPath (Join-Path $Folder 'TicketNote.txt') -Value $ticket -Encoding UTF8
    Set-Content -LiteralPath (Join-Path $Folder 'ClientEmail.txt') -Value $email -Encoding UTF8
    Set-Content -LiteralPath (Join-Path $Folder 'Report.html') -Value $report -Encoding UTF8
    $warning = ''
    $warnPath = Join-Path $Folder 'PullWarning.txt'
    if (Test-Path -LiteralPath $warnPath) { $warning = [string](Get-Content -LiteralPath $warnPath -Raw) }
    return [pscustomobject]@{
        Folder     = $Folder
        Warning    = $warning.Trim()
        Rows       = @($rows)
        Findings   = @($findings)
        Orca       = $orca
        TicketNote = $ticket
        ClientEmail = $email
    }
}

function Read-TriageSavedPackView {
    param(
        [Parameter(Mandatory)][string]$Folder,
        [int]$Limit = 0,
        [string]$Verdict = '',
        [string[]]$VipNames = @()
    )
    $overrideMap = @{}
    $reviewMap = @{}
    $ov = Read-TriageJson -Path (Join-Path $Folder 'overrides.json')
    if ($ov) {
        foreach ($item in (ConvertTo-TriageArray (Get-TriageProp $ov @('items')))) {
            $k = [string](Get-TriageProp $item @('rowKey', 'RowKey'))
            $v = [string](Get-TriageProp $item @('verdict', 'Verdict'))
            if ($k -and $v) { $overrideMap[$k] = $v }
        }
        foreach ($item in (ConvertTo-TriageArray (Get-TriageProp $ov @('review')))) {
            $k = [string](Get-TriageProp $item @('rowKey', 'RowKey'))
            if (-not $k) { continue }
            $flag = Get-TriageProp $item @('include', 'Include')
            if ($flag -is [bool]) { $reviewMap[$k] = [bool]$flag }
            elseif ([string]$flag -eq 'True') { $reviewMap[$k] = $true }
            elseif ([string]$flag -eq 'False') { $reviewMap[$k] = $false }
        }
    }
    $headerMap = @{}
    $headerDoc = Read-TriageJson -Path (Join-Path $Folder 'QuarantineHeaders.json')
    foreach ($h in (ConvertTo-TriageArray $headerDoc)) {
        $id = [string](Get-TriageProp $h @('Identity'))
        $raw = [string](Get-TriageProp $h @('Raw'))
        if ($id -and $raw) { $headerMap[$id] = $raw }
    }
    $counts = @{ Total = 0; Malicious = 0; Legit = 0; Noise = 0; Unknown = 0; Reached = 0; Impersonation = 0 }
    $truncatedHeaders = $false
    $rows = New-Object System.Collections.Generic.List[object]
    $actionRows = New-Object System.Collections.Generic.List[object]
    $delivery = New-Object System.Collections.Generic.List[object]
    $impRows = New-Object System.Collections.Generic.List[object]
    foreach ($item in (ConvertTo-TriageArray (Import-Csv -LiteralPath (Join-Path $Folder 'Triage.csv')))) {
        if (-not $item) { continue }
        $counts.Total++
        $key = [string]$item.RowKey
        $rowVerdict = [string]$item.Verdict
        if ($overrideMap.ContainsKey($key)) { $rowVerdict = $overrideMap[$key] }
        if ($counts.ContainsKey($rowVerdict)) { $counts[$rowVerdict]++ }
        $reached = ([string]$item.ReachedMailbox) -eq 'True'
        if ($reached) { $counts.Reached++ }
        $impCat = [string]$item.Impersonation
        if ($impCat) { $counts.Impersonation++ }
        $aiReview = @('Malicious', 'Unknown') -contains $rowVerdict
        if ($reviewMap.ContainsKey($key)) { $aiReview = [bool]$reviewMap[$key] }
        $raw = ''
        $id = [string]$item.Identity
        if ($id -and $headerMap.ContainsKey($id)) {
            $candidate = [string]$headerMap[$id]
            $raw = Get-TriageHeaderText $candidate
            if ($candidate -and -not $raw) { $truncatedHeaders = $true }
        }
        $row = [pscustomobject]@{
            RowKey           = $key
            Source           = [string]$item.Source
            ReceivedUtc      = [string]$item.ReceivedUtc
            SenderAddress    = [string]$item.SenderAddress
            RecipientAddress = [string]$item.RecipientAddress
            Subject          = [string]$item.Subject
            MessageId        = [string]$item.MessageId
            Identity         = $id
            Layer            = [string]$item.Layer
            SuggestedVerdict = [string]$item.SuggestedVerdict
            Verdict          = $rowVerdict
            Evidence         = [string]$item.Evidence
            ReachedMailbox   = $reached
            Outcome          = [string]$item.Outcome
            Impersonation    = $impCat
            RawHeader        = $raw
            AiReview         = $aiReview
        }
        $actionRows.Add($row)
        $cap = if ($Limit -le 0) { [int]::MaxValue } else { $Limit }
        if ($rows.Count -lt $cap -and ((-not $Verdict) -or $rowVerdict -eq $Verdict)) { $rows.Add($row) }
        if ($reached -and $delivery.Count -lt $cap) { $delivery.Add($row) }
        if ($impCat -and $impRows.Count -lt $cap) { $impRows.Add($row) }
    }
    $skipKeys = Get-TriageAiBatchSkipKeys -Rows $actionRows.ToArray()
    foreach ($row in $actionRows) {
        if ($reviewMap.ContainsKey([string]$row.RowKey)) { continue }
        if ($skipKeys.ContainsKey([string]$row.RowKey)) { $row.AiReview = $false }
    }
    $findings = @()
    $findingDoc = Read-TriageJson -Path (Join-Path $Folder 'Findings.json')
    if ($findingDoc) { $findings = @(ConvertTo-TriageArray (Get-TriageProp $findingDoc @('Findings'))) }
    $warning = ''
    $warnPath = Join-Path $Folder 'PullWarning.txt'
    if (Test-Path -LiteralPath $warnPath) { $warning = [string](Get-Content -LiteralPath $warnPath -Raw) }
    if ($truncatedHeaders) {
        $warning = ($warning.Trim() + ' Saved quarantine headers are a truncated table, not the message. Run Quarantine again, then copy a row into Message Header Analyzer.').Trim()
    }
    $ticket = Read-TriageBoundedText -Folder $Folder -Name 'TicketNote.txt' -Label 'Ticket note'
    $email = Read-TriageBoundedText -Folder $Folder -Name 'ClientEmail.txt' -Label 'Client email'
    $orcaSummary = Get-TriageOrcaSummary -Orca (Read-TriageJson -Path (Join-Path $Folder 'ORCA.json'))
    $accepted = Read-TriageJson -Path (Join-Path $Folder 'AcceptedDomains.json')
    $internal = @()
    if ($accepted) { $internal = @(ConvertTo-TriageArray (Get-TriageProp $accepted @('InternalDomains'))) }
    $config = Read-TriageJson -Path (Join-Path $Folder 'Config.json')
    $health = Read-TriageJson -Path (Join-Path $Folder 'DomainHealth.json')
    if ($config -and $health -and -not ($config.PSObject.Properties['DomainHealth'])) {
        $config | Add-Member -NotePropertyName DomainHealth -NotePropertyValue @(ConvertTo-TriageArray $health) -Force
    }
    $actionItems = @(Get-TriageRemediationItems -Rows @($actionRows.ToArray()) -Orca $orcaSummary -Config $config -InternalDomains $internal -VipNames $VipNames)
    return [pscustomobject]@{
        Folder            = $Folder
        Warning           = $warning.Trim()
        Rows              = @($rows.ToArray())
        DeliveryRows      = @($delivery.ToArray())
        ImpersonationRows = @($impRows.ToArray())
        Counts            = $counts
        Slice             = $(if ($Verdict) { $Verdict } else { 'rows' })
        Findings          = @($findings)
        Orca              = $orcaSummary
        Actions           = @($actionItems)
        TicketNote        = $ticket
        ClientEmail       = $email
    }
}

function Read-TriageBoundedText {
    param([string]$Folder, [string]$Name, [string]$Label)
    $path = Join-Path $Folder $Name
    if (-not (Test-Path -LiteralPath $path)) { return '' }
    $len = (Get-Item -LiteralPath $path).Length
    if ($len -gt 262144) {
        return ('{0} is {1} MB. Use the copy button to read the full file.' -f $Label, ([math]::Round($len / 1MB, 1)))
    }
    return [string](Get-Content -LiteralPath $path -Raw)
}

function Get-TriagePackView {
    param(
        [Parameter(Mandatory)][string]$Folder,
        [string[]]$VipNames = @(),
        [string[]]$InternalDomains = @(),
        [string[]]$ExtraFreemail = @(),
        [string]$Company = '',
        [int]$Limit = 0,
        [string]$Verdict = ''
    )
    $saved = Join-Path $Folder 'Triage.csv'
    if ((Test-Path -LiteralPath $saved) -and ($Limit -gt 0 -or (Test-Path -LiteralPath (Join-Path $Folder 'TicketNote.txt')))) {
        return Read-TriageSavedPackView -Folder $Folder -Limit $Limit -Verdict $Verdict -VipNames $VipNames
    }
    return Update-TriagePackOutputs -Folder $Folder -VipNames $VipNames -InternalDomains $InternalDomains -ExtraFreemail $ExtraFreemail -Company $Company
}

function Set-TriageVerdictOverride {
    param(
        [Parameter(Mandatory)][string]$Folder,
        [Parameter(Mandatory)][string]$RowKey,
        [Parameter(Mandatory)][string]$Verdict
    )
    $path = Join-Path $Folder 'overrides.json'
    $doc = $null
    $items = @()
    if (Test-Path -LiteralPath $path) {
        $doc = Get-Content -LiteralPath $path -Raw | ConvertFrom-Json
        if ($doc -and $doc.PSObject.Properties['items']) { $items = @(ConvertTo-TriageArray $doc.items) }
    }
    $review = New-Object System.Collections.ArrayList
    if ($doc -and $doc.PSObject.Properties['review']) {
        foreach ($entry in @(ConvertTo-TriageArray $doc.review)) {
            if ($entry) { [void]$review.Add($entry) }
        }
    }
    if ($Verdict -eq 'Include' -or $Verdict -eq 'Omit') {
        $include = ($Verdict -eq 'Include')
        $keys = New-Object System.Collections.Generic.List[string]
        foreach ($part in ($RowKey -split '\|')) {
            $part = $part.Trim()
            if ($part) { [void]$keys.Add($part) }
        }
        if ($keys.Count -eq 0) { throw 'Missing row key.' }
        $keptReview = New-Object System.Collections.ArrayList
        foreach ($entry in $review) {
            $k = [string](Get-TriageProp $entry @('rowKey', 'RowKey'))
            if (-not ($keys -contains $k)) { [void]$keptReview.Add($entry) }
        }
        foreach ($k in $keys) {
            [void]$keptReview.Add([pscustomobject]@{ rowKey = $k; include = $include })
        }
        @{ items = @($items); review = @($keptReview.ToArray()) } | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $path -Encoding UTF8
        return
    }
    $allowed = @('Malicious', 'Legit', 'Noise', 'Unknown')
    if ($allowed -notcontains $Verdict) { throw "Verdict must be one of: $($allowed -join ', ')" }
    $kept = New-Object System.Collections.Generic.List[object]
    foreach ($item in $items) {
        if ([string](Get-TriageProp $item @('rowKey')) -ne $RowKey) { $kept.Add($item) }
    }
    $kept.Add([pscustomobject]@{ rowKey = $RowKey; verdict = $Verdict; at = (Get-Date).ToUniversalTime().ToString('o') })
    @{ items = @($kept.ToArray()); review = @($review.ToArray()) } | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $path -Encoding UTF8
}

function Get-TriagePacks {
    param([string]$Root = '')
    if (-not $Root) {
        $docs = [Environment]::GetFolderPath('MyDocuments')
        $Root = Join-Path (Join-Path $docs 'ExchangeOnlineAnalyzer') 'SecurityInvestigation'
    }
    if (-not (Test-Path -LiteralPath $Root)) { return @() }
    $packs = New-Object System.Collections.Generic.List[object]
    foreach ($company in @(Get-ChildItem -LiteralPath $Root -Directory -ErrorAction SilentlyContinue)) {
        foreach ($pack in @(Get-ChildItem -LiteralPath $company.FullName -Directory -Filter 'Triage_*' -ErrorAction SilentlyContinue)) {
            $packs.Add([pscustomobject]@{
                    Company = $company.Name
                    Name    = $pack.Name
                    Folder  = $pack.FullName
                    Stamp   = $pack.Name.Substring(7)
                })
        }
    }
    return @($packs | Sort-Object Folder -Descending)
}

function Get-TriagePackFileContent {
    param([Parameter(Mandatory)][string]$Folder, [Parameter(Mandatory)][string]$Name)
    $allowed = @{
        'actions.ps1' = $true; 'Triage.csv' = $true; 'Report.html' = $true
        'TicketNote.txt' = $true; 'ClientEmail.txt' = $true; 'Quarantine.csv' = $true
        'Trace.csv' = $true; 'Findings.json' = $true; 'ORCA.json' = $true
        'DomainHealth.json' = $true; 'Config.json' = $true
    }
    if (-not $allowed.ContainsKey($Name)) { throw 'File is not available for download' }
    $path = Join-Path $Folder $Name
    if (-not (Test-Path -LiteralPath $path)) { throw "File not found: $Name" }
    return [pscustomobject]@{ FileName = $Name; Content = [string](Get-Content -LiteralPath $path -Raw) }
}

function Export-TriageRowsCsv {
    param($Rows, [string]$Path, [string[]]$Properties)
    ConvertTo-TriageArray $Rows | Select-Object $Properties | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding UTF8
}

function Invoke-TriageQuarantinePull {
    param(
        [Parameter(Mandatory)][string]$Folder,
        [int]$Days = 7,
        [string]$SenderDomain = '',
        [bool]$FreemailOnly = $false,
        [int]$HeaderTopN = 5,
        [string[]]$FreemailDomains = @()
    )
    if (-not (Get-Command Get-QuarantineMessage -ErrorAction SilentlyContinue)) {
        throw 'Get-QuarantineMessage is not available. Connect to Exchange Online first.'
    }
    $start = (Get-Date).AddDays(-[Math]::Max(1, $Days))
    $end = Get-Date
    $collected = New-Object System.Collections.Generic.List[object]
    $page = 1
    do {
        $batch = ConvertTo-TriageArray (Get-QuarantineMessage -StartReceivedDate $start -EndReceivedDate $end -Page $page -PageSize 1000 -ErrorAction Stop)
        foreach ($item in $batch) { if ($item) { $collected.Add($item) } }
        $page++
    } while ($batch.Count -eq 1000 -and $page -le 40)
    $freemail = @($FreemailDomains)
    $projected = New-Object System.Collections.Generic.List[object]
    foreach ($item in $collected) {
        $sender = [string](Get-TriageProp $item @('SenderAddress'))
        $dom = Get-TriageDomain $sender
        if ($SenderDomain -and $dom -ne $SenderDomain.Trim().ToLowerInvariant()) { continue }
        if ($FreemailOnly) {
            $ok = $false
            foreach ($d in $freemail) { if ($dom -eq $d) { $ok = $true } }
            if (-not $ok) { continue }
        }
        $projected.Add([pscustomobject]@{
                ReceivedTime     = [string](Get-TriageProp $item @('ReceivedTime'))
                SenderAddress    = $sender
                RecipientAddress = [string](Get-TriageProp $item @('RecipientAddress'))
                Subject          = [string](Get-TriageProp $item @('Subject'))
                Type             = [string](Get-TriageProp $item @('Type'))
                PolicyType       = [string](Get-TriageProp $item @('PolicyType'))
                PolicyName       = [string](Get-TriageProp $item @('PolicyName'))
                MessageId        = [string](Get-TriageProp $item @('MessageId'))
                Identity         = [string](Get-TriageProp $item @('Identity'))
                Direction        = [string](Get-TriageProp $item @('Direction'))
            })
    }
    Export-TriageRowsCsv -Rows $projected -Path (Join-Path $Folder 'Quarantine.csv') -Properties @(
        'ReceivedTime', 'SenderAddress', 'RecipientAddress', 'Subject', 'Type', 'PolicyType', 'PolicyName', 'MessageId', 'Identity', 'Direction'
    )
    $headerDocs = New-Object System.Collections.Generic.List[object]
    $take = [Math]::Max(0, $HeaderTopN)
    $n = 0
    if ($take -gt 0 -and (Get-Command Get-QuarantineMessageHeader -ErrorAction SilentlyContinue)) {
        foreach ($row in $projected) {
            if ($n -ge $take) { break }
            if (-not $row.Identity) { continue }
            $raw = ''
            try { $raw = Get-TriageHeaderText (Get-QuarantineMessageHeader -Identity $row.Identity -ErrorAction Stop) } catch { $raw = '' }
            $headerDocs.Add([pscustomobject]@{
                    Identity = $row.Identity
                    Raw      = $raw.Trim()
                    Parsed   = (ConvertFrom-TriageMessageHeader -Header $raw)
                })
            $n++
        }
    }
    @($headerDocs.ToArray()) | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $Folder 'QuarantineHeaders.json') -Encoding UTF8
    return $projected.Count
}

function Invoke-TriageDeliveryCheck {
    param(
        [Parameter(Mandatory)][string]$Folder,
        [int]$Days = 7,
        [string]$SenderAddress = '',
        [string]$Subject = '',
        [int]$WindowSeconds = 120,
        [string]$ScriptRoot = ''
    )
    $requested = $Days
    $effective = [Math]::Min([Math]::Max(1, $Days), 10)
    if ($requested -gt 10) {
        Set-Content -LiteralPath (Join-Path $Folder 'PullWarning.txt') -Encoding UTF8 -Value 'Message trace was limited to 10 days. Older mail needs a historical trace, which this pull does not run.'
    }
    if (-not (Get-Command Get-MessageTraceV2 -ErrorAction SilentlyContinue)) {
        throw 'Get-MessageTraceV2 is not available. Connect to Exchange Online first.'
    }
    $euPath = $null
    if ($ScriptRoot) { $euPath = Join-Path $ScriptRoot 'Modules\ExportUtils.psm1' }
    if ($euPath -and (Test-Path -LiteralPath $euPath)) { Import-Module $euPath -Force -ErrorAction Stop }
    $mod = Get-Module ExportUtils
    if (-not $mod) { throw 'ExportUtils is not loaded' }
    $end = (Get-Date).ToUniversalTime()
    $start = $end.AddDays(-$effective)
    $v2 = Get-Command Get-MessageTraceV2
    $page = & $mod {
        param($V2, $Start, $End, $Sender, $SubjectFilter)
        Invoke-EoaMessageTraceV2Paged -V2Command $V2 -WinStart $Start -WinEnd $End -SenderAddress $Sender -Subject $SubjectFilter -Label 'triage delivery'
    } $v2 $start $end $(if ($SenderAddress) { $SenderAddress } else { $null }) $(if ($Subject) { $Subject } else { $null })
    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($item in (ConvertTo-TriageArray $page.Rows)) {
        if (-not $item) { continue }
        $rows.Add([pscustomobject]@{
                Received         = [string](Get-TriageProp $item @('Received'))
                SenderAddress    = [string](Get-TriageProp $item @('SenderAddress'))
                RecipientAddress = [string](Get-TriageProp $item @('RecipientAddress'))
                Subject          = [string](Get-TriageProp $item @('Subject'))
                Status           = [string](Get-TriageProp $item @('Status'))
                MessageId        = [string](Get-TriageProp $item @('MessageId'))
                MessageTraceId   = [string](Get-TriageProp $item @('MessageTraceId'))
            })
    }
    Export-TriageRowsCsv -Rows $rows -Path (Join-Path $Folder 'Trace.csv') -Properties @(
        'Received', 'SenderAddress', 'RecipientAddress', 'Subject', 'Status', 'MessageId', 'MessageTraceId'
    )
    $legs = @(Merge-TriageTraceLegs -Rows $rows -WindowSeconds $WindowSeconds)
    $details = New-Object System.Collections.Generic.List[object]
    if (Get-Command Get-MessageTraceDetailV2 -ErrorAction SilentlyContinue) {
        $n = 0
        foreach ($leg in $legs) {
            if (-not $leg.ReachedMailbox -or -not $leg.MessageTraceId) { continue }
            if ($n -ge 40) { break }
            try {
                $detail = ConvertTo-TriageArray (Get-MessageTraceDetailV2 -MessageTraceId $leg.MessageTraceId -RecipientAddress $leg.RecipientAddress -ErrorAction Stop)
                foreach ($d in $detail) {
                    $details.Add([pscustomobject]@{
                            MessageTraceId   = $leg.MessageTraceId
                            RecipientAddress = $leg.RecipientAddress
                            Event            = [string](Get-TriageProp $d @('Event'))
                            Action           = [string](Get-TriageProp $d @('Action'))
                            Detail           = [string](Get-TriageProp $d @('Detail', 'Data'))
                        })
                }
            }
            catch { }
            $n++
        }
    }
    Export-TriageRowsCsv -Rows $details -Path (Join-Path $Folder 'TraceDetail.csv') -Properties @('MessageTraceId', 'RecipientAddress', 'Event', 'Action', 'Detail')
    return [pscustomobject]@{ Trace = $rows.Count; Reached = @($legs | Where-Object { $_.ReachedMailbox }).Count; Limited = ($requested -gt 10) }
}

function Invoke-TriageSaveAcceptedDomains {
    param([string]$Folder)
    $domains = @()
    if (Get-Command Get-AcceptedDomain -ErrorAction SilentlyContinue) {
        foreach ($d in @(Get-AcceptedDomain -ErrorAction SilentlyContinue)) {
            $name = [string](Get-TriageProp $d @('DomainName'))
            if ($name) { $domains += $name }
        }
    }
    @{ InternalDomains = @($domains) } | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $Folder 'AcceptedDomains.json') -Encoding UTF8
    return @($domains)
}

function Invoke-TriageConfigAudit {
    param(
        [Parameter(Mandatory)][string]$Folder
    )
    $missing = Test-TriageConfigModuleInstalled
    if ($missing) { throw $missing.Substring('TRIAGE_FAILED:'.Length) }
    # -Global: ORCA's version check calls Get-Module. A module-scoped import is invisible there, and the fallback indexes a missing ORCAPreview module.
    Import-Module ORCA -Global -ErrorAction Stop
    Import-Module DomainHealthChecker -Global -ErrorAction Stop
    $orcaParams = @{
        Connect             = $false
        Output              = @('JSON')
        OutputOptions       = @{ JSON = @{ OutputDirectory = $Folder } }
        PerformVersionCheck = $false
        InstallModules      = $false
        ShowSurvey          = $false
        SCC                 = $false
    }
    $produced = ConvertTo-TriageArray (Invoke-ORCA @orcaParams)
    $src = $null
    foreach ($item in $produced) {
        $path = [string](Get-TriageProp $item @('Result'))
        if ($path -and (Test-Path -LiteralPath $path)) { $src = $path }
    }
    if (-not $src) {
        $found = Get-ChildItem -LiteralPath $Folder -Filter 'ORCA-*.json' -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if ($found) { $src = $found.FullName }
    }
    if (-not $src) { throw 'ORCA did not write a JSON report' }
    Copy-Item -LiteralPath $src -Destination (Join-Path $Folder 'ORCA.json') -Force

    $domains = @(Invoke-TriageSaveAcceptedDomains -Folder $Folder)
    $health = New-Object System.Collections.Generic.List[object]
    $dnsRows = New-Object System.Collections.Generic.List[object]
    $mxRows = New-Object System.Collections.Generic.List[object]
    foreach ($domain in $domains) {
        try {
            foreach ($row in @(Invoke-SpfDkimDmarc -Name $domain -ErrorAction Stop)) { $health.Add($row) }
        }
        catch {
            $health.Add([pscustomobject]@{ Name = $domain; DmarcRecord = ''; Error = $_.Exception.Message })
        }
        $ns = @()
        $soa = ''
        try {
            $ns = @(Resolve-DnsName -Name $domain -Type NS -ErrorAction Stop | Where-Object { $_.NameHost } | ForEach-Object { [string]$_.NameHost })
        }
        catch { }
        try {
            $soaRec = Resolve-DnsName -Name $domain -Type SOA -ErrorAction Stop | Where-Object { $_.NameAdministrator -or $_.PrimaryServer } | Select-Object -First 1
            if ($soaRec) {
                $soa = [string]$soaRec.PrimaryServer
                if (-not $soa) { $soa = [string]$soaRec.NameAdministrator }
            }
        }
        catch { }
        $dnsRows.Add([pscustomobject]@{
                Domain = $domain
                Ns     = @($ns)
                Soa    = $soa
                Host   = (Get-TriageDnsHost -Ns $ns -Soa $soa)
            })
        $ex = @()
        try {
            $ex = @(Resolve-DnsName -Name $domain -Type MX -ErrorAction Stop | Where-Object { $_.NameExchange } | ForEach-Object { [string]$_.NameExchange })
        }
        catch { }
        $mxRows.Add([pscustomobject]@{ Domain = $domain; Exchanges = @($ex) })
    }
    @($health.ToArray()) | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $Folder 'DomainHealth.json') -Encoding UTF8

    $rules = @()
    if (Get-Command Get-TransportRule -ErrorAction SilentlyContinue) {
        foreach ($rule in @(Get-TransportRule -ErrorAction SilentlyContinue)) {
            $rules += [pscustomobject]@{
                Name                   = [string](Get-TriageProp $rule @('Name'))
                ExceptIfSenderDomainIs = @(ConvertTo-TriageArray (Get-TriageProp $rule @('ExceptIfSenderDomainIs')))
            }
        }
    }
    $connectors = @()
    if (Get-Command Get-InboundConnector -ErrorAction SilentlyContinue) {
        foreach ($c in @(Get-InboundConnector -ErrorAction SilentlyContinue)) {
            $connectors += [pscustomobject]@{
                Name         = [string](Get-TriageProp $c @('Name'))
                EFSkipIPs    = @(ConvertTo-TriageArray (Get-TriageProp $c @('EFSkipIPs')))
                EFSkipLastIP = [bool](Get-TriageProp $c @('EFSkipLastIP'))
            }
        }
    }
    $policies = @()
    if (Get-Command Get-AntiPhishPolicy -ErrorAction SilentlyContinue) {
        foreach ($p in @(Get-AntiPhishPolicy -ErrorAction SilentlyContinue)) {
            $policies += [pscustomobject]@{
                Name                   = [string](Get-TriageProp $p @('Name'))
                TargetedUsersToProtect = @(ConvertTo-TriageArray (Get-TriageProp $p @('TargetedUsersToProtect')))
            }
        }
    }
    $tabl = @()
    if (Get-Command Get-TenantAllowBlockListItems -ErrorAction SilentlyContinue) {
        foreach ($item in @(Get-TenantAllowBlockListItems -ListType Sender -ErrorAction SilentlyContinue)) {
            $tabl += [pscustomobject]@{
                Value  = [string](Get-TriageProp $item @('Value'))
                Action = [string](Get-TriageProp $item @('Action'))
            }
        }
    }
    $config = [ordered]@{
        TransportRules     = @($rules)
        InboundConnectors  = @($connectors)
        AntiPhishPolicies  = @($policies)
        TablItems          = @($tabl)
        Dns                = @($dnsRows.ToArray())
        Mx                 = @($mxRows.ToArray())
    }
    $config | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $Folder 'Config.json') -Encoding UTF8
    return @{ domains = $domains.Count }
}

function ConvertTo-TriagePsLiteral {
    param([string]$Value)
    if ($null -eq $Value) { $Value = '' }
    return "'" + ($Value.Replace("'", "''")) + "'"
}

function Get-TriageEvidenceText {
    param($Row)
    $ev = Get-TriageProp $Row @('Evidence')
    if ($null -eq $ev) { return '' }
    if ($ev -is [string]) { return $ev }
    return ((ConvertTo-TriageArray $ev) -join '; ')
}

function Test-TriageRowReached {
    param($Row)
    $flag = Get-TriageProp $Row @('ReachedMailbox')
    return (($flag -eq $true) -or ([string]$flag -eq 'True'))
}

function Test-TriageFreemailAddress {
    param([string]$Address, [string[]]$FreemailDomains)
    $dom = Get-TriageDomain $Address
    foreach ($d in @($FreemailDomains)) {
        if ($dom -and $dom -eq $d.ToLowerInvariant()) { return $true }
    }
    return $false
}

function Test-TriageCustomerDomain {
    param([string]$Name)
    $n = ([string]$Name).Trim().ToLowerInvariant().TrimEnd('.')
    if (-not $n) { return $false }
    if ($n -match '(^|\.)onmicrosoft\.com$') { return $false }
    if ($n -match 'inkyphishfence\.com$') { return $false }
    return $true
}

function Test-TriageBlockCandidate {
    param($Row, [string[]]$FreemailDomains)
    if (-not (Test-TriageRowReached $Row)) { return $false }
    $sender = [string](Get-TriageProp $Row @('SenderAddress'))
    $subject = [string](Get-TriageProp $Row @('Subject'))
    $evidence = Get-TriageEvidenceText $Row
    $imp = [string](Get-TriageProp $Row @('Impersonation'))
    $free = Test-TriageFreemailAddress -Address $sender -FreemailDomains $FreemailDomains
    $dom = Get-TriageDomain $sender
    $authFail = $evidence -match '(?i)(compauth=fail|\b(spf|dkim|dmarc)=fail)'
    $catPattern = '(?i)Forefront CAT: (' + (($script:TriageThreatCat.Keys | ForEach-Object { [regex]::Escape($_) }) -join '|') + ')\b'
    $headerThreat = ($evidence -match $catPattern) -or ($evidence -match '(?i)Forefront SFTY:') -or ($evidence -match '(?i)Forefront SFV: (SPM|SKB|BLK|SKS)\b')
    $credential = $evidence -match 'Credential-lure wording'
    $payment = $subject -match '(?i)wire transfer|bank details|updated banking'
    $knownService = @('vonage.com', 'ringcentral.com', 'docusign.com', 'docusign.net', 'adobesign.com', 'adobe.com') -contains $dom
    if ($knownService -and -not $authFail -and -not $imp -and -not $headerThreat) { return $false }
    if (($subject -match '(?i)voicemail') -and -not $free -and -not $imp -and -not $headerThreat -and -not $authFail) { return $false }
    if ($imp) { return $true }
    if ($free -and ($credential -or $payment)) { return $true }
    if ($headerThreat) { return $true }
    if ($authFail -and ($credential -or $payment)) { return $true }
    return $false
}

function Test-TriageAllowCandidate {
    param($Row, [string[]]$FreemailDomains, [string[]]$InternalDomains)
    $source = [string](Get-TriageProp $Row @('Source'))
    if ($source -ne 'Quarantine') { return $false }
    if (Test-TriageRowReached $Row) { return $false }
    if (Test-TriageBlockCandidate -Row $Row -FreemailDomains $FreemailDomains) { return $false }
    $verdict = [string](Get-TriageProp $Row @('Verdict'))
    if ($verdict -eq 'Malicious' -or $verdict -eq 'Noise') { return $false }
    if ($verdict -eq 'Legit') { return $true }
    $sender = [string](Get-TriageProp $Row @('SenderAddress'))
    $dom = Get-TriageDomain $sender
    foreach ($d in @($InternalDomains)) {
        if ($dom -and $dom -eq $d.ToLowerInvariant()) { return $true }
    }
    if (Test-TriageFreemailAddress -Address $sender -FreemailDomains $FreemailDomains) { return $false }
    $evidence = Get-TriageEvidenceText $Row
    $pass = ($evidence -match '(?i)\bSPF=pass\b') -and ($evidence -match '(?i)\bDKIM=pass\b') -and ($evidence -match '(?i)\bDMARC=pass\b')
    return [bool]$pass
}

function New-TriageRemediationItem {
    param(
        [string]$Id,
        [string]$Kind,
        [string]$Title,
        [string]$Detail,
        [string]$License,
        [string]$Apply,
        [string]$Operation = '',
        [string]$Entry = '',
        [string]$RuleName = '',
        [string]$Script = '',
        [string[]]$ReleaseIdentities = @()
    )
    return [pscustomobject]@{
        Id                 = ($Id -replace ';', '')
        Kind               = $Kind
        Title              = $Title
        Detail             = $Detail
        License            = $License
        Apply              = $Apply
        Operation          = $Operation
        Entry              = $Entry
        RuleName           = $RuleName
        ReleaseIdentities  = @($ReleaseIdentities)
        Script             = $Script
    }
}

function Get-TriageOrcaFailNames {
    param($Orca)
    $names = New-Object System.Collections.Generic.List[string]
    foreach ($group in @(ConvertTo-TriageArray (Get-TriageProp $Orca @('Groups')))) {
        if (-not $group) { continue }
        foreach ($item in @(ConvertTo-TriageArray (Get-TriageProp $group @('Items')))) {
            if (-not $item) { continue }
            $result = [string](Get-TriageProp $item @('Result'))
            if ($result -ne 'Fail') { continue }
            $name = [string](Get-TriageProp $item @('Name'))
            if ($name) { $names.Add($name) }
        }
    }
    return @($names.ToArray())
}

function Get-TriageRemediationItems {
    param(
        $Rows,
        $Orca,
        $Config,
        [string[]]$InternalDomains = @(),
        [string[]]$VipNames = @(),
        [string[]]$FreemailDomains = @()
    )
    if (-not $FreemailDomains -or @($FreemailDomains).Count -eq 0) { $FreemailDomains = @(Get-TriageFreemailDomains) }
    $hasMdo = $false
    $mdoFlag = Get-TriageProp $Orca @('DefenderForOffice365')
    if ($mdoFlag -eq $true -or [string]$mdoFlag -eq 'True') { $hasMdo = $true }
    $fails = @(Get-TriageOrcaFailNames -Orca $Orca)
    $items = New-Object System.Collections.Generic.List[object]
    $blocks = @{}
    $allows = @{}
    foreach ($row in @(ConvertTo-TriageArray $Rows)) {
        if (-not $row) { continue }
        $sender = [string](Get-TriageProp $row @('SenderAddress')).Trim()
        if (-not $sender) { continue }
        $key = $sender.ToLowerInvariant()
        $subject = [string](Get-TriageProp $row @('Subject'))
        if (Test-TriageBlockCandidate -Row $row -FreemailDomains $FreemailDomains) {
            if (-not $blocks.ContainsKey($key)) {
                $blocks[$key] = [pscustomobject]@{ Sender = $sender; Count = 0; Subjects = (New-Object System.Collections.Generic.List[string]); Why = (New-Object System.Collections.Generic.List[string]) }
            }
            $blocks[$key].Count++
            if ($subject -and $blocks[$key].Subjects.Count -lt 3 -and -not $blocks[$key].Subjects.Contains($subject)) { $blocks[$key].Subjects.Add($subject) }
            $why = Get-TriageEvidenceText $row
            if ($why -and $blocks[$key].Why.Count -lt 2 -and -not $blocks[$key].Why.Contains($why)) { $blocks[$key].Why.Add($why) }
            continue
        }
        if (Test-TriageAllowCandidate -Row $row -FreemailDomains $FreemailDomains -InternalDomains $InternalDomains) {
            if (-not $allows.ContainsKey($key)) {
                $allows[$key] = [pscustomobject]@{ Sender = $sender; Count = 0; Subjects = (New-Object System.Collections.Generic.List[string]); Identities = (New-Object System.Collections.Generic.List[string]) }
            }
            $allows[$key].Count++
            if ($subject -and $allows[$key].Subjects.Count -lt 3 -and -not $allows[$key].Subjects.Contains($subject)) { $allows[$key].Subjects.Add($subject) }
            $qid = [string](Get-TriageProp $row @('Identity'))
            if ($qid -and -not $allows[$key].Identities.Contains($qid)) { $allows[$key].Identities.Add($qid) }
        }
    }
    foreach ($key in @($blocks.Keys | Sort-Object)) {
        $b = $blocks[$key]
        $lit = ConvertTo-TriagePsLiteral $b.Sender
        $why = @($b.Why.ToArray()) -join ' '
        $sample = @($b.Subjects.ToArray()) -join '; '
        $detail = "$($b.Count) delivered message(s). $why"
        if ($sample) { $detail = "$detail Example: $sample" }
        $script = "New-TenantAllowBlockListItems -ListType Sender -Block -Entries $lit -NoExpiration -Notes 'Email triage'"
        $items.Add((New-TriageRemediationItem -Id "block:$key" -Kind 'Block' -Title $b.Sender -Detail $detail.Trim() -License 'EOP' -Apply 'Worker' -Operation 'BlockSender' -Entry $b.Sender -Script $script))
    }
    foreach ($key in @($allows.Keys | Sort-Object)) {
        $a = $allows[$key]
        $lit = ConvertTo-TriagePsLiteral $a.Sender
        $lines = New-Object System.Collections.Generic.List[string]
        $lines.Add("New-TenantAllowBlockListItems -ListType Sender -Allow -Entries $lit -ExpirationDate (Get-Date).ToUniversalTime().AddDays(30) -Notes 'Email triage'")
        foreach ($qid in @($a.Identities.ToArray())) { $lines.Add("Release-QuarantineMessage -Identity $(ConvertTo-TriagePsLiteral $qid)") }
        $sample = @($a.Subjects.ToArray()) -join '; '
        $detail = "$($a.Count) held message(s) look legitimate."
        if ($sample) { $detail = "$detail Example: $sample" }
        $items.Add((New-TriageRemediationItem -Id "allow:$key" -Kind 'Allow' -Title $a.Sender -Detail $detail -License 'EOP' -Apply 'Worker' -Operation 'AllowSender' -Entry $a.Sender -ReleaseIdentities @($a.Identities.ToArray()) -Script ($lines -join "`r`n")))
    }
    function Add-ConfigItem([string]$Id, [string]$Title, [string]$Detail, [string]$License, [string]$Apply, [string]$Operation, [string]$Entry, [string]$RuleName, [string]$Script) {
        $useApply = $Apply
        $useScript = $Script
        $useDetail = $Detail
        if ($License -eq 'MDO' -and -not $hasMdo) {
            $useApply = 'Manual'
            $useDetail = "Needs Defender for Office 365. $Detail"
            $useScript = "# Defender for Office 365 was not detected, so the worker does not run this.`r`n" + $Script
        }
        $items.Add((New-TriageRemediationItem -Id $Id -Kind 'Config' -Title $Title -Detail $useDetail.Trim() -License $License -Apply $useApply -Operation $Operation -Entry $Entry -RuleName $RuleName -Script $useScript))
    }
    if ($fails -contains 'Outbound spam filter policy settings') {
        Add-ConfigItem 'config:outbound-spam' 'Outbound spam limits' 'ORCA: set the default outbound spam policy to 500 external recipients per hour, 1000 internal, and block the sender when the limit is hit.' 'EOP' 'Worker' 'SetOutboundSpam' '' '' @'
Set-HostedOutboundSpamFilterPolicy -Identity 'Default' -RecipientLimitExternalPerHour 500 -RecipientLimitInternalPerHour 1000 -ActionWhenThresholdReached BlockUser
'@
    }
    if ($fails -contains 'External Tags') {
        Add-ConfigItem 'config:external-tags' 'External sender tags' 'ORCA: Outlook is not tagging mail that comes from outside the organization.' 'EOP' 'Worker' 'SetExternalTags' '' '' 'Set-ExternalInOutlook -Enabled $true'
    }
    $domainAllow = $fails -contains 'Domain Allow Listing'
    foreach ($rule in @(ConvertTo-TriageArray $(if ($Config) { Get-TriageProp $Config @('TransportRules') }))) {
        if (-not $rule) { continue }
        $name = [string](Get-TriageProp $rule @('Name', 'Identity'))
        foreach ($raw in @(ConvertTo-TriageArray (Get-TriageProp $rule @('ExceptIfSenderDomainIs', 'Exceptions')))) {
            $dom = [string]$raw
            if (-not $dom) { continue }
            $wholeFree = Test-TriageWholeFreemailDomain -Value $dom -FreemailDomains $FreemailDomains
            if (-not $wholeFree -and -not $domainAllow) { continue }
            $litRule = ConvertTo-TriagePsLiteral $name
            $litDom = ConvertTo-TriagePsLiteral $dom
            $why = "Transport rule '$name' skips sender domain $dom."
            if ($wholeFree) { $why = "Transport rule '$name' exempts the whole freemail domain $dom." }
            Add-ConfigItem "config:rule:$name`:$dom" "Remove rule exception $dom" $why 'EOP' 'Worker' 'RemoveRuleException' $dom $name "Set-TransportRule -Identity $litRule -ExceptIfSenderDomainIs @{Remove=$litDom}"
        }
    }
    foreach ($tabl in @(ConvertTo-TriageArray $(if ($Config) { Get-TriageProp $Config @('TablItems') }))) {
        if (-not $tabl) { continue }
        $action = [string](Get-TriageProp $tabl @('Action'))
        if ($action -notmatch '(?i)^Allow$') { continue }
        $value = [string](Get-TriageProp $tabl @('Value'))
        if (-not (Test-TriageWholeFreemailDomain -Value $value -FreemailDomains $FreemailDomains)) { continue }
        $lit = ConvertTo-TriagePsLiteral $value
        Add-ConfigItem "config:tabl:$value" "Remove allow entry $value" "Tenant allow/block list allows the freemail domain $value." 'EOP' 'Worker' 'RemoveTablAllow' $value '' "Remove-TenantAllowBlockListItems -ListType Sender -Entries $lit"
    }
    $mxAway = $false
    $mxSeen = $false
    foreach ($mx in @(ConvertTo-TriageArray $(if ($Config) { Get-TriageProp $Config @('Mx') }))) {
        if (-not $mx) { continue }
        $mxSeen = $true
        $ex = @(ConvertTo-TriageArray (Get-TriageProp $mx @('Exchanges')))
        if (Test-TriageMxPointsAtEop -Exchanges $ex) { $mxAway = $false; break }
        $mxAway = $true
    }
    if ($mxSeen -and $mxAway) {
        foreach ($conn in @(ConvertTo-TriageArray $(if ($Config) { Get-TriageProp $Config @('InboundConnectors') }))) {
            if (-not $conn) { continue }
            $cname = [string](Get-TriageProp $conn @('Name', 'Identity'))
            if ($cname -notmatch '(?i)inky|proofpoint|barracuda') { continue }
            $ips = @(ConvertTo-TriageArray (Get-TriageProp $conn @('EFSkipIPs')) | Where-Object { $_ })
            $skip = Get-TriageProp $conn @('EFSkipLastIP')
            if (@($ips).Count -gt 0 -or $skip -eq $true -or [string]$skip -eq 'True') { continue }
            $lit = ConvertTo-TriagePsLiteral $cname
            Add-ConfigItem "config:connector:$cname" "Enhanced filtering on $cname" "Inbound connector '$cname' has no enhanced filtering, and MX does not point at Exchange Online Protection." 'EOP' 'Worker' 'SetConnectorFiltering' '' $cname "Set-InboundConnector -Identity $lit -EFSkipLastIP `$true"
        }
    }
    $policies = @(ConvertTo-TriageArray $(if ($Config) { Get-TriageProp $Config @('AntiPhishPolicies') }))
    $emptyPolicy = ''
    foreach ($pol in $policies) {
        if (-not $pol) { continue }
        $users = @(ConvertTo-TriageArray (Get-TriageProp $pol @('TargetedUsersToProtect')) | Where-Object { $_ })
        if (@($users).Count -gt 0) { continue }
        $pname = [string](Get-TriageProp $pol @('Name', 'Identity'))
        if ($pname -eq 'Office365 AntiPhish Default' -or -not $emptyPolicy) { $emptyPolicy = $pname }
    }
    if ($emptyPolicy) {
        $vipNote = ''
        if (@($VipNames).Count -gt 0) { $vipNote = ' Client notes list ' + (@($VipNames) -join ', ') + '.' }
        $lit = ConvertTo-TriagePsLiteral $emptyPolicy
        $script = @"
# TargetedUsersToProtect needs each person's primary SMTP address.
Set-AntiPhishPolicy -Identity $lit -EnableTargetedUserProtection `$true -TargetedUsersToProtect @{Add='exec@contoso.com'}
"@
        Add-ConfigItem 'config:protected-users' 'Protect executive mailboxes' "Anti-phish policy '$emptyPolicy' does not list protected users.$vipNote Replace the example address before running." 'MDO' 'Manual' '' '' '' $script.Trim()
    }
    $dkimDomains = New-Object System.Collections.Generic.List[string]
    foreach ($health in @(ConvertTo-TriageArray $(if ($Config) { Get-TriageProp $Config @('DomainHealth') }))) {
        if (-not $health) { continue }
        $name = [string](Get-TriageProp $health @('Name'))
        if (-not (Test-TriageCustomerDomain $name)) { continue }
        if ($fails -contains 'Signing Configuration') { $dkimDomains.Add($name) }
        $rec = [string](Get-TriageProp $health @('DmarcRecord'))
        $missing = [string]::IsNullOrWhiteSpace($rec) -or ($rec -match '(?i)no dmarc')
        $none = $rec -match '(?i)(^|;)\s*p=none(\s|;|$)'
        $noRua = -not $missing -and ($rec -notmatch '(?i)\brua=')
        if (-not $missing -and -not $none -and -not $noRua) { continue }
        $suggested = "v=DMARC1; p=quarantine; rua=mailto:dmarc@$name"
        $why = "DMARC for $name is missing. Publish a TXT record at _dmarc.$name."
        if ($none) { $why = "DMARC for $name is p=none. Move to p=quarantine after you recognize the reports." }
        elseif ($noRua) { $why = "DMARC for $name has no rua address, so nobody receives the reports." }
        $script = @"
# DNS TXT record. The worker cannot publish this.
# Host: _dmarc.$name
# Value: $suggested
"@
        Add-ConfigItem "config:dmarc:$name" "DMARC for $name" $why 'DNS' 'Manual' '' '' '' $script.Trim()
    }
    if (($fails -contains 'Signing Configuration') -and $dkimDomains.Count -gt 0) {
        $lines = New-Object System.Collections.Generic.List[string]
        $lines.Add('# Publish the CNAME records from Get-DkimSigningConfig before you enable signing.')
        $lines.Add('Get-DkimSigningConfig | Format-List Identity, Enabled, Status, Selector1CNAME, Selector2CNAME')
        foreach ($d in @($dkimDomains.ToArray())) {
            $lines.Add("Set-DkimSigningConfig -Identity $(ConvertTo-TriagePsLiteral $d) -Enabled `$true")
        }
        Add-ConfigItem 'config:dkim' 'DKIM signing' 'ORCA: DKIM is not signing one or more customer domains. Publish the selector CNAMEs, then enable signing.' 'DNS' 'Manual' '' '' '' ($lines -join "`r`n")
    }
    if ($fails -contains 'SPF Records') {
        Add-ConfigItem 'config:spf' 'SPF records' 'ORCA: at least one domain is missing an SPF record that authorizes Exchange Online.' 'DNS' 'Manual' '' '' '' "# DNS TXT on each customer domain that fails SPF.`r`n# v=spf1 include:spf.protection.outlook.com -all`r`n# If a filter vendor sits in front of Exchange, include that vendor's SPF mechanism too."
    }
    foreach ($pair in @(
            @('Safe Links Policy Rules', 'config:safe-links-policy', 'Safe Links does not cover every domain.'),
            @('Safe Attachments Policy Rules', 'config:safe-attachments-policy', 'Safe Attachments does not cover every domain.'),
            @('Safe Links Allow Listing', 'config:safe-links-allow', 'A mail flow rule bypasses Safe Links.'),
            @('Safe Attachments Allow listing', 'config:safe-attachments-allow', 'A mail flow rule bypasses Safe Attachments.')
        )) {
        if ($fails -notcontains $pair[0]) { continue }
        $script = @"
# $($pair[2]) Review the existing policy in the Defender portal before adding another one.
# Get-SafeLinksPolicy; Get-SafeLinksRule
# Get-SafeAttachmentPolicy; Get-SafeAttachmentRule
"@
        Add-ConfigItem $pair[1] $pair[0] $pair[2] 'MDO' 'Manual' '' '' '' $script.Trim()
    }
    return @($items.ToArray())
}

function Invoke-TriageRemediationAction {
    param($Item)
    $op = [string](Get-TriageProp $Item @('Operation'))
    $entry = [string](Get-TriageProp $Item @('Entry'))
    $rule = [string](Get-TriageProp $Item @('RuleName'))
    switch ($op) {
        'BlockSender' {
            if (-not (Get-Command New-TenantAllowBlockListItems -ErrorAction SilentlyContinue)) { throw 'New-TenantAllowBlockListItems is not in this session. Copy the script into an Exchange Online window.' }
            New-TenantAllowBlockListItems -ListType Sender -Block -Entries @($entry) -NoExpiration -Notes 'Email triage' -ErrorAction Stop | Out-Null
        }
        'AllowSender' {
            if (-not (Get-Command New-TenantAllowBlockListItems -ErrorAction SilentlyContinue)) { throw 'New-TenantAllowBlockListItems is not in this session. Copy the script into an Exchange Online window.' }
            New-TenantAllowBlockListItems -ListType Sender -Allow -Entries @($entry) -ExpirationDate (Get-Date).ToUniversalTime().AddDays(30) -Notes 'Email triage' -ErrorAction Stop | Out-Null
            if (-not (Get-Command Release-QuarantineMessage -ErrorAction SilentlyContinue)) { throw 'The allow entry was created, but Release-QuarantineMessage is not in this session.' }
            foreach ($qid in @(ConvertTo-TriageArray (Get-TriageProp $Item @('ReleaseIdentities')))) {
                if ($qid) { Release-QuarantineMessage -Identity ([string]$qid) -ErrorAction Stop | Out-Null }
            }
        }
        'SetOutboundSpam' {
            if (-not (Get-Command Set-HostedOutboundSpamFilterPolicy -ErrorAction SilentlyContinue)) { throw 'Set-HostedOutboundSpamFilterPolicy is not in this session. Copy the script into an Exchange Online window.' }
            Set-HostedOutboundSpamFilterPolicy -Identity 'Default' -RecipientLimitExternalPerHour 500 -RecipientLimitInternalPerHour 1000 -ActionWhenThresholdReached BlockUser -ErrorAction Stop | Out-Null
        }
        'SetExternalTags' {
            if (-not (Get-Command Set-ExternalInOutlook -ErrorAction SilentlyContinue)) { throw 'Set-ExternalInOutlook is not in this session. Copy the script into an Exchange Online window.' }
            Set-ExternalInOutlook -Enabled $true -ErrorAction Stop | Out-Null
        }
        'RemoveRuleException' {
            if (-not (Get-Command Set-TransportRule -ErrorAction SilentlyContinue)) { throw 'Set-TransportRule is not in this session. Copy the script into an Exchange Online window.' }
            Set-TransportRule -Identity $rule -ExceptIfSenderDomainIs @{ Remove = $entry } -ErrorAction Stop | Out-Null
        }
        'RemoveTablAllow' {
            if (-not (Get-Command Remove-TenantAllowBlockListItems -ErrorAction SilentlyContinue)) { throw 'Remove-TenantAllowBlockListItems is not in this session. Copy the script into an Exchange Online window.' }
            Remove-TenantAllowBlockListItems -ListType Sender -Entries @($entry) -ErrorAction Stop | Out-Null
        }
        'SetConnectorFiltering' {
            if (-not (Get-Command Set-InboundConnector -ErrorAction SilentlyContinue)) { throw 'Set-InboundConnector is not in this session. Copy the script into an Exchange Online window.' }
            Set-InboundConnector -Identity $rule -EFSkipLastIP $true -ErrorAction Stop | Out-Null
        }
        default { throw "This item is copy-only." }
    }
}

function Invoke-TriageRemediation {
    param(
        [Parameter(Mandatory)][string]$Folder,
        [string[]]$Ids = @(),
        [string[]]$VipNames = @()
    )
    $resolved = [IO.Path]::GetFullPath($Folder)
    $docs = [Environment]::GetFolderPath('MyDocuments')
    $root = [IO.Path]::GetFullPath((Join-Path (Join-Path $docs 'ExchangeOnlineAnalyzer') 'SecurityInvestigation'))
    if (-not $resolved.StartsWith($root, [StringComparison]::OrdinalIgnoreCase)) { throw 'Triage folder must be under Documents\ExchangeOnlineAnalyzer\SecurityInvestigation' }
    if ((Split-Path $resolved -Leaf) -notmatch '^Triage_') { throw 'Not a triage pack folder' }
    if (-not (Test-Path -LiteralPath (Join-Path $resolved 'Triage.csv'))) { throw 'This pack has no triage list yet.' }
    $view = Get-TriagePackView -Folder $resolved -VipNames $VipNames -Limit 500
    $map = @{}
    foreach ($item in @(ConvertTo-TriageArray (Get-TriageProp $view @('Actions')))) {
        if (-not $item) { continue }
        $id = [string](Get-TriageProp $item @('Id'))
        if ($id) { $map[$id] = $item }
    }
    $results = New-Object System.Collections.Generic.List[object]
    foreach ($id in @($Ids)) {
        $clean = [string]$id
        if (-not $clean) { continue }
        if (-not $map.ContainsKey($clean)) {
            $results.Add([pscustomobject]@{ Id = $clean; Status = 'Skipped'; Detail = 'That item is not in this pack.' })
            continue
        }
        $item = $map[$clean]
        if ([string](Get-TriageProp $item @('Apply')) -ne 'Worker') {
            $results.Add([pscustomobject]@{ Id = $clean; Status = 'Manual'; Detail = 'Copy the PowerShell. The worker does not apply this one.' })
            continue
        }
        try {
            Invoke-TriageRemediationAction -Item $item
            $results.Add([pscustomobject]@{ Id = $clean; Status = 'Applied'; Detail = '' })
        }
        catch {
            $results.Add([pscustomobject]@{ Id = $clean; Status = 'Failed'; Detail = ($_.Exception.Message -replace '[\r\n]+', ' ') })
        }
    }
    return @($results.ToArray())
}

function Invoke-TriageWorkerCommand {
    param(
        [Parameter(Mandatory)][string]$Command,
        [string]$CompanyName = '',
        [string]$ClientNumber = '',
        [string]$ScriptRoot = ''
    )
    try {
        $kind = (($Command -split '\|')[0]).Trim()
        $known = @('TRIAGE_QUARANTINE', 'TRIAGE_DELIVERY', 'TRIAGE_IMPERSONATION', 'TRIAGE_CONFIG', 'TRIAGE_ALL', 'TRIAGE_APPLY')
        if ($known -notcontains $kind) { return 'TRIAGE_FAILED:Unknown triage command' }
        if ($kind -eq 'TRIAGE_APPLY') {
            $existing = [string](Get-TriageCommandToken -Command $Command -Name 'FOLDER')
            if (-not $existing) { return 'TRIAGE_FAILED:Missing folder' }
            $idToken = [string](Get-TriageCommandToken -Command $Command -Name 'IDS')
            if (-not $idToken) { return 'TRIAGE_FAILED:Select at least one item' }
            $ids = @($idToken -split ';' | Where-Object { $_ })
            $vips = @()
            if ($ScriptRoot -and $ClientNumber) {
                try { $vips = @((Get-TriageClientNotes -NotesRoot (Join-Path $ScriptRoot 'ClientNotes') -Client $ClientNumber).vips) } catch { }
            }
            $applied = @(Invoke-TriageRemediation -Folder $existing -Ids $ids -VipNames $vips)
            $payload = @{ results = @($applied) } | ConvertTo-Json -Compress -Depth 5
            return "TRIAGE_APPLY:$payload"
        }
        if ($kind -eq 'TRIAGE_CONFIG' -or $kind -eq 'TRIAGE_ALL') {
            $modFail = Test-TriageConfigModuleInstalled
            if ($modFail) { return $modFail }
        }
        $days = 7
        $daysToken = Get-TriageCommandToken -Command $Command -Name 'DAYS'
        if ($daysToken) { $days = [int]$daysToken }
        $senderDomain = [string](Get-TriageCommandToken -Command $Command -Name 'SENDER_DOMAIN')
        $freemail = ((Get-TriageCommandToken -Command $Command -Name 'FREEMAIL') -eq '1')
        $topN = 5
        $topToken = Get-TriageCommandToken -Command $Command -Name 'HEADER_TOPN'
        if ($topToken) { $topN = [int]$topToken }
        $window = 120
        $windowToken = Get-TriageCommandToken -Command $Command -Name 'WINDOW'
        if ($windowToken) { $window = [int]$windowToken }
        $subject = [string](Get-TriageCommandToken -Command $Command -Name 'SUBJECT')
        $company = $CompanyName
        $companyToken = Get-TriageCommandToken -Command $Command -Name 'COMPANY'
        if ($companyToken) { $company = $companyToken }
        $existing = [string](Get-TriageCommandToken -Command $Command -Name 'FOLDER')
        $folder = New-TriagePackFolder -CompanyName $company -ExistingFolder $existing
        $notesRoot = $null
        if ($ScriptRoot) { $notesRoot = Join-Path $ScriptRoot 'ClientNotes' }
        $vips = @()
        $extra = @()
        if ($notesRoot -and $ClientNumber) {
            try {
                $notes = Get-TriageClientNotes -NotesRoot $notesRoot -Client $ClientNumber
                $vips = @($notes.vips)
                $extra = @($notes.extraFreemail)
            }
            catch { }
        }
        $freemailDomains = @(Get-TriageFreemailDomains -Extra $extra)
        $counts = @{ quarantine = 0; trace = 0; reached = 0 }
        $limited = $false
        if ($kind -eq 'TRIAGE_QUARANTINE' -or $kind -eq 'TRIAGE_ALL' -or $kind -eq 'TRIAGE_IMPERSONATION') {
            $counts.quarantine = Invoke-TriageQuarantinePull -Folder $folder -Days $days -SenderDomain $senderDomain -FreemailOnly:$freemail -HeaderTopN $topN -FreemailDomains $freemailDomains
        }
        if ($kind -eq 'TRIAGE_DELIVERY' -or $kind -eq 'TRIAGE_ALL' -or $kind -eq 'TRIAGE_IMPERSONATION') {
            $delivery = Invoke-TriageDeliveryCheck -Folder $folder -Days $days -SenderAddress $(if ($senderDomain) { "*@$senderDomain" } else { '' }) -Subject $subject -WindowSeconds $window -ScriptRoot $ScriptRoot
            $counts.trace = $delivery.Trace
            $counts.reached = $delivery.Reached
            $limited = [bool]$delivery.Limited
        }
        if ($kind -eq 'TRIAGE_CONFIG' -or $kind -eq 'TRIAGE_ALL') {
            Invoke-TriageConfigAudit -Folder $folder | Out-Null
        }
        else {
            Invoke-TriageSaveAcceptedDomains -Folder $folder | Out-Null
        }
        $internal = @()
        $accepted = Read-TriageJson -Path (Join-Path $folder 'AcceptedDomains.json')
        if ($accepted) { $internal = @($accepted.InternalDomains) }
        $view = Update-TriagePackOutputs -Folder $folder -VipNames $vips -InternalDomains $internal -ExtraFreemail $extra -Company $company
        $prefix = $kind.Substring('TRIAGE_'.Length)
        $payload = @{
            kind         = $prefix
            folder       = $folder
            quarantine   = $counts.quarantine
            trace        = $counts.trace
            reached      = $counts.reached
            findings     = @($view.Findings).Count
            rows         = @($view.Rows).Count
            traceLimited = $limited
            warning      = [string]$view.Warning
        } | ConvertTo-Json -Compress -Depth 5
        return ("TRIAGE_{0}:{1}" -f $prefix, $payload)
    }
    catch {
        $msg = $_.Exception.Message -replace '[\r\n]+', ' '
        if ($msg -like 'TRIAGE_FAILED:*') { return $msg }
        return "TRIAGE_FAILED:$msg"
    }
}

function Get-TriageAiDefaultModel {
    param([string]$Provider)
    switch ($Provider) {
        'chatgpt' { return 'gpt-4.1-mini' }
        'claude' { return 'claude-sonnet-5' }
        'hatz' { return 'gpt-5.2' }
        default { return '' }
    }
}

function Get-TriageAiKeyHint {
    param([string]$ApiKey)
    $key = [string]$ApiKey
    if ($key.Length -lt 4) { return '' }
    return ('····' + $key.Substring($key.Length - 4))
}

function Read-TriageAiSettingsFile {
    param([Parameter(Mandatory)][string]$Path)
    $stored = [ordered]@{ provider = ''; model = ''; apiKey = '' }
    if (-not (Test-Path -LiteralPath $Path)) { return $stored }
    $doc = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop | ConvertFrom-Json
    $provider = [string](Get-TriageProp $doc @('provider', 'Provider'))
    $provider = $provider.Trim().ToLowerInvariant()
    if ($provider -eq 'off') { $provider = '' }
    if ($provider -and $provider -notin @('chatgpt', 'claude', 'hatz')) { $provider = '' }
    $stored.provider = $provider
    $stored.model = ([string](Get-TriageProp $doc @('model', 'Model'))).Trim()
    $stored.apiKey = [string](Get-TriageProp $doc @('apiKey', 'ApiKey'))
    return $stored
}

function Get-TriageAiStaticModels {
    param([string]$Provider)
    switch ($Provider) {
        'chatgpt' {
            return @(
                [pscustomobject]@{ name = 'gpt-4.1-mini'; displayName = 'GPT-4.1 mini'; developer = 'OpenAI' }
                [pscustomobject]@{ name = 'gpt-4.1'; displayName = 'GPT-4.1'; developer = 'OpenAI' }
            )
        }
        'claude' {
            return @(
                [pscustomobject]@{ name = 'claude-sonnet-5'; displayName = 'Claude Sonnet 5'; developer = 'Anthropic' }
                [pscustomobject]@{ name = 'claude-haiku-4-5'; displayName = 'Claude Haiku 4.5'; developer = 'Anthropic' }
            )
        }
        default { return @() }
    }
}

function Get-TriageHatzModels {
    param([string]$ApiKey)
    $key = [string]$ApiKey
    if ([string]::IsNullOrWhiteSpace($key) -or $key -match '\s') { return @() }
    try {
        $resp = Invoke-RestMethod -Method Get -Uri 'https://ai.hatz.ai/v1/chat/models' -Headers @{ 'X-API-Key' = $key } -TimeoutSec 20
    }
    catch {
        return @()
    }
    $items = New-Object System.Collections.ArrayList
    $seen = @{}
    foreach ($row in @($resp.data)) {
        $name = [string](Get-TriageProp $row @('name', 'Name'))
        $name = $name.Trim()
        if (-not $name -or $seen.ContainsKey($name)) { continue }
        $seen[$name] = $true
        $display = [string](Get-TriageProp $row @('display_name', 'displayName', 'DisplayName'))
        $developer = [string](Get-TriageProp $row @('developer', 'Developer'))
        [void]$items.Add([pscustomobject]@{
            name        = $name
            displayName = $(if ($display.Trim()) { $display.Trim() } else { $name })
            developer   = $(if ($developer.Trim()) { $developer.Trim() } else { 'Other' })
        })
    }
    if (-not $seen.ContainsKey('auto')) {
        [void]$items.Insert(0, [pscustomobject]@{ name = 'auto'; displayName = 'Auto'; developer = 'Hatz' })
    }
    return @($items.ToArray())
}

function Get-TriageAiModelChoices {
    param($Stored)
    $provider = [string]$Stored.provider
    if ($provider -eq 'hatz') {
        $live = @(Get-TriageHatzModels -ApiKey ([string]$Stored.apiKey))
        if ($live.Count) { return $live }
        return @(
            [pscustomobject]@{ name = 'gpt-5.2'; displayName = 'GPT-5.2'; developer = 'OpenAI' }
            [pscustomobject]@{ name = 'gpt-4o'; displayName = 'GPT-4o'; developer = 'OpenAI' }
            [pscustomobject]@{ name = 'auto'; displayName = 'Auto'; developer = 'Hatz' }
        )
    }
    return @(Get-TriageAiStaticModels -Provider $provider)
}

function ConvertTo-TriageAiSettingsPublic {
    param($Stored)
    $provider = [string]$Stored.provider
    $model = [string]$Stored.model
    if (-not $model -and $provider) { $model = Get-TriageAiDefaultModel -Provider $provider }
    $key = [string]$Stored.apiKey
    return [pscustomobject]@{
        provider = $provider
        model    = $model
        hasKey   = -not [string]::IsNullOrWhiteSpace($key)
        keyHint  = (Get-TriageAiKeyHint -ApiKey $key)
        models   = @(Get-TriageAiModelChoices -Stored $Stored)
    }
}

function Get-TriageAiSettings {
    param([Parameter(Mandatory)][string]$Path)
    return ConvertTo-TriageAiSettingsPublic -Stored (Read-TriageAiSettingsFile -Path $Path)
}

function Test-TriageAiKeyToken {
    param([string]$ApiKey)
    $key = [string]$ApiKey
    if ($key -match '\s') {
        throw 'Paste the API token only. The value has spaces in it, so it is not a key.'
    }
}

function Set-TriageAiSettings {
    param(
        [Parameter(Mandatory)][string]$Path,
        [string]$Provider = '',
        [string]$Model = '',
        [string]$ApiKey = ''
    )
    $provider = ([string]$Provider).Trim().ToLowerInvariant()
    if ($provider -eq 'off') { $provider = '' }
    if ($provider -and $provider -notin @('chatgpt', 'claude', 'hatz')) {
        throw 'Provider must be ChatGPT, Claude, or Hatz AI.'
    }
    $prior = Read-TriageAiSettingsFile -Path $Path
    $model = ([string]$Model).Trim()
    if (-not $model -and $provider) { $model = Get-TriageAiDefaultModel -Provider $provider }
    $key = [string]$ApiKey
    if (-not [string]::IsNullOrWhiteSpace($key)) {
        $key = $key.Trim()
        Test-TriageAiKeyToken -ApiKey $key
    }
    if ([string]::IsNullOrWhiteSpace($key)) { $key = [string]$prior.apiKey }
    Test-TriageAiKeyToken -ApiKey $key
    $dir = Split-Path -Parent $Path
    if ($dir -and -not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
    $payload = [ordered]@{
        provider = $provider
        model    = $model
        apiKey   = $key
    }
    $payload | ConvertTo-Json -Compress | Set-Content -LiteralPath $Path -Encoding UTF8
    return ConvertTo-TriageAiSettingsPublic -Stored $payload
}

function Format-TriageAiList {
    param($Value)
    if ($null -eq $Value) { return '' }
    if ($Value -is [string]) { return $Value.Trim() }
    $parts = New-Object System.Collections.Generic.List[string]
    foreach ($item in $Value) {
        if ($null -eq $item) { continue }
        $text = ([string]$item).Trim()
        if ($text) { [void]$parts.Add($text) }
    }
    return ($parts -join '; ')
}

function Get-TriageAiBatchSkipKeys {
    param($Rows)
    $records = New-Object System.Collections.Generic.List[object]
    $i = 0
    foreach ($row in @(ConvertTo-TriageArray $Rows)) {
        if (-not $row) { continue }
        $key = [string](Get-TriageProp $row @('RowKey'))
        if (-not $key) { $key = "row$i" }
        $i++
        $sender = ([string](Get-TriageProp $row @('SenderAddress'))).Trim().ToLowerInvariant()
        $subject = ([string](Get-TriageProp $row @('Subject'))).Trim().ToLowerInvariant()
        $verdict = [string](Get-TriageProp $row @('Verdict'))
        $imp = [string](Get-TriageProp $row @('Impersonation'))
        $records.Add([pscustomobject]@{
                Key         = $key
                Sender      = $sender
                Subject     = $subject
                Interesting = (($verdict -eq 'Malicious') -or -not [string]::IsNullOrWhiteSpace($imp))
            })
    }
    $bySender = @{}
    foreach ($rec in $records) {
        if (-not $rec.Sender) { continue }
        if (-not $bySender.ContainsKey($rec.Sender)) {
            $bySender[$rec.Sender] = New-Object System.Collections.Generic.List[object]
        }
        $bySender[$rec.Sender].Add($rec)
    }
    $skip = @{}
    foreach ($sender in @($bySender.Keys)) {
        $items = $bySender[$sender]
        $boringCount = 0
        foreach ($rec in $items) { if (-not $rec.Interesting) { $boringCount++ } }
        if ($boringCount -ge 5 -and $boringCount -eq $items.Count) {
            foreach ($rec in $items) { $skip[$rec.Key] = $true }
            continue
        }
        $bySubject = @{}
        foreach ($rec in $items) {
            if (-not $rec.Subject) { continue }
            if (-not $bySubject.ContainsKey($rec.Subject)) {
                $bySubject[$rec.Subject] = New-Object System.Collections.Generic.List[object]
            }
            $bySubject[$rec.Subject].Add($rec)
        }
        foreach ($subject in @($bySubject.Keys)) {
            $group = $bySubject[$subject]
            $interesting = $false
            foreach ($rec in $group) { if ($rec.Interesting) { $interesting = $true; break } }
            if ($interesting -or $group.Count -lt 3) { continue }
            foreach ($rec in $group) { $skip[$rec.Key] = $true }
        }
    }
    return $skip
}

function Test-TriageAiRowIncluded {
    param($Row, $SkipKeys, [string]$RowKey = '')
    $flag = Get-TriageProp $Row @('AiReview')
    if ($flag -is [bool]) { return [bool]$flag }
    $text = [string]$flag
    if ($text -eq 'True') { return $true }
    if ($text -eq 'False') { return $false }
    $verdict = [string](Get-TriageProp $Row @('Verdict'))
    $imp = [string](Get-TriageProp $Row @('Impersonation'))
    if ($verdict -eq 'Malicious' -or -not [string]::IsNullOrWhiteSpace($imp)) { return $true }
    if (-not $RowKey) { $RowKey = [string](Get-TriageProp $Row @('RowKey')) }
    if ($RowKey -and $SkipKeys -and $SkipKeys.ContainsKey($RowKey)) { return $false }
    return $verdict -eq 'Unknown'
}

function New-TriageAiReviewBrief {
    param(
        $View,
        [int]$MaxRows = 60,
        [int]$MaxChars = 20000
    )
    if (-not $View) { throw 'Nothing to review.' }
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('Email triage pack. Local rules assigned these verdicts. They miss lookalikes and mixed cases, so treat the verdict as a starting point.')
    $warning = [string](Get-TriageProp $View @('Warning'))
    if ($warning.Trim()) { [void]$sb.AppendLine(('Pull warning: {0}' -f $warning.Trim())) }
    $rows = @(Get-TriageProp $View @('Rows'))
    $folder = [string](Get-TriageProp $View @('Folder'))
    if ($folder -and (Test-Path -LiteralPath (Join-Path $folder 'Triage.csv'))) {
        $full = Read-TriageSavedPackView -Folder $folder -Limit 0
        if ($full) { $rows = @(Get-TriageProp $full @('Rows')) }
    }
    $shown = 0
    $leftOut = 0
    $capped = 0
    $skipKeys = Get-TriageAiBatchSkipKeys -Rows $rows
    $batchLeft = @{}
    $leftSenders = @{}
    [void]$sb.AppendLine('Messages selected for review:')
    $index = 0
    foreach ($row in $rows) {
        if (-not $row) { continue }
        $key = [string](Get-TriageProp $row @('RowKey'))
        if (-not $key) { $key = "row$index" }
        $index++
        if (-not (Test-TriageAiRowIncluded -Row $row -SkipKeys $skipKeys -RowKey $key)) {
            $leftOut++
            $sender = ([string](Get-TriageProp $row @('SenderAddress'))).Trim().ToLowerInvariant()
            if ($sender) {
                $leftSenders[$sender] = $true
                if ($skipKeys.ContainsKey($key)) {
                    if (-not $batchLeft.ContainsKey($sender)) { $batchLeft[$sender] = 0 }
                    $batchLeft[$sender]++
                }
            }
            continue
        }
        if ($shown -ge $MaxRows) { $capped++; continue }
        $evidence = Format-TriageAiList -Value (Get-TriageProp $row @('Evidence'))
        if ($evidence.Length -gt 400) { $evidence = $evidence.Substring(0, 400) }
        $line = '{0} (suggested {1}) | {2} -> {3} | {4} | {5} | mailbox={6} | {7} | {8}' -f `
            (Get-TriageProp $row @('Verdict')), `
            (Get-TriageProp $row @('SuggestedVerdict')), `
            (Get-TriageProp $row @('SenderAddress')), `
            (Get-TriageProp $row @('RecipientAddress')), `
            (Get-TriageProp $row @('Subject')), `
            $evidence, `
            (Get-TriageProp $row @('ReachedMailbox')), `
            (Get-TriageProp $row @('Outcome')), `
            (Get-TriageProp $row @('Impersonation'))
        [void]$sb.AppendLine($line.Trim())
        $shown++
    }
    if ($leftOut) {
        [void]$sb.AppendLine(('({0} messages were left out of this review.)' -f $leftOut))
    }
    if ($batchLeft.Count) {
        $labels = New-Object System.Collections.Generic.List[string]
        foreach ($sender in @($batchLeft.Keys | Sort-Object { $batchLeft[$_] } -Descending)) {
            if ($labels.Count -ge 12) { break }
            [void]$labels.Add(('{0} ({1})' -f $sender, $batchLeft[$sender]))
        }
        $line = 'Repeated mail left out because it was not marked malicious: ' + ($labels -join ', ')
        $moreSenders = $batchLeft.Count - $labels.Count
        if ($moreSenders -gt 0) { $line += (', and {0} more senders' -f $moreSenders) }
        [void]$sb.AppendLine($line + '.')
    }
    if ($capped) {
        [void]$sb.AppendLine(('({0} more included messages were not sent because the review is capped at {1}.)' -f $capped, $MaxRows))
    }
    $findings = @(Get-TriageProp $View @('Findings'))
    if ($findings.Count) {
        [void]$sb.AppendLine('Config findings:')
        foreach ($finding in $findings) {
            [void]$sb.AppendLine(('- {0}: {1}' -f (Get-TriageProp $finding @('Severity')), (Get-TriageProp $finding @('Detail'))))
        }
    }
    $orca = Get-TriageProp $View @('Orca')
    if ($orca) {
        $label = [string](Get-TriageProp $orca @('ServiceLabel'))
        if ($label) { [void]$sb.AppendLine(('Protection service: {0}' -f $label)) }
        $failShown = 0
        foreach ($group in @(Get-TriageProp $orca @('Groups'))) {
            foreach ($item in @(Get-TriageProp $group @('Items'))) {
                $result = [string](Get-TriageProp $item @('Result'))
                if ($result -notmatch '(?i)fail') { continue }
                if ($failShown -ge 20) { break }
                [void]$sb.AppendLine(('- ORCA {0} / {1}: {2}' -f (Get-TriageProp $item @('Area')), (Get-TriageProp $item @('Name')), (Get-TriageProp $item @('FailRecommendation'))))
                $failShown++
            }
        }
    }
    $actions = @(Get-TriageProp $View @('Actions'))
    if ($actions.Count) {
        $actionLines = New-Object System.Collections.Generic.List[string]
        foreach ($action in $actions) {
            if ($actionLines.Count -ge 30) { break }
            $title = ([string](Get-TriageProp $action @('Title'))).Trim().ToLowerInvariant()
            if ($title -and $leftSenders.ContainsKey($title)) { continue }
            [void]$actionLines.Add(('- {0}: {1}: {2} ({3})' -f (Get-TriageProp $action @('Kind')), (Get-TriageProp $action @('Title')), (Get-TriageProp $action @('Detail')), (Get-TriageProp $action @('Apply'))))
        }
        if ($actionLines.Count) {
            [void]$sb.AppendLine('Suggested actions (not yet applied):')
            foreach ($line in $actionLines) { [void]$sb.AppendLine($line) }
        }
    }
    $text = $sb.ToString().Trim()
    if ($text.Length -gt $MaxChars) {
        $text = $text.Substring(0, $MaxChars) + "`n(Pack summary truncated.)"
    }
    return $text
}

function Get-TriageAiReviewRequest {
    param(
        [Parameter(Mandatory)][string]$Provider,
        [Parameter(Mandatory)][string]$Model,
        [Parameter(Mandatory)][string]$ApiKey,
        [Parameter(Mandatory)][string]$Brief,
        [string]$SystemPrompt = '',
        [int]$MaxTokens = 1200,
        [switch]$Stream
    )
    $system = @'
You review an email-security triage pack for an analyst. Use only the facts in the pack. For the important messages, say whether you agree with the verdict and why. Then list what to block, what to release, and which configuration items to fix first. Call out anything uncertain. Do not invent message bodies, headers, recipients, or incidents that are not in the pack. The reply must be plaintext, not Markdown. Do not use HTML, asterisks for emphasis, hash headings, or code fences. Never use an en dash, an em dash, or a double dash. Use a comma, a colon, parentheses, or a single hyphen.
'@
    if (-not [string]::IsNullOrWhiteSpace($SystemPrompt)) { $system = $SystemPrompt }
    $providerName = $Provider.Trim().ToLowerInvariant()
    if ($providerName -eq 'chatgpt') {
        $messages = New-Object System.Collections.ArrayList
        [void]$messages.Add([ordered]@{ role = 'system'; content = $system.Trim() })
        [void]$messages.Add([ordered]@{ role = 'user'; content = $Brief })
        $body = [ordered]@{ model = $Model; temperature = 0.2; max_tokens = $MaxTokens; stream = [bool]$Stream; messages = $messages }
        return [pscustomobject]@{
            Uri     = 'https://api.openai.com/v1/chat/completions'
            Headers = @{ Authorization = ('Bearer ' + $ApiKey) }
            Body    = ($body | ConvertTo-Json -Depth 6 -Compress)
        }
    }
    if ($providerName -eq 'claude') {
        $messages = New-Object System.Collections.ArrayList
        [void]$messages.Add([ordered]@{ role = 'user'; content = $Brief })
        $body = [ordered]@{ model = $Model; max_tokens = $MaxTokens; stream = [bool]$Stream; system = $system.Trim(); messages = $messages }
        return [pscustomobject]@{
            Uri     = 'https://api.anthropic.com/v1/messages'
            Headers = @{ 'x-api-key' = $ApiKey; 'anthropic-version' = '2023-06-01' }
            Body    = ($body | ConvertTo-Json -Depth 6 -Compress)
        }
    }
    if ($providerName -eq 'hatz') {
        $messages = New-Object System.Collections.ArrayList
        [void]$messages.Add([ordered]@{ role = 'system'; content = $system.Trim() })
        [void]$messages.Add([ordered]@{ role = 'user'; content = $Brief })
        $body = [ordered]@{ model = $Model; stream = [bool]$Stream; max_tokens = $MaxTokens; messages = $messages }
        return [pscustomobject]@{
            Uri     = 'https://ai.hatz.ai/v1/chat/completions'
            Headers = @{ 'X-API-Key' = $ApiKey }
            Body    = ($body | ConvertTo-Json -Depth 6 -Compress)
        }
    }
    throw 'Choose ChatGPT, Claude, or Hatz AI in AI review settings.'
}

function ConvertFrom-TriageAiResponse {
    param($Response, [string]$Provider)
    if (-not $Response) { return '' }
    if ($Provider -eq 'chatgpt' -or $Provider -eq 'hatz') {
        $choice = @($Response.choices) | Select-Object -First 1
        if ($choice -and $choice.message) { return [string]$choice.message.content }
        return ''
    }
    $block = @($Response.content) | Select-Object -First 1
    if ($block -and $block.text) { return [string]$block.text }
    return ''
}

function Get-TriageEvidenceRules {
    return @'
Never conclude true positive without concrete supporting evidence in the pulled data. An alert, a suspicion, a failed sign-in, or an empty result is not that evidence. When the evidence is inconclusive, do not declare a true positive. Ask the client to confirm whether the activity was theirs or authorized. The verdict and the client email should ask the client first.
Never recommend or suggest a password reset, a password change, or signing the user out everywhere in place of a password reset, unless specific evidence supports compromise. That evidence is a malicious sign-in, a malicious inbox rule, or confirmed credential exposure. Failed attempts alone are not enough. A password found in a breach list is confirmed credential exposure.
'@.Trim()
}

function Get-TriageBecVerdict {
    param([string]$Text)
    $match = [regex]::Match([string]$Text, '(?im)^\s*VERDICT:\s*(true positive|false positive|authorized activity|ask the client)\b')
    if (-not $match.Success) { return '' }
    return $match.Groups[1].Value.ToLowerInvariant()
}

function ConvertFrom-TriageAiStreamDelta {
    param(
        [string]$Provider,
        [string]$Data
    )
    $data = ([string]$Data).Trim()
    if ([string]::IsNullOrWhiteSpace($data) -or $data -eq '[DONE]') { return '' }
    try { $obj = $data | ConvertFrom-Json } catch { return '' }
    if ($Provider -eq 'claude') {
        $kind = [string]$obj.type
        if ($kind -ne 'content_block_delta' -or -not $obj.delta) { return '' }
        if ([string]$obj.delta.type -ne 'text_delta') { return '' }
        return [string]$obj.delta.text
    }
    $choice = @($obj.choices) | Select-Object -First 1
    if ($choice -and $choice.delta) { return [string]$choice.delta.content }
    return ''
}

function Invoke-TriageAiStream {
    param(
        [Parameter(Mandatory)]$Request,
        [Parameter(Mandatory)][string]$Provider,
        [scriptblock]$OnDelta
    )
    $http = [System.Net.HttpWebRequest]::Create([string]$Request.Uri)
    $http.Method = 'POST'
    $http.ContentType = 'application/json'
    $http.Accept = 'text/event-stream'
    $http.Timeout = 300000
    $http.ReadWriteTimeout = 300000
    foreach ($name in @($Request.Headers.Keys)) {
        $http.Headers[[string]$name] = [string]$Request.Headers[$name]
    }
    $payload = [System.Text.Encoding]::UTF8.GetBytes([string]$Request.Body)
    $http.ContentLength = $payload.Length
    $reqStream = $http.GetRequestStream()
    try { $reqStream.Write($payload, 0, $payload.Length) }
    finally { $reqStream.Close() }
    $response = $http.GetResponse()
    $reader = New-Object System.IO.StreamReader($response.GetResponseStream())
    $full = New-Object System.Text.StringBuilder
    try {
        while (-not $reader.EndOfStream) {
            $line = $reader.ReadLine()
            if ([string]::IsNullOrWhiteSpace($line)) { continue }
            if (-not $line.StartsWith('data:')) { continue }
            $piece = ConvertFrom-TriageAiStreamDelta -Provider $Provider -Data $line.Substring(5)
            if (-not $piece) { continue }
            [void]$full.Append($piece)
            if ($OnDelta) { & $OnDelta $piece }
        }
    }
    finally {
        $reader.Dispose()
        $response.Close()
    }
    return $full.ToString()
}

function Get-TriageAiErrorMessage {
    param($ErrorRecord)
    $status = 0
    try { $status = [int]$ErrorRecord.Exception.Response.StatusCode } catch { }
    if ($status -eq 401 -or $status -eq 403) {
        return 'The API key was rejected. Check the provider and key in AI review settings.'
    }
    $detail = ''
    if ($ErrorRecord.ErrorDetails -and $ErrorRecord.ErrorDetails.Message) {
        $detail = [string]$ErrorRecord.ErrorDetails.Message
    }
    if (-not $detail) { $detail = [string]$ErrorRecord.Exception.Message }
    $detail = $detail -replace '(?i)sk-ant-[A-Za-z0-9_\-]+', 'sk-ant-…'
    $detail = $detail -replace '(?i)sk-[A-Za-z0-9_\-]{8,}', 'sk-…'
    if ($detail.Length -gt 500) { $detail = $detail.Substring(0, 500) }
    return ('AI review failed: {0}' -f $detail)
}

function ConvertTo-TriageAiPlainText {
    param([string]$Text)
    $text = [string]$Text
    if ([string]::IsNullOrWhiteSpace($text)) { return '' }
    $text = $text -replace '(?m)^```[A-Za-z0-9_-]*\s*$', ''
    $text = $text -replace '```', ''
    $text = $text -replace '`+', ''
    $text = $text -replace '\[([^\]]+)\]\([^)]+\)', '$1'
    $text = $text -replace '\*\*([^*]+)\*\*', '$1'
    $text = $text -replace '__([^_\r\n]+)__', '$1'
    $text = $text -replace '(?m)^#{1,6}\s+', ''
    $text = $text -replace '(?m)^(\s*)[*+]\s+', '${1}- '
    $text = $text -replace '(?<![\w*])\*([^*\r\n]+)\*(?![\w*])', '$1'
    return $text.Trim()
}

function Invoke-TriageAiReview {
    param(
        [Parameter(Mandatory)][string]$SettingsPath,
        [Parameter(Mandatory)]$View
    )
    $stored = Read-TriageAiSettingsFile -Path $SettingsPath
    $provider = [string]$stored.provider
    if ($provider -notin @('chatgpt', 'claude', 'hatz')) {
        throw 'Choose ChatGPT, Claude, or Hatz AI in AI review settings.'
    }
    $key = [string]$stored.apiKey
    if ([string]::IsNullOrWhiteSpace($key)) {
        throw 'Save an API key in AI review settings. It stays on this PC.'
    }
    Test-TriageAiKeyToken -ApiKey $key
    $model = [string]$stored.model
    if (-not $model) { $model = Get-TriageAiDefaultModel -Provider $provider }
    $brief = New-TriageAiReviewBrief -View $View
    $request = Get-TriageAiReviewRequest -Provider $provider -Model $model -ApiKey $key -Brief $brief
    try {
        $response = Invoke-RestMethod -Method Post -Uri $request.Uri -Headers $request.Headers -Body $request.Body -ContentType 'application/json' -TimeoutSec 90
    }
    catch {
        throw (Get-TriageAiErrorMessage -ErrorRecord $_)
    }
    $text = (ConvertTo-TriageAiPlainText -Text (ConvertFrom-TriageAiResponse -Response $response -Provider $provider)).Trim()
    if (-not $text) { throw 'The model returned an empty review.' }
    return [pscustomobject]@{
        provider = $provider
        model    = $model
        text     = $text
    }
}

function Test-TriageBecRuleSuspicious {
    param($Rule)
    if (-not $Rule) { return $false }
    $hidden = Get-TriageProp $Rule @('Hidden', 'hidden')
    if ($hidden -is [bool] -and $hidden) { return $true }
    if ([string]$hidden -eq 'True') { return $true }
    foreach ($name in @('RedirectTo', 'redirectTo', 'ForwardTo', 'forwardTo', 'ForwardAsAttachmentTo', 'forwardAsAttachmentTo')) {
        $value = [string](Get-TriageProp $Rule @($name))
        if ($value.Trim()) { return $true }
    }
    $delete = Get-TriageProp $Rule @('DeleteMessage', 'deleteMessage')
    if ($delete -is [bool] -and $delete) { return $true }
    if ([string]$delete -eq 'True') { return $true }
    return $false
}

function Get-TriageBecRemediations {
    param($Facts)
    $items = New-Object System.Collections.Generic.List[object]
    $suspicious = New-Object System.Collections.Generic.List[string]
    foreach ($rule in @(ConvertTo-TriageArray (Get-TriageProp $Facts @('rules', 'Rules')))) {
        if (-not (Test-TriageBecRuleSuspicious -Rule $rule)) { continue }
        $name = [string](Get-TriageProp $rule @('name', 'Name', 'identity', 'Identity'))
        if (-not $name.Trim()) { $name = 'unnamed rule' }
        [void]$suspicious.Add($name.Trim())
    }
    if ($suspicious.Count) {
        $detail = ($suspicious -join ', ')
        if ($detail.Length -gt 300) { $detail = $detail.Substring(0, 300) }
        [void]$items.Add([pscustomobject]@{
                Id      = 'delete-rules'
                Title   = 'Delete inbox rules that hide or redirect mail'
                Detail  = $detail
                Kind    = 'delete-rule'
                Checked = $true
            })
    }
    $forwards = New-Object System.Collections.Generic.List[string]
    foreach ($row in @(ConvertTo-TriageArray (Get-TriageProp $Facts @('forwards', 'Forwards')))) {
        $address = [string](Get-TriageProp $row @('address', 'Address', 'target', 'Target'))
        if ($address.Trim()) { [void]$forwards.Add($address.Trim()) }
    }
    if ($forwards.Count) {
        [void]$items.Add([pscustomobject]@{
                Id      = 'remove-forward'
                Title   = 'Remove mailbox forwarding'
                Detail  = ($forwards -join ', ')
                Kind    = 'remove-forward'
                Checked = $true
            })
    }
    $delegates = New-Object System.Collections.Generic.List[string]
    foreach ($row in @(ConvertTo-TriageArray (Get-TriageProp $Facts @('delegates', 'Delegates')))) {
        $user = [string](Get-TriageProp $row @('user', 'User', 'target', 'Target'))
        $right = [string](Get-TriageProp $row @('right', 'Right'))
        if (-not $user.Trim()) { continue }
        $label = $user.Trim()
        if ($right.Trim()) { $label = '{0} ({1})' -f $label, $right.Trim() }
        [void]$delegates.Add($label)
    }
    if ($delegates.Count) {
        [void]$items.Add([pscustomobject]@{
                Id      = 'remove-delegates'
                Title   = 'Remove mailbox delegates'
                Detail  = ($delegates -join ', ')
                Kind    = 'remove-delegate'
                Checked = $true
            })
    }
    if ($items.Count) {
        [void]$items.Add([pscustomobject]@{
                Id      = 'enable-hold'
                Title   = 'Turn on litigation hold and mailbox audit'
                Detail  = 'Preserves mail before the other cleanup runs.'
                Kind    = 'enable-hold'
                Checked = $true
            })
        [void]$items.Add([pscustomobject]@{
                Id      = 'block'
                Title   = 'Block sign-in'
                Detail  = 'Off until you check it. The user cannot sign in until you unblock them.'
                Kind    = 'block'
                Checked = $false
            })
    }
    if ($suspicious.Count -or $forwards.Count) {
        [void]$items.Add([pscustomobject]@{
                Id      = 'revoke'
                Title   = 'Revoke sessions'
                Detail  = 'Off until you check it. Signs the user out everywhere. Shown because a suspicious inbox rule or mailbox forward is in the pulled data.'
                Kind    = 'revoke'
                Checked = $false
            })
        [void]$items.Add([pscustomobject]@{
                Id      = 'reset-password'
                Title   = 'Reset to a random password'
                Detail  = 'Off until you check it. The new password is not shown. The user sets their own at the reset page. Shown because a suspicious inbox rule or mailbox forward is in the pulled data.'
                Kind    = 'reset-password'
                Checked = $false
            })
    }
    return $items.ToArray()
}

function Get-TriageBecWindow {
    param(
        [string]$Ticket = '',
        [datetime]$Now = (Get-Date),
        [int]$MinimumDays = 2,
        [int]$MaximumDays = 90
    )
    if ($MinimumDays -lt 1) { $MinimumDays = 1 }
    if ($MaximumDays -lt $MinimumDays) { $MaximumDays = $MinimumDays }
    $eventUtc = $null
    $hadTime = $false
    $source = ''
    $iso = [regex]::Match([string]$Ticket, '\d{4}-\d{2}-\d{2}[T ]\d{2}:\d{2}(?::\d{2})?(?:Z|[+-]\d{2}:?\d{2})?')
    if ($iso.Success) {
        $parsed = [datetimeoffset]::MinValue
        if ([datetimeoffset]::TryParse($iso.Value.Replace(' ', 'T'), [ref]$parsed)) {
            $eventUtc = $parsed.UtcDateTime
            $hadTime = $true
            $source = $iso.Value
        }
    }
    if (-not $hadTime) {
        $us = [regex]::Match([string]$Ticket, '\b(\d{1,2})/(\d{1,2})/(\d{4})(?:\s+(\d{1,2}):(\d{2})\s*([AaPp][Mm])?)?')
        if ($us.Success) {
            $hour = 0
            $minute = 0
            if ($us.Groups[4].Success) {
                $hour = [int]$us.Groups[4].Value
                $minute = [int]$us.Groups[5].Value
                $ampm = $us.Groups[6].Value
                if ($ampm -match '^[Pp]' -and $hour -lt 12) { $hour += 12 }
                if ($ampm -match '^[Aa]' -and $hour -eq 12) { $hour = 0 }
            }
            $local = Get-Date -Year ([int]$us.Groups[3].Value) -Month ([int]$us.Groups[1].Value) -Day ([int]$us.Groups[2].Value) -Hour $hour -Minute $minute -Second 0
            $eventUtc = $local.ToUniversalTime()
            $hadTime = $true
            $source = $us.Value.Trim()
        }
    }
    $endUtc = $Now.ToUniversalTime()
    if ($hadTime) {
        $padded = $eventUtc.AddDays(-1)
        if ($padded -lt $endUtc) {
            $spanDays = [int][Math]::Ceiling(($endUtc - $padded).TotalDays)
        } else {
            $spanDays = $MinimumDays
        }
    } else {
        $spanDays = $MinimumDays
    }
    if ($spanDays -lt $MinimumDays) { $spanDays = $MinimumDays }
    if ($spanDays -gt $MaximumDays) { $spanDays = $MaximumDays }
    $startUtc = $endUtc.AddDays(-$spanDays)
    $tz = Get-TriageLocalZoneLabel
    $startLocal = $startUtc.ToLocalTime().ToString('yyyy-MM-dd h:mm tt')
    $endLocal = $endUtc.ToLocalTime().ToString('yyyy-MM-dd h:mm tt')
    $note = if ($hadTime) {
        "Ticket time '$source' was read as $tz. Search window is $startLocal through $endLocal $tz ($spanDays days). Log times below are $tz."
    } else {
        "No event time was found on the ticket. Search window is the last $spanDays days, $startLocal through $endLocal $tz. Log times below are $tz."
    }
    return [pscustomobject]@{
        StartUtc = $startUtc
        EndUtc   = $endUtc
        Days     = $spanDays
        HadTime  = $hadTime
        Note     = $note
    }
}

function ConvertFrom-TriageBecRepull {
    param([string]$Text, [int]$MaximumDays = 90)
    $days = 0
    if ([string]$Text -match '(?im)^\s*REPULL:\s*(\d+)\b') {
        $days = [int]$Matches[1]
    }
    if ($days -lt 0) { $days = 0 }
    if ($days -gt $MaximumDays) { $days = $MaximumDays }
    return $days
}

function ConvertFrom-TriageBecRepullLogs {
    param([string]$Text)
    $allowed = @{
        trace = 'trace'; 'message trace' = 'trace'; messagetrace = 'trace'
        ual = 'ual'; audit = 'ual'; 'unified audit' = 'ual'
        sharing = 'sharing'; onedrive = 'sharing'
        files = 'files'; file = 'files'
        dlp = 'dlp'
        alerts = 'alerts'; alert = 'alerts'
        signins = 'signins'; 'sign-ins' = 'signins'; signin = 'signins'
        apps = 'apps'; app = 'apps'
    }
    $found = New-Object System.Collections.Generic.List[string]
    if ([string]$Text -notmatch '(?im)^\s*REPULL:\s*\d+\b[ \t]*([^\r\n]*)$') { return @() }
    $rest = $Matches[1].Trim()
    if (-not $rest) { return @() }
    $seen = @{}
    foreach ($part in ($rest -split '[,\s]+')) {
        $key = $part.Trim().ToLowerInvariant()
        if (-not $key -or -not $allowed.ContainsKey($key)) { continue }
        $name = $allowed[$key]
        if ($seen.ContainsKey($name)) { continue }
        $seen[$name] = $true
        [void]$found.Add($name)
    }
    return $found.ToArray()
}

function Test-TriageBecWindowGap {
    param([string]$Text)
    return [bool]([regex]::IsMatch([string]$Text, '(?i)before the [\w\s-]{0,40}window|outside the (current |search |review )?window|older than (the |this )?window|window we could|exceeds (the |a )?window|window (was |is )?too short|not in the (current |search |two day |2 day )?window'))
}

function Get-TriageBecGapLogs {
    param([string]$Text)
    $slice = [string]$Text
    $hit = [regex]::Match($slice, '(?i)before the [\w\s-]{0,40}window|outside the (current |search |review )?window|older than (the |this )?window|window we could|exceeds (the |a )?window|window (was |is )?too short|not in the (current |search |two day |2 day )?window')
    if ($hit.Success) {
        $start = [Math]::Max(0, $hit.Index - 400)
        $len = [Math]::Min(800, $slice.Length - $start)
        $slice = $slice.Substring($start, $len)
    }
    $found = New-Object System.Collections.Generic.List[string]
    $rules = @(
        @{ Pattern = '(?i)onedrive|sharepoint|shar(?:e|ing|eable)|anonymous link'; Name = 'sharing' }
        @{ Pattern = '(?i)file download|downloaded a file|file activity'; Name = 'files' }
        @{ Pattern = '(?i)message trace|outbound (?:mail|email)|email blast'; Name = 'trace' }
        @{ Pattern = '(?i)sign-?in'; Name = 'signins' }
        @{ Pattern = '(?i)\bdlp\b'; Name = 'dlp' }
        @{ Pattern = '(?i)security alert'; Name = 'alerts' }
        @{ Pattern = '(?i)app registration'; Name = 'apps' }
        @{ Pattern = '(?i)unified audit|inbox rule'; Name = 'ual' }
    )
    foreach ($rule in $rules) {
        if ($slice -match $rule.Pattern -and -not $found.Contains($rule.Name)) {
            [void]$found.Add($rule.Name)
        }
    }
    return $found.ToArray()
}

function Get-TriageBecFollowUpPlan {
    param(
        [string]$Text,
        [int]$CurrentDays = 0,
        [int]$MaximumDays = 90
    )
    if ($CurrentDays -lt 0) { $CurrentDays = 0 }
    if ($MaximumDays -lt 1) { $MaximumDays = 1 }
    if ($MaximumDays -gt 90) { $MaximumDays = 90 }
    $hasRepull = [string]$Text -match '(?im)^\s*REPULL:\s*\d+\b'
    $explicit = ConvertFrom-TriageBecRepull -Text $Text -MaximumDays $MaximumDays
    $logs = @(ConvertFrom-TriageBecRepullLogs -Text $Text)
    $days = 0
    if ($hasRepull) {
        if ($explicit -gt $CurrentDays) { $days = $explicit }
    }
    elseif ((Test-TriageBecWindowGap -Text $Text) -and $CurrentDays -lt $MaximumDays) {
        if ($CurrentDays -lt 30 -and $MaximumDays -ge 30) { $days = 30 }
        else { $days = $MaximumDays }
        if (-not $logs.Count) { $logs = @(Get-TriageBecGapLogs -Text $Text) }
    }
    if ($days -gt $MaximumDays) { $days = $MaximumDays }
    return [pscustomobject]@{
        Days = $days
        Logs = @($logs)
    }
}

function Test-TriageConfirmedIncident {
    param([string]$Text)
    $t = [regex]::Replace([string]$Text, '(?i)\b(unconfirmed|not\s+confirmed)\b', ' ')
    if ($t -match '(?i)\bconfirmed\b.{0,60}\b(bec|business email compromise|compromise|compromised|incident|breach|exfiltration|ransomware)\b') { return $true }
    if ($t -match '(?i)\b(bec|business email compromise|incident|compromise|breach|exfiltration|ransomware)\b.{0,40}\bconfirmed\b') { return $true }
    if ($t -match '(?i)\b(account|mailbox|user)\b.{0,40}\b(was |is |has been |been )?compromised\b') { return $true }
    if ($t -match '(?i)\bcompromised (account|mailbox|user|tenant)\b') { return $true }
    return $false
}

function Test-TriageInvestigatorAppText {
    param([string]$Text)
    return [string]$Text -match 'River Run Security Investigator'
}

function Add-TriageBecFactLines {
    param($Builder, [string]$Title, $Rows, [int]$Max = 25, [scriptblock]$Format)
    [void]$Builder.AppendLine($Title)
    $count = 0
    foreach ($row in @(ConvertTo-TriageArray $Rows)) {
        if (-not $row) { continue }
        if ($count -ge $Max) { break }
        $line = & $Format $row
        if ([string]::IsNullOrWhiteSpace([string]$line)) { continue }
        [void]$Builder.AppendLine(('- {0}' -f ([string]$line).Trim()))
        $count++
    }
    if (-not $count) { [void]$Builder.AppendLine('- none') }
}

function New-TriageBecReviewBrief {
    param($Facts)
    $mailbox = [string](Get-TriageProp $Facts @('mailbox', 'Mailbox'))
    if ([string]::IsNullOrWhiteSpace($mailbox)) { throw 'Missing mailbox.' }
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine(('Mailbox: {0}' -f $mailbox.Trim()))
    [void]$sb.AppendLine('Inbox rules:')
    $ruleCount = 0
    foreach ($rule in @(ConvertTo-TriageArray (Get-TriageProp $Facts @('rules', 'Rules')))) {
        if ($ruleCount -ge 40) { break }
        $name = [string](Get-TriageProp $rule @('name', 'Name'))
        $bits = New-Object System.Collections.Generic.List[string]
        if ((Get-TriageProp $rule @('hidden', 'Hidden')) -eq $true -or [string](Get-TriageProp $rule @('hidden', 'Hidden')) -eq 'True') { [void]$bits.Add('hidden') }
        foreach ($pair in @(
                @('redirectTo', 'RedirectTo', 'redirect'),
                @('forwardTo', 'ForwardTo', 'forward'),
                @('forwardAsAttachmentTo', 'ForwardAsAttachmentTo', 'forward as attachment')
            )) {
            $value = [string](Get-TriageProp $rule @($pair[0], $pair[1]))
            if ($value.Trim()) { [void]$bits.Add(('{0} {1}' -f $pair[2], $value.Trim())) }
        }
        if ((Get-TriageProp $rule @('deleteMessage', 'DeleteMessage')) -eq $true -or [string](Get-TriageProp $rule @('deleteMessage', 'DeleteMessage')) -eq 'True') {
            [void]$bits.Add('deletes messages')
        }
        $desc = [string](Get-TriageProp $rule @('description', 'Description'))
        if ($desc.Length -gt 180) { $desc = $desc.Substring(0, 180) }
        if ($desc.Trim()) { [void]$bits.Add($desc.Trim()) }
        $extra = if ($bits.Count) { $bits -join '; ' } else { 'no redirect, forward, or delete action listed' }
        [void]$sb.AppendLine(('- {0}: {1}' -f $(if ($name.Trim()) { $name.Trim() } else { 'unnamed rule' }), $extra))
        $ruleCount++
    }
    if (-not $ruleCount) { [void]$sb.AppendLine('- none') }
    [void]$sb.AppendLine('Forwarding:')
    $fwdCount = 0
    foreach ($row in @(ConvertTo-TriageArray (Get-TriageProp $Facts @('forwards', 'Forwards')))) {
        if ($fwdCount -ge 20) { break }
        $address = [string](Get-TriageProp $row @('address', 'Address'))
        if (-not $address.Trim()) { continue }
        $field = [string](Get-TriageProp $row @('field', 'Field'))
        [void]$sb.AppendLine(('- {0} {1}' -f $(if ($field.Trim()) { $field.Trim() } else { 'forward' }), $address.Trim()))
        $fwdCount++
    }
    if (-not $fwdCount) { [void]$sb.AppendLine('- none') }
    [void]$sb.AppendLine('Delegates:')
    $delCount = 0
    foreach ($row in @(ConvertTo-TriageArray (Get-TriageProp $Facts @('delegates', 'Delegates')))) {
        if ($delCount -ge 20) { break }
        $user = [string](Get-TriageProp $row @('user', 'User'))
        if (-not $user.Trim()) { continue }
        $right = [string](Get-TriageProp $row @('right', 'Right'))
        [void]$sb.AppendLine(('- {0} {1}' -f $(if ($right.Trim()) { $right.Trim() } else { 'delegate' }), $user.Trim()))
        $delCount++
    }
    if (-not $delCount) { [void]$sb.AppendLine('- none') }
    $windowNote = [string](Get-TriageProp $Facts @('windowNote', 'WindowNote'))
    if ($windowNote.Trim()) { [void]$sb.AppendLine($windowNote.Trim()) }
    $posture = Get-TriageProp $Facts @('mfaPosture', 'MfaPosture')
    if ($posture) {
        [void]$sb.AppendLine('MFA posture:')
        foreach ($pair in @(
                @('perUser', 'PerUser', 'Per-user MFA'),
                @('securityDefaults', 'SecurityDefaults', 'Security Defaults'),
                @('caRequiresMfa', 'CARequiresMfa', 'Conditional Access requires MFA'),
                @('covered', 'Covered', 'Covered by one of those')
            )) {
            $value = Get-TriageProp $posture @($pair[0], $pair[1])
            if ($null -eq $value -or [string]::IsNullOrWhiteSpace([string]$value)) { continue }
            [void]$sb.AppendLine(('- {0}: {1}' -f $pair[2], $value))
        }
        $defaultsOn = [string](Get-TriageProp $posture @('securityDefaults', 'SecurityDefaults'))
        if ($defaultsOn -match '(?i)^true$') {
            [void]$sb.AppendLine('- Security Defaults is on, so Microsoft already enforces MFA for this tenant. Conditional Access requires MFA being False is normal in that case and is not a gap. Do not say MFA is missing, off, or not required for every sign-in. Do not recommend turning MFA on, stronger or additional MFA, tightening sign-in security, or Conditional Access; a tenant on Security Defaults usually does not have the license for Conditional Access. Do not comment on how many methods are registered.')
        }
    }
    Add-TriageBecFactLines -Builder $sb -Title 'MFA methods:' -Rows (Get-TriageProp $Facts @('authMethods', 'AuthMethods')) -Format {
        param($row)
        $type = [string](Get-TriageProp $row @('type', 'Type'))
        $created = ConvertTo-TriagePresentedTime (Get-TriageProp $row @('created', 'Created'))
        $detail = [string](Get-TriageProp $row @('details', 'Details'))
        if (-not $type -and -not $detail) { return '' }
        return ('{0} created {1} {2}' -f $type, $(if ($created) { $created } else { 'unknown' }), $detail).Trim()
    }
    Add-TriageBecFactLines -Builder $sb -Title 'Registered devices:' -Rows (Get-TriageProp $Facts @('devices', 'Devices')) -Format {
        param($row)
        $name = [string](Get-TriageProp $row @('displayName', 'DisplayName'))
        $os = [string](Get-TriageProp $row @('operatingSystem', 'OperatingSystem'))
        $seen = ConvertTo-TriagePresentedTime (Get-TriageProp $row @('lastSignIn', 'LastSignIn'))
        if (-not $name -and -not $os) { return '' }
        return ('{0} {1} last sign-in {2}' -f $name, $os, $(if ($seen) { $seen } else { 'unknown' })).Trim()
    }
    Add-TriageBecFactLines -Builder $sb -Title 'Intune devices:' -Rows (Get-TriageProp $Facts @('intune', 'Intune')) -Format {
        param($row)
        $name = [string](Get-TriageProp $row @('deviceName', 'DeviceName', 'displayName', 'DisplayName'))
        $seen = ConvertTo-TriagePresentedTime (Get-TriageProp $row @('lastSync', 'LastSync', 'lastSyncDateTime'))
        if (-not $name -or $name -eq '(none)') { return '' }
        return ('{0} last sync {1}' -f $name, $(if ($seen) { $seen } else { 'unknown' }))
    }
    Add-TriageBecFactLines -Builder $sb -Title 'OAuth consents:' -Rows (Get-TriageProp $Facts @('oauth', 'Oauth')) -Format {
        param($row)
        $app = [string](Get-TriageProp $row @('app', 'App', 'displayName', 'DisplayName'))
        $scope = [string](Get-TriageProp $row @('scope', 'Scope', 'scopes', 'Scopes'))
        if (-not $app) { return '' }
        if ($scope.Length -gt 120) { $scope = $scope.Substring(0, 120) }
        return ('{0} {1}' -f $app, $scope).Trim()
    }
    Add-TriageBecFactLines -Builder $sb -Title 'Directory roles:' -Rows (Get-TriageProp $Facts @('roles', 'Roles')) -Format {
        param($row)
        $name = [string](Get-TriageProp $row @('name', 'Name', 'role', 'Role', 'displayName', 'DisplayName'))
        if (-not $name) { return '' }
        return $name
    }
    Add-TriageBecFactLines -Builder $sb -Title 'App registrations created in the window:' -Rows (Get-TriageProp $Facts @('apps', 'Apps')) -Format {
        param($row)
        $name = [string](Get-TriageProp $row @('displayName', 'DisplayName'))
        $created = ConvertTo-TriagePresentedTime (Get-TriageProp $row @('created', 'Created'))
        if (-not $name -or (Test-TriageInvestigatorAppText $name)) { return '' }
        return ('{0} created {1}' -f $name, $(if ($created) { $created } else { 'unknown' }))
    }
    Add-TriageBecFactLines -Builder $sb -Title 'Transport rules (new, undated, or mail-routing):' -Rows (Get-TriageProp $Facts @('transport', 'Transport')) -Max 40 -Format {
        param($row)
        $name = [string](Get-TriageProp $row @('name', 'Name'))
        $state = [string](Get-TriageProp $row @('state', 'State'))
        $when = ConvertTo-TriagePresentedTime (Get-TriageProp $row @('whenChanged', 'WhenChanged'))
        $desc = [string](Get-TriageProp $row @('description', 'Description'))
        if (-not $name) { return '' }
        if ($desc.Length -gt 180) { $desc = $desc.Substring(0, 180) }
        return ('{0} {1} changed {2} {3}' -f $name, $state, $(if ($when) { $when } else { 'unknown' }), $desc).Trim()
    }
    Add-TriageBecFactLines -Builder $sb -Title 'Connectors (new, undated, or enabled outbound):' -Rows (Get-TriageProp $Facts @('connectors', 'Connectors')) -Max 30 -Format {
        param($row)
        $name = [string](Get-TriageProp $row @('name', 'Name'))
        $direction = [string](Get-TriageProp $row @('direction', 'Direction'))
        $enabled = [string](Get-TriageProp $row @('enabled', 'Enabled'))
        $hosts = [string](Get-TriageProp $row @('hosts', 'Hosts', 'smartHosts', 'SmartHosts'))
        $domains = [string](Get-TriageProp $row @('domains', 'Domains'))
        $when = ConvertTo-TriagePresentedTime (Get-TriageProp $row @('whenCreated', 'WhenCreated'))
        if (-not $name) { return '' }
        return ('{0} {1} {2} created {3} hosts {4} domains {5}' -f $direction, $name, $enabled, $(if ($when) { $when } else { 'unknown' }), $hosts, $domains).Trim()
    }
    $auditFailed = $false
    foreach ($note in @(ConvertTo-TriageArray (Get-TriageProp $Facts @('pullNotes', 'PullNotes')))) {
        if ([string]$note -match 'Unknown remediation command|REMEDIATE_FAILED') { $auditFailed = $true; break }
    }
    if ($auditFailed) {
        [void]$sb.AppendLine('Message trace, unified audit, DLP, file activity, security alerts, and sign-ins were not collected. The audit command was not recognized. Do not describe the mailbox as inactive from those sources.')
    }
    else {
    $zoneLabel = Get-TriageLocalZoneLabel
    [void]$sb.AppendLine(('Log times are {0}.' -f $zoneLabel))
    Add-TriageBecFactLines -Builder $sb -Title ('Message trace ({0}):' -f $zoneLabel) -Rows (Get-TriageProp $Facts @('messageTrace', 'MessageTrace')) -Max 40 -Format {
        param($row)
        $when = ConvertTo-TriagePresentedTime (Get-TriageProp $row @('received', 'Received'))
        $sender = [string](Get-TriageProp $row @('sender', 'Sender', 'senderAddress', 'SenderAddress'))
        $recip = [string](Get-TriageProp $row @('recipient', 'Recipient', 'recipientAddress', 'RecipientAddress'))
        $subject = [string](Get-TriageProp $row @('subject', 'Subject'))
        $status = [string](Get-TriageProp $row @('status', 'Status'))
        if (-not $sender -and -not $recip) { return '' }
        if ($subject.Length -gt 80) { $subject = $subject.Substring(0, 80) }
        $size = [string](Get-TriageProp $row @('size', 'Size'))
        $sizeText = if ($size.Trim()) { "size $size" } else { '' }
        return ('{0} {1} -> {2} {3} {4} {5}' -f $when, $sender, $recip, $status, $subject, $sizeText).Trim()
    }
    $reported = @(ConvertTo-TriageArray (Get-TriageProp $Facts @('reportedMessages', 'ReportedMessages')))
    if ($reported.Count -gt 0) {
        [void]$sb.AppendLine('Reported messages opened from the mailbox:')
        $shown = 0
        foreach ($msg in $reported) {
            if ($shown -ge 3) { break }
            if (-not $msg) { continue }
            $when = ConvertTo-TriagePresentedTime (Get-TriageProp $msg @('received', 'Received'))
            $from = [string](Get-TriageProp $msg @('from', 'From'))
            $fromName = [string](Get-TriageProp $msg @('fromName', 'FromName'))
            $reply = [string](Get-TriageProp $msg @('replyTo', 'ReplyTo'))
            $subject = [string](Get-TriageProp $msg @('subject', 'Subject'))
            $auth = [string](Get-TriageProp $msg @('auth', 'Auth'))
            $preview = [string](Get-TriageProp $msg @('preview', 'Preview'))
            $who = $from
            if ($fromName.Trim() -and $fromName.Trim() -ne $from) { $who = ('{0} <{1}>' -f $fromName.Trim(), $from) }
            [void]$sb.AppendLine(('- {0} {1} subject {2}' -f $(if ($when) { $when } else { 'undated' }), $who, $subject).Trim())
            if ($reply.Trim()) { [void]$sb.AppendLine(('  Reply-To: {0}' -f $reply.Trim())) }
            if ($auth.Trim()) { [void]$sb.AppendLine(('  Authentication: {0}' -f $auth.Trim())) }
            $links = @(ConvertTo-TriageArray (Get-TriageProp $msg @('links', 'Links')) | ForEach-Object { [string]$_ } | Where-Object { $_.Trim() })
            if ($links.Count) { [void]$sb.AppendLine(('  Links: {0}' -f ($links -join ' '))) }
            else { [void]$sb.AppendLine('  Links: none') }
            $files = @(ConvertTo-TriageArray (Get-TriageProp $msg @('files', 'Files')) | ForEach-Object { [string]$_ } | Where-Object { $_.Trim() })
            if ($files.Count) { [void]$sb.AppendLine(('  Attachments: {0}' -f ($files -join '; '))) }
            else { [void]$sb.AppendLine('  Attachments: none') }
            if ($preview.Trim()) { [void]$sb.AppendLine(('  Body: {0}' -f ($preview -replace '\s+', ' ').Trim())) }
            $shown++
        }
    }
    Add-TriageBecFactLines -Builder $sb -Title ('Unified audit ({0}, not Exchange item aggregates):' -f $zoneLabel) -Rows (Get-TriageProp $Facts @('ual', 'Ual')) -Max 40 -Format {
        param($row)
        $when = ConvertTo-TriagePresentedTime (Get-TriageProp $row @('created', 'Created', 'creationDate', 'CreationDate'))
        $op = [string](Get-TriageProp $row @('operation', 'Operation', 'operations', 'Operations'))
        $user = [string](Get-TriageProp $row @('user', 'User'))
        $recordType = [string](Get-TriageProp $row @('recordType', 'RecordType'))
        $target = [string](Get-TriageProp $row @('target', 'Target'))
        if (-not $op) { return '' }
        if ((Test-TriageInvestigatorAppText $target) -or (Test-TriageInvestigatorAppText $user)) { return '' }
        if ($target.Length -gt 120) { $target = $target.Substring(0, 120) }
        return ('{0} {1} {2} {3} {4}' -f $when, $recordType, $op, $user, $target).Trim()
    }
    $sharingRows = Get-TriageProp $Facts @('sharing', 'Sharing')
    if ($null -ne $sharingRows) {
        Add-TriageBecFactLines -Builder $sb -Title ('Sharing and OneDrive ({0}):' -f $zoneLabel) -Rows $sharingRows -Max 25 -Format {
            param($row)
            $when = ConvertTo-TriagePresentedTime (Get-TriageProp $row @('created', 'Created'))
            $op = [string](Get-TriageProp $row @('operation', 'Operation'))
            $recordType = [string](Get-TriageProp $row @('recordType', 'RecordType'))
            $target = [string](Get-TriageProp $row @('target', 'Target'))
            $who = [string](Get-TriageProp $row @('user', 'User'))
            if (-not $op -and -not $target) { return '' }
            if ($target.Length -gt 120) { $target = $target.Substring(0, 120) }
            return ('{0} {1} {2} {3} {4}' -f $when, $recordType, $op, $who, $target).Trim()
        }
    }
    $fileRows = Get-TriageProp $Facts @('files', 'Files')
    if ($null -ne $fileRows) {
        Add-TriageBecFactLines -Builder $sb -Title ('File and sharing ({0}):' -f $zoneLabel) -Rows $fileRows -Max 20 -Format {
            param($row)
            $when = ConvertTo-TriagePresentedTime (Get-TriageProp $row @('created', 'Created'))
            $op = [string](Get-TriageProp $row @('operation', 'Operation'))
            $recordType = [string](Get-TriageProp $row @('recordType', 'RecordType'))
            $target = [string](Get-TriageProp $row @('target', 'Target'))
            if (-not $op -and -not $target) { return '' }
            if ($target.Length -gt 120) { $target = $target.Substring(0, 120) }
            return ('{0} {1} {2} {3}' -f $when, $recordType, $op, $target).Trim()
        }
    }
    $dlpRows = Get-TriageProp $Facts @('dlp', 'Dlp')
    if ($null -ne $dlpRows) {
        Add-TriageBecFactLines -Builder $sb -Title ('DLP violations ({0}):' -f $zoneLabel) -Rows $dlpRows -Max 15 -Format {
            param($row)
            $when = ConvertTo-TriagePresentedTime (Get-TriageProp $row @('created', 'Created'))
            $op = [string](Get-TriageProp $row @('operation', 'Operation'))
            $policy = [string](Get-TriageProp $row @('policy', 'Policy'))
            $target = [string](Get-TriageProp $row @('target', 'Target'))
            if (-not $op -and -not $policy -and -not $target) { return '' }
            return ('{0} {1} {2} {3}' -f $when, $op, $policy, $target).Trim()
        }
    }
    $alertRows = Get-TriageProp $Facts @('alerts', 'Alerts')
    if ($null -ne $alertRows) {
        Add-TriageBecFactLines -Builder $sb -Title ('Security alerts ({0}):' -f $zoneLabel) -Rows $alertRows -Max 15 -Format {
            param($row)
            $when = ConvertTo-TriagePresentedTime (Get-TriageProp $row @('created', 'Created'))
            $title = [string](Get-TriageProp $row @('title', 'Title'))
            $severity = [string](Get-TriageProp $row @('severity', 'Severity'))
            $category = [string](Get-TriageProp $row @('category', 'Category'))
            if (-not $title) { return '' }
            return ('{0} {1} {2} {3}' -f $when, $severity, $category, $title).Trim()
        }
    }
    Add-TriageBecFactLines -Builder $sb -Title ('Sign-ins ({0}):' -f $zoneLabel) -Rows (Get-TriageProp $Facts @('signIns', 'SignIns')) -Max 25 -Format {
        param($row)
        $when = ConvertTo-TriagePresentedTime (Get-TriageProp $row @('created', 'Created'))
        $ip = [string](Get-TriageProp $row @('ip', 'Ip'))
        $where = [string](Get-TriageProp $row @('location', 'Location'))
        $app = [string](Get-TriageProp $row @('app', 'App'))
        $result = [string](Get-TriageProp $row @('result', 'Result'))
        if (-not $when -and -not $ip) { return '' }
        if (Test-TriageInvestigatorAppText $app) { return '' }
        $resultLabels = @{
            '0'      = 'success'
            '50126'  = 'failed: wrong password'
            '50053'  = 'failed: account locked or sign-in IP blocked'
            '50055'  = 'failed: password expired'
            '50057'  = 'failed: account disabled'
            '50074'  = 'interrupted: MFA required'
            '50076'  = 'interrupted: MFA required'
            '50079'  = 'interrupted: MFA registration required'
            '50140'  = 'interrupted: stay signed in prompt'
            '50158'  = 'interrupted: external security challenge'
            '53003'  = 'failed: blocked by Conditional Access'
            '500121' = 'failed: MFA not completed'
            '700082' = 'failed: refresh token expired from inactivity'
        }
        if ($resultLabels.ContainsKey($result)) { $result = ('{0} ({1})' -f $result, $resultLabels[$result]) }
        elseif ($result) { $result = ('{0} (failed or interrupted)' -f $result) }
        $kind = [string](Get-TriageProp $row @('kind', 'Kind'))
        if ($kind -eq 'non-interactive') {
            $count = [string](Get-TriageProp $row @('count', 'Count'))
            $first = ConvertTo-TriagePresentedTime (Get-TriageProp $row @('firstSeen', 'FirstSeen'))
            $kind = ('non-interactive x{0}, first {1}, last' -f $count, $first)
        }
        return ('{0} {1} {2} {3} {4} {5}' -f $kind, $when, $ip, $where, $app, $result).Trim()
    }
    $blastRows = Get-TriageProp $Facts @('blasts', 'Blasts')
    if ($null -ne $blastRows) {
        [void]$sb.AppendLine('Outbound email blasts delivered from this mailbox:')
        $blastShown = 0
        foreach ($blast in @(ConvertTo-TriageArray $blastRows)) {
            $countText = [string](Get-TriageProp $blast @('count', 'Count'))
            $n = 0
            if ($countText -match '^\d+$') { $n = [int]$countText }
            else { $n = @(ConvertTo-TriageArray (Get-TriageProp $blast @('recipients', 'Recipients'))).Count }
            if ($n -lt 5) { continue }
            $subject = [string](Get-TriageProp $blast @('subject', 'Subject'))
            $when = ConvertTo-TriagePresentedTime (Get-TriageProp $blast @('received', 'Received'))
            if (-not $subject.Trim()) { $subject = '(no subject)' }
            [void]$sb.AppendLine(('- {0} {1}: {2} recipients delivered. The incident report lists each address on its own line.' -f $(if ($when) { $when } else { 'undated' }), $subject.Trim(), $n))
            $blastShown++
        }
        if (-not $blastShown) { [void]$sb.AppendLine('- none') }
    }
    }
    $pullNotes = Get-TriageProp $Facts @('pullNotes', 'PullNotes')
    foreach ($note in @(ConvertTo-TriageArray $pullNotes)) {
        $text = [string]$note
        if ($text -match 'Unknown remediation command|REMEDIATE_FAILED') { continue }
        if ($text.Trim()) { [void]$sb.AppendLine($text.Trim()) }
    }
    $ticketNumber = [string](Get-TriageProp $Facts @('ticketNumber', 'TicketNumber'))
    $ticket = [string](Get-TriageProp $Facts @('ticket', 'Ticket'))
    if ($ticketNumber.Trim() -or $ticket.Trim()) {
        [void]$sb.AppendLine('Ticket (optional context):')
        if ($ticketNumber.Trim()) { [void]$sb.AppendLine(('Ticket number: {0}' -f $ticketNumber.Trim())) }
        if ($ticket.Trim()) { [void]$sb.AppendLine((Get-TriageClippedText $ticket 4000)) }
    }
    $company = [string](Get-TriageProp $Facts @('company', 'Company'))
    $gap = Get-TriageProp $Facts @('collectionGap', 'CollectionGap')
    if ($gap) {
        $unchecked = @(ConvertTo-TriageArray (Get-TriageProp $gap @('unchecked', 'Unchecked')) | ForEach-Object { [string]$_ } | Where-Object { $_.Trim() })
        $lead = [string](Get-TriageProp $gap @('lead', 'Lead'))
        if ($unchecked.Count -or $lead.Trim()) {
            [void]$sb.AppendLine('')
            [void]$sb.AppendLine('Collection gap for this tenant:')
            if ($unchecked.Count) {
                [void]$sb.AppendLine(('Not yet checked: {0}. Empty rows for those sources are not a clean result.' -f ($unchecked -join ', ')))
            }
            if ($lead.Trim()) {
                [void]$sb.AppendLine(('Reported lead: {0}' -f $lead.Trim()))
                [void]$sb.AppendLine('The lead is what was reported. It is not proof by itself. Use collected facts to confirm or deny it. If the sources above are still absent, say the lead is still unchecked.')
            }
        }
    }
    $context = Get-TriageMemberberryReviewContext -Company $company -Ticket $ticket
    if ($context) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine($context)
    }
    $revision = [string](Get-TriageProp $Facts @('revision', 'Revision'))
    if ($revision.Trim()) {
        $prior = [string](Get-TriageProp $Facts @('priorReview', 'PriorReview'))
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('Technician revision. Rewrite the review to follow this. The evidence rules still apply. Do not add a true positive or a password reset unless the pack already contains the required evidence, or this revision explicitly asks for that wording. Do not add facts that are not already in the pack.')
        [void]$sb.AppendLine((Get-TriageEvidenceRules))
        if ($prior.Trim()) {
            [void]$sb.AppendLine('Previous review:')
            [void]$sb.AppendLine((Get-TriageClippedText $prior 6000))
        }
        [void]$sb.AppendLine('Change the review as follows:')
        [void]$sb.AppendLine((Get-TriageClippedText $revision 2000))
    }
    return $sb.ToString().Trim()
}

function Get-TriageBecReviewSystemPrompt {
    param([bool]$ConfirmedIncident)
    $system = @'
The first line of the reply must be exactly one of these: VERDICT: true positive, VERDICT: false positive, VERDICT: authorized activity, or VERDICT: ask the client. true positive means the pulled data contains concrete evidence of malicious activity. false positive means the alert fired and the activity is not malicious. authorized activity means someone was allowed to do it. VERDICT: ask the client means the pulled data is inconclusive, so ask the client whether the activity was theirs or authorized. Do not use true positive for that case.
EVIDENCE RULES
You review one mailbox for malicious activity. Cover business email compromise, data exfiltration, DLP policy violations, and any other malicious activity the facts support. Use only the facts in the pack. If a section is not in the pack, do not say that source was checked and do not say none were found. A ticket may be included as context for what was reported. If no ticket is included, review the mailbox anyway. Memberberry client exceptions in the pack are reference for preferred names, VIP status, known tools, and known false positive patterns. They are not evidence that this incident happened. Do not copy another client's incident into this review. Say whether compromise, exfiltration, a DLP violation, or other malicious activity is supported, not supported, or unknown. Name which of those the facts support. An empty list was checked and had no rows. Say that it was checked. Do not tell anyone to collect unified audit, file sharing, DLP, or sign-in logs. The check downloads those. If the pack says a source was not collected, say that in one sentence. Do not treat a missing source as a clean result. REMEDIATE_FAILED:Unknown remediation command means the audit command did not run. Empty logs from that failure are not an inactive mailbox. Say that those sources were not collected, in one sentence, and do not write an operational note about reconciling the command. If the pack includes a collection gap, those named sources have not been checked for this tenant. Say that, and do not treat them as clean. A reported lead is the event to confirm or deny. While the named sources are absent, say the lead is still unchecked. If the pack includes reported messages opened from the mailbox, that is the suspicious mail itself. Judge that lead from the sender name, reply-to, authentication, links, and attachments in that section. Do not say the message content was not inspected. Links include addresses that were only in a link or image target, not only text the reader would see. A body that says it is zero-width characters has no visible text. Do not call that a tracking message unless a link is listed. Evidence includes inbox rules, forwarding, delegates, transport rules, mail connectors, new app registrations, new MFA methods, devices, Intune devices, OAuth consents, directory roles, message trace including large messages, unified audit including file downloads, sharing, and DLP, security alerts, and sign-ins. River Run Security Investigator is the app this check uses to pull logs. Ignore it. Do not mention it, and do not ask anyone to confirm who created it. If the pack lists an outbound email blast, say the subject and how many recipients were delivered. Do not paste those recipient addresses in the review. Log times in the pack are the technician local time, named in the pack. Write every time in that local time. Do not write UTC, a trailing Z, or a numeric offset. A ticket clock time is already that local time. If a fact the ticket needs is missing and it may be older than this window, that is not a finding. Do not write that the record was outside the window, and do not say you cannot tell who did it because of the window. End with one line, REPULL: <days> <logs>. Days is the smallest lookback that would include the missing event. If you cannot tell how old it is, use 30. Do not ask for more days than the limit stated in the pack. Logs are only the sources that can answer it, separated by commas, from this list: trace, ual, sharing, files, dlp, alerts, signins, apps. A link that was used in the window, when the sharing rows do not show who created that link, needs sharing. Add files only when the file activity itself is missing. Do not ask for trace or signins unless those are the rows that are missing. If this window already answers the ticket, end with REPULL: 0. The check pulls the logs you name and reviews again before the investigation is finished. A follow-up pull is not evidence of a true positive. If those logs are still inconclusive, keep VERDICT: ask the client and ask whether the activity was theirs or authorized. Do not switch to true positive just because the window got longer. If that window still does not contain the event, ask again with a longer look. If the pack says not to ask for another pull, end with REPULL: 0 and you may say the record was not in the logs that were pulled. If a pull note says a longer window was refused by the license, that shorter window is as far as this tenant can go for that log. Do not ask for the refused length again. Review the window that was pulled. Do not invent rules, recipients, files, policies, or incidents that are not in the pack. The reply must be plaintext, not Markdown. Do not use HTML, asterisks for emphasis, hash headings, or code fences. Never use an en dash, an em dash, or a double dash. Use a comma, a colon, parentheses, or a single hyphen. After the review, add a suggestion line only when this investigation showed a concrete pattern worth keeping. Omit a line when you have nothing concrete.
MEMBERBERRY EXCEPTION SUGGESTION: <client name> - <one sentence for that client's notes>
MEMORY VAULT: <one sentence the next review should remember>
If the pack includes a technician revision, rewrite the review to follow it. The evidence rules still apply. Do not add a true positive or a password reset unless the pack already contains the required evidence, or the technician revision explicitly asks for that wording. Do not add facts that are not already in the pack.
'@
    $system = $system.Replace('EVIDENCE RULES', (Get-TriageEvidenceRules))
    if ($ConfirmedIncident) {
        $system += @'

The ticket already confirms a business email compromise or another security incident. Do not spend the review deciding whether that incident happened. Focus on whether remediation is complete. For inbox rules, forwarding, delegates, OAuth consents, MFA methods, app registrations, and anything the ticket says was already done, say what is still in place and what is already gone. A step is done only when the pack or the ticket says it was done. Before the REPULL line, add one line: REMEDIATION: complete or REMEDIATION: incomplete. Do not recommend a password reset or session revocation unless the pack or the ticket shows a malicious sign-in, a malicious inbox rule, or confirmed credential exposure.
'@
    }
    return $system
}

function Invoke-TriageBecReview {
    param(
        [Parameter(Mandatory)][string]$SettingsPath,
        [Parameter(Mandatory)]$Facts,
        [scriptblock]$OnDelta
    )
    $stored = Read-TriageAiSettingsFile -Path $SettingsPath
    $provider = [string]$stored.provider
    if ($provider -notin @('chatgpt', 'claude', 'hatz')) {
        throw 'Choose ChatGPT, Claude, or Hatz AI in AI review settings on the Email triage page.'
    }
    $key = [string]$stored.apiKey
    if ([string]::IsNullOrWhiteSpace($key)) {
        throw 'Save an API key in AI review settings on the Email triage page. It stays on this PC.'
    }
    Test-TriageAiKeyToken -ApiKey $key
    $model = [string]$stored.model
    if (-not $model) { $model = Get-TriageAiDefaultModel -Provider $provider }
    $brief = New-TriageBecReviewBrief -Facts $Facts
    $confirmedIncident = Test-TriageConfirmedIncident -Text ([string](Get-TriageProp $Facts @('ticket', 'Ticket')))
    $system = Get-TriageBecReviewSystemPrompt -ConfirmedIncident $confirmedIncident
    $text = ''
    if ($OnDelta) {
        $streamRequest = Get-TriageAiReviewRequest -Provider $provider -Model $model -ApiKey $key -Brief $brief -SystemPrompt $system -MaxTokens 1600 -Stream
        try {
            $text = [string](Invoke-TriageAiStream -Request $streamRequest -Provider $provider -OnDelta $OnDelta)
        }
        catch {
            $text = ''
        }
    }
    if ([string]::IsNullOrWhiteSpace($text)) {
        $request = Get-TriageAiReviewRequest -Provider $provider -Model $model -ApiKey $key -Brief $brief -SystemPrompt $system -MaxTokens 1600
        try {
            $response = Invoke-RestMethod -Method Post -Uri $request.Uri -Headers $request.Headers -Body $request.Body -ContentType 'application/json' -TimeoutSec 300
        }
        catch {
            throw (Get-TriageAiErrorMessage -ErrorRecord $_)
        }
        $text = (ConvertFrom-TriageAiResponse -Response $response -Provider $provider)
    }
    $text = (ConvertTo-TriageAiPlainText -Text $text).Trim()
    if (-not $text) { throw 'The model returned an empty review.' }
    $currentDays = 0
    $rawDays = [string](Get-TriageProp $Facts @('windowDays', 'WindowDays'))
    if ($rawDays -match '^\d+$') { $currentDays = [int]$rawDays }
    $follow = Get-TriageBecFollowUpPlan -Text $text -CurrentDays $currentDays
    $repullDays = [int]$follow.Days
    $repullLogs = @($follow.Logs)
    $suggestions = ConvertFrom-TriageStoreSuggestions -Text $text
    $remediations = @(Get-TriageBecRemediations -Facts $Facts)
    return [pscustomobject]@{
        provider      = $provider
        model         = $model
        text          = [string]$suggestions.Text
        suspicious    = ($remediations.Count -gt 0)
        remediations  = $remediations
        repullDays    = $repullDays
        repullLogs    = $repullLogs
        verdict       = (Get-TriageBecVerdict -Text ([string]$suggestions.Text))
        storeClient   = [string]$suggestions.Client
        storeNote     = [string]$suggestions.Note
        storeMemory   = [string]$suggestions.Memory
    }
}

function Get-TriageLogPlanCatalog {
    return @(
        [pscustomobject]@{ Key = 'IncludeMessageTrace'; Label = 'Message trace'; Volume = 'medium' }
        [pscustomobject]@{ Key = 'IncludeInboxRules'; Label = 'Inbox rules'; Volume = 'small' }
        [pscustomobject]@{ Key = 'IncludeTransportRules'; Label = 'Transport rules'; Volume = 'small' }
        [pscustomobject]@{ Key = 'IncludeMailFlowConnectors'; Label = 'Mail flow connectors'; Volume = 'small' }
        [pscustomobject]@{ Key = 'IncludeMailboxForwarding'; Label = 'Mailbox forwarding'; Volume = 'small' }
        [pscustomobject]@{ Key = 'IncludeUnifiedAuditLogs'; Label = 'Unified audit logs (Exchange)'; Volume = 'high' }
        [pscustomobject]@{ Key = 'IncludeExchangeItemAggregated'; Label = 'UAL ExchangeItemAggregated'; Volume = 'very high' }
        [pscustomobject]@{ Key = 'IncludeSignInLogs'; Label = 'Sign-in logs'; Volume = 'medium' }
        [pscustomobject]@{ Key = 'IncludeAuditLogs'; Label = 'Directory audit logs'; Volume = 'medium' }
        [pscustomobject]@{ Key = 'IncludeMfaCoverage'; Label = 'MFA coverage'; Volume = 'small' }
        [pscustomobject]@{ Key = 'IncludeConditionalAccessPolicies'; Label = 'Conditional Access policies'; Volume = 'small' }
        [pscustomobject]@{ Key = 'IncludeAppRegistrations'; Label = 'App registrations'; Volume = 'small' }
        [pscustomobject]@{ Key = 'IncludeIntuneDevices'; Label = 'Intune devices'; Volume = 'small' }
        [pscustomobject]@{ Key = 'IncludeSecurityAlerts'; Label = 'Security alerts'; Volume = 'small' }
        [pscustomobject]@{ Key = 'IncludeSecurityIncidents'; Label = 'Security incidents'; Volume = 'small' }
        [pscustomobject]@{ Key = 'IncludeSharePointActivity'; Label = 'SharePoint activity'; Volume = 'medium' }
        [pscustomobject]@{ Key = 'IncludeOneDriveActivity'; Label = 'OneDrive activity'; Volume = 'medium' }
        [pscustomobject]@{ Key = 'IncludeTeamsActivity'; Label = 'Teams activity'; Volume = 'medium' }
        [pscustomobject]@{ Key = 'IncludeSharePointSharing'; Label = 'SharePoint sharing links'; Volume = 'medium' }
        [pscustomobject]@{ Key = 'IncludeAnonymousSharePointSharing'; Label = 'Anonymous SharePoint sharing'; Volume = 'medium' }
        [pscustomobject]@{ Key = 'IncludeSharePointFileSharingLinks'; Label = 'SharePoint file sharing links'; Volume = 'medium' }
        [pscustomobject]@{ Key = 'IncludeDLPViolations'; Label = 'DLP violations'; Volume = 'small' }
        [pscustomobject]@{ Key = 'IncludeSharePointOneDriveFileActions'; Label = 'SharePoint and OneDrive file actions'; Volume = 'very high' }
    )
}

function ConvertFrom-TriageLogPlanText {
    param([string]$Text)
    $allowed = @{}
    foreach ($item in @(Get-TriageLogPlanCatalog)) { $allowed[$item.Key] = $true }
    $days = 0
    $logs = New-Object System.Collections.Generic.List[string]
    $notes = New-Object System.Collections.Generic.List[string]
    $seen = @{}
    $more = $false
    $sawMore = $false
    foreach ($line in ([string]$Text -split '\r?\n')) {
        $line = $line.Trim()
        if ($line -match '^(?i)DAYS:\s*(\d+)') {
            $days = [int]$Matches[1]
        }
        elseif ($line -match '^(?i)LOG:\s*([A-Za-z0-9]+)') {
            $key = $Matches[1]
            if ($allowed.ContainsKey($key) -and -not $seen.ContainsKey($key)) {
                $seen[$key] = $true
                [void]$logs.Add($key)
            }
        }
        elseif ($line -match '^(?i)NOTE:\s*(.+)') {
            [void]$notes.Add($Matches[1].Trim())
        }
        elseif ($line -match '^(?i)MORE:\s*(yes|no)\b') {
            $sawMore = $true
            $more = $Matches[1].ToLowerInvariant() -eq 'yes'
        }
    }
    if ($days -lt 1) { $days = 1 }
    if ($days -gt 90) { $days = 90 }
    return [pscustomobject]@{
        Days  = $days
        Logs  = $logs.ToArray()
        Note  = (($notes -join ' ').Trim())
        More  = $(if ($sawMore) { $more } else { $logs.Count -gt 0 })
    }
}

function Test-TriageLogReportFolder {
    param([Parameter(Mandatory)][string]$Folder)
    $full = [IO.Path]::GetFullPath($Folder)
    $root = [IO.Path]::GetFullPath((Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'ExchangeOnlineAnalyzer'))
    $prefix = $root.TrimEnd('\') + '\'
    $inside = $full.Equals($root, [StringComparison]::OrdinalIgnoreCase) -or $full.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)
    if (-not $inside) { throw 'That folder is outside ExchangeOnlineAnalyzer reports.' }
    if (-not (Test-Path -LiteralPath $full -PathType Container)) { throw 'Report folder was not found.' }
    return $full
}

function Get-TriageLogFolderInventory {
    param([Parameter(Mandatory)][string]$Folder)
    $rows = New-Object System.Collections.Generic.List[string]
    $files = @(Get-ChildItem -LiteralPath $Folder -File -Recurse -Depth 2 -ErrorAction SilentlyContinue |
        Sort-Object Length -Descending |
        Select-Object -First 60)
    foreach ($file in $files) {
        $kb = [math]::Round($file.Length / 1KB, 1)
        [void]$rows.Add(('{0} ({1} KB)' -f $file.Name, $kb))
    }
    return $rows.ToArray()
}

function New-TriageLogPlanBrief {
    param(
        [Parameter(Mandatory)][string]$Ticket,
        [ValidateSet('initial', 'followup')]
        [string]$Phase = 'initial',
        [string[]]$Inventory = @()
    )
    $ticket = $Ticket.Trim()
    if ($ticket.Length -gt 6000) { $ticket = $ticket.Substring(0, 6000) + "`n(Ticket text truncated.)" }
    $sb = New-Object System.Text.StringBuilder
    if ($Phase -eq 'followup') {
        [void]$sb.AppendLine('The report pack is already downloaded. Suggest only extra logs that would answer something the ticket still leaves open. If the pack is enough, say MORE: no and do not list LOG lines.')
    }
    else {
        [void]$sb.AppendLine('Choose the smallest Entra and Microsoft 365 pull that can investigate this ticket. Prefer a short window. Skip a log when the ticket does not need it. Very high volume logs need a specific reason.')
    }
    [void]$sb.AppendLine('Catalog (use the key on LOG lines):')
    foreach ($item in @(Get-TriageLogPlanCatalog)) {
        [void]$sb.AppendLine(('- {0}: {1} ({2})' -f $item.Key, $item.Label, $item.Volume))
    }
    if ($Phase -eq 'followup') {
        [void]$sb.AppendLine('Files already downloaded (names and sizes only):')
        $shown = 0
        foreach ($row in @($Inventory)) {
            if ($shown -ge 60) { break }
            if ($row) { [void]$sb.AppendLine(('- {0}' -f $row)); $shown++ }
        }
        if (-not $shown) { [void]$sb.AppendLine('- none') }
    }
    [void]$sb.AppendLine('Ticket:')
    [void]$sb.AppendLine($ticket)
    $context = Get-TriageMemberberryReviewContext -Ticket $ticket
    if ($context) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine($context)
    }
    return $sb.ToString().Trim()
}

function Invoke-TriageLogPlan {
    param(
        [Parameter(Mandatory)][string]$SettingsPath,
        [string]$Ticket = '',
        [ValidateSet('initial', 'followup')]
        [string]$Phase = 'initial',
        [string]$Folder = ''
    )
    if ([string]::IsNullOrWhiteSpace($Ticket)) { throw 'Paste a ticket or fetch it from Manage first.' }
    $stored = Read-TriageAiSettingsFile -Path $SettingsPath
    $provider = [string]$stored.provider
    if ($provider -notin @('chatgpt', 'claude', 'hatz')) {
        throw 'Choose ChatGPT, Claude, or Hatz AI in AI review settings on the Email triage page.'
    }
    $key = [string]$stored.apiKey
    if ([string]::IsNullOrWhiteSpace($key)) {
        throw 'Save an API key in AI review settings on the Email triage page. It stays on this PC.'
    }
    Test-TriageAiKeyToken -ApiKey $key
    $model = [string]$stored.model
    if (-not $model) { $model = Get-TriageAiDefaultModel -Provider $provider }
    $inventory = @()
    if ($Phase -eq 'followup') {
        if ([string]::IsNullOrWhiteSpace($Folder)) { throw 'Generate a report pack before asking for more logs.' }
        $safe = Test-TriageLogReportFolder -Folder $Folder
        $inventory = @(Get-TriageLogFolderInventory -Folder $safe)
    }
    $brief = New-TriageLogPlanBrief -Ticket $Ticket -Phase $Phase -Inventory $inventory
    $system = @'
You choose which Entra and Microsoft 365 logs to download for one ticket. Use the ticket, the catalog, and any Memberberry client exceptions in the pack. Those exceptions are reference for how this client is usually handled. They are not evidence about this ticket. Do not copy another client's incident. Reply in plaintext, not Markdown. Do not use HTML, asterisks, hash headings, or code fences. Never use an en dash, an em dash, or a double dash. Use this shape:
DAYS: <1 to 90>
LOG: <catalog key>
NOTE: <one or two sentences>
MORE: yes or no
List one LOG line per log. On a follow-up, MORE: no means no LOG lines. After that shape, add a suggestion line only when the ticket shows a concrete pattern worth keeping. Omit a line when you have nothing concrete.
MEMBERBERRY EXCEPTION SUGGESTION: <client name> - <one sentence for that client's notes>
MEMORY VAULT: <one sentence the next review should remember>
'@
    $request = Get-TriageAiReviewRequest -Provider $provider -Model $model -ApiKey $key -Brief $brief -SystemPrompt $system
    try {
        $response = Invoke-RestMethod -Method Post -Uri $request.Uri -Headers $request.Headers -Body $request.Body -ContentType 'application/json' -TimeoutSec 90
    }
    catch {
        throw (Get-TriageAiErrorMessage -ErrorRecord $_)
    }
    $text = (ConvertTo-TriageAiPlainText -Text (ConvertFrom-TriageAiResponse -Response $response -Provider $provider)).Trim()
    if (-not $text) { throw 'The model returned an empty log plan.' }
    $suggestions = ConvertFrom-TriageStoreSuggestions -Text $text
    $plan = ConvertFrom-TriageLogPlanText -Text $suggestions.Text
    return [pscustomobject]@{
        provider    = $provider
        model       = $model
        text        = [string]$suggestions.Text
        days        = $plan.Days
        logs        = @($plan.Logs)
        note        = $plan.Note
        more        = [bool]$plan.More
        phase       = $Phase
        storeClient = [string]$suggestions.Client
        storeNote   = [string]$suggestions.Note
        storeMemory = [string]$suggestions.Memory
    }
}

function Get-TriageClippedText {
    param([string]$Text, [int]$Max = 500)
    $text = ([string]$Text).Trim()
    if ($text.Length -le $Max) { return $text }
    return $text.Substring(0, $Max).Trim()
}

function Get-TriageMemberberryMatchKey {
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return '' }
    return ([regex]::Replace($Name.ToLowerInvariant(), '[^a-z0-9]', ''))
}

function Get-TriageTicketCompany {
    param([string]$Ticket)
    if ([string]::IsNullOrWhiteSpace($Ticket)) { return '' }
    $labeled = [regex]::Match($Ticket, '(?im)^\s*Company:\s*(\S.+?)\s*$')
    if ($labeled.Success) {
        $name = $labeled.Groups[1].Value.Trim()
        if ($name -and $name -notmatch '^(?i)company$') { return $name }
    }
    $lines = $Ticket -split "`r?`n"
    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($lines[$i].Trim() -ne 'Company') { continue }
        $stop = [Math]::Min($i + 6, $lines.Count)
        for ($j = $i + 1; $j -lt $stop; $j++) {
            $candidate = $lines[$j].Trim().TrimStart('*').Trim()
            if ($candidate -and $candidate -ne 'Company') { return $candidate }
        }
    }
    return ''
}

function Get-TriageTicketContact {
    param([string]$Ticket)
    if ([string]::IsNullOrWhiteSpace($Ticket)) { return '' }
    $labeled = [regex]::Match($Ticket, '(?im)^\s*Contact:\s*(\S.+?)\s*$')
    if ($labeled.Success) {
        $name = $labeled.Groups[1].Value.Trim()
        if ($name -and $name -notmatch '^(?i)contact$') { return $name }
    }
    return ''
}

function Get-TriageMemberberryPaths {
    param([string]$Root = '')
    $candidate = $Root
    if ([string]::IsNullOrWhiteSpace($candidate)) { $candidate = 'C:\Git\memberberry' }
    $configPath = Join-Path $candidate 'config.json'
    if (-not (Test-Path -LiteralPath $configPath)) {
        $fallback = Join-Path $env:USERPROFILE 'Documents\memberberry'
        $fallbackConfig = Join-Path $fallback 'config.json'
        if (-not (Test-Path -LiteralPath $fallbackConfig)) { return $null }
        $candidate = $fallback
        $configPath = $fallbackConfig
    }
    try {
        $config = Get-Content -LiteralPath $configPath -Raw -Encoding UTF8 | ConvertFrom-Json
    }
    catch { return $null }
    $data = [string](Get-TriageProp $config @('data_path'))
    $user = [string](Get-TriageProp $config @('user_data_path'))
    if ([string]::IsNullOrWhiteSpace($user)) { $user = $data }
    return [pscustomobject]@{
        Root       = $candidate
        Exceptions = $(if ($data) { Join-Path $data 'exceptions.json' } else { '' })
        Settings   = $(if ($user) { Join-Path $user 'settings.json' } else { '' })
    }
}

function Get-TriageMemberberrySettings {
    param([string]$SettingsPath)
    $settings = [ordered]@{
        TechnicianName = ''
        CompanyName    = ''
        ContactNumber  = ''
        Timezone       = ''
        SupportHours   = ''
    }
    if ([string]::IsNullOrWhiteSpace($SettingsPath) -or -not (Test-Path -LiteralPath $SettingsPath)) { return [pscustomobject]$settings }
    try {
        $doc = Get-Content -LiteralPath $SettingsPath -Raw -Encoding UTF8 | ConvertFrom-Json
    }
    catch { return [pscustomobject]$settings }
    $tech = Get-TriageProp $doc @('technician_name')
    $company = Get-TriageProp $doc @('company_name')
    $phone = Get-TriageProp $doc @('contact_number')
    $zone = Get-TriageProp $doc @('timezone')
    $hours = Get-TriageProp $doc @('support_hours')
    if ($tech) { $settings.TechnicianName = [string]$tech }
    if ($company) { $settings.CompanyName = [string]$company }
    if ($phone) { $settings.ContactNumber = [string]$phone }
    if ($zone) { $settings.Timezone = [string]$zone }
    if ($hours) { $settings.SupportHours = [string]$hours }
    return [pscustomobject]$settings
}

function Format-TriageMemberberryNameItem {
    param($Item)
    if ($null -eq $Item -or $Item -is [string]) { return '' }
    $from = [string](Get-TriageProp $Item @('name', 'Name'))
    $to = [string](Get-TriageProp $Item @('preferred', 'Preferred', 'preferred_name'))
    if ($from -and $to) { return ('{0} -> {1}' -f $from.Trim(), $to.Trim()) }
    if ($to) { return $to.Trim() }
    if ($from) { return $from.Trim() }
    return ''
}

function Get-TriageMemberberryNamePairs {
    param($Names)
    $pairs = New-Object System.Collections.Generic.List[object]
    if ($null -eq $Names -or $Names -is [string]) { return ,$pairs }
    if ($Names -is [System.Collections.IDictionary]) {
        foreach ($k in @($Names.Keys)) {
            if ($Names[$k] -is [string]) { [void]$pairs.Add(@([string]$k, [string]$Names[$k])) }
        }
        return ,$pairs
    }
    if (($Names -is [System.Array]) -or ($Names -is [System.Collections.IList])) {
        foreach ($item in @($Names)) {
            if ($null -eq $item -or $item -is [string]) { continue }
            $from = [string](Get-TriageProp $item @('name', 'Name'))
            $to = [string](Get-TriageProp $item @('preferred', 'Preferred', 'preferred_name'))
            if ($from -and $to) { [void]$pairs.Add(@($from, $to)) }
        }
        return ,$pairs
    }
    $from = [string](Get-TriageProp $Names @('name', 'Name'))
    $to = [string](Get-TriageProp $Names @('preferred', 'Preferred', 'preferred_name'))
    if ($from -and $to) {
        [void]$pairs.Add(@($from, $to))
        return ,$pairs
    }
    # Memberberry stores names as { "Full Name": "Preferred" }.
    foreach ($p in $Names.PSObject.Properties) {
        if ($p.Value -is [string]) { [void]$pairs.Add(@([string]$p.Name, [string]$p.Value)) }
    }
    return ,$pairs
}

function Test-TriageNameSpellingClose {
    param([string]$A, [string]$B)
    if ($A -eq $B) { return $true }
    if ($A.Length -lt 4 -or $B.Length -lt 4) { return $false }
    if ([Math]::Abs($A.Length - $B.Length) -gt 1) { return $false }
    # One inserted, deleted, or changed letter (Wilms vs Willms).
    $i = 0
    $j = 0
    $edits = 0
    while ($i -lt $A.Length -and $j -lt $B.Length) {
        if ($A[$i] -eq $B[$j]) { $i++; $j++; continue }
        $edits++
        if ($edits -gt 1) { return $false }
        if ($A.Length -gt $B.Length) { $i++ }
        elseif ($B.Length -gt $A.Length) { $j++ }
        else { $i++; $j++ }
    }
    $edits += ($A.Length - $i) + ($B.Length - $j)
    return ($edits -le 1)
}

function Get-TriageMemberberryGreetingName {
    param($Record, [string]$Contact)
    $contactText = ([string]$Contact).Trim()
    if (-not $contactText -or $null -eq $Record) { return '' }
    $names = $null
    if ($Record -is [System.Collections.IDictionary]) {
        if ($Record.Contains('names')) { $names = $Record['names'] }
    }
    elseif ($Record.PSObject.Properties['names']) {
        $names = $Record.PSObject.Properties['names'].Value
    }
    $pairs = Get-TriageMemberberryNamePairs $names
    $contactWords = @($contactText.ToLowerInvariant() -split '\s+' | Where-Object { $_ })
    $firstOnly = New-Object System.Collections.Generic.List[string]
    foreach ($pair in $pairs) {
        $from = ([string]$pair[0]).Trim()
        $to = ([string]$pair[1]).Trim()
        if (-not $from -or -not $to) { continue }
        if ($from -ieq $contactText) { return $to }
        $fromWords = @($from.ToLowerInvariant() -split '\s+' | Where-Object { $_ })
        if ($fromWords.Count -eq 0 -or $contactWords.Count -eq 0) { continue }
        if ($fromWords[0] -ne $contactWords[0]) { continue }
        if ($contactWords.Count -gt 1 -and $fromWords.Count -gt 1 -and (Test-TriageNameSpellingClose $fromWords[-1] $contactWords[-1])) { return $to }
        if ($contactWords.Count -eq 1 -or $fromWords.Count -eq 1) {
            if (-not $firstOnly.Contains($to)) { [void]$firstOnly.Add($to) }
        }
    }
    # A bare first name shared by people with different preferred names is ambiguous.
    if ($firstOnly.Count -eq 1) { return $firstOnly[0] }
    return ''
}

function ConvertTo-TriageMemberberryNames {
    param($Names)
    if ($null -eq $Names) { return '' }
    if ($Names -is [string]) { return $Names.Trim() }
    if (-not ($Names -is [System.Array]) -and -not ($Names -is [System.Collections.IList]) -and -not (Format-TriageMemberberryNameItem $Names)) {
        $mapped = New-Object System.Collections.Generic.List[string]
        $namePairs = Get-TriageMemberberryNamePairs $Names
        foreach ($pair in $namePairs) { [void]$mapped.Add(('{0} -> {1}' -f $pair[0], $pair[1])) }
        if ($mapped.Count -gt 0) { return ($mapped -join '; ') }
    }
    $parts = New-Object System.Collections.Generic.List[string]
    $items = @($Names)
    if (($Names -is [System.Array]) -or ($Names -is [System.Collections.IList])) {
        foreach ($item in $items) {
            if ($null -eq $item) { continue }
            if ($item -is [string]) {
                if ($item.Trim()) { [void]$parts.Add($item.Trim()) }
                continue
            }
            $named = Format-TriageMemberberryNameItem $item
            if ($named) { [void]$parts.Add($named) }
        }
        return ($parts -join '; ')
    }
    return (Format-TriageMemberberryNameItem $Names)
}

function Format-TriageMemberberryValue {
    param($Value)
    if ($null -eq $Value) { return '' }
    if ($Value -is [string]) { return $Value.Trim() }
    if ($Value -is [bool]) { return $(if ($Value) { 'true' } else { 'false' }) }
    if ($Value -is [System.Array] -or $Value -is [System.Collections.IList]) {
        $parts = New-Object System.Collections.Generic.List[string]
        foreach ($item in @($Value)) {
            if ($null -eq $item) { continue }
            if ($item -is [string]) {
                if ($item.Trim()) { [void]$parts.Add($item.Trim()) }
            }
            else {
                $named = ConvertTo-TriageMemberberryNames $item
                if ($named) { [void]$parts.Add($named) }
            }
        }
        return ($parts -join '; ')
    }
    return ([string]$Value).Trim()
}

function ConvertTo-TriageMemberberryClientBrief {
    param($Record, [string]$Name, [switch]$OmitLists)
    if (-not $Record) { return '' }
    $lines = New-Object System.Collections.Generic.List[string]
    [void]$lines.Add(('Client: {0}' -f $Name))
    foreach ($pair in @(
            @('detail', 'Detail'),
            @('include_what_happened', 'Include what happened'),
            @('escalation_threshold', 'Escalation threshold'),
            @('industry', 'Industry'),
            @('preferred_communication', 'Preferred communication'),
            @('mfa', 'MFA'),
            @('vpn', 'VPN'),
            @('onsite_it', 'Onsite IT')
        )) {
        $text = Format-TriageMemberberryValue (Get-TriageProp $Record @($pair[0]))
        if ($text) { [void]$lines.Add(('{0}: {1}' -f $pair[1], (Get-TriageClippedText $text 300))) }
    }
    $names = ConvertTo-TriageMemberberryNames (Get-TriageProp $Record @('names'))
    if ($names) { [void]$lines.Add(('Preferred names: {0}' -f (Get-TriageClippedText $names 400))) }
    $listPairs = @(
            @('vips', 'VIPs'),
            @('authorized_tools', 'Authorized tools'),
            @('blocked_tools', 'Blocked tools'),
            @('false_positive_patterns', 'False positive patterns'),
            @('expected_remote_locations', 'Expected remote locations'),
            @('common_ip_ranges', 'Common IP ranges'),
            @('notes', 'Notes')
        )
    if ($OmitLists) {
        $listPairs = @(
            @('vips', 'VIPs'),
            @('expected_remote_locations', 'Expected remote locations'),
            @('common_ip_ranges', 'Common IP ranges')
        )
    }
    foreach ($pair in $listPairs) {
        $text = Format-TriageMemberberryValue (Get-TriageProp $Record @($pair[0]))
        if ($text) { [void]$lines.Add(('{0}: {1}' -f $pair[1], (Get-TriageClippedText $text 600))) }
    }
    return ($lines -join "`n")
}

function Find-TriageMemberberryClient {
    param([string]$Company, [string]$ExceptionsPath)
    if ([string]::IsNullOrWhiteSpace($Company) -or [string]::IsNullOrWhiteSpace($ExceptionsPath)) { return $null }
    if (-not (Test-Path -LiteralPath $ExceptionsPath)) { return $null }
    $want = Get-TriageMemberberryMatchKey $Company
    if (-not $want) { return $null }
    $doc = Read-TriageMemberberryDocument -Path $ExceptionsPath
    if (-not $doc) { return $null }
    $exact = $null
    $associated = $null
    foreach ($prop in @($doc.PSObject.Properties)) {
        if ($prop.Name -eq '_global' -or $prop.Name -eq 'global') { continue }
        if ((Get-TriageMemberberryMatchKey $prop.Name) -eq $want) {
            $exact = $prop
            break
        }
        if ($associated) { continue }
        $companies = Format-TriageMemberberryValue (Get-TriageProp $prop.Value @('associated_companies'))
        foreach ($item in @($companies -split '; ')) {
            if ((Get-TriageMemberberryMatchKey $item) -eq $want) {
                $associated = $prop
                break
            }
        }
    }
    $hit = if ($exact) { $exact } else { $associated }
    if (-not $hit) { return $null }
    return [pscustomobject]@{ Name = [string]$hit.Name; Record = $hit.Value }
}

function Get-TriageMemberberryReviewContext {
    param(
        [string]$Company = '',
        [string]$Ticket = '',
        [string]$MemberberryRoot = ''
    )
    $name = $Company.Trim()
    if (-not $name) { $name = Get-TriageTicketCompany -Ticket $Ticket }
    if (-not $name) { return '' }
    $paths = Get-TriageMemberberryPaths -Root $MemberberryRoot
    if (-not $paths -or -not $paths.Exceptions -or -not (Test-Path -LiteralPath $paths.Exceptions)) {
        return ''
    }
    $client = Find-TriageMemberberryClient -Company $name -ExceptionsPath $paths.Exceptions
    if (-not $client) {
        return ('Memberberry client exceptions: no client matched {0}. Do not invent client facts. These exceptions are not evidence about this incident.' -f $name)
    }
    $brief = ConvertTo-TriageMemberberryClientBrief -Record $client.Record -Name $client.Name
    return ("Memberberry client exceptions for {0} (reference for names, VIP, tools, and known patterns, not evidence about this incident):`n{1}" -f $client.Name, $brief)
}

function ConvertFrom-TriageStoreSuggestions {
    param([string]$Text)
    $client = ''
    $note = ''
    $memory = ''
    $kept = New-Object System.Collections.Generic.List[string]
    foreach ($line in ([string]$Text -split "`r?`n")) {
        if ($line -match '^(?i)\s*MEMBERBERRY EXCEPTION SUGGESTION:\s*(.+)$') {
            $rest = $Matches[1].Trim()
            $split = $rest.IndexOf(' - ')
            if ($split -gt 0) {
                $client = $rest.Substring(0, $split).Trim()
                $note = $rest.Substring($split + 3).Trim()
            }
            else { $note = $rest }
            continue
        }
        if ($line -match '^(?i)\s*MEMORY VAULT:\s*(.+)$') {
            $memory = $Matches[1].Trim()
            continue
        }
        [void]$kept.Add($line)
    }
    return [pscustomobject]@{
        Client = $client
        Note   = $note
        Memory = $memory
        Text   = (($kept -join "`n").Trim())
    }
}

function Get-TriageMemberberryGreeting {
    param([datetime]$Now = (Get-Date))
    if ($Now.Hour -lt 12) { return 'Good morning' }
    if ($Now.Hour -lt 17) { return 'Good afternoon' }
    return 'Good evening'
}

function ConvertTo-TriageQuotedArg {
    param([string]$Value)
    $clean = [regex]::Replace([string]$Value, '[^\w\s@.,:;()/#''+-]', ' ')
    $clean = ($clean -replace '\s+', ' ').Trim()
    return '"' + ($clean -replace '"', '') + '"'
}

function ConvertFrom-TriageMemoryVaultText {
    param([string]$Text)
    $hits = New-Object System.Collections.Generic.List[object]
    $lines = [string]$Text -split "`r?`n"
    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($lines[$i] -match '^\s+\[(\d+)\]\s+\S+\s+\[([^\]]+)\]\s*(.*)$') {
            $preview = ''
            if (($i + 1) -lt $lines.Count) { $preview = $lines[$i + 1].Trim() }
            if (-not $preview) { continue }
            [void]$hits.Add([pscustomobject]@{
                    Space  = $Matches[2].Trim()
                    Source = $Matches[3].Trim()
                    Text   = $preview
                })
        }
    }
    return $hits.ToArray()
}

function Invoke-TriageMemoryVaultSearch {
    param(
        [string]$Query,
        [string]$PythonPath,
        [string]$WorkingDirectory,
        [string]$Space = 'soc',
        [int]$Limit = 4
    )
    $skipped = [pscustomobject]@{ Hits = @(); Status = 'skipped' }
    if ([string]::IsNullOrWhiteSpace($Query)) { return $skipped }
    if ([string]::IsNullOrWhiteSpace($PythonPath) -or -not (Test-Path -LiteralPath $PythonPath)) {
        return [pscustomobject]@{ Hits = @(); Status = 'unavailable' }
    }
    if ([string]::IsNullOrWhiteSpace($WorkingDirectory) -or -not (Test-Path -LiteralPath $WorkingDirectory)) {
        return [pscustomobject]@{ Hits = @(); Status = 'unavailable' }
    }
    $q = Get-TriageClippedText (($Query -replace '[\r\n]+', ' ').Trim()) 400
    if ($q.StartsWith('-')) { $q = "ticket $q" }
    if ($Limit -lt 1) { $Limit = 1 }
    if ($Limit -gt 5) { $Limit = 5 }
    $spaceName = if ([string]::IsNullOrWhiteSpace($Space)) { 'soc' } else { $Space.Trim() }
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $PythonPath
    $psi.WorkingDirectory = $WorkingDirectory
    $psi.UseShellExecute = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.CreateNoWindow = $true
    $launch = @'
import asyncio, os, sys
policy = getattr(asyncio, "WindowsSelectorEventLoopPolicy", None)
if policy is not None:
    asyncio.set_event_loop_policy(policy())
sys.argv = ["cli", "search", os.environ["MV_QUERY"], "--space", os.environ.get("MV_SPACE") or "soc", "--limit", os.environ.get("MV_LIMIT") or "4"]
from src.cli import main
main()
'@
    if ($null -ne $psi.ArgumentList) {
        [void]$psi.ArgumentList.Add('-c')
        [void]$psi.ArgumentList.Add($launch)
    }
    else {
        $psi.Arguments = '-c "' + ($launch -replace '"', '\"') + '"'
    }
    try { $psi.EnvironmentVariables['PYTHONPATH'] = $WorkingDirectory } catch { }
    try { $psi.EnvironmentVariables['MV_QUERY'] = $q } catch { }
    try { $psi.EnvironmentVariables['MV_SPACE'] = $spaceName } catch { }
    try { $psi.EnvironmentVariables['MV_LIMIT'] = [string]$Limit } catch { }
    Add-TriageMemoryVaultDatabaseEnv -StartInfo $psi -Root $WorkingDirectory
    $proc = $null
    try {
        $proc = New-Object System.Diagnostics.Process
        $proc.StartInfo = $psi
        [void]$proc.Start()
        $stdoutTask = $proc.StandardOutput.ReadToEndAsync()
        $stderrTask = $proc.StandardError.ReadToEndAsync()
        # The web listener handles one request at a time, so a cold search has to fail open.
        if (-not $proc.WaitForExit(50000)) {
            try { $proc.Kill() } catch { }
            return [pscustomobject]@{ Hits = @(); Status = 'unavailable' }
        }
        $stdout = [string]$stdoutTask.Result
        [void]$stderrTask.Result
        if ($proc.ExitCode -ne 0) {
            return [pscustomobject]@{ Hits = @(); Status = 'unavailable' }
        }
        return [pscustomobject]@{
            Hits   = @(ConvertFrom-TriageMemoryVaultText -Text $stdout)
            Status = 'ok'
        }
    }
    catch {
        return [pscustomobject]@{ Hits = @(); Status = 'unavailable' }
    }
    finally {
        if ($proc) { $proc.Dispose() }
    }
}

function Add-TriageMemoryVaultDatabaseEnv {
    param($StartInfo, [string]$Root)
    $keys = @('DB_HOST', 'DB_PORT', 'DB_NAME', 'DB_USER', 'DB_PASSWORD')
    $apply = {
        param($Name, $Value)
        if ([string]::IsNullOrWhiteSpace($Name) -or [string]::IsNullOrWhiteSpace([string]$Value)) { return }
        if ($keys -notcontains $Name) { return }
        try { $StartInfo.EnvironmentVariables[$Name] = [string]$Value } catch { }
    }
    $envFile = Join-Path $Root '.env'
    if (Test-Path -LiteralPath $envFile) {
        foreach ($line in @(Get-Content -LiteralPath $envFile -ErrorAction SilentlyContinue)) {
            if ($line -match '^\s*(DB_[A-Za-z0-9_]+)\s*=\s*(.*)$') {
                & $apply $Matches[1] ($Matches[2].Trim().Trim('"').Trim("'"))
            }
        }
    }
    $currentHost = ''
    try { $currentHost = [string]$StartInfo.EnvironmentVariables['DB_HOST'] } catch { }
    if ($currentHost -and $currentHost -ne 'localhost') { return }
    $mcpPath = Join-Path $env:USERPROFILE '.cursor\mcp.json'
    if (-not (Test-Path -LiteralPath $mcpPath)) { return }
    try {
        $doc = Get-Content -LiteralPath $mcpPath -Raw -Encoding UTF8 | ConvertFrom-Json
        $servers = Get-TriageProp $doc @('mcpServers')
        $vault = Get-TriageProp $servers @('memory-vault')
        $envBlock = Get-TriageProp $vault @('env')
        if (-not $envBlock) { return }
        foreach ($key in $keys) {
            & $apply $key (Get-TriageProp $envBlock @($key))
        }
    }
    catch { }
}

function Add-TriageMemberberryClientNote {
    param(
        [Parameter(Mandatory)][string]$Client,
        [Parameter(Mandatory)][string]$Note,
        [string]$Root = 'C:\Git\memberberry'
    )
    $clientName = $Client.Trim()
    $line = ($Note -replace '[\r\n]+', ' ').Trim()
    if (-not $clientName) { throw 'Enter the Memberberry client name before saving.' }
    if (-not $line) { throw 'Enter the note to add before saving.' }
    $python = Join-Path $Root 'mcp\.venv\Scripts\python.exe'
    $package = Join-Path $Root 'mcp\src'
    if (-not (Test-Path -LiteralPath $python)) { throw 'Memberberry is not available on this PC.' }
    if (-not (Test-Path -LiteralPath (Join-Path $package 'memberberry_mcp\store.py'))) { throw 'Memberberry is not available on this PC.' }
    $code = @'
import json, os, sys
sys.path.insert(0, os.environ["MB_SRC"])
from memberberry_mcp import store
data = store.load_exceptions()
key = store.find_client_key(data, os.environ["MB_CLIENT"])
if not key:
    print(json.dumps({"saved": False, "error": "No Memberberry client matched."}))
    raise SystemExit(2)
entry = data.get(key) if isinstance(data.get(key), dict) else {}
notes = str(entry.get("notes") or "")
line = os.environ["MB_NOTE"].strip()
if line and line not in notes:
    notes = (notes.rstrip() + "\n" + line).strip()
data[key] = store.normalize_client_payload({"notes": notes}, existing=entry)
store.atomic_write_exceptions(data)
print(json.dumps({"saved": True, "client": key}))
'@
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $python
    $psi.WorkingDirectory = $Root
    $psi.UseShellExecute = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.CreateNoWindow = $true
    [void]$psi.ArgumentList.Add('-c')
    [void]$psi.ArgumentList.Add($code)
    $psi.EnvironmentVariables['MB_SRC'] = $package
    $psi.EnvironmentVariables['MB_CLIENT'] = $clientName
    $psi.EnvironmentVariables['MB_NOTE'] = $line
    $psi.EnvironmentVariables['MEMBERBERRY_ROOT'] = $Root
    $proc = New-Object System.Diagnostics.Process
    $proc.StartInfo = $psi
    try {
        [void]$proc.Start()
        $outTask = $proc.StandardOutput.ReadToEndAsync()
        $errTask = $proc.StandardError.ReadToEndAsync()
        if (-not $proc.WaitForExit(20000)) {
            try { $proc.Kill() } catch { }
            throw 'Memberberry did not finish saving.'
        }
        $out = [string]$outTask.Result
        [void]$errTask.Result
        if ($proc.ExitCode -ne 0) {
            $parsed = $null
            try { $parsed = $out | ConvertFrom-Json } catch { }
            $message = ''
            if ($parsed) { $message = [string](Get-TriageProp $parsed @('error')) }
            if (-not $message) { $message = 'No Memberberry client matched.' }
            throw $message
        }
        return ($out | ConvertFrom-Json)
    }
    finally {
        $proc.Dispose()
    }
}

function Add-TriageMemoryVaultNote {
    param(
        [Parameter(Mandatory)][string]$Text,
        [string]$Space = 'soc',
        [string]$Root = ''
    )
    $line = ($Text -replace '[\r\n]+', ' ').Trim()
    if (-not $line) { throw 'Enter the memory to add before saving.' }
    $vault = Get-TriageMemoryVaultCommand -Root $Root
    if (-not $vault) { throw 'Memory vault is not available on this PC.' }
    $spaceName = if ([string]::IsNullOrWhiteSpace($Space)) { 'soc' } else { $Space.Trim() }
    $code = @'
import asyncio, json, os
policy = getattr(asyncio, "WindowsSelectorEventLoopPolicy", None)
if policy is not None:
    asyncio.set_event_loop_policy(policy())
from src.mcp.server import remember
print(asyncio.run(remember(os.environ["MV_TEXT"], space=os.environ.get("MV_SPACE") or "soc", source="eoa-web", speaker="human")))
'@
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $vault.Python
    $psi.WorkingDirectory = $vault.Root
    $psi.UseShellExecute = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.CreateNoWindow = $true
    [void]$psi.ArgumentList.Add('-c')
    [void]$psi.ArgumentList.Add($code)
    try { $psi.EnvironmentVariables['PYTHONPATH'] = $vault.Root } catch { }
    try { $psi.EnvironmentVariables['MV_TEXT'] = $line } catch { }
    try { $psi.EnvironmentVariables['MV_SPACE'] = $spaceName } catch { }
    Add-TriageMemoryVaultDatabaseEnv -StartInfo $psi -Root $vault.Root
    $proc = New-Object System.Diagnostics.Process
    $proc.StartInfo = $psi
    try {
        [void]$proc.Start()
        $outTask = $proc.StandardOutput.ReadToEndAsync()
        $errTask = $proc.StandardError.ReadToEndAsync()
        if (-not $proc.WaitForExit(50000)) {
            try { $proc.Kill() } catch { }
            throw 'Memory vault did not finish saving.'
        }
        $out = [string]$outTask.Result
        [void]$errTask.Result
        if ($proc.ExitCode -ne 0 -or -not $out.Trim()) { throw 'Memory vault did not store that note.' }
        $parsed = $out | ConvertFrom-Json
        $stored = Get-TriageProp $parsed @('stored')
        $duplicate = Get-TriageProp $parsed @('duplicate')
        if ($stored -ne $true -and $duplicate -ne $true) {
            $message = [string](Get-TriageProp $parsed @('error', 'message'))
            if (-not $message) { $message = 'Memory vault did not store that note.' }
            throw $message
        }
        return $parsed
    }
    finally {
        $proc.Dispose()
    }
}

function Get-TriageMemoryVaultCommand {
    param([string]$Root = '')
    $candidate = $Root
    if ([string]::IsNullOrWhiteSpace($candidate)) { $candidate = 'C:\tools\memory-vault' }
    $python = Join-Path $candidate '.venv\Scripts\python.exe'
    if (-not (Test-Path -LiteralPath $python)) { return $null }
    if (-not (Test-Path -LiteralPath (Join-Path $candidate 'src\cli.py'))) { return $null }
    return [pscustomobject]@{ Python = $python; Root = $candidate }
}

function ConvertTo-TriageActionText {
    param($Actions)
    if ($null -eq $Actions) { return '' }
    if ($Actions -is [string]) { return $Actions.Trim() }
    $rows = @($Actions)
    $lines = New-Object System.Collections.Generic.List[string]
    foreach ($row in $rows) {
        if (-not $row) { continue }
        if ($row -is [string]) {
            if ($row.Trim()) { [void]$lines.Add(('- {0}' -f $row.Trim())) }
            continue
        }
        $when = [string](Get-TriageProp $row @('Timestamp', 'timestamp'))
        $upn = [string](Get-TriageProp $row @('UPN', 'upn'))
        $action = [string](Get-TriageProp $row @('Action', 'action'))
        $result = [string](Get-TriageProp $row @('Result', 'result'))
        $detail = [string](Get-TriageProp $row @('Detail', 'detail'))
        if (-not $action.Trim()) { continue }
        [void]$lines.Add(('- {0} {1} {2} {3} {4}' -f $when, $upn, $action, $result, $detail).Trim())
    }
    return ($lines -join "`n").Trim()
}

function Remove-TriageOperationalNotes {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return '' }
    $paragraphs = [regex]::Split([string]$Text, "(?:`r?`n){2,}")
    $kept = New-Object System.Collections.Generic.List[string]
    foreach ($para in $paragraphs) {
        $block = $para.Trim()
        if (-not $block) { continue }
        if ($block -match '(?i)REMEDIATE_FAILED|Unknown remediation command|Operational note:') { continue }
        if ($block -match '(?i)remediation command' -and $block -match '(?i)unknown command|was not recognized|reconcile') { continue }
        if ($block -match '(?i)not yet checked|have not been checked|were not collected|not been collected' -and $block -notmatch '(?i)OneDrive|SIEM') { continue }
        $lines = New-Object System.Collections.Generic.List[string]
        foreach ($line in ($block -split "`r?`n")) {
            if ($line -match '(?i)REMEDIATE_FAILED|Unknown remediation command|Operational note:') { continue }
            if ($line -match '(?i)^\s*REPULL:\s*\d+\b') { continue }
            [void]$lines.Add($line)
        }
        $clean = (($lines -join "`n").Trim())
        if ($clean) { [void]$kept.Add($clean) }
    }
    return (($kept -join "`n`n").Trim())
}

function Test-TriageCompromiseEvidence {
    param([string]$Text)
    $t = [string]$Text
    if ($t -match '(?i)malicious sign-?in|unauthorized sign-?in|attacker signed in') { return $true }
    if ($t -match '(?i)malicious inbox rule|(hidden|suspicious) inbox rule') { return $true }
    if ($t -match '(?i)inbox rule.{0,120}(forward|redirect|delet)') { return $true }
    if ($t -match '(?i)credential exposure|credentials? (were |was )?(exposed|leaked|stolen)|password (was |is )?(found |listed )?(in|on) a (breach|dump)|breached password|password in a breach') { return $true }
    return $false
}

function Remove-TriageUngatedCredentialReset {
    param([string]$Text)
    $raw = [string]$Text
    if (Test-TriageCompromiseEvidence -Text $raw) { return $raw.Trim() }
    $kept = New-Object System.Collections.Generic.List[string]
    foreach ($line in ($raw -split "`r?`n")) {
        if ($line -match '(?i)\b(reset|change)\b.{0,40}\bpassword\b') { continue }
        if ($line -match '(?i)\bpassword\b.{0,40}\b(reset|change)\b') { continue }
        if ($line -match '(?i)session revocation|revoke .{0,30}sessions|sign out everywhere|signing the user out') { continue }
        [void]$kept.Add($line)
    }
    return (($kept -join "`n").Trim())
}

function New-TriageClientEmailBody {
    param(
        [string]$Review,
        [string]$Contact = '',
        [string]$Greeting = 'Hello',
        [string]$Technician = '',
        [string]$Phone = '',
        [string]$GreetingName = ''
    )
    $clean = Remove-TriageOperationalNotes $Review
    $clean = ($clean -replace '(?im)^REPULL:\s*\d+\b.*$', '').Trim()
    $clean = Remove-TriageUngatedCredentialReset -Text $clean
    if ($clean.Length -gt 1200) { $clean = $clean.Substring(0, 1200).Trim() }
    $name = ''
    if ($GreetingName.Trim()) { $name = $GreetingName.Trim() }
    elseif ($Contact.Trim()) { $name = (($Contact.Trim() -split '\s+')[0]) }
    $hello = if ($name) { "$Greeting $name," } else { "$Greeting," }
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine($hello)
    [void]$sb.AppendLine('')
    if ($clean) { [void]$sb.AppendLine($clean) } else { [void]$sb.AppendLine('We reviewed the mailbox from the ticket. The findings are in the internal note.') }
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('Thanks,')
    if ($Technician.Trim()) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine($Technician.Trim())
    }
    if ($Phone.Trim()) { [void]$sb.AppendLine($Phone.Trim()) }
    return $sb.ToString().Trim()
}

function New-TriageMemberberryNoteBrief {
    param(
        [string]$Review,
        [string]$Ticket = '',
        [string]$TicketNumber = '',
        [string]$Company = '',
        [string]$Contact = '',
        [string]$Greeting = 'Hello',
        [string]$LocalTime = '',
        $Settings,
        [string]$ClientBrief = '',
        [string]$MemoryBrief = '',
        [string]$Audience = 'client',
        [string]$ActionsText = '',
        [string]$GreetingName = '',
        [string]$Revision = '',
        [string]$CurrentDrafts = ''
    )
    $audienceName = switch ($Audience) {
        'internal' { 'internal' }
        'incident' { 'incident' }
        default { 'client' }
    }
    $sb = New-Object System.Text.StringBuilder
    if ($audienceName -eq 'internal') {
        [void]$sb.AppendLine('Write the internal ConnectWise ticket note from this finished investigation. This is not the client email. Do not greet the reader.')
    }
    elseif ($audienceName -eq 'incident') {
        [void]$sb.AppendLine('Write an incident report for legal, management, and insurance. This is not a client email. Do not greet the reader.')
    }
    else {
        [void]$sb.AppendLine('Write the client email from this finished investigation. The reader is not technical. Use a friendly, everyday tone.')
        [void]$sb.AppendLine('The review below may name logs that failed, were not collected, or were limited. Leave every one of those out of this email: no technical problem, hiccup, missing record, or piece we could not confirm. Base the email only on what was checked, then give the recommendation.')
        [void]$sb.AppendLine('Do not ask the client to change or reset a password, and do not suggest signing them out everywhere, unless the investigation outcome shows a malicious sign-in, a malicious inbox rule, or confirmed credential exposure. A password in a breach list is confirmed credential exposure. Failed attempts alone are not. Only ask the client to change MFA when the review shows someone actually got in, such as a successful unauthorized sign-in or a rule, forward, or app added by an attacker. Never suggest stronger MFA or Conditional Access when nothing got in.
The goal is to close the ticket, not to worry the client. Keep the email short and calm. State the outcome plainly. Do not speculate about what might have happened, do not describe possible risks that were not found, and do not use words like strongly, concerned, worrying, or at risk when nothing got in. Do not offer to walk them through anything or invite more questions beyond the phone number in the signature.')
    }
    [void]$sb.AppendLine((Get-TriageEvidenceRules))
    if ($TicketNumber.Trim()) { [void]$sb.AppendLine(('Ticket number: {0}' -f $TicketNumber.Trim())) }
    if ($Company.Trim()) { [void]$sb.AppendLine(('Company: {0}' -f $Company.Trim())) }
    if ($Contact.Trim()) {
        [void]$sb.AppendLine(('Client contact: {0}' -f $Contact.Trim()))
    }
    if ($audienceName -eq 'client') {
        if ($Contact.Trim()) {
            $hi = if ($GreetingName.Trim()) { $GreetingName.Trim() } else { (($Contact.Trim() -split '\s+')[0]) }
            [void]$sb.AppendLine(('Greeting line: {0} {1},' -f $Greeting, $hi))
        }
        else {
            [void]$sb.AppendLine(('Greeting line: {0},' -f $Greeting))
            [void]$sb.AppendLine('No client contact was supplied. Use only that greeting. Do not invent a name. Do not greet the mailbox that was investigated.')
        }
    }
    if ($LocalTime.Trim()) { [void]$sb.AppendLine(('Local time: {0}' -f $LocalTime.Trim())) }
    $tech = ''
    $phone = ''
    $tz = ''
    if ($Settings) {
        $tech = [string](Get-TriageProp $Settings @('TechnicianName'))
        $phone = [string](Get-TriageProp $Settings @('ContactNumber'))
        $tz = [string](Get-TriageProp $Settings @('Timezone'))
    }
    if ($tech) { [void]$sb.AppendLine(('Technician name: {0}' -f $tech)) }
    if ($phone) { [void]$sb.AppendLine(('Contact number: {0}' -f $phone)) }
    if ($tz) { [void]$sb.AppendLine(('Timezone: {0}' -f $tz)) }
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('Memberberry client exceptions (reference for names, VIP, and known patterns):')
    if ($ClientBrief.Trim()) { [void]$sb.AppendLine($ClientBrief.Trim()) }
    else { [void]$sb.AppendLine('No Memberberry client matched. Do not invent client facts.') }
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('Memory vault, soc space (reference only, not facts about this ticket):')
    if ($MemoryBrief.Trim()) { [void]$sb.AppendLine($MemoryBrief.Trim()) }
    else { [void]$sb.AppendLine('No memory vault notes were available.') }
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('Remediation steps taken in this session:')
    if ($ActionsText.Trim()) { [void]$sb.AppendLine($ActionsText.Trim()) }
    else { [void]$sb.AppendLine('No remediation steps were run in this session.') }
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('Investigation outcome:')
    [void]$sb.AppendLine((Get-TriageClippedText (Remove-TriageOperationalNotes $Review) 8000))
    if ($Revision.Trim()) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('Technician revision. Rewrite the notes to follow this. It overrides the default tone where they conflict. The evidence rules still apply unless this revision explicitly asks for other wording. Do not add a true positive or a password reset unless the investigation outcome already contains the required evidence, or this revision explicitly asks for that wording. Do not add facts that are not in the investigation outcome.')
        [void]$sb.AppendLine((Get-TriageEvidenceRules))
        if ($CurrentDrafts.Trim()) {
            [void]$sb.AppendLine('Current drafts:')
            [void]$sb.AppendLine((Get-TriageClippedText $CurrentDrafts 8000))
        }
        [void]$sb.AppendLine('Change the notes as follows:')
        [void]$sb.AppendLine((Get-TriageClippedText $Revision 2000))
    }
    if ($Ticket.Trim()) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('Ticket text:')
        [void]$sb.AppendLine((Get-TriageClippedText $Ticket 4000))
    }
    return $sb.ToString().Trim()
}

function ConvertFrom-TriageMemberberryNote {
    param([string]$Text)
    $subject = ''
    $classification = ''
    $suggestion = ''
    $bodyLines = New-Object System.Collections.Generic.List[string]
    $inSuggestion = $false
    foreach ($line in ([string]$Text -split "`r?`n")) {
        if ($line -match '^(?i)\s*MEMBERBERRY EXCEPTION SUGGESTION:\s*(.*)$') {
            $inSuggestion = $true
            $suggestion = $Matches[1].Trim()
            continue
        }
        if ($inSuggestion) {
            if ($line.Trim()) {
                $suggestion = $(if ($suggestion) { $suggestion + ' ' + $line.Trim() } else { $line.Trim() })
            }
            continue
        }
        if (-not $subject -and $line -match '^(?i)\s*Subject:\s*(.+)$') {
            $subject = $Matches[1].Trim()
            continue
        }
        if (-not $classification -and $line -match '^(?i)\s*Classification:\s*(.+)$') {
            $classification = $Matches[1].Trim()
            continue
        }
        [void]$bodyLines.Add($line)
    }
    $body = (($bodyLines -join "`n").Trim())
    $canonical = ''
    foreach ($label in @('True Positive', 'False Positive', 'Authorized Activity')) {
        if ($classification -match [regex]::Escape($label)) { $canonical = $label; break }
    }
    if (-not $canonical) { $canonical = $classification }
    return [pscustomobject]@{
        Subject        = $subject
        Classification = $canonical
        Body           = $body
        Suggestion     = $suggestion
    }
}

function Get-TriageAlertLogKey {
    param([string]$Line)
    $key = ([string]$Line).ToLowerInvariant()
    $key = [regex]::Replace($key, '\b[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\b', '{id}')
    $key = [regex]::Replace($key, '\d{4}-\d{2}-\d{2}[t ]\d{2}:\d{2}:\d{2}(?:\.\d+)?(?:z|[+-]\d{2}:?\d{2})?', '{time}')
    $key = [regex]::Replace($key, '\b\d{1,2}/\d{1,2}/\d{2,4}\s+\d{1,2}:\d{2}(?::\d{2})?\s*(?:am|pm)?\b', '{time}')
    $key = [regex]::Replace($key, '\b[0-9a-f]{32,}\b', '{hex}')
    return $key.Trim()
}

function Compress-TriageAlertLogs {
    param([string]$Text, [int]$MaxChars = 18000)
    $raw = ([string]$Text).Replace("`r`n", "`n").Replace("`r", "`n").Trim()
    if (-not $raw) {
        return [pscustomobject]@{ Text = ''; Grouped = 0; Kept = 0 }
    }
    $groups = [ordered]@{}
    $order = New-Object System.Collections.Generic.List[string]
    foreach ($line in ($raw -split "`n")) {
        $shown = $line.Trim()
        if (-not $shown) { continue }
        $key = Get-TriageAlertLogKey $shown
        if (-not $key) { $key = $shown.ToLowerInvariant() }
        if (-not $groups.Contains($key)) {
            $groups[$key] = [pscustomobject]@{ Count = 0; Sample = $shown }
            [void]$order.Add($key)
        }
        $groups[$key].Count++
    }
    $sb = New-Object System.Text.StringBuilder
    $grouped = 0
    $kept = 0
    foreach ($key in $order) {
        $item = $groups[$key]
        $piece = $item.Sample
        if ($item.Count -gt 1) {
            $grouped++
            $piece = '[{0} similar lines] {1}' -f $item.Count, $item.Sample
        }
        if ($sb.Length -gt 0 -and ($sb.Length + $piece.Length + 1) -gt $MaxChars) {
            [void]$sb.AppendLine()
            [void]$sb.Append('[Further alert text was cut after the lines above.]')
            break
        }
        if ($sb.Length -gt 0) { [void]$sb.AppendLine() }
        [void]$sb.Append($piece)
        $kept++
    }
    return [pscustomobject]@{ Text = $sb.ToString().Trim(); Grouped = $grouped; Kept = $kept }
}

function Get-TriageAlertReviewBanner {
    return @'
HATZ-AI POWERED REVIEW OF THE ALERT ONLY
This note reviews the alert text on the ticket. It is not an investigation. No mailbox, sign-in, or device logs were pulled for this note. The technician must perform their own investigation before acting, closing the ticket, or emailing the client.
'@.Trim()
}

function Add-TriageAlertReviewBanner {
    param([string]$Text)
    $banner = Get-TriageAlertReviewBanner
    $body = ([string]$Text).Trim()
    if ($body -match '(?i)^HATZ-AI POWERED REVIEW OF THE ALERT ONLY') {
        $lines = @($body -split "`r?`n")
        $start = 1
        while ($start -lt $lines.Count -and -not $lines[$start].Trim()) { $start++ }
        if ($start -lt $lines.Count -and $lines[$start] -match '(?i)^This note reviews the alert text') { $start++ }
        while ($start -lt $lines.Count -and -not $lines[$start].Trim()) { $start++ }
        $body = if ($start -lt $lines.Count) { ($lines[$start..($lines.Count - 1)] -join "`n").Trim() } else { '' }
    }
    if (-not $body) { return $banner }
    return ($banner + "`n`n" + $body).Trim()
}

function Read-TriageMemberberryDocument {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path)) { return $null }
    $item = Get-Item -LiteralPath $Path
    $ticks = $item.LastWriteTimeUtc.Ticks
    $cached = $script:TriageMemberberryDoc[$Path]
    if ($cached -and $cached.Ticks -eq $ticks) { return $cached.Doc }
    try {
        $doc = Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json
    }
    catch { return $null }
    $script:TriageMemberberryDoc[$Path] = @{ Ticks = $ticks; Doc = $doc }
    return $doc
}

function Get-TriageMemberberryGlobalRecord {
    param([string]$ExceptionsPath)
    $doc = Read-TriageMemberberryDocument -Path $ExceptionsPath
    if (-not $doc) { return $null }
    foreach ($prop in @($doc.PSObject.Properties)) {
        if ($prop.Name -eq '_global' -or $prop.Name -eq 'global') { return $prop.Value }
    }
    return $null
}

function ConvertTo-TriageMemberberryExceptionBrief {
    param($Record)
    if (-not $Record) { return '' }
    $lines = New-Object System.Collections.Generic.List[string]
    foreach ($pair in @(
            @('false_positive_patterns', 'Known false positives'),
            @('authorized_tools', 'Authorized tools'),
            @('blocked_tools', 'Blocked tools')
        )) {
        $value = Get-TriageProp $Record @($pair[0])
        $items = @()
        if ($value -is [System.Array] -or ($null -ne $value -and $value -is [System.Collections.IList])) { $items = @($value) }
        elseif ($value) { $items = @([string]$value) }
        $clean = @(foreach ($item in $items) { $text = ([string]$item).Trim(); if ($text) { $text } })
        if (-not $clean.Count) { continue }
        [void]$lines.Add(('{0}:' -f $pair[1]))
        foreach ($item in $clean) { [void]$lines.Add(('- {0}' -f $item)) }
    }
    $notes = [string](Get-TriageProp $Record @('notes'))
    $accountLines = @(foreach ($line in ($notes -split "`r?`n")) {
            $text = $line.Trim()
            if ($text -match '(?i)river run accounts') { $text }
        })
    if ($accountLines.Count) {
        [void]$lines.Add('River Run accounts:')
        foreach ($item in $accountLines) { [void]$lines.Add(('- {0}' -f $item)) }
    }
    return ($lines -join "`n").Trim()
}

function New-TriageAlertReviewBrief {
    param(
        [string]$Alert = '',
        [string]$TicketNumber = '',
        [string]$Company = '',
        [string]$Contact = '',
        [string]$LocalTime = '',
        $Settings,
        [string]$GlobalBrief = '',
        [string]$ClientBrief = '',
        [string]$ClientExceptions = '',
        [string]$MemoryBrief = ''
    )
    $packed = Compress-TriageAlertLogs -Text $Alert -MaxChars 7000
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('Check Memberberry exceptions before you decide. Compare the alert to the global lists first, then this client. A match is an exception, not a new finding. Then say only whether the alert itself already shows a true positive, a false positive, or authorized activity, and the check that would settle it. Do not retell the alert. Use only the alert text and these notes. Do not treat any of it as proof that a log was pulled.')
    if ($TicketNumber.Trim()) { [void]$sb.AppendLine(('Ticket number: {0}' -f $TicketNumber.Trim())) }
    if ($Company.Trim()) { [void]$sb.AppendLine(('Company: {0}' -f $Company.Trim())) }
    if ($Contact.Trim()) { [void]$sb.AppendLine(('Ticket contact: {0}' -f $Contact.Trim())) }
    if ($LocalTime.Trim()) { [void]$sb.AppendLine(('Local time: {0}' -f $LocalTime.Trim())) }
    $tech = ''
    if ($Settings) { $tech = [string](Get-TriageProp $Settings @('TechnicianName')) }
    if ($tech) { [void]$sb.AppendLine(('Technician name: {0}' -f $tech)) }
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('Global Memberberry exceptions (every client):')
    if ($GlobalBrief.Trim()) { [void]$sb.AppendLine($GlobalBrief.Trim()) }
    else { [void]$sb.AppendLine('No global Memberberry exceptions were available.') }
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('This client in Memberberry:')
    if ($ClientBrief.Trim()) { [void]$sb.AppendLine($ClientBrief.Trim()) }
    else { [void]$sb.AppendLine('No Memberberry client matched. Do not invent client facts.') }
    if ($ClientExceptions.Trim()) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('This client''s exception lists:')
        [void]$sb.AppendLine($ClientExceptions.Trim())
    }
    if ($MemoryBrief.Trim()) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('Memory vault, soc space (reference only, not a finding about this alert):')
        [void]$sb.AppendLine($MemoryBrief.Trim())
    }
    [void]$sb.AppendLine('')
    if ($packed.Grouped -gt 0) {
        [void]$sb.AppendLine(('Repeated log lines were grouped first. {0} pattern(s) occurred more than once. A line that starts with [N similar lines] means that pattern appeared N times in the alert. Different IP addresses, users, and hosts were left as separate lines.' -f $packed.Grouped))
        [void]$sb.AppendLine('')
    }
    [void]$sb.AppendLine('Alert text:')
    if ($packed.Text) { [void]$sb.AppendLine($packed.Text) }
    else { [void]$sb.AppendLine('The ticket had no alert text.') }
    return $sb.ToString().Trim()
}

function Split-TriageTechnicianRevision {
    param([string]$Review)
    $revision = ''
    $drafts = ''
    $body = [string]$Review
    if ($body -match '(?s)^(.*?)\r?\n<<<TECHNICIAN_REVISION>>>\r?\n(.*)$') {
        $body = $Matches[1].Trim()
        $rest = $Matches[2]
        if ($rest -match '(?s)^(.*?)\r?\n<<<CURRENT_DRAFTS>>>\r?\n(.*)$') {
            $revision = $Matches[1].Trim()
            $drafts = $Matches[2].Trim()
        }
        else { $revision = $rest.Trim() }
    }
    return [pscustomobject]@{ Review = $body; Revision = $revision; Drafts = $drafts }
}

function Invoke-TriageMemberberryNote {
    param(
        [Parameter(Mandatory)][string]$SettingsPath,
        [string]$Review = '',
        [string]$Ticket = '',
        [string]$TicketNumber = '',
        [string]$Company = '',
        [string]$Contact = '',
        [string]$Audience = 'all',
        $Actions,
        [string]$MemberberryRoot = '',
        [string]$MemoryVaultRoot = '',
        [string]$Revision = '',
        [string]$CurrentDrafts = ''
    )
    $wantedFirst = $Audience.Trim().ToLowerInvariant()
    if ($wantedFirst -eq 'alert') {
        if ([string]::IsNullOrWhiteSpace($Review) -and [string]::IsNullOrWhiteSpace($Ticket)) {
            throw 'Open a ticket with alert text before triaging.'
        }
    }
    elseif ([string]::IsNullOrWhiteSpace($Review)) {
        throw 'Finish a review before writing the ticket note.'
    }
    $split = Split-TriageTechnicianRevision -Review $Review
    $Review = [string]$split.Review
    if (-not $Revision.Trim()) { $Revision = [string]$split.Revision }
    if (-not $CurrentDrafts.Trim()) { $CurrentDrafts = [string]$split.Drafts }
    $stored = Read-TriageAiSettingsFile -Path $SettingsPath
    $provider = [string]$stored.provider
    if ($provider -notin @('chatgpt', 'claude', 'hatz')) {
        throw 'Choose ChatGPT, Claude, or Hatz AI in AI review settings.'
    }
    $key = [string]$stored.apiKey
    if ([string]::IsNullOrWhiteSpace($key)) {
        throw 'Save an API key in AI review settings. It stays on this PC.'
    }
    Test-TriageAiKeyToken -ApiKey $key
    $model = [string]$stored.model
    if (-not $model) { $model = Get-TriageAiDefaultModel -Provider $provider }
    $companyName = $Company.Trim()
    if (-not $companyName) { $companyName = Get-TriageTicketCompany -Ticket $Ticket }
    $contactName = $Contact.Trim()
    if (-not $contactName) { $contactName = Get-TriageTicketContact -Ticket $Ticket }
    $actionsText = ConvertTo-TriageActionText -Actions $Actions
    $wanted = $Audience.Trim().ToLowerInvariant()
    if ($wanted -in @('all', 'both')) { $wanted = 'all' }
    elseif ($wanted -eq 'alert') { $wanted = 'alert' }
    elseif ($wanted -notin @('internal', 'client', 'incident')) { $wanted = 'internal' }
    $audiences = if ($wanted -eq 'all') { @('internal', 'client', 'incident') } else { @($wanted) }
    $paths = Get-TriageMemberberryPaths -Root $MemberberryRoot
    $mbSettings = $null
    $clientBrief = ''
    $clientName = ''
    $client = $null
    $greetName = ''
    if ($paths) {
        if ($paths.Settings) { $mbSettings = Get-TriageMemberberrySettings -SettingsPath $paths.Settings }
        if ($companyName -and $paths.Exceptions) {
            $client = Find-TriageMemberberryClient -Company $companyName -ExceptionsPath $paths.Exceptions
            if ($client) {
                $clientName = [string]$client.Name
                $clientBrief = ConvertTo-TriageMemberberryClientBrief -Record $client.Record -Name $clientName -OmitLists:($wanted -eq 'alert')
                $greetName = Get-TriageMemberberryGreetingName -Record $client.Record -Contact $contactName
            }
        }
    }
    $globalBrief = ''
    $clientExceptionBrief = ''
    if ($wanted -eq 'alert' -and $paths -and $paths.Exceptions) {
        $globalBrief = ConvertTo-TriageMemberberryExceptionBrief -Record (Get-TriageMemberberryGlobalRecord -ExceptionsPath $paths.Exceptions)
        if ($client -and $client.Record) {
            $clientExceptionBrief = ConvertTo-TriageMemberberryExceptionBrief -Record $client.Record
        }
    }
    $memoryBrief = ''
    $memoryCount = 0
    $memoryStatus = 'skipped'
    $vault = if ($wanted -eq 'alert') { $null } else { Get-TriageMemoryVaultCommand -Root $MemoryVaultRoot }
    if ($vault) {
        $memorySeed = if ($Review.Trim()) { $Review } else { $Ticket }
        $query = (($companyName, $TicketNumber, (Get-TriageClippedText $memorySeed 240)) | Where-Object { $_ }) -join ' '
        $found = Invoke-TriageMemoryVaultSearch -Query $query -PythonPath $vault.Python -WorkingDirectory $vault.Root -Space 'soc' -Limit 4
        $memoryStatus = [string]$found.Status
        $blocks = New-Object System.Collections.Generic.List[string]
        foreach ($hit in @($found.Hits)) {
            if (-not $hit -or -not $hit.Text) { continue }
            $memoryCount++
            [void]$blocks.Add(('- [{0}] {1}' -f $hit.Space, $hit.Text))
        }
        $memoryBrief = ($blocks -join "`n")
    }
    $now = Get-Date
    $greeting = Get-TriageMemberberryGreeting -Now $now
    $tz = ''
    if ($mbSettings) { $tz = [string]$mbSettings.Timezone }
    $local = '{0:yyyy-MM-dd h:mm tt} {1}' -f $now, $tz
    $plain = @'
The reply is plaintext, not Markdown. Do not use HTML, asterisks, hash headings, or code fences. Never use an en dash, an em dash, or a double dash. Use a comma, a colon, parentheses, or a single hyphen.

Print exactly one Classification line.
Classification values allowed: True Positive, False Positive, Authorized Activity.
Use False Positive for a false positive, Authorized Activity for authorized activity, and True Positive when the event is a real security issue.
If the review begins with VERDICT, the Classification line must use that same choice.
If that verdict is true positive, copy it only when the investigation outcome contains concrete evidence of malicious activity. If the verdict is ask the client, or the outcome is inconclusive, do not write True Positive. Ask the client to confirm whether the activity was theirs or authorized.
The first sentence of the body states FALSE POSITIVE, AUTHORIZED ACTIVITY, LOW RISK, MEDIUM RISK, or HIGH RISK.
Use only the investigation outcome for what happened. Memberberry exceptions and memory vault notes are reference only. Do not invent names, actions, times, or incidents. A remediation step is taken only when it is listed under remediation steps taken. Anything else is recommended, not completed. Do not mention investigation tool errors, command names, REMEDIATE_FAILED, an unknown remediation command, or a recommendation to reconcile a failed command. Those lines are not findings and do not belong in the note, the client email, or the incident report. Every clock time you write is local time on this PC. Use the local time and timezone from the brief. Do not write UTC, a trailing Z, or a numeric offset. If the investigation outcome shows a time in UTC, rewrite it in that local time before it appears in the note, the client email, or the incident report.
'@
    $plain = $plain + "`n`n" + (Get-TriageEvidenceRules)
    $systems = @{
        internal = @"
You write the internal ConnectWise ticket note after an investigation is finished. This note stays on the ticket. It is not the email to the client. Do not greet anyone.

$plain

Subject: <subject, include the ticket number when one was supplied>
Classification: True Positive

Then the note. State the mailbox, the evidence, and what was done. Close with the technician name from the brief on its own line. Do not add the company name under the signature.

MEMBERBERRY EXCEPTION SUGGESTION: <client> - <one sentence>
MEMORY VAULT: <one sentence the next review should remember>
Omit a line when you have no concrete addition.
"@
        client = @"
You write the email the technician sends to the client after an investigation. The reader works at the company and is not an IT person. Write it the way a friendly person would explain it on the phone: warm, calm, and easy to follow.

If a contact name is supplied, greet that contact by first name, or by the preferred name from exceptions. If no contact name is supplied, the greeting is only the time-of-day greeting and a comma, such as Good morning,. Do not invent a name. Do not greet the mailbox that was investigated unless that person is the contact.

The reply is plaintext, not Markdown. Do not use HTML, asterisks, hash headings, or code fences. Never use an en dash, an em dash, or a double dash. Use a comma, a colon, parentheses, or a single hyphen.

Print exactly one Classification line. That line is for our records. Do not repeat it, or any risk label, in the email the client reads.
Classification values allowed: True Positive, False Positive, Authorized Activity.
Use False Positive for a false positive, Authorized Activity for authorized activity, and True Positive when the event is a real security issue.
If the review begins with VERDICT, the Classification line must use that same choice.
If that verdict is true positive, copy it only when the investigation outcome contains concrete evidence of malicious activity. If the verdict is ask the client, or the outcome is inconclusive, do not write True Positive. Ask the client to confirm whether the activity was theirs or authorized.
Use only the investigation outcome for what happened. Memberberry exceptions and memory vault notes are reference only. Do not invent names, actions, times, or incidents. A step is done only when it is listed under remediation steps taken. Anything else is something we are suggesting, not something we already did. Do not mention investigation tool errors, command names, REMEDIATE_FAILED, an unknown remediation command, or a recommendation to reconcile a failed command. Those lines are not findings and do not belong in the email. Every clock time you write is local time on this PC. Use the local time and timezone from the brief. Do not write UTC, a trailing Z, or a numeric offset. If the investigation outcome shows a time in UTC, rewrite it in that local time before it appears in the email.

Classification: True Positive

Then the email body. Do not write a subject line. This is the message the technician pastes to the client. Start with the greeting line from the brief. Say what we looked into, what it means for them, and what they should do next. LOW detail clients get two or three short sentences. HIGH detail clients may use a few short paragraphs, still in plain language. Use everyday words. Do not write FALSE POSITIVE, TRUE POSITIVE, AUTHORIZED ACTIVITY, LOW RISK, MEDIUM RISK, HIGH RISK, BEC, IOC, remediation, inbox rule, message trace, audit, SPF, DKIM, DMARC, or sign-in. Do not mention a search window, a lookback, or that a record was older than the period that was reviewed. Do not offer to review a longer stretch of time; the check already pulled the history the account keeps. The investigation is finished when this email is written. Do not say we will keep looking, take a closer look, dig deeper, keep working on it, or follow up as we learn more. Do not say some checks did not come through, could not be completed, or were limited by the account or its license; those gaps belong in the internal note, not the client email. State what we found and what we recommend, based on what was checked. If a record was not found, say you could not find it. When you need one of those ideas, say it plainly, for example: we checked the mailbox, we did not find a problem, a setting was sending copies of their mail to another address and we turned that off, please do not use the link in that message, or please give us a call. Do not use the headings ACTION TAKEN or ACTION REQUIRED. If we already did something, say it in a normal sentence, and only when that step is listed under remediation steps taken. If the client still needs to do something because someone actually got in, ask for it kindly, and do not say that nothing else is needed. If nothing got in, do not suggest a password change, a password reset, signing out everywhere, an MFA change, Conditional Access, or a call to tighten security. The email should close the matter: short, calm, no speculation about risks that were not found, and no offers to walk through anything. If no remediation steps were run, do not say we already fixed it. Close with Thanks, then the technician name from the brief on its own line, then the contact number from the brief on its own line. That number is the River Run Support line. Do not add the company name under the signature. The body must not be empty.

MEMBERBERRY EXCEPTION SUGGESTION: <client> - <one sentence>
MEMORY VAULT: <one sentence the next review should remember>
Omit a line when you have no concrete addition.
EVIDENCE RULES
"@
        incident = @"
You write an incident report for legal, management, and insurance. It is not a client email. Do not greet anyone. If a fact is unknown, write unknown. Do not invent dollar amounts, recipients, or actions. Log times in the investigation are the technician local time named in the brief. Write the timeline in that local time. Do not write UTC, a trailing Z, or a numeric offset.

$plain

Subject: Incident report for the mailbox, and the ticket number when one was supplied
Classification: True Positive

Then these sections. Put each label on its own line in capital letters, then the text. Leave a blank line between sections.

INCIDENT SUMMARY
AFFECTED ACCOUNT
TIMELINE
WHAT WAS FOUND
REMEDIATION TAKEN
Only steps listed as taken. If none were run, write that no remediation steps were run in this session.
REMEDIATION RECOMMENDED
Steps the investigation suggested that are not in the taken list. Do not recommend collecting unified audit, file sharing, DLP, or sign-in logs. The check downloads those. Use the rows in the pack. An empty list was checked and had no rows. If the pack says a source was not collected, say that once under WHAT WAS FOUND, not as a task for the reader.
REMAINING RISK
PREPARED BY
The technician name and the local time from the brief.

Do not add a DELIVERED BLAST RECIPIENTS section. When an outbound email blast was delivered, the address list is appended after your report. Each delivered recipient is on its own line. Do not join those addresses with commas or semicolons, and do not invent recipients.

Do not include a MEMBERBERRY EXCEPTION SUGGESTION line.
MEMORY VAULT: <one sentence the next review should remember>
Omit that line when you have nothing concrete.
"@
        alert = @'
You write a short internal note for the technician who still has to investigate this alert. You have only the alert text. You did not pull mailbox, sign-in, or device logs. Do not greet anyone. Do not write a client email. Do not retell the alert, list every user, IP, file, or host, or paste log lines.

The reply is plaintext, not Markdown. Do not use HTML, asterisks, hash headings, or code fences. Never use an en dash, an em dash, or a double dash. Use a comma, a colon, parentheses, or a single hyphen.

Do not print a Classification line. Do not say the investigation is finished, the ticket can be closed, or the client was already contacted. Do not invent users, addresses, IPs, files, times, or log rows that are not in the alert. Check Memberberry before you decide. If a known false positive matches, the leaning is Possible false positive and you name that exception. A pattern can match a class of alerts, not only an exact product name. When the alert's finding is an unusual or impossible location and the alert identifies that IP as a mobile or cellular provider, that is the Memberberry mobile-network geo-IP false positive: shared egress addresses make the published location unreliable. Lean Possible false positive and name that exception. If the alert does not identify the provider, do not assume the IP is mobile. The next check is whether it belongs to a mobile carrier. Do not use that exception when the alert's concern is not the location. If an authorized tool matches, the leaning is Possible authorized activity and you name that tool. When the alert names a River Run account from the Memberberry list, including names like rrc, rradmin, and rrcadmin, say that actor is a River Run account. Lean Possible authorized activity only when the alert shows ordinary administrative work by that account. The username alone is not a false positive. If a blocked tool matches, say it stays blocked and treat that as evidence of a possible true positive. If nothing matches, say Memberberry has no matching exception. Memory vault notes are reference only. Do not summarize them.

Subject: Alert triage, and the ticket number when one was supplied

Then the note. Start with this line, using one of these four values only:
Alert leaning: Possible false positive
Alert leaning: Possible true positive
Alert leaning: Possible authorized activity
Alert leaning: Needs investigation

Then these labels, each on its own line. Leave a blank line between sections. Each section is at most three sentences.

MEMBERBERRY MATCH
One or two sentences. Name the global or client exception that matches, or say Memberberry has no matching exception.

EVIDENCE IN THE ALERT
Only the facts already written in the alert that support a false positive, a true positive, or authorized activity. Name the field or quote a short phrase. A repeated pattern matters only when the count itself is the evidence, and then say the count without quoting the lines. If the alert does not contain that evidence, write one sentence: the alert does not show whether this is a true positive, a false positive, or authorized activity.

NEXT CHECK
If the alert or a Memberberry match already points to one of those three, give the one confirmation the technician should still make before acting. If it does not, give the one or two checks that would settle which of the three it is. Each check is still to be done. Do not write a general investigation plan.

Close with the technician name from the brief on its own line. Do not add the company name under the signature.
Do not write the words HATZ-AI. That heading is added after you reply.
'@
    }
    $systems.client = $systems.client.Replace('EVIDENCE RULES', (Get-TriageEvidenceRules))
    $texts = @{ internal = ''; client = ''; incident = ''; alert = '' }
    $notes = @{}
    $storeClient = ''
    $storeNote = ''
    $storeMemory = ''
    $firstError = ''
    foreach ($aud in $audiences) {
        $maxTokens = 1200
        if ($aud -eq 'alert') {
            $source = if ($Review.Trim()) { $Review.Trim() } else { $Ticket.Trim() }
            if ($Ticket.Trim() -and $Review.Trim() -and $Ticket.Trim() -ne $Review.Trim()) {
                $source = $Review.Trim() + "`n`n" + $Ticket.Trim()
            }
            $brief = New-TriageAlertReviewBrief -Alert $source -TicketNumber $TicketNumber -Company $companyName -Contact $contactName -LocalTime $local.Trim() -Settings $mbSettings -GlobalBrief $globalBrief -ClientBrief $clientBrief -ClientExceptions $clientExceptionBrief -MemoryBrief $memoryBrief
            $maxTokens = 700
        }
        else {
            $brief = New-TriageMemberberryNoteBrief -Review $Review -Ticket $Ticket -TicketNumber $TicketNumber -Company $companyName -Contact $contactName -Greeting $greeting -LocalTime $local.Trim() -Settings $mbSettings -ClientBrief $clientBrief -MemoryBrief $memoryBrief -Audience $aud -ActionsText $actionsText -GreetingName $greetName -Revision $Revision -CurrentDrafts $CurrentDrafts
        }
        $systemPrompt = [string]$systems[$aud]
        if ($aud -eq 'incident' -and (Test-TriageConfirmedIncident -Text $Ticket)) {
            $systemPrompt += @'

The ticket already confirms this incident. Do not re-argue whether it happened. REMEDIATION TAKEN is only what the ticket or the investigation says was already done. REMEDIATION RECOMMENDED is what is still open. REMAINING RISK says the incident is fully remediated, or names what is still open.
'@
        }
        $request = Get-TriageAiReviewRequest -Provider $provider -Model $model -ApiKey $key -Brief $brief -SystemPrompt $systemPrompt -MaxTokens $maxTokens
        try {
            $response = Invoke-RestMethod -Method Post -Uri $request.Uri -Headers $request.Headers -Body $request.Body -ContentType 'application/json' -TimeoutSec 90
            $raw = ConvertTo-TriageAiPlainText -Text (ConvertFrom-TriageAiResponse -Response $response -Provider $provider)
            $raw = $raw -replace [char]0x2013, ', '
            $raw = $raw -replace [char]0x2014, ', '
            $raw = $raw -replace '\s+--\s+', ', '
            $raw = $raw.Trim()
            if (-not $raw) { throw 'The model returned an empty note.' }
            $suggested = ConvertFrom-TriageStoreSuggestions -Text $raw
            if (-not $storeClient -and $suggested.Client) { $storeClient = [string]$suggested.Client }
            if (-not $storeNote -and $suggested.Note) { $storeNote = [string]$suggested.Note }
            if (-not $storeMemory -and $suggested.Memory) { $storeMemory = [string]$suggested.Memory }
            $note = ConvertFrom-TriageMemberberryNote -Text $suggested.Text
            $noteBody = Remove-TriageOperationalNotes ([string]$note.Body)
            if ($aud -eq 'alert') { $noteBody = Add-TriageAlertReviewBanner $noteBody }
            if ($aud -eq 'client' -and $noteBody.Length -lt 40) {
                $techName = ''
                $phone = ''
                if ($mbSettings) {
                    $techName = [string](Get-TriageProp $mbSettings @('TechnicianName'))
                    $phone = [string](Get-TriageProp $mbSettings @('ContactNumber'))
                }
                $noteBody = New-TriageClientEmailBody -Review $Review -Contact $contactName -Greeting $greeting -Technician $techName -Phone $phone -GreetingName $greetName
            }
            $copy = New-Object System.Text.StringBuilder
            if ($aud -eq 'alert') {
                $banner = Get-TriageAlertReviewBanner
                if ($noteBody.StartsWith($banner)) { $noteBody = $noteBody.Substring($banner.Length).Trim() }
                $noteBody = [regex]::Replace($noteBody, '(?i)\s*A matched false positive is not a true positive\.?', '')
                $noteBody = $noteBody.Trim()
                [void]$copy.AppendLine($banner)
                [void]$copy.AppendLine('')
            }
            if ($aud -ne 'client' -and $note.Subject) { [void]$copy.AppendLine(('Subject: {0}' -f $note.Subject)) }
            if ($aud -ne 'client' -and $aud -ne 'alert' -and $note.Classification) { [void]$copy.AppendLine(('Classification: {0}' -f $note.Classification)) }
            if (($aud -ne 'client' -and $note.Subject) -or ($aud -ne 'client' -and $aud -ne 'alert' -and $note.Classification)) { [void]$copy.AppendLine('') }
            if ($noteBody) { [void]$copy.AppendLine($noteBody) }
            $texts[$aud] = (Remove-TriageOperationalNotes ($copy.ToString().Trim()))
            $notes[$aud] = $note
        }
        catch {
            $message = Get-TriageAiErrorMessage -ErrorRecord $_
            if (-not $firstError) { $firstError = $message }
            $texts[$aud] = ''
        }
    }
    if ($audiences -contains 'client' -and [string]::IsNullOrWhiteSpace([string]$texts.client)) {
        $techName = ''
        $phone = ''
        if ($mbSettings) {
            $techName = [string](Get-TriageProp $mbSettings @('TechnicianName'))
            $phone = [string](Get-TriageProp $mbSettings @('ContactNumber'))
        }
        $texts.client = New-TriageClientEmailBody -Review $Review -Contact $contactName -Greeting $greeting -Technician $techName -Phone $phone -GreetingName $greetName
    }
    $ready = @($audiences | Where-Object { $texts[$_] })
    if (-not $ready.Count) {
        if ($firstError) { throw $firstError }
        throw 'The model returned an empty ticket note.'
    }
    $primaryAud = if ($texts.alert) { 'alert' } elseif ($texts.internal) { 'internal' } elseif ($texts.client) { 'client' } else { 'incident' }
    $primary = $notes[$primaryAud]
    $ref = New-Object System.Collections.Generic.List[string]
    if ($clientName) { [void]$ref.Add(('Memberberry client: {0}' -f $clientName)) }
    elseif (-not $companyName) { [void]$ref.Add('No company name was available, so Memberberry was not searched') }
    else { [void]$ref.Add(('No Memberberry client matched {0}' -f $companyName)) }
    if ($memoryStatus -eq 'ok') { [void]$ref.Add(('{0} soc memory hit(s)' -f $memoryCount)) }
    elseif ($memoryStatus -eq 'unavailable') { [void]$ref.Add('Memory vault was not available') }
    else { [void]$ref.Add('Memory vault was not searched') }
    if ($firstError) { [void]$ref.Add(('One note failed: {0}' -f $firstError)) }
    return [pscustomobject]@{
        provider       = $provider
        model          = $model
        subject        = [string]$primary.Subject
        classification = [string]$primary.Classification
        body           = [string]$primary.Body
        suggestion     = [string]$primary.Suggestion
        text           = [string]$texts[$primaryAud]
        internalText   = $(if ($texts.alert) { [string]$texts.alert } else { [string]$texts.internal })
        clientText     = [string]$texts.client
        incidentText   = [string]$texts.incident
        confirmedIncident = [bool](Test-TriageConfirmedIncident -Text $Ticket)
        storeClient    = $storeClient
        storeNote      = $storeNote
        storeMemory    = $storeMemory
        client         = $clientName
        memories       = $memoryCount
        references     = ($ref -join '. ')
    }
}

Export-ModuleMember -Function Get-TriageFreemailDomains, Get-TriageLayer, Get-TriageHeaderText, ConvertFrom-TriageMessageHeader, Get-TriageClassification, Get-TriageImpersonationMatch, Merge-TriageTraceLegs, ConvertTo-TriageUtc, Get-TriageLocalZoneLabel, ConvertTo-TriagePresentedTime, Get-TriageDnsHost, Test-TriageMxPointsAtEop, Test-TriageConfigFindings, Get-TriageOrcaSummary, Get-TriageBlockTarget, New-TriageActionScript, New-TriageTicketNote, New-TriageClientEmail, Get-TriageCommandToken, Test-TriageConfigModuleInstalled, Update-TriagePackOutputs, Get-TriagePackView, Set-TriageVerdictOverride, Get-TriagePacks, Get-TriageClientNotes, Set-TriageClientNotes, Get-TriagePackFileContent, Invoke-TriageWorkerCommand, ConvertTo-TriageRows, Get-TriageRemediationItems, Invoke-TriageRemediation, Get-TriageAiSettings, Set-TriageAiSettings, New-TriageAiReviewBrief, Get-TriageAiReviewRequest, ConvertFrom-TriageAiResponse, ConvertTo-TriageAiPlainText, Invoke-TriageAiReview, Get-TriageBecRemediations, New-TriageBecReviewBrief, Invoke-TriageBecReview, Get-TriageBecWindow, ConvertFrom-TriageBecRepull, ConvertFrom-TriageBecRepullLogs, Get-TriageBecFollowUpPlan, Test-TriageConfirmedIncident, ConvertFrom-TriageStoreSuggestions, Add-TriageMemberberryClientNote, Add-TriageMemoryVaultNote, ConvertFrom-TriageLogPlanText, New-TriageLogPlanBrief, Get-TriageLogFolderInventory, Test-TriageLogReportFolder, Invoke-TriageLogPlan, Get-TriageTicketCompany, Get-TriageTicketContact, Get-TriageMemberberrySettings, ConvertTo-TriageMemberberryClientBrief, Find-TriageMemberberryClient, ConvertFrom-TriageMemoryVaultText, New-TriageMemberberryNoteBrief, ConvertFrom-TriageMemberberryNote, Invoke-TriageMemberberryNote, Remove-TriageOperationalNotes, New-TriageClientEmailBody, Compress-TriageAlertLogs, Add-TriageAlertReviewBanner, New-TriageAlertReviewBrief, Get-TriageMemberberryGlobalRecord, ConvertTo-TriageMemberberryExceptionBrief, Get-TriageBecVerdict, Get-TriageEvidenceRules, Get-TriageBecReviewSystemPrompt, Test-TriageCompromiseEvidence, Remove-TriageUngatedCredentialReset, ConvertFrom-TriageAiStreamDelta
