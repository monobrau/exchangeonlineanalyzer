# LogInvestigate.psm1
# Read-only query / mail-graph over EOA report-folder CSVs.
# Never modifies original export files.

$script:InvestigateSources = @(
    @{
        Name        = 'MessageTrace'
        TimeCols    = @('Received', 'receivedDateTime')
        AddressCols = @('SenderAddress', 'RecipientAddress', 'Sender', 'Recipient')
        IpCols      = @('FromIP', 'FromIp', 'ToIP', 'ToIp', 'IPAddress')
        Preset      = @('Received', 'SenderAddress', 'RecipientAddress', 'Subject', 'Status', 'FromIP', 'Size')
    }
    @{
        Name        = 'SignInLogs'
        TimeCols    = @('CreatedDateTime', 'createdDateTime', 'TimeGenerated')
        AddressCols = @('UserPrincipalName', 'UPN', 'User')
        IpCols      = @('IPAddress', 'IpAddress', 'IP')
        Preset      = @('CreatedDateTime', 'UserPrincipalName', 'IPAddress', 'AppDisplayName', 'Status', 'CountryOrRegion', 'ClientAppUsed')
    }
    @{
        Name        = 'UnifiedAuditLogs'
        TimeCols    = @('CreationDate', 'CreationTime', 'CreationDateTime')
        AddressCols = @('UserIds', 'UserId', 'UserPrincipalName', 'UserKey')
        IpCols      = @('ClientIP', 'ClientIp', 'IPAddress')
        Preset      = @('CreationDate', 'UserIds', 'Operations', 'RecordType', 'ClientIP', 'ResultStatus')
        EnrichUal   = $true
    }
)

$script:CsvCache = [ordered]@{}
$script:CsvCacheMax = 8

function Get-InvestigateProp {
    param($Object, [string]$Name)
    if ($null -eq $Object) { return $null }
    if ($Object -is [hashtable] -or $Object -is [System.Collections.IDictionary]) {
        if ($Object.ContainsKey($Name)) { return $Object[$Name] }
        foreach ($k in $Object.Keys) {
            if ([string]$k -eq $Name) { return $Object[$k] }
        }
        return $null
    }
    $p = $Object.PSObject.Properties[$Name]
    if ($p) { return $p.Value }
    return $null
}

function Get-InvestigateCsvPath {
    param(
        [Parameter(Mandatory)][string]$Folder,
        [Parameter(Mandatory)][string]$BaseName
    )
    if (-not (Test-Path -LiteralPath $Folder)) { return $null }

    $candidates = [System.Collections.Generic.List[string]]::new()
    $exact = Join-Path $Folder "$BaseName.csv"
    if (Test-Path -LiteralPath $exact) { [void]$candidates.Add($exact) }

    Get-ChildItem -LiteralPath $Folder -Filter "${BaseName}*.csv" -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -notmatch '_curated' } |
        Sort-Object Length -Descending |
        ForEach-Object {
            if (-not $candidates.Contains($_.FullName)) { [void]$candidates.Add($_.FullName) }
        }

    foreach ($p in $candidates) {
        try {
            $lines = Get-Content -LiteralPath $p -TotalCount 2 -ErrorAction Stop
            if ($lines -and $lines.Count -ge 2) { return $p }
        }
        catch { }
    }
    return $null
}

function Import-InvestigateCsv {
    param([Parameter(Mandatory)][string]$Path)
    try {
        return @(Import-Csv -LiteralPath $Path -Encoding UTF8 -ErrorAction Stop)
    }
    catch {
        return @(Import-Csv -LiteralPath $Path -ErrorAction Stop)
    }
}

function Get-InvestigateCachedRows {
    param(
        [Parameter(Mandatory)][string]$Path,
        [switch]$EnrichUal
    )
    $item = Get-Item -LiteralPath $Path
    $key = $item.FullName
    $hit = $script:CsvCache[$key]
    if ($hit -and $hit.Length -eq $item.Length -and $hit.WriteUtc -eq $item.LastWriteTimeUtc) {
        return $hit.Rows
    }

    $rows = Import-InvestigateCsv -Path $Path
    if ($EnrichUal) {
        foreach ($row in $rows) { Add-InvestigateUalFields -Row $row }
    }

    while ($script:CsvCache.Count -ge $script:CsvCacheMax) {
        $script:CsvCache.Remove($script:CsvCache.Keys[0])
    }
    $script:CsvCache[$key] = @{
        Length   = $item.Length
        WriteUtc = $item.LastWriteTimeUtc
        Rows     = $rows
    }
    return $rows
}

function Add-InvestigateUalFields {
    param($Row)
    if (-not $Row) { return }
    $audit = $null
    $raw = [string]$Row.AuditData
    if ($raw -and $raw.TrimStart().StartsWith('{')) {
        try { $audit = $raw | ConvertFrom-Json -ErrorAction Stop } catch { $audit = $null }
    }
    $op = [string]$Row.Operations
    if (-not $op -and $audit) { $op = [string]$audit.Operation }
    $ip = [string]$Row.ClientIP
    if (-not $ip) { $ip = [string]$Row.ClientIp }
    if (-not $ip -and $audit) { $ip = [string]$audit.ClientIP }
    $mbx = ''
    if ($audit) { $mbx = [string]$audit.MailboxOwnerUPN }
    $user = [string]$Row.UserIds
    if (-not $user -and $audit) { $user = [string]$audit.UserId }

    Add-Member -InputObject $Row -NotePropertyName '_Operation' -NotePropertyValue $op -Force
    Add-Member -InputObject $Row -NotePropertyName '_ClientIP' -NotePropertyValue $ip -Force
    Add-Member -InputObject $Row -NotePropertyName '_MailboxOwnerUPN' -NotePropertyValue $mbx -Force
    Add-Member -InputObject $Row -NotePropertyName '_User' -NotePropertyValue $user -Force
}

function Resolve-InvestigateColumn {
    param($SampleRow, [string[]]$Candidates)
    if (-not $SampleRow) { return $null }
    $names = @($SampleRow.PSObject.Properties.Name)
    foreach ($c in $Candidates) {
        $hit = $names | Where-Object { $_ -eq $c } | Select-Object -First 1
        if ($hit) { return [string]$hit }
    }
    foreach ($c in $Candidates) {
        $hit = $names | Where-Object { $_ -like "*$c*" } | Select-Object -First 1
        if ($hit) { return [string]$hit }
    }
    return $null
}

function Get-InvestigateSourceDef {
    param([string]$Name)
    $script:InvestigateSources | Where-Object { $_.Name -eq $Name } | Select-Object -First 1
}

function ConvertTo-InvestigateFilterList {
    param($Filters)
    $list = [System.Collections.Generic.List[object]]::new()
    foreach ($f in @($Filters)) {
        if ($null -eq $f) { continue }
        $col = [string](Get-InvestigateProp -Object $f -Name 'column')
        $op = ([string](Get-InvestigateProp -Object $f -Name 'op')).ToLowerInvariant()
        $val = Get-InvestigateProp -Object $f -Name 'value'
        if ([string]::IsNullOrWhiteSpace($col) -or [string]::IsNullOrWhiteSpace($op)) { continue }
        [void]$list.Add([pscustomobject]@{ column = $col; op = $op; value = $val })
    }
    return @($list)
}

function ConvertTo-InvestigateDateTime {
    param($Value)
    if ($null -eq $Value -or $Value -eq '') { return $null }
    if ($Value -is [datetime]) { return [datetime]$Value }
    $text = [string]$Value
    $dt = [datetime]::MinValue
    if ([datetime]::TryParse($text, [ref]$dt)) { return $dt }
    return $null
}

function Test-InvestigateCompare {
    param($Left, [string]$Op, $Right)
    $l = if ($null -eq $Left) { '' } else { [string]$Left }
    $r = if ($null -eq $Right) { '' } else { [string]$Right }
    switch ($Op) {
        'eq' { return $l -eq $r }
        'ne' { return $l -ne $r }
        'contains' { return $l.Contains($r) }
        'icontains' { return $l.ToLowerInvariant().Contains($r.ToLowerInvariant()) }
        'startswith' { return $l.StartsWith($r, [StringComparison]::OrdinalIgnoreCase) }
        'endswith' { return $l.EndsWith($r, [StringComparison]::OrdinalIgnoreCase) }
        'isnull' { return [string]::IsNullOrWhiteSpace($l) }
        'notnull' { return -not [string]::IsNullOrWhiteSpace($l) }
        'in' {
            $set = @()
            if ($Right -is [System.Collections.IEnumerable] -and $Right -isnot [string]) { $set = @($Right | ForEach-Object { [string]$_ }) }
            else { $set = @($r -split '\s*,\s*') }
            return $set -contains $l
        }
        'not_in' {
            $set = @()
            if ($Right -is [System.Collections.IEnumerable] -and $Right -isnot [string]) { $set = @($Right | ForEach-Object { [string]$_ }) }
            else { $set = @($r -split '\s*,\s*') }
            return $set -notcontains $l
        }
        { $_ -in @('gt', 'gte', 'lt', 'lte') } {
            $ld = ConvertTo-InvestigateDateTime $Left
            $rd = ConvertTo-InvestigateDateTime $Right
            if ($ld -and $rd) {
                switch ($Op) {
                    'gt' { return $ld -gt $rd }
                    'gte' { return $ld -ge $rd }
                    'lt' { return $ld -lt $rd }
                    'lte' { return $ld -le $rd }
                }
            }
            $ln = 0.0; $rn = 0.0
            $okL = [double]::TryParse($l, [ref]$ln)
            $okR = [double]::TryParse($r, [ref]$rn)
            if ($okL -and $okR) {
                switch ($Op) {
                    'gt' { return $ln -gt $rn }
                    'gte' { return $ln -ge $rn }
                    'lt' { return $ln -lt $rn }
                    'lte' { return $ln -le $rn }
                }
            }
            switch ($Op) {
                'gt' { return $l -gt $r }
                'gte' { return $l -ge $r }
                'lt' { return $l -lt $r }
                'lte' { return $l -le $r }
            }
        }
        default { return $true }
    }
}

function Get-InvestigateCell {
    param($Row, [string]$Column, $ColMap)
    if (-not $Row -or -not $Column) { return '' }
    if ($Column -eq '_address') {
        $parts = [System.Collections.Generic.List[string]]::new()
        foreach ($c in @($ColMap.AddressActual)) {
            $v = [string]$Row.$c
            if ($v) { [void]$parts.Add($v) }
        }
        if ($Row.PSObject.Properties['_User']) { [void]$parts.Add([string]$Row._User) }
        if ($Row.PSObject.Properties['_MailboxOwnerUPN']) { [void]$parts.Add([string]$Row._MailboxOwnerUPN) }
        return ($parts -join ' ')
    }
    if ($Column -eq '_ip') {
        $parts = [System.Collections.Generic.List[string]]::new()
        foreach ($c in @($ColMap.IpActual)) {
            $v = [string]$Row.$c
            if ($v) { [void]$parts.Add($v) }
        }
        if ($Row.PSObject.Properties['_ClientIP']) { [void]$parts.Add([string]$Row._ClientIP) }
        return ($parts -join ' ')
    }
    if ($Column -eq '_time') {
        $c = $ColMap.TimeActual
        if ($c) { return [string]$Row.$c }
        return ''
    }
    if ($Row.PSObject.Properties[$Column]) { return [string]$Row.$Column }
    $resolved = $null
    if ($ColMap.Resolved.ContainsKey($Column)) { $resolved = $ColMap.Resolved[$Column] }
    if ($resolved -and $Row.PSObject.Properties[$resolved]) { return [string]$Row.$resolved }
    return ''
}

function New-InvestigateColMap {
    param($SampleRow, $Def)
    $resolved = @{}
    $time = Resolve-InvestigateColumn -SampleRow $SampleRow -Candidates $Def.TimeCols
    $addr = @()
    foreach ($c in $Def.AddressCols) {
        $hit = Resolve-InvestigateColumn -SampleRow $SampleRow -Candidates @($c)
        if ($hit -and $addr -notcontains $hit) { $addr += $hit }
        $resolved[$c] = $hit
    }
    $ips = @()
    foreach ($c in $Def.IpCols) {
        $hit = Resolve-InvestigateColumn -SampleRow $SampleRow -Candidates @($c)
        if ($hit -and $ips -notcontains $hit) { $ips += $hit }
        $resolved[$c] = $hit
    }
    if ($time) { $resolved['_time'] = $time }
    return @{
        TimeActual    = $time
        AddressActual = $addr
        IpActual      = $ips
        Resolved      = $resolved
    }
}

function Test-InvestigateRow {
    param($Row, $Filters, $ColMap)
    foreach ($f in @($Filters)) {
        $cell = Get-InvestigateCell -Row $Row -Column $f.column -ColMap $ColMap
        if (-not (Test-InvestigateCompare -Left $cell -Op $f.op -Right $f.value)) { return $false }
    }
    return $true
}

function Select-InvestigateColumns {
    param($Row, [string[]]$Columns)
    if (-not $Columns -or $Columns.Count -eq 0) {
        $hash = [ordered]@{}
        foreach ($p in $Row.PSObject.Properties) {
            if ($p.Name -eq 'AuditData') {
                $raw = [string]$p.Value
                if ($raw.Length -gt 400) { $hash[$p.Name] = $raw.Substring(0, 400) + '...' }
                else { $hash[$p.Name] = $raw }
            }
            else {
                $hash[$p.Name] = $p.Value
            }
        }
        return [pscustomobject]$hash
    }
    $hash = [ordered]@{}
    foreach ($c in $Columns) {
        if ($Row.PSObject.Properties[$c]) { $hash[$c] = $Row.$c }
        else { $hash[$c] = '' }
    }
    return [pscustomobject]$hash
}

function Resolve-InvestigatePreset {
    param($Def, $SampleRow, [string]$Columns)
    if ([string]::IsNullOrWhiteSpace($Columns)) {
        $picked = @()
        foreach ($c in $Def.Preset) {
            $hit = Resolve-InvestigateColumn -SampleRow $SampleRow -Candidates @($c)
            if ($hit -and $picked -notcontains $hit) { $picked += $hit }
        }
        if ($Def.EnrichUal) {
            foreach ($extra in @('_Operation', '_ClientIP', '_MailboxOwnerUPN')) {
                if ($picked -notcontains $extra) { $picked += $extra }
            }
        }
        return $picked
    }
    if ($Columns -eq 'signin_triage' -or $Columns -eq 'audit_triage' -or $Columns -eq 'message_triage') {
        $picked = @()
        foreach ($c in $Def.Preset) {
            $hit = Resolve-InvestigateColumn -SampleRow $SampleRow -Candidates @($c)
            if ($hit -and $picked -notcontains $hit) { $picked += $hit }
        }
        if ($Def.EnrichUal) {
            foreach ($extra in @('_Operation', '_ClientIP', '_MailboxOwnerUPN')) {
                if ($picked -notcontains $extra) { $picked += $extra }
            }
        }
        return $picked
    }
    return @($Columns -split '\s*,\s*' | Where-Object { $_ })
}

function Get-InvestigateOpenContext {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Source
    )
    if (-not (Test-Path -LiteralPath $Path)) { throw "Report folder not found: $Path" }
    $def = Get-InvestigateSourceDef -Name $Source
    if (-not $def) { throw "Unknown source: $Source" }
    $csv = Get-InvestigateCsvPath -Folder $Path -BaseName $def.Name
    if (-not $csv) { throw "No $($def.Name) CSV in $Path" }
    $rows = Get-InvestigateCachedRows -Path $csv -EnrichUal:$def.EnrichUal
    $sample = $null
    if ($rows.Count -gt 0) { $sample = $rows[0] }
    $map = New-InvestigateColMap -SampleRow $sample -Def $def
    return @{
        Def    = $def
        Csv    = $csv
        Rows   = $rows
        Sample = $sample
        Map    = $map
    }
}

function Get-InvestigateSources {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) { throw "Report folder not found: $Path" }

    $out = [System.Collections.Generic.List[object]]::new()
    foreach ($def in $script:InvestigateSources) {
        $csv = Get-InvestigateCsvPath -Folder $Path -BaseName $def.Name
        if (-not $csv) {
            [void]$out.Add([ordered]@{
                    name      = $def.Name
                    present   = $false
                    path      = $null
                    rowCount  = 0
                    columns   = @()
                    timeCol   = $null
                    preset    = @($def.Preset)
                })
            continue
        }
        $rows = Get-InvestigateCachedRows -Path $csv -EnrichUal:$def.EnrichUal
        $cols = @()
        $time = $null
        if ($rows.Count -gt 0) {
            $cols = @($rows[0].PSObject.Properties.Name)
            $time = Resolve-InvestigateColumn -SampleRow $rows[0] -Candidates $def.TimeCols
        }
        [void]$out.Add([ordered]@{
                name     = $def.Name
                present  = $true
                path     = $csv
                rowCount = $rows.Count
                columns  = $cols
                timeCol  = $time
                preset   = @($def.Preset)
            })
    }

    return [ordered]@{
        path    = $Path
        sources = @($out)
    }
}

function Invoke-InvestigateQuery {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Source,
        $Filters,
        [string]$Columns,
        [string]$SortBy,
        [bool]$Ascending = $true,
        [int]$Limit = 50,
        [int]$Offset = 0
    )

    if ($Limit -lt 1) { $Limit = 50 }
    if ($Limit -gt 200) { $Limit = 200 }
    if ($Offset -lt 0) { $Offset = 0 }

    $ctx = Get-InvestigateOpenContext -Path $Path -Source $Source
    $filterList = ConvertTo-InvestigateFilterList -Filters $Filters
    $matched = [System.Collections.Generic.List[object]]::new()
    foreach ($row in $ctx.Rows) {
        if (Test-InvestigateRow -Row $row -Filters $filterList -ColMap $ctx.Map) {
            [void]$matched.Add($row)
        }
    }

    if ($SortBy) {
        $matched = [System.Collections.Generic.List[object]]@(
            $matched | Sort-Object {
                $v = Get-InvestigateCell -Row $_ -Column $SortBy -ColMap $ctx.Map
                $dt = ConvertTo-InvestigateDateTime $v
                if ($dt) { $dt.ToString('o') } else { $v }
            } -Descending:(-not $Ascending)
        )
    }

    $total = $matched.Count
    $page = @($matched | Select-Object -Skip $Offset -First $Limit)
    $colList = Resolve-InvestigatePreset -Def $ctx.Def -SampleRow $ctx.Sample -Columns $Columns
    $rowsOut = @(foreach ($r in $page) { Select-InvestigateColumns -Row $r -Columns $colList })

    return [ordered]@{
        path       = $Path
        source     = $Source
        csv        = $ctx.Csv
        total      = $total
        offset     = $Offset
        limit      = $Limit
        columns    = $colList
        timeCol    = $ctx.Map.TimeActual
        rows       = $rowsOut
    }
}

function Get-InvestigateEmailDomain {
    param([string]$Address)
    if ([string]::IsNullOrWhiteSpace($Address)) { return '' }
    $at = $Address.LastIndexOf('@')
    if ($at -lt 0) { return '' }
    return $Address.Substring($at + 1).ToLowerInvariant()
}

function Split-InvestigateRecipients {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return @() }
    return @(
        $Text -split '[;,\s]+' |
            Where-Object { $_ -match '@' } |
            ForEach-Object { $_.Trim().Trim('"') } |
            Where-Object { $_ }
    )
}

function Get-InvestigateMailGraph {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        $Filters,
        [int]$MinCount = 1,
        [int]$MaxNodes = 80,
        [switch]$HideInternal,
        [ValidateSet('hour', 'day')]
        [string]$Bucket = 'hour'
    )

    if ($MinCount -lt 1) { $MinCount = 1 }
    if ($MaxNodes -lt 5) { $MaxNodes = 5 }
    if ($MaxNodes -gt 200) { $MaxNodes = 200 }

    $ctx = Get-InvestigateOpenContext -Path $Path -Source 'MessageTrace'
    $filterList = ConvertTo-InvestigateFilterList -Filters $Filters
    $senderCol = Resolve-InvestigateColumn -SampleRow $ctx.Sample -Candidates @('SenderAddress', 'Sender')
    $recipCol = Resolve-InvestigateColumn -SampleRow $ctx.Sample -Candidates @('RecipientAddress', 'Recipient')
    $timeCol = $ctx.Map.TimeActual
    if (-not $senderCol -or -not $recipCol) {
        return [ordered]@{
            path         = $Path
            matched      = 0
            minCount     = $MinCount
            maxNodes     = $MaxNodes
            hideInternal = [bool]$HideInternal
            bucket       = $Bucket
            nodes        = @()
            edges        = @()
            timeline     = @()
        }
    }

    $edgeCounts = @{}
    $nodeCounts = @{}
    $buckets = @{}
    $matched = 0

    foreach ($row in $ctx.Rows) {
        if (-not (Test-InvestigateRow -Row $row -Filters $filterList -ColMap $ctx.Map)) { continue }
        $matched++
        $sender = ([string]$row.$senderCol).Trim().ToLowerInvariant()
        if (-not $sender) { continue }
        $recips = Split-InvestigateRecipients -Text ([string]$row.$recipCol)
        if ($recips.Count -eq 0) { continue }

        if ($timeCol) {
            $dt = ConvertTo-InvestigateDateTime $row.$timeCol
            if ($dt) {
                $key = if ($Bucket -eq 'day') { $dt.ToString('yyyy-MM-dd') } else { $dt.ToString('yyyy-MM-ddTHH:00') }
                if (-not $buckets.ContainsKey($key)) { $buckets[$key] = 0 }
                $buckets[$key]++
            }
        }

        if (-not $nodeCounts.ContainsKey($sender)) { $nodeCounts[$sender] = 0 }
        $nodeCounts[$sender]++

        foreach ($to in $recips) {
            $toL = $to.ToLowerInvariant()
            if ($HideInternal) {
                $d1 = Get-InvestigateEmailDomain $sender
                $d2 = Get-InvestigateEmailDomain $toL
                if ($d1 -and $d1 -eq $d2) { continue }
            }
            if (-not $nodeCounts.ContainsKey($toL)) { $nodeCounts[$toL] = 0 }
            $nodeCounts[$toL]++
            $ek = "$sender|$toL"
            if (-not $edgeCounts.ContainsKey($ek)) { $edgeCounts[$ek] = 0 }
            $edgeCounts[$ek]++
        }
    }

    $edges = @(
        $edgeCounts.GetEnumerator() |
            Where-Object { $_.Value -ge $MinCount } |
            ForEach-Object {
                $parts = $_.Key -split '\|', 2
                [ordered]@{ from = $parts[0]; to = $parts[1]; count = [int]$_.Value }
            }
    )

    $keep = @{}
    foreach ($e in $edges) {
        if (-not $keep.ContainsKey($e.from)) { $keep[$e.from] = 0 }
        if (-not $keep.ContainsKey($e.to)) { $keep[$e.to] = 0 }
        $keep[$e.from] += $e.count
        $keep[$e.to] += $e.count
    }
    $topNames = @(
        $keep.GetEnumerator() |
            Sort-Object { -$_.Value }, Name |
            Select-Object -First $MaxNodes |
            ForEach-Object { $_.Key }
    )
    $topSet = @{}
    foreach ($n in $topNames) { $topSet[$n] = $true }

    $edgesOut = @($edges | Where-Object { $topSet.ContainsKey($_.from) -and $topSet.ContainsKey($_.to) })
    $nodesOut = @(
        foreach ($n in $topNames) {
            [ordered]@{ id = $n; label = $n; count = [int]$nodeCounts[$n] }
        }
    )
    $sortedKeys = @($buckets.Keys | Sort-Object)
    if ($sortedKeys.Count -ge 2) {
        $first = ConvertTo-InvestigateDateTime $sortedKeys[0]
        $last = ConvertTo-InvestigateDateTime $sortedKeys[-1]
        if ($first -and $last) {
            $step = if ($Bucket -eq 'day') { [timespan]::FromDays(1) } else { [timespan]::FromHours(1) }
            $spanSteps = [int][math]::Floor(($last - $first).TotalHours)
            if ($Bucket -eq 'day') { $spanSteps = [int][math]::Floor(($last - $first).TotalDays) }
            $maxSteps = if ($Bucket -eq 'day') { 90 } else { 72 }
            if ($spanSteps -ge 1 -and $spanSteps -le $maxSteps) {
                $cur = $first
                while ($cur -le $last) {
                    $k = if ($Bucket -eq 'day') { $cur.ToString('yyyy-MM-dd') } else { $cur.ToString('yyyy-MM-ddTHH:00') }
                    if (-not $buckets.ContainsKey($k)) { $buckets[$k] = 0 }
                    $cur = $cur.Add($step)
                }
            }
        }
    }
    $timeline = @(
        $buckets.GetEnumerator() |
            Sort-Object Name |
            ForEach-Object { [ordered]@{ t = $_.Key; count = [int]$_.Value } }
    )

    return [ordered]@{
        path        = $Path
        matched     = $matched
        minCount    = $MinCount
        maxNodes    = $MaxNodes
        hideInternal = [bool]$HideInternal
        bucket      = $Bucket
        nodes       = $nodesOut
        edges       = $edgesOut
        timeline    = $timeline
    }
}

function Export-InvestigateQuery {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Source,
        $Filters,
        [string]$Columns,
        [string]$SortBy,
        [bool]$Ascending = $true
    )

    $ctx = Get-InvestigateOpenContext -Path $Path -Source $Source
    $filterList = ConvertTo-InvestigateFilterList -Filters $Filters
    $matched = [System.Collections.Generic.List[object]]::new()
    foreach ($row in $ctx.Rows) {
        if (Test-InvestigateRow -Row $row -Filters $filterList -ColMap $ctx.Map) {
            [void]$matched.Add($row)
        }
    }
    if ($SortBy) {
        $matched = [System.Collections.Generic.List[object]]@(
            $matched | Sort-Object {
                $v = Get-InvestigateCell -Row $_ -Column $SortBy -ColMap $ctx.Map
                $dt = ConvertTo-InvestigateDateTime $v
                if ($dt) { $dt.ToString('o') } else { $v }
            } -Descending:(-not $Ascending)
        )
    }

    $colList = Resolve-InvestigatePreset -Def $ctx.Def -SampleRow $ctx.Sample -Columns $Columns
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $outFile = Join-Path $Path ("Investigate_{0}_{1}.csv" -f $Source, $stamp)
    $exportRows = @(foreach ($r in $matched) { Select-InvestigateColumns -Row $r -Columns $colList })
    $exportRows | Export-Csv -LiteralPath $outFile -NoTypeInformation -Encoding UTF8

    return [ordered]@{
        path     = $Path
        source   = $Source
        file     = $outFile
        rowCount = $exportRows.Count
    }
}

Export-ModuleMember -Function Get-InvestigateSources, Invoke-InvestigateQuery, Get-InvestigateMailGraph, Export-InvestigateQuery, Get-InvestigateCsvPath
