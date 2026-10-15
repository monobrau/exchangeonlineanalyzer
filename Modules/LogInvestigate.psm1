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
        IpCols      = @('IPAddress', 'IpAddress', 'IP address', 'IP')
        GeoCols     = @('CountryOrRegion', 'Country', 'Location')
        Preset      = @('CreatedDateTime', 'UserPrincipalName', 'IPAddress', 'AppDisplayName', 'Status', 'ConditionalAccessStatus', 'CAPolicyDetails', 'CountryOrRegion', 'ClientAppUsed')
        EnrichSignInCa = $true
    }
    @{
        Name        = 'UnifiedAuditLogs'
        TimeCols    = @('CreationDate', 'CreationTime', 'CreationDateTime')
        AddressCols = @('UserIds', 'UserId', 'UserPrincipalName', 'UserKey')
        IpCols      = @('ClientIP', 'ClientIp', 'ClientIPAddress', 'IPAddress')
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
        [switch]$EnrichUal,
        [switch]$EnrichSignInCa
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
    if ($EnrichSignInCa) {
        foreach ($row in $rows) { Add-InvestigateSignInCaFields -Row $row }
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
    if (-not $ip) { $ip = [string]$Row.ClientIPAddress }
    if (-not $ip -and $audit) { $ip = [string]$audit.ClientIP }
    if (-not $ip -and $audit) { $ip = [string]$audit.ClientIPAddress }
    if (-not $ip -and $audit) { $ip = [string]$audit.ActorIpAddress }
    $mbx = ''
    if ($audit) { $mbx = [string]$audit.MailboxOwnerUPN }
    $user = [string]$Row.UserIds
    if (-not $user -and $audit) { $user = [string]$audit.UserId }

    Add-Member -InputObject $Row -NotePropertyName '_Operation' -NotePropertyValue $op -Force
    Add-Member -InputObject $Row -NotePropertyName '_ClientIP' -NotePropertyValue $ip -Force
    Add-Member -InputObject $Row -NotePropertyName '_MailboxOwnerUPN' -NotePropertyValue $mbx -Force
    Add-Member -InputObject $Row -NotePropertyName '_User' -NotePropertyValue $user -Force
}

function ConvertTo-InvestigateCaResultLabel {
    param([string]$Result)
    $r = ([string]$Result).Trim()
    if ([string]::IsNullOrWhiteSpace($r)) { return 'Unknown' }
    switch -Regex ($r) {
        '^(?i)success$' { return 'Success' }
        '^(?i)fail(ure|ed)$' { return 'Failed' }
        '^(?i)notApplied$' { return 'Not Applied' }
        '^(?i)not[\s-]?applied$' { return 'Not Applied' }
        '^(?i)reportOnly' { return 'Report-only' }
        '^(?i)notEnabled$' { return 'Not Enabled' }
        default {
            if ($r -match '(?i)report') { return 'Report-only' }
            return $r
        }
    }
}

function ConvertTo-InvestigateCaTone {
    param([string]$Label)
    switch -Regex ([string]$Label) {
        '^(?i)Failed$' { return 'fail' }
        '^(?i)Success$' { return 'ok' }
        '^(?i)Report-only$' { return 'report' }
        default { return 'muted' }
    }
}

function Parse-InvestigateConditionalAccess {
    param(
        [string]$Json,
        [string]$Status,
        [string]$Details,
        [string]$Names,
        [string]$Results
    )
    $policies = [System.Collections.Generic.List[object]]::new()
    $raw = [string]$Json
    if ($raw -and $raw.TrimStart() -match '^[\[\{]') {
        try {
            $parsed = $raw | ConvertFrom-Json -ErrorAction Stop
            foreach ($p in @($parsed)) {
                $id = [string](Get-InvestigateProp -Object $p -Name 'id')
                if (-not $id) { $id = [string](Get-InvestigateProp -Object $p -Name 'Id') }
                $name = [string](Get-InvestigateProp -Object $p -Name 'displayName')
                if (-not $name) { $name = [string](Get-InvestigateProp -Object $p -Name 'DisplayName') }
                $result = [string](Get-InvestigateProp -Object $p -Name 'result')
                if (-not $result) { $result = [string](Get-InvestigateProp -Object $p -Name 'Result') }
                $label = ConvertTo-InvestigateCaResultLabel -Result $result
                [void]$policies.Add([pscustomobject]@{
                        id     = $id
                        name   = $(if ($name) { $name } else { $id })
                        result = $label
                        tone   = ConvertTo-InvestigateCaTone -Label $label
                    })
            }
        }
        catch { }
    }
    if ($policies.Count -eq 0 -and ($Names -or $Results -or $Details)) {
        $nameList = @([string]$Names -split '\s*;\s*' | Where-Object { $_ })
        $resultList = @([string]$Results -split '\s*;\s*' | Where-Object { $_ })
        $detailList = @([string]$Details -split '\s*;\s*' | Where-Object { $_ })
        $n = [Math]::Max($nameList.Count, [Math]::Max($resultList.Count, $detailList.Count))
        for ($i = 0; $i -lt $n; $i++) {
            $name = if ($i -lt $nameList.Count -and $nameList[$i]) { $nameList[$i] }
            elseif ($i -lt $detailList.Count -and $detailList[$i]) { $detailList[$i] }
            else { "Policy $($i + 1)" }
            $res = if ($i -lt $resultList.Count) { $resultList[$i] } else { '' }
            $label = ConvertTo-InvestigateCaResultLabel -Result $res
            [void]$policies.Add([pscustomobject]@{
                    id     = ''
                    name   = $name
                    result = $label
                    tone   = ConvertTo-InvestigateCaTone -Label $label
                })
        }
    }
    $statusLabel = ConvertTo-InvestigateCaResultLabel -Result $Status
    if ($statusLabel -eq 'Unknown' -and $Status) { $statusLabel = $Status }
    $summary = if ($policies.Count -gt 0) {
        ($policies | ForEach-Object { "$($_.name):$($_.result)" }) -join '; '
    }
    else {
        $(if ($Status) { $Status } else { $statusLabel })
    }
    return [pscustomobject]@{
        status      = $(if ($Status) { $Status } else { $statusLabel })
        statusLabel = $statusLabel
        summary     = $summary
        policies    = @($policies)
    }
}

function Add-InvestigateSignInCaFields {
    param($Row)
    if (-not $Row) { return }
    $json = ''
    $status = ''
    $details = ''
    $names = ''
    $results = ''
    if ($Row.PSObject.Properties['ConditionalAccessPoliciesJSON']) { $json = [string]$Row.ConditionalAccessPoliciesJSON }
    if ($Row.PSObject.Properties['ConditionalAccessStatus']) { $status = [string]$Row.ConditionalAccessStatus }
    if ($Row.PSObject.Properties['CAPolicyDetails']) { $details = [string]$Row.CAPolicyDetails }
    if ($Row.PSObject.Properties['CAPolicyNames']) { $names = [string]$Row.CAPolicyNames }
    if ($Row.PSObject.Properties['CAPolicyResults']) { $results = [string]$Row.CAPolicyResults }
    $parsed = Parse-InvestigateConditionalAccess -Json $json -Status $status -Details $details -Names $names -Results $results
    $policiesJson = '[]'
    if ($parsed.policies.Count -gt 0) {
        $policiesJson = @($parsed.policies) | ConvertTo-Json -Compress -Depth 4
        if ($parsed.policies.Count -eq 1 -and $policiesJson -notmatch '^\s*\[') { $policiesJson = "[$policiesJson]" }
    }
    Add-Member -InputObject $Row -NotePropertyName '_CA' -NotePropertyValue $parsed.summary -Force
    Add-Member -InputObject $Row -NotePropertyName '_CAPolicies' -NotePropertyValue $policiesJson -Force
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
    if ($Column -eq '_geo') {
        $parts = [System.Collections.Generic.List[string]]::new()
        foreach ($c in @($ColMap.GeoActual)) {
            $v = [string]$Row.$c
            if ($v) { [void]$parts.Add($v) }
        }
        return ($parts -join ' ')
    }
    if ($Column -eq '_time') {
        $c = $ColMap.TimeActual
        if ($c) { return [string]$Row.$c }
        return ''
    }
    if ($Column -eq '_CA' -or $Column -eq '_ca') {
        if ($Row.PSObject.Properties['_CA']) { return [string]$Row._CA }
        $bits = [System.Collections.Generic.List[string]]::new()
        foreach ($c in @('ConditionalAccessStatus', 'CAPolicyDetails', 'CAPolicyNames', 'CAPolicyResults')) {
            if ($Row.PSObject.Properties[$c] -and $Row.$c) { [void]$bits.Add([string]$Row.$c) }
        }
        return ($bits -join ' ')
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
    $geo = @()
    foreach ($c in @($Def.GeoCols | Where-Object { $_ })) {
        $hit = Resolve-InvestigateColumn -SampleRow $SampleRow -Candidates @($c)
        if ($hit -and $geo -notcontains $hit) { $geo += $hit }
        $resolved[$c] = $hit
    }
    if ($time) { $resolved['_time'] = $time }
    return @{
        TimeActual    = $time
        AddressActual = $addr
        IpActual      = $ips
        GeoActual     = $geo
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
        if ($Def.EnrichSignInCa) {
            foreach ($extra in @('_CA', '_CAPolicies')) {
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
        if ($Def.EnrichSignInCa) {
            foreach ($extra in @('_CA', '_CAPolicies')) {
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
    $rows = Get-InvestigateCachedRows -Path $csv -EnrichUal:$def.EnrichUal -EnrichSignInCa:$def.EnrichSignInCa
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
        $rows = Get-InvestigateCachedRows -Path $csv -EnrichUal:$def.EnrichUal -EnrichSignInCa:$def.EnrichSignInCa
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

function Get-InvestigateNormalizedIp {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return '' }
    $ip = $Value.Trim().Trim('"')
    if ($ip -match '(?i)^(?:clientip|ipaddress|ip)\s*=\s*(.+)$') { $ip = $Matches[1].Trim() }
    if ($ip -match '^\[([^\]]+)\](?::\d+)?$') { $ip = $Matches[1] }
    elseif ($ip -match '^(\d{1,3}(?:\.\d{1,3}){3}):\d+$') { $ip = $Matches[1] }
    elseif ($ip -match '(?i)^::ffff:(\d{1,3}(?:\.\d{1,3}){3})$') { $ip = $Matches[1] }
    return $ip
}

function Get-InvestigateCountryCode {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return '' }
    $t = $Value.Trim().Trim('"')
    if ($t -match '^\s*\{') {
        try {
            $j = $t | ConvertFrom-Json -ErrorAction Stop
            foreach ($n in @('countryOrRegion', 'CountryOrRegion', 'country', 'Country')) {
                if ($j.PSObject.Properties[$n] -and $j.$n) { return ([string]$j.$n).Trim() }
            }
        } catch { }
    }
    if ($t -match '(?i)countryorregion["\s:=]+([A-Za-z]{2,3})\b') { return $Matches[1].ToUpperInvariant() }
    $parts = @($t -split '\s*,\s*' | Where-Object { $_ })
    if ($parts.Count -ge 2 -and $parts[-1] -match '^[A-Za-z]{2,3}$') { return $parts[-1].ToUpperInvariant() }
    return $t
}

function Test-InvestigatePublicIpAddress {
    param([string]$Ip)
    $ip = Get-InvestigateNormalizedIp -Value $Ip
    if (-not $ip) { return $false }
    $parsed = $null
    if (-not [System.Net.IPAddress]::TryParse($ip, [ref]$parsed)) { return $false }
    if ($parsed.AddressFamily -ne [System.Net.Sockets.AddressFamily]::InterNetwork) {
        $s = $parsed.ToString()
        if ($s.StartsWith('fe80', [System.StringComparison]::OrdinalIgnoreCase)) { return $false }
        if ($s.StartsWith('fc', [System.StringComparison]::OrdinalIgnoreCase) -or $s.StartsWith('fd', [System.StringComparison]::OrdinalIgnoreCase)) { return $false }
        return $true
    }
    $b = $parsed.GetAddressBytes()
    if ($b[0] -eq 10 -or $b[0] -eq 127 -or $b[0] -eq 0) { return $false }
    if ($b[0] -eq 169 -and $b[1] -eq 254) { return $false }
    if ($b[0] -eq 172 -and $b[1] -ge 16 -and $b[1] -le 31) { return $false }
    if ($b[0] -eq 192 -and $b[1] -eq 168) { return $false }
    if ($b[0] -eq 100 -and $b[1] -ge 64 -and $b[1] -le 127) { return $false }
    return $true
}

function Test-InvestigateFailedStatus {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $false }
    return ($Value -match '(?i)fail|interrupt|denied|block')
}

function Test-InvestigateAccessUser {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $false }
    if ($Value -match '(?i)^(NT AUTHORITY\\|S-1-|SHAREPOINT\\)') { return $false }
    return $true
}

function Get-InvestigateAccessRowBits {
    param($Row, $ColMap, $Kind)
    $user = ''
    foreach ($c in @($ColMap.AddressActual)) {
        $v = [string]$Row.$c
        if (Test-InvestigateAccessUser -Value $v) { $user = $v.Trim().ToLowerInvariant(); break }
    }
    if (-not $user -and $Row.PSObject.Properties['_User']) {
        $v = [string]$Row._User
        if (Test-InvestigateAccessUser -Value $v) { $user = $v.Trim().ToLowerInvariant() }
    }
    $ip = ''
    foreach ($c in @($ColMap.IpActual)) {
        $v = Get-InvestigateNormalizedIp -Value ([string]$Row.$c)
        if ($v) { $ip = $v; break }
    }
    if (-not $ip -and $Row.PSObject.Properties['_ClientIP']) {
        $ip = Get-InvestigateNormalizedIp -Value ([string]$Row._ClientIP)
    }
    $country = ''
    foreach ($c in @($ColMap.GeoActual)) {
        $v = Get-InvestigateCountryCode -Value ([string]$Row.$c)
        if ($v -and $v -ne 'Unknown') { $country = $v; break }
        if ($v) { $country = $v }
    }
    $failed = $false
    if ($Kind -eq 'signin') {
        $failed = Test-InvestigateFailedStatus -Value (Get-InvestigateCell -Row $Row -Column 'Status' -ColMap $ColMap)
        if (-not $failed) { $failed = Test-InvestigateFailedStatus -Value (Get-InvestigateCell -Row $Row -Column 'ResultStatus' -ColMap $ColMap) }
    } else {
        $failed = Test-InvestigateFailedStatus -Value (Get-InvestigateCell -Row $Row -Column 'ResultStatus' -ColMap $ColMap)
        if (-not $failed) { $failed = Test-InvestigateFailedStatus -Value (Get-InvestigateCell -Row $Row -Column '_Operation' -ColMap $ColMap) }
    }
    $dt = $null
    if ($ColMap.TimeActual) { $dt = ConvertTo-InvestigateDateTime $Row.($ColMap.TimeActual) }
    return @{
        User    = $user
        Ip      = $ip
        Country = $country
        Failed  = $failed
        Time    = $dt
    }
}

function Add-InvestigateAccessEvent {
    param($Agg, $Bits, [string]$Source)
    if (-not $Bits.User -or -not $Bits.Ip) { return }
    if ($Agg.hidePrivateIp -and -not (Test-InvestigatePublicIpAddress -Ip $Bits.Ip)) { return }
    $country = $Bits.Country
    if (-not $country -and $Agg.ipCountry.ContainsKey($Bits.Ip)) { $country = $Agg.ipCountry[$Bits.Ip] }
    if ($Bits.Country) { $Agg.ipCountry[$Bits.Ip] = $Bits.Country }
    if (-not $country) { $country = 'Unknown' }
    foreach ($gf in @($Agg.geoFilters)) {
        if (-not (Test-InvestigateCompare -Left $country -Op $gf.op -Right $gf.value)) { return }
    }
    $Agg.matched++
    if ($Source -eq 'signin') { $Agg.signInMatched++ } else { $Agg.ualMatched++ }

    if ($Bits.Time) {
        $key = if ($Agg.bucket -eq 'day') { $Bits.Time.ToString('yyyy-MM-dd') } else { $Bits.Time.ToString('yyyy-MM-ddTHH:00') }
        if (-not $Agg.buckets.ContainsKey($key)) { $Agg.buckets[$key] = 0; $Agg.bucketFails[$key] = 0 }
        $Agg.buckets[$key]++
        if ($Bits.Failed) { $Agg.bucketFails[$key]++ }
    }

    if (-not $Agg.userCounts.ContainsKey($Bits.User)) { $Agg.userCounts[$Bits.User] = 0; $Agg.userFails[$Bits.User] = 0 }
    $Agg.userCounts[$Bits.User]++
    if ($Bits.Failed) { $Agg.userFails[$Bits.User]++ }

    if (-not $Agg.ipCounts.ContainsKey($Bits.Ip)) { $Agg.ipCounts[$Bits.Ip] = 0; $Agg.ipFails[$Bits.Ip] = 0; $Agg.ipSources[$Bits.Ip] = @{} }
    $Agg.ipCounts[$Bits.Ip]++
    if ($Bits.Failed) { $Agg.ipFails[$Bits.Ip]++ }
    $Agg.ipSources[$Bits.Ip][$Source] = $true

    $ek = "$($Bits.User)|$($Bits.Ip)"
    if (-not $Agg.edgeCounts.ContainsKey($ek)) { $Agg.edgeCounts[$ek] = 0; $Agg.edgeFails[$ek] = 0 }
    $Agg.edgeCounts[$ek]++
    if ($Bits.Failed) { $Agg.edgeFails[$ek]++ }

    if (-not $Agg.countryCounts.ContainsKey($country)) {
        $Agg.countryCounts[$country] = 0
        $Agg.countryFails[$country] = 0
        $Agg.countryUsers[$country] = @{}
        $Agg.countryIps[$country] = @{}
    }
    $Agg.countryCounts[$country]++
    if ($Bits.Failed) { $Agg.countryFails[$country]++ }
    $Agg.countryUsers[$country][$Bits.User] = $true
    $Agg.countryIps[$country][$Bits.Ip] = $true
}

function Get-InvestigateAccessGraph {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        $Filters,
        [int]$MinCount = 1,
        [int]$MaxNodes = 80,
        [bool]$IncludeSignIn = $true,
        [bool]$IncludeUal = $true,
        [bool]$HidePrivateIp = $true,
        [ValidateSet('hour', 'day')]
        [string]$Bucket = 'hour'
    )

    if ($MinCount -lt 1) { $MinCount = 1 }
    if ($MaxNodes -lt 5) { $MaxNodes = 5 }
    if ($MaxNodes -gt 200) { $MaxNodes = 200 }

    $filterList = ConvertTo-InvestigateFilterList -Filters $Filters
    $geoFilters = @($filterList | Where-Object { $_.column -eq '_geo' })
    $baseFilters = @($filterList | Where-Object { $_.column -ne '_geo' })

    $ipCountry = @{}
    $edgeCounts = @{}
    $edgeFails = @{}
    $userCounts = @{}
    $userFails = @{}
    $ipCounts = @{}
    $ipFails = @{}
    $ipSources = @{}
    $countryCounts = @{}
    $countryFails = @{}
    $countryUsers = @{}
    $countryIps = @{}
    $buckets = @{}
    $bucketFails = @{}
    $agg = @{
        ipCountry     = $ipCountry
        edgeCounts    = $edgeCounts
        edgeFails     = $edgeFails
        userCounts    = $userCounts
        userFails     = $userFails
        ipCounts      = $ipCounts
        ipFails       = $ipFails
        ipSources     = $ipSources
        countryCounts = $countryCounts
        countryFails  = $countryFails
        countryUsers  = $countryUsers
        countryIps    = $countryIps
        buckets       = $buckets
        bucketFails   = $bucketFails
        geoFilters    = $geoFilters
        hidePrivateIp = [bool]$HidePrivateIp
        bucket        = $Bucket
        matched       = 0
        signInMatched = 0
        ualMatched    = 0
    }

    if ($IncludeSignIn) {
        try {
            $ctx = Get-InvestigateOpenContext -Path $Path -Source 'SignInLogs'
            foreach ($row in $ctx.Rows) {
                if (-not (Test-InvestigateRow -Row $row -Filters $baseFilters -ColMap $ctx.Map)) { continue }
                $bits = Get-InvestigateAccessRowBits -Row $row -ColMap $ctx.Map -Kind 'signin'
                Add-InvestigateAccessEvent -Agg $agg -Bits $bits -Source 'signin'
            }
        } catch {
            if ($_.Exception.Message -notmatch 'No SignInLogs CSV') { throw }
        }
    }
    if ($IncludeUal) {
        try {
            $ctx = Get-InvestigateOpenContext -Path $Path -Source 'UnifiedAuditLogs'
            foreach ($row in $ctx.Rows) {
                if (-not (Test-InvestigateRow -Row $row -Filters $baseFilters -ColMap $ctx.Map)) { continue }
                $bits = Get-InvestigateAccessRowBits -Row $row -ColMap $ctx.Map -Kind 'ual'
                Add-InvestigateAccessEvent -Agg $agg -Bits $bits -Source 'ual'
            }
        } catch {
            if ($_.Exception.Message -notmatch 'No UnifiedAuditLogs CSV') { throw }
        }
    }

    $matched = [int]$agg.matched
    $signInMatched = [int]$agg.signInMatched
    $ualMatched = [int]$agg.ualMatched

    $edges = @(
        $edgeCounts.GetEnumerator() |
            Where-Object { $_.Value -ge $MinCount } |
            ForEach-Object {
                $parts = $_.Key -split '\|', 2
                $ip = $parts[1]
                $country = if ($ipCountry.ContainsKey($ip)) { $ipCountry[$ip] } else { 'Unknown' }
                [ordered]@{
                    from      = $parts[0]
                    to        = $ip
                    count     = [int]$_.Value
                    failCount = [int]$edgeFails[$_.Key]
                    country   = $country
                }
            }
    )

    $keepScore = @{}
    foreach ($e in $edges) {
        if (-not $keepScore.ContainsKey("u:$($e.from)")) { $keepScore["u:$($e.from)"] = 0 }
        if (-not $keepScore.ContainsKey("i:$($e.to)")) { $keepScore["i:$($e.to)"] = 0 }
        $keepScore["u:$($e.from)"] += $e.count
        $keepScore["i:$($e.to)"] += $e.count
    }
    $topIds = @(
        $keepScore.GetEnumerator() |
            Sort-Object { -$_.Value }, Name |
            Select-Object -First $MaxNodes |
            ForEach-Object { $_.Key }
    )
    $topSet = @{}
    foreach ($id in $topIds) { $topSet[$id] = $true }

    $edgesOut = @($edges | Where-Object { $topSet.ContainsKey("u:$($_.from)") -and $topSet.ContainsKey("i:$($_.to)") })
    $nodesOut = @(
        foreach ($id in $topIds) {
            if ($id.StartsWith('u:')) {
                $name = $id.Substring(2)
                [ordered]@{
                    id        = $id
                    label     = $name
                    kind      = 'user'
                    count     = [int]$userCounts[$name]
                    failCount = [int]$userFails[$name]
                    country   = ''
                }
            } else {
                $ip = $id.Substring(2)
                $country = if ($ipCountry.ContainsKey($ip)) { $ipCountry[$ip] } else { 'Unknown' }
                $src = @()
                if ($ipSources.ContainsKey($ip)) { $src = @($ipSources[$ip].Keys) }
                [ordered]@{
                    id        = $id
                    label     = $ip
                    kind      = 'ip'
                    count     = [int]$ipCounts[$ip]
                    failCount = [int]$ipFails[$ip]
                    country   = $country
                    sources   = $src
                }
            }
        }
    )

    $sortedKeys = @($buckets.Keys | Sort-Object)
    if ($sortedKeys.Count -ge 2) {
        $first = ConvertTo-InvestigateDateTime $sortedKeys[0]
        $last = ConvertTo-InvestigateDateTime $sortedKeys[-1]
        if ($first -and $last) {
            $step = if ($Bucket -eq 'day') { [timespan]::FromDays(1) } else { [timespan]::FromHours(1) }
            $spanSteps = if ($Bucket -eq 'day') { [int][math]::Floor(($last - $first).TotalDays) } else { [int][math]::Floor(($last - $first).TotalHours) }
            $maxSteps = if ($Bucket -eq 'day') { 90 } else { 72 }
            if ($spanSteps -ge 1 -and $spanSteps -le $maxSteps) {
                $cur = $first
                while ($cur -le $last) {
                    $k = if ($Bucket -eq 'day') { $cur.ToString('yyyy-MM-dd') } else { $cur.ToString('yyyy-MM-ddTHH:00') }
                    if (-not $buckets.ContainsKey($k)) { $buckets[$k] = 0; $bucketFails[$k] = 0 }
                    $cur = $cur.Add($step)
                }
            }
        }
    }
    $timeline = @(
        $buckets.GetEnumerator() |
            Sort-Object Name |
            ForEach-Object { [ordered]@{ t = $_.Key; count = [int]$_.Value; failCount = [int]$bucketFails[$_.Key] } }
    )
    $countries = @(
        $countryCounts.GetEnumerator() |
            Sort-Object { -$_.Value }, Name |
            ForEach-Object {
                $c = $_.Key
                [ordered]@{
                    country   = $c
                    count     = [int]$_.Value
                    failCount = [int]$countryFails[$c]
                    users     = @($countryUsers[$c].Keys).Count
                    ips       = @($countryIps[$c].Keys).Count
                }
            }
    )

    return [ordered]@{
        path          = $Path
        matched       = $matched
        signInMatched = $signInMatched
        ualMatched    = $ualMatched
        minCount      = $MinCount
        maxNodes      = $MaxNodes
        includeSignIn = [bool]$IncludeSignIn
        includeUal    = [bool]$IncludeUal
        hidePrivateIp = [bool]$HidePrivateIp
        bucket        = $Bucket
        nodes         = $nodesOut
        edges         = $edgesOut
        countries     = $countries
        timeline      = $timeline
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

function Get-InvestigateNormalizedPath {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return $null }
    $t = $Path.Trim().TrimEnd('\', '/')
    try { return [System.IO.Path]::GetFullPath($t) } catch { return $t }
}

function Get-InvestigatePackRoots {
    @(
        Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'ExchangeOnlineAnalyzer\SecurityInvestigation'
    )
}

function Get-InvestigatePackSourceNames {
    param([Parameter(Mandatory)][string]$Folder)
    $names = @(Get-ChildItem -LiteralPath $Folder -Filter '*.csv' -File -ErrorAction SilentlyContinue | ForEach-Object { $_.Name })
    $found = [System.Collections.Generic.List[string]]::new()
    foreach ($def in $script:InvestigateSources) {
        $hit = $names | Where-Object { $_ -like "$($def.Name)*.csv" -and $_ -notmatch '_curated' } | Select-Object -First 1
        if ($hit) { [void]$found.Add($def.Name) }
    }
    return @($found)
}

function Get-InvestigatePackInfo {
    param(
        [Parameter(Mandatory)][string]$Path,
        [switch]$AllowMissing
    )
    $full = Get-InvestigateNormalizedPath -Path $Path
    if (-not $full) { return $null }
    $leaf = Split-Path $full -Leaf
    $parent = Split-Path $full -Parent
    $company = if ($parent) { Split-Path $parent -Leaf } else { '' }
    $exists = Test-Path -LiteralPath $full -PathType Container
    $curated = ($leaf -like 'Curated_*')
    if (-not $exists) {
        if (-not $AllowMissing) { return $null }
        return [ordered]@{
            path      = $full
            exists    = $false
            name      = $leaf
            company   = $company
            sources   = @()
            curated   = $curated
            writtenAt = $null
        }
    }
    $sources = @(Get-InvestigatePackSourceNames -Folder $full)
    if ($sources.Count -lt 1) {
        if (-not $AllowMissing) { return $null }
        return [ordered]@{
            path      = $full
            exists    = $true
            name      = $leaf
            company   = $company
            sources   = @()
            curated   = $curated
            writtenAt = $null
        }
    }
    $item = Get-Item -LiteralPath $full
    return [ordered]@{
        path      = $full
        exists    = $true
        name      = $leaf
        company   = $company
        sources   = $sources
        curated   = $curated
        writtenAt = $item.LastWriteTimeUtc.ToString('o')
    }
}

function Add-InvestigatePackCandidate {
    param(
        $Map,
        [string]$Path,
        $Hint,
        [switch]$AllowMissing
    )
    $full = Get-InvestigateNormalizedPath -Path $Path
    if (-not $full) { return }
    $key = $full.ToLowerInvariant()
    $info = Get-InvestigatePackInfo -Path $full -AllowMissing:$AllowMissing
    if (-not $info) { return }
    if (-not $Map.ContainsKey($key)) {
        $Map[$key] = [ordered]@{
            path         = [string]$info.path
            exists       = [bool]$info.exists
            name         = [string]$info.name
            company      = [string]$info.company
            sources      = @($info.sources)
            curated      = [bool]$info.curated
            writtenAt    = $info.writtenAt
            sessionId    = $null
            clientNumber = $null
            organization = $null
            ticket       = $null
            origin       = 'disk'
        }
    }
    $row = $Map[$key]
    if ($info.exists -and @($info.sources).Count -gt 0) {
        $row.exists = $true
        $row.sources = @($info.sources)
        $row.writtenAt = $info.writtenAt
        $row.curated = [bool]$info.curated
    }
    if ($Hint) {
        $sid = Get-InvestigateProp -Object $Hint -Name 'sessionId'
        $cn = Get-InvestigateProp -Object $Hint -Name 'clientNumber'
        $org = Get-InvestigateProp -Object $Hint -Name 'organization'
        $ticket = Get-InvestigateProp -Object $Hint -Name 'ticket'
        if ($sid) { $row.sessionId = [string]$sid }
        if ($null -ne $cn -and "$cn" -ne '') { $row.clientNumber = [int]$cn }
        if ($org) { $row.organization = [string]$org }
        if ($ticket) { $row.ticket = [string]$ticket }
        $row.origin = 'history'
    }
}

function Get-InvestigatePacks {
    [CmdletBinding()]
    param(
        [object[]]$Hints,
        [string[]]$ExtraRoots,
        [string]$Query,
        [string]$SessionId,
        [int]$Limit = 80
    )
    if ($Limit -lt 1) { $Limit = 80 }
    if ($Limit -gt 200) { $Limit = 200 }

    $map = @{}
    foreach ($h in @($Hints)) {
        $p = [string](Get-InvestigateProp -Object $h -Name 'path')
        if ($p) { Add-InvestigatePackCandidate -Map $map -Path $p -Hint $h -AllowMissing }
    }

    $roots = [System.Collections.Generic.List[string]]::new()
    foreach ($r in @(@($ExtraRoots) + (Get-InvestigatePackRoots))) {
        $n = Get-InvestigateNormalizedPath -Path $r
        if ($n -and (Test-Path -LiteralPath $n -PathType Container) -and -not ($roots.Contains($n))) {
            [void]$roots.Add($n)
        }
    }

    $walked = 0
    foreach ($root in $roots) {
        $perRoot = 0
        $companies = @(Get-ChildItem -LiteralPath $root -Directory -ErrorAction SilentlyContinue)
        foreach ($company in $companies) {
            if ($walked -gt 500 -or $perRoot -gt 200) { break }
            Add-InvestigatePackCandidate -Map $map -Path $company.FullName
            $packs = @(Get-ChildItem -LiteralPath $company.FullName -Directory -ErrorAction SilentlyContinue)
            foreach ($pack in $packs) {
                $walked++
                $perRoot++
                if ($walked -gt 500 -or $perRoot -gt 200) { break }
                Add-InvestigatePackCandidate -Map $map -Path $pack.FullName
                Get-ChildItem -LiteralPath $pack.FullName -Directory -Filter 'Curated_*' -ErrorAction SilentlyContinue |
                    ForEach-Object { Add-InvestigatePackCandidate -Map $map -Path $_.FullName }
            }
        }
    }

    $rows = @($map.Values)
    if ($SessionId) {
        $sid = $SessionId.Trim()
        $rows = @($rows | Where-Object { [string]$_.sessionId -eq $sid })
    }
    $q = if ($Query) { $Query.Trim().ToLowerInvariant() } else { '' }
    if ($q) {
        $rows = @($rows | Where-Object {
                $hay = @(
                    $_.path, $_.name, $_.company, $_.organization, $_.ticket, $_.sessionId, ($_.sources -join ' ')
                ) -join ' '
                $hay.ToLowerInvariant().Contains($q)
            })
    }

    $sorted = @($rows | Sort-Object {
            if ($_.writtenAt) {
                try { [datetime]$_.writtenAt } catch { [datetime]::MinValue }
            } else { [datetime]::MinValue }
        } -Descending)
    $existing = @($sorted | Where-Object { $_.exists -and @($_.sources).Count -gt 0 })
    $missing = @($sorted | Where-Object { -not $_.exists })
    $combined = @($existing + $missing)
    if ($combined.Count -gt $Limit) { $combined = @($combined | Select-Object -First $Limit) }

    return [ordered]@{
        packs = @($combined)
        roots = @($roots)
        count = $combined.Count
    }
}

Export-ModuleMember -Function Get-InvestigateSources, Invoke-InvestigateQuery, Get-InvestigateMailGraph, Get-InvestigateAccessGraph, Export-InvestigateQuery, Get-InvestigateCsvPath, Get-InvestigatePacks, Get-InvestigatePackRoots, Parse-InvestigateConditionalAccess, ConvertTo-InvestigateCaResultLabel
