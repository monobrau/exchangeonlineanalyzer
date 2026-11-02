# LogCuration.psm1
# Facet-based include/exclude curation of EOA report-folder CSVs.
# Never modifies original export files; writes Curated_<timestamp>\ beside them.

$script:CurationSources = @(
    @{
        Name   = 'SignInLogs'
        Facets = @(
            @{ Name = 'UserPrincipalName'; Columns = @('UserPrincipalName', 'UPN', 'User') }
            @{ Name = 'IPAddress'; Columns = @('IPAddress', 'IpAddress', 'IP') }
            @{ Name = 'CountryOrRegion'; Columns = @('CountryOrRegion', 'Country', 'Location') }
            @{ Name = 'AppDisplayName'; Columns = @('AppDisplayName', 'AppId', 'ResourceDisplayName', 'Application') }
            @{ Name = 'ClientAppUsed'; Columns = @('ClientAppUsed', 'ClientApp', 'ClientAppUsed') }
            @{ Name = 'Status'; Columns = @('Status', 'ResultType', 'ErrorCode') }
            @{ Name = 'FailureReason'; Columns = @('FailureReason', 'Status.FailureReason', 'AdditionalDetails') }
            @{ Name = 'OperatingSystem'; Columns = @('OperatingSystem', 'DeviceDetail.OperatingSystem', 'OS') }
            @{ Name = 'Browser'; Columns = @('Browser', 'DeviceDetail.Browser') }
        )
    }
    @{
        Name   = 'GraphAuditLogs'
        Facets = @(
            @{ Name = 'InitiatedBy'; Columns = @('InitiatedBy', 'InitiatedByUser', 'UserPrincipalName', 'Actor') }
            @{ Name = 'ActivityDisplayName'; Columns = @('ActivityDisplayName', 'Activity', 'Operation') }
            @{ Name = 'Category'; Columns = @('Category', 'CategoryDisplayName') }
            @{ Name = 'Result'; Columns = @('Result', 'ResultReason', 'Status') }
            @{ Name = 'TargetResources'; Columns = @('TargetResources', 'Target', 'TargetDisplayName') }
        )
    }
    @{
        Name   = 'UnifiedAuditLogs'
        Facets = @(
            @{ Name = 'UserIds'; Columns = @('UserIds', 'UserId', 'UserPrincipalName', 'UserKey') }
            @{ Name = 'Operations'; Columns = @('Operations', 'Operation') }
            @{ Name = 'RecordType'; Columns = @('RecordType', 'RecordTypeName') }
            @{ Name = 'ClientIP'; Columns = @('ClientIP', 'ClientIp', 'IPAddress') }
            @{ Name = 'ResultStatus'; Columns = @('ResultStatus', 'Result', 'Status') }
        )
    }
    @{
        Name   = 'MessageTrace'
        Facets = @(
            @{ Name = 'SenderAddress'; Columns = @('SenderAddress', 'Sender') }
            @{ Name = 'RecipientAddress'; Columns = @('RecipientAddress', 'Recipient') }
            @{ Name = 'Status'; Columns = @('Status') }
            @{ Name = 'FromIP'; Columns = @('FromIP', 'FromIp', 'IPAddress') }
        )
    }
    @{
        Name   = 'InboxRules'
        Facets = @(
            @{ Name = 'Mailbox'; Columns = @('Mailbox', 'UserPrincipalName', 'Identity', 'MailboxOwnerId') }
            @{ Name = 'Name'; Columns = @('Name', 'RuleName', 'DisplayName') }
            @{ Name = 'ForwardTo'; Columns = @('ForwardTo', 'RedirectTo', 'ForwardAsAttachmentTo') }
            @{ Name = 'Enabled'; Columns = @('Enabled', 'State') }
        )
    }
)

function Get-CurationCsvPath {
    param(
        [Parameter(Mandatory)]
        [string]$Folder,
        [Parameter(Mandatory)]
        [string]$BaseName
    )
    if (-not (Test-Path -LiteralPath $Folder)) { return $null }
    $exact = Join-Path $Folder "$BaseName.csv"
    $candidates = [System.Collections.Generic.List[string]]::new()
    if (Test-Path -LiteralPath $exact) { [void]$candidates.Add($exact) }
    Get-ChildItem -LiteralPath $Folder -Filter "${BaseName}*.csv" -File -ErrorAction SilentlyContinue |
        Where-Object {
            $_.Name -notmatch '_curated' -and
            $_.DirectoryName -notmatch '[\\/]Curated_'
        } |
        ForEach-Object { if (-not $candidates.Contains($_.FullName)) { [void]$candidates.Add($_.FullName) } }

    foreach ($p in $candidates) {
        try {
            $lines = Get-Content -LiteralPath $p -TotalCount 2 -ErrorAction Stop
            if ($lines -and $lines.Count -ge 2) { return $p }
        }
        catch { }
    }
    return $null
}

function Resolve-CurationColumn {
    param(
        [Parameter(Mandatory)]
        [object]$SampleRow,
        [Parameter(Mandatory)]
        [string[]]$Candidates
    )
    $names = @($SampleRow.PSObject.Properties.Name)
    foreach ($c in $Candidates) {
        $hit = $names | Where-Object { $_ -eq $c } | Select-Object -First 1
        if ($hit) { return [string]$hit }
    }
    foreach ($c in $Candidates) {
        $hit = $names | Where-Object { $_ -like "*$c*" -or $c -like "*$_*" } | Select-Object -First 1
        if ($hit) { return [string]$hit }
    }
    return $null
}

function Get-CurationCellValue {
    param(
        [object]$Row,
        [string]$Column
    )
    if (-not $Row -or -not $Column) { return '' }
    $raw = $Row.$Column
    if ($null -eq $raw) { return '' }
    $text = [string]$raw
    if ([string]::IsNullOrWhiteSpace($text)) { return '' }
    # Location-style "City, State, Country" — keep full string as facet value
    return $text.Trim()
}

function Import-CurationCsv {
    param([Parameter(Mandatory)][string]$Path)
    try {
        return @(Import-Csv -LiteralPath $Path -Encoding UTF8 -ErrorAction Stop)
    }
    catch {
        try {
            return @(Import-Csv -LiteralPath $Path -ErrorAction Stop)
        }
        catch {
            Write-Warning "Failed to import $Path : $($_.Exception.Message)"
            return @()
        }
    }
}

function Get-LogCurationFacets {
    <#
    .SYNOPSIS
        Inventory EOA report CSVs and return facet value counts for curation UI.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path,

        [int]$TopValues = 40
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Report folder not found: $Path"
    }

    $sourcesOut = [System.Collections.Generic.List[object]]::new()

    foreach ($src in $script:CurationSources) {
        $csvPath = Get-CurationCsvPath -Folder $Path -BaseName $src.Name
        if (-not $csvPath) {
            [void]$sourcesOut.Add([ordered]@{
                    name     = $src.Name
                    present  = $false
                    path     = $null
                    rowCount = 0
                    facets   = @()
                })
            continue
        }

        $rows = Import-CurationCsv -Path $csvPath
        $facetOut = [System.Collections.Generic.List[object]]::new()

        if ($rows.Count -gt 0) {
            foreach ($facetDef in $src.Facets) {
                $col = Resolve-CurationColumn -SampleRow $rows[0] -Candidates $facetDef.Columns
                if (-not $col) { continue }

                $counts = @{}
                foreach ($row in $rows) {
                    $val = Get-CurationCellValue -Row $row -Column $col
                    if ($val -eq '') { $val = '(blank)' }
                    if (-not $counts.ContainsKey($val)) { $counts[$val] = 0 }
                    $counts[$val]++
                }

                $values = @(
                    $counts.GetEnumerator() |
                        Sort-Object { -$_.Value }, Name |
                        Select-Object -First $TopValues |
                        ForEach-Object {
                            [ordered]@{ value = $_.Key; count = [int]$_.Value }
                        }
                )

                [void]$facetOut.Add([ordered]@{
                        name       = $facetDef.Name
                        column     = $col
                        valueCount = $counts.Count
                        values     = $values
                    })
            }
        }

        [void]$sourcesOut.Add([ordered]@{
                name     = $src.Name
                present  = $true
                path     = $csvPath
                rowCount = $rows.Count
                facets   = @($facetOut)
            })
    }

    return [ordered]@{
        path    = $Path
        sources = @($sourcesOut)
    }
}

function Get-CurationRuleProperty {
    param($Rule, [string]$Name)
    if ($null -eq $Rule) { return $null }
    if ($Rule -is [hashtable] -or $Rule -is [System.Collections.IDictionary]) {
        if ($Rule.ContainsKey($Name)) { return $Rule[$Name] }
        foreach ($k in $Rule.Keys) {
            if ([string]$k -eq $Name) { return $Rule[$k] }
        }
        return $null
    }
    return $Rule.$Name
}

function ConvertTo-CurationRuleList {
    param($Rules)
    $list = [System.Collections.Generic.List[object]]::new()
    if ($null -eq $Rules) { return @() }
    foreach ($r in @($Rules)) {
        $source = [string](Get-CurationRuleProperty -Rule $r -Name 'source')
        $facet = [string](Get-CurationRuleProperty -Rule $r -Name 'facet')
        $op = ([string](Get-CurationRuleProperty -Rule $r -Name 'op')).ToLowerInvariant()
        if ($op -notin @('include', 'exclude')) { continue }
        $rawValues = Get-CurationRuleProperty -Rule $r -Name 'values'
        $values = @()
        if ($null -ne $rawValues) {
            $values = @($rawValues | ForEach-Object { [string]$_ })
        }
        if (-not $source -or -not $facet -or $values.Count -eq 0) { continue }
        [void]$list.Add([pscustomobject]@{
                source = $source
                facet  = $facet
                op     = $op
                values = $values
            })
    }
    return @($list)
}

function Test-CurationRowMatch {
    param(
        [Parameter(Mandatory)]
        [object]$Row,

        [Parameter(Mandatory)]
        [hashtable]$ColumnMap,

        [Parameter(Mandatory)]
        [ValidateSet('exclude', 'include')]
        [string]$Mode,

        [AllowEmptyCollection()]
        [array]$RulesForSource = @()
    )

    if (-not $RulesForSource -or $RulesForSource.Count -eq 0) {
        return $true
    }

    $includes = @($RulesForSource | Where-Object { $_.op -eq 'include' })
    $excludes = @($RulesForSource | Where-Object { $_.op -eq 'exclude' })

    foreach ($rule in $excludes) {
        $col = $ColumnMap[$rule.facet]
        if (-not $col) { continue }
        $val = Get-CurationCellValue -Row $Row -Column $col
        if ($val -eq '') { $val = '(blank)' }
        foreach ($v in $rule.values) {
            if ($val -eq $v) { return $false }
        }
    }

    if ($Mode -eq 'include' -or $includes.Count -gt 0) {
        if ($includes.Count -eq 0) {
            # include mode with only excludes already applied
            return ($Mode -ne 'include')
        }
        # Group include rules by facet: AND across facets that have includes; OR within facet
        $byFacet = $includes | Group-Object facet
        foreach ($g in $byFacet) {
            $col = $ColumnMap[$g.Name]
            if (-not $col) { return $false }
            $val = Get-CurationCellValue -Row $Row -Column $col
            if ($val -eq '') { $val = '(blank)' }
            $allowed = @()
            foreach ($rule in $g.Group) { $allowed += $rule.values }
            $allowed = $allowed | Select-Object -Unique
            if ($allowed -notcontains $val) { return $false }
        }
    }

    return $true
}

function Invoke-LogCurationFilter {
    param(
        [Parameter(Mandatory)]
        [string]$Path,

        [Parameter(Mandatory)]
        [ValidateSet('exclude', 'include')]
        [string]$Mode,

        [AllowEmptyCollection()]
        [array]$Rules = @(),

        [switch]$WriteFiles,

        [string]$OutputFolder
    )

    $ruleList = ConvertTo-CurationRuleList -Rules $Rules
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    if ($WriteFiles) {
        if ([string]::IsNullOrWhiteSpace($OutputFolder)) {
            $OutputFolder = Join-Path $Path "Curated_$stamp"
        }
        if (-not (Test-Path -LiteralPath $OutputFolder)) {
            New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
        }
    }

    $fileResults = [System.Collections.Generic.List[object]]::new()
    $written = [System.Collections.Generic.List[string]]::new()

    foreach ($src in $script:CurationSources) {
        $csvPath = Get-CurationCsvPath -Folder $Path -BaseName $src.Name
        if (-not $csvPath) { continue }

        $rows = Import-CurationCsv -Path $csvPath
        $rulesForSource = @($ruleList | Where-Object { $_.source -eq $src.Name })

        $columnMap = @{}
        if ($rows.Count -gt 0) {
            foreach ($facetDef in $src.Facets) {
                $col = Resolve-CurationColumn -SampleRow $rows[0] -Candidates $facetDef.Columns
                if ($col) { $columnMap[$facetDef.Name] = $col }
            }
        }

        $kept = [System.Collections.Generic.List[object]]::new()
        foreach ($row in $rows) {
            if (Test-CurationRowMatch -Row $row -ColumnMap $columnMap -Mode $Mode -RulesForSource $rulesForSource) {
                [void]$kept.Add($row)
            }
        }

        $outName = '{0}_curated.csv' -f $src.Name
        $outPath = $null
        if ($WriteFiles) {
            $outPath = Join-Path $OutputFolder $outName
            if ($kept.Count -gt 0) {
                $kept | Export-Csv -LiteralPath $outPath -NoTypeInformation -Encoding UTF8
            }
            elseif ($rows.Count -gt 0) {
                $headers = @(
                    $rows[0].PSObject.Properties.Name | ForEach-Object {
                        '"' + ([string]$_).Replace('"', '""') + '"'
                    }
                ) -join ','
                Set-Content -LiteralPath $outPath -Value $headers -Encoding UTF8
            }
            else {
                Set-Content -LiteralPath $outPath -Value '' -Encoding UTF8
            }
            [void]$written.Add($outPath)
        }

        [void]$fileResults.Add([ordered]@{
                source      = $src.Name
                inputPath   = $csvPath
                outputPath  = $outPath
                beforeCount = $rows.Count
                afterCount  = $kept.Count
                dropped     = [Math]::Max(0, $rows.Count - $kept.Count)
            })
    }

    if ($WriteFiles) {
        $rulesObj = [ordered]@{
            mode      = $Mode
            generated = (Get-Date).ToString('o')
            sourcePath = $Path
            rules     = @($ruleList)
        }
        $rulesPath = Join-Path $OutputFolder 'CurationRules.json'
        $rulesObj | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $rulesPath -Encoding UTF8
        [void]$written.Add($rulesPath)

        $sb = New-Object System.Text.StringBuilder
        [void]$sb.AppendLine('# Curation Manifest')
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine(('Generated: {0}' -f (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')))
        [void]$sb.AppendLine(('Source: {0}' -f $Path))
        [void]$sb.AppendLine(('Mode: {0}' -f $Mode))
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('## Rules')
        if ($ruleList.Count -eq 0) {
            [void]$sb.AppendLine('_No rules (copy-through)._')
        }
        else {
            foreach ($r in $ruleList) {
                $vals = ($r.values -join ', ')
                [void]$sb.AppendLine(('- **{0}** {1}.{2}: {3}' -f $r.op, $r.source, $r.facet, $vals))
            }
        }
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('## Counts')
        [void]$sb.AppendLine('| Source | Before | After | Dropped |')
        [void]$sb.AppendLine('| --- | ---: | ---: | ---: |')
        foreach ($f in $fileResults) {
            [void]$sb.AppendLine(('| {0} | {1} | {2} | {3} |' -f $f.source, $f.beforeCount, $f.afterCount, $f.dropped))
        }
        $manifestPath = Join-Path $OutputFolder 'CurationManifest.md'
        Set-Content -LiteralPath $manifestPath -Value $sb.ToString() -Encoding UTF8
        [void]$written.Add($manifestPath)
    }

    return [ordered]@{
        path         = $Path
        mode         = $Mode
        outputFolder = if ($WriteFiles) { $OutputFolder } else { $null }
        files        = @($fileResults)
        written      = @($written)
        ruleCount    = $ruleList.Count
    }
}

function Invoke-LogCurationPreview {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path,

        [ValidateSet('exclude', 'include')]
        [string]$Mode = 'exclude',

        [Parameter(Mandatory)]
        [array]$Rules
    )
    return Invoke-LogCurationFilter -Path $Path -Mode $Mode -Rules $Rules
}

function Export-LogCurationSet {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path,

        [ValidateSet('exclude', 'include')]
        [string]$Mode = 'exclude',

        [Parameter(Mandatory)]
        [array]$Rules,

        [string]$OutputFolder
    )
    return Invoke-LogCurationFilter -Path $Path -Mode $Mode -Rules $Rules -WriteFiles -OutputFolder $OutputFolder
}

function Test-CurationPublicIpAddress {
    param([string]$Ip)
    if ([string]::IsNullOrWhiteSpace($Ip)) { return $false }
    $ip = $Ip.Trim()
    # Strip port / brackets if present (e.g. [2001:db8::1]:443 or 1.2.3.4:443)
    if ($ip -match '^\[([^\]]+)\]') { $ip = $Matches[1] }
    elseif ($ip -match '^(\d{1,3}(?:\.\d{1,3}){3}):\d+$') { $ip = $Matches[1] }

    $parsed = $null
    if (-not [System.Net.IPAddress]::TryParse($ip, [ref]$parsed)) { return $false }
    if ($parsed.AddressFamily -ne [System.Net.Sockets.AddressFamily]::InterNetwork) {
        # Treat non-empty IPv6 as public candidate (skip link-local/unique-local later if needed)
        $s = $parsed.ToString()
        if ($s.StartsWith('fe80', [System.StringComparison]::OrdinalIgnoreCase)) { return $false }
        if ($s.StartsWith('fc', [System.StringComparison]::OrdinalIgnoreCase) -or $s.StartsWith('fd', [System.StringComparison]::OrdinalIgnoreCase)) { return $false }
        return $true
    }

    $b = $parsed.GetAddressBytes()
    # 10/8
    if ($b[0] -eq 10) { return $false }
    # 127/8
    if ($b[0] -eq 127) { return $false }
    # 0.0.0.0/8
    if ($b[0] -eq 0) { return $false }
    # 169.254/16
    if ($b[0] -eq 169 -and $b[1] -eq 254) { return $false }
    # 172.16/12
    if ($b[0] -eq 172 -and $b[1] -ge 16 -and $b[1] -le 31) { return $false }
    # 192.168/16
    if ($b[0] -eq 192 -and $b[1] -eq 168) { return $false }
    # 100.64/10 CGNAT
    if ($b[0] -eq 100 -and $b[1] -ge 64 -and $b[1] -le 127) { return $false }
    return $true
}

function Get-LogCurationWanIpSuggestions {
    <#
    .SYNOPSIS
        Rank likely tenant physical WAN / office egress IPs from EOA export CSVs.
    .DESCRIPTION
        Uses successful SignInLogs public IPs (primary), cross-checks UnifiedAuditLogs ClientIP
        and MessageTrace FromIP. Private/RFC1918 addresses are ignored. Does not call Liongard.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path,

        [int]$Top = 12
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Report folder not found: $Path"
    }

    $stats = @{}

    function Add-WanStat {
        param(
            [string]$Ip,
            [string]$Source,
            [bool]$Success,
            [string]$Country,
            [string]$User
        )
        if (-not (Test-CurationPublicIpAddress -Ip $Ip)) { return }
        $key = $Ip.Trim()
        if ($key -match '^(\d{1,3}(?:\.\d{1,3}){3}):\d+$') { $key = $Matches[1] }
        if (-not $stats.ContainsKey($key)) {
            $stats[$key] = [ordered]@{
                ip            = $key
                total         = 0
                success       = 0
                failure       = 0
                signIn        = 0
                ual           = 0
                messageTrace  = 0
                users         = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
                countries     = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
                sources       = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
            }
        }
        $s = $stats[$key]
        $s.total++
        [void]$s.sources.Add($Source)
        if ($Source -eq 'SignInLogs') { $s.signIn++ }
        elseif ($Source -eq 'UnifiedAuditLogs') { $s.ual++ }
        elseif ($Source -eq 'MessageTrace') { $s.messageTrace++ }
        if ($Success) { $s.success++ } else { $s.failure++ }
        if ($User) { [void]$s.users.Add($User) }
        if ($Country -and $Country -ne '(blank)') { [void]$s.countries.Add($Country) }
    }

    # Sign-in logs (primary)
    $slPath = Get-CurationCsvPath -Folder $Path -BaseName 'SignInLogs'
    if ($slPath) {
        $rows = Import-CurationCsv -Path $slPath
        if ($rows.Count -gt 0) {
            $ipCol = Resolve-CurationColumn -SampleRow $rows[0] -Candidates @('IPAddress', 'IpAddress', 'IP')
            $stCol = Resolve-CurationColumn -SampleRow $rows[0] -Candidates @('Status', 'ResultType', 'ErrorCode')
            $coCol = Resolve-CurationColumn -SampleRow $rows[0] -Candidates @('CountryOrRegion', 'Country', 'Location')
            $upCol = Resolve-CurationColumn -SampleRow $rows[0] -Candidates @('UserPrincipalName', 'UPN', 'User')
            foreach ($row in $rows) {
                $ip = Get-CurationCellValue -Row $row -Column $ipCol
                if (-not $ip) { continue }
                $status = Get-CurationCellValue -Row $row -Column $stCol
                $success = $status -match '(?i)success|0' -and $status -notmatch '(?i)fail'
                if ($status -match '(?i)fail') { $success = $false }
                Add-WanStat -Ip $ip -Source 'SignInLogs' -Success:$success `
                    -Country (Get-CurationCellValue -Row $row -Column $coCol) `
                    -User (Get-CurationCellValue -Row $row -Column $upCol)
            }
        }
    }

    # UAL client IPs
    $ualPath = Get-CurationCsvPath -Folder $Path -BaseName 'UnifiedAuditLogs'
    if ($ualPath) {
        $rows = Import-CurationCsv -Path $ualPath
        if ($rows.Count -gt 0) {
            $ipCol = Resolve-CurationColumn -SampleRow $rows[0] -Candidates @('ClientIP', 'ClientIp', 'IPAddress')
            $upCol = Resolve-CurationColumn -SampleRow $rows[0] -Candidates @('UserIds', 'UserId', 'UserPrincipalName')
            foreach ($row in $rows) {
                $ip = Get-CurationCellValue -Row $row -Column $ipCol
                if (-not $ip) { continue }
                Add-WanStat -Ip $ip -Source 'UnifiedAuditLogs' -Success:$true `
                    -Country '' -User (Get-CurationCellValue -Row $row -Column $upCol)
            }
        }
    }

    # MessageTrace FromIP (often gateway / related; weaker signal)
    $mtPath = Get-CurationCsvPath -Folder $Path -BaseName 'MessageTrace'
    if ($mtPath) {
        $rows = Import-CurationCsv -Path $mtPath
        if ($rows.Count -gt 0) {
            $ipCol = Resolve-CurationColumn -SampleRow $rows[0] -Candidates @('FromIP', 'FromIp', 'IPAddress')
            foreach ($row in $rows) {
                $ip = Get-CurationCellValue -Row $row -Column $ipCol
                if (-not $ip) { continue }
                Add-WanStat -Ip $ip -Source 'MessageTrace' -Success:$true -Country '' -User ''
            }
        }
    }

    $suggestions = foreach ($entry in $stats.GetEnumerator()) {
        $s = $entry.Value
        # Prefer IPs seen in successful sign-ins; require at least one SignIn hit for "likely WAN"
        if ($s.signIn -le 0 -and $s.success -le 0) { continue }
        if ($s.signIn -le 0) { continue }

        $score = ($s.success * 10) + ($s.users.Count * 5) + ($s.signIn * 2) + $s.ual
        if ($s.failure -gt $s.success -and $s.success -eq 0) { $score = $score - 20 }

        $reasons = [System.Collections.Generic.List[string]]::new()
        if ($s.success -gt 0) { [void]$reasons.Add(('{0} successful sign-in(s)' -f $s.success)) }
        if ($s.users.Count -gt 0) { [void]$reasons.Add(('{0} user(s)' -f $s.users.Count)) }
        if ($s.ual -gt 0) { [void]$reasons.Add('also in UAL') }
        if ($s.messageTrace -gt 0) { [void]$reasons.Add('also in MessageTrace FromIP') }
        if ($s.countries.Count -gt 0) { [void]$reasons.Add(('countries: {0}' -f (($s.countries | Select-Object -First 3) -join ', '))) }

        [ordered]@{
            ip           = $s.ip
            score        = [int]$score
            successCount = [int]$s.success
            failureCount = [int]$s.failure
            signInCount  = [int]$s.signIn
            ualCount     = [int]$s.ual
            messageTraceCount = [int]$s.messageTrace
            userCount    = [int]$s.users.Count
            countries    = @($s.countries)
            sources      = @($s.sources)
            reason       = ($reasons -join '; ')
            suggested    = ($s.success -gt 0)
        }
    }

    $ranked = @($suggestions | Sort-Object { -$_.score }, { -$_.successCount }, ip | Select-Object -First $Top)

    return [ordered]@{
        path        = $Path
        count       = $ranked.Count
        suggestions = $ranked
        note        = 'Public IPs with successful SignInLogs activity, ranked as likely office/WAN egress. Verify before excluding. Private RFC1918 addresses are omitted. Paste known WAN IPs manually if Liongard/firewall values are authoritative.'
    }
}

# Fold WAN suggestions into facets payload for one-shot UI load
function Get-LogCurationFacetsWithWan {
    param(
        [Parameter(Mandatory)]
        [string]$Path,
        [int]$TopValues = 40,
        [int]$WanTop = 12
    )
    $facets = Get-LogCurationFacets -Path $Path -TopValues $TopValues
    $wan = Get-LogCurationWanIpSuggestions -Path $Path -Top $WanTop
    return [ordered]@{
        path           = $facets.path
        sources        = $facets.sources
        wanSuggestions = $wan
    }
}

Export-ModuleMember -Function Get-LogCurationFacets, Get-LogCurationWanIpSuggestions, Get-LogCurationFacetsWithWan, Invoke-LogCurationPreview, Export-LogCurationSet
