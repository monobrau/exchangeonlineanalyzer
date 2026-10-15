#Requires -Version 5.1
#Requires -Modules Pester

$modulePath = Join-Path (Split-Path $PSScriptRoot -Parent) 'Modules\LogInvestigate.psm1'
Import-Module $modulePath -Force

Describe 'LogInvestigate' {
    BeforeAll {
        $script:Folder = Join-Path $TestDrive 'ContosoPack'
        New-Item -ItemType Directory -Path $script:Folder | Out-Null

        @(
            'Received,SenderAddress,RecipientAddress,Subject,Status,FromIP,Size'
            '2026-01-15T10:00:00Z,alice@contoso.com,bob@contoso.com,Hello,Delivered,203.0.113.10,1234'
            '2026-01-15T11:00:00Z,alice@contoso.com,mallory@fabrikam.com,Invoice,Delivered,203.0.113.10,999'
            '2026-01-15T11:30:00Z,eve@fabrikam.com,alice@contoso.com,Re: Invoice,Delivered,198.51.100.8,200'
        ) | Set-Content -LiteralPath (Join-Path $script:Folder 'MessageTrace.csv') -Encoding UTF8

        @(
            'CreatedDateTime,UserPrincipalName,IPAddress,AppDisplayName,Status,CountryOrRegion,ClientAppUsed'
            '2026-01-15T10:05:00Z,alice@contoso.com,203.0.113.10,Outlook,Success,US,Browser'
            '2026-01-15T12:00:00Z,eve@fabrikam.com,198.51.100.8,Browser,Failure,DE,Browser'
        ) | Set-Content -LiteralPath (Join-Path $script:Folder 'SignInLogs.csv') -Encoding UTF8

        $audit = '{"Operation":"MailItemsAccessed","ClientIP":"203.0.113.10","MailboxOwnerUPN":"alice@contoso.com","UserId":"alice@contoso.com"}'
        $audit = $audit.Replace('"', '""')
        @(
            'CreationDate,UserIds,Operations,RecordType,ClientIP,ResultStatus,AuditData'
            "2026-01-15T10:10:00Z,alice@contoso.com,MailItemsAccessed,ExchangeItem,203.0.113.10,Success,`"$audit`""
            '2026-01-16T08:00:00Z,bob@contoso.com,UpdateInboxRules,ExchangeAdmin,203.0.113.20,Success,{}'
        ) | Set-Content -LiteralPath (Join-Path $script:Folder 'UnifiedAuditLogs.csv') -Encoding UTF8
    }

    It 'inventories MessageTrace, SignInLogs, and UnifiedAuditLogs' {
        $src = Get-InvestigateSources -Path $script:Folder
        ($src.sources | Where-Object { $_.present }).Count | Should -Be 3
        ($src.sources | Where-Object { $_.name -eq 'MessageTrace' }).rowCount | Should -Be 3
    }

    It 'resolves ticket-suffixed MessageTrace CSV names' {
        $alt = Join-Path $TestDrive 'TicketPack'
        New-Item -ItemType Directory -Path $alt | Out-Null
        Copy-Item (Join-Path $script:Folder 'MessageTrace.csv') (Join-Path $alt 'MessageTrace_12345.csv')
        $p = Get-InvestigateCsvPath -Folder $alt -BaseName 'MessageTrace'
        $p | Should -Match 'MessageTrace_12345.csv$'
    }

    It 'filters message trace by address and paginates' {
        $q = Invoke-InvestigateQuery -Path $script:Folder -Source 'MessageTrace' -Filters @(
            @{ column = '_address'; op = 'icontains'; value = 'mallory' }
        ) -Limit 50 -Offset 0
        $q.total | Should -Be 1
        $q.rows[0].RecipientAddress | Should -Be 'mallory@fabrikam.com'
    }

    It 'builds a mail graph and can hide internal edges' {
        $all = Get-InvestigateMailGraph -Path $script:Folder -MinCount 1 -MaxNodes 80
        $all.matched | Should -Be 3
        $all.edges.Count | Should -Be 3

        $ext = Get-InvestigateMailGraph -Path $script:Folder -MinCount 1 -MaxNodes 80 -HideInternal
        $pairs = @($ext.edges | ForEach-Object { "$($_.from)|$($_.to)" })
        ($pairs -contains 'alice@contoso.com|bob@contoso.com') | Should -Be $false
        ($pairs -contains 'alice@contoso.com|mallory@fabrikam.com') | Should -Be $true
    }

    It 'exposes UAL AuditData fields and filters by operation' {
        $q = Invoke-InvestigateQuery -Path $script:Folder -Source 'UnifiedAuditLogs' -Filters @(
            @{ column = '_Operation'; op = 'icontains'; value = 'MailItemsAccessed' }
        )
        $q.total | Should -Be 1
        $q.rows[0]._MailboxOwnerUPN | Should -Be 'alice@contoso.com'
    }

    It 'exports a filtered CSV beside the pack' {
        $ex = Export-InvestigateQuery -Path $script:Folder -Source 'SignInLogs' -Filters @(
            @{ column = '_address'; op = 'icontains'; value = 'alice' }
        )
        (Test-Path -LiteralPath $ex.file) | Should -Be $true
        $ex.rowCount | Should -Be 1
        $ex.file | Should -Match 'Investigate_SignInLogs_'
    }

    It 'builds a login graph from sign-in and UAL IPs' {
        $g = Get-InvestigateAccessGraph -Path $script:Folder -MinCount 1 -MaxNodes 80
        $g.signInMatched | Should -Be 2
        $g.ualMatched | Should -Be 2
        $pairs = @($g.edges | ForEach-Object { "$($_.from)|$($_.to)" })
        ($pairs -contains 'alice@contoso.com|203.0.113.10') | Should -Be $true
        ($pairs -contains 'eve@fabrikam.com|198.51.100.8') | Should -Be $true
        ($pairs -contains 'bob@contoso.com|203.0.113.20') | Should -Be $true
        $de = @($g.countries | Where-Object { $_.country -eq 'DE' })
        $de.Count | Should -Be 1
        $de[0].failCount | Should -Be 1
        $eve = @($g.nodes | Where-Object { $_.id -eq 'i:198.51.100.8' })
        $eve[0].failCount | Should -Be 1
        $eve[0].country | Should -Be 'DE'
    }

    It 'can omit UAL and skip private IPs on the login graph' {
        $noUal = Get-InvestigateAccessGraph -Path $script:Folder -IncludeUal:$false
        $ips = @($noUal.nodes | Where-Object { $_.kind -eq 'ip' } | ForEach-Object { $_.label })
        ($ips -contains '203.0.113.20') | Should -Be $false
        ($ips -contains '203.0.113.10') | Should -Be $true
    }

    It 'parses country from Location city/state/US strings' {
        $alt = Join-Path $TestDrive 'GeoPack'
        New-Item -ItemType Directory -Path $alt | Out-Null
        @(
            'CreatedDateTime,UserPrincipalName,IPAddress,Location,Status'
            '2026-01-15T10:05:00Z,alice@contoso.com,203.0.113.10,"Austin, TX, US",Success'
        ) | Set-Content -LiteralPath (Join-Path $alt 'SignInLogs.csv') -Encoding UTF8
        $g = Get-InvestigateAccessGraph -Path $alt -IncludeUal:$false
        $us = @($g.countries | Where-Object { $_.country -eq 'US' })
        $us.Count | Should -Be 1
    }

    It 'lists timestamped packs and curated folders under an extra root' {
        $root = Join-Path $TestDrive 'SecurityInvestigation'
        $pack = Join-Path $root 'Fabrikam\20260115_100000'
        New-Item -ItemType Directory -Path $pack | Out-Null
        Copy-Item (Join-Path $script:Folder 'MessageTrace.csv') (Join-Path $pack 'MessageTrace.csv')
        $curated = Join-Path $pack 'Curated_20260115_110000'
        New-Item -ItemType Directory -Path $curated | Out-Null
        Copy-Item (Join-Path $script:Folder 'SignInLogs.csv') (Join-Path $curated 'SignInLogs.csv')
        $listed = Get-InvestigatePacks -ExtraRoots $root -Query '20260115_100000'
        (@($listed.packs | Where-Object { $_.name -eq '20260115_100000' })).Count | Should -Be 1
        (@($listed.packs | Where-Object { $_.curated -eq $true })).Count | Should -Be 1
    }

    It 'attaches session hints to known pack paths' {
        $listed = Get-InvestigatePacks -Hints @(@{
                path         = $script:Folder
                sessionId    = 'EOA_BulkWeb_hint'
                clientNumber = 1
                ticket       = '99'
            }) -SessionId 'EOA_BulkWeb_hint'
        $hit = @($listed.packs | Where-Object { $_.ticket -eq '99' }) | Select-Object -First 1
        $hit.sessionId | Should -Be 'EOA_BulkWeb_hint'
        (@($hit.sources) -contains 'MessageTrace') | Should -Be $true
    }
}

Describe 'Investigate Conditional Access parse' {
    It 'normalizes Graph appliedConditionalAccessPolicies results' {
        ConvertTo-InvestigateCaResultLabel -Result 'success' | Should -Be 'Success'
        ConvertTo-InvestigateCaResultLabel -Result 'failure' | Should -Be 'Failed'
        ConvertTo-InvestigateCaResultLabel -Result 'notApplied' | Should -Be 'Not Applied'
        ConvertTo-InvestigateCaResultLabel -Result 'reportOnly' | Should -Be 'Report-only'
        ConvertTo-InvestigateCaResultLabel -Result 'reportOnlyFailure' | Should -Be 'Report-only'
    }

    It 'parses appliedConditionalAccessPolicies JSON' {
        $json = '[{"id":"p1","displayName":"Require MFA","result":"success"},{"id":"p2","displayName":"Block legacy","result":"notApplied"},{"id":"p3","displayName":"Geo report","result":"reportOnly"}]'
        $parsed = Parse-InvestigateConditionalAccess -Json $json -Status 'success'
        $parsed.policies.Count | Should -Be 3
        $parsed.policies[0].result | Should -Be 'Success'
        $parsed.policies[0].tone | Should -Be 'ok'
        $parsed.policies[1].result | Should -Be 'Not Applied'
        $parsed.policies[2].result | Should -Be 'Report-only'
        $parsed.policies[2].tone | Should -Be 'report'
        $parsed.summary | Should -Match 'Require MFA:Success'
    }

    It 'falls back to semicolon CAPolicy columns' {
        $parsed = Parse-InvestigateConditionalAccess -Status 'failure' -Names 'Require MFA; Block countries' -Results 'Failed; Not Applied'
        $parsed.policies.Count | Should -Be 2
        $parsed.policies[0].name | Should -Be 'Require MFA'
        $parsed.policies[0].result | Should -Be 'Failed'
        $parsed.policies[0].tone | Should -Be 'fail'
        $parsed.policies[1].result | Should -Be 'Not Applied'
    }

    It 'exposes CA columns and filters _CA on SignInLogs query' {
        $folder = Join-Path $TestDrive 'CaPack'
        New-Item -ItemType Directory -Path $folder | Out-Null
        $json = '[{"id":"p1","displayName":"Require MFA","result":"failure"}]'
        $jsonCsv = $json.Replace('"', '""')
        @(
            'CreatedDateTime,UserPrincipalName,IPAddress,AppDisplayName,Status,CountryOrRegion,ClientAppUsed,ConditionalAccessStatus,CAPolicyNames,CAPolicyResults,CAPolicyDetails,ConditionalAccessPoliciesJSON'
            "2026-01-15T10:05:00Z,alice@contoso.com,203.0.113.10,Outlook,Success,US,Browser,failure,Require MFA,Failed,Require MFA (BLOCKED),`"$jsonCsv`""
            '2026-01-15T12:00:00Z,eve@fabrikam.com,198.51.100.8,Browser,Success,DE,Browser,notApplied,,,,'
        ) | Set-Content -LiteralPath (Join-Path $folder 'SignInLogs.csv') -Encoding UTF8

        $q = Invoke-InvestigateQuery -Path $folder -Source 'SignInLogs'
        ($q.columns -contains 'ConditionalAccessStatus') | Should -Be $true
        ($q.columns -contains '_CA') | Should -Be $true
        ($q.columns -contains '_CAPolicies') | Should -Be $true
        $q.total | Should -Be 2
        $alice = @($q.rows | Where-Object { $_.UserPrincipalName -eq 'alice@contoso.com' })[0]
        $alice._CA | Should -Match 'Require MFA'
        $alice._CAPolicies | Should -Match 'Failed'

        $blocked = Invoke-InvestigateQuery -Path $folder -Source 'SignInLogs' -Filters @(
            @{ column = '_CA'; op = 'icontains'; value = 'Require MFA' }
        )
        $blocked.total | Should -Be 1
        $blocked.rows[0].UserPrincipalName | Should -Be 'alice@contoso.com'
    }
}
