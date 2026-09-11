#Requires -Version 5.1
#Requires -Modules Pester

$modulePath = Join-Path (Split-Path $PSScriptRoot -Parent) 'Modules\LogInvestigate.psm1'

Describe 'LogInvestigate' {
    BeforeAll {
        Import-Module $modulePath -Force
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
        ($src.sources | Where-Object { $_.present }).Count | Should Be 3
        ($src.sources | Where-Object { $_.name -eq 'MessageTrace' }).rowCount | Should Be 3
    }

    It 'resolves ticket-suffixed MessageTrace CSV names' {
        $alt = Join-Path $TestDrive 'TicketPack'
        New-Item -ItemType Directory -Path $alt | Out-Null
        Copy-Item (Join-Path $script:Folder 'MessageTrace.csv') (Join-Path $alt 'MessageTrace_12345.csv')
        $p = Get-InvestigateCsvPath -Folder $alt -BaseName 'MessageTrace'
        $p | Should Match 'MessageTrace_12345.csv$'
    }

    It 'filters message trace by address and paginates' {
        $q = Invoke-InvestigateQuery -Path $script:Folder -Source 'MessageTrace' -Filters @(
            @{ column = '_address'; op = 'icontains'; value = 'mallory' }
        ) -Limit 50 -Offset 0
        $q.total | Should Be 1
        $q.rows[0].RecipientAddress | Should Be 'mallory@fabrikam.com'
    }

    It 'builds a mail graph and can hide internal edges' {
        $all = Get-InvestigateMailGraph -Path $script:Folder -MinCount 1 -MaxNodes 80
        $all.matched | Should Be 3
        $all.edges.Count | Should Be 3

        $ext = Get-InvestigateMailGraph -Path $script:Folder -MinCount 1 -MaxNodes 80 -HideInternal
        $pairs = @($ext.edges | ForEach-Object { "$($_.from)|$($_.to)" })
        ($pairs -contains 'alice@contoso.com|bob@contoso.com') | Should Be $false
        ($pairs -contains 'alice@contoso.com|mallory@fabrikam.com') | Should Be $true
    }

    It 'exposes UAL AuditData fields and filters by operation' {
        $q = Invoke-InvestigateQuery -Path $script:Folder -Source 'UnifiedAuditLogs' -Filters @(
            @{ column = '_Operation'; op = 'icontains'; value = 'MailItemsAccessed' }
        )
        $q.total | Should Be 1
        $q.rows[0]._MailboxOwnerUPN | Should Be 'alice@contoso.com'
    }

    It 'exports a filtered CSV beside the pack' {
        $ex = Export-InvestigateQuery -Path $script:Folder -Source 'SignInLogs' -Filters @(
            @{ column = '_address'; op = 'icontains'; value = 'alice' }
        )
        (Test-Path -LiteralPath $ex.file) | Should Be $true
        $ex.rowCount | Should Be 1
        $ex.file | Should Match 'Investigate_SignInLogs_'
    }
}
