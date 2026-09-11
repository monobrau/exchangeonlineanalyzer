#Requires -Version 5.1
#Requires -Modules Pester

$modulePath = Join-Path (Split-Path $PSScriptRoot -Parent) 'Modules\ExportUtils.psm1'
Import-Module $modulePath -Force

function global:Fake-GetMessageTraceV2 {
    param(
        $StartDate,
        $EndDate,
        $StartingRecipientAddress,
        $ResultSize = 1000,
        $SenderAddress,
        $RecipientAddress,
        $ErrorAction
    )
    [void]$global:EoaMtTestCalls.Add([pscustomobject]@{
        StartDate                = $StartDate
        EndDate                  = $EndDate
        StartingRecipientAddress = $StartingRecipientAddress
        ResultSize               = $ResultSize
        SenderAddress            = $SenderAddress
        RecipientAddress         = $RecipientAddress
    })

    if ($global:EoaMtTestStall) {
        $t = [datetime]::SpecifyKind([datetime]'2026-08-31T12:00:00', 'Utc')
        $out = New-Object System.Collections.ArrayList
        for ($i = 0; $i -lt [int]$ResultSize; $i++) {
            [void]$out.Add([pscustomobject]@{
                Received         = $t
                RecipientAddress = 'victim@contoso.com'
                SenderAddress    = "sender$i@fabrikam.com"
                MessageTraceId   = "stall-$i"
                MessageId        = "<stall-$i@contoso.com>"
                Subject          = "stall $i"
            })
        }
        return @($out.ToArray())
    }

    $end = $EndDate
    if ($end -is [string]) { $end = [datetime]$end }
    $keep = New-Object System.Collections.ArrayList
    foreach ($r in @($global:EoaMtTestRows)) {
        if ($null -eq $r -or $null -eq $r.Received) { continue }
        if ($r.Received -lt $StartDate -or $r.Received -gt $end) { continue }
        if ($StartingRecipientAddress) {
            if ($r.Received -eq $end -and [string]::Compare([string]$r.RecipientAddress, [string]$StartingRecipientAddress, $true) -le 0) { continue }
        }
        if ($SenderAddress -and $r.SenderAddress -ne $SenderAddress) { continue }
        if ($RecipientAddress -and $r.RecipientAddress -ne $RecipientAddress) { continue }
        [void]$keep.Add($r)
    }
    $sorted = @($keep | Sort-Object @{ Expression = { $_.Received }; Descending = $true }, @{ Expression = { $_.RecipientAddress }; Descending = $false })
    if ($sorted.Count -eq 0) { return @() }
    $take = [Math]::Min([int]$ResultSize, $sorted.Count)
    return @($sorted[0..($take - 1)])
}

function global:New-EoaMtOrgWideRows {
    param([datetime]$WinEnd, [int]$Total = 6200, [int]$OtherCount = 3800, [string]$Victim = 'victim@contoso.com')
    $rows = New-Object System.Collections.ArrayList
    for ($i = 0; $i -lt $Total; $i++) {
        $received = $WinEnd.AddSeconds(-$i - 1)
        $recip = if ($i -lt $OtherCount) { "user{0:D4}@contoso.com" -f ($i % 50) } else { $Victim }
        [void]$rows.Add([pscustomobject]@{
            Received         = $received
            RecipientAddress = $recip
            SenderAddress    = 'sender@fabrikam.com'
            MessageTraceId   = "id-$i"
            MessageId        = "<id-$i@contoso.com>"
            Subject          = "msg $i"
        })
    }
    return @($rows.ToArray())
}

Describe 'MessageTrace V2 paging' {
    BeforeAll {
        $script:V2Cmd = Get-Command Fake-GetMessageTraceV2
        $script:WinStart = [datetime]::SpecifyKind([datetime]'2026-08-31T00:00:00', 'Utc')
        $script:WinEnd = [datetime]::SpecifyKind([datetime]'2026-09-01T00:00:00', 'Utc')
    }

    BeforeEach {
        $global:EoaMtTestCalls = New-Object System.Collections.ArrayList
        $global:EoaMtTestStall = $false
        $global:EoaMtTestRows = @()
        $global:EOA_ExchangeHasUserMailboxes = $true
    }

    It 'walks EndDate and StartingRecipientAddress for an org-wide day with 6200 rows' {
        $global:EoaMtTestRows = New-EoaMtOrgWideRows -WinEnd $script:WinEnd
        $page = InModuleScope ExportUtils -Parameters @{ Cmd = $script:V2Cmd; WinStart = $script:WinStart; WinEnd = $script:WinEnd } {
            $script:EOA_MessageTraceCallCount = 0
            Invoke-EoaMessageTraceV2Paged -V2Command $Cmd -WinStart $WinStart -WinEnd $WinEnd -Label '2026-08-31'
        }
        $page.Rows.Count | Should -Be 6200
        $page.Incomplete | Should -Be $false
        $global:EoaMtTestCalls.Count | Should -Be 2
        $global:EoaMtTestCalls[0].ResultSize | Should -Be 5000
        $global:EoaMtTestCalls[0].StartingRecipientAddress | Should -Be $null
        $global:EoaMtTestCalls[0].EndDate | Should -Be $script:WinEnd

        $sorted = @($global:EoaMtTestRows | Sort-Object @{ Expression = { $_.Received }; Descending = $true }, @{ Expression = { $_.RecipientAddress }; Descending = $false })
        $lastPage1 = $sorted[4999]
        $global:EoaMtTestCalls[1].StartingRecipientAddress | Should -Be $lastPage1.RecipientAddress
        $global:EoaMtTestCalls[1].EndDate | Should -Be $lastPage1.Received
        $global:EoaMtTestCalls[1].ResultSize | Should -Be 5000
    }

    It 'keeps 2400 rows to one recipient across the page boundary (mail-bomb case)' {
        $global:EoaMtTestRows = New-EoaMtOrgWideRows -WinEnd $script:WinEnd
        $page = InModuleScope ExportUtils -Parameters @{ Cmd = $script:V2Cmd; WinStart = $script:WinStart; WinEnd = $script:WinEnd } {
            $script:EOA_MessageTraceCallCount = 0
            Invoke-EoaMessageTraceV2Paged -V2Command $Cmd -WinStart $WinStart -WinEnd $WinEnd -Label '2026-08-31'
        }
        $victimRows = @($page.Rows | Where-Object { $_.RecipientAddress -eq 'victim@contoso.com' })
        $victimRows.Count | Should -Be 2400
        $page.Incomplete | Should -Be $false
    }

    It 'returns all 6200 rows for a selected-user RecipientAddress filter, not the first 5000' {
        $rows = New-EoaMtOrgWideRows -WinEnd $script:WinEnd -Total 6200 -OtherCount 0 -Victim 'victim@contoso.com'
        $global:EoaMtTestRows = $rows
        $page = InModuleScope ExportUtils -Parameters @{ Cmd = $script:V2Cmd; WinStart = $script:WinStart; WinEnd = $script:WinEnd } {
            $script:EOA_MessageTraceCallCount = 0
            Invoke-EoaMessageTraceV2Paged -V2Command $Cmd -WinStart $WinStart -WinEnd $WinEnd -RecipientAddress 'victim@contoso.com' -Label 'recipient victim@contoso.com 2026-08-31'
        }
        $page.Rows.Count | Should -Be 6200
        $page.Incomplete | Should -Be $false
        $global:EoaMtTestCalls.Count | Should -Be 2
        $global:EoaMtTestCalls[0].RecipientAddress | Should -Be 'victim@contoso.com'
        $global:EoaMtTestCalls[1].EndDate | Should -Be $rows[4999].Received
        $global:EoaMtTestCalls[1].StartingRecipientAddress | Should -Be 'victim@contoso.com'
    }

    It 'makes one call and stops when the first page is short' {
        $global:EoaMtTestRows = @(
            [pscustomobject]@{
                Received         = $script:WinEnd.AddHours(-1)
                RecipientAddress = 'bob@contoso.com'
                SenderAddress    = 'alice@fabrikam.com'
                MessageTraceId   = 'short-1'
                MessageId        = '<short-1@contoso.com>'
                Subject          = 'hi'
            }
        )
        $page = InModuleScope ExportUtils -Parameters @{ Cmd = $script:V2Cmd; WinStart = $script:WinStart; WinEnd = $script:WinEnd } {
            $script:EOA_MessageTraceCallCount = 0
            Invoke-EoaMessageTraceV2Paged -V2Command $Cmd -WinStart $WinStart -WinEnd $WinEnd -Label '2026-08-31'
        }
        $page.Rows.Count | Should -Be 1
        $page.Incomplete | Should -Be $false
        $global:EoaMtTestCalls.Count | Should -Be 1
        $global:EoaMtTestCalls[0].StartingRecipientAddress | Should -Be $null
    }

    It 'stops and marks incomplete when a full page cannot advance the cursor' {
        $global:EoaMtTestStall = $true
        $stallStart = [datetime]::SpecifyKind([datetime]'2026-08-31T12:00:00', 'Utc')
        $stallEnd = $stallStart.AddMinutes(10)
        $page = InModuleScope ExportUtils -Parameters @{ Cmd = $script:V2Cmd; StallStart = $stallStart; StallEnd = $stallEnd } {
            $script:EOA_MessageTraceCallCount = 0
            Invoke-EoaMessageTraceV2Paged -V2Command $Cmd -WinStart $StallStart -WinEnd $StallEnd -Label 'stall' -MaxRounds 8
        }
        $page.Incomplete | Should -Be $true
        $page.IncompleteReason | Should -Match 'stalled'
        ($global:EoaMtTestCalls.Count -ge 1 -and $global:EoaMtTestCalls.Count -lt 5) | Should -Be $true
    }

    It 'pages a selected-user Get-ExchangeMessageTrace path past 5000 rows' {
        $global:EoaMtTestRows = New-EoaMtOrgWideRows -WinEnd $script:WinEnd -Total 6200 -OtherCount 0 -Victim 'victim@contoso.com'
        $got = InModuleScope ExportUtils -Parameters @{ Cmd = $script:V2Cmd; WinStart = $script:WinStart; WinEnd = $script:WinEnd } {
            $script:EOA_MessageTraceCallCount = 0
            Mock Get-ExchangeMessageTraceV2CommandInfo { $Cmd }
            Mock Test-ExchangeUserMailboxesPresent { $true }
            Mock Add-EoaHistoricalMessageTraceRows { }
            Get-ExchangeMessageTrace -StartDate $WinStart -EndDate $WinEnd -SelectedUsers @('victim@contoso.com')
        }
        @($got).Count | Should -Be 6200
        ($global:EoaMtTestCalls | Where-Object { $_.RecipientAddress -eq 'victim@contoso.com' }).Count | Should -Be 2
        ($global:EoaMtTestCalls | Where-Object { $_.ResultSize -eq 5000 }).Count | Should -Be $global:EoaMtTestCalls.Count
    }

    It 'does not report N of N days when a day was truncated' {
        $warn = InModuleScope ExportUtils {
            $script:LastWarn = $null
            Mock Write-Warning { $script:LastWarn = [string]$Message }
            $trunc = New-Object 'System.Collections.Generic.List[string]'
            [void]$trunc.Add('2026-08-31')
            Write-EoaMessageTraceCoverage -EntryCount 1000 -DaysRequested 7 -TruncatedDays $trunc
            $script:LastWarn
        }
        $warn | Should -Match 'INCOMPLETE'
        $warn | Should -Match 'Truncated'
        $warn | Should -Match '6 of 7 days'
        $warn | Should -Not -Match '7 of 7 days'
        $warn | Should -Not -Match 'No data for'
    }
}
