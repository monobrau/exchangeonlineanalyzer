#Requires -Version 5.1
#Requires -Modules Pester

$modulePath = Join-Path (Split-Path $PSScriptRoot -Parent) 'Modules\ConnectWiseManage.psm1'
Import-Module $modulePath -Force

Describe 'Threat hunter Manage post plan' {
    It 'uses the live Security board, status, and work type ids' {
        $catalog = Get-ConnectWiseManageHunterCatalog
        $catalog.BoardId | Should -Be 97
        $catalog.BoardName | Should -Be 'Security'
        $catalog.MemberIdentifier | Should -Be 'cknospe'
        $catalog.MemberId | Should -Be 301
        $catalog.New.Id | Should -Be 1831
        $catalog.New.Name | Should -Be 'New - Save 1st to send user Ticket Information'
        $catalog.Dispatch.Id | Should -Be 1825
        $catalog.Dispatch.Name | Should -Be 'Dispatch'
        $catalog.InProgress.Id | Should -Be 1829
        $catalog.ResourceSchedule.TypeId | Should -Be 4
        $catalog.ResourceSchedule.StatusId | Should -Be 2
        $catalog.ResourceSchedule.SpanId | Should -Be 1
        $catalog.ResourceSchedule.WhereId | Should -Be 4
        $catalog.InProgress.Name | Should -Be 'In Progress'
        ($catalog.Statuses | Where-Object Key -eq 'review').Id | Should -Be 1841
        ($catalog.Statuses | Where-Object Key -eq 'waiting').Id | Should -Be 1854
        ($catalog.WorkTypes | Where-Object Key -eq 'remote').Id | Should -Be 3
        ($catalog.WorkTypes | Where-Object Key -eq 'remote').BillableOption | Should -Be 'Billable'
        ($catalog.WorkTypes | Where-Object Key -eq 'communication').Id | Should -Be 82
        ($catalog.WorkTypes | Where-Object Key -eq 'communication').BillableOption | Should -Be 'DoNotBill'
    }

    It 'plans a Waiting on Client update with Communication time on the client email only' {
        $plan = Resolve-ConnectWiseManageHunterPost -Request @{
            ticketId          = '1901053'
            status            = 'waiting'
            includeRecipients = $true
            time              = 'communication'
            hours             = '0.5'
            notes             = @{ client = 'Hello from the check.'; internal = 'Technical detail.'; incident = '' }
        }
        $plan.StatusId | Should -Be 1854
        $plan.StatusName | Should -Be 'Waiting on Client'
        $plan.WorkTypeId | Should -Be 82
        $plan.BillableOption | Should -Be 'DoNotBill'
        $plan.Hours | Should -Be 0.5
        $plan.NotifyClient | Should -BeTrue
        $plan.NotifyIncident | Should -BeFalse
        $plan.ClientText | Should -Be 'Hello from the check.'
        $plan.InternalText | Should -Be 'Technical detail.'
        $plan.IncidentText | Should -Be ''
        $plan.MemberIdentifier | Should -Be 'cknospe'
        $plan.PSObject.Properties.Name | Should -Not -Contain 'ClosedFlag'
    }

    It 'plans Review with no time and no email when those are off' {
        $plan = Resolve-ConnectWiseManageHunterPost -Request @{
            ticketId          = '#1841'
            status            = 'review'
            includeRecipients = $false
            time              = 'none'
            notes             = @{ client = 'Client note'; internal = ''; incident = 'Incident note' }
        }
        $plan.StatusId | Should -Be 1841
        $plan.TimeMode | Should -Be 'none'
        $plan.WorkTypeId | Should -Be 0
        $plan.NotifyClient | Should -BeFalse
        $plan.NotifyIncident | Should -BeFalse
        $plan.IncidentText | Should -Be 'Incident note'
    }

    It 'pins the Security statuses and drops zClosed names' {
        $names = @(Get-ConnectWiseManageHunterStatusNames -Statuses @(
            [pscustomobject]@{ name = 'zClosed Loop - Contact' }
            [pscustomobject]@{ name = 'Pre-Process' }
            [pscustomobject]@{ name = 'Review' }
            [pscustomobject]@{ name = 'Dispatch' }
            [pscustomobject]@{ name = 'New - Save 1st to send user Ticket Information' }
            [pscustomobject]@{ name = 'Waiting on Client'; inactive = $true }
        ))
        $names[0] | Should -Be 'New - Save 1st to send user Ticket Information'
        $names[1] | Should -Be 'Dispatch'
        $names[2] | Should -Be 'Review'
        $names | Should -Contain 'Pre-Process'
        $names | Should -Not -Contain 'zClosed Loop - Contact'
        $names | Should -Not -Contain 'Waiting on Client'
    }

    It 'classifies notes the way Manage groups them' {
        ConvertTo-ConnectWiseManageHunterNoteKind ([pscustomobject]@{ resolutionFlag = $true; internalAnalysisFlag = $true }) | Should -Be 'resolution'
        ConvertTo-ConnectWiseManageHunterNoteKind ([pscustomobject]@{ internalAnalysisFlag = $true; detailDescriptionFlag = $true }) | Should -Be 'internal'
        ConvertTo-ConnectWiseManageHunterNoteKind ([pscustomobject]@{ detailDescriptionFlag = $true }) | Should -Be 'discussion'
        ConvertTo-ConnectWiseManageHunterNoteKind ([pscustomobject]@{}) | Should -Be 'other'
    }

    It 'records a failed Manage write and reads it back for that ticket' {
        $path = Join-Path $TestDrive 'hunter-manage-log.jsonl'
        Add-ConnectWiseManageHunterAudit -Path $path -TicketId '1901051' -Action 'Status' -Outcome 'failed' -Summary 'Status: Manage left the status at Dispatch.' -Member 'cknospe'
        Add-ConnectWiseManageHunterAudit -Path $path -TicketId '1901057' -Action 'Owner' -Outcome 'ok' -Summary 'Owner set to cknospe.' -Member 'cknospe'
        $rows = @(Get-ConnectWiseManageHunterAudit -Path $path -TicketId '1901051')
        $rows.Count | Should -Be 1
        $rows[0].outcome | Should -Be 'failed'
        $rows[0].summary | Should -Be 'Status: Manage left the status at Dispatch.'
        (Get-Content -LiteralPath $path -Raw) | Should -Not -Match 'privateKey|publicKey'
    }

    It 'plans a closing status with the start time and hours' {
        $plan = Resolve-ConnectWiseManageHunterStatusChange -Request @{
            ticketId  = '1901051'
            statusId  = 1814
            time      = 'remote'
            hours     = '0.5'
            timeStart = '2026-09-26T15:00'
        } -Choices @(
            [pscustomobject]@{ id = 1814; name = 'Closed'; closed = $true }
            [pscustomobject]@{ id = 1841; name = 'Review'; closed = $false }
        )
        $plan.StatusName | Should -Be 'Closed'
        $plan.Closed | Should -BeTrue
        $plan.WorkTypeId | Should -Be 3
        $plan.BillableOption | Should -Be 'Billable'
        $plan.Hours | Should -Be 0.5
        $plan.TimeStart.Hour | Should -Be 15
        $plan.TimeEnd | Should -Be $plan.TimeStart.AddMinutes(30)
    }

    It 'does not add time when the status leaves the ticket open' {
        $plan = Resolve-ConnectWiseManageHunterStatusChange -Request @{
            ticketId = '1901051'
            statusId = 1841
            time     = 'remote'
            hours    = '1'
            timeStart = '2026-09-26T15:00'
        } -Choices @([pscustomobject]@{ id = 1841; name = 'Review'; closed = $false })
        $plan.Closed | Should -BeFalse
        $plan.TimeMode | Should -Be 'none'
        $plan.WorkTypeId | Should -Be 0
    }

    It 'plans time on a discussion note and refuses an internal note' {
        $plan = Resolve-ConnectWiseManageHunterNoteTime -Request @{
            ticketId  = '1901051'
            noteId    = '55'
            time      = 'communication'
            hours     = '0.25'
            timeStart = '2026-09-26T15:00'
        } -Note ([pscustomobject]@{ detailDescriptionFlag = $true; text = 'Hello from the existing note.' })
        $plan.NoteText | Should -Be 'Hello from the existing note.'
        $plan.WorkTypeId | Should -Be 82
        $plan.BillableOption | Should -Be 'DoNotBill'
        $plan.Hours | Should -Be 0.25
        $plan.TimeEnd | Should -Be $plan.TimeStart.AddMinutes(15)
        { Resolve-ConnectWiseManageHunterNoteTime -Request @{
            ticketId = '1901051'; noteId = '9'; time = 'remote'; hours = '0.25'; timeStart = '2026-09-26T15:00'
        } -Note ([pscustomobject]@{ internalAnalysisFlag = $true; text = 'Internal only.' }) } | Should -Throw '*discussion*'
    }

    It 'plans an alert triage note without changing the ticket status' {
        $plan = Resolve-ConnectWiseManageHunterPost -Request @{
            ticketId = '1901107'
            status   = 'keep'
            time     = 'none'
            notes    = @{ client = ''; internal = 'HATZ-AI POWERED REVIEW OF THE ALERT ONLY'; incident = '' }
        }
        $plan.KeepStatus | Should -BeTrue
        $plan.StatusId | Should -Be 0
        $plan.InternalText | Should -Match 'HATZ-AI POWERED REVIEW'
        $plan.ClientText | Should -BeNullOrEmpty
    }

    It 'formats a time entry in UTC with a Z so Manage accepts timeStart' {
        $local = [datetimeoffset]::new(2026, 9, 27, 9, 3, 0, [timespan]::FromHours(-5))
        ConvertTo-ConnectWiseManageApiTime $local | Should -Be '2026-09-27T14:03:00Z'
        (ConvertTo-ConnectWiseManageApiTime $local) | Should -Not -Match '[-+]\d\d:\d\d$'
    }

    It 'recognizes the insecure-redirect failure that still saved the note' {
        InModuleScope ConnectWiseManage {
            $msg = 'ConnectWise Manage API request failed: Operation is not valid due to the current state of the object. - Cannot follow an insecure redirection by default. Reissue the command specifying the -AllowInsecureRedirect switch.'
            Test-ConnectWiseManageRedirectFailure $msg | Should -BeTrue
            Test-ConnectWiseManageRedirectFailure 'ConnectWise Manage API request failed: 400 bad request' | Should -BeFalse
        }
    }

    It 'keeps an https call on https when Manage redirects to http' {
        InModuleScope ConnectWiseManage {
            $next = Resolve-ConnectWiseManageRedirectUri -RequestUri ([Uri]'https://na.myconnectwise.net/v2025_1/apis/3.0/service/tickets/1/notes') -Location 'http://na.myconnectwise.net/v2025_1/apis/3.0/service/tickets/1/notes/9'
            $next.AbsoluteUri | Should -Be 'https://na.myconnectwise.net/v2025_1/apis/3.0/service/tickets/1/notes/9'
        }
    }

    It 'returns a created note from the response body and does not follow an http location' {
        $tcp = [System.Net.Sockets.TcpListener]::new([System.Net.IPAddress]::Loopback, 0)
        $tcp.Start()
        $port = $tcp.LocalEndpoint.Port
        $tcp.Stop()
        $listener = [System.Net.HttpListener]::new()
        $listener.Prefixes.Add("http://127.0.0.1:$port/")
        $listener.Start()
        $gate = [hashtable]::Synchronized(@{ hit = ''; stop = $false })
        $worker = [powershell]::Create()
        $worker.Runspace = [runspacefactory]::CreateRunspace()
        $worker.Runspace.Open()
        [void]$worker.AddScript({
            param($listener, $gate)
            while (-not $gate.stop) {
                $pending = $listener.GetContextAsync()
                while (-not $pending.IsCompleted) {
                    if ($gate.stop) { return }
                    Start-Sleep -Milliseconds 50
                }
                if ($gate.stop) { return }
                $ctx = $pending.Result
                $gate.hit = $ctx.Request.Url.AbsolutePath
                $json = '{"id":42,"text":"saved"}'
                $bytes = [Text.Encoding]::UTF8.GetBytes($json)
                $ctx.Response.StatusCode = 201
                $ctx.Response.Headers['Location'] = 'http://127.0.0.1/do-not-follow'
                $ctx.Response.ContentType = 'application/json'
                $ctx.Response.OutputStream.Write($bytes, 0, $bytes.Length)
                $ctx.Response.Close()
            }
        }).AddArgument($listener).AddArgument($gate)
        $handle = $worker.BeginInvoke()
        try {
            $note = Invoke-ConnectWiseManageApi -Credentials ([pscustomobject]@{
                ApiUrl = "http://127.0.0.1:$port"; CompanyId = 'co'; PublicKey = 'pub'; PrivateKey = 'priv'; ClientId = 'client'
            }) -Method POST -RelativePath 'service/tickets/9/notes' -Body @{ text = 'saved' }
            $note.id | Should -Be 42
            $gate.hit | Should -Be '/apis/3.0/service/tickets/9/notes'
        } finally {
            $gate.stop = $true
            $listener.Stop()
            $listener.Close()
            $worker.EndInvoke($handle) | Out-Null
            $worker.Runspace.Close()
            $worker.Dispose()
        }
    }

    It 'refuses time when the client email is not selected' {
        { Resolve-ConnectWiseManageHunterPost -Request @{
            ticketId = '1'
            status   = 'review'
            time     = 'remote'
            hours    = 1
            notes    = @{ client = ''; internal = 'Only internal' }
        } } | Should -Throw '*client email*'
    }
}
