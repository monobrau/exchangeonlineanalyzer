#Requires -Version 5.1
#Requires -Modules Pester

$modulePath = Join-Path (Split-Path $PSScriptRoot -Parent) 'Modules\EmailTriage.psm1'
Import-Module $modulePath -Force

Describe 'Email triage pure functions' {
    BeforeAll {
        $script:free = @(Get-TriageFreemailDomains)
    }

    It 'includes gmail.com and an extra freemail domain from notes' {
        $script:free | Should -Contain 'gmail.com'
        $more = @(Get-TriageFreemailDomains -Extra @('notes.example'))
        $more | Should -Contain 'notes.example'
    }

    It 'maps quarantine type and policy to a layer' {
        Get-TriageLayer -Type 'Phish' -PolicyType 'AntiPhishPolicy' | Should -Be 'anti-phish'
        Get-TriageLayer -Type 'Spam' -PolicyType 'HostedContentFilterPolicy' | Should -Be 'HCP'
        Get-TriageLayer -Type 'Spam' | Should -Be 'EOP anti-spam'
        Get-TriageLayer -Type 'TransportRule' -PolicyName 'Block external' | Should -Be 'transport rule'
        Get-TriageLayer -Type 'TransportRule' -PolicyName 'Bypass for Proofpoint' | Should -Be 'third-party filter'
    }

    It 'classifies Contoso exec, lure, invoice, reply, noise, spray, and rotation' {
        $vip = @('Jane Doe')
        $internal = @('contoso.com')
        $exec = Get-TriageClassification -Subject 'Jane Doe shared a file with you' -SenderAddress 'jane.doe.exec@gmail.com' -VipNames $vip -InternalDomains $internal -FreemailDomains $script:free
        $exec.Verdict | Should -Be 'Malicious'
        ($exec.Evidence -join ' ') | Should -Match 'Jane Doe'

        $lure = Get-TriageClassification -Subject 'Your password expires today' -SenderAddress 'notice@example.com' -FreemailDomains $script:free
        $lure.Verdict | Should -Be 'Malicious'

        $invoice = Get-TriageClassification -Subject 'Updated accounts payable bank details' -SenderAddress 'ap@example.com' -FreemailDomains $script:free
        $invoice.Verdict | Should -Be 'Malicious'

        $reply = Get-TriageClassification -Subject 'Re: Q3 engagement letter' -SenderAddress 'alex@client.example' -InternalDomains $internal -FreemailDomains $script:free
        $reply.Verdict | Should -Be 'Legit'

        $noise = Get-TriageClassification -Subject 'Merchant cash advance for your firm' -SenderAddress 'sales@example.com' -FreemailDomains $script:free
        $noise.Verdict | Should -Be 'Noise'

        $spray = Get-TriageClassification -Subject 'Hello team' -SenderAddress 'one@example.com' -SenderRecipientCount 8 -FreemailDomains $script:free
        $spray.Verdict | Should -Be 'Malicious'
        ($spray.Evidence -join ' ') | Should -Match 'Spray'

        $rot = Get-TriageClassification -Subject 'Invoice copy' -SenderAddress 'billing@lookalike.example' -RotationDomainCount 3 -RotationLocalPart 'billing' -FreemailDomains $script:free
        $rot.Verdict | Should -Be 'Malicious'
        ($rot.Evidence -join ' ') | Should -Match 'Domain rotation'
    }

    It 'keeps a spoofed internal exec as its own category and never drops it' {
        $match = Get-TriageImpersonationMatch -Subject 'Jane Doe wire' -SenderAddress 'ceo.spoof@contoso.com' -VipNames @('Jane Doe') -InternalDomains @('contoso.com')
        $match.Category | Should -Be 'SpoofedInternal'
        $cls = Get-TriageClassification -Subject 'Jane Doe wire' -SenderAddress 'ceo.spoof@contoso.com' -VipNames @('Jane Doe') -InternalDomains @('contoso.com') -FreemailDomains $script:free
        $cls.Verdict | Should -Be 'Malicious'
    }

    It 'keeps a quarantine Header property and drops the formatted table' {
        $obj = [pscustomobject]@{ Identity = 'qid'; Header = "Authentication-Results: spf=fail`r`nFrom: a@example.com" }
        (Get-TriageHeaderText $obj) | Should -Match 'Authentication-Results'
        $table = @"
Identity                                                                  Organization                         Header
--------                                                                  ------------                         ------
30e534f7-8e5b-48e8-3842-08df1b463f1d b2519536-04a8-466b-92f8-0ad8af6f1e1e ARC-Seal…
"@
        Get-TriageHeaderText $table | Should -Be ''
    }

    It 'parses synthetic EOP headers for example.com' {
        $raw = @"
Received: from mail.example.com
Authentication-Results: spf=fail (sender IP is 203.0.113.10)
 smtp.mailfrom=example.com; dkim=none; dmarc=fail action=none
 header.from=example.com; compauth=fail reason=001
X-Forefront-Antispam-Report: CIP:203.0.113.10;CTRY:US;LANG:en;SCL:9;SRV:;IPV:NLI;SFV:SPM;H:mail.example.com;PTR:mail.example.com;CAT:PHSH;SFS:;SFTY:9.20
X-Microsoft-Antispam: BCL:8;ARA:13230040
"@
        $parsed = ConvertFrom-TriageMessageHeader -Header $raw
        $parsed.SCL | Should -Be '9'
        $parsed.SFV | Should -Be 'SPM'
        $parsed.CAT | Should -Be 'PHSH'
        $parsed.CIP | Should -Be '203.0.113.10'
        $parsed.CTRY | Should -Be 'US'
        $parsed.IPV | Should -Be 'NLI'
        $parsed.SFTY | Should -Be '9.20'
        $parsed.Spf | Should -Be 'fail'
        $parsed.Dkim | Should -Be 'none'
        $parsed.Dmarc | Should -Be 'fail'
        $parsed.CompAuth | Should -Be 'fail'
        $parsed.CompAuthReason | Should -Be '001'
        $parsed.Bcl | Should -Be '8'

        $cls = Get-TriageClassification -Subject 'Please review' -SenderAddress 'user@example.com' -ParsedHeader $parsed -FreemailDomains $script:free
        $cls.Verdict | Should -Be 'Malicious'
        ($cls.Evidence -join ' ') | Should -Match 'CAT: PHSH'
        ($cls.Evidence -join ' ') | Should -Match 'compauth=fail reason 001'
        ($cls.Evidence -join ' ') | Should -Match '203.0.113.10'
    }

    It 'does not treat a passed Contoso authentication header as malicious' {
        $raw = @"
Authentication-Results: spf=pass (sender IP is 203.0.113.20) smtp.mailfrom=contoso.com; dkim=pass header.d=contoso.com; dmarc=pass header.from=contoso.com; compauth=pass reason=100
X-Forefront-Antispam-Report: CIP:203.0.113.20;CTRY:US;SCL:1;IPV:NLI;SFV:NSPM
X-Microsoft-Antispam: BCL:0
"@
        $parsed = ConvertFrom-TriageMessageHeader -Header $raw
        $parsed.Spf | Should -Be 'pass'
        $parsed.Dmarc | Should -Be 'pass'
        $cls = Get-TriageClassification -Subject 'Lunch Thursday' -SenderAddress 'alex@contoso.com' -ParsedHeader $parsed -FreemailDomains $script:free
        $cls.Verdict | Should -Be 'Unknown'
    }

    It 'merges a delivered then quarantined leg and a lone delivery' {
        $t0 = [datetime]::SpecifyKind([datetime]'2026-09-01T15:00:00', 'Utc')
        $filtered = @(Merge-TriageTraceLegs -WindowSeconds 120 -Rows @(
                [pscustomobject]@{ MessageId = '<m1@contoso.com>'; RecipientAddress = 'user@contoso.com'; SenderAddress = 'a@example.com'; Subject = 'loop'; Status = 'Delivered'; Received = $t0 },
                [pscustomobject]@{ MessageId = '<m1@contoso.com>'; RecipientAddress = 'user@contoso.com'; SenderAddress = 'a@example.com'; Subject = 'loop'; Status = 'Quarantined'; Received = $t0.AddSeconds(30) }
            ))
        $filtered.Count | Should -Be 1
        $filtered[0].ReachedMailbox | Should -Be $false
        $filtered[0].Outcome | Should -Be 'Filtered'

        $lone = @(Merge-TriageTraceLegs -WindowSeconds 120 -Rows @(
                [pscustomobject]@{ MessageId = '<m2@contoso.com>'; RecipientAddress = 'user@contoso.com'; SenderAddress = 'a@example.com'; Subject = 'ok'; Status = 'Delivered'; Received = $t0 }
            ))
        $lone.Count | Should -Be 1
        $lone[0].ReachedMailbox | Should -Be $true

        $late = @(Merge-TriageTraceLegs -WindowSeconds 120 -Rows @(
                [pscustomobject]@{ MessageId = '<m3@contoso.com>'; RecipientAddress = 'user@contoso.com'; SenderAddress = 'a@example.com'; Subject = 'late'; Status = 'Delivered'; Received = $t0 },
                [pscustomobject]@{ MessageId = '<m3@contoso.com>'; RecipientAddress = 'user@contoso.com'; SenderAddress = 'a@example.com'; Subject = 'late'; Status = 'Quarantined'; Received = $t0.AddSeconds(180) }
            ))
        $late[0].ReachedMailbox | Should -Be $true
    }

    It 'merges trace legs held in a List[object]' {
        $t0 = [datetime]::SpecifyKind([datetime]'2026-09-01T15:00:00', 'Utc')
        $rows = New-Object System.Collections.Generic.List[object]
        $rows.Add([pscustomobject]@{ MessageId = '<m-list@contoso.com>'; RecipientAddress = 'user@contoso.com'; SenderAddress = 'a@example.com'; Subject = 'list wrap'; Status = 'Delivered'; Received = $t0 })
        $legs = @(Merge-TriageTraceLegs -WindowSeconds 120 -Rows $rows)
        $legs.Count | Should -Be 1
        $legs[0].ReachedMailbox | Should -Be $true
    }

    It 'folds Expanded and Resolved into the final recipient' {
        $t0 = [datetime]::SpecifyKind([datetime]'2026-09-01T15:00:00', 'Utc')
        $legs = @(Merge-TriageTraceLegs -WindowSeconds 120 -Rows @(
                [pscustomobject]@{ MessageId = '<m4@contoso.com>'; RecipientAddress = 'dl@contoso.com'; SenderAddress = 'a@example.com'; Subject = 'list'; Status = 'Expanded'; Received = $t0 },
                [pscustomobject]@{ MessageId = '<m4@contoso.com>'; RecipientAddress = 'alice@contoso.com'; SenderAddress = 'a@example.com'; Subject = 'list'; Status = 'Resolved'; Received = $t0.AddSeconds(2) },
                [pscustomobject]@{ MessageId = '<m4@contoso.com>'; RecipientAddress = 'alice@contoso.com'; SenderAddress = 'a@example.com'; Subject = 'list'; Status = 'Delivered'; Received = $t0.AddSeconds(5) }
            ))
        $legs.Count | Should -Be 1
        $legs[0].RecipientAddress | Should -Be 'alice@contoso.com'
        $legs[0].ReachedMailbox | Should -Be $true
        ($legs[0].Statuses -join ',') | Should -Match 'Expanded'
        ($legs[0].Statuses -join ',') | Should -Match 'Resolved'
    }

    It 'converts an unspecified local time to UTC with a fixed offset' {
        $tz = [TimeZoneInfo]::CreateCustomTimeZone('TestFixed', [TimeSpan]::FromHours(-5), 'Test', 'Test')
        $local = [datetime]::SpecifyKind([datetime]'2026-01-15T12:00:00', 'Unspecified')
        $utc = ConvertTo-TriageUtc -Time $local -TimeZone $tz
        $utc.Hour | Should -Be 17
        $utc.Kind | Should -Be ([DateTimeKind]::Utc)
        $already = [datetime]::SpecifyKind([datetime]'2026-01-15T12:00:00', 'Utc')
        (ConvertTo-TriageUtc -Time $already -TimeZone $tz).Hour | Should -Be 12
    }

    It 'keeps only the custom config checks and reads DMARC from DomainHealthChecker output' {
        $config = @{
            TransportRules = @(@{ Name = 'Allow consumer'; ExceptIfSenderDomainIs = @('gmail.com') })
            TablItems = @(@{ Value = 'yahoo.com'; Action = 'Allow' })
            AntiPhishPolicies = @(@{ Name = 'Office365 AntiPhish Default'; EnableTargetedUserProtection = $false; TargetedUsersToProtect = @() })
            InboundConnectors = @(@{ Name = 'INKY inbound'; EFSkipIPs = @(); EFSkipLastIP = $false })
            Mx = @(@{ Domain = 'contoso.com'; Exchanges = @('mx1.inky.example') })
            DomainHealth = @(@{ Name = 'contoso.com'; DmarcRecord = 'v=DMARC1; p=none' })
            Dns = @(@{ Domain = 'contoso.com'; Ns = @('ns1.cloudflare.com'); Soa = 'ns1.cloudflare.com' })
        }
        $found = @(Test-TriageConfigFindings -Config $config -FreemailDomains $script:free)
        $ids = @($found | ForEach-Object { $_.Check })
        $ids | Should -Contain 'TransportRuleConsumerException'
        $ids | Should -Contain 'TablFreemailAllow'
        $ids | Should -Contain 'InboundConnectorEnhancedFiltering'
        $ids | Should -Contain 'AntiPhishNoProtectedUsers'
        $ids | Should -Contain 'DmarcPolicyNone'
        $ids | Should -Contain 'DmarcMissingRua'
        $ids | Should -Contain 'DnsHost'
        $ids | Should -Not -Contain 'SafeLinksDefender'
        $ids | Should -Not -Contain 'MailboxIntelligence'
        @($found | Where-Object { $_.Check -eq 'AntiPhishNoProtectedUsers' }).Count | Should -Be 1

        $eop = @{
            InboundConnectors = @(@{ Name = 'Proofpoint inbound'; EFSkipIPs = @(); EFSkipLastIP = $false })
            Mx = @(@{ Domain = 'contoso.com'; Exchanges = @('contoso-com.mail.protection.outlook.com') })
            DomainHealth = @(@{ Name = 'contoso.com'; DmarcRecord = 'v=DMARC1; p=reject; rua=mailto:dmarc@contoso.com' })
        }
        $eopIds = @($(Test-TriageConfigFindings -Config $eop -FreemailDomains $script:free) | ForEach-Object { $_.Check })
        $eopIds | Should -Not -Contain 'InboundConnectorEnhancedFiltering'
        $eopIds | Should -Not -Contain 'DmarcPolicyNone'
        $eopIds | Should -Not -Contain 'DmarcMissingRua'
        (Get-TriageDnsHost -Ns @('ns1.cloudflare.com') -Soa 'ns1.cloudflare.com') | Should -Be 'Cloudflare'
        (Test-TriageMxPointsAtEop -Exchanges @('contoso-com.mail.protection.outlook.com')) | Should -Be $true
    }

    It 'detects Defender for Office 365 from an MDO-only ORCA result' {
        $orca = @{
            ResultDate = '2026-09-25T00:00:00'
            Tenant = 'contoso'
            Results = @(
                @{ Name = 'Safe Links'; Result = 3; Services = 2; ChiValue = 15; Area = 'Microsoft Defender for Office 365 Policies'; FailRecommendation = 'Enable Safe Links' },
                @{ Name = 'Spam action'; Result = 1; Services = 1; ChiValue = 10; Area = 'EOP' }
            )
        }
        $sum = Get-TriageOrcaSummary -Orca $orca
        $sum.DefenderForOffice365 | Should -Be $true
        $sum.ServiceLabel | Should -Be 'Defender for Office 365'
        $sum.FailCount | Should -Be 1
        $sum.Groups[0].Severity | Should -Be 'High'
        $sum.Groups[0].Items[0].Name | Should -Be 'Safe Links'

        $eopOnly = Get-TriageOrcaSummary -Orca @{ Results = @(@{ Name = 'Spam'; Result = 'Fail'; Services = 'EOP'; ChiValue = 'Low' }) }
        $eopOnly.DefenderForOffice365 | Should -Be $false
        $eopOnly.ServiceLabel | Should -Match 'not detected'
    }

    It 'comments every action command and blocks freemail per address' {
        $address = Get-TriageBlockTarget -SenderAddress 'badguy@gmail.com' -FreemailDomains $script:free
        $domain = Get-TriageBlockTarget -SenderAddress 'attacker@evil.example' -FreemailDomains $script:free
        $address.Scope | Should -Be 'Address'
        $domain.Scope | Should -Be 'Domain'
        $script = New-TriageActionScript -ReleaseIdentities @('qid-1') -Blocks @($address, $domain) -ProtectUsers @('ceo@contoso.com') -SoftDeleteMessageIds @('<m@evil.example>')
        $bare = @($script -split "`r?`n" | Where-Object { $_ -match '^\s*[A-Za-z]+-[A-Za-z]' })
        $bare.Count | Should -Be 0
        $script | Should -Match "# Release-QuarantineMessage -Identity 'qid-1'"
        $script | Should -Match 'badguy@gmail.com'
        $script | Should -Match 'evil\.example'
        $script | Should -Not -Match "Entries 'gmail.com'"
        $script | Should -Match 'Connect-IPPSSession -EnableSearchOnlySession'
        $script | Should -Match 'Explorer'
    }

    It 'writes a client email without filter jargon' {
        $rows = @([pscustomobject]@{ Verdict = 'Malicious'; SenderAddress = 'a@gmail.com'; Subject = 'DocuSign'; Evidence = @('lure') })
        $mail = New-TriageClientEmail -Rows $rows -Company 'Contoso' -Contacts @('it@contoso.com')
        $mail | Should -Match 'phone call'
        $mail | Should -Match 'password'
        $mail | Should -Not -Match 'EFSkip|TABL|DMARC|HostedContentFilter|AntiPhish'
        $note = New-TriageTicketNote -Rows $rows -Findings @([pscustomobject]@{ Severity = 'High'; Detail = 'example finding' }) -Company 'Contoso'
        $note | Should -Match 'Malicious: 1'
        $note | Should -Match 'example finding'
    }

    It 'fails config commands when ORCA or DomainHealthChecker is missing' {
        $both = Test-TriageConfigModuleInstalled -InstalledNames @()
        $both | Should -Match '^TRIAGE_FAILED:'
        $both | Should -Match 'Install-Module ORCA,DomainHealthChecker -Scope CurrentUser'
        $one = Test-TriageConfigModuleInstalled -InstalledNames @('ORCA')
        $one | Should -Match 'DomainHealthChecker'
        $one | Should -Match 'Install-Module ORCA,DomainHealthChecker -Scope CurrentUser'
        Test-TriageConfigModuleInstalled -InstalledNames @('ORCA', 'DomainHealthChecker') | Should -Be $null
        Invoke-TriageWorkerCommand -Command 'TRIAGE_NOPE' | Should -Be 'TRIAGE_FAILED:Unknown triage command'
    }

    It 'reviews quarantine rows and delivered lures, and skips ordinary trace mail' {
        $t0 = [datetime]::SpecifyKind([datetime]'2026-09-01T15:00:00', 'Utc')
        $rows = @(ConvertTo-TriageRows -FreemailDomains $script:free -Quarantine @(
                [pscustomobject]@{ SenderAddress = 'bad@gmail.com'; RecipientAddress = 'alex@contoso.com'; Subject = 'Hello'; MessageId = '<q@contoso.com>'; Identity = 'qid-1'; ReceivedTime = '2026-09-01T15:00:00Z'; Type = 'Phish' }
            ) -Legs @(
                [pscustomobject]@{ SenderAddress = 'bad@gmail.com'; RecipientAddress = 'alex@contoso.com'; Subject = 'Hello'; MessageId = '<q@contoso.com>'; Received = $t0; ReachedMailbox = $true; Outcome = 'ReachedMailbox' },
                [pscustomobject]@{ SenderAddress = 'noreply@news.example'; RecipientAddress = 'alex@contoso.com'; Subject = 'Weekly newsletter'; MessageId = '<n@news.example>'; Received = $t0; ReachedMailbox = $true; Outcome = 'ReachedMailbox' },
                [pscustomobject]@{ SenderAddress = 'billing@evil.example'; RecipientAddress = 'ap@contoso.com'; Subject = 'Updated banking details'; MessageId = '<b@evil.example>'; Received = $t0; ReachedMailbox = $true; Outcome = 'ReachedMailbox' }
            ))
        $rows.Count | Should -Be 2
        $held = @($rows | Where-Object { $_.Source -eq 'Quarantine' })
        $held.Count | Should -Be 1
        $held[0].ReachedMailbox | Should -Be $true
        ($held[0].Evidence -join ' ') | Should -Match 'reached a mailbox'
        @($rows | Where-Object { $_.MessageId -eq '<b@evil.example>' }).Count | Should -Be 1
        @($rows | Where-Object { $_.MessageId -eq '<n@news.example>' }).Count | Should -Be 0
    }

    It 'marks a quarantined message that was delivered before the hold' {
        $t0 = [datetime]::SpecifyKind([datetime]'2026-09-01T15:00:00', 'Utc')
        $rows = @(ConvertTo-TriageRows -FreemailDomains $script:free -Quarantine @(
                [pscustomobject]@{ SenderAddress = 'bad@gmail.com'; RecipientAddress = 'alex@contoso.com'; Subject = 'Invoice'; MessageId = '<held@x>'; Identity = 'qid-2'; ReceivedTime = '2026-09-01T15:00:24Z'; Type = 'Phish' }
            ) -Legs @(
                [pscustomobject]@{ SenderAddress = 'bad@gmail.com'; RecipientAddress = 'alex@contoso.com'; Subject = 'Invoice'; MessageId = '<held@x>'; Received = $t0; ReachedMailbox = $false; Outcome = 'Filtered'; Statuses = @('Delivered', 'Quarantined') }
            ))
        $rows.Count | Should -Be 1
        $rows[0].ReachedMailbox | Should -Be $true
        ($rows[0].Evidence -join ' ') | Should -Match 'reached a mailbox'
    }

    It 'saves an override without changing the original quarantine CSV' {
        $folder = Join-Path $TestDrive 'Triage_fixture'
        New-Item -ItemType Directory -Path $folder | Out-Null
        $csv = Join-Path $folder 'Quarantine.csv'
        @(
            'ReceivedTime,SenderAddress,RecipientAddress,Subject,Type,PolicyType,PolicyName,MessageId,Identity,Direction'
            '2026-09-25T12:00:00,bad@gmail.com,alex@contoso.com,Jane Doe password expires,Phish,AntiPhishPolicy,Default,<a@evil.example>,qid-1,Inbound'
        ) | Set-Content -LiteralPath $csv -Encoding UTF8
        '{}' | Set-Content -LiteralPath (Join-Path $folder 'Config.json') -Encoding UTF8
        $before = (Get-FileHash -LiteralPath $csv -Algorithm SHA256).Hash
        $view = Get-TriagePackView -Folder $folder -VipNames @('Jane Doe') -Company 'Contoso'
        @($view.Rows).Count | Should -Be 1
        $view.Rows[0].Verdict | Should -Be 'Malicious'
        $key = $view.Rows[0].RowKey
        Set-TriageVerdictOverride -Folder $folder -RowKey $key -Verdict 'Legit'
        $after = (Get-FileHash -LiteralPath $csv -Algorithm SHA256).Hash
        $after | Should -Be $before
        $again = Get-TriagePackView -Folder $folder -VipNames @('Jane Doe') -Company 'Contoso'
        $again.Rows[0].Verdict | Should -Be 'Legit'
        $again.Rows[0].SuggestedVerdict | Should -Be 'Malicious'
        $again.ClientEmail | Should -Not -Match 'EFSkip|TABL'
        Test-Path -LiteralPath (Join-Path $folder 'actions.ps1') | Should -Be $true
        Test-Path -LiteralPath (Join-Path $folder 'Report.html') | Should -Be $true
        $html = Get-Content -LiteralPath (Join-Path $folder 'Report.html') -Raw
        $html | Should -Match '<html'
    }

    It 'builds block, allow, and license-gated config actions' {
        $rows = @(
            [pscustomobject]@{ Source = 'Trace'; SenderAddress = 'bad@gmail.com'; RecipientAddress = 'alex@contoso.com'; Subject = 'Password expires tonight'; Verdict = 'Malicious'; ReachedMailbox = $true; Evidence = 'Credential-lure wording in the subject'; Impersonation = ''; Identity = '' },
            [pscustomobject]@{ Source = 'Trace'; SenderAddress = 'donotreply@vonage.com'; RecipientAddress = 'alex@contoso.com'; Subject = 'New Voicemail from LUIS'; Verdict = 'Malicious'; ReachedMailbox = $true; Evidence = 'Credential-lure wording in the subject'; Impersonation = ''; Identity = '' },
            [pscustomobject]@{ Source = 'Quarantine'; SenderAddress = 'ap@vendor.example'; RecipientAddress = 'alex@contoso.com'; Subject = 'Re: March invoice'; Verdict = 'Legit'; ReachedMailbox = $false; Evidence = 'Reply subject from a non-freemail sender'; Impersonation = ''; Identity = 'qid-9' }
        )
        $orca = @{
            DefenderForOffice365 = $false
            Groups = @(@{ Items = @(
                @{ Name = 'Outbound spam filter policy settings'; Result = 'Fail' },
                @{ Name = 'Safe Links Policy Rules'; Result = 'Fail' },
                @{ Name = 'Signing Configuration'; Result = 'Fail' }
            ) })
        }
        $config = @{
            TransportRules = @(@{ Name = 'Skip Gmail'; ExceptIfSenderDomainIs = @('gmail.com') })
            AntiPhishPolicies = @(@{ Name = 'Office365 AntiPhish Default'; TargetedUsersToProtect = @() })
            DomainHealth = @(
                @{ Name = 'contoso.com'; DmarcRecord = 'v=DMARC1; p=none; rua=mailto:dmarc@contoso.com' },
                @{ Name = 'contoso.onmicrosoft.com'; DmarcRecord = '' }
            )
        }
        $items = @(Get-TriageRemediationItems -Rows $rows -Orca $orca -Config $config -FreemailDomains $script:free -VipNames @('Jane Doe'))
        $block = @($items | Where-Object { $_.Kind -eq 'Block' })
        $block.Count | Should -Be 1
        $block[0].Title | Should -Be 'bad@gmail.com'
        $block[0].Apply | Should -Be 'Worker'
        $block[0].Script | Should -Match "New-TenantAllowBlockListItems -ListType Sender -Block -Entries 'bad@gmail.com'"
        @($items | Where-Object { $_.Title -eq 'donotreply@vonage.com' }).Count | Should -Be 0
        $allow = @($items | Where-Object { $_.Kind -eq 'Allow' })
        $allow.Count | Should -Be 1
        $allow[0].Script | Should -Match 'Release-QuarantineMessage'
        $allow[0].Script | Should -Match 'qid-9'
        $spam = @($items | Where-Object { $_.Id -eq 'config:outbound-spam' })
        $spam.Count | Should -Be 1
        $spam[0].Apply | Should -Be 'Worker'
        $spam[0].Script | Should -Match 'Set-HostedOutboundSpamFilterPolicy'
        $safe = @($items | Where-Object { $_.Id -eq 'config:safe-links-policy' })
        $safe[0].Apply | Should -Be 'Manual'
        $safe[0].Detail | Should -Match 'Defender for Office 365'
        $rule = @($items | Where-Object { $_.Operation -eq 'RemoveRuleException' })
        $rule[0].Script | Should -Match 'gmail.com'
        @($items | Where-Object { $_.Id -eq 'config:dmarc:contoso.com' }).Count | Should -Be 1
        @($items | Where-Object { $_.Id -like 'config:dmarc:contoso.onmicrosoft.com' }).Count | Should -Be 0
        @($items | Where-Object { $_.Id -eq 'config:protected-users' })[0].Detail | Should -Match 'Jane Doe'
        Invoke-TriageWorkerCommand -Command 'TRIAGE_APPLY' | Should -Be 'TRIAGE_FAILED:Missing folder'
    }
}

Describe 'Email triage AI review' {
    It 'builds a brief from verdicts and findings and leaves out headers and scripts' {
        $view = [pscustomobject]@{
            Warning = 'Trace was limited.'
            Rows    = @(
                [pscustomobject]@{
                    Verdict = 'Malicious'; SuggestedVerdict = 'Malicious'
                    SenderAddress = 'jane.doe.exec@gmail.com'; RecipientAddress = 'alex@contoso.com'
                    Subject = 'Jane Doe password expires'
                    Evidence = @('Exec name in subject from a freemail sender')
                    ReachedMailbox = $false; Outcome = 'Quarantine'; Impersonation = 'External'
                    RawHeader = 'Authentication-Results: secret-header'
                }
            )
            Findings = @([pscustomobject]@{ Severity = 'High'; Detail = 'DMARC for contoso.com is p=none.' })
            Orca     = [pscustomobject]@{
                ServiceLabel = 'Defender for Office 365'
                Groups       = @([pscustomobject]@{
                    Items = @([pscustomobject]@{ Name = 'Safe Links'; Area = 'Defender'; Result = 'Fail'; FailRecommendation = 'Enable Safe Links.' })
                })
            }
            Actions  = @([pscustomobject]@{ Kind = 'Block'; Title = 'bad@gmail.com'; Detail = 'Credential lure.'; Apply = 'Worker'; Script = "New-TenantAllowBlockListItems -ListType Sender -Block -Entries 'bad@gmail.com'" })
        }
        $brief = New-TriageAiReviewBrief -View $view
        $brief | Should -Match 'jane.doe.exec@gmail.com'
        $brief | Should -Match 'Jane Doe password expires'
        $brief | Should -Match 'DMARC for contoso.com is p=none'
        $brief | Should -Match 'Enable Safe Links'
        $brief | Should -Match 'bad@gmail.com'
        $brief | Should -Not -Match 'secret-header'
        $brief | Should -Not -Match 'New-TenantAllowBlockListItems'
    }

    It 'leaves unchecked and non-interesting messages out of the AI brief' {
        $view = [pscustomobject]@{
            Rows = @(
                [pscustomobject]@{ Verdict = 'Malicious'; SenderAddress = 'bad@gmail.com'; Subject = 'wire'; Evidence = 'lure' }
                [pscustomobject]@{ Verdict = 'Noise'; SenderAddress = 'news@letters.example'; Subject = 'Weekly digest'; Evidence = 'newsletter' }
                [pscustomobject]@{ Verdict = 'Legit'; SenderAddress = 'boss@contoso.com'; Subject = 'status'; Evidence = 'internal'; AiReview = $true }
                [pscustomobject]@{ Verdict = 'Malicious'; SenderAddress = 'skip@gmail.com'; Subject = 'skip me'; Evidence = 'lure'; AiReview = $false }
                [pscustomobject]@{ Verdict = 'Unknown'; SenderAddress = 'maybe@gmail.com'; Subject = 'question'; Evidence = 'unclear' }
            )
        }
        $brief = New-TriageAiReviewBrief -View $view
        $brief | Should -Match 'bad@gmail.com'
        $brief | Should -Match 'boss@contoso.com'
        $brief | Should -Match 'maybe@gmail.com'
        $brief | Should -Not -Match 'news@letters.example'
        $brief | Should -Not -Match 'skip@gmail.com'
        $brief | Should -Match '2 messages were left out of this review'
    }

    It 'leaves repeated non-malicious mail out of the AI brief' {
        $rows = @(
            [pscustomobject]@{ RowKey = 'ship1'; Verdict = 'Unknown'; SenderAddress = 'notices@shop.example'; Subject = 'Your order has shipped'; Evidence = 'notice' }
            [pscustomobject]@{ RowKey = 'ship2'; Verdict = 'Unknown'; SenderAddress = 'notices@shop.example'; Subject = 'Your order has shipped'; Evidence = 'notice' }
            [pscustomobject]@{ RowKey = 'ship3'; Verdict = 'Unknown'; SenderAddress = 'notices@shop.example'; Subject = 'Your order has shipped'; Evidence = 'notice' }
            [pscustomobject]@{ RowKey = 'quote'; Verdict = 'Unknown'; SenderAddress = 'notices@shop.example'; Subject = 'Quote for the dock repair'; Evidence = 'one off' }
            [pscustomobject]@{ RowKey = 'phish'; Verdict = 'Malicious'; SenderAddress = 'notices@shop.example'; Subject = 'Password expires'; Evidence = 'lure' }
            [pscustomobject]@{ RowKey = 'ship4'; Verdict = 'Unknown'; SenderAddress = 'notices@shop.example'; Subject = 'Your order has shipped'; Evidence = 'notice'; AiReview = $true }
        )
        1..5 | ForEach-Object {
            $rows += [pscustomobject]@{ RowKey = "news$_"; Verdict = 'Unknown'; SenderAddress = 'digest@news.example'; Subject = "Headline $_"; Evidence = 'digest' }
        }
        $brief = New-TriageAiReviewBrief -View ([pscustomobject]@{ Rows = $rows })
        $brief | Should -Match 'Quote for the dock repair'
        $brief | Should -Match 'Password expires'
        $brief | Should -Match 'Your order has shipped'
        $brief | Should -Not -Match 'Headline 1'
        $brief | Should -Match 'Repeated mail left out because it was not marked malicious'
        $brief | Should -Match 'digest@news.example \(5\)'
        $brief | Should -Match 'notices@shop.example \(3\)'
    }

    It 'keeps a verdict when a message is left out of AI review' {
        $folder = Join-Path $TestDrive 'ai-pack'
        New-Item -ItemType Directory -Path $folder | Out-Null
        @(
            [pscustomobject]@{ RowKey = 'noise1'; Source = 'Quarantine'; ReceivedUtc = ''; SenderAddress = 'news@letters.example'; RecipientAddress = 'a@contoso.com'; Subject = 'digest'; MessageId = ''; Identity = ''; Layer = ''; SuggestedVerdict = 'Noise'; Verdict = 'Noise'; Evidence = 'newsletter'; ReachedMailbox = 'False'; Outcome = ''; Impersonation = '' }
            [pscustomobject]@{ RowKey = 'bad1'; Source = 'Quarantine'; ReceivedUtc = ''; SenderAddress = 'bad@gmail.com'; RecipientAddress = 'a@contoso.com'; Subject = 'wire'; MessageId = ''; Identity = ''; Layer = ''; SuggestedVerdict = 'Malicious'; Verdict = 'Malicious'; Evidence = 'lure'; ReachedMailbox = 'False'; Outcome = ''; Impersonation = '' }
            [pscustomobject]@{ RowKey = 'ok1'; Source = 'Quarantine'; ReceivedUtc = ''; SenderAddress = 'boss@contoso.com'; RecipientAddress = 'a@contoso.com'; Subject = 'status'; MessageId = ''; Identity = ''; Layer = ''; SuggestedVerdict = 'Legit'; Verdict = 'Legit'; Evidence = 'internal'; ReachedMailbox = 'False'; Outcome = ''; Impersonation = '' }
        ) | Export-Csv -LiteralPath (Join-Path $folder 'Triage.csv') -NoTypeInformation -Encoding UTF8
        Set-TriageVerdictOverride -Folder $folder -RowKey 'bad1' -Verdict 'Legit'
        Set-TriageVerdictOverride -Folder $folder -RowKey 'bad1|ok1' -Verdict 'Omit'
        Set-TriageVerdictOverride -Folder $folder -RowKey 'ok1' -Verdict 'Include'
        $doc = Get-Content -LiteralPath (Join-Path $folder 'overrides.json') -Raw | ConvertFrom-Json
        @($doc.items | Where-Object { $_.rowKey -eq 'bad1' }).verdict | Should -Be 'Legit'
        @($doc.review | Where-Object { $_.rowKey -eq 'bad1' }).include | Should -BeFalse
        @($doc.review | Where-Object { $_.rowKey -eq 'ok1' }).include | Should -BeTrue
        $saved = Get-TriagePackView -Folder $folder -Limit 20
        @($saved.Rows | Where-Object { $_.RowKey -eq 'noise1' }).AiReview | Should -BeFalse
        @($saved.Rows | Where-Object { $_.RowKey -eq 'bad1' }).AiReview | Should -BeFalse
        @($saved.Rows | Where-Object { $_.RowKey -eq 'bad1' }).Verdict | Should -Be 'Legit'
        @($saved.Rows | Where-Object { $_.RowKey -eq 'ok1' }).AiReview | Should -BeTrue
        $brief = New-TriageAiReviewBrief -View $saved
        $brief | Should -Match 'boss@contoso.com'
        $brief | Should -Not -Match 'bad@gmail.com'
        $brief | Should -Not -Match 'news@letters.example'
    }

    It 'stores the key on disk and does not return it' {
        $path = Join-Path $TestDrive 'ai-review.settings.json'
        $saved = Set-TriageAiSettings -Path $path -Provider 'Claude' -Model '' -ApiKey 'sk-ant-testkey1234'
        $saved.provider | Should -Be 'claude'
        $saved.model | Should -Be 'claude-sonnet-5'
        $saved.hasKey | Should -BeTrue
        $saved.keyHint | Should -Be '····1234'
        $saved.PSObject.Properties.Name | Should -Not -Contain 'apiKey'
        $raw = Get-Content -LiteralPath $path -Raw
        $raw | Should -Match 'sk-ant-testkey1234'
        $kept = Set-TriageAiSettings -Path $path -Provider 'chatgpt' -Model 'gpt-4.1' -ApiKey ''
        $kept.provider | Should -Be 'chatgpt'
        $kept.model | Should -Be 'gpt-4.1'
        $kept.hasKey | Should -BeTrue
        @($kept.models | ForEach-Object { $_.name }) | Should -Contain 'gpt-4.1'
        (Get-Content -LiteralPath $path -Raw) | Should -Match 'sk-ant-testkey1234'
        { Set-TriageAiSettings -Path $path -Provider 'hatz' -Model 'gpt-4o' -ApiKey 'not a real token' } | Should -Throw -ExpectedMessage 'Paste the API token only. The value has spaces in it, so it is not a key.'
        $public = Get-TriageAiSettings -Path $path
        ($public | ConvertTo-Json -Compress) | Should -Not -Match 'sk-ant-testkey1234'
    }

    It 'builds ChatGPT and Claude requests without putting the key in the body' {
        $chat = Get-TriageAiReviewRequest -Provider 'chatgpt' -Model 'gpt-4.1-mini' -ApiKey 'sk-testkey9876' -Brief 'Pack facts'
        $chat.Uri | Should -Be 'https://api.openai.com/v1/chat/completions'
        $chat.Headers.Authorization | Should -Be 'Bearer sk-testkey9876'
        $chat.Body | Should -Match 'Pack facts'
        $chat.Body | Should -Match 'not Markdown'
        $chat.Body | Should -Not -Match 'sk-testkey9876'
        $claude = Get-TriageAiReviewRequest -Provider 'claude' -Model 'claude-sonnet-5' -ApiKey 'sk-ant-testkey9876' -Brief 'Pack facts'
        $claude.Uri | Should -Be 'https://api.anthropic.com/v1/messages'
        $claude.Headers['x-api-key'] | Should -Be 'sk-ant-testkey9876'
        $claude.Headers['anthropic-version'] | Should -Be '2023-06-01'
        $claude.Body | Should -Match 'claude-sonnet-5'
        $claude.Body | Should -Not -Match 'sk-ant-testkey9876'
        $hatz = Get-TriageAiReviewRequest -Provider 'hatz' -Model 'gpt-4o' -ApiKey 'hatz-testkey9876' -Brief 'Pack facts'
        $hatz.Uri | Should -Be 'https://ai.hatz.ai/v1/chat/completions'
        $hatz.Headers['X-API-Key'] | Should -Be 'hatz-testkey9876'
        $hatz.Body | Should -Match 'gpt-4o'
        $hatz.Body | Should -Not -Match 'hatz-testkey9876'
        $hatz.Headers.Authorization | Should -BeNullOrEmpty
        { Get-TriageAiReviewRequest -Provider 'off' -Model 'x' -ApiKey 'sk-testkey9876' -Brief 'Pack facts' } | Should -Throw
    }

    It 'reads the model text from either provider response' {
        $chat = [pscustomobject]@{ choices = @([pscustomobject]@{ message = [pscustomobject]@{ content = 'Agree.' } }) }
        (ConvertFrom-TriageAiResponse -Response $chat -Provider 'chatgpt') | Should -Be 'Agree.'
        $claude = [pscustomobject]@{ content = @([pscustomobject]@{ type = 'text'; text = 'Release the reply.' }) }
        (ConvertFrom-TriageAiResponse -Response $claude -Provider 'claude') | Should -Be 'Release the reply.'
    }

    It 'turns a Markdown review into plaintext' {
        $raw = @"
# What to fix

**Block** ``jane.doe.exec@gmail.com``.

See [DMARC](https://example.com).

* Release the reply
"@
        $plain = ConvertTo-TriageAiPlainText -Text $raw
        $plain | Should -Match 'What to fix'
        $plain | Should -Match 'Block jane.doe.exec@gmail.com'
        $plain | Should -Match 'See DMARC'
        $plain | Should -Match '- Release the reply'
        $plain | Should -Not -Match '\*\*'
        $plain | Should -Not -Match '#'
        $plain | Should -Not -Match 'https://example.com'
        $plain | Should -Not -Match '`'
    }
}

Describe 'Email triage BEC review' {
    It 'offers cleanup only when rules, forwarding, or delegates look like persistence' {
        $quiet = Get-TriageBecRemediations -Facts ([pscustomobject]@{
            mailbox = 'alex@contoso.com'
            rules = @([pscustomobject]@{ name = 'Move newsletters'; hidden = $false })
            forwards = @()
            delegates = @()
        })
        @($quiet).Count | Should -Be 0
        $facts = [pscustomobject]@{
            mailbox = 'alex@contoso.com'
            rules = @(
                [pscustomobject]@{ name = '.;'; hidden = $true; redirectTo = 'attacker@evil.example'; deleteMessage = $true }
                [pscustomobject]@{ name = 'Move newsletters'; hidden = $false }
            )
            forwards = @([pscustomobject]@{ address = 'attacker@evil.example'; field = 'Smtp' })
            delegates = @([pscustomobject]@{ user = 'helper@contoso.com'; right = 'FullAccess' })
        }
        $items = @(Get-TriageBecRemediations -Facts $facts)
        @($items | ForEach-Object { $_.Id }) | Should -Contain 'delete-rules'
        @($items | ForEach-Object { $_.Id }) | Should -Contain 'remove-forward'
        @($items | ForEach-Object { $_.Id }) | Should -Contain 'remove-delegates'
        @($items | ForEach-Object { $_.Id }) | Should -Contain 'enable-hold'
        @($items | ForEach-Object { $_.Id }) | Should -Contain 'reset-password'
        @($items | ForEach-Object { $_.Id }) | Should -Contain 'revoke'
        ($items | Where-Object { $_.Id -eq 'block' }).Checked | Should -BeFalse
        ($items | Where-Object { $_.Id -eq 'reset-password' }).Checked | Should -BeFalse
        ($items | Where-Object { $_.Id -eq 'revoke' }).Checked | Should -BeFalse
        ($items | Where-Object { $_.Id -eq 'delete-rules' }).Checked | Should -BeTrue
        $delegatesOnly = @(Get-TriageBecRemediations -Facts ([pscustomobject]@{
            mailbox = 'alex@contoso.com'
            rules = @()
            forwards = @()
            delegates = @([pscustomobject]@{ user = 'helper@contoso.com'; right = 'FullAccess' })
        }))
        @($delegatesOnly | ForEach-Object { $_.Id }) | Should -Contain 'remove-delegates'
        @($delegatesOnly | ForEach-Object { $_.Id }) | Should -Not -Contain 'reset-password'
        @($delegatesOnly | ForEach-Object { $_.Id }) | Should -Not -Contain 'revoke'
        $brief = New-TriageBecReviewBrief -Facts $facts
        $brief | Should -Match 'attacker@evil.example'
        $brief | Should -Match 'helper@contoso.com'
        $brief | Should -Not -Match 'Remove-InboxRule'
        $brief | Should -Not -Match 'Ticket'
        $withTicket = New-TriageBecReviewBrief -Facts ([pscustomobject]@{
            mailbox = 'alex@contoso.com'
            rules = @()
            forwards = @()
            delegates = @()
            ticketNumber = '1001'
            ticket = 'User reported a suspicious forward this morning.'
        })
        $withTicket | Should -Match 'Ticket number: 1001'
        $withTicket | Should -Match 'suspicious forward'
        $noTicket = New-TriageBecReviewBrief -Facts ([pscustomobject]@{
            mailbox = 'alex@contoso.com'
            rules = @()
            forwards = @()
            delegates = @()
            ticketNumber = ' '
            ticket = ''
        })
        $noTicket | Should -Match 'alex@contoso.com'
        $noTicket | Should -Not -Match 'Ticket'
        $window = Get-TriageBecWindow -Ticket 'User reported a forward on 9/20/2026 3:15 PM.' -Now ([datetime]'2026-09-26T12:00:00')
        $window.HadTime | Should -BeTrue
        $window.Days | Should -BeGreaterOrEqual 2
        $window.Note | Should -Match ([regex]::Escape((Get-TriageLocalZoneLabel)))
        $window.Note | Should -Not -Match 'UTC'
        $quietWindow = Get-TriageBecWindow -Ticket 'No clock time here.' -Now ([datetime]'2026-09-26T12:00:00')
        $quietWindow.HadTime | Should -BeFalse
        $quietWindow.Days | Should -Be 2
        ConvertFrom-TriageBecRepull -Text "Unknown.`nREPULL: 14" | Should -Be 14
        ConvertFrom-TriageBecRepull -Text "REPULL: 0" | Should -Be 0
        ConvertFrom-TriageBecRepull -Text "REPULL: 200" | Should -Be 90
        ConvertFrom-TriageBecRepull -Text "The link was created earlier.`nREPULL: 30 sharing" | Should -Be 30
        (@(ConvertFrom-TriageBecRepullLogs -Text "REPULL: 30 sharing, files") -join ',') | Should -Be 'sharing,files'
        (@(ConvertFrom-TriageBecRepullLogs -Text "REPULL: 14") -join ',') | Should -Be ''
        (@(ConvertFrom-TriageBecRepullLogs -Text "REPULL: 30 onedrive") -join ',') | Should -Be 'sharing'
        $gap = @'
What we did find is that this mailbox has been creating shareable links on OneDrive files.
We were not able to find when that particular link was created. It was set up before the two day window we could review.
'@
        $short = Get-TriageBecFollowUpPlan -Text $gap -CurrentDays 2
        $short.Days | Should -Be 30
        (@($short.Logs) -join ',') | Should -Be 'sharing'
        (Get-TriageBecFollowUpPlan -Text $gap -CurrentDays 30).Days | Should -Be 90
        (Get-TriageBecFollowUpPlan -Text $gap -CurrentDays 90).Days | Should -Be 0
        $named = Get-TriageBecFollowUpPlan -Text "REPULL: 14 trace" -CurrentDays 2
        $named.Days | Should -Be 14
        (@($named.Logs) -join ',') | Should -Be 'trace'
        (Get-TriageBecFollowUpPlan -Text "The mailbox is quiet.`nREPULL: 0" -CurrentDays 2).Days | Should -Be 0
        (Get-TriageBecFollowUpPlan -Text "Transport rules are dated well before the window.`nREPULL: 0" -CurrentDays 2).Days | Should -Be 0
        $packed = New-TriageBecReviewBrief -Facts ([pscustomobject]@{
            mailbox = 'alex@contoso.com'
            rules = @()
            forwards = @()
            delegates = @()
            windowNote = 'Search window is local.'
            messageTrace = @([pscustomobject]@{ Received = '2026-09-20T20:15:00Z'; Sender = 'alex@contoso.com'; Recipient = 'attacker@evil.example'; Status = 'Delivered'; Subject = 'wire'; Size = '9000000' })
            ual = @([pscustomobject]@{ Created = '2026-09-20T20:10:00Z'; Operation = 'New-InboxRule'; User = 'alex@contoso.com'; RecordType = 'ExchangeAdmin' })
            dlp = @([pscustomobject]@{ Created = '2026-09-20T20:00:00Z'; Operation = 'DlpRuleMatch'; Policy = 'Credit cards'; Target = 'payroll.xlsx' })
            alerts = @([pscustomobject]@{ Created = '2026-09-20T20:05:00Z'; Title = 'Malware detected'; Severity = 'high'; Category = 'Malware' })
            transport = @([pscustomobject]@{ Name = 'Redirect payroll'; State = 'Enabled'; WhenChanged = '2026-09-20T20:00:00Z'; Description = 'redirect to attacker@evil.example' })
            connectors = @([pscustomobject]@{ Direction = 'Outbound'; Name = 'Smart host'; Enabled = 'enabled'; WhenCreated = '2026-09-20T19:00:00Z'; Hosts = 'mail.evil.example' })
        })
        $packed | Should -Match 'attacker@evil.example'
        $packed | Should -Match 'New-InboxRule'
        $packed | Should -Match 'size 9000000'
        $packed | Should -Match 'Credit cards'
        $packed | Should -Match 'payroll.xlsx'
        $packed | Should -Match 'Malware detected'
        $packed | Should -Match 'Redirect payroll'
        $packed | Should -Match 'mail.evil.example'
        $packed | Should -Match ([regex]::Escape((ConvertTo-TriagePresentedTime '2026-09-20T20:15:00Z')))
        $packed | Should -Match ([regex]::Escape((Get-TriageLocalZoneLabel)))
        $packed | Should -Not -Match '20:15:00Z'
        $blastBrief = New-TriageBecReviewBrief -Facts ([pscustomobject]@{
            mailbox = 'alex@contoso.com'
            rules = @()
            forwards = @()
            delegates = @()
            blasts = @([pscustomobject]@{
                Subject = 'Payroll update'
                Received = '2026-09-20T20:15:00Z'
                Count = 5
                Recipients = @('a@evil.example', 'b@evil.example', 'c@evil.example', 'd@evil.example', 'e@evil.example')
            })
        })
        $blastBrief | Should -Match 'Payroll update'
        $blastBrief | Should -Match '5 recipients delivered'
        $blastBrief | Should -Match 'own line'
        $blastBrief | Should -Not -Match 'a@evil.example'
        $failed = New-TriageBecReviewBrief -Facts ([pscustomobject]@{
            mailbox = 'tspitaletto@metroeye.biz'
            rules = @()
            forwards = @()
            delegates = @()
            pullNotes = @('REMEDIATE_FAILED:Unknown remediation command')
        })
        $failed | Should -Match 'were not collected'
        $failed | Should -Not -Match 'Unknown remediation command'
        $failed | Should -Not -Match 'Message trace \(UTC\)'
    }
}

Describe 'Email triage log plan' {
    It 'keeps only catalog logs and clamps the window' {
        $plan = ConvertFrom-TriageLogPlanText -Text @"
DAYS: 200
LOG: IncludeSignInLogs
LOG: IncludeSignInLogs
LOG: IncludeNotARealLog
NOTE: Sign-in only.
MORE: no
"@
        $plan.Days | Should -Be 90
        @($plan.Logs) | Should -Be @('IncludeSignInLogs')
        $plan.Note | Should -Be 'Sign-in only.'
        $plan.More | Should -BeFalse
        $short = ConvertFrom-TriageLogPlanText -Text "DAYS: 0`nLOG: IncludeInboxRules"
        $short.Days | Should -Be 1
        $short.More | Should -BeTrue
    }

    It 'sends the ticket and file names, not file contents' {
        $folder = Join-Path $TestDrive 'pack'
        New-Item -ItemType Directory -Path $folder | Out-Null
        "UserPrincipalName`nsecret-row@contoso.com" | Set-Content -LiteralPath (Join-Path $folder 'SignInLogs.csv') -Encoding UTF8
        $brief = New-TriageLogPlanBrief -Ticket 'User reported a suspicious sign-in this morning.' -Phase 'followup' -Inventory (Get-TriageLogFolderInventory -Folder $folder)
        $brief | Should -Match 'suspicious sign-in'
        $brief | Should -Match 'SignInLogs.csv'
        $brief | Should -Match 'IncludeSignInLogs'
        $brief | Should -Not -Match 'secret-row@contoso.com'
        { Test-TriageLogReportFolder -Folder 'C:\Windows' } | Should -Throw -ExpectedMessage 'That folder is outside ExchangeOnlineAnalyzer reports.'
    }
}

Describe 'Memberberry ticket note' {
    It 'reads the company from a labeled line or a Company heading' {
        Get-TriageTicketCompany -Ticket "Board`nCompany: Contoso Ltd`nContact: Alex Kim" | Should -Be 'Contoso Ltd'
        $stacked = @"
Company

*
Contoso Ltd
Contact
"@
        Get-TriageTicketCompany -Ticket $stacked | Should -Be 'Contoso Ltd'
    }

    It 'matches a client by name or associated company and keeps the brief short' {
        $path = Join-Path $TestDrive 'exceptions.json'
        @{
            '_global' = @{ notes = 'ignore' }
            'Contoso Ltd' = @{
                detail = 'LOW'
                include_what_happened = $false
                names = @(@{ name = 'Alex Kim'; preferred = 'Alex' })
                associated_companies = @('Contoso Manufacturing')
                notes = 'VIP desk'
            }
            'Other Co' = @{ detail = 'HIGH'; notes = 'do not leak' }
        } | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $path -Encoding UTF8
        $hit = Find-TriageMemberberryClient -Company 'Contoso Manufacturing' -ExceptionsPath $path
        $hit.Name | Should -Be 'Contoso Ltd'
        $brief = ConvertTo-TriageMemberberryClientBrief -Record $hit.Record -Name $hit.Name
        $brief | Should -Match 'Detail: LOW'
        $brief | Should -Match 'Alex Kim -> Alex'
        $brief | Should -Not -Match 'do not leak'
        $missing = Find-TriageMemberberryClient -Company 'Missing Co' -ExceptionsPath $path
        $missing | Should -BeNullOrEmpty
    }

    It 'loads technician settings without requiring other fields' {
        $path = Join-Path $TestDrive 'settings.json'
        @{ technician_name = 'Avery Case'; timezone = 'CST'; contact_number = '555-0100' } | ConvertTo-Json | Set-Content -LiteralPath $path -Encoding UTF8
        $settings = Get-TriageMemberberrySettings -SettingsPath $path
        $settings.TechnicianName | Should -Be 'Avery Case'
        $settings.Timezone | Should -Be 'CST'
        $settings.ContactNumber | Should -Be '555-0100'
        $settings.CompanyName | Should -Be ''
    }

    It 'parses memory vault search text and a finished note' {
        $raw = @"
Search: "contoso"
Results: 1 (10ms)

  [1] 0.8120  [soc]  notes.md
      Hidden inbox rules are treated as account compromise.
"@
        $hits = @(ConvertFrom-TriageMemoryVaultText -Text $raw)
        $hits.Count | Should -Be 1
        $hits[0].Space | Should -Be 'soc'
        $hits[0].Text | Should -Match 'Hidden inbox rules'

        $note = ConvertFrom-TriageMemberberryNote -Text @"
Subject: [Resolved]: Inbox rule - Ticket #1001
Classification: False Positive

Good afternoon Alex,

We have classified this as FALSE POSITIVE.

Thanks,

Avery Case

MEMBERBERRY EXCEPTION SUGGESTION: Contoso Ltd - add the forwarding partner.
"@
        $note.Subject | Should -Match 'Ticket #1001'
        $note.Classification | Should -Be 'False Positive'
        $note.Body | Should -Match 'FALSE POSITIVE'
        $note.Body | Should -Not -Match 'EXCEPTION SUGGESTION'
        $note.Suggestion | Should -Match 'forwarding partner'
    }

    It 'puts the review in the brief and labels memories as reference' {
        $brief = New-TriageMemberberryNoteBrief -Review 'Hidden rule to attacker@evil.example. No remediation was run.' -TicketNumber '1001' -Company 'Contoso Ltd' -Contact 'Alex Kim' -Greeting 'Good afternoon' -ClientBrief "Client: Contoso Ltd`nDetail: LOW" -MemoryBrief '- [soc] Hidden rules are compromise.'
        $brief | Should -Match 'attacker@evil.example'
        $brief | Should -Match 'reference only'
        $brief | Should -Match 'Good afternoon Alex,'
        $brief | Should -Match 'Ticket number: 1001'
        { Invoke-TriageMemberberryNote -SettingsPath (Join-Path $TestDrive 'missing.json') -Review ' ' } | Should -Throw -ExpectedMessage 'Finish a review before writing the ticket note.'
        { Invoke-TriageMemberberryNote -SettingsPath (Join-Path $TestDrive 'missing.json') -Audience 'alert' -Review ' ' -Ticket ' ' } | Should -Throw -ExpectedMessage 'Open a ticket with alert text before triaging.'
    }

    It 'groups repeated alert log lines and keeps different addresses' {
        $lines = 1..4 | ForEach-Object { "2026-09-27T14:0${_}:00Z alice signed in from 203.0.113.10" }
        $lines += '2026-09-27T15:00:00Z alice signed in from 198.51.100.20'
        $packed = Compress-TriageAlertLogs -Text ($lines -join "`n")
        $packed.Grouped | Should -Be 1
        $packed.Text | Should -Match '\[4 similar lines\]'
        $packed.Text | Should -Match '198.51.100.20'
        $packed.Text | Should -Not -Match '\[5 similar lines\]'
        $note = Add-TriageAlertReviewBanner @"
HATZ-AI POWERED REVIEW OF THE ALERT ONLY
This note reviews the alert text on the ticket. It is not an investigation.

Alert leaning: Needs investigation
"@
        $note | Should -Match '^HATZ-AI POWERED REVIEW OF THE ALERT ONLY'
        ([regex]::Matches($note, 'HATZ-AI POWERED REVIEW OF THE ALERT ONLY')).Count | Should -Be 1
        $note | Should -Match 'Needs investigation'
        $lists = ConvertTo-TriageMemberberryExceptionBrief -Record ([pscustomobject]@{
                false_positive_patterns = @('HP Touchpoint Analytics is a known false positive.')
                authorized_tools        = @()
                blocked_tools           = @('PCAppStore')
            })
        $lists | Should -Match 'HP Touchpoint Analytics'
        $lists | Should -Match 'PCAppStore'
        $accounts = ConvertTo-TriageMemberberryExceptionBrief -Record ([pscustomobject]@{
                notes = "Advanced IP Scanner stays blocked.`nUsernames like rrc, rradmin, and rrcadmin are River Run accounts."
            })
        $accounts | Should -Match 'rrcadmin'
        $accounts | Should -Not -Match 'Advanced IP Scanner'
        $brief = New-TriageAlertReviewBrief -Alert ($lines -join "`n") -TicketNumber '1901107' -Company 'Contoso' -GlobalBrief $lists
        $brief | Should -Match 'Ticket number: 1901107'
        $brief | Should -Match 'similar lines'
        $brief | Should -Match '198.51.100.20'
        $brief | Should -Match 'Check Memberberry exceptions before'
        $brief | Should -Match 'HP Touchpoint Analytics'
        $brief.IndexOf('Global Memberberry exceptions') | Should -BeLessThan $brief.IndexOf('Alert text:')
    }

    It 'reads the ticket contact and writes a greeting with no name when none is supplied' {
        Get-TriageTicketContact -Ticket "Company: Contoso Ltd`nContact: Alex Kim" | Should -Be 'Alex Kim'
        $quiet = New-TriageMemberberryNoteBrief -Review 'No hidden rules.' -Greeting 'Good morning' -Audience 'client'
        $quiet | Should -Match 'Good morning,'
        $quiet | Should -Match 'Do not invent a name'
        $internal = New-TriageMemberberryNoteBrief -Review 'No hidden rules.' -Audience 'internal' -ActionsText '- 2026-09-26T17:00:00Z alex@contoso.com revoke success sessions revoked'
        $internal | Should -Match 'internal ConnectWise'
        $internal | Should -Match 'Do not greet'
        $internal | Should -Match 'revoke success'
        $report = New-TriageMemberberryNoteBrief -Review 'Forwarding to attacker@evil.example.' -Audience 'incident'
        $report | Should -Match 'legal, management, and insurance'
        $report | Should -Match 'No remediation steps were run'
        $report | Should -Match 'attacker@evil.example'
    }

    It 'splits store suggestions out of a review and keeps the repull line' {
        $parsed = ConvertFrom-TriageStoreSuggestions -Text @"
Unknown.
MEMBERBERRY EXCEPTION SUGGESTION: Contoso Ltd - VIP desk uses Teams.
MEMORY VAULT: Hidden inbox rules on this tenant were created by the user.
REPULL: 0
"@
        $parsed.Client | Should -Be 'Contoso Ltd'
        $parsed.Note | Should -Match 'Teams'
        $parsed.Memory | Should -Match 'Hidden inbox rules'
        $parsed.Text | Should -Match 'REPULL: 0'
        $parsed.Text | Should -Not -Match 'MEMORY VAULT'
        $parsed.Text | Should -Not -Match 'EXCEPTION SUGGESTION'
        $brief = New-TriageBecReviewBrief -Facts ([pscustomobject]@{
            mailbox = 'a@b.test'
            company = 'Zz No Such Client 999'
            rules = @()
            forwards = @()
            delegates = @()
        })
        $brief | Should -Match 'no client matched Zz No Such Client 999'
        $brief | Should -Match 'not evidence'
        $ops = @"
Hidden inbox rule forwarded mail to attacker@evil.example.

Operational note: the pack contains the line REMEDIATE_FAILED:Unknown remediation command. This indicates a remediation command was attempted and failed with an unknown command error, but the pack gives no facts about what was being remediated or why. This does not by itself evidence compromise, and no supporting artifacts appear in any pulled source. It is recommended that the failed remediation be reconciled with the ticketing or remediation system to confirm what action was intended. No remediation steps were run in this session.
"@
        $clean = Remove-TriageOperationalNotes $ops
        $clean | Should -Match 'attacker@evil.example'
        $clean | Should -Not -Match 'REMEDIATE_FAILED'
        $clean | Should -Not -Match 'Operational note'
        $clean | Should -Not -Match 'reconciled'
        $email = New-TriageClientEmailBody -Review $ops -Greeting 'Good afternoon' -Technician 'Avery Case'
        $email | Should -Match 'Good afternoon,'
        $email | Should -Match 'attacker@evil.example'
        $email | Should -Match 'Avery Case'
        $email | Should -Not -Match 'REMEDIATE_FAILED'
        $quietWindow = New-TriageClientEmailBody -Review "We could not find when the link was created.`n`nREPULL: 30 sharing" -Greeting 'Good morning' -Technician 'Avery Case'
        $quietWindow | Should -Match 'could not find'
        $quietWindow | Should -Not -Match 'REPULL'
        $noteBrief = New-TriageMemberberryNoteBrief -Review $ops -Audience 'internal'
        $noteBrief | Should -Match 'attacker@evil.example'
        $noteBrief | Should -Not -Match 'REMEDIATE_FAILED'
        $noteBrief | Should -Not -Match 'Operational note'
    }

    It 'leaves the log-pull app registration out of the mailbox review' {
        $brief = New-TriageBecReviewBrief -Facts ([pscustomobject]@{
            mailbox = 'alex@contoso.com'
            apps = @(
                [pscustomobject]@{ displayName = 'River Run Security Investigator'; created = '2026-09-26T16:59:00Z' }
                [pscustomobject]@{ displayName = 'River Run Security Investigator'; created = '2026-09-26T16:59:30Z' }
                [pscustomobject]@{ displayName = 'Odd Grant'; created = '2026-09-26T16:59:00Z' }
            )
            ual = @(
                [pscustomobject]@{ operation = 'Add service principal'; target = 'River Run Security Investigator'; user = 'admin@contoso.com'; created = '2026-09-26T16:59:00Z' }
                [pscustomobject]@{ operation = 'MailItemsAccessed'; target = 'Inbox'; user = 'alex@contoso.com'; created = '2026-09-26T15:00:00Z' }
            )
            signIns = @(
                [pscustomobject]@{ created = '2026-09-26T16:59:00Z'; ip = '1.2.3.4'; app = 'River Run Security Investigator'; result = 'success' }
                [pscustomobject]@{ created = '2026-09-26T15:00:00Z'; ip = '5.6.7.8'; app = 'Office 365'; result = 'success' }
            )
        })
        $brief | Should -Not -Match 'River Run Security Investigator'
        $brief | Should -Match 'Odd Grant'
        $brief | Should -Match 'MailItemsAccessed'
        $brief | Should -Match 'Office 365'
    }

    It 'treats a confirmed incident in the ticket as already established' {
        Test-TriageConfirmedIncident -Text 'Confirmed BEC. The mailbox was used to send invoices.' | Should -Be $true
        Test-TriageConfirmedIncident -Text 'The account was compromised on Monday.' | Should -Be $true
        Test-TriageConfirmedIncident -Text 'Please check this suspicious email from sales@example.com.' | Should -Be $false
        Test-TriageConfirmedIncident -Text 'The incident is not confirmed.' | Should -Be $false
    }

    It 'passes a technician revision into the review and the notes' {
        $brief = New-TriageBecReviewBrief -Facts ([pscustomobject]@{
            mailbox = 'a@contoso.com'
            revision = 'Treat the blast as the mailbox owner.'
            priorReview = 'Compromise: unknown.'
        })
        $brief | Should -Match 'Treat the blast as the mailbox owner'
        $brief | Should -Match 'Compromise: unknown'
        $brief | Should -Match 'Do not add facts'
        $brief | Should -Match 'evidence rules still apply'
        $brief | Should -Match 'Never conclude true positive'
        $notes = New-TriageMemberberryNoteBrief -Review 'Nothing got in.' -Audience 'client' -Revision 'Do not mention MFA.' -CurrentDrafts 'Client email: please change the password.'
        $notes | Should -Match 'Nothing got in'
        $notes | Should -Match 'Do not mention MFA'
        $notes | Should -Match 'please change the password'
        $notes | Should -Match 'evidence rules still apply'
        $notes | Should -Match 'Never conclude true positive'
        $notes | Should -Not -Match 'always fine'
        $internalRules = New-TriageMemberberryNoteBrief -Review 'No hidden rules.' -Audience 'internal'
        $incidentRules = New-TriageMemberberryNoteBrief -Review 'No hidden rules.' -Audience 'incident'
        $internalRules | Should -Match 'Never recommend or suggest a password reset'
        $incidentRules | Should -Match 'Never recommend or suggest a password reset'
        $ask = New-TriageClientEmailBody -Review "The sign-in failed.`nPlease reset the password and revoke their sessions." -Greeting 'Good morning'
        $ask | Should -Not -Match 'reset the password'
        $ask | Should -Not -Match 'revoke their sessions'
        $ask | Should -Match 'The sign-in failed'
        $kept = New-TriageClientEmailBody -Review "A malicious inbox rule forwarded mail to attacker@evil.example.`nReset the password." -Greeting 'Good morning'
        $kept | Should -Match 'Reset the password'
        $reviewPrompt = Get-TriageBecReviewSystemPrompt -ConfirmedIncident $false
        $reviewPrompt | Should -Match 'VERDICT: ask the client'
        $reviewPrompt | Should -Match 'follow-up pull is not evidence'
        $reviewPrompt | Should -Match 'Never conclude true positive'
        $reviewPrompt | Should -Match 'evidence rules still apply'
        $reviewPrompt | Should -Not -Match '(?m)^EVIDENCE RULES$'
        $confirmedPrompt = Get-TriageBecReviewSystemPrompt -ConfirmedIncident $true
        $confirmedPrompt | Should -Match 'Do not recommend a password reset or session revocation'
        $moduleText = Get-Content -LiteralPath (Get-Module EmailTriage).Path -Raw
        $moduleText | Should -Match 'do not suggest a password change, a password reset, signing out everywhere'
        $moduleText | Should -Not -Match 'a password change is still fine'
        $moduleText | Should -Not -Match 'Asking the client to change a password is always fine'
        Get-TriageBecVerdict -Text "VERDICT: ask the client`nAsk whether this was theirs." | Should -Be 'ask the client'
        Get-TriageBecVerdict -Text "VERDICT: true positive" | Should -Be 'true positive'
        Test-TriageCompromiseEvidence -Text 'Failed sign-in only.' | Should -BeFalse
        Test-TriageCompromiseEvidence -Text 'Hidden inbox rule forwarded mail.' | Should -BeTrue
        InModuleScope EmailTriage {
            $split = Split-TriageTechnicianRevision -Review "Nothing got in.`n`n<<<TECHNICIAN_REVISION>>>`nDo not mention MFA.`n`n<<<CURRENT_DRAFTS>>>`nClient email: hello"
            $split.Review | Should -Be 'Nothing got in.'
            $split.Revision | Should -Be 'Do not mention MFA.'
            $split.Drafts | Should -Be 'Client email: hello'
        }
    }

    It 'greets the contact by the Memberberry preferred name' {
        InModuleScope EmailTriage {
            $record = '{"detail":"STANDARD","names":{"Christine Wilms":"Chris"}}' | ConvertFrom-Json
            Get-TriageMemberberryGreetingName -Record $record -Contact 'Christine Wilms' | Should -Be 'Chris'
            Get-TriageMemberberryGreetingName -Record $record -Contact 'Christine' | Should -Be 'Chris'
            Get-TriageMemberberryGreetingName -Record $record -Contact 'Christine Smith' | Should -Be ''
            Get-TriageMemberberryGreetingName -Record $record -Contact 'Dana Wilms' | Should -Be ''
            Get-TriageMemberberryGreetingName -Record $record -Contact 'Christine Willms' | Should -Be 'Chris'
            Get-TriageMemberberryGreetingName -Record $record -Contact 'Christine Wills' | Should -Be 'Chris'
            Get-TriageMemberberryGreetingName -Record $record -Contact 'Christine Williams' | Should -Be ''
            $shared = '{"names":{"Christopher Purtell":"Chrispy","Christopher Knospe":"Chris"}}' | ConvertFrom-Json
            Get-TriageMemberberryGreetingName -Record $shared -Contact 'Christopher' | Should -Be ''
            Get-TriageMemberberryGreetingName -Record $shared -Contact 'Christopher Knospe' | Should -Be 'Chris'
            ConvertTo-TriageMemberberryNames $record.names | Should -Be 'Christine Wilms -> Chris'
            $brief = New-TriageMemberberryNoteBrief -Review 'Review text.' -Contact 'Christine Wilms' -Greeting 'Good morning' -Audience client -GreetingName 'Chris'
            $brief | Should -Match 'Greeting line: Good morning Chris,'
            $body = New-TriageClientEmailBody -Review 'We checked the mailbox.' -Contact 'Christine Wilms' -Greeting 'Good morning' -GreetingName 'Chris'
            $body | Should -Match '^Good morning Chris,'
        }
    }
}
