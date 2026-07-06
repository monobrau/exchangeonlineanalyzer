# ReportAnalysis.psm1
# Rule-based analysis of security investigation reports - reduces LLM reliance
# Supports both in-memory Report objects and folder-based CSV analysis

$script:DefaultSuspiciousKeywords = @(
    'invoice','payment','bank','wire','transfer','refund','urgent','verify','confirm',
    'password','credential','login','account','suspicious','forward','external'
)

# Common BEC attacker inbox rule names - used to hide malicious rules (exact match, case-insensitive)
$script:BECSuspiciousRuleNames = @(
    '!', '!!', '1', '2', '3', 'a', 'aa', 'ab', 'x', 'xx', 'test', 'test1', 'test2', 'rule', 'rule1', 'rule2',
    'new', 'temp', 'tmp', 'asdf', 'qwerty', 'fgh', 'xyz', 'ok', 'n', 'y', 'no', 'yes', 'copy', 'move',
    'read', 'unread', 'read mail', 'mark read', 'organize', 'sort', 'filter', 'clean', 'cleanup',
    '.', '..', '...', '-', '--', '---', 'zzz', 'aaa', 'bbb', 'asd', 'qwe', 'zxc', 'abc', '123'
)

# Typosquatting patterns (0/o, 1/l/i, rn/m, etc.)
$script:TyposquatPatterns = @(
    @{ Original='microsoft'; Pattern='m[i1l]cr[o0]s[o0]ft|micr[o0]s[o0]ft' },
    @{ Original='google'; Pattern='g[o0][o0]gle|g[o0]ogl[e3]' },
    @{ Original='outlook'; Pattern='[o0]utl[o0][o0]k' },
    @{ Original='office'; Pattern='[o0]ff[i1l]ce' },
    @{ Original='login'; Pattern='l[o0]g[i1l]n' },
    @{ Original='account'; Pattern='acc[o0]unt' }
)

# Known-good sender patterns for external_message_spike (excluded from findings)
$script:ExternalMessageSpikeAllowlist = @(
    'govdelivery\.com', 'service\.govdelivery', 'newsletters\.', '\.newsletter',
    'squareup\.com', 'messaging\.squareup', 'mailchimp\.com', 'constantcontact\.com',
    'sendgrid\.net', 'mailgun\.', 'amazonses\.com', 'notifications\.',
    'noreply@', 'no-reply@', 'donotreply@', 'mailer-daemon@'
)

# Remediation text per finding type
$script:RemediationByType = @{
    'suspicious_inbox_rule' = 'Review and remove if unauthorized. Check for BEC compromise.'
    'hidden_inbox_rule' = 'Verify rule purpose. Hidden rules can hide malicious activity.'
    'shared_mailbox_forward' = 'Shared mailboxes are high-value targets. Verify forwarding is authorized.'
    'suspicious_domain_forward' = 'Domain may be typosquatting. Verify recipient before allowing.'
    'transport_forward' = 'Review transport rule. Ensure redirect/forward is business-justified.'
    'transport_spoof_bypass' = 'Spoofing bypass weakens protection. Review necessity.'
    'transport_scope_all' = 'Rule applies to all mail. Ensure actions are appropriate for broad scope.'
    'high_risk_app' = 'Review app permissions. Revoke if not required.'
    'risky_app_permissions' = 'App has Mail.Read, Mail.Send, or other high-privilege scopes. Review and revoke if unauthorized.'
    'unverified_app' = 'App has no verified publisher. Verify legitimacy before use.'
    'user_consent_app' = 'User consent apps can be phished. Consider admin consent only.'
    'high_risk_ca_policy' = 'Review CA policy. Restrict or disable if overly permissive.'
    'ca_no_mfa_all_users' = 'Policy applies to all users without MFA. Add MFA requirement.'
    'ca_broad_exclusions' = 'Policy excludes all locations or has broad exclusions. Review trusted locations.'
    'no_mfa' = 'Enable MFA for user. Consider Conditional Access to enforce.'
    'external_message_spike' = 'Review for business justification. May indicate compromise or data exfil.'
    'high_risk_signin' = 'Investigate user. Consider password reset and session revocation.'
    'failed_signin' = 'Check for credential stuffing or brute force. Consider blocking or MFA.'
    'trivial_inbox_rule_condition' = 'BEC pattern: rule condition matches almost all mail. Review and remove if unauthorized.'
    'suspicious_rule_name' = 'BEC pattern: rule name commonly used by attackers to hide malicious rules. Review and remove if unauthorized.'
    'move_to_rss_or_conversation_history' = 'BEC pattern: mail moved to RSS Feeds or Conversation History to hide. Review and remove if unauthorized.'
    'move_and_delete' = 'BEC pattern: mail moved to folder and deleted. Review and remove if unauthorized.'
    'missing_data' = 'Download or collect missing data for complete analysis.'
    'stale_data' = 'Data may be outdated. Re-run collection for current assessment.'
}

function Test-TrivialInboxRuleCondition {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $false }
    # Check each semicolon-separated part (conditions can be ".;" or ".," or ". , ")
    $parts = $Value -split ';' | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    foreach ($p in $parts) {
        # Trivial: . , ,, .. ... ; : - etc. - BEC attackers use these to match almost all mail
        if ($p -match '^[\s.,;:\-_''"…]+$') { return $true }
    }
    return $false
}

function Test-SuspiciousDomain {
    param([string]$Email)
    if (-not $Email -or $Email -notmatch '@') { return $false }
    $domain = ($Email -split '@')[-1].ToLower()
    foreach ($t in $script:TyposquatPatterns) {
        if ($domain -match $t.Pattern) { return $true }
    }
    # Check for homograph-style (lookalike chars)
    if ($domain -match '[0-9]' -and $domain -match 'microsoft|google|outlook|office|login|account') { return $true }
    return $false
}

function Get-InboxRuleFindings {
    param(
        [Parameter(Mandatory=$false)]
        [array]$Rules,
        [Parameter(Mandatory=$false)]
        [string]$CsvPath,
        [Parameter(Mandatory=$false)]
        [array]$SuspiciousKeywords = $script:DefaultSuspiciousKeywords,
        [Parameter(Mandatory=$false)]
        [hashtable]$MailboxRecipientTypeMap = @{}
    )
    $items = $null
    if ($Rules -and $Rules.Count -gt 0) {
        $items = $Rules
    } elseif ($CsvPath -and (Test-Path $CsvPath)) {
        try { $items = Import-Csv -Path $CsvPath -ErrorAction Stop } catch { return @{ Findings = @(); Count = 0 } }
    }
    if (-not $items -or $items.Count -eq 0) { return @{ Findings = @(); Count = 0 } }

    $findings = @()
    foreach ($r in $items) {
        $forwardTo = if ($r.ForwardTo) { $r.ForwardTo } else { '' }
        $name = if ($r.Name) { $r.Name } else { '' }
        $mailbox = if ($r.MailboxOwner) { $r.MailboxOwner } else { 'Unknown' }
        $isHidden = $false
        if ($r.PSObject.Properties['IsHidden']) { $isHidden = $r.IsHidden -eq $true -or $r.IsHidden -eq 'True' }
        $recipientType = if ($MailboxRecipientTypeMap[$mailbox]) { $MailboxRecipientTypeMap[$mailbox] } else { '' }
        $isSharedMailbox = $recipientType -match 'SharedMailbox|Shared'

        if ($forwardTo -match '@' -and $forwardTo -notmatch ';') {
            $extractAddr = ($forwardTo -replace '^.*\[SMTP:(.+)\].*$','$1') -replace '^["'']|["'']$',''
            if (-not $extractAddr -or $extractAddr -notmatch '@') { $extractAddr = $forwardTo }
            if (Test-SuspiciousDomain -Email $extractAddr) {
                $findings += [PSCustomObject]@{ Type = 'suspicious_domain_forward'; Severity = 'High'; Detail = "Forward to suspicious domain (possible typosquatting): $mailbox -> $extractAddr"; Source = $name }
            } elseif ($isSharedMailbox) {
                $findings += [PSCustomObject]@{ Type = 'shared_mailbox_forward'; Severity = 'High'; Detail = "Shared mailbox external forwarding: $mailbox -> $forwardTo"; Source = $name }
            } else {
                $findings += [PSCustomObject]@{ Type = 'suspicious_inbox_rule'; Severity = 'High'; Detail = "External forwarding: $mailbox -> $forwardTo"; Source = $name }
            }
        } elseif ($forwardTo -match '@') {
            if ($isSharedMailbox) {
                $findings += [PSCustomObject]@{ Type = 'shared_mailbox_forward'; Severity = 'High'; Detail = "Shared mailbox multiple external forwards: $mailbox"; Source = $name }
            } else {
                $findings += [PSCustomObject]@{ Type = 'suspicious_inbox_rule'; Severity = 'Medium'; Detail = "Multiple external forwards: $mailbox"; Source = $name }
            }
        }
        if ($isHidden -and $name -and $name -notmatch 'system|default|outlook|microsoft|junk|clutter|archive') {
            $findings += [PSCustomObject]@{ Type = 'hidden_inbox_rule'; Severity = 'Medium'; Detail = "Hidden rule: $mailbox"; Source = $name }
        }
        # BEC pattern: trivial conditions (. , ,, .. etc.) match almost all mail - flag as High
        $subjectContains = if ($r.SubjectContains) { $r.SubjectContains } else { '' }
        $fromContains = if ($r.FromAddressContains) { $r.FromAddressContains } else { '' }
        if ((Test-TrivialInboxRuleCondition -Value $subjectContains) -or (Test-TrivialInboxRuleCondition -Value $fromContains)) {
            $condNote = @()
            if (Test-TrivialInboxRuleCondition -Value $subjectContains) { $condNote += "SubjectContains='$subjectContains'" }
            if (Test-TrivialInboxRuleCondition -Value $fromContains) { $condNote += "FromAddressContains='$fromContains'" }
            $findings += [PSCustomObject]@{ Type = 'trivial_inbox_rule_condition'; Severity = 'High'; Detail = "Inbox rule with trivial condition (matches almost all mail): $mailbox - $($condNote -join '; ')"; Source = $name }
        }
        # BEC pattern: suspicious rule names attackers use to hide malicious rules
        $nameTrimmed = $name.Trim().ToLower()
        if ($nameTrimmed -and ($script:BECSuspiciousRuleNames | Where-Object { $_ -eq $nameTrimmed })) {
            $findings += [PSCustomObject]@{ Type = 'suspicious_rule_name'; Severity = 'High'; Detail = "Inbox rule with suspicious BEC-style name: $mailbox - '$name'"; Source = $name }
        }
        # BEC pattern: move to RSS Feeds or Conversation History (common hiding spots)
        $moveToFolder = if ($r.MoveToFolder) { $r.MoveToFolder } elseif ($r.MoveToFolderName) { $r.MoveToFolderName } else { '' }
        if ($moveToFolder -and $moveToFolder -match 'RSS\s*Feeds|Conversation\s*History') {
            $findings += [PSCustomObject]@{ Type = 'move_to_rss_or_conversation_history'; Severity = 'High'; Detail = "Inbox rule moves mail to '$($moveToFolder -replace '^[^\\]+:\\', '')': $mailbox"; Source = $name }
        }
        # BEC pattern: move to any folder AND delete
        $deleteMsg = $r.DeleteMessage -eq $true -or $r.DeleteMessage -eq 'True'
        if ($moveToFolder -and $deleteMsg) {
            $findings += [PSCustomObject]@{ Type = 'move_and_delete'; Severity = 'High'; Detail = "Inbox rule moves mail to folder and deletes: $mailbox - '$($moveToFolder -replace '^[^\\]+:\\', '')'"; Source = $name }
        }
    }
    return @{ Findings = $findings; Count = $findings.Count }
}

function Get-TransportRuleFindings {
    param(
        [Parameter(Mandatory=$false)]
        [array]$Rules,
        [Parameter(Mandatory=$false)]
        [string]$CsvPath
    )
    $items = $null
    if ($Rules -and $Rules.Count -gt 0) {
        $items = $Rules
    } elseif ($CsvPath -and (Test-Path $CsvPath)) {
        try { $items = Import-Csv -Path $CsvPath -ErrorAction Stop } catch { return @{ Findings = @(); Count = 0 } }
    }
    if (-not $items -or $items.Count -eq 0) { return @{ Findings = @(); Count = 0 } }

    $findings = @()
    foreach ($r in $items) {
        $actions = if ($r.ActionsSummary) { $r.ActionsSummary } else { '' }
        $conditions = if ($r.ConditionsSummary) { $r.ConditionsSummary } else { '' }
        $name = if ($r.Name) { $r.Name } else { 'Unknown' }
        $combined = "$actions $conditions"
        $hasForward = $combined -match 'ForwardTo|RedirectMessageTo|RedirectTo'
        $scopeAll = $combined -match '"All"|"Everyone"|"InOrganization"|RecipientScope.*All'

        if ($hasForward) {
            if ($scopeAll) {
                $findings += [PSCustomObject]@{ Type = 'transport_scope_all'; Severity = 'High'; Detail = "Transport rule forwards/redirects mail and applies to all"; Source = $name }
            } else {
                $findings += [PSCustomObject]@{ Type = 'transport_forward'; Severity = 'High'; Detail = "Transport rule forwards/redirects mail"; Source = $name }
            }
        }
        if ($combined -match 'BypassSpoofing') {
            $findings += [PSCustomObject]@{ Type = 'transport_spoof_bypass'; Severity = 'High'; Detail = "Spoofing protection bypass"; Source = $name }
        }
    }
    return @{ Findings = $findings; Count = $findings.Count }
}

function Get-AppRegistrationFindings {
    param(
        [Parameter(Mandatory=$false)]
        [array]$Apps,
        [Parameter(Mandatory=$false)]
        [string]$CsvPath
    )
    $items = $null
    if ($Apps -and $Apps.Count -gt 0) {
        $items = $Apps
    } elseif ($CsvPath -and (Test-Path $CsvPath)) {
        try { $items = Import-Csv -Path $CsvPath -ErrorAction Stop } catch { return @{ Findings = @(); Count = 0 } }
    }
    if (-not $items -or $items.Count -eq 0) { return @{ Findings = @(); Count = 0 } }

    $riskyPerms = @('Mail.Read', 'Mail.ReadWrite', 'Mail.Send', 'User.Read.All', 'Directory.Read.All', 'Mailbox.ReadWrite')
    $findings = @()
    foreach ($a in $items) {
        $riskLevel = if ($a.RiskLevel) { $a.RiskLevel } else { '' }
        $name = if ($a.DisplayName) { $a.DisplayName } else { 'Unknown' }
        $publisher = if ($a.PublisherDomain) { $a.PublisherDomain } else { '' }
        $userConsent = $a.HasUserConsent -eq $true -or $a.HasUserConsent -eq 'True'
        $hasHighPriv = $a.HasHighPrivilegePermissions -eq $true -or $a.HasHighPrivilegePermissions -eq 'True'
        $hasSuspicious = $a.HasSuspiciousPermissions -eq $true -or $a.HasSuspiciousPermissions -eq 'True'
        $reqPerms = if ($a.RequiredPermissions) { $a.RequiredPermissions } else { '' }

        if ($riskLevel -eq 'High') {
            $findings += [PSCustomObject]@{ Type = 'high_risk_app'; Severity = 'High'; Detail = "High-risk app: $name"; Source = $publisher }
        }
        $hasRiskyPerm = $reqPerms -and ($riskyPerms | Where-Object { $reqPerms -match [regex]::Escape($_) })
        if ($hasHighPriv -or $hasSuspicious -or $hasRiskyPerm) {
            $permNote = if ($reqPerms) { " ($reqPerms)" } else { '' }
            $findings += [PSCustomObject]@{ Type = 'risky_app_permissions'; Severity = 'High'; Detail = "App with risky permissions: $name$permNote"; Source = $name }
        }
        if ([string]::IsNullOrWhiteSpace($publisher) -or $publisher -eq '') {
            $findings += [PSCustomObject]@{ Type = 'unverified_app'; Severity = 'Medium'; Detail = "Unverified publisher: $name"; Source = $name }
        }
        if ($userConsent) {
            $findings += [PSCustomObject]@{ Type = 'user_consent_app'; Severity = 'Medium'; Detail = "User consent enabled: $name"; Source = $name }
        }
    }
    return @{ Findings = $findings; Count = $findings.Count }
}

function Get-CAPolicyFindings {
    param(
        [Parameter(Mandatory=$false)]
        [array]$Policies,
        [Parameter(Mandatory=$false)]
        [string]$CsvPath
    )
    $items = $null
    if ($Policies -and $Policies.Count -gt 0) {
        $items = $Policies
    } elseif ($CsvPath -and (Test-Path $CsvPath)) {
        try { $items = Import-Csv -Path $CsvPath -ErrorAction Stop } catch { return @{ Findings = @(); Count = 0 } }
    }
    if (-not $items -or $items.Count -eq 0) { return @{ Findings = @(); Count = 0 } }

    $findings = @()
    foreach ($p in $items) {
        $riskLevel = if ($p.RiskLevel) { $p.RiskLevel } else { '' }
        $name = if ($p.DisplayName) { $p.DisplayName } else { 'Unknown' }
        $requiresMfa = $p.RequiresMfa -eq $true -or $p.RequiresMfa -eq 'True'
        $userAll = $p.UserIncludeAll -eq $true -or $p.UserIncludeAll -eq 'True'
        $locationAll = $p.LocationIncludeAll -eq $true -or $p.LocationIncludeAll -eq 'True'

        if ($riskLevel -eq 'High') {
            $findings += [PSCustomObject]@{ Type = 'high_risk_ca_policy'; Severity = 'High'; Detail = "High-risk CA policy: $name"; Source = $name }
        }
        if ($userAll -and -not $requiresMfa) {
            $findings += [PSCustomObject]@{ Type = 'ca_no_mfa_all_users'; Severity = 'High'; Detail = "CA applies to all users without MFA: $name"; Source = $name }
        }
        if ($locationAll -or ($p.UserExcludeCount -and [int]$p.UserExcludeCount -gt 10)) {
            $findings += [PSCustomObject]@{ Type = 'ca_broad_exclusions'; Severity = 'Medium'; Detail = "CA policy has broad exclusions (all locations or many excluded users): $name"; Source = $name }
        }
    }
    return @{ Findings = $findings; Count = $findings.Count }
}

function Get-MfaGapFindings {
    param(
        [Parameter(Mandatory=$false)]
        [array]$UserPosture,
        [Parameter(Mandatory=$false)]
        [string]$CsvPath
    )
    $items = $null
    if ($UserPosture -and $UserPosture.Count -gt 0) {
        $items = $UserPosture
    } elseif ($CsvPath -and (Test-Path $CsvPath)) {
        try { $items = Import-Csv -Path $CsvPath -ErrorAction Stop } catch { return @{ Findings = @(); Count = 0 } }
    }
    if (-not $items -or $items.Count -eq 0) { return @{ Findings = @(); Count = 0 } }

    $findings = @()
    $mfaCol = $null
    if ($items[0].PSObject.Properties['MfaCovered']) { $mfaCol = 'MfaCovered' }
    elseif ($items[0].PSObject.Properties['PerUserMfaStatus']) { $mfaCol = 'PerUserMfaStatus' }
    elseif ($items[0].PSObject.Properties['MfaStatus']) { $mfaCol = 'MfaStatus' }
    elseif ($items[0].PSObject.Properties['HasMfa']) { $mfaCol = 'HasMfa' }
    elseif ($items[0].PSObject.Properties['MfaEnabled']) { $mfaCol = 'MfaEnabled' }
    if (-not $mfaCol) { return @{ Findings = @(); Count = 0 } }

    $upnCol = if ($items[0].PSObject.Properties['UserPrincipalName']) { 'UserPrincipalName' } else { 'UPN' }
    foreach ($u in $items) {
        $mfaVal = $u.$mfaCol
        $hasMfa = $mfaVal -eq $true -or $mfaVal -eq 'True' -or $mfaVal -eq 'Yes' -or $mfaVal -match 'Enabled|Covered'
        if (-not $hasMfa) {
            $upn = $u.$upnCol
            $findings += [PSCustomObject]@{ Type = 'no_mfa'; Severity = 'Medium'; Detail = "User without MFA: $upn"; Source = $upn }
        }
    }
    return @{ Findings = $findings; Count = $findings.Count }
}

function Get-MessageTraceFindings {
    param(
        [Parameter(Mandatory=$false)]
        [array]$Trace,
        [Parameter(Mandatory=$false)]
        [string]$CsvPath,
        [Parameter(Mandatory=$false)]
        [int]$ExternalPercentileThreshold = 95,
        [Parameter(Mandatory=$false)]
        [int]$MinExternalCount = 25,
        [Parameter(Mandatory=$false)]
        [array]$AllowlistPatterns = $script:ExternalMessageSpikeAllowlist,
        [Parameter(Mandatory=$false)]
        [array]$TrustedDomainPatterns = @()
    )
    $items = $null
    if ($Trace -and $Trace.Count -gt 0) {
        $items = $Trace
    } elseif ($CsvPath -and (Test-Path $CsvPath)) {
        try { $items = Import-Csv -Path $CsvPath -ErrorAction Stop } catch { return @{ Findings = @(); Count = 0 } }
    }
    if (-not $items -or $items.Count -eq 0) { return @{ Findings = @(); Count = 0 } }

    $internalPattern = '\.onmicrosoft\.com|\.mail\.protection\.outlook'
    if ($TrustedDomainPatterns -and $TrustedDomainPatterns.Count -gt 0) {
        $internalPattern += '|' + ($TrustedDomainPatterns -join '|')
    }

    $findings = @()
    $recipientCol = $null
    foreach ($prop in $items[0].PSObject.Properties.Name) {
        if ($prop -match 'Recipient|To|RecipientAddress') { $recipientCol = $prop; break }
    }
    if (-not $recipientCol) { return @{ Findings = @(); Count = 0 } }

    $externalBySender = @{}
    foreach ($t in $items) {
        $recipients = $t.$recipientCol
        if (-not $recipients) { continue }
        $senderCol = if ($items[0].PSObject.Properties['SenderAddress']) { 'SenderAddress' } else { 'From' }
        $senderAddr = $t.$senderCol
        if (-not $senderAddr) { continue }
        if (-not $externalBySender.ContainsKey($senderAddr)) { $externalBySender[$senderAddr] = 0 }
        if ($recipients -match '@' -and $recipients -notmatch $internalPattern) {
            $externalBySender[$senderAddr]++
        }
    }
    if ($externalBySender.Count -eq 0) { return @{ Findings = @(); Count = 0 } }
    $counts = @($externalBySender.Values | Where-Object { $_ -gt 0 })
    if ($counts.Count -eq 0) { return @{ Findings = @(); Count = 0 } }
    $threshold = [Math]::Max(1, [int]([Math]::Ceiling($counts.Count * $ExternalPercentileThreshold / 100)))
    $sorted = $counts | Sort-Object -Descending
    $percentileVal = if ($threshold -le $sorted.Count) { $sorted[$threshold - 1] } else { $sorted[-1] }

    foreach ($senderAddr in $externalBySender.Keys) {
        $count = $externalBySender[$senderAddr]
        if ($count -lt $percentileVal -or $count -lt $MinExternalCount) { continue }
        $senderLower = $senderAddr.ToLower()
        $isAllowlisted = $false
        foreach ($pat in $AllowlistPatterns) {
            if ($senderLower -match $pat) { $isAllowlisted = $true; break }
        }
        if ($isAllowlisted) { continue }
        $findings += [PSCustomObject]@{ Type = 'external_message_spike'; Severity = 'Low'; Detail = "High external message volume: $senderAddr ($count external)"; Source = $senderAddr }
    }
    return @{ Findings = $findings; Count = $findings.Count }
}

function Get-SignInLogFindings {
    param(
        [Parameter(Mandatory=$false)]
        [array]$Logs,
        [Parameter(Mandatory=$false)]
        [string]$CsvPath,
        [Parameter(Mandatory=$false)]
        [int]$FailedSignInThreshold = 5
    )
    $items = $null
    if ($Logs -and $Logs.Count -gt 0) {
        $items = $Logs
    } elseif ($CsvPath -and (Test-Path $CsvPath)) {
        try { $items = Import-Csv -Path $CsvPath -ErrorAction Stop } catch { return @{ Findings = @(); Count = 0 } }
    }
    if (-not $items -or $items.Count -eq 0) { return @{ Findings = @(); Count = 0 } }

    $findings = @()
    $riskCol = $null
    foreach ($prop in $items[0].PSObject.Properties.Name) {
        if ($prop -match 'RiskLevel|Risk') { $riskCol = $prop; break }
    }
    $statusCol = if ($items[0].PSObject.Properties['Status']) { 'Status' } else { 'ResultType' }
    $upnCol = if ($items[0].PSObject.Properties['UserPrincipalName']) { 'UserPrincipalName' } else { 'UPN' }

    # Aggregate failed sign-ins by user (one finding per user with count)
    $failedByUser = @{}
    foreach ($log in $items) {
        $risk = if ($riskCol) { $log.$riskCol } else { '' }
        $status = $log.$statusCol
        $upn = $log.$upnCol
        if (-not $upn) { continue }
        if ($risk -match 'High|high') {
            $findings += [PSCustomObject]@{ Type = 'high_risk_signin'; Severity = 'High'; Detail = "High-risk sign-in: $upn"; Source = $upn }
        }
        if ($status -match 'Failure|Failed|0') {
            if (-not $failedByUser.ContainsKey($upn)) { $failedByUser[$upn] = 0 }
            $failedByUser[$upn]++
        }
    }
    foreach ($upn in $failedByUser.Keys) {
        $count = $failedByUser[$upn]
        if ($count -lt $FailedSignInThreshold) { continue }
        $severity = if ($count -ge 20) { 'High' } elseif ($count -ge 10) { 'Medium' } else { 'Low' }
        $findings += [PSCustomObject]@{ Type = 'failed_signin'; Severity = $severity; Detail = "User had $count failed sign-ins in period: $upn"; Source = $upn }
    }
    return @{ Findings = $findings; Count = $findings.Count }
}

$script:SeverityWeights = @{ High = 10; Medium = 5; Low = 2 }

function Get-ReportRiskScore {
    param(
        [Parameter(Mandatory=$true)]
        [array]$Findings,
        [Parameter(Mandatory=$false)]
        [hashtable]$Weights = $script:SeverityWeights
    )
    $score = 0
    $breakdown = @{ High = 0; Medium = 0; Low = 0 }
    foreach ($f in $Findings) {
        $sev = if ($f.Severity) { $f.Severity } else { 'Low' }
        # Reduce weight for external_message_spike to avoid inflating score
        $w = if ($f.Type -eq 'external_message_spike') { 1 } elseif ($Weights[$sev]) { $Weights[$sev] } else { 2 }
        $score += $w
        if ($breakdown.ContainsKey($sev)) { $breakdown[$sev]++ }
    }
    $level = if ($score -ge 30) { 'HIGH' } elseif ($score -ge 10) { 'MEDIUM' } else { 'LOW' }
    return @{ Score = $score; Level = $level; Breakdown = $breakdown }
}

function Get-ReportTemplateSummary {
    param(
        [Parameter(Mandatory=$true)]
        [array]$Findings,
        [Parameter(Mandatory=$true)]
        [hashtable]$RiskScore,
        [Parameter(Mandatory=$false)]
        [string]$Company = 'Organization',
        [Parameter(Mandatory=$false)]
        [string]$Timestamp = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),
        [Parameter(Mandatory=$false)]
        [hashtable]$UserSummaries = @{},
        [Parameter(Mandatory=$false)]
        [bool]$HasUAL = $false
    )
    $highCount = ($Findings | Where-Object { $_.Severity -eq 'High' }).Count
    $medCount = ($Findings | Where-Object { $_.Severity -eq 'Medium' }).Count
    $lowCount = ($Findings | Where-Object { $_.Severity -eq 'Low' }).Count
    $summary = @"
# Automated Security Analysis Summary

**Generated:** $Timestamp
**Organization:** $Company
**Risk Level:** $($RiskScore.Level) (Score: $($RiskScore.Score))

## Findings Summary
- **High Severity:** $highCount
- **Medium Severity:** $medCount
- **Low Severity:** $lowCount
- **Total:** $($Findings.Count)

## BEC / Anomalous Login Triage Checklist
Use this checklist when investigating impossible travel, anomalous logins, or potential BEC alerts. Answer each item from the data to determine false positive vs. authorized vs. true positive.

| Check | Data Source | Notes |
|-------|-------------|-------|
| Forwarding rules present? | InboxRules.csv | Look for ForwardTo with external addresses |
| Rule changes in alert window? | UnifiedAuditLogs.csv | Search Operations/AuditData for: New-InboxRule, Set-InboxRule, Set-MailboxForwarding |
| Unusual external send volume? | MessageTrace.csv | Compare to user baseline |
| Sign-in from new country? | SignInLogs.csv | Check CountryOrRegion, Location |
| MFA used on sign-in? | SignInLogs.csv | Check AuthMethods column |
| High-risk or failed sign-ins? | SignInLogs.csv | RiskLevelDuringSignIn, Status |
| UAL available for search? | UnifiedAuditLogs.csv | $(if ($HasUAL) { 'Yes - search AuditData for operation details' } else { 'No - enable UAL collection for full triage' }) |

## Per-User Investigation Summary
"@
    $userList = @($UserSummaries.Keys | Sort-Object)
    $shown = 0
    $maxUsers = 50
    foreach ($upn in $userList) {
        if ($shown -ge $maxUsers) {
            $summary += "`n`n... and $($userList.Count - $maxUsers) more users (see data sources for full list)"
            break
        }
        $us = $UserSummaries[$upn]
        $summary += "`n`n### $upn"
        $summary += "`n- **Sign-in locations:** $($us.SignInLocations)"
        $summary += "`n- **MFA status:** $($us.MfaStatus)"
        $summary += "`n- **Forwarding rules:** $($us.ForwardingRules)"
        $summary += "`n- **External sends (period):** $($us.ExternalSends)"
        $summary += "`n- **High-risk sign-in:** $($us.HighRiskSignIn)"
        $summary += "`n- **Failed sign-ins:** $($us.FailedSignInCount)"
        if ($HasUAL -and $us.UALBECCount -gt 0) {
            $summary += "`n- **UAL BEC-relevant ops:** $($us.UALBECOperations)"
        }
        $shown++
    }
    if ($userList.Count -eq 0) {
        $summary += "`nNo user data available. Ensure SignInLogs, InboxRules, MessageTrace, and/or UserSecurityPosture are collected."
    }

    $summary += "`n`n## Findings by Category`n"
    $byType = $Findings | Group-Object -Property Type | ForEach-Object {
        $worst = 2
        foreach ($g in $_.Group) { if ($g.Severity -eq 'High') { $worst = 0; break } elseif ($g.Severity -eq 'Medium' -and $worst -gt 1) { $worst = 1 } }
        [PSCustomObject]@{ Grp = $_; Worst = $worst }
    } | Sort-Object Worst | ForEach-Object { $_.Grp }
    foreach ($grp in $byType) {
        $count = $grp.Count
        $summary += "`n`n### $($grp.Name) ($count)"
        $topInGroup = $grp.Group | Sort-Object { @{ High=0; Medium=1; Low=2 }[$_.Severity] }, { $_.Detail } | Select-Object -First 5
        foreach ($f in $topInGroup) {
            $rem = if ($script:RemediationByType[$f.Type]) { " | Remediation: $($script:RemediationByType[$f.Type])" } else { '' }
            $summary += "`n- [$($f.Severity)] $($f.Detail)$rem"
        }
        if ($count -gt 5) { $summary += "`n- ... and $($count - 5) more" }
    }

    $summary += "`n`n## Top Findings (by severity)"
    $top = $Findings | Sort-Object { @{ High=0; Medium=1; Low=2 }[$_.Severity] }, { $_.Detail } | Select-Object -First 20
    foreach ($f in $top) {
        $rem = if ($script:RemediationByType[$f.Type]) { " | $($script:RemediationByType[$f.Type])" } else { '' }
        $summary += "`n- [$($f.Severity)] $($f.Type): $($f.Detail)$rem"
    }
    if ($Findings.Count -gt 20) {
        $summary += "`n- ... and $($Findings.Count - 20) more (see Findings.csv)"
    }
    return $summary
}

function Get-ReportFolderCsvPath {
    param([string]$Folder, [string]$BaseName)
    if (-not $Folder -or -not (Test-Path $Folder)) { return $null }
    $candidates = @()
    $candidates += Join-Path $Folder "$BaseName.csv"
    Get-ChildItem -Path $Folder -Filter "${BaseName}*.csv" -ErrorAction SilentlyContinue | ForEach-Object { $candidates += $_.FullName }
    foreach ($p in $candidates) {
        if (Test-Path $p) {
            $lines = (Get-Content $p -TotalCount 2 -ErrorAction SilentlyContinue)
            if ($lines -and $lines.Count -ge 2) { return $p }
        }
    }
    return $null
}

# BEC-relevant UAL operations for triage (rule changes, forwarding, mailbox access)
$script:UALBECOperations = @('New-InboxRule', 'Set-InboxRule', 'Remove-InboxRule', 'Set-MailboxForwarding', 'UpdateInboxRules', 'MailboxLogin', 'MailSend')

function Get-UALBECOperationsByUser {
    param(
        [array]$UAL,
        [string]$UALCsvPath
    )
    $byUser = @{}
    $items = $null
    if ($UAL -and $UAL.Count -gt 0) {
        $items = $UAL
    } elseif ($UALCsvPath -and (Test-Path $UALCsvPath)) {
        try { $items = Import-Csv -Path $UALCsvPath -ErrorAction Stop } catch { return $byUser }
    }
    if (-not $items -or $items.Count -eq 0) { return $byUser }

    foreach ($r in $items) {
        $ops = if ($r.Operations) { $r.Operations } else { '' }
        $op = if ($r.Operation) { $r.Operation } else { '' }
        $auditData = if ($r.AuditData) { $r.AuditData } else { '' }
        $combined = "$ops $op $auditData"
        $isBEC = $false
        foreach ($becOp in $script:UALBECOperations) {
            if ($combined -match [regex]::Escape($becOp)) { $isBEC = $true; break }
        }
        if (-not $isBEC) { continue }

        $userIds = if ($r.UserIds) { ($r.UserIds -split ';') | ForEach-Object { $_.Trim() } | Where-Object { $_ } } else { @() }
        $mailboxUpn = if ($r.MailboxOwnerUPN) { $r.MailboxOwnerUPN } else { $null }
        $creationDate = if ($r.CreationDate) { $r.CreationDate } else { '' }
        $recordType = if ($r.RecordType) { $r.RecordType } else { '' }

        $targetUsers = @()
        if ($userIds -and $userIds.Count -gt 0) { $targetUsers += $userIds }
        if ($mailboxUpn -and $targetUsers -notcontains $mailboxUpn) { $targetUsers += $mailboxUpn }
        if ($targetUsers.Count -eq 0) { $targetUsers = @('Unknown') }

        foreach ($u in $targetUsers) {
            if (-not $byUser.ContainsKey($u)) { $byUser[$u] = [System.Collections.ArrayList]::new() }
            $desc = "$op"
            if ($recordType) { $desc += " ($recordType)" }
            if ($creationDate) { $desc += " @ $creationDate" }
            [void]$byUser[$u].Add($desc)
        }
    }
    return $byUser
}

function Get-UserInvestigationSummaries {
    param(
        [array]$SignInLogs,
        [string]$SignInLogsPath,
        [array]$InboxRules,
        [string]$InboxRulesPath,
        [array]$MessageTrace,
        [string]$MessageTracePath,
        [hashtable]$UALBECByUser,
        [array]$UserSecurityPosture,
        [string]$UserSecurityPosturePath
    )
    $summaries = @{}
    $allUsers = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)

    # Collect users from SignInLogs
    $slItems = $null
    if ($SignInLogs -and $SignInLogs.Count -gt 0) { $slItems = $SignInLogs }
    elseif ($SignInLogsPath -and (Test-Path $SignInLogsPath)) { try { $slItems = Import-Csv -Path $SignInLogsPath -ErrorAction Stop } catch {} }
    if ($slItems) {
        $upnCol = if ($slItems[0].PSObject.Properties['UserPrincipalName']) { 'UserPrincipalName' } else { 'UPN' }
        foreach ($s in $slItems) {
            $upn = $s.$upnCol
            if ($upn) { [void]$allUsers.Add($upn) }
        }
    }

    # Collect users from InboxRules
    $irItems = $null
    if ($InboxRules -and $InboxRules.Count -gt 0) { $irItems = $InboxRules }
    elseif ($InboxRulesPath -and (Test-Path $InboxRulesPath)) { try { $irItems = Import-Csv -Path $InboxRulesPath -ErrorAction Stop } catch {} }
    if ($irItems) {
        $mbxCol = if ($irItems[0].PSObject.Properties['MailboxOwner']) { 'MailboxOwner' } else { $null }
        if ($mbxCol) {
            foreach ($r in $irItems) {
                $mbx = $r.$mbxCol
                if ($mbx) { [void]$allUsers.Add($mbx) }
            }
        }
    }

    # Collect users from MessageTrace
    $mtItems = $null
    if ($MessageTrace -and $MessageTrace.Count -gt 0) { $mtItems = $MessageTrace }
    elseif ($MessageTracePath -and (Test-Path $MessageTracePath)) { try { $mtItems = Import-Csv -Path $MessageTracePath -ErrorAction Stop } catch {} }
    if ($mtItems) {
        $senderCol = if ($mtItems[0].PSObject.Properties['SenderAddress']) { 'SenderAddress' } else { 'From' }
        foreach ($t in $mtItems) {
            $s = $t.$senderCol
            if ($s) { [void]$allUsers.Add($s) }
        }
    }

    # Add UAL users
    if ($UALBECByUser) {
        foreach ($u in $UALBECByUser.Keys) {
            if ($u -and $u -ne 'Unknown') { [void]$allUsers.Add($u) }
        }
    }

    # Collect users from UserSecurityPosture
    $uspItems = $null
    if ($UserSecurityPosture -and $UserSecurityPosture.Count -gt 0) { $uspItems = $UserSecurityPosture }
    elseif ($UserSecurityPosturePath -and (Test-Path $UserSecurityPosturePath)) { try { $uspItems = Import-Csv -Path $UserSecurityPosturePath -ErrorAction Stop } catch {} }
    if ($uspItems) {
        $upnCol = if ($uspItems[0].PSObject.Properties['UserPrincipalName']) { 'UserPrincipalName' } else { 'UPN' }
        foreach ($u in $uspItems) {
            $upn = $u.$upnCol
            if ($upn) { [void]$allUsers.Add($upn) }
        }
    }

    # Build MFA map
    $mfaMap = @{}
    if ($uspItems) {
        $mfaCol = $null
        if ($uspItems[0].PSObject.Properties['MfaCovered']) { $mfaCol = 'MfaCovered' }
        elseif ($uspItems[0].PSObject.Properties['PerUserMfaStatus']) { $mfaCol = 'PerUserMfaStatus' }
        if ($mfaCol) {
            foreach ($u in $uspItems) {
                $upn = if ($u.UserPrincipalName) { $u.UserPrincipalName } elseif ($u.UPN) { $u.UPN } else { $null }
                $val = $u.$mfaCol
                $mfaMap[$upn] = $val -eq $true -or $val -eq 'True' -or $val -match 'Enabled|Covered'
            }
        }
    }

    foreach ($upn in $allUsers) {
        if (-not $upn -or $upn -eq 'Unknown') { continue }
        $loc = @()
        $highRisk = $false
        $failedCount = 0
        if ($slItems) {
            $upnCol = if ($slItems[0].PSObject.Properties['UserPrincipalName']) { 'UserPrincipalName' } else { 'UPN' }
            $statusCol = if ($slItems[0].PSObject.Properties['Status']) { 'Status' } else { 'ResultType' }
            $riskCol = $null
            foreach ($p in $slItems[0].PSObject.Properties.Name) { if ($p -match 'RiskLevel|Risk') { $riskCol = $p; break } }
            $locCol = if ($slItems[0].PSObject.Properties['CountryOrRegion']) { 'CountryOrRegion' } else { 'Location' }
            foreach ($s in $slItems) {
                if (($s.$upnCol -ne $upn)) { continue }
                $c = if ($locCol -and $s.$locCol) { $s.$locCol } else { $null }
                if ($c -and $loc -notcontains $c) { $loc += $c }
                if ($riskCol -and $s.$riskCol -match 'High|high') { $highRisk = $true }
                if ($s.$statusCol -match 'Failure|Failed|0') { $failedCount++ }
            }
        }
        $forwardingRules = 0
        if ($irItems) {
            $mbxCol = if ($irItems[0].PSObject.Properties['MailboxOwner']) { 'MailboxOwner' } else { $null }
            $ftCol = if ($irItems[0].PSObject.Properties['ForwardTo']) { 'ForwardTo' } else { $null }
            if ($mbxCol -and $ftCol) {
                foreach ($r in $irItems) {
                    if ($r.$mbxCol -eq $upn -and $r.$ftCol -match '@') { $forwardingRules++ }
                }
            }
        }
        $externalSends = 0
        if ($mtItems) {
            $senderCol = if ($mtItems[0].PSObject.Properties['SenderAddress']) { 'SenderAddress' } else { 'From' }
            $recCol = $null
            foreach ($p in $mtItems[0].PSObject.Properties.Name) { if ($p -match 'Recipient|To|RecipientAddress') { $recCol = $p; break } }
            if ($senderCol -and $recCol) {
                foreach ($t in $mtItems) {
                    if ($t.$senderCol -ne $upn) { continue }
                    $rec = $t.$recCol
                    if ($rec -match '@' -and $rec -notmatch '\.onmicrosoft\.com|\.mail\.protection\.outlook') { $externalSends++ }
                }
            }
        }
        $ualBECOps = @()
        if ($UALBECByUser -and $UALBECByUser[$upn]) {
            $ualBECOps = @($UALBECByUser[$upn] | Select-Object -Unique)
        }
        $mfa = if ($mfaMap[$upn]) { 'Yes' } else { 'Unknown' }

        $summaries[$upn] = [PSCustomObject]@{
            User = $upn
            SignInLocations = ($loc | Select-Object -Unique) -join '; '
            HighRiskSignIn = $highRisk
            FailedSignInCount = $failedCount
            ForwardingRules = $forwardingRules
            ExternalSends = $externalSends
            UALBECOperations = ($ualBECOps | Select-Object -First 10) -join '; '
            UALBECCount = $ualBECOps.Count
            MfaStatus = $mfa
        }
    }
    return $summaries
}

function Get-MailboxRecipientTypeMap {
    param(
        [array]$UserPosture,
        [string]$UspCsvPath
    )
    $map = @{}
    if ($UserPosture -and $UserPosture.Count -gt 0) {
        $upnCol = if ($UserPosture[0].PSObject.Properties['UserPrincipalName']) { 'UserPrincipalName' } else { 'UPN' }
        $rtCol = if ($UserPosture[0].PSObject.Properties['RecipientType']) { 'RecipientType' } else { $null }
        if ($rtCol) {
            foreach ($u in $UserPosture) {
                $upn = $u.$upnCol
                $rt = $u.$rtCol
                if ($upn) { $map[$upn] = $rt }
            }
        }
    }
    if ($UspCsvPath -and (Test-Path $UspCsvPath)) {
        try {
            $rows = Import-Csv -Path $UspCsvPath -ErrorAction Stop
            foreach ($r in $rows) {
                $upn = if ($r.UserPrincipalName) { $r.UserPrincipalName } elseif ($r.UPN) { $r.UPN } else { $null }
                $rt = if ($r.RecipientType) { $r.RecipientType } else { $null }
                if ($upn -and $rt) { $map[$upn] = $rt }
            }
        } catch {}
    }
    return $map
}

function Get-FindingsWithDeduplication {
    param([array]$Findings)
    $seen = @{}
    $deduped = @()
    foreach ($f in $Findings) {
        $key = "$($f.Type)|$($f.Source)|$($f.Detail)"
        if ($seen[$key]) { continue }
        $seen[$key] = $true
        $deduped += $f
    }
    return $deduped
}

function Get-SecurityIntegrationFindings {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FolderPath
    )

    $findings = [System.Collections.Generic.List[object]]::new()
    if (-not (Test-Path -LiteralPath $FolderPath)) { return @() }

    $patterns = @(
        @{ Glob = 'HuntressSignals*.csv'; Label = 'Huntress signals'; Type = 'huntress_signals' }
        @{ Glob = 'HuntressIncidents*.csv'; Label = 'Huntress incidents'; Type = 'huntress_incidents' }
        @{ Glob = 'HuntressAgents*.csv'; Label = 'Huntress agents'; Type = 'huntress_agents' }
        @{ Glob = 'S1Threats*.csv'; Label = 'SentinelOne threats'; Type = 's1_threats' }
        @{ Glob = 'S1Agents*.csv'; Label = 'SentinelOne agents'; Type = 's1_agents' }
        @{ Glob = 'LiongardDetections*.csv'; Label = 'Liongard detections'; Type = 'liongard_detections' }
    )

    foreach ($p in $patterns) {
        $files = Get-ChildItem -LiteralPath $FolderPath -Filter $p.Glob -File -ErrorAction SilentlyContinue
        foreach ($f in $files) {
            try {
                $rows = @(Import-Csv -LiteralPath $f.FullName -ErrorAction Stop)
                $count = $rows.Count
                if ($count -gt 0) {
                    [void]$findings.Add([PSCustomObject]@{
                        Type     = $p.Type
                        Severity = if ($p.Type -match 'threat|incident|detection') { 'Medium' } else { 'Low' }
                        Detail   = "$($p.Label): $count row(s) in $($f.Name)"
                        Source   = $f.Name
                    })
                }
            } catch {}
        }
    }

    return @($findings)
}

function Get-ReportFindings {
    param(
        [Parameter(Mandatory=$false)]
        [object]$Report,
        [Parameter(Mandatory=$false)]
        [string]$FolderPath,
        [Parameter(Mandatory=$false)]
        [array]$SuspiciousKeywords = $script:DefaultSuspiciousKeywords,
        [Parameter(Mandatory=$false)]
        [int]$StaleDataDaysThreshold = 7
    )
    $allFindings = @()
    $missingData = @()
    $staleData = @()
    $ual = $null
    $ualPath = $null
    $usp = $null
    $uspPath = $null
    $inboxRules = $null
    $inboxPath = $null
    $signInLogs = $null
    $slPath = $null
    $messageTrace = $null
    $mtPath = $null

    if ($Report) {
        $usp = $Report.UserSecurityPosture
        if (-not $usp -and $Report.MfaCoverage -and $Report.MfaCoverage.Users) {
            $usp = $Report.MfaCoverage.Users | ForEach-Object { [PSCustomObject]@{ UserPrincipalName = $_.UserPrincipalName; MfaCovered = $_.MfaCovered } }
        }
        $mbxMap = Get-MailboxRecipientTypeMap -UserPosture $usp
        $ir = Get-InboxRuleFindings -Rules $Report.InboxRules -SuspiciousKeywords $SuspiciousKeywords -MailboxRecipientTypeMap $mbxMap
        $allFindings += $ir.Findings
        $tr = Get-TransportRuleFindings -Rules $Report.TransportRules
        $allFindings += $tr.Findings
        $ar = Get-AppRegistrationFindings -Apps $Report.AppRegistrations
        $allFindings += $ar.Findings
        $ca = Get-CAPolicyFindings -Policies $Report.ConditionalAccessPolicies
        $allFindings += $ca.Findings
        $mfa = Get-MfaGapFindings -UserPosture $usp
        $allFindings += $mfa.Findings
        $mt = Get-MessageTraceFindings -Trace $Report.MessageTrace
        $allFindings += $mt.Findings
        $sl = Get-SignInLogFindings -Logs $Report.SignInLogs
        $allFindings += $sl.Findings
        $ual = $Report.UnifiedAuditLogs
        $inboxRules = $Report.InboxRules
        $signInLogs = $Report.SignInLogs
        $messageTrace = $Report.MessageTrace
        if (-not $Report.SignInLogs -or $Report.SignInLogs.Count -eq 0) { $missingData += 'SignInLogs' }
        if (-not $Report.MessageTrace -or $Report.MessageTrace.Count -eq 0) { $missingData += 'MessageTrace' }
        if (-not $ual -or $ual.Count -eq 0) { $missingData += 'UnifiedAuditLogs' }
    } elseif ($FolderPath -and (Test-Path $FolderPath)) {
        $uspPath = Get-ReportFolderCsvPath -Folder $FolderPath -BaseName 'UserSecurityPosture'
        $mbxMap = Get-MailboxRecipientTypeMap -UspCsvPath $uspPath
        $inboxPath = Get-ReportFolderCsvPath -Folder $FolderPath -BaseName 'InboxRules'
        if ($inboxPath) { $allFindings += (Get-InboxRuleFindings -CsvPath $inboxPath -SuspiciousKeywords $SuspiciousKeywords -MailboxRecipientTypeMap $mbxMap).Findings }
        $transPath = Get-ReportFolderCsvPath -Folder $FolderPath -BaseName 'TransportRules'
        if ($transPath) { $allFindings += (Get-TransportRuleFindings -CsvPath $transPath).Findings }
        $appPath = Get-ReportFolderCsvPath -Folder $FolderPath -BaseName 'AppRegistrations'
        if ($appPath) { $allFindings += (Get-AppRegistrationFindings -CsvPath $appPath).Findings }
        $caPath = Get-ReportFolderCsvPath -Folder $FolderPath -BaseName 'ConditionalAccessPolicies'
        if ($caPath) { $allFindings += (Get-CAPolicyFindings -CsvPath $caPath).Findings }
        if ($uspPath) { $allFindings += (Get-MfaGapFindings -CsvPath $uspPath).Findings }
        $mtPath = Get-ReportFolderCsvPath -Folder $FolderPath -BaseName 'MessageTrace'
        if ($mtPath) {
            $allFindings += (Get-MessageTraceFindings -CsvPath $mtPath).Findings
            $mtFile = Get-Item -Path $mtPath -ErrorAction SilentlyContinue
            if ($mtFile -and $mtFile.LastWriteTime -lt (Get-Date).AddDays(-$StaleDataDaysThreshold)) { $staleData += 'MessageTrace' }
        } else { $missingData += 'MessageTrace' }
        $slPath = Get-ReportFolderCsvPath -Folder $FolderPath -BaseName 'SignInLogs'
        if ($slPath) {
            $allFindings += (Get-SignInLogFindings -CsvPath $slPath).Findings
            $slFile = Get-Item -Path $slPath -ErrorAction SilentlyContinue
            if ($slFile -and $slFile.LastWriteTime -lt (Get-Date).AddDays(-$StaleDataDaysThreshold)) { $staleData += 'SignInLogs' }
        } else { $missingData += 'SignInLogs' }
        $ualPath = Get-ReportFolderCsvPath -Folder $FolderPath -BaseName 'UnifiedAuditLogs'
        if (-not $ualPath) { $missingData += 'UnifiedAuditLogs' }
        $allFindings += Get-SecurityIntegrationFindings -FolderPath $FolderPath
    }

    $hasUAL = ($ual -and $ual.Count -gt 0) -or ($ualPath -and (Test-Path $ualPath))
    $ualBECByUser = Get-UALBECOperationsByUser -UAL $ual -UALCsvPath $ualPath
    $userSummaries = Get-UserInvestigationSummaries -SignInLogs $signInLogs -SignInLogsPath $slPath -InboxRules $inboxRules -InboxRulesPath $inboxPath -MessageTrace $messageTrace -MessageTracePath $mtPath -UALBECByUser $ualBECByUser -UserSecurityPosture $usp -UserSecurityPosturePath $uspPath

    foreach ($m in ($missingData | Select-Object -Unique)) {
        $allFindings += [PSCustomObject]@{ Type = 'missing_data'; Severity = 'Low'; Detail = "Missing data: $m - download or collect for complete analysis"; Source = $m }
    }
    foreach ($s in ($staleData | Select-Object -Unique)) {
        $allFindings += [PSCustomObject]@{ Type = 'stale_data'; Severity = 'Low'; Detail = "Data may be outdated: $s - consider re-running collection"; Source = $s }
    }

    $allFindings = Get-FindingsWithDeduplication -Findings $allFindings
    $riskScore = Get-ReportRiskScore -Findings $allFindings
    $company = if ($Report -and $Report.Company) { $Report.Company } else { 'Organization' }
    $timestamp = if ($Report -and $Report.Timestamp) { $Report.Timestamp } else { (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') }
    $summary = Get-ReportTemplateSummary -Findings $allFindings -RiskScore $riskScore -Company $company -Timestamp $timestamp -UserSummaries $userSummaries -HasUAL $hasUAL
    return @{
        Findings = $allFindings
        RiskScore = $riskScore
        Summary = $summary
        UserSummaries = $userSummaries
        HasUAL = $hasUAL
    }
}

function Invoke-ReportFolderAnalysis {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Path,
        [Parameter(Mandatory=$false)]
        [switch]$WriteOutputFiles
    )
    $result = Get-ReportFindings -FolderPath $Path
    if ($WriteOutputFiles -and $Path -and (Test-Path $Path)) {
        $findingsPath = Join-Path $Path 'Findings.csv'
        $summaryPath = Join-Path $Path '_Automated_Summary.txt'
        $jsonPath = Join-Path $Path 'Findings.json'
        if ($result.Findings -and $result.Findings.Count -gt 0) {
            $result.Findings | Export-Csv -Path $findingsPath -NoTypeInformation -Encoding UTF8
        }
        $result.Summary | Out-File -Path $summaryPath -Encoding utf8
        $jsonObj = @{ Findings = $result.Findings; RiskScore = $result.RiskScore; HasUAL = $result.HasUAL }
        if ($result.UserSummaries) {
            $userSummariesForJson = @{}
            foreach ($k in $result.UserSummaries.Keys) {
                $userSummariesForJson[$k] = $result.UserSummaries[$k]
            }
            $jsonObj['UserSummaries'] = $userSummariesForJson
        }
        $jsonObj | ConvertTo-Json -Depth 5 | Out-File -Path $jsonPath -Encoding utf8
    }
    return $result
}

function Get-BulkTenantAnalysis {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ParentFolder,
        [Parameter(Mandatory=$false)]
        [switch]$WriteOutputFiles
    )
    if (-not (Test-Path $ParentFolder)) { return @() }
    $tenantFolders = Get-ChildItem -Path $ParentFolder -Directory -ErrorAction SilentlyContinue
    $results = @()
    foreach ($t in $tenantFolders) {
        $runs = Get-ChildItem -Path $t.FullName -Directory -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending
        $targetFolder = if ($runs) { $runs[0].FullName } else { $t.FullName }
        $analysis = Get-ReportFindings -FolderPath $targetFolder
        $topFinding = ''
        if ($analysis.Findings -and $analysis.Findings.Count -gt 0) {
            $top = $analysis.Findings | Sort-Object { @{ High=0; Medium=1; Low=2 }[$_.Severity] } | Select-Object -First 1
            $topFinding = $top.Detail
        }
        $results += [PSCustomObject]@{
            Tenant = $t.Name
            Folder = $targetFolder
            RiskScore = $analysis.RiskScore.Score
            RiskLevel = $analysis.RiskScore.Level
            FindingCount = $analysis.Findings.Count
            TopFinding = $topFinding
        }
    }
    $results = $results | Sort-Object RiskScore -Descending
    if ($WriteOutputFiles) {
        $rankingPath = Join-Path $ParentFolder 'BulkTenantRanking.csv'
        $results | Export-Csv -Path $rankingPath -NoTypeInformation -Encoding UTF8
    }
    return $results
}

Export-ModuleMember -Function Get-InboxRuleFindings, Get-TransportRuleFindings, Get-AppRegistrationFindings, Get-CAPolicyFindings
Export-ModuleMember -Function Get-MfaGapFindings, Get-MessageTraceFindings, Get-SignInLogFindings, Get-SecurityIntegrationFindings
Export-ModuleMember -Function Get-ReportRiskScore, Get-ReportTemplateSummary, Get-ReportFindings
Export-ModuleMember -Function Invoke-ReportFolderAnalysis, Get-BulkTenantAnalysis
