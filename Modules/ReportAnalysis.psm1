# ReportAnalysis.psm1
# Rule-based analysis of security investigation reports - reduces LLM reliance
# Supports both in-memory Report objects and folder-based CSV analysis

$script:DefaultSuspiciousKeywords = @(
    'invoice','payment','bank','wire','transfer','refund','urgent','verify','confirm',
    'password','credential','login','account','suspicious','forward','external'
)

function Get-InboxRuleFindings {
    param(
        [Parameter(Mandatory=$false)]
        [array]$Rules,
        [Parameter(Mandatory=$false)]
        [string]$CsvPath,
        [Parameter(Mandatory=$false)]
        [array]$SuspiciousKeywords = $script:DefaultSuspiciousKeywords
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

        if ($forwardTo -match '@' -and $forwardTo -notmatch ';') {
            $findings += [PSCustomObject]@{ Type = 'suspicious_inbox_rule'; Severity = 'High'; Detail = "External forwarding: $mailbox -> $forwardTo"; Source = $name }
        } elseif ($forwardTo -match '@') {
            $findings += [PSCustomObject]@{ Type = 'suspicious_inbox_rule'; Severity = 'Medium'; Detail = "Multiple external forwards: $mailbox"; Source = $name }
        }
        if ($isHidden -and $name -and $name -notmatch 'system|default|outlook|microsoft|junk|clutter|archive') {
            $findings += [PSCustomObject]@{ Type = 'hidden_inbox_rule'; Severity = 'Medium'; Detail = "Hidden rule: $mailbox"; Source = $name }
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
    $riskyPatterns = @('RedirectMessageTo|ForwardTo|RedirectTo', 'BypassSpoofing', 'SetHeader|ModifyHeader', 'Quarantine')
    foreach ($r in $items) {
        $actions = if ($r.ActionsSummary) { $r.ActionsSummary } else { '' }
        $conditions = if ($r.ConditionsSummary) { $r.ConditionsSummary } else { '' }
        $name = if ($r.Name) { $r.Name } else { 'Unknown' }
        $combined = "$actions $conditions"
        if ($combined -match 'ForwardTo|RedirectMessageTo|RedirectTo') {
            $findings += [PSCustomObject]@{ Type = 'transport_forward'; Severity = 'High'; Detail = "Transport rule forwards/redirects mail"; Source = $name }
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

    $findings = @()
    foreach ($a in $items) {
        $riskLevel = if ($a.RiskLevel) { $a.RiskLevel } else { '' }
        $name = if ($a.DisplayName) { $a.DisplayName } else { 'Unknown' }
        $publisher = if ($a.PublisherDomain) { $a.PublisherDomain } else { '' }
        $userConsent = if ($a.HasUserConsent) { $a.HasUserConsent -eq $true -or $a.HasUserConsent -eq 'True' } else { $false }

        if ($riskLevel -eq 'High') {
            $findings += [PSCustomObject]@{ Type = 'high_risk_app'; Severity = 'High'; Detail = "High-risk app: $name"; Source = $publisher }
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
        $requiresMfa = if ($p.RequiresMfa) { $p.RequiresMfa -eq $true -or $p.RequiresMfa -eq 'True' } else { $false }
        $userAll = if ($p.UserIncludeAll) { $p.UserIncludeAll -eq $true -or $p.UserIncludeAll -eq 'True' } else { $false }

        if ($riskLevel -eq 'High') {
            $findings += [PSCustomObject]@{ Type = 'high_risk_ca_policy'; Severity = 'High'; Detail = "High-risk CA policy: $name"; Source = $name }
        }
        if ($userAll -and -not $requiresMfa) {
            $findings += [PSCustomObject]@{ Type = 'ca_no_mfa_all_users'; Severity = 'High'; Detail = "CA applies to all users without MFA: $name"; Source = $name }
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
        [int]$ExternalPercentileThreshold = 95
    )
    $items = $null
    if ($Trace -and $Trace.Count -gt 0) {
        $items = $Trace
    } elseif ($CsvPath -and (Test-Path $CsvPath)) {
        try { $items = Import-Csv -Path $CsvPath -ErrorAction Stop } catch { return @{ Findings = @(); Count = 0 } }
    }
    if (-not $items -or $items.Count -eq 0) { return @{ Findings = @(); Count = 0 } }

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
        $sender = $t.$senderCol
        if (-not $sender) { continue }
        if (-not $externalBySender.ContainsKey($sender)) { $externalBySender[$sender] = 0 }
        if ($recipients -match '@' -and $recipients -notmatch '\.onmicrosoft\.com|\.mail\.protection\.outlook') {
            $externalBySender[$sender]++
        }
    }
    if ($externalBySender.Count -eq 0) { return @{ Findings = @(); Count = 0 } }
    $counts = @($externalBySender.Values | Where-Object { $_ -gt 0 })
    if ($counts.Count -eq 0) { return @{ Findings = @(); Count = 0 } }
    $threshold = [Math]::Max(1, [int]([Math]::Ceiling($counts.Count * $ExternalPercentileThreshold / 100)))
    $sorted = $counts | Sort-Object -Descending
    $percentileVal = if ($threshold -le $sorted.Count) { $sorted[$threshold - 1] } else { $sorted[-1] }
    foreach ($sender in $externalBySender.Keys) {
        if ($externalBySender[$sender] -ge $percentileVal -and $externalBySender[$sender] -gt 5) {
            $findings += [PSCustomObject]@{ Type = 'external_message_spike'; Severity = 'Low'; Detail = "High external message volume: $sender ($($externalBySender[$sender]) external)"; Source = $sender }
        }
    }
    return @{ Findings = $findings; Count = $findings.Count }
}

function Get-SignInLogFindings {
    param(
        [Parameter(Mandatory=$false)]
        [array]$Logs,
        [Parameter(Mandatory=$false)]
        [string]$CsvPath
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
    foreach ($log in $items) {
        $risk = if ($riskCol) { $log.$riskCol } else { '' }
        $status = $log.$statusCol
        $upn = $log.$upnCol
        if ($risk -match 'High|high') {
            $findings += [PSCustomObject]@{ Type = 'high_risk_signin'; Severity = 'High'; Detail = "High-risk sign-in: $upn"; Source = $upn }
        }
        if ($status -match 'Failure|Failed|0') {
            $findings += [PSCustomObject]@{ Type = 'failed_signin'; Severity = 'Low'; Detail = "Failed sign-in: $upn"; Source = $upn }
        }
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
        $w = if ($Weights[$sev]) { $Weights[$sev] } else { 2 }
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
        [string]$Timestamp = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
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

## Top Findings
"@
    $top = $Findings | Sort-Object { @{ H=0; M=1; L=2 }[$_.Severity] }, { $_.Detail } | Select-Object -First 15
    foreach ($f in $top) {
        $summary += "`n- [$($f.Severity)] $($f.Type): $($f.Detail)"
    }
    if ($Findings.Count -gt 15) {
        $summary += "`n- ... and $($Findings.Count - 15) more (see Findings.csv)"
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

function Get-ReportFindings {
    param(
        [Parameter(Mandatory=$false)]
        [object]$Report,
        [Parameter(Mandatory=$false)]
        [string]$FolderPath,
        [Parameter(Mandatory=$false)]
        [array]$SuspiciousKeywords = $script:DefaultSuspiciousKeywords
    )
    $allFindings = @()
    if ($Report) {
        $ir = Get-InboxRuleFindings -Rules $Report.InboxRules -SuspiciousKeywords $SuspiciousKeywords
        $allFindings += $ir.Findings
        $tr = Get-TransportRuleFindings -Rules $Report.TransportRules
        $allFindings += $tr.Findings
        $ar = Get-AppRegistrationFindings -Apps $Report.AppRegistrations
        $allFindings += $ar.Findings
        $ca = Get-CAPolicyFindings -Policies $Report.ConditionalAccessPolicies
        $allFindings += $ca.Findings
        $usp = $Report.UserSecurityPosture
        if (-not $usp -and $Report.MfaCoverage -and $Report.MfaCoverage.Users) {
            $usp = $Report.MfaCoverage.Users | ForEach-Object { [PSCustomObject]@{ UserPrincipalName = $_.UserPrincipalName; MfaCovered = $_.MfaCovered } }
        }
        $mfa = Get-MfaGapFindings -UserPosture $usp
        $allFindings += $mfa.Findings
        $mt = Get-MessageTraceFindings -Trace $Report.MessageTrace
        $allFindings += $mt.Findings
        $sl = Get-SignInLogFindings -Logs $Report.SignInLogs
        $allFindings += $sl.Findings
    } elseif ($FolderPath -and (Test-Path $FolderPath)) {
        $inboxPath = Get-ReportFolderCsvPath -Folder $FolderPath -BaseName 'InboxRules'
        if ($inboxPath) { $allFindings += (Get-InboxRuleFindings -CsvPath $inboxPath -SuspiciousKeywords $SuspiciousKeywords).Findings }
        $transPath = Get-ReportFolderCsvPath -Folder $FolderPath -BaseName 'TransportRules'
        if ($transPath) { $allFindings += (Get-TransportRuleFindings -CsvPath $transPath).Findings }
        $appPath = Get-ReportFolderCsvPath -Folder $FolderPath -BaseName 'AppRegistrations'
        if ($appPath) { $allFindings += (Get-AppRegistrationFindings -CsvPath $appPath).Findings }
        $caPath = Get-ReportFolderCsvPath -Folder $FolderPath -BaseName 'ConditionalAccessPolicies'
        if ($caPath) { $allFindings += (Get-CAPolicyFindings -CsvPath $caPath).Findings }
        $uspPath = Get-ReportFolderCsvPath -Folder $FolderPath -BaseName 'UserSecurityPosture'
        if ($uspPath) { $allFindings += (Get-MfaGapFindings -CsvPath $uspPath).Findings }
        $mtPath = Get-ReportFolderCsvPath -Folder $FolderPath -BaseName 'MessageTrace'
        if ($mtPath) { $allFindings += (Get-MessageTraceFindings -CsvPath $mtPath).Findings }
        $slPath = Get-ReportFolderCsvPath -Folder $FolderPath -BaseName 'SignInLogs'
        if ($slPath) { $allFindings += (Get-SignInLogFindings -CsvPath $slPath).Findings }
    }
    $riskScore = Get-ReportRiskScore -Findings $allFindings
    $company = if ($Report -and $Report.Company) { $Report.Company } else { 'Organization' }
    $timestamp = if ($Report -and $Report.Timestamp) { $Report.Timestamp } else { (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') }
    $summary = Get-ReportTemplateSummary -Findings $allFindings -RiskScore $riskScore -Company $company -Timestamp $timestamp
    return @{
        Findings = $allFindings
        RiskScore = $riskScore
        Summary = $summary
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
        @{ Findings = $result.Findings; RiskScore = $result.RiskScore } | ConvertTo-Json -Depth 5 | Out-File -Path $jsonPath -Encoding utf8
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
            $top = $analysis.Findings | Sort-Object { @{ H=0; M=1; L=2 }[$_.Severity] } | Select-Object -First 1
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
Export-ModuleMember -Function Get-MfaGapFindings, Get-MessageTraceFindings, Get-SignInLogFindings
Export-ModuleMember -Function Get-ReportRiskScore, Get-ReportTemplateSummary, Get-ReportFindings
Export-ModuleMember -Function Invoke-ReportFolderAnalysis, Get-BulkTenantAnalysis
