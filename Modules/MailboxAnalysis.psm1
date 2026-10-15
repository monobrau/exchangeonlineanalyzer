function Test-ExternalForwarding {
    <#
    .SYNOPSIS
        Tests if recipient addresses include external (non-organizational) domains.
    .DESCRIPTION
        Restored original function signature for backward compatibility.
        Use Test-InboxRuleRisk for comprehensive rule risk evaluation.
    .PARAMETER RecipientAddresses
        Array of recipient addresses to check
    .PARAMETER InternalDomains
        Array of internal/organizational domains
    #>
    param(
        [Parameter(Mandatory=$true)]
        [array]$RecipientAddresses,
        [Parameter(Mandatory=$true)]
        [array]$InternalDomains
    )
    
    if ($null -eq $RecipientAddresses -or $RecipientAddresses.Count -eq 0) { 
        return $false 
    }
    
    foreach ($recipient in $RecipientAddresses) {
        $emailAddress = $null
        
        if ($recipient -is [string]) { 
            $emailAddress = $recipient 
        }
        elseif ($recipient.PropertyNames -contains 'Address') { 
            $emailAddress = $recipient.Address 
        } 
        elseif ($recipient.PropertyNames -contains 'EmailAddress') { 
            $emailAddress = $recipient.EmailAddress 
        }
        
        if (-not [string]::IsNullOrWhiteSpace($emailAddress)) {
            # Parse email using regex to handle SMTP: prefix and EX:/o= internal recipients
            $parsedEmail = $null
            
            # Handle SMTP: prefix
            if ($emailAddress -match 'SMTP:([^\]]+)') {
                $parsedEmail = $matches[1]
            }
            # Handle EX:/o= (Exchange internal format) - treat as internal
            elseif ($emailAddress -match '^EX:/o=') {
                continue # Internal recipient
            }
            # Plain email address
            elseif ($emailAddress -match '@') {
                $parsedEmail = $emailAddress
            }
            
            if ($parsedEmail -and $parsedEmail -match '@') {
                try {
                    $domain = ($parsedEmail -split '@')[1].ToLowerInvariant()
                    if ($domain -notin $InternalDomains) { 
                        Write-Verbose "External domain detected: $domain"
                        return $true 
                    }
                } catch { 
                    Write-Verbose "Could not parse domain from: $parsedEmail" 
                }
            }
        }
    }
    
    return $false
}

function Test-InboxRuleRisk {
    <#
    .SYNOPSIS
        Evaluates inbox rule risk using scoring-based heuristics.
    .DESCRIPTION
        Analyzes inbox rules for suspicious patterns with risk scoring instead of blanket flagging.
        Handles SMTP: prefixed addresses and EX:/o= internal recipients.
    .PARAMETER Rule
        The inbox rule object to evaluate
    .PARAMETER InternalDomains
        Array of internal/organizational domains
    .OUTPUTS
        PSCustomObject with IsExternal, IsSuspicious, RiskScore, RiskLevel, Recipients, ForwardingType, SuspiciousReasons
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        $Rule,
        [Parameter(Mandatory=$true)]
        [array]$InternalDomains
    )
    
    $result = [PSCustomObject]@{
        IsExternal = $false
        IsSuspicious = $false
        RiskScore = 0
        RiskLevel = "Low"
        Recipients = @()
        ForwardingType = $null
        SuspiciousReasons = [System.Collections.ArrayList]::new()
    }
    
    # Check forwarding/redirect properties
    $recipientProperties = @(
        @{ Name = 'ForwardTo'; Type = 'Forward' },
        @{ Name = 'ForwardAsAttachmentTo'; Type = 'ForwardAsAttachment' },
        @{ Name = 'RedirectTo'; Type = 'Redirect' }
    )
    
    foreach ($prop in $recipientProperties) {
        $recipients = $Rule.($prop.Name)
        if ($null -eq $recipients -or $recipients.Count -eq 0) { continue }
        
        foreach ($recipient in $recipients) {
            $emailAddress = $null
            
            if ($recipient -is [string]) { 
                $emailAddress = $recipient 
            }
            elseif ($recipient.PropertyNames -contains 'Address') { 
                $emailAddress = $recipient.Address 
            } 
            elseif ($recipient.PropertyNames -contains 'EmailAddress') { 
                $emailAddress = $recipient.EmailAddress 
            }
            
            if (-not [string]::IsNullOrWhiteSpace($emailAddress)) {
                # Parse email using regex to handle SMTP: and EX:/o= formats
                $parsedEmail = $null
                
                # Handle SMTP:address format
                if ($emailAddress -match 'SMTP:([^\]]+)') {
                    $parsedEmail = $matches[1]
                }
                # Handle EX:/o=ExchangeLegacyDN (internal recipient)
                elseif ($emailAddress -match '^EX:/o=') {
                    Write-Verbose "Internal Exchange recipient: $emailAddress"
                    $result.Recipients += $emailAddress
                    continue # Internal, skip domain check
                }
                # Plain email
                elseif ($emailAddress -match '@') {
                    $parsedEmail = $emailAddress
                }
                
                if ($parsedEmail -and $parsedEmail -match '@') {
                    try {
                        $domain = ($parsedEmail -split '@')[1].ToLowerInvariant()
                        $result.Recipients += $parsedEmail
                        
                        if ($domain -notin $InternalDomains) {
                            $result.IsExternal = $true
                            $result.ForwardingType = $prop.Type
                            $result.RiskScore += 4
                            [void]$result.SuspiciousReasons.Add("$($prop.Type) to external domain: $parsedEmail")
                        }
                    } catch { 
                        Write-Verbose "Could not parse domain from: $parsedEmail" 
                    }
                } else {
                    # Could not parse email, add to recipients anyway
                    $result.Recipients += $emailAddress
                }
            }
        }
    }
    
    # Scoring: Check for delete actions
    if ($Rule.DeleteMessage -eq $true) {
        # Plain delete is low risk (score +1) unless combined with other factors
        $result.RiskScore += 1
        
        # Delete + keyword condition = higher risk (evidence hiding)
        if ($Rule.SubjectOrBodyContainsWords -or $Rule.FromAddressContainsWords) {
            $result.RiskScore += 3
            [void]$result.SuspiciousReasons.Add("Deletes messages matching specific patterns (possible evidence hiding)")
        } else {
            [void]$result.SuspiciousReasons.Add("Deletes messages (low risk alone)")
        }
    }
    
    # Scoring: Check for move-to-folder actions
    if ($Rule.MoveToFolder) {
        $folderPath = $Rule.MoveToFolder.ToString()
        $folderName = $folderPath -replace '.*\\', '' # Get last segment
        
        # High-risk folder names
        $highRiskFolders = @('RSS Subscriptions', 'RSS Feeds', 'Conversation History')
        $isHighRiskFolder = $false
        foreach ($folder in $highRiskFolders) {
            if ($folderPath -like "*$folder*") {
                $isHighRiskFolder = $true
                break
            }
        }
        
        # Check for non-printable or suspicious characters in folder name
        $hasNonPrintable = $folderName -match "[^\x20-\x7E]"
        $isWhitespaceOnly = $folderName -match "^\s+$"
        
        # Scoring for suspicious folder moves
        if ($isHighRiskFolder -and ($Rule.SubjectOrBodyContainsWords -or $Rule.FromAddressContainsWords)) {
            # RSS/Conversation History + keyword = high risk (score +4)
            $result.RiskScore += 4
            [void]$result.SuspiciousReasons.Add("Moves matching messages to rarely-checked folder: $folderPath")
        } elseif ($hasNonPrintable -or $isWhitespaceOnly) {
            # Non-printable folder name = high risk (score +4)
            $result.RiskScore += 4
            [void]$result.SuspiciousReasons.Add("Moves to folder with suspicious name: $folderPath")
        } elseif ($folderPath -match "\\Archive|\\Junk|\\Deleted Items") {
            # Archive/Junk/Deleted = medium risk (score +2)
            $result.RiskScore += 2
            [void]$result.SuspiciousReasons.Add("Moves to cleanup folder: $folderPath")
        } else {
            # Plain move to normal folder = low risk (score +1)
            $result.RiskScore += 1
            [void]$result.SuspiciousReasons.Add("Moves messages to: $folderPath")
        }
    }
    
    # Determine overall risk level and suspicious flag
    # Score >= 5: High risk (suspicious)
    # Score 3-4: Medium risk
    # Score < 3: Low risk
    if ($result.RiskScore -ge 5) {
        $result.IsSuspicious = $true
        $result.RiskLevel = "High"
    } elseif ($result.RiskScore -ge 3) {
        $result.RiskLevel = "Medium"
    } else {
        $result.RiskLevel = "Low"
    }
    
    return $result
}

# Legacy function wrapper for backward compatibility
function Test-ExternalForwardingLegacy {
    param(
        [Parameter(Mandatory=$true)]
        [array]$RecipientAddresses,
        [Parameter(Mandatory=$true)]
        [array]$InternalDomains
    )
    if ($null -eq $RecipientAddresses -or $RecipientAddresses.Count -eq 0) { return $false }
    foreach ($recipient in $RecipientAddresses) {
        $emailAddress = $null
        if ($recipient -is [string]) { $emailAddress = $recipient }
        elseif ($recipient.PropertyNames -contains 'Address') { $emailAddress = $recipient.Address } 
        elseif ($recipient.PropertyNames -contains 'EmailAddress') { $emailAddress = $recipient.EmailAddress } 
        if (-not [string]::IsNullOrWhiteSpace($emailAddress)) {
            # Parse email using regex to handle SMTP: and EX:/o=
            $parsedEmail = $null
            if ($emailAddress -match 'SMTP:([^\]]+)') {
                $parsedEmail = $matches[1]
            } elseif ($emailAddress -match '^EX:/o=') {
                continue # Internal
            } elseif ($emailAddress -match '@') {
                $parsedEmail = $emailAddress
            }
            
            if ($parsedEmail -and $parsedEmail -match '@') {
                try {
                    $domain = ($parsedEmail -split '@')[1].ToLowerInvariant()
                    if ($domain -notin $InternalDomains) { 
                        Write-Verbose "External domain: $domain"
                        return $true 
                    }
                } catch { 
                    Write-Verbose "Could not parse domain: $parsedEmail" 
                }
            }
        }
    }
    return $false
}

function Get-AutoDetectedDomains {
    param(
        [Parameter(Mandatory=$true)]
        [array]$MailboxUPNs,
        [Parameter(Mandatory=$false)]
        [int]$MaxSampleSize = 100
    )
    Write-Host "Auto-detecting organization domains from loaded mailboxes..." -ForegroundColor Cyan
    if (-not $MailboxUPNs -or $MailboxUPNs.Count -eq 0) {
        Write-Warning "No mailbox UPNs provided for domain detection"
        return @()
    }
    $sampleSize = [Math]::Min($MaxSampleSize, $MailboxUPNs.Count)
    $samplesToAnalyze = if ($MailboxUPNs.Count -gt $MaxSampleSize) {
        $step = [Math]::Floor($MailboxUPNs.Count / $MaxSampleSize)
        $samples = @()
        for ($i = 0; $i -lt $MailboxUPNs.Count; $i += $step) {
            $samples += $MailboxUPNs[$i]
            if ($samples.Count -ge $MaxSampleSize) { break }
        }
        $samples
    } else {
        $MailboxUPNs
    }
    Write-Host "Analyzing $($samplesToAnalyze.Count) mailbox UPNs for domain patterns..." -ForegroundColor Yellow
    $domainCounts = @{}
    $onMicrosoftDomains = @{}
    foreach ($upn in $samplesToAnalyze) {
        if ($upn -like "*@*") {
            try {
                $domain = ($upn.Split('@')[1]).ToLowerInvariant()
                if (-not [string]::IsNullOrWhiteSpace($domain)) {
                    if ($domain -like "*.onmicrosoft.com") {
                        if (-not $onMicrosoftDomains.ContainsKey($domain)) {
                            $onMicrosoftDomains[$domain] = 0
                        }
                        $onMicrosoftDomains[$domain]++
                    } else {
                        if (-not $domainCounts.ContainsKey($domain)) {
                            $domainCounts[$domain] = 0
                        }
                        $domainCounts[$domain]++
                    }
                }
            } catch {
                Write-Warning "Could not parse domain from UPN: $upn"
            }
        }
    }
    $detectedDomains = @()
    if ($domainCounts.Count -gt 0) {
        $sortedDomains = $domainCounts.GetEnumerator() | Sort-Object Value -Descending
        foreach ($domainEntry in $sortedDomains) {
            $detectedDomains += $domainEntry.Key
        }
    }
    if ($detectedDomains.Count -eq 0 -and $onMicrosoftDomains.Count -gt 0) {
        $sortedOnMicrosoftDomains = $onMicrosoftDomains.GetEnumerator() | Sort-Object Value -Descending
        foreach ($domainEntry in $sortedOnMicrosoftDomains) {
            $detectedDomains += $domainEntry.Key
        }
    }
    if ($detectedDomains.Count -gt 5) {
        $detectedDomains = $detectedDomains[0..4]
    }
    return $detectedDomains
}

# PERFORMANCE: Combined function to get both delegates and permissions in one API call
function Get-MailboxDelegatesAndPermissions {
    param(
        [Parameter(Mandatory=$true)]
        [string]$UserPrincipalName
    )
    
    # ROBUSTNESS: Use API call wrapper with logging and error handling
    $apiHelpersPath = Join-Path (Split-Path $PSScriptRoot -Parent) 'Scripts\Common\ApiRobustnessHelpers.psm1'
    $useRobustWrapper = $false
    if (Test-Path $apiHelpersPath) {
        try {
            Import-Module $apiHelpersPath -Force -ErrorAction SilentlyContinue
            if (Get-Command Invoke-ApiCallWithLogging -ErrorAction SilentlyContinue) {
                $useRobustWrapper = $true
            }
        } catch {
            # Fall back to standard error handling
        }
    }
    
    try {
        # ROBUSTNESS: Use wrapper with logging instead of SilentlyContinue
        if ($useRobustWrapper) {
            $permissions = Invoke-ApiCallWithLogging -ScriptBlock {
                Get-MailboxPermission -Identity $UserPrincipalName -ErrorAction Stop | 
                Where-Object { $_.User -notlike "*NT AUTHORITY*" -and $_.User -notlike "*S-1-*" }
            } -OperationName "Get-MailboxPermission" -ItemIdentifier $UserPrincipalName -ErrorAction 'Continue' -LogErrors $true
        } else {
            # Fallback: Log errors properly
            try {
                $permissions = Get-MailboxPermission -Identity $UserPrincipalName -ErrorAction Stop | 
                             Where-Object { $_.User -notlike "*NT AUTHORITY*" -and $_.User -notlike "*S-1-*" }
            } catch {
                $errorMsg = "Failed to get mailbox permissions for $UserPrincipalName : $($_.Exception.Message)"
                Write-Warning $errorMsg
                if (Get-Command Write-Log -ErrorAction SilentlyContinue) {
                    Write-Log -Message $errorMsg -Level Warning -Data @{ UserPrincipalName = $UserPrincipalName }
                }
                return @{
                    Delegates = "Error"
                    FullAccess = "Error"
                }
            }
        }
        
        if (-not $permissions) {
            return @{
                Delegates = "None"
                FullAccess = "None"
            }
        }
        
        # Extract delegates (FullAccess only)
        $delegateNames = [System.Collections.ArrayList]::new()
        $fullAccessUsers = [System.Collections.ArrayList]::new()
        $sendAsUsers = [System.Collections.ArrayList]::new()
        
        foreach ($perm in $permissions) {
            if ($perm.AccessRights -contains "FullAccess") {
                [void]$fullAccessUsers.Add($perm.User)
                [void]$delegateNames.Add($perm.User)
            }
            if ($perm.AccessRights -contains "SendAs") {
                [void]$sendAsUsers.Add($perm.User)
            }
        }
        
        # Remove duplicates and sort
        $uniqueDelegates = ($delegateNames | Sort-Object -Unique) -join ", "
        $uniqueFullAccess = ($fullAccessUsers | Sort-Object -Unique) -join ", "
        $uniqueSendAs = ($sendAsUsers | Sort-Object -Unique) -join ", "
        
        $delegates = if ($uniqueDelegates) { $uniqueDelegates } else { "None" }
        
        $resultParts = [System.Collections.ArrayList]::new()
        if ($uniqueFullAccess) { [void]$resultParts.Add("FullAccess: $uniqueFullAccess") }
        if ($uniqueSendAs) { [void]$resultParts.Add("SendAs: $uniqueSendAs") }
        $fullAccess = if ($resultParts.Count -gt 0) { ($resultParts -join "; ") } else { "None" }
        
        return @{
            Delegates = $delegates
            FullAccess = $fullAccess
        }
    } catch {
        $errorMsg = "Get-MailboxDelegatesAndPermissions failed for $UserPrincipalName : $($_.Exception.Message)"
        Write-Warning $errorMsg
        if (Get-Command Write-Log -ErrorAction SilentlyContinue) {
            Write-Log -Message $errorMsg -Level Error -Data @{ UserPrincipalName = $UserPrincipalName }
        }
        return @{
            Delegates = "Error"
            FullAccess = "Error"
        }
    }
}

function Analyze-MailboxDelegates {
    param(
        [Parameter(Mandatory=$true)]
        [string]$UserPrincipalName
    )
    
    try {
        $result = Get-MailboxDelegatesAndPermissions -UserPrincipalName $UserPrincipalName
        return $result.Delegates
    } catch {
        return "Error"
    }
}

function Analyze-MailboxPermissions {
    param(
        [Parameter(Mandatory=$true)]
        [string]$UserPrincipalName
    )
    
    try {
        $result = Get-MailboxDelegatesAndPermissions -UserPrincipalName $UserPrincipalName
        return $result.FullAccess
    } catch {
        return "Error"
    }
}

function Analyze-MailboxRulesEnhanced {
    <#
    .SYNOPSIS
        Analyzes mailbox rules with improved hidden rule detection and risk scoring.
    #>
    param(
        [Parameter(Mandatory=$true)]
        [array]$Rules,
        [Parameter(Mandatory=$true)]
        [array]$BaseSuspiciousKeywords,
        [Parameter(Mandatory=$true)]
        [array]$InternalDomains
    )
    
    $totalRules = $Rules.Count
    $suspiciousHiddenCount = 0
    $suspiciousVisibleCount = 0
    $hasExternalForwarding = $false
    
    # PERFORMANCE: Pre-build hashtable for legitimate patterns (O(1) lookup instead of O(n) loop)
    $legitimatePatterns = @{
        "system" = $true; "default" = $true; "outlook" = $true; "microsoft" = $true; "office" = $true; "exchange" = $true;
        "shared" = $true; "team" = $true; "group" = $true; "distribution" = $true; "dl" = $true; "mailbox" = $true;
        "automatic" = $true; "auto" = $true; "sync" = $true; "migration" = $true; "upgrade" = $true;
        "clutter" = $true; "focused" = $true; "junk" = $true; "spam" = $true; "archive" = $true; "retention" = $true;
        "compliance" = $true; "legal" = $true; "hold" = $true; "litigation" = $true; "ediscovery" = $true
    }
    
    # PERFORMANCE: Pre-build regex pattern for suspicious keywords (single regex match instead of loop)
    $suspiciousKeywordsPattern = $null
    if ($BaseSuspiciousKeywords -and $BaseSuspiciousKeywords.Count -gt 0) {
        $escapedKeywords = $BaseSuspiciousKeywords | ForEach-Object { [regex]::Escape($_) }
        $suspiciousKeywordsPattern = [regex]::new(($escapedKeywords -join '|'), [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    }
    
    foreach ($rule in $Rules) {
        # Enhanced hidden rule detection - only count truly suspicious hidden rules
        $isSuspiciousHidden = $false
        if ($rule.IsHidden) {
            # Check if this is a legitimate hidden rule or potentially malicious
            $ruleName = if ($rule.Name) { $rule.Name.ToLower() } else { "" }
            
            # PERFORMANCE: Check legitimate patterns using hashtable lookup (much faster)
            $isLegitimate = $false
            foreach ($pattern in $legitimatePatterns.Keys) {
                if ($ruleName -like "*$pattern*") {
                    $isLegitimate = $true
                    break
                }
            }
            
            # Additional checks for suspicious hidden rules
            if (-not $isLegitimate) {
                # PERFORMANCE: Use pre-built regex pattern instead of loop
                if ($suspiciousKeywordsPattern -and $ruleName -match $suspiciousKeywordsPattern) {
                    $isSuspiciousHidden = $true
                }
                
                # Check for symbols-only names in hidden rules
                if (-not $isSuspiciousHidden -and $ruleName.Length -gt 0) {
                    $textCharacters = $ruleName -replace '[^\p{L}\p{N}\s]', ''
                    if ([string]::IsNullOrWhiteSpace($textCharacters)) {
                        $isSuspiciousHidden = $true
                    }
                }
                
                # Check for external forwarding or suspicious actions using risk scoring
                if (-not $isSuspiciousHidden) {
                    $riskTest = Test-InboxRuleRisk -Rule $rule -InternalDomains $InternalDomains
                    if ($riskTest.IsExternal -or $riskTest.IsSuspicious) {
                        $isSuspiciousHidden = $true
                    }
                }
            }
            
            # Only count as suspicious hidden if it meets suspicious criteria
            if ($isSuspiciousHidden) {
                $suspiciousHiddenCount++
            }
        }
        
        # Check for suspicious keywords in visible rules
        $isSuspiciousVisible = $false
        if (-not $rule.IsHidden -and $rule.Name) {  # Only check visible rules for suspicious keywords
            # PERFORMANCE: Use pre-built regex pattern instead of loop
            if ($suspiciousKeywordsPattern -and $rule.Name -match $suspiciousKeywordsPattern) {
                $isSuspiciousVisible = $true
            }
            
            # Check for symbols-only names in visible rules
            if (-not $isSuspiciousVisible -and $rule.Name.Length -gt 0) {
                $textCharacters = $rule.Name -replace '[^\p{L}\p{N}\s]', ''
                if ([string]::IsNullOrWhiteSpace($textCharacters)) {
                    $isSuspiciousVisible = $true
                }
            }
        }
        
        # Check for external forwarding or suspicious actions using risk scoring
        $riskTest = Test-InboxRuleRisk -Rule $rule -InternalDomains $InternalDomains
        if ($riskTest.IsExternal) {
            $hasExternalForwarding = $true
        }
        
        # Count suspicious visible rules (avoid double-counting)
        if (-not $rule.IsHidden) {
            # Either has suspicious keyword/symbol pattern OR high risk score
            if ($isSuspiciousVisible -or ($riskTest.RiskScore -ge 5 -or $riskTest.IsSuspicious)) {
                $suspiciousVisibleCount++
            }
        }
    }
    
    return @{
        TotalRules = $totalRules
        SuspiciousHidden = $suspiciousHiddenCount
        SuspiciousVisible = $suspiciousVisibleCount
        HasExternalForwarding = $hasExternalForwarding
    }
}

Export-ModuleMember -Function Test-ExternalForwarding,Test-InboxRuleRisk,Test-ExternalForwardingLegacy,Analyze-MailboxDelegates,Analyze-MailboxPermissions,Get-AutoDetectedDomains,Get-MailboxDelegatesAndPermissions,Analyze-MailboxRulesEnhanced 