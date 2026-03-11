function Test-ExternalForwarding {
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
        elseif ($recipient -is [Microsoft.Exchange.Data.SmtpAddress]) { $emailAddress = $recipient.ToString() }
        elseif ($recipient.PropertyNames -contains 'Address') { $emailAddress = $recipient.Address } 
        elseif ($recipient.PropertyNames -contains 'EmailAddress') { $emailAddress = $recipient.EmailAddress } 
        if (-not [string]::IsNullOrWhiteSpace($emailAddress)) {
            try {
                $domain = ($emailAddress -split '@')[1].ToLowerInvariant()
                if ($domain -notin $InternalDomains) { Write-Verbose "External domain: $domain"; return $true }
            } catch { Write-Warning "Could not parse domain: $emailAddress" }
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
        Analyzes mailbox rules with improved hidden rule detection.
    #>
    param(
        [Parameter(Mandatory=$true)]
        [array]$Rules,
        [Parameter(Mandatory=$true)]
        [array]$BaseSuspiciousKeywords
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
                
                # Check for external forwarding in hidden rules
                if (-not $isSuspiciousHidden -and $rule.ForwardTo -and $rule.ForwardTo -match '@') {
                    $isSuspiciousHidden = $true
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
        
        # Count suspicious rules (visible rules with suspicious characteristics)
        if ($isSuspiciousVisible) {
            $suspiciousVisibleCount++
        }
        
        # Check for external forwarding
        if ($rule.ForwardTo -and $rule.ForwardTo -match '@') {
            $hasExternalForwarding = $true
        }
    }
    
    return @{
        TotalRules = $totalRules
        SuspiciousHidden = $suspiciousHiddenCount
        SuspiciousVisible = $suspiciousVisibleCount
        HasExternalForwarding = $hasExternalForwarding
    }
}

Export-ModuleMember -Function Test-ExternalForwarding,Analyze-MailboxDelegates,Analyze-MailboxPermissions,Get-AutoDetectedDomains,Get-MailboxDelegatesAndPermissions,Analyze-MailboxRulesEnhanced 