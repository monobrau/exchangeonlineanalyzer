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
    
    try {
        # Single API call to get all permissions
        $permissions = Get-MailboxPermission -Identity $UserPrincipalName -ErrorAction SilentlyContinue | 
                     Where-Object { $_.User -notlike "*NT AUTHORITY*" -and $_.User -notlike "*S-1-*" }
        
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

Export-ModuleMember -Function Test-ExternalForwarding,Analyze-MailboxDelegates,Analyze-MailboxPermissions,Get-AutoDetectedDomains,Get-MailboxDelegatesAndPermissions 