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

Export-ModuleMember -Function Test-ExternalForwarding,Get-InternalDomains,Analyze-MailboxRules,Analyze-MailboxDelegates,Analyze-MailboxPermissions,Get-AutoDetectedDomains 