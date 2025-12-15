function Get-SettingsLocationConfig {
    # Get the script root directory (where ExchangeOnlineAnalyzer.ps1 and BulkTenantExporter.ps1 are located)
    # This ensures both scripts use the same config file
    
    # Try to get the calling script's directory (the script that imported this module)
    $scriptRoot = $null
    
    # Check call stack to find the script that imported this module
    $callStack = Get-PSCallStack
    foreach ($frame in $callStack) {
        if ($frame.ScriptName -and $frame.ScriptName -notlike '*\Modules\*') {
            $scriptRoot = Split-Path -Parent $frame.ScriptName
            break
        }
    }
    
    # Fallback: try to find ExchangeOnlineAnalyzer.ps1 or BulkTenantExporter.ps1 in parent directories
    if (-not $scriptRoot) {
        $moduleDir = $PSScriptRoot
        if ($moduleDir) {
            $parentDir = Split-Path -Parent $moduleDir
            if ($parentDir) {
                # Check if ExchangeOnlineAnalyzer.ps1 or BulkTenantExporter.ps1 exists in parent
                $exchPath = Join-Path $parentDir 'ExchangeOnlineAnalyzer.ps1'
                $bulkPath = Join-Path $parentDir 'BulkTenantExporter.ps1'
                if ((Test-Path $exchPath) -or (Test-Path $bulkPath)) {
                    $scriptRoot = $parentDir
                }
            }
        }
    }
    
    # Final fallback: use module directory
    if (-not $scriptRoot) {
        $scriptRoot = $PSScriptRoot
        if (-not $scriptRoot) {
            $scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
        }
    }
    
    $configFile = Join-Path $scriptRoot 'settings-location.config'
    return $configFile
}

function Get-SettingsPath {
    # Check for custom settings location
    $configFile = Get-SettingsLocationConfig
    if (Test-Path $configFile) {
        try {
            $customPath = Get-Content -Path $configFile -Raw -ErrorAction Stop | ForEach-Object { $_.Trim() }
            if (-not [string]::IsNullOrWhiteSpace($customPath) -and (Test-Path (Split-Path $customPath -Parent))) {
                return $customPath
            }
        } catch {
            # If config file is invalid, fall back to default
        }
    }
    
    # Default location
    $dir = Join-Path ([Environment]::GetFolderPath('ApplicationData')) 'ExchangeOnlineAnalyzer'
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    return (Join-Path $dir 'settings.json')
}

function Set-SettingsLocation {
    param(
        [Parameter(Mandatory=$true)]
        [string]$SettingsPath
    )
    
    try {
        # Validate the path
        $parentDir = Split-Path -Parent $SettingsPath
        if ([string]::IsNullOrWhiteSpace($parentDir)) {
            throw "Invalid path: parent directory cannot be determined"
        }
        
        # Ensure parent directory exists
        if (-not (Test-Path $parentDir)) {
            New-Item -ItemType Directory -Path $parentDir -Force -ErrorAction Stop | Out-Null
        }
        
        # Save the custom path to config file
        $configFile = Get-SettingsLocationConfig
        $SettingsPath | Out-File -FilePath $configFile -Encoding UTF8 -Force -ErrorAction Stop
        
        return $true
    } catch {
        Write-Error "Failed to set settings location: $($_.Exception.Message)"
        return $false
    }
}

function Get-AppSettings {
    try {
        $path = Get-SettingsPath
        if (Test-Path $path) {
            $raw = Get-Content -Path $path -Raw -ErrorAction Stop
            if ($raw.Trim().Length -gt 0) {
                $loaded = $raw | ConvertFrom-Json
                # Merge with defaults to ensure all fields exist and null values are replaced
                $defaults = Get-DefaultSettings
                foreach ($prop in $defaults.PSObject.Properties.Name) {
                    if (-not $loaded.PSObject.Properties.Name -contains $prop) {
                        $loaded | Add-Member -MemberType NoteProperty -Name $prop -Value $defaults.$prop
                    } elseif ($null -eq $loaded.$prop -or [string]::IsNullOrWhiteSpace($loaded.$prop)) {
                        $loaded.$prop = $defaults.$prop
                    }
                }
                return $loaded
            }
        }
    } catch {}
    return Get-DefaultSettings
}

function Get-DefaultSettings {
    return [pscustomobject]@{
        InvestigatorName = 'Security Administrator'
        InvestigatorTitle = 'Security Engineer'
        CompanyName = 'Organization'
        GeminiApiKey = ''
        ClaudeApiKey = ''
        TimeZone = 'CST'
        AdminUsernames = 'rrc,rradmin,rrcadmin,rmmadmin'
        InternalTeamDisplayNames = 'River Run,RRC Admin,Managed Services'
        AuthorizedISPs = 'Comcast,Charter,CenturyLink,Verizon,Brightspeed,AT&T,T-Mobile'
        InFlightWiFiProviders = 'Anuvu,Gogo,Viasat,Panasonic Avionics'
        ServicePrincipalNames = 'Microsoft Graph Command Line Tools'
        KnownAdmins = 'Jeff Beyer'
        ClientContactOverrides = '{}'
        ThirdPartyMFA = ''
        MemberberryEnabled = $false
        MemberberryPath = 'C:\git\memberberry\memberberry-complete-output.txt'
        MemberberryExceptionsPath = 'C:\git\memberberry\exceptions.json'
    }
}

function Save-AppSettings {
    param([Parameter(Mandatory=$true)][object]$Settings)
    try {
        $json = $Settings | ConvertTo-Json -Depth 4
        $path = Get-SettingsPath
        $json | Out-File -FilePath $path -Encoding utf8
        return $true
    } catch {
        Write-Error "Failed to save settings: $($_.Exception.Message)"; return $false
    }
}

function Get-MemberberryContent {
    param(
        [Parameter(Mandatory=$true)]
        [string]$MemberberryPath,
        [Parameter(Mandatory=$false)]
        [string]$MemberberryExceptionsPath = '',
        [Parameter(Mandatory=$false)]
        [string]$CompanyName = ''
    )
    
    $result = @{
        GlobalInstructions = ''
        ClientExceptions = ''
        Procedures = ''
        Success = $false
        ErrorMessage = ''
    }
    
    try {
        if (-not (Test-Path $MemberberryPath)) {
            $result.ErrorMessage = "Memberberry file not found: $MemberberryPath"
            return $result
        }
        
        $content = Get-Content -Path $MemberberryPath -Raw -ErrorAction Stop
        
        # File structure: 
        # - General instructions (lines 1-252)
        # - "# CLIENT EXCEPTIONS" placeholder marker (line 254)
        # - "# PROCEDURES" section (line 261+)
        # - "# CLIENT EXCEPTIONS" actual section (line 790+) - we ignore this, use JSON instead
        # - Ticket information may follow (we exclude this)
        
        $proceduresMarker = '(?m)^# PROCEDURES'
        $clientExceptionsMarker = '(?m)^# CLIENT EXCEPTIONS'
        
        # Common ticket markers - stop parsing at any of these (ticket content starts here)
        $ticketMarkers = @(
            '(?m)^Please analyze the following',
            '(?m)^Ticket #[0-9]',
            '(?m)^Security Alert',
            '(?m)^Alert Type',
            '(?m)^User:',
            '(?m)^Timestamp:',
            '(?m)^IP Address:',
            '(?m)^Contact:',
            '(?m)^Subject:',
            '(?m)^From:',
            '(?m)^To:',
            '(?m)^Date:',
            '(?m)^Message ID:',
            '(?m)^Incident',
            '(?m)^Case'
        )
        
        # Find the earliest ticket marker position (where ticket content starts)
        $ticketStartPos = $content.Length
        foreach ($marker in $ticketMarkers) {
            if ($content -match $marker) {
                $matchPos = $content.IndexOf($matches[0])
                if ($matchPos -ge 0 -and $matchPos -lt $ticketStartPos) {
                    $ticketStartPos = $matchPos
                }
            }
        }
        
        # Extract only useful content (before ticket starts)
        $usefulContent = if ($ticketStartPos -lt $content.Length) {
            $content.Substring(0, $ticketStartPos).TrimEnd()
        } else {
            $content
        }
        
        # Extract global instructions (everything before PROCEDURES)
        # This includes ANALYSIS PRINCIPLES, THREAT CLASSIFICATION RULES, REMEDIATION PROTOCOLS, DRAFTING STANDARDS, DEFAULTS
        if ($usefulContent -match "(.+?)$proceduresMarker") {
            $result.GlobalInstructions = $matches[1].Trim()
        } else {
            # If no PROCEDURES marker, try to find where CLIENT EXCEPTIONS starts (second occurrence)
            # We want everything before the actual CLIENT EXCEPTIONS section
            if ($usefulContent -match "(.+?)$clientExceptionsMarker") {
                $beforeFirst = $matches[1]
                # Check if there's a second CLIENT EXCEPTIONS marker (the actual one)
                $remaining = $usefulContent.Substring($beforeFirst.Length + $matches[0].Length)
                if ($remaining -match "$clientExceptionsMarker") {
                    # There's a second one - use everything before the first PROCEDURES or first CLIENT EXCEPTIONS
                    $result.GlobalInstructions = $beforeFirst.Trim()
                } else {
                    $result.GlobalInstructions = $beforeFirst.Trim()
                }
            } else {
                # Use everything we have (already filtered ticket content)
                $result.GlobalInstructions = $usefulContent.Trim()
            }
        }
        
        # Extract procedures section (between PROCEDURES and end, or before second CLIENT EXCEPTIONS)
        # The procedures section contains all the detailed procedures (Impossible Travel, Inbox Rule Anomaly, etc.)
        if ($usefulContent -match "$proceduresMarker(.+)") {
            $proceduresContent = $matches[1]
            
            # Remove the CLIENT EXCEPTIONS section if it appears after PROCEDURES (we use JSON for that)
            if ($proceduresContent -match "(.+?)$clientExceptionsMarker") {
                $proceduresContent = $matches[1]
            }
            
            # Clean up any trailing ticket markers that might have slipped through
            foreach ($marker in $ticketMarkers) {
                if ($proceduresContent -match "(.+?)$marker") {
                    $proceduresContent = $matches[1]
                }
            }
            
            $result.Procedures = $proceduresContent.Trim()
        }
        
        # Load client exceptions from JSON file if provided
        if ($CompanyName -and -not [string]::IsNullOrWhiteSpace($CompanyName) -and $MemberberryExceptionsPath) {
            try {
                if (Test-Path $MemberberryExceptionsPath) {
                    $exceptionsJson = Get-Content -Path $MemberberryExceptionsPath -Raw -ErrorAction Stop | ConvertFrom-Json
                    
                    # Try to find matching client (exact match first, then partial)
                    $matchedClient = $null
                    $matchedKey = $null
                    
                    # First try exact match
                    foreach ($key in $exceptionsJson.PSObject.Properties.Name) {
                        if ($key -eq $CompanyName) {
                            $matchedClient = $exceptionsJson.$key
                            $matchedKey = $key
                            break
                        }
                    }
                    
                    # If no exact match, try case-insensitive partial match
                    if (-not $matchedClient) {
                        $companyNameLower = $CompanyName.ToLower()
                        foreach ($key in $exceptionsJson.PSObject.Properties.Name) {
                            if ($key.ToLower() -eq $companyNameLower -or $key.ToLower().Contains($companyNameLower) -or $companyNameLower.Contains($key.ToLower())) {
                                $matchedClient = $exceptionsJson.$key
                                $matchedKey = $key
                                break
                            }
                        }
                    }
                    
                    # Format client exceptions if found
                    if ($matchedClient) {
                        $exceptionText = "## CLIENT EXCEPTIONS`n`n## $matchedKey`n`n"
                        
                        if ($matchedClient.detail) {
                            $exceptionText += "**Detail Level**: $($matchedClient.detail)`n"
                        }
                        if ($matchedClient.mfa) {
                            $exceptionText += "**MFA Provider**: $($matchedClient.mfa)`n"
                        }
                        if ($matchedClient.vpn) {
                            $exceptionText += "**VPN**: $($matchedClient.vpn)`n"
                        } else {
                            $exceptionText += "**VPN**: None`n"
                        }
                        if ($matchedClient.PSObject.Properties.Name -contains 'onsite_it') {
                            $exceptionText += "**Onsite IT**: $(if ($matchedClient.onsite_it) { 'Yes' } else { 'No' })`n"
                        }
                        if ($matchedClient.industry) {
                            $exceptionText += "**Industry**: $($matchedClient.industry)`n"
                        }
                        if ($matchedClient.authorized_tools -and $matchedClient.authorized_tools.Count -gt 0) {
                            $exceptionText += "**Authorized Tools**: $($matchedClient.authorized_tools -join ', ')`n"
                        }
                        if ($matchedClient.vips -and $matchedClient.vips.Count -gt 0) {
                            $exceptionText += "**VIPs**: $($matchedClient.vips -join ', ')`n"
                        }
                        if ($matchedClient.names -and $matchedClient.names.PSObject.Properties.Count -gt 0) {
                            $nameMappings = @()
                            foreach ($nameProp in $matchedClient.names.PSObject.Properties) {
                                $nameMappings += "$($nameProp.Name): $($nameProp.Value)"
                            }
                            if ($nameMappings.Count -gt 0) {
                                $exceptionText += "**Name Mappings**: $($nameMappings -join '; ')`n"
                            }
                        }
                        if ($matchedClient.notes) {
                            $exceptionText += "**Notes**: $($matchedClient.notes)`n"
                        }
                        
                        $result.ClientExceptions = $exceptionText.Trim()
                    }
                    
                    # Also include global exceptions if present
                    if ($exceptionsJson._global) {
                        $globalText = "`n`n## _global`n`n"
                        if ($exceptionsJson._global.notes) {
                            $globalText += "**Notes**: $($exceptionsJson._global.notes)`n"
                        }
                        if ($exceptionsJson._global.authorized_tools -and $exceptionsJson._global.authorized_tools.Count -gt 0) {
                            $globalText += "**Authorized Tools**: $($exceptionsJson._global.authorized_tools -join ', ')`n"
                        }
                        if ($exceptionsJson._global.vips -and $exceptionsJson._global.vips.Count -gt 0) {
                            $globalText += "**VIPs**: $($exceptionsJson._global.vips -join ', ')`n"
                        }
                        if ($result.ClientExceptions) {
                            $result.ClientExceptions += $globalText.Trim()
                        } else {
                            $result.ClientExceptions = $globalText.Trim()
                        }
                    }
                }
            } catch {
                Write-Warning "Failed to load memberberry exceptions JSON: $($_.Exception.Message)"
            }
        }
        
        $result.Success = $true
    } catch {
        $result.ErrorMessage = "Error reading memberberry file: $($_.Exception.Message)"
    }
    
    return $result
}

function New-AIReadme {
    param(
        [Parameter(Mandatory=$false)]
        [object]$Settings,
        [Parameter(Mandatory=$false)]
        [string[]]$TicketNumbers = @(),
        [Parameter(Mandatory=$false)]
        [string]$TicketContent = ''
    )
    
    if (-not $Settings) {
        $Settings = Get-AppSettings
    }
    
    # Check if memberberry is enabled and file exists
    $useMemberberry = $false
    $memberberryContent = $null
    $memberberryWarning = ''
    
    if ($Settings.MemberberryEnabled -eq $true -and $Settings.MemberberryPath) {
        try {
            $exceptionsPath = if ($Settings.MemberberryExceptionsPath) { $Settings.MemberberryExceptionsPath } else { '' }
            $memberberryContent = Get-MemberberryContent -MemberberryPath $Settings.MemberberryPath -MemberberryExceptionsPath $exceptionsPath -CompanyName $Settings.CompanyName
            if ($memberberryContent.Success) {
                $useMemberberry = $true
            } else {
                $memberberryWarning = "Warning: $($memberberryContent.ErrorMessage). Using default instructions."
            }
        } catch {
            $memberberryWarning = "Warning: Failed to load memberberry content: $($_.Exception.Message). Using default instructions."
        }
    }
    
    # Build ticket information section if tickets are provided
    $ticketSection = ''
    # Ensure TicketNumbers is an array
    $ticketNumsArray = @()
    if ($TicketNumbers) {
        if ($TicketNumbers -is [string]) {
            $ticketNumsArray = @($TicketNumbers)
        } elseif ($TicketNumbers -is [array]) {
            $ticketNumsArray = $TicketNumbers
        } else {
            $ticketNumsArray = @($TicketNumbers)
        }
    }
    Write-Host "New-AIReadme: Received TicketNumbers=$($ticketNumsArray.Count) ($($ticketNumsArray -join ', ')), TicketContent length=$($TicketContent.Length)" -ForegroundColor Gray
    # Add ticket section if we have ticket numbers OR ticket content
    if (($ticketNumsArray.Count -gt 0) -or (-not [string]::IsNullOrWhiteSpace($TicketContent))) {
        Write-Host "New-AIReadme: Building ticket section (has numbers: $($ticketNumsArray.Count -gt 0), has content: $(-not [string]::IsNullOrWhiteSpace($TicketContent)))" -ForegroundColor Gray
        $ticketNums = if ($ticketNumsArray.Count -gt 0) { $ticketNumsArray -join ', #' } else { '[Ticket Number]' }
        $ticketSection = @"

## ConnectWise Ticket Information

**Ticket Number(s)**: #$ticketNums

**Instructions**: Analyze the security alert based on the ticket information provided below. Use this ticket context to understand the specific alert details, user involved, timeline, and any relevant discussion or resolution notes.

**Ticket Content**:
$TicketContent

---

"@
        Write-Host "New-AIReadme: Ticket section built (length: $($ticketSection.Length) chars, content preview: $($TicketContent.Substring(0, [Math]::Min(100, $TicketContent.Length)))...)" -ForegroundColor Gray
    } else {
        Write-Host "New-AIReadme: No ticket section (no numbers and no content)" -ForegroundColor Gray
    }
    
    # If memberberry is enabled and loaded successfully, use it exclusively
    if ($useMemberberry) {
        $readme = $memberberryContent.GlobalInstructions
        
        # Prepend ticket information if provided
        if ($ticketSection) {
            Write-Host "New-AIReadme: Prepending ticket section to memberberry content (ticket section length: $($ticketSection.Length), readme length before: $($readme.Length))" -ForegroundColor Gray
            $readme = "$ticketSection$readme"
            Write-Host "New-AIReadme: Readme length after prepending: $($readme.Length)" -ForegroundColor Gray
        } else {
            Write-Host "New-AIReadme: No ticket section to prepend (memberberry enabled)" -ForegroundColor Yellow
        }
        
        # Append client exceptions if found
        if ($memberberryContent.ClientExceptions) {
            $readme += "`n`n$($memberberryContent.ClientExceptions)"
        }
        
        # Append procedures if available
        if ($memberberryContent.Procedures) {
            $readme += "`n`n$($memberberryContent.Procedures)"
        }
        
        # Add warning if present (prepend)
        if ($memberberryWarning) {
            $readme = "$memberberryWarning`n`n$readme"
        }
        
        # Update subject line to include ticket number if provided
        if ($ticketNumsArray.Count -gt 0) {
            $subjectLine = "Subject: Security Alert: Ticket $(($ticketNumsArray | ForEach-Object { "#$_" }) -join ', ') - [Brief Subject]"
            $readme = $readme -replace '(?m)^Subject: Security Alert:.*$', $subjectLine
        }
        
        return $readme
    }
    
    # Fall back to default instructions (existing logic)
    if ($memberberryWarning) {
        Write-Warning $memberberryWarning
    }
    
    # Parse comma-separated lists
    $adminUsers = if ($Settings.AdminUsernames) { ($Settings.AdminUsernames -split ',' | ForEach-Object { $_.Trim() }) -join ', ' } else { '[admin], [service_account], [rmm_account]' }
    $internalTeams = if ($Settings.InternalTeamDisplayNames) { ($Settings.InternalTeamDisplayNames -split ',' | ForEach-Object { $_.Trim() }) -join ', ' } else { 'Managed Services' }
    $authorizedISPs = if ($Settings.AuthorizedISPs) { ($Settings.AuthorizedISPs -split ',' | ForEach-Object { $_.Trim() }) -join ', ' } else { 'Comcast, Charter, CenturyLink, Verizon, Brightspeed, AT&T, T-Mobile' }
    $inFlightWiFi = if ($Settings.InFlightWiFiProviders) { ($Settings.InFlightWiFiProviders -split ',' | ForEach-Object { $_.Trim() }) -join ', ' } else { 'Anuvu, Gogo, Viasat, Panasonic Avionics' }
    $servicePrincipals = if ($Settings.ServicePrincipalNames) { ($Settings.ServicePrincipalNames -split ',' | ForEach-Object { $_.Trim() }) -join ', ' } else { 'Microsoft Graph Command Line Tools' }
    $knownAdmins = if ($Settings.KnownAdmins) { ($Settings.KnownAdmins -split ',' | ForEach-Object { $_.Trim() }) -join ', ' } else { '' }
    
    # Parse 3rd Party MFA clients (comma-separated list)
    $thirdPartyMFA = if ($Settings.ThirdPartyMFA) { ($Settings.ThirdPartyMFA -split ',' | ForEach-Object { $_.Trim() }) -join ', ' } else { '' }
    
    # Parse client contact overrides (JSON format: {"ClientName": {"Contact": "Full Name", "Greeting": "First Name"}})
    $contactOverrides = ''
    try {
        if ($Settings.ClientContactOverrides -and $Settings.ClientContactOverrides -ne '{}') {
            $overrides = $Settings.ClientContactOverrides | ConvertFrom-Json
            if ($overrides) {
                $contactOverrides = "`n`nClient Specific Nuances:`n"
                foreach ($client in $overrides.PSObject.Properties.Name) {
                    $contactOverrides += "$client`: The contact name is listed as `"$($overrides.$client.Contact)`", but always address them as `"$($overrides.$client.Greeting)`" in the email greeting.`n"
                }
            }
        }
    } catch {}
    
    $timeZone = if ($Settings.TimeZone) { $Settings.TimeZone } else { 'CST (Central Standard Time)' }
    $companyName = if ($Settings.CompanyName) { $Settings.CompanyName } else { '[Your Company Name]' }
    $investigatorName = if ($Settings.InvestigatorName) { $Settings.InvestigatorName } else { '[Your Name]' }
    $investigatorTitle = if ($Settings.InvestigatorTitle) { $Settings.InvestigatorTitle } else { '[Your Title]' }
    
    $readme = @"
Master Prompt - Generic Template (Copy and Save This)
$ticketSection
Role & Objective You are a Security Engineer acting on behalf of $companyName. Your task is to analyze security alert tickets, cross-reference them with attached CSV logs/text files, and classify the event as True Positive, False Positive, or Authorized Activity.



You will then draft a non-technical, professional email response to the client contact.



I. Data Ingestion & Analysis Rules

1. Analyze the Ticket Context



Ticket Body: Extract the User, Timestamp (UTC), IP Address, and Alert Type.



Ticket Notes/Configs: Look for notes like "Remote Employees," "Office Key," or specific authorized devices which indicate authorized activity.



Contact Name: Extract the contact from the "Contact" field. Check the "Client Specific Nuances" section below for any naming overrides.$contactOverrides



2. Verify with Logs (The "Evidence" Rule)



Crucial: Do not rely solely on the ticket description. You must find the corresponding event in the attached CSVs (SignInLogs, GraphAudit, etc.) to confirm the activity.



Time Zone: Convert all UTC timestamps to $timeZone for the email.



II. Classification Logic

A. Authorized Activity (White-Listed)

Internal Admin Accounts: Usernames like $adminUsers.



Verification: Check UserSecurityPosture.csv. If the Display Name matches your internal team (e.g., $internalTeams), treat as Authorized.



Action: Classify as Authorized Activity (Administrative Maintenance).



Travel (Residential/Mobile): Logins from standard ISPs ($authorizedISPs) in a different city/state.



Action: Classify as Authorized Activity (User Travel/Remote Work).



In-Flight Wi-Fi: IPs from $inFlightWiFi.



Action: Classify as Authorized Activity.



Service Principals: "MFA Disabled" alerts where the Actor is "$servicePrincipals"$(if ($knownAdmins) { " or a known Admin (e.g., $knownAdmins)" } else { "" }).



Action: Classify as Authorized Activity (Maintenance Script).


MFA Disabled Alerts: When reviewing tickets like "MFA Disabled" types, always check the logs to see if MFA was re-enabled after being disabled. Review the audit logs (GraphAudit.csv) for subsequent "MFA Enabled" or "MFA Registration" events for the same user. If MFA was re-enabled shortly after being disabled, this may indicate a temporary administrative action or user self-service re-enrollment rather than a security incident.


Action: Verify in logs whether MFA was re-enabled. If re-enabled, classify appropriately based on the context and timing.


3rd Party MFA:$(if ($thirdPartyMFA) { " The following clients use 3rd party MFA solutions (e.g., Duo Security) that won't show up in Entra exports: $thirdPartyMFA. When analyzing MFA status for these clients, note that they are covered by MFA even if the Entra exports show MFA as disabled." } else { " Note: Some clients may use 3rd party MFA solutions (e.g., Duo Security) that won't show up in Entra exports. When analyzing MFA status, verify if the client uses 3rd party MFA before classifying as a security issue." })


Action: When reviewing MFA status for these clients, treat them as having MFA enabled even if Entra exports indicate otherwise.



B. False Positives (System Noise)

Endpoint Protection: Alerts for TrustedInstaller.exe, `$`$DeleteMe..., or files in \Windows\WinSxS\Temp\.



Action: Classify as False Positive (System Update/Cleanup).



C. True Positives (Compromise Indicators)

Inbox Rules:



Name consists only of non-alphanumeric characters (e.g., ., .., ,,, ).



Action moves mail to "RSS Feeds" or "Conversation History" folders.



Action: Classify as True Positive. Recommend immediate password reset & session revocation.



D. Suspicious (Requires Confirmation)

Hosting Providers: Logins from AWS, DigitalOcean, Linode (unless the user has a known hosted workflow).



Consumer VPNs: NordVPN, ProtonVPN, Private Internet Access.



Action: Draft email asking for confirmation.



III. Output Format

Subject: Security Alert: Ticket #$(if ($ticketNumsArray.Count -gt 0) { $ticketNumsArray[0] } else { '[Ticket Number]' }) - [Brief Subject]



Hi [Contact First Name],



[Opening: State the alert type and the user involved.]



[Verdict: Explicitly state: "We have classified this as [Category]."]



[Analysis:



Source: [ISP Name / Location] (IP: [IP Address])



Evidence: Explain why it is classified this way (e.g., "This is a standard residential ISP," or "The rule name '.' is a known indicator of compromise"). Cite the specific log file used (e.g., ``).]



[Action Taken/Required:



If Authorized/False Positive: "No further action is required. We have closed this ticket."



If Suspicious: "Please confirm if [User] is currently [Traveling/Using a VPN]."



If True Positive: "We recommend immediately resetting the password and revoking sessions."]



Best,



$investigatorName $investigatorTitle



Clarification Questions [Ask 2 questions here regarding tuning, specific client policies, or missing data.]
"@
    
    return $readme
}

function Extract-TicketNumbers {
    param(
        [Parameter(Mandatory=$true)]
        [string]$TicketContent
    )
    
    $ticketNumbers = @()
    
    if ([string]::IsNullOrWhiteSpace($TicketContent)) {
        Write-Host "Extract-TicketNumbers: Ticket content is empty or whitespace" -ForegroundColor Yellow
        return $ticketNumbers
    }
    
    Write-Host "Extract-TicketNumbers: Processing ticket content (length: $($TicketContent.Length))" -ForegroundColor Gray
    Write-Host "Extract-TicketNumbers: Content preview (first 200 chars): $($TicketContent.Substring(0, [Math]::Min(200, $TicketContent.Length)))" -ForegroundColor Gray
    
    # Pattern 1: "Ticket #1809100" or "Service Ticket #1809100"
    $pattern1 = 'Ticket\s+#(\d+)'
    $matches1 = [regex]::Matches($TicketContent, $pattern1, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    Write-Host "Extract-TicketNumbers: Pattern 1 (Ticket #) found $($matches1.Count) match(es)" -ForegroundColor Gray
    foreach ($match in $matches1) {
        if ($match.Groups.Count -gt 1) {
            $ticketNum = $match.Groups[1].Value
            $ticketNumbers += $ticketNum
            Write-Host "Extract-TicketNumbers: Found ticket number (Pattern 1): $ticketNum" -ForegroundColor Green
        }
    }
    
    # Pattern 2: "Ticket: 126575144" or "Ticket 126575144"
    $pattern2 = 'Ticket[:\s]+(\d{7,})'
    $matches2 = [regex]::Matches($TicketContent, $pattern2, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    Write-Host "Extract-TicketNumbers: Pattern 2 (Ticket: or Ticket ) found $($matches2.Count) match(es)" -ForegroundColor Gray
    foreach ($match in $matches2) {
        if ($match.Groups.Count -gt 1) {
            $ticketNum = $match.Groups[1].Value
            # Only add if not already found by pattern 1
            if ($ticketNumbers -notcontains $ticketNum) {
                $ticketNumbers += $ticketNum
                Write-Host "Extract-TicketNumbers: Found ticket number (Pattern 2): $ticketNum" -ForegroundColor Green
            }
        }
    }
    
    # Return unique ticket numbers
    $uniqueTickets = $ticketNumbers | Select-Object -Unique
    Write-Host "Extract-TicketNumbers: Extracted $($uniqueTickets.Count) unique ticket number(s): $($uniqueTickets -join ', ')" -ForegroundColor $(if ($uniqueTickets.Count -gt 0) { 'Green' } else { 'Yellow' })
    return $uniqueTickets
}

function Filter-TicketContent {
    param(
        [Parameter(Mandatory=$true)]
        [string]$TicketContent
    )
    
    if ([string]::IsNullOrWhiteSpace($TicketContent)) {
        return $TicketContent
    }
    
    $lines = $TicketContent -split "`r?`n"
    $filteredLines = @()
    $skipConfigSection = $false
    $configListStarted = $false
    $blankLineCount = 0
    
    foreach ($line in $lines) {
        # Detect start of Configurations/Config List section
        # Look for various markers that indicate config section
        if ($line -match '^Configurations\s*$' -or 
            $line -match '^Configurations\s+\d+' -or 
            $line -match '^Config List' -or 
            $line -match '^Configurations\s*:' -or
            $line -match 'Config List\s*\(No View\)' -or
            ($line -match '\(No View\)' -and ($line -match 'Config' -or $line -match 'Attachments'))) {
            $skipConfigSection = $true
            $configListStarted = $true
            $blankLineCount = 0
            continue
        }
        
        # Detect config table headers (more comprehensive patterns)
        if (-not $skipConfigSection -and (
            $line -match 'Configuration Name\s+Configuration Type\s+Serial Number' -or
            $line -match 'Configuration Name\s+Configuration Type\s+Serial Number\s+Model Number' -or
            $line -match '^\s*RMITs are Remote' -or
            $line -match '^\s*Drag a pod here')) {
            $skipConfigSection = $true
            $configListStarted = $true
            $blankLineCount = 0
            continue
        }
        
        # Detect end of Configurations section
        if ($skipConfigSection) {
            # Stop skipping when we hit a new major section header or meaningful content
            if ($line -match '^(Additional Details|Knowledge Base|Resources|Team|Ticket Where|Board Icon|Ticket Type|Notes|Discussion|Resolution|Request ID|Impact|Request Status|Source IP|Contact:|Resources:|Cc:|Tasks|Attachments|Category|Subcategory|Allow all clients|Do not show|Drag a pod here|Search)' -or
                ($line -match '^[A-Z][a-z]+\s*:' -and -not ($line -match 'Config|Configuration'))) {
                $skipConfigSection = $false
                $configListStarted = $false
                $blankLineCount = 0
                # Don't add the section header line itself if it's just a label
                if ($line -match '^\s*[A-Z][a-z]+\s*:\s*$') {
                    continue
                }
            }
            
            # Skip lines that look like configuration table entries
            # Pattern: Device name/ID followed by device type (Managed Workstation, Managed Server, etc.)
            if ($line -match '^\s*[A-Z0-9\-_]+\s+(Managed Workstation|Managed Server|APPLICATION|BACKUP|Managed Network Switch|FIREWALL|HYPERVISOR|Managed Network Firewall)' -or
                # Pattern: Device name followed by serial/model numbers and dates
                $line -match '^\s*[A-Z0-9\-_]+\s+[A-Z0-9]+\s+[A-Z0-9\-]+\s+\d{1,2}/\d{1,2}/\d{2,4}' -or
                # Pattern: Device name, type, serial, model, contact name, date
                $line -match '^\s*[A-Z0-9\-_]+\s+(Managed Workstation|Managed Server|FIREWALL)\s+[A-Z0-9]+\s+[A-Z0-9\-]+\s+[A-Z][a-z]+\s+[A-Z][a-z]+\s+\d{1,2}/\d{1,2}/\d{2,4}' -or
                # Pattern: Device name, type, serial, Active status, MAC address
                $line -match '^\s*[A-Z0-9\-_]+\s+(Managed Workstation|Managed Server|FIREWALL)\s+[A-Z0-9]+\s+Active\s+[0-9A-F]{2}-[0-9A-F]{2}-' -or
                # Pattern: Device name, type, serial, Active, numeric ID
                $line -match '^\s*[A-Z0-9\-_]+\s+(Managed Workstation|Managed Server|FIREWALL)\s+[A-Z0-9]+\s+Active\s+\d+$' -or
                # Pattern: Device name, type, serial, domain\username
                $line -match '^\s*[A-Z0-9\-_]+\s+(Managed Workstation|Managed Server|FIREWALL)\s+[A-Z0-9]+\s+[A-Z]\\[A-Z0-9]' -or
                # Pattern: Device name, type, serial, IP address
                $line -match '^\s*[A-Z0-9\-_]+\s+(Managed Workstation|Managed Server|FIREWALL)\s+[A-Z0-9]+\s+\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}' -or
                # Pattern: Device name, type, OS version
                $line -match '^\s*[A-Z0-9\-_]+\s+(Managed Workstation|Managed Server|FIREWALL)\s+[A-Z0-9]+\s+Microsoft Windows' -or
                # Pattern: Device name, type, macOS
                $line -match '^\s*[A-Z0-9\-_]+\s+(Managed Workstation|Managed Server|FIREWALL)\s+[A-Z0-9]+\s+macOS' -or
                # Pattern: Device name, type, serial, contact name, date, Active
                $line -match '^\s*[A-Z0-9\-_]+\s+(Managed Workstation|Managed Server|FIREWALL)\s+[A-Z0-9]+\s+[A-Z0-9\-]+\s+[A-Z][a-z]+\s+[A-Z][a-z]+\s+\d{1,2}/\d{1,2}/\d{2,4}\s+Active' -or
                # Pattern: Device name, type, serial, IP, OS
                $line -match '^\s*[A-Z0-9\-_]+\s+(Managed Workstation|Managed Server|FIREWALL)\s+[A-Z0-9]+\s+\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\s+Microsoft Windows' -or
                # Pattern: Multiple tab-separated columns (config table row)
                ($line -match '^\s*[A-Z0-9\-_]+\s+' -and ($line -split '\s+').Count -ge 4 -and $line -match '(Managed|Active|Microsoft Windows|macOS|FIREWALL|HYPERVISOR)')) {
                $blankLineCount = 0
                continue
            }
            
            # Skip table header rows
            if ($line -match 'Configuration Name\s+Configuration Type\s+Serial Number' -or
                $line -match 'Configuration Name\s+Configuration Type\s+Serial Number\s+Model Number' -or
                $line -match '^\s*RMITs are Remote' -or
                $line -match '^\s*Drag a pod here' -or
                $line -match '^\s*Show All Active') {
                $blankLineCount = 0
                continue
            }
            
            # Track blank lines - if we see 3+ blank lines in a row, likely end of config section
            if ($line.Trim().Length -eq 0) {
                $blankLineCount++
                if ($blankLineCount -ge 3) {
                    # Multiple blank lines - might be end of config section, but continue skipping
                    continue
                }
            } else {
                $blankLineCount = 0
            }
            
            # If we hit a meaningful line that doesn't look like a config entry, stop skipping
            # Meaningful content: has substantial text, doesn't start with device ID pattern
            if ($line.Trim().Length -gt 20 -and 
                -not ($line -match '^\s*[A-Z0-9\-_]+\s+') -and
                -not ($line -match '^\s*[A-Z0-9\-_]+\s+(Managed|FIREWALL|HYPERVISOR)')) {
                # This looks like real content, stop skipping
                $skipConfigSection = $false
                $configListStarted = $false
                $blankLineCount = 0
                # Add this line
                $filteredLines += $line
            } else {
                # Still looks like config content, skip it
                continue
            }
        } else {
            # Not in config section, add the line
            $blankLineCount = 0
            $filteredLines += $line
        }
    }
    
    # Remove excessive blank lines (3+ consecutive blank lines become 2)
    $result = ($filteredLines -join "`n") -replace "(`r?`n){3,}", "`n`n"
    
    return $result.Trim()
}

Export-ModuleMember -Function Get-AppSettings,Save-AppSettings,Get-SettingsPath,Set-SettingsLocation,Get-SettingsLocationConfig,New-AIReadme,Get-MemberberryContent,Extract-TicketNumbers,Filter-TicketContent



