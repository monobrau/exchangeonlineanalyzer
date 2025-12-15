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
        MemberberryPath = 'C:\git\memberberry'
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
                            $exceptionText += "**Name Preferences** (CRITICAL - Always use these preferred names when addressing users):`n"
                            foreach ($nameProp in $matchedClient.names.PSObject.Properties) {
                                $exceptionText += "  - **$($nameProp.Name)** → Use preferred name: **$($nameProp.Value)**`n"
                            }
                            $exceptionText += "`n"
                        }
                        if ($matchedClient.notes) {
                            $exceptionText += "**Notes**: $($matchedClient.notes)`n"
                        }
                        
                        $result.ClientExceptions = $exceptionText.Trim()
                    }
                    
                    # Also include global exceptions if present
                    if ($exceptionsJson._global) {
                        $globalText = "`n`n## ⚠️ GLOBAL EXCEPTIONS (APPLIES TO ALL CLIENTS) ⚠️`n`n"
                        if ($exceptionsJson._global.notes) {
                            $globalText += "### CRITICAL GLOBAL NOTES`n`n$($exceptionsJson._global.notes)`n`n"
                        }
                        if ($exceptionsJson._global.authorized_tools -and $exceptionsJson._global.authorized_tools.Count -gt 0) {
                            $globalText += "**Authorized Tools**: $($exceptionsJson._global.authorized_tools -join ', ')`n`n"
                        }
                        if ($exceptionsJson._global.vips -and $exceptionsJson._global.vips.Count -gt 0) {
                            $globalText += "**VIPs**: $($exceptionsJson._global.vips -join ', ')`n`n"
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
    $memberberryContent = @{
        GlobalInstructions = ''
        ClientExceptions = ''
        Procedures = ''
        Success = $false
        ErrorMessage = ''
    }
    $memberberryWarning = ''
    
    # Debug: Log memberberry settings
    Write-Host "New-AIReadme: MemberberryEnabled=$($Settings.MemberberryEnabled) (type: $($Settings.MemberberryEnabled.GetType().Name)), MemberberryPath='$($Settings.MemberberryPath)'" -ForegroundColor Gray
    
    # Check if memberberry is enabled (handle both boolean true and string "true")
    $memberberryEnabled = $false
    if ($Settings.MemberberryEnabled) {
        if ($Settings.MemberberryEnabled -eq $true -or $Settings.MemberberryEnabled -eq "true" -or $Settings.MemberberryEnabled.ToString() -eq "True") {
            $memberberryEnabled = $true
        }
    }
    
    if ($memberberryEnabled -and $Settings.MemberberryPath) {
        # Validate that MemberberryPath is a directory, not a file
        if (Test-Path $Settings.MemberberryPath -PathType Leaf) {
            Write-Warning "MemberberryPath points to a file, not a directory: $($Settings.MemberberryPath). Please set MemberberryPath to the directory containing the compile.ps1 script (e.g., 'C:\git\memberberry')."
            $memberberryWarning = "MemberberryPath is configured incorrectly (points to file instead of directory). Using default instructions."
        } elseif (-not (Test-Path $Settings.MemberberryPath -PathType Container)) {
            Write-Warning "MemberberryPath directory does not exist: $($Settings.MemberberryPath). Using default instructions."
            $memberberryWarning = "MemberberryPath directory not found: $($Settings.MemberberryPath). Using default instructions."
        } else {
            try {
                # Run memberberry script to generate fresh output
                Write-Host "Running memberberry script to generate updated output..." -ForegroundColor Cyan
            
            # Try to find and execute memberberry script
            # Common script names: compile.ps1, memberberry.ps1, run.ps1, main.ps1, memberberry.py, run.py, main.py
            $memberberryScript = $null
            $scriptNames = @('compile.ps1', 'memberberry.ps1', 'run.ps1', 'main.ps1', 'memberberry.py', 'run.py', 'main.py', 'memberberry.bat', 'run.bat')
            
            foreach ($scriptName in $scriptNames) {
                $scriptPath = Join-Path $Settings.MemberberryPath $scriptName
                if (Test-Path $scriptPath) {
                    $memberberryScript = $scriptPath
                    break
                }
            }
            
            if ($memberberryScript) {
                Write-Host "Found memberberry script: $memberberryScript" -ForegroundColor Gray
                try {
                    $scriptExtension = [System.IO.Path]::GetExtension($memberberryScript).ToLower()
                    
                    if ($scriptExtension -eq '.ps1') {
                        # Execute PowerShell script using call operator to show output in current console
                        try {
                            & $memberberryScript
                            Write-Host "Memberberry script completed successfully" -ForegroundColor Green
                        } catch {
                            Write-Warning "Memberberry script execution failed: $($_.Exception.Message). Continuing with existing output file."
                        }
                    } elseif ($scriptExtension -eq '.py') {
                        # Execute Python script
                        $pythonExe = Get-Command python -ErrorAction SilentlyContinue
                        if (-not $pythonExe) {
                            $pythonExe = Get-Command python3 -ErrorAction SilentlyContinue
                        }
                        if ($pythonExe) {
                            $process = Start-Process -FilePath $pythonExe.Path -ArgumentList "`"$memberberryScript`"" -Wait -PassThru -NoNewWindow
                            if ($process.ExitCode -ne 0) {
                                Write-Warning "Memberberry script exited with code $($process.ExitCode)"
                            } else {
                                Write-Host "Memberberry script completed successfully" -ForegroundColor Green
                            }
                        } else {
                            Write-Warning "Python not found. Skipping memberberry script execution."
                        }
                    } elseif ($scriptExtension -eq '.bat' -or $scriptExtension -eq '.cmd') {
                        # Execute batch file
                        $process = Start-Process -FilePath "cmd.exe" -ArgumentList "/c", "`"$memberberryScript`"" -Wait -PassThru -NoNewWindow
                        if ($process.ExitCode -ne 0) {
                            Write-Warning "Memberberry script exited with code $($process.ExitCode)"
                        } else {
                            Write-Host "Memberberry script completed successfully" -ForegroundColor Green
                        }
                    }
                } catch {
                    Write-Warning "Failed to execute memberberry script: $($_.Exception.Message). Continuing with existing output file."
                }
            } else {
                Write-Host "Memberberry script not found in $($Settings.MemberberryPath). Looking for: $($scriptNames -join ', ')" -ForegroundColor Yellow
                Write-Host "Continuing with existing output file if available." -ForegroundColor Yellow
            }
            
            # Construct path to memberberry output file
            # Expected: $MemberberryPath\output\memberberry.md
            # Example: C:\Git\memberberry\output\memberberry.md
            $memberberryOutputFile = Join-Path $Settings.MemberberryPath "output\memberberry.md"
            
            if (Test-Path $memberberryOutputFile) {
                Write-Host "Reading memberberry content from: $memberberryOutputFile" -ForegroundColor Gray
                $rawContent = Get-Content $memberberryOutputFile -Raw -ErrorAction Stop
                
                # Filter out ticket information (same logic as before)
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
                
                # Find the earliest ticket marker position
                $ticketStartPos = $rawContent.Length
                foreach ($marker in $ticketMarkers) {
                    if ($rawContent -match $marker) {
                        $matchPos = $rawContent.IndexOf($matches[0])
                        if ($matchPos -ge 0 -and $matchPos -lt $ticketStartPos) {
                            $ticketStartPos = $matchPos
                        }
                    }
                }
                
                # Extract only useful content (before ticket starts)
                # Use the ENTIRE memberberry.md content as-is (minus ticket information)
                $usefulContent = if ($ticketStartPos -lt $rawContent.Length) {
                    $rawContent.Substring(0, $ticketStartPos).TrimEnd()
                } else {
                    $rawContent
                }
                
                # Remove CLIENT EXCEPTIONS section from the file (we use JSON for that instead)
                # But keep everything else including PROCEDURES
                $clientExceptionsMarker = '(?m)^# CLIENT EXCEPTIONS'
                if ($usefulContent -match "(.+?)$clientExceptionsMarker") {
                    $usefulContent = $matches[1].TrimEnd()
                }
                
                # Use the entire filtered content as GlobalInstructions (includes PROCEDURES)
                # This preserves the complete memberberry instructions as intended
                $memberberryContent.GlobalInstructions = $usefulContent.Trim()
                
                # Procedures are already included in GlobalInstructions, so set to empty
                $memberberryContent.Procedures = ''
                
                # Load client exceptions from JSON file if provided
                $exceptionsPath = if ($Settings.MemberberryExceptionsPath) { $Settings.MemberberryExceptionsPath } else { '' }
                
                # Validate exceptions path - must be a file, not a directory
                if ($exceptionsPath) {
                    if (Test-Path $exceptionsPath -PathType Container) {
                        Write-Warning "MemberberryExceptionsPath points to a directory, not a file: $exceptionsPath. Expected path to exceptions.json file (e.g., 'C:\git\memberberry\exceptions.json'). Skipping exceptions loading."
                        $exceptionsPath = ''
                    } elseif (-not (Test-Path $exceptionsPath -PathType Leaf)) {
                        Write-Warning "MemberberryExceptionsPath file not found: $exceptionsPath. Skipping exceptions loading."
                        $exceptionsPath = ''
                    }
                }
                
                # Load exceptions JSON if path is valid (global exceptions should ALWAYS be loaded if they exist)
                if ($exceptionsPath) {
                    try {
                        if (Test-Path $exceptionsPath -PathType Leaf) {
                            $exceptionsJson = Get-Content -Path $exceptionsPath -Raw -ErrorAction Stop | ConvertFrom-Json
                            
                            # ALWAYS load global exceptions first (they apply to all clients)
                            if ($exceptionsJson._global) {
                                Write-Host "New-AIReadme: Loading global exceptions from JSON" -ForegroundColor Gray
                                $globalText = "`n`n## ⚠️ GLOBAL EXCEPTIONS (APPLIES TO ALL CLIENTS) ⚠️`n`n"
                                if ($exceptionsJson._global.notes) {
                                    # Process notes to add explicit reboot prohibition for Advanced IP Scanner
                                    $notes = $exceptionsJson._global.notes
                                    # If notes mention Advanced IP Scanner and mitigation/reboot, add explicit prohibition
                                    if ($notes -match 'Advanced IP Scanner' -and ($notes -match 'mitigation|reboot|Quarantined|Mitigated')) {
                                        # Replace the technical note to make it clearer
                                        $notes = $notes -replace '(?m)(Technical Note:.*?reboot\.)', "`$1`n`n**🚫 ABSOLUTE PROHIBITION: DO NOT REQUEST A REBOOT. DO NOT ASK THE CLIENT TO REBOOT. DO NOT INCLUDE 'Action Required: Please reboot' OR ANY VARIATION. Mitigation happens automatically - you may inform the client of this fact, but you MUST NOT request any action from them.**"
                                        # Also add a prominent warning at the end of Advanced IP Scanner section
                                        if ($notes -match 'Advanced IP Scanner:') {
                                            $notes = $notes -replace '(?m)(STRICT CONSTRAINT:.*?DO NOT NEED.*?allow-list.*?SOC)', "`$1`n`n**🚫 REBOOT PROHIBITION: When Advanced IP Scanner is detected, your 'Action Required' section MUST say 'No action required' or be omitted entirely. DO NOT request a reboot, even if mitigation is pending. Mitigation completes automatically.**"
                                        }
                                    }
                                    $globalText += "### CRITICAL GLOBAL NOTES`n`n$notes`n`n"
                                }
                                if ($exceptionsJson._global.authorized_tools -and $exceptionsJson._global.authorized_tools.Count -gt 0) {
                                    $globalText += "**Authorized Tools**: $($exceptionsJson._global.authorized_tools -join ', ')`n`n"
                                }
                                if ($exceptionsJson._global.vips -and $exceptionsJson._global.vips.Count -gt 0) {
                                    $globalText += "**VIPs**: $($exceptionsJson._global.vips -join ', ')`n`n"
                                }
                                $memberberryContent.ClientExceptions = $globalText.Trim()
                                Write-Host "New-AIReadme: Global exceptions loaded successfully (length: $($globalText.Length) chars)" -ForegroundColor Green
                            }
                            
                            # Load client-specific exceptions if company name is provided
                            if ($Settings.CompanyName -and -not [string]::IsNullOrWhiteSpace($Settings.CompanyName)) {
                                Write-Host "New-AIReadme: Searching for client exceptions for: $($Settings.CompanyName)" -ForegroundColor Gray
                                
                                # Try to find matching client (exact match first, then partial)
                                $matchedClient = $null
                                $matchedKey = $null
                                
                                # First try exact match
                                foreach ($key in $exceptionsJson.PSObject.Properties.Name) {
                                    if ($key -eq $Settings.CompanyName -and $key -ne '_global') {
                                        $matchedClient = $exceptionsJson.$key
                                        $matchedKey = $key
                                        break
                                    }
                                }
                                
                                # If no exact match, try case-insensitive partial match
                                if (-not $matchedClient) {
                                    $companyNameLower = $Settings.CompanyName.ToLower()
                                    foreach ($key in $exceptionsJson.PSObject.Properties.Name) {
                                        if ($key -ne '_global' -and ($key.ToLower() -eq $companyNameLower -or $key.ToLower().Contains($companyNameLower) -or $companyNameLower.Contains($key.ToLower()))) {
                                            $matchedClient = $exceptionsJson.$key
                                            $matchedKey = $key
                                            break
                                        }
                                    }
                                }
                                
                                # Format client exceptions if found
                                if ($matchedClient) {
                                    Write-Host "New-AIReadme: Found client exceptions for: $matchedKey" -ForegroundColor Green
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
                                        $exceptionText += "**Name Preferences** (CRITICAL - Always use these preferred names when addressing users):`n"
                                        foreach ($nameProp in $matchedClient.names.PSObject.Properties) {
                                            $exceptionText += "  - **$($nameProp.Name)** → Use preferred name: **$($nameProp.Value)**`n"
                                        }
                                        $exceptionText += "`n"
                                    }
                                    if ($matchedClient.notes) {
                                        $exceptionText += "**Notes**: $($matchedClient.notes)`n"
                                    }
                                    
                                    # Append client exceptions to global exceptions (if they exist)
                                    if ($memberberryContent.ClientExceptions) {
                                        $memberberryContent.ClientExceptions += "`n`n$($exceptionText.Trim())"
                                    } else {
                                        $memberberryContent.ClientExceptions = $exceptionText.Trim()
                                    }
                                } else {
                                    Write-Host "New-AIReadme: No client-specific exceptions found for: $($Settings.CompanyName)" -ForegroundColor Yellow
                                }
                            }
                        }
                    } catch {
                        Write-Warning "Failed to load memberberry exceptions JSON: $($_.Exception.Message)"
                    }
                }
                
                $memberberryContent.Success = $true
                $useMemberberry = $true
                Write-Host "New-AIReadme: Memberberry content loaded successfully (length: $($memberberryContent.GlobalInstructions.Length) chars)" -ForegroundColor Green
            } else {
                $memberberryWarning = "Warning: Memberberry output file not found: $memberberryOutputFile. Using default instructions."
                Write-Warning $memberberryWarning
            }
            } catch {
                $memberberryWarning = "Warning: Failed to load memberberry content: $($_.Exception.Message). Using default instructions."
                Write-Warning $memberberryWarning
            }
        }
    } elseif ($Settings.MemberberryEnabled -eq $true) {
        Write-Warning "Memberberry is enabled but MemberberryPath is not configured. Using default instructions."
        $memberberryWarning = "Memberberry is enabled but MemberberryPath is not configured."
    } elseif ($Settings.MemberberryPath -and -not $Settings.MemberberryEnabled) {
        Write-Host "New-AIReadme: MemberberryPath is configured but MemberberryEnabled is false. Using default instructions." -ForegroundColor Yellow
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
        Write-Host "New-AIReadme: Using memberberry content (GlobalInstructions length: $($memberberryContent.GlobalInstructions.Length) chars)" -ForegroundColor Green
        $readme = $memberberryContent.GlobalInstructions
        
        # Validate that we actually have content
        if ([string]::IsNullOrWhiteSpace($readme)) {
            Write-Warning "New-AIReadme: WARNING - Memberberry enabled but GlobalInstructions is empty! Falling back to default template."
            $useMemberberry = $false
        } else {
            # Add critical reminders section at the beginning to ensure LLMs don't miss important details
            $criticalReminders = @"

---

## ⚠️ CRITICAL REMINDERS - READ BEFORE DRAFTING ⚠️

### Name Preferences (MANDATORY CHECK)
- **ALWAYS check the CLIENT EXCEPTIONS section for Name Preferences before addressing users**
- If a name preference exists (e.g., "Joseph Nedvidek → Use preferred name: Joe"), you MUST use the preferred name
- Example: If you see "**Joseph Nedvidek** → Use preferred name: **Joe**", address the user as "Hi Joe," NOT "Hi Joseph,"
- Name preferences are listed in the CLIENT EXCEPTIONS section under "**Name Preferences**"
- **Failure to use preferred names is a critical error**

### Advanced IP Scanner - CRITICAL GLOBAL EXCEPTION (MANDATORY CHECK)
- **⚠️ ALWAYS check the GLOBAL EXCEPTIONS section for Advanced IP Scanner instructions**
- **When Advanced IP Scanner is detected:**
  - **🚫 ABSOLUTELY DO NOT request a server reboot** - This is EXPLICITLY FORBIDDEN. NO EXCEPTIONS. NO REBOOT REQUESTS ARE ALLOWED, REGARDLESS OF WORDING OR REASONING.
  - **❌ FORBIDDEN - DO NOT INCLUDE ANY OF THESE IN YOUR RESPONSE:**
    - "Action Required: Please reboot"
    - "Please reboot the server"
    - "Reboot required"
    - "Action Required: Please reboot the server to clear any temporary files"
    - "Please reboot the server to clear any temporary files locked by the security agent"
    - "Please reboot the server... as the original installation was interrupted"
    - "Action Required" section containing ANY mention of reboot
    - ANY phrase containing "reboot" + "required" or "please reboot" or "reboot the"
    - ANY request, suggestion, or instruction to reboot the server
  - **✅ CORRECT APPROACH**: 
    - If mitigation is pending, you may say: "Mitigation will complete automatically on the next system reboot" (informational only, no action requested)
    - Your "Action Required" section should say: "No action required" or be omitted entirely
    - DO NOT mention rebooting in any context that implies the client should do it
  - **🚫 ABSOLUTELY DO NOT draft an allow-list request to the SOC** - Advanced IP Scanner does NOT need an allow-list request. This is explicitly stated in GLOBAL EXCEPTIONS. Do NOT create any SOC exception request emails.
  - **🚫 DO NOT offer to install Angry IP Scanner** - Only suggest it as a recommendation for the client to implement if they need IP scanning functionality. Explicitly state this is a recommendation for them to implement if needed.
  - **Narrative**: Explain that while River Run historically used this tool, it is now deprecated because it has not received updates in over 3 years and contains unpatched vulnerabilities.
  - **Replacement**: Suggest "Angry IP Scanner" ONLY if the client requires IP scanning functionality on this device. Explicitly state this is a recommendation for them to implement if needed.
  - **Action Required Section**: MUST say "No action required" or be omitted entirely - DO NOT include ANY reboot-related instructions or suggestions.
- **These constraints are in the GLOBAL EXCEPTIONS section - check there for full details**
- **CRITICAL: If you include ANY reboot request in your response, you have FAILED this task. NO REBOOT REQUESTS ARE ALLOWED.**

### Software Recommendations (MANDATORY CHECK)
- **ALWAYS check for software recommendations in the instructions**
- When Advanced IP Scanner is detected, you MUST mention the recommended replacement: **Angry IP Scanner**
- The instructions specify: "Recommend using alternative network scanning tools that are actively maintained and have valid certificates" - the specific recommendation is **Angry IP Scanner**
- **Always include software recommendations when relevant to the alert**

---

"@
            # Insert reminders after the main header but before the main content
            # Try multiple patterns to catch different header formats
            if ($readme -match '^(# MEMBERBERRY[^\n]*\r?\n)') {
                $readme = $readme -replace '^(# MEMBERBERRY[^\n]*\r?\n)', "`$1$criticalReminders"
                Write-Host "New-AIReadme: CRITICAL REMINDERS section inserted after header" -ForegroundColor Green
            } elseif ($readme -match '^# MEMBERBERRY') {
                # If header exists but no newline after, insert after first line
                $readme = $readme -replace '^(# MEMBERBERRY[^\n]*)', "`$1`n$criticalReminders"
                Write-Host "New-AIReadme: CRITICAL REMINDERS section inserted after header (no newline)" -ForegroundColor Green
            } else {
                # If no header match, prepend reminders
                $readme = "$criticalReminders`n`n$readme"
                Write-Host "New-AIReadme: CRITICAL REMINDERS section prepended (no header found)" -ForegroundColor Yellow
            }
            # Append client exceptions if found (procedures are already included in GlobalInstructions)
            if ($memberberryContent.ClientExceptions) {
                $readme += "`n`n$($memberberryContent.ClientExceptions)"
            }
            
            # Append ticket information AFTER the instructions if provided
            if ($ticketSection) {
                Write-Host "New-AIReadme: Appending ticket section after memberberry content (ticket section length: $($ticketSection.Length), readme length before: $($readme.Length))" -ForegroundColor Gray
                $readme += "`n`n$ticketSection"
                Write-Host "New-AIReadme: Readme length after appending: $($readme.Length)" -ForegroundColor Gray
            } else {
                Write-Host "New-AIReadme: No ticket section to append (memberberry enabled)" -ForegroundColor Yellow
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
    
    # Append ticket information AFTER the instructions if provided
    if ($ticketSection) {
        Write-Host "New-AIReadme: Appending ticket section after default template (ticket section length: $($ticketSection.Length), readme length before: $($readme.Length))" -ForegroundColor Gray
        $readme += "`n`n$ticketSection"
        Write-Host "New-AIReadme: Readme length after appending: $($readme.Length)" -ForegroundColor Gray
    }
    
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
    $inMultiLineConfigEntry = $false
    $configEntryLineCount = 0
    
    for ($i = 0; $i -lt $lines.Count; $i++) {
        $line = $lines[$i]
        $nextLine = if ($i -lt $lines.Count - 1) { $lines[$i + 1] } else { '' }
        $prevLine = if ($i -gt 0) { $lines[$i - 1] } else { '' }
        
        # Detect start of Configurations/Config List section
        if ($line -match '^Configurations\s*$' -or 
            $line -match '^Configurations\s+\d+' -or 
            $line -match '^Config List' -or 
            $line -match '^Configurations\s*:' -or
            $line -match 'Config List\s*\(No View\)' -or
            ($line -match '\(No View\)' -and ($line -match 'Config' -or $line -match 'Attachments'))) {
            $skipConfigSection = $true
            $configListStarted = $true
            $blankLineCount = 0
            $inMultiLineConfigEntry = $false
            continue
        }
        
        # Detect config table headers
        if (-not $skipConfigSection -and (
            $line -match 'Configuration Name\s+Configuration Type\s+Serial Number' -or
            $line -match 'Configuration Name\s+Configuration Type\s+Serial Number\s+Model Number' -or
            $line -match '^\s*RMITs are Remote' -or
            $line -match '^\s*Drag a pod here')) {
            $skipConfigSection = $true
            $configListStarted = $true
            $blankLineCount = 0
            $inMultiLineConfigEntry = $false
            continue
        }
        
        # Detect multi-line config entries: Device name/ID on one line, followed by "Managed Server" or "Managed Workstation" on next line
        # Pattern: Line with device ID (IP-AC1F92B9, JAMIE, JASONNEW-PC, etc.) followed by device type on next line
        # Also detect standalone device IDs that are likely config entries (even without next line check)
        if (-not $skipConfigSection) {
            $isDeviceId = $line.Trim() -match '^(IP-[A-Z0-9]+|[A-Z0-9\-_]+(?:-PC|_WORK|DT\d{4}-\d{2})?)$'
            $nextIsDeviceType = -not [string]::IsNullOrWhiteSpace($nextLine) -and
                $nextLine.Trim() -match '^(Managed Server|Managed Workstation|Managed Network Switch|Managed Network Firewall|FIREWALL|HYPERVISOR|APPLICATION|BACKUP|Domain Controller|Windows Server)$'
            
            if ($isDeviceId -and $nextIsDeviceType) {
                $skipConfigSection = $true
                $configListStarted = $true
                $inMultiLineConfigEntry = $true
                $configEntryLineCount = 0
                $blankLineCount = 0
                continue
            }
            
            # Also detect if this looks like a config entry based on surrounding context
            # If we see a device ID followed by lines that look like config data, start skipping
            if ($isDeviceId -and $i + 2 -lt $lines.Count) {
                $line2 = $lines[$i + 1]
                $line3 = $lines[$i + 2]
                # Check if next few lines look like config data (UUID, date, status, etc.)
                if ($line2.Trim() -match '^(Managed Server|Managed Workstation|Managed Network Switch|Managed Network Firewall|FIREWALL|HYPERVISOR|APPLICATION|BACKUP|Domain Controller|Windows Server|[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}|t3\.|c7i-|ec2[a-z0-9-]+|\d{1,2}/\d{1,2}/\d{2,4}|Active|Inactive)$' -or
                    $line3.Trim() -match '^(Managed Server|Managed Workstation|[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}|t3\.|c7i-|ec2[a-z0-9-]+|\d{1,2}/\d{1,2}/\d{2,4}|Active|Inactive)$') {
                    $skipConfigSection = $true
                    $configListStarted = $true
                    $inMultiLineConfigEntry = $true
                    $configEntryLineCount = 0
                    $blankLineCount = 0
                    continue
                }
            }
        }
        
        # Detect single-line config entries (device name followed by type on same line)
        if (-not $skipConfigSection -and (
            $line -match '^\s*[A-Z0-9\-_]+\s+(Managed Workstation|Managed Server|APPLICATION|BACKUP|Managed Network Switch|FIREWALL|HYPERVISOR|Managed Network Firewall)' -or
            $line -match '^\s*[A-Z0-9\-_]+\s+[A-Z0-9]+\s+[A-Z0-9\-]+\s+\d{1,2}/\d{1,2}/\d{2,4}' -or
            ($line -match '^\s*[A-Z0-9\-_]+\s+' -and ($line -split '\s+').Count -ge 4 -and $line -match '(Managed|Active|Microsoft Windows|macOS|FIREWALL|HYPERVISOR)'))) {
            $skipConfigSection = $true
            $configListStarted = $true
            $inMultiLineConfigEntry = $false
            $blankLineCount = 0
            continue
        }
        
        # Handle lines when we're in a config section
        if ($skipConfigSection) {
            # If we're in a multi-line config entry, count lines and stop after reasonable number (e.g., 20 lines per device)
            if ($inMultiLineConfigEntry) {
                $configEntryLineCount++
                
                # Check if next line starts a new device entry (device ID pattern)
                if (-not [string]::IsNullOrWhiteSpace($nextLine) -and 
                    $nextLine.Trim() -match '^(IP-[A-Z0-9]+|[A-Z0-9\-_]+(?:-PC|_WORK)?)$' -and
                    $i + 1 -lt $lines.Count -and
                    $lines[$i + 1].Trim() -match '^(Managed Server|Managed Workstation|Managed Network Switch|Managed Network Firewall|FIREWALL|HYPERVISOR|APPLICATION|BACKUP|Domain Controller|Windows Server)$') {
                    # Next line starts a new device entry, reset counter
                    $configEntryLineCount = 0
                    continue
                }
                
                # Stop multi-line entry after 20 lines (should cover one device entry)
                if ($configEntryLineCount -ge 20) {
                    $inMultiLineConfigEntry = $false
                    $configEntryLineCount = 0
                }
            }
            
            # Stop skipping when we hit a new major section header or meaningful content
            # Also stop if we see "Search" at the end of config section (common pattern)
            if ($line -match '^(Additional Details|Knowledge Base|Resources|Team|Ticket Where|Board Icon|Ticket Type|Notes|Discussion|Resolution|Request ID|Impact|Request Status|Source IP|Destination IP|Contact:|Resources:|Cc:|Tasks|Attachments|Category|Subcategory|Allow all clients|Do not show|Drag a pod here|---)' -or
                ($line -match '^[A-Z][a-z]+\s*:' -and -not ($line -match 'Config|Configuration')) -or
                ($line.Trim() -eq '---' -and $prevLine.Trim().Length -gt 0) -or
                ($line.Trim() -eq 'Search' -and $prevLine.Trim().Length -eq 0)) {  # "Search" at end of config section
                $skipConfigSection = $false
                $configListStarted = $false
                $inMultiLineConfigEntry = $false
                $configEntryLineCount = 0
                $blankLineCount = 0
                # Don't add the section header line itself if it's just a label
                if ($line -match '^\s*[A-Z][a-z]+\s*:\s*$' -or $line.Trim() -eq 'Search') {
                    continue
                }
                # Add meaningful section headers
                $filteredLines += $line
                continue
            }
            
            # Skip lines that look like configuration entries (single-line patterns)
            if ($line -match '^\s*[A-Z0-9\-_]+\s+(Managed Workstation|Managed Server|APPLICATION|BACKUP|Managed Network Switch|FIREWALL|HYPERVISOR|Managed Network Firewall)' -or
                $line -match '^\s*[A-Z0-9\-_]+\s+[A-Z0-9]+\s+[A-Z0-9\-]+\s+\d{1,2}/\d{1,2}/\d{2,4}' -or
                $line -match '^\s*[A-Z0-9\-_]+\s+(Managed Workstation|Managed Server|FIREWALL)\s+[A-Z0-9]+\s+[A-Z0-9\-]+\s+[A-Z][a-z]+\s+[A-Z][a-z]+\s+\d{1,2}/\d{1,2}/\d{2,4}' -or
                $line -match '^\s*[A-Z0-9\-_]+\s+(Managed Workstation|Managed Server|FIREWALL)\s+[A-Z0-9]+\s+Active\s+[0-9A-F]{2}-[0-9A-F]{2}-' -or
                $line -match '^\s*[A-Z0-9\-_]+\s+(Managed Workstation|Managed Server|FIREWALL)\s+[A-Z0-9]+\s+Active\s+\d+$' -or
                $line -match '^\s*[A-Z0-9\-_]+\s+(Managed Workstation|Managed Server|FIREWALL)\s+[A-Z0-9]+\s+[A-Z]\\[A-Z0-9]' -or
                $line -match '^\s*[A-Z0-9\-_]+\s+(Managed Workstation|Managed Server|FIREWALL)\s+[A-Z0-9]+\s+\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}' -or
                $line -match '^\s*[A-Z0-9\-_]+\s+(Managed Workstation|Managed Server|FIREWALL)\s+[A-Z0-9]+\s+Microsoft Windows' -or
                $line -match '^\s*[A-Z0-9\-_]+\s+(Managed Workstation|Managed Server|FIREWALL)\s+[A-Z0-9]+\s+macOS' -or
                $line -match '^\s*[A-Z0-9\-_]+\s+(Managed Workstation|Managed Server|FIREWALL)\s+[A-Z0-9]+\s+[A-Z0-9\-]+\s+[A-Z][a-z]+\s+[A-Z][a-z]+\s+\d{1,2}/\d{1,2}/\d{2,4}\s+Active' -or
                $line -match '^\s*[A-Z0-9\-_]+\s+(Managed Workstation|Managed Server|FIREWALL)\s+[A-Z0-9]+\s+\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\s+Microsoft Windows' -or
                ($line -match '^\s*[A-Z0-9\-_]+\s+' -and ($line -split '\s+').Count -ge 4 -and $line -match '(Managed|Active|Microsoft Windows|macOS|FIREWALL|HYPERVISOR)')) {
                $blankLineCount = 0
                continue
            }
            
            # Skip device type lines (Managed Server, Managed Workstation, etc.) when in multi-line entry
            if ($inMultiLineConfigEntry -and $line.Trim() -match '^(Managed Server|Managed Workstation|Managed Network Switch|Managed Network Firewall|FIREWALL|HYPERVISOR|APPLICATION|BACKUP|Domain Controller|Windows Server)$') {
                continue
            }
            
            # Skip lines that look like config entry data (UUIDs, serial numbers, model numbers, dates, status, MAC addresses, IPs, OS versions)
            # Also skip device IDs when in config section (they might appear standalone)
            if ($inMultiLineConfigEntry -and (
                $line.Trim() -match '^(IP-[A-Z0-9]+|[A-Z0-9\-_]+(?:-PC|_WORK|DT\d{4}-\d{2})?)$' -or  # Device IDs
                $line.Trim() -match '^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$' -or  # UUID
                $line.Trim() -match '^(t3\.|c7i-|ec2[a-z0-9-]+)$' -or  # AWS instance types
                $line.Trim() -match '^\d{1,2}/\d{1,2}/\d{2,4}$' -or  # Dates
                $line.Trim() -match '^(Active|Inactive)$' -or  # Status
                $line.Trim() -match '^[0-9A-F]{2}-[0-9A-F]{2}-[0-9A-F]{2}-[0-9A-F]{2}-[0-9A-F]{2}-[0-9A-F]{2}$' -or  # MAC addresses
                $line.Trim() -match '^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$' -or  # IP addresses
                $line.Trim() -match '^Microsoft Windows.*x64$' -or  # OS versions
                $line.Trim() -match '^[A-Z]\\[A-Z0-9]+$' -or  # Domain\username
                $line.Trim() -match '^\d{5,}$' -or  # Numeric IDs
                $line.Trim() -match '^(Parent|Child)$' -or  # Relationship indicators
                $line.Trim() -match '^[A-Z]{2}$' -or  # State codes
                $line.Trim() -match '^\d{5}$' -or  # ZIP codes
                $line.Trim() -match '^[A-Z][a-z]+\s+[A-Z][a-z]+\s+[A-Z][a-z]+$' -or  # Full names (3 words)
                $line.Trim() -match '^[A-Z][a-z]+\s+[A-Z][a-z]+$' -or  # Two-word names
                $line.Trim() -match '^\d{3}\s+W\s+[A-Z]' -or  # Addresses like "700 W Virginia"
                $line.Trim() -match '^Suite\s+\d+$' -or  # Suite numbers
                $line.Trim() -match '^[A-Z][a-z]+\s+&\s+[A-Z]' -or  # Company names like "Miller & Miller"
                $line.Trim() -match '^[A-Z][a-z]+\s+[A-Z][a-z]+\s+St' -or  # Street names
                $line.Trim() -match '^[A-Z][a-z]+$' -and $line.Trim() -match '^(Milwaukee|Wisconsin|United States)$')) {  # City/State names
                continue
            }
            
            # Also skip device IDs when in config section (even if not in multi-line entry mode)
            if ($skipConfigSection -and $line.Trim() -match '^(IP-[A-Z0-9]+|[A-Z0-9\-_]+(?:-PC|_WORK|DT\d{4}-\d{2})?)$') {
                # Check if next line is a device type - if so, enter multi-line mode
                if (-not [string]::IsNullOrWhiteSpace($nextLine) -and
                    $nextLine.Trim() -match '^(Managed Server|Managed Workstation|Managed Network Switch|Managed Network Firewall|FIREWALL|HYPERVISOR|APPLICATION|BACKUP|Domain Controller|Windows Server)$') {
                    $inMultiLineConfigEntry = $true
                    $configEntryLineCount = 0
                }
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
            
            # Track blank lines
            if ($line.Trim().Length -eq 0) {
                $blankLineCount++
                if ($blankLineCount -ge 3) {
                    # Multiple blank lines might indicate end of config section
                    # But continue skipping to be safe
                    continue
                }
            } else {
                $blankLineCount = 0
            }
            
            # If we hit meaningful content that doesn't look like config, stop skipping
            if ($line.Trim().Length -gt 20 -and 
                -not ($line -match '^\s*[A-Z0-9\-_]+\s*$') -and  # Not just a device ID
                -not ($line -match '^\s*[A-Z0-9\-_]+\s+(Managed|FIREWALL|HYPERVISOR)') -and  # Not device type
                -not ($line.Trim() -match '^(Managed Server|Managed Workstation|FIREWALL|HYPERVISOR)$') -and  # Not standalone device type
                -not ($line.Trim() -match '^[0-9a-f]{8}-') -and  # Not UUID
                -not ($line.Trim() -match '^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$') -and  # Not IP
                -not ($line.Trim() -match '^\d{1,2}/\d{1,2}/\d{2,4}$')) {  # Not date
                # This looks like real content, stop skipping
                $skipConfigSection = $false
                $configListStarted = $false
                $inMultiLineConfigEntry = $false
                $configEntryLineCount = 0
                $blankLineCount = 0
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



