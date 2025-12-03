function Get-SettingsPath {
    $dir = Join-Path ([Environment]::GetFolderPath('ApplicationData')) 'ExchangeOnlineAnalyzer'
    if (-not (Test-Path $dir)) {
        try {
            New-Item -ItemType Directory -Path $dir -Force -ErrorAction Stop | Out-Null
        } catch {
            Write-Error "Failed to create settings directory: $($_.Exception.Message)"
            throw
        }
    }
    return (Join-Path $dir 'settings.json')
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
    }
}

function Save-AppSettings {
    param([Parameter(Mandatory=$true)][object]$Settings)
    try {
        $path = Get-SettingsPath
        $json = $Settings | ConvertTo-Json -Depth 4

        # Ensure the directory exists before writing
        $dir = Split-Path -Parent $path
        if (-not (Test-Path $dir)) {
            New-Item -ItemType Directory -Path $dir -Force -ErrorAction Stop | Out-Null
        }

        # Write to a temp file first, then move (atomic operation)
        $tempPath = "$path.tmp"
        $json | Out-File -FilePath $tempPath -Encoding utf8 -ErrorAction Stop
        Move-Item -Path $tempPath -Destination $path -Force -ErrorAction Stop

        Write-Verbose "Settings saved successfully to: $path"
        return $true
    } catch {
        Write-Error "Failed to save settings to $path : $($_.Exception.Message)"
        return $false
    }
}

function New-AIReadme {
    param(
        [Parameter(Mandatory=$false)]
        [object]$Settings
    )
    
    if (-not $Settings) {
        $Settings = Get-AppSettings
    }
    
    # Parse comma-separated lists
    $adminUsers = if ($Settings.AdminUsernames) { ($Settings.AdminUsernames -split ',' | ForEach-Object { $_.Trim() }) -join ', ' } else { '[admin], [service_account], [rmm_account]' }
    $internalTeams = if ($Settings.InternalTeamDisplayNames) { ($Settings.InternalTeamDisplayNames -split ',' | ForEach-Object { $_.Trim() }) -join ', ' } else { 'Managed Services' }
    $authorizedISPs = if ($Settings.AuthorizedISPs) { ($Settings.AuthorizedISPs -split ',' | ForEach-Object { $_.Trim() }) -join ', ' } else { 'Comcast, Charter, CenturyLink, Verizon, Brightspeed, AT&T, T-Mobile' }
    $inFlightWiFi = if ($Settings.InFlightWiFiProviders) { ($Settings.InFlightWiFiProviders -split ',' | ForEach-Object { $_.Trim() }) -join ', ' } else { 'Anuvu, Gogo, Viasat, Panasonic Avionics' }
    $servicePrincipals = if ($Settings.ServicePrincipalNames) { ($Settings.ServicePrincipalNames -split ',' | ForEach-Object { $_.Trim() }) -join ', ' } else { 'Microsoft Graph Command Line Tools' }
    $knownAdmins = if ($Settings.KnownAdmins) { ($Settings.KnownAdmins -split ',' | ForEach-Object { $_.Trim() }) -join ', ' } else { '' }
    
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

Subject: Security Alert: Ticket #[Ticket Number] - [Brief Subject]



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

function Export-AppSettings {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ExportPath
    )
    try {
        $settings = Get-AppSettings
        $json = $settings | ConvertTo-Json -Depth 4
        $json | Out-File -FilePath $ExportPath -Encoding utf8 -ErrorAction Stop
        return $true
    } catch {
        Write-Error "Failed to export settings: $($_.Exception.Message)"
        return $false
    }
}

function Import-AppSettings {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ImportPath
    )
    try {
        if (-not (Test-Path $ImportPath)) {
            Write-Error "Import file not found: $ImportPath"
            return $false
        }
        $raw = Get-Content -Path $ImportPath -Raw -ErrorAction Stop
        $imported = $raw | ConvertFrom-Json
        return Save-AppSettings -Settings $imported
    } catch {
        Write-Error "Failed to import settings: $($_.Exception.Message)"
        return $false
    }
}

Export-ModuleMember -Function Get-AppSettings,Save-AppSettings,Get-SettingsPath,New-AIReadme,Export-AppSettings,Import-AppSettings


