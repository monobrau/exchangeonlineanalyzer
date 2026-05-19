param(
    [int]$ClientNumber,
    [string]$ScriptRoot,
    [string]$InvestigatorName,
    [string]$CompanyName,
    [int]$DaysBack,
    [string]$ReportSelectionsFile,
    [string]$StatusFile,
    [string]$ResultFile,
    [string]$CommandDir,
    [string[]]$SelectedUsers = @()
)

# CRITICAL: Set error action preference immediately after param block
$ErrorActionPreference = "Continue"


# Pause immediately to see if script starts at all
Write-Host "==========================================" -ForegroundColor Green
Write-Host "Worker script starting..." -ForegroundColor Green
Write-Host "==========================================" -ForegroundColor Green
Write-Host "Parameters received:" -ForegroundColor Green
Write-Host "  ClientNumber: $ClientNumber" -ForegroundColor Gray
Write-Host "  ScriptRoot: $ScriptRoot" -ForegroundColor Gray
Write-Host "  StatusFile: $StatusFile" -ForegroundColor Gray
Write-Host "  ResultFile: $ResultFile" -ForegroundColor Gray
Write-Host "  CommandDir: $CommandDir" -ForegroundColor Gray
Start-Sleep -Seconds 3

function Write-Status {
    param([string]$Message)
    try {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        "[$timestamp] $Message" | Out-File -FilePath $StatusFile -Append -Encoding UTF8 -ErrorAction SilentlyContinue
        Write-Host "[Client $ClientNumber] $Message" -ForegroundColor Cyan
    } catch {
        Write-Host "[Client $ClientNumber] $Message" -ForegroundColor Cyan
    }
}

function Write-CommandResponse {
    param([string]$Response)
    try {
        $responseFile = Join-Path $CommandDir "Client$($ClientNumber)_Response.txt"
        $Response | Out-File -FilePath $responseFile -Encoding UTF8 -Force -ErrorAction SilentlyContinue
    } catch {
        Write-Host "WARNING: Could not write command response: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

# Write initial error to result file immediately in case of early failure
Write-Host "Writing initial status to result file..." -ForegroundColor Gray
try {
    if ($ResultFile) {
        "STARTING" | Out-File -FilePath $ResultFile -Encoding UTF8 -ErrorAction SilentlyContinue
        Write-Host "Result file written successfully" -ForegroundColor Green
    } else {
        Write-Host "WARNING: ResultFile parameter is null!" -ForegroundColor Red
    }
} catch {
    Write-Host "WARNING: Could not write result file: $($_.Exception.Message)" -ForegroundColor Yellow
    Write-Host "ResultFile value: $ResultFile" -ForegroundColor Yellow
}
Start-Sleep -Seconds 1

Write-Host "Entering main try block..." -ForegroundColor Green
try {
    # DEBUGGING: Pause at start to see any immediate errors
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host "CLIENT $ClientNumber - PowerShell Session" -ForegroundColor Cyan
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host "PowerShell session starting for Client $ClientNumber..." -ForegroundColor Yellow
    Write-Host "ScriptRoot: $ScriptRoot" -ForegroundColor Gray
    Write-Host "StatusFile: $StatusFile" -ForegroundColor Gray
    Write-Host "ResultFile: $ResultFile" -ForegroundColor Gray
    Write-Host "CommandDir: $CommandDir" -ForegroundColor Gray
    Write-Host ""
    
    # Set window title
    try {
        $Host.UI.RawUI.WindowTitle = "Client $ClientNumber - Waiting for Authentication Commands"
        Write-Host "Window title set successfully" -ForegroundColor Gray
    } catch {
        Write-Host "WARNING: Could not set window title: $($_.Exception.Message)" -ForegroundColor Yellow
    }
    
    # Initialize status file IMMEDIATELY
    try {
        "STARTING" | Out-File -FilePath $ResultFile -Encoding UTF8 -ErrorAction Stop
        Write-Host "Result file initialized: $ResultFile" -ForegroundColor Green
    } catch {
        $errMsg = "CRITICAL: Could not write result file: $($_.Exception.Message)"
        Write-Host $errMsg -ForegroundColor Red
        Write-Host "Result file path: $ResultFile" -ForegroundColor Red
        Write-Host "Directory exists: $((Test-Path (Split-Path $ResultFile -Parent)))" -ForegroundColor Red
        Start-Sleep -Seconds 10
        exit 1
    }
    
    # Initialize status file
    try {
        Write-Status "Client $ClientNumber PowerShell session started"
        Write-Host "Status file initialized successfully" -ForegroundColor Green
    } catch {
        Write-Host "WARNING: Could not write status file: $($_.Exception.Message)" -ForegroundColor Yellow
        Write-Host "Status file path: $StatusFile" -ForegroundColor Yellow
    }
    
    Write-Host "This window is associated with Client $ClientNumber" -ForegroundColor Yellow
    Write-Host "Waiting for authentication commands from GUI..." -ForegroundColor Yellow
    Write-Host ""
    
    # Create isolated cache directory for this client
    Write-Host "Creating cache directory..." -ForegroundColor Gray
    try {
        $cacheDir = Join-Path $env:TEMP "ExchangeOnlineAnalyzer_Client$ClientNumber_Cache_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
        $null = New-Item -ItemType Directory -Path $cacheDir -Force -ErrorAction Stop
        $env:IDENTITY_SERVICE_CACHE_DIR = $cacheDir
        $env:MSAL_CACHE_DIR = $cacheDir
        $env:AZURE_IDENTITY_DISABLE_BROKER = "true"
        $env:MSAL_DISABLE_BROKER = "1"
        $env:MSAL_EXPERIMENTAL_DISABLE_BROKER = "1"
        Write-Status "Using isolated cache directory: $cacheDir"
        Write-Host "Cache directory created: $cacheDir" -ForegroundColor Green
        Write-Host ""
    } catch {
        $errMsg = "CRITICAL: Failed to create cache directory: $($_.Exception.Message)"
        Write-Host $errMsg -ForegroundColor Red
        Write-Status $errMsg
        "ERROR: $errMsg" | Out-File -FilePath $ResultFile -Encoding UTF8 -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 10
        exit 1
    }
    
    # Import required modules
    Write-Status "Importing modules..."
    Write-Host "Importing modules..." -ForegroundColor Cyan
    Write-Host "ScriptRoot path: $ScriptRoot" -ForegroundColor Gray
    
    # Check if ScriptRoot exists
    if (-not (Test-Path $ScriptRoot)) {
        $errorMsg = "CRITICAL: ScriptRoot directory does not exist: $ScriptRoot"
        Write-Host $errorMsg -ForegroundColor Red
        Write-Status $errorMsg
        "ERROR: $errorMsg" | Out-File -FilePath $ResultFile -Encoding UTF8 -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 10
        exit 1
    }
    
    # Import Logging module
    Write-Host "Importing Logging module..." -ForegroundColor Gray
    try {
        $loggingPath = Join-Path $ScriptRoot "Modules\Logging.psm1"
        if (Test-Path $loggingPath) {
            Import-Module $loggingPath -Force -ErrorAction SilentlyContinue
            Write-Host "Logging module imported" -ForegroundColor Green
        } else {
            Write-Host "WARNING: Logging.psm1 not found at $loggingPath" -ForegroundColor Yellow
        }
        try { Initialize-Logger -MinLevel Info -ConsoleOutput $true -SessionId "Client$ClientNumber" -CompanyName $CompanyName -Component ExportUtils | Out-Null } catch {}
    } catch {
        Write-Host "WARNING: Failed to import Logging module: $($_.Exception.Message)" -ForegroundColor Yellow
    }
    
    # ROBUSTNESS: Better error handling for critical module import
    Write-Host "Importing ExportUtils module..." -ForegroundColor Gray
    try {
        $exportUtilsPath = Join-Path $ScriptRoot "Modules\ExportUtils.psm1"
        if (-not (Test-Path $exportUtilsPath)) {
            throw "ExportUtils.psm1 not found at $exportUtilsPath"
        }
        Import-Module $exportUtilsPath -Force -ErrorAction Stop
        Write-Host "ExportUtils module imported successfully" -ForegroundColor Green
    } catch {
        $errorMsg = "CRITICAL: Failed to import ExportUtils.psm1 - $($_.Exception.Message)"
        Write-Host $errorMsg -ForegroundColor Red
        Write-Host "Full error: $($_.Exception | Out-String)" -ForegroundColor Red
        Write-Status $errorMsg
        "ERROR: $errorMsg`n`nFull details:`n$($_.Exception | Out-String)" | Out-File -FilePath $ResultFile -Encoding UTF8 -ErrorAction SilentlyContinue
        Write-Host "Press any key to exit..."
        try {
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        } catch {
            Start-Sleep -Seconds 10
        }
        exit 1
    }
    
    Write-Host "Importing GraphOnline module..." -ForegroundColor Gray
    Import-Module "$ScriptRoot\Modules\GraphOnline.psm1" -Force -ErrorAction SilentlyContinue
    Write-Host "Importing BrowserIntegration module..." -ForegroundColor Gray
    Import-Module "$ScriptRoot\Modules\BrowserIntegration.psm1" -Force -ErrorAction SilentlyContinue
    Write-Host "Importing Settings module..." -ForegroundColor Gray
    Import-Module "$ScriptRoot\Modules\Settings.psm1" -Force -ErrorAction SilentlyContinue
    Write-Host "Importing GraphAppCredential module (for WCM app-only auth)..." -ForegroundColor Gray
    $graphAppMod = Join-Path $ScriptRoot "Modules\GraphAppCredential.psm1"
    if (Test-Path $graphAppMod) { Import-Module $graphAppMod -Force -ErrorAction SilentlyContinue }
    Write-Status "Modules imported successfully"
    Write-Host "All modules imported successfully" -ForegroundColor Green
    Write-Host ""
    
    # CRITICAL: Disconnect any existing Graph session before starting
    # This ensures each tenant starts with a fresh authentication state
    # Even though each tenant has its own process, WAM might cache credentials globally
    try {
        $existingContext = Get-MgContext -ErrorAction SilentlyContinue
        if ($existingContext) {
            Write-Host "Found existing Graph session - disconnecting to ensure fresh authentication..." -ForegroundColor Yellow
            Disconnect-MgGraph -ErrorAction SilentlyContinue
            Start-Sleep -Milliseconds 500  # Brief pause to ensure disconnection completes
        }
    } catch {
        # Ignore errors - no session exists or module not loaded yet
    }
    
    # Disconnect any existing Exchange session for clean slate (module may not be loaded yet)
    try {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
        Get-PSSession | Where-Object { $_.ConfigurationName -eq "Microsoft.Exchange" } | Remove-PSSession -ErrorAction SilentlyContinue
    } catch {
        # Ignore - Exchange module may not be loaded yet
    }
    
    # Load report selections from JSON
    $reportSelections = @{}
    if (Test-Path $ReportSelectionsFile) {
        $jsonObj = Get-Content $ReportSelectionsFile -Raw | ConvertFrom-Json
        $reportSelections = @{
            IncludeMessageTrace = if ($null -ne $jsonObj.IncludeMessageTrace) { $jsonObj.IncludeMessageTrace } else { $false }
            IncludeInboxRules = if ($null -ne $jsonObj.IncludeInboxRules) { $jsonObj.IncludeInboxRules } else { $false }
            IncludeTransportRules = if ($null -ne $jsonObj.IncludeTransportRules) { $jsonObj.IncludeTransportRules } else { $false }
            IncludeMailFlowConnectors = if ($null -ne $jsonObj.IncludeMailFlowConnectors) { $jsonObj.IncludeMailFlowConnectors } else { $false }
            IncludeMailboxForwarding = if ($null -ne $jsonObj.IncludeMailboxForwarding) { $jsonObj.IncludeMailboxForwarding } else { $false }
            IncludeAuditLogs = if ($null -ne $jsonObj.IncludeAuditLogs) { $jsonObj.IncludeAuditLogs } else { $false }
            IncludeConditionalAccessPolicies = if ($null -ne $jsonObj.IncludeConditionalAccessPolicies) { $jsonObj.IncludeConditionalAccessPolicies } else { $false }
            IncludeAppRegistrations = if ($null -ne $jsonObj.IncludeAppRegistrations) { $jsonObj.IncludeAppRegistrations } else { $false }
            IncludeSignInLogs = (($jsonObj.IncludeSignInLogs -eq $true) -or ("$($jsonObj.IncludeSignInLogs)" -match '^(?i)true$|^(?i)yes$|^1$'))
            IncludeIntuneDevices = if ($null -ne $jsonObj.IncludeIntuneDevices -and $jsonObj.IncludeIntuneDevices -ne "") { [bool]$jsonObj.IncludeIntuneDevices } else { $false }
            IncludeMfaCoverage = if ($null -ne $jsonObj.IncludeMfaCoverage -and $jsonObj.IncludeMfaCoverage -ne "") { [bool]$jsonObj.IncludeMfaCoverage } else { $false }
            IncludeSharePointActivity = if ($null -ne $jsonObj.IncludeSharePointActivity) { $jsonObj.IncludeSharePointActivity } else { $true }
            IncludeOneDriveActivity = if ($null -ne $jsonObj.IncludeOneDriveActivity) { $jsonObj.IncludeOneDriveActivity } else { $true }
            IncludeTeamsActivity = if ($null -ne $jsonObj.IncludeTeamsActivity) { $jsonObj.IncludeTeamsActivity } else { $true }
            IncludeSharePointSharing = if ($null -ne $jsonObj.IncludeSharePointSharing) { $jsonObj.IncludeSharePointSharing } else { $true }
            IncludeSecurityAlerts = if ($null -ne $jsonObj.IncludeSecurityAlerts) { $jsonObj.IncludeSecurityAlerts } else { $true }
            IncludeSecurityIncidents = if ($null -ne $jsonObj.IncludeSecurityIncidents) { $jsonObj.IncludeSecurityIncidents } else { $false }
            IncludeAnonymousSharePointSharing = if ($null -ne $jsonObj.IncludeAnonymousSharePointSharing) { $jsonObj.IncludeAnonymousSharePointSharing } else { $true }
            IncludeSharePointFileSharingLinks = if ($null -ne $jsonObj.IncludeSharePointFileSharingLinks) { $jsonObj.IncludeSharePointFileSharingLinks } else { $true }
            IncludeDLPViolations = if ($null -ne $jsonObj.IncludeDLPViolations) { $jsonObj.IncludeDLPViolations } else { $true }
            IncludeUnifiedAuditLogs = if ($null -ne $jsonObj.IncludeUnifiedAuditLogs) { $jsonObj.IncludeUnifiedAuditLogs } else { $false }
            IncludeSharePointOneDriveFileActions = if ($null -ne $jsonObj.IncludeSharePointOneDriveFileActions) { $jsonObj.IncludeSharePointOneDriveFileActions } else { $true }
            SignInLogsDaysBack = if ($null -ne $jsonObj.SignInLogsDaysBack) { $jsonObj.SignInLogsDaysBack } else { 7 }
            MessageTraceDaysBack = if ($null -ne $jsonObj.MessageTraceDaysBack) { $jsonObj.MessageTraceDaysBack } else { 10 }
        }
    }
    
    $graphAuthenticated = $false
    $exchangeAuthenticated = $false
    $tenantDisplayName = "Client$ClientNumber"
    
    # Main command loop - wait for commands from GUI
    $commandFile = Join-Path $CommandDir "Client$($ClientNumber)_Command.txt"
    $pollInterval = 200  # milliseconds
    
    Write-Host "Ready! Waiting for Graph Auth command from GUI..." -ForegroundColor Green
    Write-Status "Ready! Waiting for Graph Auth command from GUI..."
    Write-Host "Command file: $commandFile" -ForegroundColor Gray
    Write-Host "Polling every $pollInterval ms for commands..." -ForegroundColor Gray
    Write-Host ""
    Write-Host "Worker script is now running and ready!" -ForegroundColor Green
    Write-Host "This window will stay open. Do not close it manually." -ForegroundColor Yellow
    Write-Host ""
    
    Write-Status "Command polling loop started - ready to receive commands"
    $pollCount = 0
    while ($true) {
        $pollCount++
        if ($pollCount % 100 -eq 0) {
            Write-Host "Still polling... (checked $pollCount times)" -ForegroundColor DarkGray
        }
        
        if (Test-Path $commandFile) {
            Write-Host "==========================================" -ForegroundColor Yellow
            Write-Host "Command file detected! Reading command..." -ForegroundColor Yellow
            Write-Host "Command file path: $commandFile" -ForegroundColor Cyan
            Start-Sleep -Milliseconds 300  # Brief delay to ensure file is fully written
            # SECURITY: Use safe command file reading with validation
            if (Get-Command Read-CommandFile -ErrorAction SilentlyContinue) {
                $command = Read-CommandFile -CommandFilePath $commandFile
            } else {
                $command = (Get-Content $commandFile -Raw -ErrorAction SilentlyContinue).Trim().TrimStart([char]0xFEFF)
            }
            if ($command) {
                $commandType = if ($command -match '^([^|]+)') { $Matches[1] } else { "Unknown" }
                Write-Host "Command type: $commandType" -ForegroundColor Cyan
                Write-Host "Command length: $($command.Length)" -ForegroundColor Gray
            } else {
                Write-Host "No valid command found in file" -ForegroundColor Yellow
            }
            Remove-Item $commandFile -Force -ErrorAction SilentlyContinue
            Write-Host "Command file removed" -ForegroundColor Gray
            
            if ($command -eq "GRAPH_AUTH" -or $command -like "GRAPH_AUTH|*") {
                Write-Host "==========================================" -ForegroundColor Yellow
                Write-Host "GRAPH AUTHENTICATION COMMAND RECEIVED" -ForegroundColor Yellow
                Write-Host "==========================================" -ForegroundColor Yellow
                Write-Status "Graph authentication command received"
                Write-CommandResponse "GRAPH_AUTH_STARTED"
                
                # Clear any existing sessions and token caches
                # NOTE: Each tenant runs in its own isolated PowerShell process, so disconnecting only affects
                # this tenant's session, not other tenants' sessions running in separate processes.
                Write-Status "Clearing existing sessions and token caches..."
                Write-Host "Clearing existing sessions and token caches..." -ForegroundColor Cyan
                
                # Disconnect Graph session first (only if one exists in this process)
                # CRITICAL: This must happen BEFORE clearing cache to ensure session is fully released
                try { 
                    $mgContext = Get-MgContext -ErrorAction SilentlyContinue
                    if ($mgContext) {
                        Write-Host "Found existing Graph context - Tenant: $($mgContext.TenantId), Account: $($mgContext.Account)" -ForegroundColor Yellow
                        Disconnect-MgGraph -ErrorAction SilentlyContinue 
                        Write-Host "Disconnected existing Graph session for this tenant" -ForegroundColor Gray
                        # Wait for session to fully release before re-auth (reduces reuse of cached credentials)
                        Start-Sleep -Milliseconds 1500
                    } else {
                        Write-Host "No existing Graph session to disconnect" -ForegroundColor Gray
                    }
                } catch {
                    # Ignore errors - session may not exist
                }
                
                # Clear Graph token cache and reset GraphSession singleton
                try {
                    $graphSession = [Microsoft.Graph.PowerShell.Authentication.GraphSession]::Instance
                    if ($graphSession) {
                        if ($graphSession.AuthContext) {
                            $graphSession.AuthContext.ClearTokenCache()
                            Write-Host "Cleared Graph token cache" -ForegroundColor Gray
                        }
                        # Try to reset the session instance to ensure fresh state
                        try {
                            $graphSession.Reset() | Out-Null
                            Write-Host "Reset GraphSession instance" -ForegroundColor Gray
                        } catch {
                            # Reset() method may not exist in all versions - ignore if not available
                        }
                    }
                } catch {
                    # Ignore errors clearing token cache
                }
                
                # Clear ALL files in the MSAL cache directory (not just "*cache*" files)
                # This ensures no cached tokens from previous tenants remain
                try {
                    if ($env:MSAL_CACHE_DIR -and (Test-Path $env:MSAL_CACHE_DIR)) {
                        $allCacheFiles = Get-ChildItem -Path $env:MSAL_CACHE_DIR -File -Recurse -ErrorAction SilentlyContinue
                        $fileCount = $allCacheFiles.Count
                        foreach ($file in $allCacheFiles) {
                            Remove-Item $file.FullName -Force -ErrorAction SilentlyContinue
                        }
                        Write-Host "Cleared all files from MSAL cache directory ($fileCount files)" -ForegroundColor Gray
                    }
                } catch {
                    # Ignore errors clearing MSAL cache
                }
                
                # Also clear IDENTITY_SERVICE_CACHE_DIR if it exists
                try {
                    if ($env:IDENTITY_SERVICE_CACHE_DIR -and (Test-Path $env:IDENTITY_SERVICE_CACHE_DIR)) {
                        $allIdentityFiles = Get-ChildItem -Path $env:IDENTITY_SERVICE_CACHE_DIR -File -Recurse -ErrorAction SilentlyContinue
                        $identityFileCount = $allIdentityFiles.Count
                        foreach ($file in $allIdentityFiles) {
                            Remove-Item $file.FullName -Force -ErrorAction SilentlyContinue
                        }
                        Write-Host "Cleared all files from Identity cache directory ($identityFileCount files)" -ForegroundColor Gray
                    }
                } catch {
                    # Ignore errors clearing Identity cache
                }
                
                # Graph Authentication
                # NOTE: Microsoft Graph Authentication Behavior
                # Microsoft.Graph.Authentication version 2.33.0+ defaults to using WAM (Web Account Manager) on Windows,
                # which shows a popup dialog instead of opening the system browser. Unlike Connect-ExchangeOnline which
                # has a -DisableWAM parameter, Connect-MgGraph does not have this option. Environment variables to disable
                # WAM are set below, but newer module versions may ignore them. The authentication will still work via
                # the WAM popup if the browser doesn't open automatically.
                # TODO: Revisit this implementation if/when Microsoft.Graph.Authentication adds a -DisableWAM parameter
                #       or provides another mechanism to force system browser authentication.
                Write-Host ""
                Write-Host "Starting Microsoft Graph authentication..." -ForegroundColor Yellow
                Write-Host "Note: A popup may appear instead of your browser (this is a limitation of Microsoft.Graph.Authentication)." -ForegroundColor Yellow
                Write-Host ""
                Write-Status "Waiting for authentication window to appear (this may take 10-30 seconds)..."

                # Disable broker/WAM so authentication uses the system browser instead of an embedded popup
                $env:AZURE_IDENTITY_DISABLE_BROKER = "true"
                $env:MSAL_DISABLE_BROKER = "1"
                $env:MSAL_EXPERIMENTAL_DISABLE_BROKER = "1"

                # Create a FRESH cache directory for THIS auth attempt (new path = no cached tokens)
                # This is critical to prevent reusing tokens from previous tenants or prior attempts
                $authCacheDir = Join-Path $env:TEMP "ExchangeOnlineAnalyzer_Client$ClientNumber_Auth_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
                try {
                    if (Test-Path $authCacheDir) { Remove-Item -Path $authCacheDir -Recurse -Force -ErrorAction SilentlyContinue }
                    New-Item -ItemType Directory -Path $authCacheDir -Force -ErrorAction Stop | Out-Null
                    $env:MSAL_CACHE_DIR = $authCacheDir
                    $env:IDENTITY_SERVICE_CACHE_DIR = $authCacheDir
                    Write-Host "Using fresh auth cache directory: $authCacheDir" -ForegroundColor Gray
                } catch {
                    Write-Warning "Could not create fresh cache dir, using existing: $($_.Exception.Message)"
                }

                # Clear default MSAL cache location in user profile (in addition to custom cache dir)
                try {
                    $defaultMsalCache = Join-Path $env:LOCALAPPDATA ".IdentityService"
                    if (Test-Path $defaultMsalCache) {
                        Get-ChildItem -Path $defaultMsalCache -Recurse -File -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
                        Write-Host "Cleared default IdentityService cache in user profile" -ForegroundColor Gray
                    }
                } catch {
                    # Ignore errors
                }

                # Clear Microsoft.Graph module's own cache
                try {
                    $graphModuleCache = Join-Path $env:LOCALAPPDATA "Microsoft\Graph"
                    if (Test-Path $graphModuleCache) {
                        Get-ChildItem -Path $graphModuleCache -Recurse -File -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
                        Write-Host "Cleared Microsoft Graph module cache" -ForegroundColor Gray
                    }
                } catch {
                    # Ignore errors
                }

                # Clear Windows WAM (Web Account Manager) token cache
                # This helps prevent reusing cached credentials from previous sessions
                try {
                    $wamCache = Join-Path $env:LOCALAPPDATA "Packages\Microsoft.AAD.BrokerPlugin_cw5n1h2txyewy\AC\TokenBroker\Cache"
                    if (Test-Path $wamCache) {
                        Get-ChildItem -Path $wamCache -Recurse -File -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
                        Write-Host "Cleared WAM token broker cache" -ForegroundColor Gray
                    }
                } catch {
                    # Ignore errors - WAM cache may not exist or may be in use
                }

                # Also try alternative WAM cache location
                try {
                    $wamCache2 = Join-Path $env:LOCALAPPDATA "Packages\Microsoft.AAD.BrokerPlugin_cw5n1h2txyewy\LocalState"
                    if (Test-Path $wamCache2) {
                        Get-ChildItem -Path $wamCache2 -Recurse -File -Filter "*.dat" -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
                        Write-Host "Cleared WAM local state cache" -ForegroundColor Gray
                    }
                } catch {
                    # Ignore errors
                }

                $scopes = @(
                    "AuditLog.Read.All",
                    "User.Read.All",
                    "Directory.Read.All",
                    "Policy.Read.All",
                    "Application.Read.All",
                    "Reports.Read.All"
                )

                # Try WCM (app-only) first when we have tenant ID(s) to try - skip if INTERACTIVE:1 (use browser instead)
                $tenantIdsToTry = @()
                $graphAuthExplicitTenantId = $null
                $useInteractiveOnly = ($command -match 'INTERACTIVE:1')
                if (-not $useInteractiveOnly) {
                    if ($command -match '\|TENANT_ID:([a-fA-F0-9\-]{36})') {
                        $graphAuthExplicitTenantId = $matches[1]
                        $tenantIdsToTry = @($graphAuthExplicitTenantId)
                    } else {
                        $graphAppMod = Join-Path $ScriptRoot "Modules\GraphAppCredential.psm1"
                        if (Test-Path $graphAppMod) {
                            Import-Module $graphAppMod -Force -ErrorAction SilentlyContinue
                            if (Get-Command Get-WCMTenantIds -ErrorAction SilentlyContinue) {
                                $allIds = Get-WCMTenantIds
                                if ($allIds -and $allIds.Count -gt 0) { $tenantIdsToTry = @($allIds) }
                            }
                        }
                    }
                }
                if ($useInteractiveOnly) {
                    Write-Host "Use interactive Graph selected - skipping app credentials, connecting via browser..." -ForegroundColor Cyan
                }
                if ($tenantIdsToTry.Count -gt 0) {
                    if ($graphAuthExplicitTenantId) {
                        Write-Host "Trying app-only for selected tenant $graphAuthExplicitTenantId (Credential Manager target: EOA-GraphApp-$graphAuthExplicitTenantId)..." -ForegroundColor Cyan
                    } else {
                        Write-Host "Trying app-only for $($tenantIdsToTry.Count) tenant(s) listed in Credential Manager..." -ForegroundColor Cyan
                    }
                }
                foreach ($tid in $tenantIdsToTry) {
                    try {
                        $wcmErr = $null
                        $wcmToken = Get-GraphAppTokenFromWCM -TenantId $tid -FailureVariable wcmErr
                        if (-not $wcmToken) {
                            if ($wcmErr) {
                                Write-Host "App-only token failed for tenant $tid : $wcmErr" -ForegroundColor Yellow
                                Write-Status "WCM app-only failed: $wcmErr"
                            }
                            continue
                        }
                        $headers = @{ Authorization = "Bearer $wcmToken" }
                        $orgResp = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/organization" -Headers $headers -Method GET -ErrorAction Stop
                        if ($orgResp -and $orgResp.value -and $orgResp.value.Count -gt 0) {
                            $tenantDisplayName = $orgResp.value[0].displayName
                            if (-not $tenantDisplayName) { $tenantDisplayName = "Tenant" }
                            $graphAuthenticated = $true
                            $script:graphTokenFromWCM = $wcmToken
                            $script:currentTenantId = $tid
                            Write-Status "Using app-only credentials from Windows Credential Manager"
                            try {
                                $env:AZURE_IDENTITY_DISABLE_BROKER = "true"
                                $env:MSAL_DISABLE_BROKER = "1"
                                $env:MSAL_EXPERIMENTAL_DISABLE_BROKER = "1"
                                $secGraphTokWcm = ConvertTo-SecureString $wcmToken -AsPlainText -Force
                                Connect-MgGraph -AccessToken $secGraphTokWcm -NoWelcome -ErrorAction Stop
                                Write-Host "Graph PowerShell session established with app-only token (no interactive scopes flow)." -ForegroundColor DarkGreen
                            } catch {
                                Write-Warning "Connect-MgGraph -AccessToken after WCM success failed: $($_.Exception.Message)"
                            }
                            Write-Host "Using app-only credentials from Windows Credential Manager (Tenant: $tid)" -ForegroundColor Green
                            $verifiedDomains = @()
                            try {
                                $domResp = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/domains" -Headers $headers -Method GET -ErrorAction Stop
                                if ($domResp -and $domResp.value) {
                                    $verifiedDomains = $domResp.value | Where-Object { $_.isVerified -eq $true } | ForEach-Object { $_.id }
                                    Write-Host "Found $($verifiedDomains.Count) verified domain(s): $($verifiedDomains -join ', ')" -ForegroundColor Cyan
                                }
                            } catch {}
                            if ($verifiedDomains -and $verifiedDomains.Count -gt 0) {
                                Write-CommandResponse "GRAPH_AUTH_SUCCESS:$tenantDisplayName|TENANT_ID:$tid|DOMAINS:$($verifiedDomains -join ',')"
                            } else {
                                Write-CommandResponse "GRAPH_AUTH_SUCCESS:$tenantDisplayName|TENANT_ID:$tid"
                            }
                            break
                        } else {
                            Write-Host "App-only token was obtained but GET /organization returned no data (app may lack Directory.Read.All or admin consent). Tenant: $tid" -ForegroundColor Yellow
                            Write-Status "WCM: token OK, organization API empty or forbidden"
                        }
                    } catch {
                        Write-Host "App-only Graph verification failed for tenant $tid (token or API): $($_.Exception.Message)" -ForegroundColor Yellow
                        continue
                    }
                }

                if (-not $graphAuthenticated) {
                Write-Host "No WCM credentials found or validation failed. Connecting interactively (browser/popup/WAM)..." -ForegroundColor Yellow
                Write-Host "If you expected app-only: add EOA-GraphApp-<tenantId> in Credential Manager (Import App Creds or Create Graph App), or pick the tenant that matches an existing EOA-GraphApp-* entry." -ForegroundColor Gray
                if ($graphAuthExplicitTenantId -and (Get-Command Get-WCMTenantIds -ErrorAction SilentlyContinue)) {
                    try {
                        $wcmOtherIds = @(Get-WCMTenantIds)
                        if ($wcmOtherIds.Count -gt 0 -and ($wcmOtherIds -notcontains $graphAuthExplicitTenantId)) {
                            Write-Host "Credential Manager has app secrets for different tenant ID(s): $($wcmOtherIds -join ', '). Not for selected $graphAuthExplicitTenantId." -ForegroundColor Yellow
                            Write-Host "Import/create credentials for the selected tenant, or choose the matching tenant in the app-reg dropdown." -ForegroundColor Gray
                        }
                    } catch {}
                }
                try {
                    # Use standard Connect-MgGraph authentication
                    # LIMITATION: Microsoft.Graph.Authentication version 2.33.0+ defaults to WAM (Web Account Manager) on Windows.
                    # Unlike Connect-ExchangeOnline which has a -DisableWAM parameter, Connect-MgGraph does not provide
                    # this option. Environment variables are set below to attempt disabling WAM, but newer module versions
                    # may ignore them. The authentication will still function correctly via the WAM popup if the system
                    # browser doesn't open automatically. This is a known limitation of the Microsoft.Graph.Authentication
                    # module and not a bug in this script.
                    # TODO: Revisit this implementation if/when Microsoft.Graph.Authentication adds a -DisableWAM parameter
                    #       or provides another mechanism to force system browser authentication.
                    # Set environment variables to try to disable WAM (may not work with newer module versions)
                    $env:AZURE_IDENTITY_DISABLE_BROKER = "true"
                    $env:MSAL_DISABLE_BROKER = "1"
                    $env:MSAL_EXPERIMENTAL_DISABLE_BROKER = "1"
                    Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Yellow
                    Connect-MgGraph -Scopes $scopes -ContextScope Process -NoWelcome -ErrorAction Stop
                    $mgContext = Get-MgContext -ErrorAction Stop
                    $graphAuthenticated = $true
                    $script:currentTenantId = $mgContext.TenantId
                    Write-Status "Graph authentication successful! Tenant: $($mgContext.TenantId)"
                    Write-Host "Graph authentication successful!" -ForegroundColor Green
                    Write-Host "Tenant ID: $($mgContext.TenantId)" -ForegroundColor Cyan
                    
                    # Get tenant name
                    try {
                        $ti = $null
                        try { $ti = Get-TenantIdentity } catch {}
                        if ($ti -and $ti.TenantDisplayName) {
                            $tenantDisplayName = $ti.TenantDisplayName
                        } elseif ($ti -and $ti.PrimaryDomain) {
                            $tenantDisplayName = $ti.PrimaryDomain
                        } else {
                            try {
                                $org = Get-MgOrganization -ErrorAction SilentlyContinue | Select-Object -First 1
                                if ($org -and $org.DisplayName) {
                                    $tenantDisplayName = $org.DisplayName
                                }
                            } catch {}
                        }
                    } catch {}
                    
                    Write-Status "Tenant identified as: $tenantDisplayName"
                    Write-Host "Tenant: $tenantDisplayName" -ForegroundColor Cyan

                    # Query all verified domains for the tenant
                    $verifiedDomains = @()
                    try {
                        Write-Host "Querying tenant domains..." -ForegroundColor Gray
                        $domainsResponse = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/domains" -ErrorAction Stop
                        if ($domainsResponse -and $domainsResponse.value) {
                            $verifiedDomains = $domainsResponse.value |
                                               Where-Object { $_.isVerified -eq $true } |
                                               ForEach-Object { $_.id }
                            Write-Host "Found $($verifiedDomains.Count) verified domain(s): $($verifiedDomains -join ', ')" -ForegroundColor Cyan
                        }
                    } catch {
                        Write-Host "Warning: Failed to query tenant domains: $($_.Exception.Message)" -ForegroundColor Yellow
                        Write-Host "Falling back to tenant name as primary domain" -ForegroundColor Yellow
                    }

                    # Build response with tenant name, tenant ID, and domains (tenant ID enables WCM-first on re-auth)
                    if ($verifiedDomains -and $verifiedDomains.Count -gt 0) {
                        $domainsString = $verifiedDomains -join ','
                        Write-CommandResponse "GRAPH_AUTH_SUCCESS:$tenantDisplayName|TENANT_ID:$($mgContext.TenantId)|DOMAINS:$domainsString"
                    } else {
                        Write-CommandResponse "GRAPH_AUTH_SUCCESS:$tenantDisplayName|TENANT_ID:$($mgContext.TenantId)"
                    }
                } catch {
                    Write-Status "ERROR: Graph authentication failed - $($_.Exception.Message)"
                    Write-Host "ERROR: Graph authentication failed - $($_.Exception.Message)" -ForegroundColor Red
                    # If PromptToCreateGraphApp is on, offer to create app and save to WCM
                    $promptToCreate = $false
                    try {
                        Import-Module "$ScriptRoot\Modules\Settings.psm1" -Force -ErrorAction SilentlyContinue
                        $s = Get-AppSettings -ErrorAction SilentlyContinue
                        if ($s -and $s.PromptToCreateGraphApp) { $promptToCreate = $true }
                    } catch {}
                    if ($promptToCreate) {
                        Write-Host ""
                        Write-Host "Create app registration (River Run Security Investigator) and save to Windows Credential Manager? (y/n): " -ForegroundColor Yellow -NoNewline
                        $create = Read-Host
                        if ($create -eq 'y' -or $create -eq 'Y') {
                            try {
                                Write-Host "Running app creation script (browser will open for admin sign-in)..." -ForegroundColor Cyan
                                $appScript = Join-Path $ScriptRoot "New-GraphInboxRulesApp.ps1"
                                if (Test-Path $appScript) {
                                    & $appScript -SaveToWCM
                                    $mgContext = Get-MgContext -ErrorAction SilentlyContinue
                                    if ($mgContext -and $mgContext.TenantId) {
                                        Import-Module (Join-Path $ScriptRoot "Modules\GraphAppCredential.psm1") -Force -ErrorAction SilentlyContinue
                                        $wcmToken = Get-GraphAppTokenFromWCM -TenantId $mgContext.TenantId
                                        if ($wcmToken) {
                                            $graphAuthenticated = $true
                                            $script:graphTokenFromWCM = $wcmToken
                                            $tenantDisplayName = "Tenant"
                                            try { $ti = Get-TenantIdentity -ErrorAction SilentlyContinue; if ($ti -and $ti.TenantDisplayName) { $tenantDisplayName = $ti.TenantDisplayName } elseif ($ti -and $ti.PrimaryDomain) { $tenantDisplayName = $ti.PrimaryDomain } } catch {}
                                            Write-CommandResponse "GRAPH_AUTH_SUCCESS:$tenantDisplayName"
                                            Write-Host "Using app-only credentials from Windows Credential Manager." -ForegroundColor Green
                                        }
                                    }
                                }
                            } catch {
                                Write-Host "App creation failed: $($_.Exception.Message)" -ForegroundColor Red
                            }
                        }
                    }
                    if (-not $graphAuthenticated) {
                        Write-CommandResponse "GRAPH_AUTH_FAILED:$($_.Exception.Message)"
                    }
                }
                }

                Write-Host ""
                Write-Host "Waiting for Exchange Online Auth command from GUI..." -ForegroundColor Green
                Write-Host ""
                
            } elseif ($command -eq "EXCHANGE_AUTH") {
                Write-Host "==========================================" -ForegroundColor Yellow
                Write-Host "EXCHANGE ONLINE AUTHENTICATION COMMAND RECEIVED" -ForegroundColor Yellow
                Write-Host "==========================================" -ForegroundColor Yellow
                Write-Status "Exchange Online authentication command received"
                Write-CommandResponse "EXCHANGE_AUTH_STARTED"
                
                # Disconnect any existing Exchange session to ensure fresh authentication per tenant
                Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
                Get-PSSession | Where-Object { $_.ConfigurationName -eq "Microsoft.Exchange" } | Remove-PSSession -ErrorAction SilentlyContinue
                Start-Sleep -Milliseconds 500  # Allow session to release before re-auth (reduced from 1000ms)
                
                # Exchange Online Authentication
                Write-Host ""
                Write-Host "Connecting to Exchange Online..." -ForegroundColor Yellow
                Write-Host "A browser window will open for authentication - typically 15-60 seconds total." -ForegroundColor Yellow
                Write-Host "Please wait for the browser popup and complete the sign-in." -ForegroundColor Yellow
                Write-Host ""
                Write-Status "Waiting for browser popup (typically 15-60 seconds)..."
    
                try {
                    # Note: Connect-ExchangeOnline may take 15-60s (browser popup + sign-in)
                    # -DisableWAM prevents WAM issues; -SkipLoadingCmdletHelp speeds up connection (ExchangeOnlineManagement v3.3+)
                    $connectParams = @{ ShowBanner = $false; DisableWAM = $true; ErrorAction = 'Stop' }
                    if ((Get-Command Connect-ExchangeOnline -ErrorAction SilentlyContinue).Parameters.Keys -contains 'SkipLoadingCmdletHelp') {
                        $connectParams['SkipLoadingCmdletHelp'] = $true
                    }
                    Connect-ExchangeOnline @connectParams
                    $exchangeAuthenticated = $true
                    Write-Status "Exchange Online authentication successful!"
                    Write-Host "Exchange Online authentication successful!" -ForegroundColor Green
                    Write-CommandResponse "EXCHANGE_AUTH_SUCCESS"
                } catch {
                    Write-Status "ERROR: Exchange Online authentication failed - $($_.Exception.Message)"
                    Write-Host "ERROR: Exchange Online authentication failed - $($_.Exception.Message)" -ForegroundColor Red
                    Write-CommandResponse "EXCHANGE_AUTH_FAILED:$($_.Exception.Message)"
                }
                
                Write-Host ""
                Write-Host "Waiting for Generate Reports command from GUI..." -ForegroundColor Green
                Write-Host ""
                
            } elseif ($command -match "^VALIDATE_USERS") {
                if (-not $graphAuthenticated) {
                    Write-Host "ERROR: Graph authentication must be completed first!" -ForegroundColor Red
                    Write-CommandResponse "VALIDATE_USERS_FAILED:Graph authentication not completed"
                    continue
                }
                
                # Ensure Graph session exists for Get-MgUser (WCM app-only stores token but doesn't call Connect-MgGraph)
                $mgCtx = Get-MgContext -ErrorAction SilentlyContinue
                if (-not $mgCtx) {
                    $tokenToUse = $script:graphTokenFromWCM
                    if (-not $tokenToUse -and $script:currentTenantId -and (Get-Command Get-GraphAppTokenFromWCM -ErrorAction SilentlyContinue)) {
                        $tokenToUse = Get-GraphAppTokenFromWCM -TenantId $script:currentTenantId
                        if ($tokenToUse) { $script:graphTokenFromWCM = $tokenToUse }
                    }
                    if ($tokenToUse) {
                        try {
                            $secToken = ConvertTo-SecureString $tokenToUse -AsPlainText -Force
                            Connect-MgGraph -AccessToken $secToken -NoWelcome -ErrorAction Stop
                            Write-Host "Connected Graph session for user validation (app-only)" -ForegroundColor Gray
                        } catch {
                            Write-Host "ERROR: Failed to connect Graph for validation: $($_.Exception.Message)" -ForegroundColor Red
                            Write-CommandResponse "VALIDATE_USERS_FAILED:Failed to connect Graph - $($_.Exception.Message)"
                            continue
                        }
                    } else {
                        Write-Host "ERROR: No Graph session and no app token. Re-authenticate with Graph." -ForegroundColor Red
                        Write-CommandResponse "VALIDATE_USERS_FAILED:No Graph session. Please complete Graph authentication first."
                        continue
                    }
                }
                
                Write-Host "==========================================" -ForegroundColor Yellow
                Write-Host "VALIDATE USERS COMMAND RECEIVED" -ForegroundColor Yellow
                Write-Host "==========================================" -ForegroundColor Yellow
                Write-Status "User validation command received"
                Write-CommandResponse "VALIDATE_USERS_STARTED"
                
                try {
                    # Parse search terms from command (format: VALIDATE_USERS|SEARCH_TERMS:term1,term2)
                    $searchTerms = @()
                    if ($command -match '\|SEARCH_TERMS:(.+)$') {
                        $searchTermsJson = $Matches[1]
                        try {
                            $searchTermsArray = $searchTermsJson | ConvertFrom-Json -ErrorAction Stop
                            if ($searchTermsArray -is [array]) {
                                $searchTerms = @($searchTermsArray | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
                            } elseif ($searchTermsArray -is [string] -and -not [string]::IsNullOrWhiteSpace($searchTermsArray)) {
                                $searchTerms = @($searchTermsArray)
                            } elseif ($searchTermsArray -ne $null) {
                                $searchTerms = @($searchTermsArray | Where-Object { $_ -ne $null -and -not [string]::IsNullOrWhiteSpace($_) })
                            } else {
                                $searchTerms = @()
                            }
                        } catch {
                            # If JSON parsing fails, try splitting as comma-separated string
                            $searchTerms = @($searchTermsJson -split ',' | ForEach-Object { if ($_ -ne $null) { $_.Trim() } } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
                        }
                    } else {
                        Write-Warning "No search terms found in VALIDATE_USERS command"
                        Write-CommandResponse "VALIDATE_USERS_FAILED:No search terms provided"
                        continue
                    }
                    
                    Write-Host "Search terms received: $($searchTerms -join ', ')" -ForegroundColor Cyan
                    Write-Status "Validating users for search terms: $($searchTerms -join ', ')"
                    
                    # Perform user search - minimal API calls to avoid extra auth prompts
                    $allFoundUsers = [System.Collections.ArrayList]::new()
                    foreach ($searchTerm in $searchTerms) {
                        Write-Host "  Searching for users matching: '$searchTerm'" -ForegroundColor Gray
                        $users = @()
                        try {
                            # For full UPNs/emails (contains @), try direct lookup first - works with both app-only and delegated auth
                            if ($searchTerm.Trim() -match '@') {
                                try {
                                    $u = Get-MgUser -UserId $searchTerm.Trim() -Property Id, UserPrincipalName, DisplayName -ErrorAction Stop
                                    if ($u) { $users = @($u); Write-Host "    Found via direct UPN lookup" -ForegroundColor Gray }
                                } catch {
                                    Write-Host "    Direct UPN lookup failed: $($_.Exception.Message)" -ForegroundColor DarkGray
                                }
                            }
                            # For domain names only (e.g. contoso.com) - not partial UPN like john.smith
                            # Domain heuristic: has dot, no @, and part after last dot is 2-4 chars (TLD)
                            if ((-not $users -or $users.Count -eq 0) -and $searchTerm.Trim() -match '\.' -and $searchTerm.Trim() -notmatch '@') {
                                $lastPart = $searchTerm.Trim().Split('.')[-1]
                                $looksLikeDomain = ($lastPart.Length -ge 2 -and $lastPart.Length -le 4)
                                if ($looksLikeDomain) {
                                    $domainPart = '@' + $searchTerm.Trim().Replace("'","''")
                                    $users = @(Get-MgUser -Filter "endsWith(userPrincipalName,'$domainPart')" -Top 999 -Property Id, UserPrincipalName, DisplayName -ConsistencyLevel eventual -CountVariable userCount -ErrorAction SilentlyContinue)
                                    if ($users.Count -gt 0) { Write-Host "    Found $($users.Count) user(s) in domain" -ForegroundColor Gray }
                                }
                            }
                            if (-not $users -or $users.Count -eq 0) {
                                # startsWith requires ConsistencyLevel eventual for advanced queries
                                $escaped = $searchTerm.Replace("'","''")
                                $users = @(Get-MgUser -Filter "startsWith(DisplayName,'$escaped') or startsWith(UserPrincipalName,'$escaped')" -Top 999 -Property Id, UserPrincipalName, DisplayName -ConsistencyLevel eventual -CountVariable userCount -ErrorAction SilentlyContinue)
                            }
                            if (-not $users -or $users.Count -eq 0) {
                                $escapedLower = $searchTerm.ToLower().Replace("'","''")
                                $users = @(Get-MgUser -Filter "startsWith(DisplayName,'$escapedLower') or startsWith(UserPrincipalName,'$escapedLower')" -Top 999 -Property Id, UserPrincipalName, DisplayName -ConsistencyLevel eventual -CountVariable userCount -ErrorAction SilentlyContinue)
                            }
                            if (-not $users -or $users.Count -eq 0) {
                                # Exact match fallback (eq works without eventual)
                                $escaped = $searchTerm.Replace("'","''")
                                $users = @(Get-MgUser -Filter "DisplayName eq '$escaped' or UserPrincipalName eq '$escaped'" -Top 10 -Property Id, UserPrincipalName, DisplayName -ErrorAction SilentlyContinue)
                            }
                            # Client-side fallback when API filters fail (partial UPN, partial name, app-only without advanced query)
                            # Use contains for both DisplayName and UPN - handles john.smith, Smith, user@domain.com
                            if ((-not $users -or $users.Count -eq 0) -and $searchTerm.Trim().Length -gt 0) {
                                try {
                                    $term = $searchTerm.Trim()
                                    $all = @(Get-MgUser -All -Top 1000 -Property Id, UserPrincipalName, DisplayName -ErrorAction Stop)
                                    $users = @($all | Where-Object { ($_.DisplayName -and $_.DisplayName -like "*$term*") -or ($_.UserPrincipalName -and $_.UserPrincipalName -like "*$term*") })
                                    if ($users.Count -gt 0) { Write-Host "    Found $($users.Count) user(s) via client-side search" -ForegroundColor Gray }
                                } catch {
                                    Write-Host "    Client-side fallback failed: $($_.Exception.Message)" -ForegroundColor DarkGray
                                }
                            }
                            Write-Host "    Found $($users.Count) user(s)" -ForegroundColor Gray
                        } catch {
                            Write-Warning "Search failed for '$searchTerm': $($_.Exception.Message)"
                        }
                        if ($users.Count -gt 0) {
                            foreach ($user in $users) {
                                [void]$allFoundUsers.Add($user)
                            }
                        }
                    }
                    
                    # Get unique UserPrincipalNames
                    $validatedUsers = ($allFoundUsers | Sort-Object UserPrincipalName -Unique | ForEach-Object { $_.UserPrincipalName })
                    
                    if ($validatedUsers.Count -gt 0) {
                        Write-Host "Validation successful: Found $($validatedUsers.Count) user(s)" -ForegroundColor Green
                        Write-Status "Validation successful: Found $($validatedUsers.Count) user(s)"
                        $responseJson = @{
                            Success = $true
                            UserCount = $validatedUsers.Count
                            Users = $validatedUsers
                        } | ConvertTo-Json -Compress
                        Write-CommandResponse "VALIDATE_USERS_SUCCESS:$responseJson"
                    } else {
                        Write-Host "Validation completed: No users found matching search terms" -ForegroundColor Yellow
                        Write-Status "Validation completed: No users found matching search terms"
                        $responseJson = @{
                            Success = $false
                            UserCount = 0
                            Users = @()
                            Message = "No users found matching the search terms"
                        } | ConvertTo-Json -Compress
                        Write-CommandResponse "VALIDATE_USERS_SUCCESS:$responseJson"
                    }
                } catch {
                    Write-Host "ERROR: User validation failed - $($_.Exception.Message)" -ForegroundColor Red
                    Write-Status "ERROR: User validation failed - $($_.Exception.Message)"
                    Write-CommandResponse "VALIDATE_USERS_FAILED:$($_.Exception.Message)"
                }
                
                Write-Host ""
                Write-Host "Waiting for next command from GUI..." -ForegroundColor Green
                Write-Host ""
                
            } elseif ($command -match "^GENERATE_REPORTS") {
                $required = @{ NeedsGraph = $true; NeedsExchange = $true }
                if (Get-Command Get-RequiredAuthFromReportSelections -ErrorAction SilentlyContinue) {
                    $required = Get-RequiredAuthFromReportSelections -ReportSelections $reportSelections
                }
                $missing = @()
                if ($required.NeedsGraph -and -not $graphAuthenticated) { $missing += "Graph" }
                if ($required.NeedsExchange -and -not $exchangeAuthenticated) { $missing += "Exchange Online" }
                if ($missing.Count -gt 0) {
                    Write-Host "ERROR: Selected reports require: $($missing -join ', '). Please complete authentication first." -ForegroundColor Red
                    Write-CommandResponse "GENERATE_REPORTS_FAILED:Authentication required: $($missing -join ', ')"
                    continue
                }
                
                # Parse SelectedUsers, TicketData, and DATE_RANGE from command if provided
                $selectedUsersForReport = @()
                $ticketNumbers = @()
                $ticketContent = ''
                $reportStartDate = [DateTime]::MinValue
                $reportEndDate = [DateTime]::MinValue
                
                # Parse DATE_RANGE from command (format: |DATE_RANGE:{"StartDate":"...","EndDate":"..."})
                if ($command -match '\|DATE_RANGE:(\{[^|]+\})') {
                    try {
                        $dateRangeJson = $Matches[1]
                        $dateRange = $dateRangeJson | ConvertFrom-Json -ErrorAction Stop
                        if ($dateRange.StartDate -and $dateRange.EndDate) {
                            $reportStartDate = [DateTime]::Parse($dateRange.StartDate)
                            $reportEndDate = [DateTime]::Parse($dateRange.EndDate)
                            if ($reportEndDate -ge $reportStartDate) {
                                Write-Host "Date range from command: $($reportStartDate.ToString('yyyy-MM-dd')) to $($reportEndDate.ToString('yyyy-MM-dd'))" -ForegroundColor Cyan
                                Write-Status "Date range: $($reportStartDate.ToString('yyyy-MM-dd')) to $($reportEndDate.ToString('yyyy-MM-dd'))"
                            } else {
                                Write-Warning "Invalid date range (End < Start), ignoring"
                                $reportStartDate = [DateTime]::MinValue
                                $reportEndDate = [DateTime]::MinValue
                            }
                        }
                    } catch {
                        Write-Warning "Could not parse DATE_RANGE from command: $($_.Exception.Message)"
                    }
                }
                
                # Parse ticket data from command (format: |TICKET_DATA:{"TicketNumbers":["12345"],"TicketContent":"..."}|DATE_RANGE:...)
                # Extract only the ticket JSON - stop at |DATE_RANGE: to avoid passing trailing content to ConvertFrom-Json
                # (TicketContent can contain | and other chars; |DATE_RANGE: delimiter would cause "Additional text" JSON error)
                Write-Host "Parsing ticket data from command. Command length: $($command.Length)" -ForegroundColor Gray
                Write-Host "Command preview (first 500 chars): $($command.Substring(0, [Math]::Min(500, $command.Length)))" -ForegroundColor Gray
                if ($command -match '\|TICKET_DATA:(.+?)(?:\|DATE_RANGE:|$)') {
                    Write-Host "TICKET_DATA regex matched!" -ForegroundColor Green
                    try {
                        $ticketDataJson = $Matches[1]
                        Write-Host "Ticket data JSON extracted (length: $($ticketDataJson.Length))" -ForegroundColor Gray
                        Write-Host "Ticket data JSON preview (first 300 chars): $($ticketDataJson.Substring(0, [Math]::Min(300, $ticketDataJson.Length)))" -ForegroundColor Gray
                        $ticketData = $ticketDataJson | ConvertFrom-Json -ErrorAction Stop
                        Write-Host "Ticket data JSON parsed successfully" -ForegroundColor Green
                        if ($ticketData.TicketNumbers) {
                            Write-Host "TicketNumbers property found: $($ticketData.TicketNumbers)" -ForegroundColor Gray
                            # Ensure TicketNumbers is always an array
                            if ($ticketData.TicketNumbers -is [string]) {
                                $ticketNumbers = @($ticketData.TicketNumbers)
                                Write-Host "TicketNumbers was string, converted to array: $ticketNumbers" -ForegroundColor Gray
                            } elseif ($ticketData.TicketNumbers -is [array]) {
                                $ticketNumbers = $ticketData.TicketNumbers
                                Write-Host "TicketNumbers was array: $ticketNumbers" -ForegroundColor Gray
                            } else {
                                $ticketNumbers = @($ticketData.TicketNumbers)
                                Write-Host "TicketNumbers was other type, converted to array: $ticketNumbers" -ForegroundColor Gray
                            }
                        } else {
                            Write-Host "TicketNumbers property not found in parsed data" -ForegroundColor Yellow
                        }
                        if ($ticketData.TicketContent) {
                            $ticketContent = $ticketData.TicketContent
                            Write-Host "TicketContent property found (length: $($ticketContent.Length))" -ForegroundColor Gray
                        } else {
                            Write-Host "TicketContent property not found in parsed data" -ForegroundColor Yellow
                        }
                        Write-Host "Ticket data parsed: TicketNumbers=$($ticketNumbers.Count) ($($ticketNumbers -join ', ')), TicketContent length=$($ticketContent.Length)" -ForegroundColor Cyan
                        Write-Host "Ticket data found: $($ticketNumbers.Count) ticket number(s): $($ticketNumbers -join ', ')" -ForegroundColor Cyan
                        Write-Status "Ticket data found: $($ticketNumbers.Count) ticket number(s): $($ticketNumbers -join ', ')"
                    } catch {
                        Write-Warning "Could not parse ticket data from command: $($_.Exception.Message)"
                        Write-Host "Ticket data JSON that failed to parse: $ticketDataJson" -ForegroundColor Yellow
                        Write-Host "Full command was: $command" -ForegroundColor Yellow
                        Write-Host "Exception details: $($_.Exception | Out-String)" -ForegroundColor Red
                    }
                } else {
                    Write-Host "No TICKET_DATA found in command. Command preview: $($command.Substring(0, [Math]::Min(500, $command.Length)))" -ForegroundColor Yellow
                    Write-Host "Checking if command contains 'TICKET_DATA': $($command.Contains('TICKET_DATA'))" -ForegroundColor Yellow
                }
                
                # Check if this is a search terms command (GENERATE_REPORTS_SEARCH:["term1","term2"])
                # Extract search terms before TICKET_DATA if present
                if ($command -match "^GENERATE_REPORTS_SEARCH:(.+?)(?:\|TICKET_DATA:|$)") {
                    try {
                        $searchTermsJson = $Matches[1]
                        $searchTermsParsed = $searchTermsJson | ConvertFrom-Json -ErrorAction Stop
                        # Ensure searchTerms is always an array (ConvertFrom-Json might return a string for single values)
                        if ($searchTermsParsed -is [string]) {
                            $searchTerms = @($searchTermsParsed)
                        } elseif ($searchTermsParsed -is [array]) {
                            $searchTerms = $searchTermsParsed
                        } else {
                            $searchTerms = @($searchTermsParsed)
                        }
                        Write-Host "User filtering enabled with search terms. Validating users..." -ForegroundColor Cyan
                        Write-Status "User filtering enabled with search terms. Validating users..."
                        
                        # Validate search terms using Graph API
                        $allFoundUsers = [System.Collections.ArrayList]::new()
                        foreach ($searchTerm in $searchTerms) {
                            Write-Host "  Searching for users matching: '$searchTerm'" -ForegroundColor Gray
                            $users = @()
                            try {
                                # Try server-side filtering first (startsWith) - try multiple case variations
                                $users1 = Get-MgUser -Filter "startsWith(DisplayName,'$searchTerm') or startsWith(UserPrincipalName,'$searchTerm')" -All -Property Id, UserPrincipalName, DisplayName -ConsistencyLevel eventual -CountVariable userCount -ErrorAction SilentlyContinue
                                $searchTermLower = $searchTerm.ToLower()
                                $searchTermUpper = $searchTerm.ToUpper()
                                $searchTermTitle = (Get-Culture).TextInfo.ToTitleCase($searchTermLower)
                                $users2 = Get-MgUser -Filter "startsWith(DisplayName,'$searchTermLower') or startsWith(UserPrincipalName,'$searchTermLower')" -All -Property Id, UserPrincipalName, DisplayName -ConsistencyLevel eventual -CountVariable userCount -ErrorAction SilentlyContinue
                                $users3 = Get-MgUser -Filter "startsWith(DisplayName,'$searchTermUpper') or startsWith(UserPrincipalName,'$searchTermUpper')" -All -Property Id, UserPrincipalName, DisplayName -ConsistencyLevel eventual -CountVariable userCount -ErrorAction SilentlyContinue
                                $users4 = Get-MgUser -Filter "startsWith(DisplayName,'$searchTermTitle') or startsWith(UserPrincipalName,'$searchTermTitle')" -All -Property Id, UserPrincipalName, DisplayName -ConsistencyLevel eventual -CountVariable userCount -ErrorAction SilentlyContinue
                                $users = @($users1) + @($users2) + @($users3) + @($users4) | Sort-Object UserPrincipalName -Unique
                                Write-Host "    Found $($users.Count) users with startsWith filter (tried multiple case variations)" -ForegroundColor Gray
                            } catch {
                                Write-Host "    startsWith filter failed: $($_.Exception.Message), trying alternatives..." -ForegroundColor Yellow
                            }
                            
                            if ($users.Count -eq 0) {
                                # Try alternative search methods
                                try {
                                    # Try exact match (case-sensitive first, then variations)
                                    $usersAlt1 = Get-MgUser -Filter "DisplayName eq '$searchTerm'" -All -Property Id, UserPrincipalName, DisplayName -ErrorAction SilentlyContinue
                                    $usersAlt1 += Get-MgUser -Filter "DisplayName eq '$searchTermLower'" -All -Property Id, UserPrincipalName, DisplayName -ErrorAction SilentlyContinue
                                    $usersAlt1 = $usersAlt1 | Sort-Object UserPrincipalName -Unique
                                    
                                    $usersAlt2 = Get-MgUser -Filter "UserPrincipalName eq '$searchTerm'" -All -Property Id, UserPrincipalName, DisplayName -ErrorAction SilentlyContinue
                                    $usersAlt2 += Get-MgUser -Filter "UserPrincipalName eq '$searchTermLower'" -All -Property Id, UserPrincipalName, DisplayName -ErrorAction SilentlyContinue
                                    $usersAlt2 = $usersAlt2 | Sort-Object UserPrincipalName -Unique
                                    
                                    # Try case-insensitive search by getting all users and filtering client-side
                                    Write-Host "    Fetching all users for client-side filtering..." -ForegroundColor Gray
                                    try {
                                        $allUsers = Get-MgUser -All -Property Id, UserPrincipalName, DisplayName -ErrorAction Stop
                                        Write-Host "    Retrieved $($allUsers.Count) total users from tenant" -ForegroundColor Gray
                                        
                                        # Use case-insensitive matching with -ilike
                                        $searchTermPattern = "*$searchTerm*"
                                        $usersAlt3 = $allUsers | Where-Object { 
                                            ($_.DisplayName -and $_.DisplayName -ilike $searchTermPattern) -or 
                                            ($_.UserPrincipalName -and $_.UserPrincipalName -ilike $searchTermPattern)
                                        }
                                        Write-Host "    Client-side filtering: Found $($usersAlt3.Count) users matching '$searchTerm'" -ForegroundColor Gray
                                    } catch {
                                        Write-Warning "Failed to retrieve all users for client-side filtering: $($_.Exception.Message)"
                                        $usersAlt3 = @()
                                    }
                                    
                                    # Combine all results
                                    $users = @($usersAlt1) + @($usersAlt2) + @($usersAlt3) | Sort-Object UserPrincipalName -Unique
                                    Write-Host "    Combined alternative searches: Found $($users.Count) users" -ForegroundColor Gray
                                } catch {
                                    Write-Warning "Could not search for users matching '$searchTerm': $($_.Exception.Message)"
                                }
                            }
                            if ($users.Count -gt 0) {
                                $allFoundUsers += $users
                            }
                        }
                        
                        # Get unique UserPrincipalNames
                        $selectedUsersForReport = ($allFoundUsers | Sort-Object UserPrincipalName -Unique | ForEach-Object { $_.UserPrincipalName })
                        Write-Host "User filtering enabled: Found $($selectedUsersForReport.Count) user(s) from search terms" -ForegroundColor Cyan
                        Write-Status "User filtering enabled: Found $($selectedUsersForReport.Count) user(s) from search terms"
                        
                        # Warn if search terms were provided but no users found
                        if ($selectedUsersForReport.Count -eq 0) {
                            Write-Warning "No users found matching the search terms. Report will be generated without user filtering."
                            Write-Status "WARNING: No users found matching search terms - generating report without filtering"
                        }
                    } catch {
                        Write-Warning "Could not parse or validate search terms from command: $($_.Exception.Message)"
                        Write-Status "ERROR: Failed to validate search terms - $($_.Exception.Message)"
                        # Set to empty array so report continues without filtering
                        $selectedUsersForReport = @()
                    }
                }
                # Check if this is a direct users command (GENERATE_REPORTS|SelectedUsers:["user1","user2"])
                elseif ($command -match '\|SelectedUsers:(.+?)(?:\||$)') {
                    try {
                        $usersJson = $Matches[1]
                        $parsed = $usersJson | ConvertFrom-Json -ErrorAction Stop
                        # ConvertFrom-Json returns scalar for single-element array in PS 5.1 - ensure array
                        $selectedUsersForReport = @()
                        foreach ($p in @($parsed)) {
                            $upn = if ($p -is [string]) { $p } elseif ($p.UserPrincipalName) { $p.UserPrincipalName } else { $p.ToString() }
                            if (-not [string]::IsNullOrWhiteSpace($upn)) { $selectedUsersForReport += $upn }
                        }
                        Write-Host "User filtering enabled: $($selectedUsersForReport.Count) user(s) selected: $($selectedUsersForReport -join ', ')" -ForegroundColor Cyan
                        Write-Status "User filtering enabled: $($selectedUsersForReport.Count) user(s)"
                    } catch {
                        Write-Warning "Could not parse SelectedUsers from command: $($_.Exception.Message)"
                    }
                }
                
                Write-Host "==========================================" -ForegroundColor Yellow
                Write-Host "GENERATE REPORTS COMMAND RECEIVED" -ForegroundColor Yellow
                Write-Host "==========================================" -ForegroundColor Yellow
                Write-Status "Report generation command received"
                Write-CommandResponse "GENERATE_REPORTS_STARTED"
                
                # Generate Reports
                Write-Host ""
                Write-Host "Generating security investigation report..." -ForegroundColor Cyan
                
                # Generate security investigation report (will use default folder structure matching non-bulk)
                # OutputFolder will be automatically determined by ExportUtils using:
                # Documents\ExchangeOnlineAnalyzer\SecurityInvestigation\{TenantName}\{Timestamp}
                Write-Status "Generating security investigation report..."
                Write-Host "Starting report generation..." -ForegroundColor Yellow
                # Filter ticket content to remove configuration sections
                if ($ticketContent -and -not [string]::IsNullOrWhiteSpace($ticketContent)) {
                    try {
                        Import-Module "$ScriptRoot\Modules\Settings.psm1" -Force -ErrorAction SilentlyContinue
                        if (Get-Command Filter-TicketContent -ErrorAction SilentlyContinue) {
                            $originalLength = $ticketContent.Length
                            $ticketContent = Filter-TicketContent -TicketContent $ticketContent
                            Write-Host "Ticket content filtered: $originalLength -> $($ticketContent.Length) characters" -ForegroundColor Gray
                        } else {
                            Write-Warning "Filter-TicketContent function not found, using raw ticket content"
                        }
                    } catch {
                        Write-Warning "Failed to filter ticket content: $($_.Exception.Message). Using raw content."
                    }
                }
                
                Write-Host "Ticket data being passed: TicketNumbers=$($ticketNumbers.Count) ($($ticketNumbers -join ', ')), TicketContent length=$($ticketContent.Length)" -ForegroundColor Cyan
                if (Get-Command Set-LogContext -ErrorAction SilentlyContinue) { Set-LogContext -CompanyName $CompanyName -TicketNumbers $ticketNumbers }
                # POC: Get Graph token for parallel collection (avoids extra auth prompts in runspaces)
                $graphToken = $null
                if ($script:graphTokenFromWCM) {
                    $graphToken = $script:graphTokenFromWCM
                    Write-Status "Using app-only token from WCM - parallel collection"
                } elseif (Get-Command Get-GraphAccessToken -ErrorAction SilentlyContinue) {
                    try {
                        $diag = { param($m) Write-Status $m; if (Get-Command Write-Log -ErrorAction SilentlyContinue) { Write-Log -Message "Get-GraphAccessToken: $m" -Level Debug } }
                        $graphToken = Get-GraphAccessToken -DiagnosticCallback $diag
                    } catch {
                        Write-Status "Graph token acquisition failed: $($_.Exception.Message)"
                        if (Get-Command Write-Log -ErrorAction SilentlyContinue) { Write-Log -Message "Get-GraphAccessToken failed: $($_.Exception.Message)" -Level Warning }
                    }
                }
                if ($graphToken) { Write-Status "Graph token acquired - using parallel collection" } else { Write-Status "No Graph token - using sequential collection" }
                try {
                    $messageTraceDays = if ($reportSelections.MessageTraceDaysBack) { $reportSelections.MessageTraceDaysBack } else { $DaysBack }
                    $reportParams = @{
                        InvestigatorName = $InvestigatorName
                        CompanyName = $CompanyName
                        DaysBack = $DaysBack
                        StatusLabel = $null
                        MainForm = $null
                        NoParallel = (-not $graphToken)
                        GraphAccessToken = $graphToken
                        ProgressCallback = { param($m) Write-Status $m }
                        SessionId = "Client$ClientNumber"
                        StatusFile = $StatusFile
                        IncludeMessageTrace = $reportSelections.IncludeMessageTrace
                        IncludeInboxRules = $reportSelections.IncludeInboxRules
                        IncludeTransportRules = $reportSelections.IncludeTransportRules
                        IncludeMailFlowConnectors = $reportSelections.IncludeMailFlowConnectors
                        IncludeMailboxForwarding = $reportSelections.IncludeMailboxForwarding
                        IncludeAuditLogs = $reportSelections.IncludeAuditLogs
                        IncludeConditionalAccessPolicies = $reportSelections.IncludeConditionalAccessPolicies
                        IncludeAppRegistrations = $reportSelections.IncludeAppRegistrations
                        IncludeSignInLogs = $reportSelections.IncludeSignInLogs
                        IncludeIntuneDevices = $reportSelections.IncludeIntuneDevices
                        IncludeMfaCoverage = $reportSelections.IncludeMfaCoverage
                        IncludeSharePointActivity = $reportSelections.IncludeSharePointActivity
                        IncludeOneDriveActivity = $reportSelections.IncludeOneDriveActivity
                        IncludeTeamsActivity = $reportSelections.IncludeTeamsActivity
                        IncludeSharePointSharing = $reportSelections.IncludeSharePointSharing
                        IncludeSecurityAlerts = $reportSelections.IncludeSecurityAlerts
                        IncludeSecurityIncidents = $reportSelections.IncludeSecurityIncidents
                        IncludeUnifiedAuditLogs = $reportSelections.IncludeUnifiedAuditLogs
                        SignInLogsDaysBack = $reportSelections.SignInLogsDaysBack
                        MessageTraceDaysBack = $messageTraceDays
                        SelectedUsers = $selectedUsersForReport
                        TicketNumbers = $ticketNumbers
                        TicketContent = $ticketContent
                    }
                    if ($reportStartDate -and $reportEndDate -and $reportStartDate -ne [DateTime]::MinValue -and $reportEndDate -ne [DateTime]::MinValue -and $reportEndDate -ge $reportStartDate) {
                        $reportParams.StartDate = $reportStartDate
                        $reportParams.EndDate = $reportEndDate
                    }
                    $report = New-SecurityInvestigationReport @reportParams
                    Write-Status "Report generation function completed"
                    Write-Host "Report generation function completed successfully" -ForegroundColor Green
                } catch {
                    # SECURITY: Use safe error handling - don't expose full exception details
                    if (Get-Command Get-SafeErrorMessage -ErrorAction SilentlyContinue) {
                        $safeError = Get-SafeErrorMessage -Error $_ -UserMessage "Failed to generate report"
                        Write-Status "ERROR: Failed to generate report - $safeError"
                        Write-Host "ERROR: Failed to generate report - $safeError" -ForegroundColor Red
                        Write-CommandResponse "GENERATE_REPORTS_FAILED:$safeError"
                    } else {
                        $errMsg = if ($_.Exception.Message) { $_.Exception.Message } else { "Report generation failed" }
                        Write-Status "ERROR: Failed to generate report"
                        Write-Host "ERROR: Failed to generate report" -ForegroundColor Red
                        Write-CommandResponse "GENERATE_REPORTS_FAILED:$errMsg"
                    }
                    continue
                }
                
                if ($report -and $report.OutputFolder) {
                    Write-Status "Report generation successful!"
                    Write-Host "`nReport generation successful!" -ForegroundColor Green
                    Write-Host "Reports saved to: $($report.OutputFolder)" -ForegroundColor Green
                    "SUCCESS: $($report.OutputFolder)" | Out-File -FilePath $ResultFile -Encoding UTF8
                    Write-CommandResponse "GENERATE_REPORTS_SUCCESS:$($report.OutputFolder)"
                } else {
                    Write-Status "Warning: Report generation returned no data"
                    Write-Host "Warning: Report generation returned no data" -ForegroundColor Yellow
                    $defaultOutput = Join-Path ([Environment]::GetFolderPath('MyDocuments')) "ExchangeOnlineAnalyzer\SecurityInvestigation"
                    "NO_DATA: $defaultOutput" | Out-File -FilePath $ResultFile -Encoding UTF8
                    Write-CommandResponse "GENERATE_REPORTS_NO_DATA:$defaultOutput"
                }
                
                # Update status FIRST so completion is recorded even if disconnect hangs
                Write-Status "Processing complete!"
                
                # Disconnect sessions (attempt but don't block if it hangs)
                Write-Host "Disconnecting sessions..." -ForegroundColor Cyan
                try {
                    Disconnect-MgGraph -ErrorAction SilentlyContinue
                } catch {}
                
                # Attempt Exchange disconnect with timeout (non-blocking)
                try {
                    if (Get-Command Disconnect-ExchangeOnline -ErrorAction SilentlyContinue) {
                        # Use runspace with module import and timeout
                        $runspace = [runspacefactory]::CreateRunspace()
                        $runspace.Open()
                        $ps = [PowerShell]::Create()
                        $ps.Runspace = $runspace
                        # Import module and disconnect
                        $script = "Import-Module ExchangeOnlineManagement -ErrorAction SilentlyContinue; Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue"
                        $null = $ps.AddScript($script)
                        $handle = $ps.BeginInvoke()
                        $waited = $handle.AsyncWaitHandle.WaitOne(5000)  # 5 second timeout
                        if ($waited) {
                            try { $ps.EndInvoke($handle) | Out-Null } catch {}
                        } else {
                            Write-Host "Exchange disconnect timed out (non-critical, continuing...)" -ForegroundColor Yellow
                            $ps.Stop()
                        }
                        $ps.Dispose()
                        $runspace.Close()
                        $runspace.Dispose()
                    }
                } catch {
                    Write-Host "Disconnect completed with warnings (non-critical)" -ForegroundColor Yellow
                }
                Write-Host ""
                Write-Host "==========================================" -ForegroundColor Green
                Write-Host "Client $ClientNumber processing complete!" -ForegroundColor Green
                Write-Host "==========================================" -ForegroundColor Green
                Write-Host "This window will remain open. You may close it manually." -ForegroundColor Yellow
                Write-Host ""
                
                # Keep window open but stop polling
                break
            } elseif ($command -eq "CANCEL_AUTH") {
                Write-Host "==========================================" -ForegroundColor Yellow
                Write-Host "CANCEL AUTHENTICATION COMMAND RECEIVED" -ForegroundColor Yellow
                Write-Host "==========================================" -ForegroundColor Yellow
                Write-Status "Cancelling authentication and resetting state..."
                
                # Reset authentication state
                $graphAuthenticated = $false
                $exchangeAuthenticated = $false
                $tenantDisplayName = $null
                
                # Disconnect any active sessions
                try {
                    Disconnect-MgGraph -ErrorAction SilentlyContinue
                    Start-Sleep -Milliseconds 500
                } catch {}
                try {
                    if (Get-Command Disconnect-ExchangeOnline -ErrorAction SilentlyContinue) {
                        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
                    }
                    Get-PSSession | Where-Object { $_.ConfigurationName -eq "Microsoft.Exchange" } | Remove-PSSession -ErrorAction SilentlyContinue
                } catch {}
                
                # Clear authentication context and token cache (same as GRAPH_AUTH)
                try {
                    $graphSession = [Microsoft.Graph.PowerShell.Authentication.GraphSession]::Instance
                    if ($graphSession -and $graphSession.AuthContext) {
                        $graphSession.AuthContext.ClearTokenCache()
                        Write-Host "Cleared Graph token cache" -ForegroundColor Cyan
                    }
                } catch {}
                
                # Clear MSAL and Identity cache directories
                try {
                    if ($env:MSAL_CACHE_DIR -and (Test-Path $env:MSAL_CACHE_DIR)) {
                        Get-ChildItem -Path $env:MSAL_CACHE_DIR -File -Recurse -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
                        Write-Host "Cleared MSAL cache directory" -ForegroundColor Cyan
                    }
                    if ($env:IDENTITY_SERVICE_CACHE_DIR -and (Test-Path $env:IDENTITY_SERVICE_CACHE_DIR)) {
                        Get-ChildItem -Path $env:IDENTITY_SERVICE_CACHE_DIR -File -Recurse -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
                        Write-Host "Cleared Identity cache directory" -ForegroundColor Cyan
                    }
                } catch {}
                
                # Clear default MSAL cache, Graph module cache, WAM cache
                try {
                    $paths = @(
                        (Join-Path $env:LOCALAPPDATA ".IdentityService"),
                        (Join-Path $env:LOCALAPPDATA "Microsoft\Graph"),
                        (Join-Path $env:LOCALAPPDATA "Packages\Microsoft.AAD.BrokerPlugin_cw5n1h2txyewy\AC\TokenBroker\Cache"),
                        (Join-Path $env:LOCALAPPDATA "Packages\Microsoft.AAD.BrokerPlugin_cw5n1h2txyewy\LocalState")
                    )
                    foreach ($p in $paths) {
                        if (Test-Path $p) {
                            Get-ChildItem -Path $p -Recurse -File -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
                        }
                    }
                    Write-Host "Cleared IdentityService, Graph, and WAM caches" -ForegroundColor Cyan
                } catch {}
    
                Write-Status "Authentication cancelled and reset"
                Write-Host "Authentication cancelled and reset. All token caches cleared. Ready for new authentication attempt." -ForegroundColor Green
                Write-CommandResponse "CANCEL_AUTH_SUCCESS"
            } elseif ($command -eq "GRAPH_DISCONNECT") {
                Write-Host "GRAPH_DISCONNECT: signing out Microsoft Graph only (Exchange stays connected)..." -ForegroundColor Yellow
                try { Disconnect-MgGraph -ErrorAction SilentlyContinue } catch {}
                try {
                    $gs = [Microsoft.Graph.PowerShell.Authentication.GraphSession]::Instance
                    if ($gs -and $gs.AuthContext) { $gs.AuthContext.ClearTokenCache(); Write-Host "Cleared Graph token cache" -ForegroundColor Cyan }
                } catch {}
                try {
                    $msalCache = [Microsoft.Identity.Client.TokenCacheHelper]::GetCacheFilePath()
                    if ($msalCache -and (Test-Path $msalCache)) { Remove-Item $msalCache -Force -ErrorAction SilentlyContinue }
                } catch {}
                try {
                    if ($env:MSAL_CACHE_DIR -and (Test-Path $env:MSAL_CACHE_DIR)) {
                        Get-ChildItem -Path $env:MSAL_CACHE_DIR -Recurse -File -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
                    }
                } catch {}
                try {
                    if ($env:IDENTITY_SERVICE_CACHE_DIR -and (Test-Path $env:IDENTITY_SERVICE_CACHE_DIR)) {
                        Get-ChildItem -Path $env:IDENTITY_SERVICE_CACHE_DIR -Recurse -File -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
                    }
                } catch {}
                Write-Status "Microsoft Graph disconnected"
                Write-Host "Graph session ended in this window. Use Graph Auth to sign in again." -ForegroundColor Green
                Write-CommandResponse "GRAPH_DISCONNECT_SUCCESS"
            } elseif ($command -eq "EXIT") {
                Write-Host "Exit command received. Closing window..." -ForegroundColor Yellow
                break
            }
        }
        
        Start-Sleep -Milliseconds $pollInterval
    }
    
} catch {
    Write-Host "`n==========================================" -ForegroundColor Red
    Write-Host "FATAL ERROR OCCURRED IN TRY BLOCK" -ForegroundColor Red
    Write-Host "==========================================" -ForegroundColor Red
    
    $errorMsg = if ($_.Exception.Message) { $_.Exception.Message } else { "Unknown error" }
    $errorDetails = if ($_.Exception) { $_.Exception | Out-String } else { "No exception details available" }
    $errorLocation = if ($_.InvocationInfo) { $_.InvocationInfo.PositionMessage } else { "Unknown location" }
    
    Write-Host "ERROR: $errorMsg" -ForegroundColor Red
    Write-Host "`nFull error details:" -ForegroundColor Red
    Write-Host $errorDetails -ForegroundColor Red
    Write-Host "`nError location:" -ForegroundColor Red
    Write-Host $errorLocation -ForegroundColor Red
    Write-Host "`n==========================================" -ForegroundColor Red
    
    try {
        Write-Status "ERROR: $errorMsg"
    } catch {
        Write-Host "Could not write to status file: $($_.Exception.Message)" -ForegroundColor Yellow
    }
    
    try {
        if ($ResultFile) {
            "ERROR: $errorMsg`n`nFull details:`n$errorDetails`n`nLocation:`n$errorLocation" | Out-File -FilePath $ResultFile -Encoding UTF8 -ErrorAction SilentlyContinue
        }
    } catch {
        Write-Host "Could not write to result file: $($_.Exception.Message)" -ForegroundColor Yellow
    }
    
    Write-Host "`nWindow will stay open for 60 seconds so you can read the error..." -ForegroundColor Yellow
    Write-Host "Press any key to exit immediately, or wait 60 seconds..." -ForegroundColor Yellow
    
    # Wait for keypress or timeout - longer timeout
    try {
        $keyPressed = $false
        $startTime = Get-Date
        while (((Get-Date) - $startTime).TotalSeconds -lt 60) {
            if ($Host.UI.RawUI.KeyAvailable) {
                $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                $keyPressed = $true
                break
            }
            Start-Sleep -Milliseconds 100
        }
        if (-not $keyPressed) {
            Write-Host "`nTimeout reached, exiting..." -ForegroundColor Gray
        }
    } catch {
        Write-Host "Key input not available, waiting 60 seconds..." -ForegroundColor Gray
        Start-Sleep -Seconds 60
    }
    exit 1
}

# Catch any errors that occur OUTSIDE the try block (shouldn't happen but just in case)
trap {
    Write-Host "`n==========================================" -ForegroundColor Red
    Write-Host "FATAL ERROR OUTSIDE TRY BLOCK" -ForegroundColor Red
    Write-Host "==========================================" -ForegroundColor Red
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "`nFull error:" -ForegroundColor Red
    Write-Host ($_.Exception | Out-String) -ForegroundColor Red
    Write-Host "`nWindow will stay open for 60 seconds..." -ForegroundColor Yellow
    Start-Sleep -Seconds 60
    exit 1
}

