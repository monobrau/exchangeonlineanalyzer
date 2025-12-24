function Test-GraphModules {
    foreach ($moduleInfo in $script:requiredGraphModules) {
        if (-not (Get-Module -ListAvailable -Name $moduleInfo.Name)) {
            Write-Warning "Required Graph module $($moduleInfo.Name) is missing."
            return $false
        }
    }
    Write-Host "All required Microsoft Graph modules are available." -ForegroundColor Green
    return $true
}

function Install-GraphModules {
    param($statusLabel)
    Write-Host "Attempting to install required Microsoft Graph modules..." -ForegroundColor Yellow
    if ($statusLabel) { $statusLabel.Text = "Installing Graph modules..." }
    $allInstalled = $true
    foreach ($moduleInfo in $script:requiredGraphModules) {
        if (-not (Get-Module -ListAvailable -Name $moduleInfo.Name)) {
            Write-Host "Installing module: $($moduleInfo.Name)..."
            try {
                Install-Module -Name $moduleInfo.Name -Scope CurrentUser -Repository PSGallery -Force -AllowClobber -ErrorAction Stop
                Write-Host "Module $($moduleInfo.Name) installed successfully." -ForegroundColor Green
            } catch {
                $ex = $_.Exception
                Write-Error ("Failed to install module $($moduleInfo.Name). Please install it manually: Install-Module $($moduleInfo.Name) -Scope CurrentUser. Error: {0}" -f $ex.Message)
                if ($statusLabel) { $statusLabel.Text = "Error installing $($moduleInfo.Name). See console." }
                $allInstalled = $false
            }
        }
    }
    if ($allInstalled) {
        Write-Host "All required Graph modules checked/installed. Please restart the script if prompted or if new modules were installed." -ForegroundColor Green
        if ($statusLabel) { $statusLabel.Text = "Graph modules installed/checked. Restart script if needed." }
        return $true
    } else {
        return $false
    }
}

function Fix-GraphModuleConflicts {
    param([System.Windows.Forms.Label]$statusLabel)
    Write-Host "Attempting to fix Microsoft Graph module version conflicts..." -ForegroundColor Yellow
    if ($statusLabel) { $statusLabel.Text = "Fixing Graph module conflicts..." }

    try {
        # Disconnect any existing connections
        Disconnect-MgGraph -ErrorAction SilentlyContinue

        # Uninstall all Microsoft Graph modules (no wildcard with -Name)
        Write-Host "Unloading Microsoft Graph modules from session..." -ForegroundColor Cyan
        Get-Module -Name "Microsoft.Graph*" -All | Remove-Module -Force -ErrorAction SilentlyContinue

        Write-Host "Uninstalling all installed Microsoft Graph modules..." -ForegroundColor Cyan
        $installed = @()
        try {
            $installed = Get-InstalledModule -Name 'Microsoft.Graph*' -AllVersions -ErrorAction SilentlyContinue
        } catch { $installed = @() }
        if (-not $installed -or $installed.Count -eq 0) {
            # Fallback to list-available
            $available = Get-Module -ListAvailable -Name 'Microsoft.Graph*'
            $installed = $available | Sort-Object -Property Name, Version -Unique | ForEach-Object { @{ Name=$_.Name; Version=$_.Version } }
        }
        foreach ($m in $installed) {
            try {
                $name = if ($m -is [Microsoft.PowerShell.Commands.PSRepositoryItemInfo]) { $m.Name } elseif ($m.PSObject.Properties['Name']) { $m.Name } else { $m['Name'] }
                $ver  = if ($m -is [Microsoft.PowerShell.Commands.PSRepositoryItemInfo]) { $m.Version } elseif ($m.PSObject.Properties['Version']) { $m.Version } else { $m['Version'] }
                if ($name) {
                    if ($ver) { Uninstall-Module -Name $name -RequiredVersion $ver -Force -ErrorAction SilentlyContinue }
                    else { Uninstall-Module -Name $name -AllVersions -Force -ErrorAction SilentlyContinue }
                }
            } catch {
                Write-Warning "Failed to uninstall module ${name} ${ver}: $($_.Exception.Message)"
            }
        }

        # Clear any cached modules
        Get-Module -Name "Microsoft.Graph*" -ListAvailable | ForEach-Object {
            try {
                Remove-Item -Path $_.ModuleBase -Recurse -Force -ErrorAction SilentlyContinue
            } catch {
                Write-Warning "Could not remove module directory: $($_.ModuleBase)"
            }
        }

        # Clear MSAL cache (if available)
        try {
            # Check if TokenCacheHelper class exists before using it
            if ([Microsoft.Identity.Client.TokenCacheHelper] -as [type]) {
                $msalCache = [Microsoft.Identity.Client.TokenCacheHelper]::GetCacheFilePath()
                if ($msalCache -and (Test-Path $msalCache)) {
                    Remove-Item $msalCache -Force -ErrorAction SilentlyContinue
                    Write-Host "Cleared MSAL token cache" -ForegroundColor Cyan
                }
            }
        } catch {
            # Ignore errors clearing MSAL cache - method may not be available
        }

        # Reinstall using umbrella to ensure consistent versions
        Write-Host "Installing Microsoft.Graph umbrella module for consistent versions..." -ForegroundColor Cyan
        try {
            Install-Module -Name Microsoft.Graph -Scope CurrentUser -Repository PSGallery -Force -AllowClobber -ErrorAction Stop
            Write-Host "✓ Microsoft.Graph installed successfully" -ForegroundColor Green
        } catch {
            Write-Error "Failed to install Microsoft.Graph umbrella: $($_.Exception.Message)"
            return $false
        }

        Write-Host "Microsoft Graph module conflicts fixed! Please restart PowerShell and try connecting again." -ForegroundColor Green
        if ($statusLabel) { $statusLabel.Text = "Graph module conflicts fixed. Restart PowerShell." }

        return $true

    } catch {
        Write-Error "Failed to fix Microsoft Graph module conflicts: $($_.Exception.Message)"
        if ($statusLabel) { $statusLabel.Text = "Error fixing Graph modules. See console." }
        return $false
    }
}

function Connect-GraphService {
    param(
        [Parameter(Mandatory=$false)]
        [System.Windows.Forms.ToolStripStatusLabel]$statusLabel,
        [Parameter(Mandatory=$false)]
        [System.Windows.Forms.Form]$mainForm
    )
    try {
        if ($statusLabel) { $statusLabel.Text = "Connecting to Microsoft Graph..." }
        if ($mainForm) { $mainForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor }

        # Use global script variables for scopes
        $scopes = $script:graphScopes
        if (-not $scopes) {
            # Include scopes required for audit and sign-in logs
            $scopes = @(
                "User.Read.All",
                "Directory.Read.All",
                "AuditLog.Read.All",
                "SecurityEvents.Read.All"
            )
        }

        # Clear any existing connections and cached tokens
        Disconnect-MgGraph -ErrorAction SilentlyContinue

        # Clear authentication context and token cache more thoroughly
        try {
            $graphSession = [Microsoft.Graph.PowerShell.Authentication.GraphSession]::Instance
            if ($graphSession -and $graphSession.AuthContext) {
                $graphSession.AuthContext.ClearTokenCache()
            }
        } catch {
            # Ignore errors clearing token cache
        }

        # Also try to clear any MSAL cache
        try {
            $msalCache = [Microsoft.Identity.Client.TokenCacheHelper]::GetCacheFilePath()
            if ($msalCache -and (Test-Path $msalCache)) {
                Remove-Item $msalCache -Force -ErrorAction SilentlyContinue
            }
        } catch {
            # Ignore errors clearing MSAL cache - method may not be available
        }

        # Disable broker/WAM so Connect-MgGraph uses the system browser instead of an embedded popup
        $env:AZURE_IDENTITY_DISABLE_BROKER = "true"
        $env:MSAL_DISABLE_BROKER = "1"
        $env:MSAL_EXPERIMENTAL_DISABLE_BROKER = "1"

        # Ensure the cache directory is empty before starting a new auth flow
        if ($env:MSAL_CACHE_DIR) {
            try {
                if (-not (Test-Path $env:MSAL_CACHE_DIR)) {
                    New-Item -ItemType Directory -Path $env:MSAL_CACHE_DIR -Force -ErrorAction SilentlyContinue | Out-Null
                }
                Get-ChildItem -Path $env:MSAL_CACHE_DIR -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
            } catch {
                # Ignore cache cleanup errors to avoid blocking auth
            }
        }

        # Connect to Graph with improved error handling
        # With broker disabled above, this will use interactive browser auth and automatically open the browser
        try {
            $global:graphConnection = Connect-MgGraph -Scopes $scopes -NoWelcome -ErrorAction Stop
        } catch {
            $msg = $_.Exception.Message
            # Retry if there was a parameter issue
            if ($msg -match "parameter name|matches parameter name") {
                try {
                    $global:graphConnection = Connect-MgGraph -Scopes $scopes -NoWelcome -ErrorAction Stop
                } catch {
                    throw
                }
            }
            elseif ($msg -match "Method not found|Could not load type|BaseAbstractApplicationBuilder.*WithLogging") {
                Write-Warning "Graph module conflict detected. Attempting automatic repair..."
                if ($statusLabel) { $statusLabel.Text = "Fixing Graph modules..." }
                $fixOk = $false
                try { $fixOk = Fix-GraphModuleConflicts -statusLabel $statusLabel } catch {}
                if ($fixOk) {
                    Write-Host "Retrying Graph connection after repair..." -ForegroundColor Yellow
                    try {
                        $global:graphConnection = Connect-MgGraph -Scopes $scopes -ForceRefresh -ErrorAction Stop
                    } catch {
                        try { $global:graphConnection = Connect-MgGraph -Scopes $scopes -ErrorAction Stop } catch {}
                    }
                }
                # If still not connected, fall back to Device Code flow (bypasses InteractiveBrowserCredential path)
                if (-not $global:graphConnection) {
                    Write-Warning "Falling back to Device Code authentication for Microsoft Graph..."
                    try {
                        $global:graphConnection = Connect-MgGraph -Scopes $scopes -UseDeviceCode -ErrorAction Stop
                        Write-Host "Connected to Graph via Device Code." -ForegroundColor Green
                    } catch {
                        throw
                    }
                }
            } else {
                throw
            }
        }

        # Import required Microsoft Graph modules after connection
        $modulesToImport = @(
            "Microsoft.Graph.Authentication",
            "Microsoft.Graph.Users",
            "Microsoft.Graph.Users.Actions",
            "Microsoft.Graph.Identity.SignIns",
            "Microsoft.Graph.Reports"
        )

        foreach ($module in $modulesToImport) {
            try {
                if (Get-Module -ListAvailable -Name $module) {
                    Import-Module $module -ErrorAction Stop
                    Write-Host "Successfully imported module: $module" -ForegroundColor Green
                } else {
                    Write-Warning "Module $module not available. Installing..."
                    Install-Module -Name $module -Scope CurrentUser -Repository PSGallery -Force -AllowClobber -ErrorAction Stop
                    Import-Module $module -ErrorAction Stop
                    Write-Host "Installed and imported module: $module" -ForegroundColor Green
                }
            } catch {
                Write-Error "Failed to import module $module`: $($_.Exception.Message)"
                if ($statusLabel) { $statusLabel.Text = "Warning: Failed to import $module" }
            }
        }

        # Verify that required functions are available
        $requiredFunctions = @(
            "Update-MgUser",
            "Revoke-MgUserSignInSession",
            "Get-MgUser",
            "Get-MgContext",
            "Get-MgAuditLogDirectoryAudit",
            "Get-MgAuditLogSignIn"
        )

        $missingFunctions = @()
        foreach ($function in $requiredFunctions) {
            if (-not (Get-Command $function -ErrorAction SilentlyContinue)) {
                $missingFunctions += $function
            }
        }

        if ($missingFunctions.Count -gt 0) {
            Write-Warning "Some required Microsoft Graph functions are not available: $($missingFunctions -join ', ')"
            if ($statusLabel) { $statusLabel.Text = "Warning: Missing functions: $($missingFunctions -join ', ')" }
        } else {
            Write-Host "All required Microsoft Graph functions are available." -ForegroundColor Green
        }

        $global:graphConnectionAttempted = $true
        if ($statusLabel) { $statusLabel.Text = "Connected to Microsoft Graph and modules loaded." }
        return $true
    } catch {
        $ex = $_.Exception
        $errorMessage = $ex.Message

        # Check if this is a module version conflict
        if ($errorMessage -match "Method not found|Could not load type|Assembly.*not found") {
            $fixMessage = "This appears to be a Microsoft Graph module version conflict.`n`n" +
                         "To fix this issue, please run:`n`n" +
                         "1. Uninstall-Module Microsoft.Graph* -AllVersions -Force`n" +
                         "2. Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force`n" +
                         "3. Install-Module Microsoft.Graph.Users -Scope CurrentUser -Force`n" +
                         "4. Install-Module Microsoft.Graph.Identity.SignIns -Scope CurrentUser -Force`n" +
                         "5. Install-Module Microsoft.Graph.Reports -Scope CurrentUser -Force`n" +
                         "6. Restart PowerShell and try again"

            [System.Windows.Forms.MessageBox]::Show("MODULE VERSION CONFLICT DETECTED`n`n$fixMessage", "Microsoft Graph Connection Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        } else {
            [System.Windows.Forms.MessageBox]::Show("Connect-GraphService ERROR: $($ex.Message)", "DEBUG: GraphOnline", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }

        if ($statusLabel) { $statusLabel.Text = "Microsoft Graph connection failed." }
        Write-Error "Microsoft Graph connection failed: $($ex.Message)"
        return $false
    } finally {
        if ($mainForm) { $mainForm.Cursor = [System.Windows.Forms.Cursors]::Default }
    }
}

Export-ModuleMember -Function Test-GraphModules,Install-GraphModules,Fix-GraphModuleConflicts,Connect-GraphService 