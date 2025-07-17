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
            $scopes = @("User.Read.All", "User.ReadWrite.All", "SecurityEvents.Read.All", "SecurityEvents.ReadWrite.All")
        }
        # Connect to Graph
        $global:graphConnection = Connect-MgGraph -Scopes $scopes -ErrorAction Stop
        $global:graphConnectionAttempted = $true
        if ($statusLabel) { $statusLabel.Text = "Connected to Microsoft Graph." }
        return $true
    } catch {
        $ex = $_.Exception
        if ($statusLabel) { $statusLabel.Text = "Microsoft Graph connection failed." }
        Write-Error "Microsoft Graph connection failed: $($ex.Message)"
        return $false
    } finally {
        if ($mainForm) { $mainForm.Cursor = [System.Windows.Forms.Cursors]::Default }
    }
}

Export-ModuleMember -Function Test-GraphModules,Install-GraphModules,Connect-GraphService 