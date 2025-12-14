<#
.SYNOPSIS
Bulk Tenant Report Exporter - Standalone Application

.DESCRIPTION
Standalone PowerShell GUI application for exporting security investigation reports for multiple tenants.
Allows dynamic tenant addition and sequential authentication for bulk report generation.

.NOTES
Version: 1.0
Requires: PowerShell 5.1+, ExchangeOnlineManagement, Microsoft.Graph modules
Permissions: Exchange administrative privileges and Microsoft Graph permissions

.LINK
Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force
#>

#Requires -Version 5.1

# Set error action preference
$ErrorActionPreference = "Stop"

# Load Windows Forms assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Get script root directory
$script:scriptRoot = $PSScriptRoot
if (-not $script:scriptRoot) {
    $script:scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
}

# Function to safely import modules
function Safe-ImportModule {
    param([string]$modulePath)
    
    try {
        $moduleName = [System.IO.Path]::GetFileNameWithoutExtension($modulePath)
        
        # Remove the module if it's already loaded to force reload
        if (Get-Module -Name $moduleName -ErrorAction SilentlyContinue) {
            Remove-Module -Name $moduleName -Force -ErrorAction SilentlyContinue
        }
        
        Import-Module $modulePath -Global -ErrorAction Stop
        Write-Host "Successfully imported module: $moduleName" -ForegroundColor Green
    } catch {
        $errorMsg = "Failed to import module: $modulePath`nError: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show($errorMsg, "Module Import Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        Write-Error $errorMsg
        exit 1
    }
}

# Import required modules
Write-Host "Loading required modules..." -ForegroundColor Cyan
Safe-ImportModule "$script:scriptRoot\Modules\ExportUtils.psm1"
Safe-ImportModule "$script:scriptRoot\Modules\GraphOnline.psm1"
Safe-ImportModule "$script:scriptRoot\Modules\BrowserIntegration.psm1"
Safe-ImportModule "$script:scriptRoot\Modules\Settings.psm1"
Write-Host "All modules loaded successfully." -ForegroundColor Green

# Load settings (shared with main application if it exists)
$settings = $null
$settingsPath = Join-Path $script:scriptRoot "settings.json"
if (Test-Path $settingsPath) {
    try {
        $settings = Get-AppSettings
        Write-Host "Settings loaded from: $settingsPath" -ForegroundColor Green
    } catch {
        Write-Warning "Could not load settings from $settingsPath : $($_.Exception.Message)"
        $settings = $null
    }
}

# Initialize script-scope variables
$script:clientProcesses = @{}
$script:nextClientNumber = 1
$script:readinessCheckCount = @{}
$script:clientAuthStates = @{}
$script:clientAuthControls = @{}
$script:clientCacheDirs = @{}

# Create Bulk Tenant Exporter form
$bulkForm = New-Object System.Windows.Forms.Form
$bulkForm.Text = "Bulk Tenant Report Exporter"
$bulkForm.Size = New-Object System.Drawing.Size(900, 750)
$bulkForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
$bulkForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
$bulkForm.MaximizeBox = $true

# Create main panel
$bulkMainPanel = New-Object System.Windows.Forms.Panel
$bulkMainPanel.Dock = 'Fill'
$bulkMainPanel.Padding = New-Object System.Windows.Forms.Padding(15)

# Title
$bulkTitleLabel = New-Object System.Windows.Forms.Label
$bulkTitleLabel.Text = "Bulk Tenant Report Exporter"
$bulkTitleLabel.Font = New-Object System.Drawing.Font('Segoe UI', 16, [System.Drawing.FontStyle]::Bold)
$bulkTitleLabel.Location = New-Object System.Drawing.Point(15, 15)
$bulkTitleLabel.Size = New-Object System.Drawing.Size(500, 35)

# Description
$bulkDescLabel = New-Object System.Windows.Forms.Label
$bulkDescLabel.Text = "Export security investigation reports for multiple tenants. You will be prompted to authenticate to each tenant sequentially.`nReports will be saved in separate folders for each tenant."
$bulkDescLabel.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$bulkDescLabel.Location = New-Object System.Drawing.Point(15, 55)
$bulkDescLabel.Size = New-Object System.Drawing.Size(600, 40)
$bulkDescLabel.MaximumSize = New-Object System.Drawing.Size(600, 0)
$bulkDescLabel.AutoSize = $true

# Configuration GroupBox
$bulkConfigGroupBox = New-Object System.Windows.Forms.GroupBox
$bulkConfigGroupBox.Text = "Configuration"
$bulkConfigGroupBox.Location = New-Object System.Drawing.Point(15, 110)
$bulkConfigGroupBox.Size = New-Object System.Drawing.Size(400, 150)

# Investigator Name
$bulkInvestigatorLabel = New-Object System.Windows.Forms.Label
$bulkInvestigatorLabel.Text = "Investigator Name:"
$bulkInvestigatorLabel.Location = New-Object System.Drawing.Point(20, 25)
$bulkInvestigatorLabel.Size = New-Object System.Drawing.Size(150, 20)

$bulkInvestigatorTextBox = New-Object System.Windows.Forms.TextBox
$bulkInvestigatorTextBox.Location = New-Object System.Drawing.Point(180, 23)
$bulkInvestigatorTextBox.Size = New-Object System.Drawing.Size(200, 20)
if ($settings -and $settings.InvestigatorName) { 
    $bulkInvestigatorTextBox.Text = $settings.InvestigatorName 
} else { 
    $bulkInvestigatorTextBox.Text = "Security Administrator" 
}

# Company Name
$bulkCompanyLabel = New-Object System.Windows.Forms.Label
$bulkCompanyLabel.Text = "Company Name:"
$bulkCompanyLabel.Location = New-Object System.Drawing.Point(20, 55)
$bulkCompanyLabel.Size = New-Object System.Drawing.Size(150, 20)

$bulkCompanyTextBox = New-Object System.Windows.Forms.TextBox
$bulkCompanyTextBox.Location = New-Object System.Drawing.Point(180, 53)
$bulkCompanyTextBox.Size = New-Object System.Drawing.Size(200, 20)
if ($settings -and $settings.CompanyName) { 
    $bulkCompanyTextBox.Text = $settings.CompanyName 
} else { 
    $bulkCompanyTextBox.Text = "Organization" 
}

# Days Back
$bulkDaysLabel = New-Object System.Windows.Forms.Label
$bulkDaysLabel.Text = "Days Back (Message Trace):"
$bulkDaysLabel.Location = New-Object System.Drawing.Point(20, 85)
$bulkDaysLabel.Size = New-Object System.Drawing.Size(150, 20)

$bulkDaysComboBox = New-Object System.Windows.Forms.ComboBox
$bulkDaysComboBox.Location = New-Object System.Drawing.Point(180, 83)
$bulkDaysComboBox.Size = New-Object System.Drawing.Size(100, 20)
$bulkDaysComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$bulkDaysComboBox.Items.AddRange(@("1", "3", "5", "7", "10", "14", "30"))
$bulkDaysComboBox.SelectedIndex = 4  # Default to 10 days

$bulkConfigGroupBox.Controls.AddRange(@($bulkInvestigatorLabel, $bulkInvestigatorTextBox, $bulkCompanyLabel, $bulkCompanyTextBox, $bulkDaysLabel, $bulkDaysComboBox))

# Report Selection section
$bulkReportsGroupBox = New-Object System.Windows.Forms.GroupBox
$bulkReportsGroupBox.Text = "Select Reports to Export"
$bulkReportsGroupBox.Location = New-Object System.Drawing.Point(15, 300)
$bulkReportsGroupBox.Size = New-Object System.Drawing.Size(400, 320)

# Create scrollable panel inside GroupBox
$bulkReportsScrollPanel = New-Object System.Windows.Forms.Panel
$bulkReportsScrollPanel.Location = New-Object System.Drawing.Point(10, 20)
$bulkReportsScrollPanel.Size = New-Object System.Drawing.Size(380, 290)
$bulkReportsScrollPanel.AutoScroll = $true
$bulkReportsScrollPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::None

# Select All / Deselect All buttons
$bulkSelectAllBtn = New-Object System.Windows.Forms.Button
$bulkSelectAllBtn.Text = "Select All"
$bulkSelectAllBtn.Location = New-Object System.Drawing.Point(10, 5)
$bulkSelectAllBtn.Size = New-Object System.Drawing.Size(80, 25)

$bulkDeselectAllBtn = New-Object System.Windows.Forms.Button
$bulkDeselectAllBtn.Text = "Deselect All"
$bulkDeselectAllBtn.Location = New-Object System.Drawing.Point(100, 5)
$bulkDeselectAllBtn.Size = New-Object System.Drawing.Size(90, 25)

# Checkboxes for each report type
$bulkMessageTraceCheckBox = New-Object System.Windows.Forms.CheckBox
$bulkMessageTraceCheckBox.Text = "Message Trace (last 10 days)"
$bulkMessageTraceCheckBox.Location = New-Object System.Drawing.Point(10, 40)
$bulkMessageTraceCheckBox.Size = New-Object System.Drawing.Size(360, 20)
$bulkMessageTraceCheckBox.Checked = $true

$bulkInboxRulesCheckBox = New-Object System.Windows.Forms.CheckBox
$bulkInboxRulesCheckBox.Text = "Inbox Rules"
$bulkInboxRulesCheckBox.Location = New-Object System.Drawing.Point(10, 65)
$bulkInboxRulesCheckBox.Size = New-Object System.Drawing.Size(360, 20)
$bulkInboxRulesCheckBox.Checked = $true

$bulkTransportRulesCheckBox = New-Object System.Windows.Forms.CheckBox
$bulkTransportRulesCheckBox.Text = "Transport Rules"
$bulkTransportRulesCheckBox.Location = New-Object System.Drawing.Point(10, 90)
$bulkTransportRulesCheckBox.Size = New-Object System.Drawing.Size(360, 20)
$bulkTransportRulesCheckBox.Checked = $true

$bulkMailFlowCheckBox = New-Object System.Windows.Forms.CheckBox
$bulkMailFlowCheckBox.Text = "Mail Flow Connectors"
$bulkMailFlowCheckBox.Location = New-Object System.Drawing.Point(10, 115)
$bulkMailFlowCheckBox.Size = New-Object System.Drawing.Size(360, 20)
$bulkMailFlowCheckBox.Checked = $true

$bulkMailboxForwardingCheckBox = New-Object System.Windows.Forms.CheckBox
$bulkMailboxForwardingCheckBox.Text = "Mailbox Forwarding & Delegation"
$bulkMailboxForwardingCheckBox.Location = New-Object System.Drawing.Point(10, 140)
$bulkMailboxForwardingCheckBox.Size = New-Object System.Drawing.Size(360, 20)
$bulkMailboxForwardingCheckBox.Checked = $true

$bulkAuditLogsCheckBox = New-Object System.Windows.Forms.CheckBox
$bulkAuditLogsCheckBox.Text = "Audit Logs"
$bulkAuditLogsCheckBox.Location = New-Object System.Drawing.Point(10, 165)
$bulkAuditLogsCheckBox.Size = New-Object System.Drawing.Size(360, 20)
$bulkAuditLogsCheckBox.Checked = $true

$bulkCaPoliciesCheckBox = New-Object System.Windows.Forms.CheckBox
$bulkCaPoliciesCheckBox.Text = "Conditional Access Policies"
$bulkCaPoliciesCheckBox.Location = New-Object System.Drawing.Point(10, 190)
$bulkCaPoliciesCheckBox.Size = New-Object System.Drawing.Size(360, 20)
$bulkCaPoliciesCheckBox.Checked = $true

$bulkAppRegistrationsCheckBox = New-Object System.Windows.Forms.CheckBox
$bulkAppRegistrationsCheckBox.Text = "App Registrations"
$bulkAppRegistrationsCheckBox.Location = New-Object System.Drawing.Point(10, 215)
$bulkAppRegistrationsCheckBox.Size = New-Object System.Drawing.Size(360, 20)
$bulkAppRegistrationsCheckBox.Checked = $true

$bulkSignInLogsCheckBox = New-Object System.Windows.Forms.CheckBox
$bulkSignInLogsCheckBox.Text = "Sign-In Logs"
$bulkSignInLogsCheckBox.Location = New-Object System.Drawing.Point(10, 240)
$bulkSignInLogsCheckBox.Size = New-Object System.Drawing.Size(360, 20)
$bulkSignInLogsCheckBox.Checked = $true

$bulkSignInLogsDaysLabel = New-Object System.Windows.Forms.Label
$bulkSignInLogsDaysLabel.Text = "Sign-In Logs Days:"
$bulkSignInLogsDaysLabel.Location = New-Object System.Drawing.Point(30, 265)
$bulkSignInLogsDaysLabel.Size = New-Object System.Drawing.Size(120, 20)

$bulkSignInLogsDaysComboBox = New-Object System.Windows.Forms.ComboBox
$bulkSignInLogsDaysComboBox.Location = New-Object System.Drawing.Point(160, 263)
$bulkSignInLogsDaysComboBox.Size = New-Object System.Drawing.Size(100, 20)
$bulkSignInLogsDaysComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$bulkSignInLogsDaysComboBox.Items.AddRange(@("1 day", "7 days", "30 days"))
$bulkSignInLogsDaysComboBox.SelectedIndex = 1  # Default to 7 days
$bulkSignInLogsDaysComboBox.Enabled = $bulkSignInLogsCheckBox.Checked

$bulkSignInLogsCheckBox.add_CheckedChanged({
    $bulkSignInLogsDaysComboBox.Enabled = $bulkSignInLogsCheckBox.Checked
})

# Select All button click handler
$bulkSelectAllBtn.add_Click({
    $bulkMessageTraceCheckBox.Checked = $true
    $bulkInboxRulesCheckBox.Checked = $true
    $bulkTransportRulesCheckBox.Checked = $true
    $bulkMailFlowCheckBox.Checked = $true
    $bulkMailboxForwardingCheckBox.Checked = $true
    $bulkAuditLogsCheckBox.Checked = $true
    $bulkCaPoliciesCheckBox.Checked = $true
    $bulkAppRegistrationsCheckBox.Checked = $true
    $bulkSignInLogsCheckBox.Checked = $true
})

# Deselect All button click handler
$bulkDeselectAllBtn.add_Click({
    $bulkMessageTraceCheckBox.Checked = $false
    $bulkInboxRulesCheckBox.Checked = $false
    $bulkTransportRulesCheckBox.Checked = $false
    $bulkMailFlowCheckBox.Checked = $false
    $bulkMailboxForwardingCheckBox.Checked = $false
    $bulkAuditLogsCheckBox.Checked = $false
    $bulkCaPoliciesCheckBox.Checked = $false
    $bulkAppRegistrationsCheckBox.Checked = $false
    $bulkSignInLogsCheckBox.Checked = $false
})

# Add all controls to scrollable panel
$bulkReportsScrollPanel.Controls.AddRange(@(
    $bulkSelectAllBtn, $bulkDeselectAllBtn,
    $bulkMessageTraceCheckBox, $bulkInboxRulesCheckBox, $bulkTransportRulesCheckBox,
    $bulkMailFlowCheckBox, $bulkMailboxForwardingCheckBox, $bulkAuditLogsCheckBox,
    $bulkCaPoliciesCheckBox, $bulkAppRegistrationsCheckBox,
    $bulkSignInLogsCheckBox, $bulkSignInLogsDaysLabel, $bulkSignInLogsDaysComboBox
))

# Add scrollable panel to GroupBox
$bulkReportsGroupBox.Controls.Add($bulkReportsScrollPanel)

# Progress Label
$bulkProgressLabel = New-Object System.Windows.Forms.Label
$bulkProgressLabel.Text = "Ready to start bulk export..."
$bulkProgressLabel.Location = New-Object System.Drawing.Point(430, 190)
$bulkProgressLabel.Size = New-Object System.Drawing.Size(400, 20)
$bulkProgressLabel.ForeColor = [System.Drawing.Color]::Blue

# Status TextBox (for detailed progress)
$bulkStatusTextBox = New-Object System.Windows.Forms.TextBox
$bulkStatusTextBox.Multiline = $true
$bulkStatusTextBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
$bulkStatusTextBox.ReadOnly = $true
$bulkStatusTextBox.Location = New-Object System.Drawing.Point(430, 220)
$bulkStatusTextBox.Size = New-Object System.Drawing.Size(400, 400)
$bulkStatusTextBox.Font = New-Object System.Drawing.Font('Consolas', 9)

# Start Export Button (opens authentication console)
$bulkStartButton = New-Object System.Windows.Forms.Button
$bulkStartButton.Text = "Open Authentication Console"
$bulkStartButton.Font = New-Object System.Drawing.Font('Segoe UI', 12, [System.Drawing.FontStyle]::Bold)
$bulkStartButton.Location = New-Object System.Drawing.Point(430, 110)
$bulkStartButton.Size = New-Object System.Drawing.Size(280, 50)
$bulkStartButton.BackColor = [System.Drawing.Color]::FromArgb(46, 125, 50)
$bulkStartButton.ForeColor = [System.Drawing.Color]::White

# Close Button
$bulkCloseButton = New-Object System.Windows.Forms.Button
$bulkCloseButton.Text = "Close"
$bulkCloseButton.Location = New-Object System.Drawing.Point(430, 640)
$bulkCloseButton.Size = New-Object System.Drawing.Size(100, 30)
$bulkCloseButton.add_Click({
    $bulkForm.Close()
})

# Add controls to main panel
$bulkMainPanel.Controls.AddRange(@(
    $bulkTitleLabel, $bulkDescLabel, $bulkConfigGroupBox, $bulkReportsGroupBox,
    $bulkProgressLabel, $bulkStatusTextBox, $bulkStartButton, $bulkCloseButton
))

# Add panel to form
$bulkForm.Controls.Add($bulkMainPanel)

# Start Export button click handler - Opens Authentication Console
$bulkStartButton.add_Click({
    $investigator = if ($bulkInvestigatorTextBox.Text -and $bulkInvestigatorTextBox.Text.Trim().Length -gt 0) { $bulkInvestigatorTextBox.Text } else { 'Security Administrator' }
    $company = if ($bulkCompanyTextBox.Text -and $bulkCompanyTextBox.Text.Trim().Length -gt 0) { $bulkCompanyTextBox.Text } else { 'Organization' }
    $days = [int]$bulkDaysComboBox.SelectedItem

    # Parse sign-in logs time range
    $signInLogsDays = 7
    $selectedRange = $bulkSignInLogsDaysComboBox.SelectedItem
    if ($selectedRange -eq "1 day") { $signInLogsDays = 1 }
    elseif ($selectedRange -eq "7 days") { $signInLogsDays = 7 }
    elseif ($selectedRange -eq "30 days") { $signInLogsDays = 30 }

    # Get report selections from checkboxes
    $reportSelections = @{
        IncludeMessageTrace = $bulkMessageTraceCheckBox.Checked
        IncludeInboxRules = $bulkInboxRulesCheckBox.Checked
        IncludeTransportRules = $bulkTransportRulesCheckBox.Checked
        IncludeMailFlowConnectors = $bulkMailFlowCheckBox.Checked
        IncludeMailboxForwarding = $bulkMailboxForwardingCheckBox.Checked
        IncludeAuditLogs = $bulkAuditLogsCheckBox.Checked
        IncludeConditionalAccessPolicies = $bulkCaPoliciesCheckBox.Checked
        IncludeAppRegistrations = $bulkAppRegistrationsCheckBox.Checked
        IncludeSignInLogs = $bulkSignInLogsCheckBox.Checked
        SignInLogsDaysBack = $signInLogsDays
    }

    # Validate at least one report is selected
    $anySelected = $false
    foreach ($key in $reportSelections.Keys) {
        if ($key -ne 'SignInLogsDaysBack' -and $reportSelections[$key]) { $anySelected = $true; break }
    }
    if (-not $anySelected) {
        [System.Windows.Forms.MessageBox]::Show("Please select at least one report to export.", "No Reports Selected", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }

    # Close the configuration form and open authentication console
    $bulkForm.Hide()
    
    # Create temp directory for scripts, status files, and command files
    $tempDir = Join-Path $env:TEMP "ExchangeOnlineAnalyzer_BulkReports_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
    try {
        $null = New-Item -ItemType Directory -Path $tempDir -Force -ErrorAction Stop
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Failed to create temp directory: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                $bulkForm.ShowDialog() | Out-Null
                return
            }

    # Save report selections to JSON file (shared by all clients)
    $reportSelectionsFile = Join-Path $tempDir "ReportSelections.json"
    try {
        $reportSelections | ConvertTo-Json -ErrorAction Stop | Out-File -FilePath $reportSelectionsFile -Encoding UTF8 -ErrorAction Stop
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Failed to create report selections file: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                $bulkForm.ShowDialog() | Out-Null
                return
            }

    # Create the worker script that waits for commands and handles auth/reports
    $workerScriptContent = @"
param(
    [int]`$ClientNumber,
    [string]`$ScriptRoot,
    [string]`$InvestigatorName,
    [string]`$CompanyName,
    [int]`$DaysBack,
    [string]`$ReportSelectionsFile,
    [string]`$StatusFile,
    [string]`$ResultFile,
    [string]`$CommandDir
)

function Write-Status {
    param([string]`$Message)
    `$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "[`$timestamp] `$Message" | Out-File -FilePath `$StatusFile -Append -Encoding UTF8
    Write-Host "[Client `$ClientNumber] `$Message" -ForegroundColor Cyan
}

function Write-CommandResponse {
    param([string]`$Response)
    `$responseFile = Join-Path `$CommandDir "Client`$(`$ClientNumber)_Response.txt"
    `$Response | Out-File -FilePath `$responseFile -Encoding UTF8 -Force
}

try {
    # Set window title
    try {
        `$Host.UI.RawUI.WindowTitle = "Client `$ClientNumber - Waiting for Authentication Commands"
    } catch {}
    
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host "CLIENT `$ClientNumber - PowerShell Session" -ForegroundColor Cyan
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host "PowerShell session starting for Client `$ClientNumber..." -ForegroundColor Yellow
    Write-Host ""
    
    # Initialize status file
    try {
        "STARTING" | Out-File -FilePath `$ResultFile -Encoding UTF8 -ErrorAction Stop
        Write-Host "Status file initialized: `$ResultFile" -ForegroundColor Gray
    } catch {
        Write-Host "WARNING: Could not write status file: `$(`$_.Exception.Message)" -ForegroundColor Yellow
    }
    
    Write-Status "Client `$ClientNumber PowerShell session started"
    Write-Host "This window is associated with Client `$ClientNumber" -ForegroundColor Yellow
    Write-Host "Waiting for authentication commands from GUI..." -ForegroundColor Yellow
    Write-Host ""
    
    # Create isolated cache directory for this client
    `$cacheDir = Join-Path `$env:TEMP "ExchangeOnlineAnalyzer_Client`$ClientNumber_Cache_`$(Get-Date -Format 'yyyyMMdd_HHmmss')"
    `$null = New-Item -ItemType Directory -Path `$cacheDir -Force -ErrorAction Stop
    `$env:IDENTITY_SERVICE_CACHE_DIR = `$cacheDir
    `$env:MSAL_CACHE_DIR = `$cacheDir
    Write-Status "Using isolated cache directory: `$cacheDir"
    Write-Host "Cache directory: `$cacheDir" -ForegroundColor Gray
    Write-Host ""
    
    # Import required modules
    Write-Status "Importing modules..."
    Write-Host "Importing modules..." -ForegroundColor Cyan
    Import-Module "`$ScriptRoot\Modules\ExportUtils.psm1" -Force -ErrorAction Stop
    Import-Module "`$ScriptRoot\Modules\GraphOnline.psm1" -Force -ErrorAction SilentlyContinue
    Import-Module "`$ScriptRoot\Modules\BrowserIntegration.psm1" -Force -ErrorAction SilentlyContinue
    Write-Status "Modules imported successfully"
    Write-Host ""
    
    # Load report selections from JSON
    `$reportSelections = @{}
    if (Test-Path `$ReportSelectionsFile) {
        `$jsonObj = Get-Content `$ReportSelectionsFile -Raw | ConvertFrom-Json
        `$reportSelections = @{
            IncludeMessageTrace = if (`$null -ne `$jsonObj.IncludeMessageTrace) { `$jsonObj.IncludeMessageTrace } else { `$false }
            IncludeInboxRules = if (`$null -ne `$jsonObj.IncludeInboxRules) { `$jsonObj.IncludeInboxRules } else { `$false }
            IncludeTransportRules = if (`$null -ne `$jsonObj.IncludeTransportRules) { `$jsonObj.IncludeTransportRules } else { `$false }
            IncludeMailFlowConnectors = if (`$null -ne `$jsonObj.IncludeMailFlowConnectors) { `$jsonObj.IncludeMailFlowConnectors } else { `$false }
            IncludeMailboxForwarding = if (`$null -ne `$jsonObj.IncludeMailboxForwarding) { `$jsonObj.IncludeMailboxForwarding } else { `$false }
            IncludeAuditLogs = if (`$null -ne `$jsonObj.IncludeAuditLogs) { `$jsonObj.IncludeAuditLogs } else { `$false }
            IncludeConditionalAccessPolicies = if (`$null -ne `$jsonObj.IncludeConditionalAccessPolicies) { `$jsonObj.IncludeConditionalAccessPolicies } else { `$false }
            IncludeAppRegistrations = if (`$null -ne `$jsonObj.IncludeAppRegistrations) { `$jsonObj.IncludeAppRegistrations } else { `$false }
            IncludeSignInLogs = if (`$null -ne `$jsonObj.IncludeSignInLogs) { `$jsonObj.IncludeSignInLogs } else { `$false }
            SignInLogsDaysBack = if (`$null -ne `$jsonObj.SignInLogsDaysBack) { `$jsonObj.SignInLogsDaysBack } else { 7 }
        }
    }
    
    `$graphAuthenticated = `$false
    `$exchangeAuthenticated = `$false
    `$tenantDisplayName = "Client`$ClientNumber"
    
    # Main command loop - wait for commands from GUI
    `$commandFile = Join-Path `$CommandDir "Client`$(`$ClientNumber)_Command.txt"
    `$pollInterval = 500  # milliseconds
    
    Write-Host "Ready! Waiting for Graph Auth command from GUI..." -ForegroundColor Green
    Write-Status "Ready! Waiting for Graph Auth command from GUI..."
    Write-Host "Command file: `$commandFile" -ForegroundColor Gray
    Write-Host "Polling every `$pollInterval ms for commands..." -ForegroundColor Gray
    Write-Host ""
    
    Write-Status "Command polling loop started - ready to receive commands"
    `$pollCount = 0
    while (`$true) {
        `$pollCount++
        if (`$pollCount % 100 -eq 0) {
            Write-Host "Still polling... (checked `$pollCount times)" -ForegroundColor DarkGray
        }
        
        if (Test-Path `$commandFile) {
            Write-Host "==========================================" -ForegroundColor Yellow
            Write-Host "Command file detected! Reading command..." -ForegroundColor Yellow
            Write-Host "Command file path: `$commandFile" -ForegroundColor Cyan
            Start-Sleep -Milliseconds 300  # Brief delay to ensure file is fully written
            `$command = Get-Content `$commandFile -Raw -ErrorAction SilentlyContinue | ForEach-Object { `$_.Trim() }
            Write-Host "Command received: '`$command'" -ForegroundColor Cyan
            Write-Host "Command length: `$(`$command.Length)" -ForegroundColor Gray
            Remove-Item `$commandFile -Force -ErrorAction SilentlyContinue
            Write-Host "Command file removed" -ForegroundColor Gray
            
            if (`$command -eq "GRAPH_AUTH") {
                Write-Host "==========================================" -ForegroundColor Yellow
                Write-Host "GRAPH AUTHENTICATION COMMAND RECEIVED" -ForegroundColor Yellow
                Write-Host "==========================================" -ForegroundColor Yellow
                Write-Status "Graph authentication command received"
                Write-CommandResponse "GRAPH_AUTH_STARTED"
                
                # Clear any existing sessions
                Write-Status "Clearing existing sessions..."
                Write-Host "Clearing existing sessions..." -ForegroundColor Cyan
                try { Disconnect-MgGraph -ErrorAction SilentlyContinue } catch {}
                
                # Graph Authentication
                Write-Host ""
                Write-Host "A browser window will open for Microsoft Graph authentication - this may take 10-30 seconds to appear." -ForegroundColor Yellow
                Write-Host "Please wait for the browser popup and complete the sign-in in your browser window." -ForegroundColor Yellow
                Write-Host ""
                Write-Status "Waiting for browser popup to appear (this may take 10-30 seconds)..."
    
                `$scopes = @(
                    "AuditLog.Read.All",
                    "User.Read.All",
                    "Directory.Read.All",
                    "Policy.Read.All",
                    "Application.Read.All",
                    "Reports.Read.All"
                )
    
                try {
                    Connect-MgGraph -Scopes `$scopes -ContextScope Process -ErrorAction Stop
                    `$mgContext = Get-MgContext -ErrorAction Stop
                    `$graphAuthenticated = `$true
                    Write-Status "Graph authentication successful! Tenant: `$(`$mgContext.TenantId)"
                    Write-Host "Graph authentication successful!" -ForegroundColor Green
                    Write-Host "Tenant ID: `$(`$mgContext.TenantId)" -ForegroundColor Cyan
                    
                    # Get tenant name
                    try {
                        `$ti = `$null
                        try { `$ti = Get-TenantIdentity } catch {}
                        if (`$ti -and `$ti.TenantDisplayName) {
                            `$tenantDisplayName = `$ti.TenantDisplayName
                        } elseif (`$ti -and `$ti.PrimaryDomain) {
                            `$tenantDisplayName = `$ti.PrimaryDomain
                        } else {
                            try {
                                `$org = Get-MgOrganization -ErrorAction SilentlyContinue | Select-Object -First 1
                                if (`$org -and `$org.DisplayName) {
                                    `$tenantDisplayName = `$org.DisplayName
                                }
                            } catch {}
                        }
                    } catch {}
                    
                    Write-Status "Tenant identified as: `$tenantDisplayName"
                    Write-Host "Tenant: `$tenantDisplayName" -ForegroundColor Cyan
                    Write-CommandResponse "GRAPH_AUTH_SUCCESS:`$tenantDisplayName"
                } catch {
                    Write-Status "ERROR: Graph authentication failed - `$(`$_.Exception.Message)"
                    Write-Host "ERROR: Graph authentication failed - `$(`$_.Exception.Message)" -ForegroundColor Red
                    Write-CommandResponse "GRAPH_AUTH_FAILED:`$(`$_.Exception.Message)"
                }
                
                Write-Host ""
                Write-Host "Waiting for Exchange Online Auth command from GUI..." -ForegroundColor Green
                Write-Host ""
                
            } elseif (`$command -eq "EXCHANGE_AUTH") {
                if (-not `$graphAuthenticated) {
                    Write-Host "ERROR: Graph authentication must be completed first!" -ForegroundColor Red
                    Write-CommandResponse "EXCHANGE_AUTH_FAILED:Graph authentication not completed"
                    continue
                }
                
                Write-Host "==========================================" -ForegroundColor Yellow
                Write-Host "EXCHANGE ONLINE AUTHENTICATION COMMAND RECEIVED" -ForegroundColor Yellow
                Write-Host "==========================================" -ForegroundColor Yellow
                Write-Status "Exchange Online authentication command received"
                Write-CommandResponse "EXCHANGE_AUTH_STARTED"
                
                # Exchange Online Authentication
                Write-Host ""
                Write-Host "Connecting to Exchange Online..." -ForegroundColor Yellow
                Write-Host "A browser window will open for authentication - this may take 10-30 seconds to appear." -ForegroundColor Yellow
                Write-Host "Please wait for the browser popup and complete the sign-in." -ForegroundColor Yellow
                Write-Host ""
                Write-Status "Waiting for browser popup to appear (this may take 10-30 seconds)..."
    
                try {
                    # Note: Connect-ExchangeOnline may take time to show the browser popup
                    Connect-ExchangeOnline -ShowBanner:`$false -ErrorAction Stop
                    `$exchangeAuthenticated = `$true
                    Write-Status "Exchange Online authentication successful!"
                    Write-Host "Exchange Online authentication successful!" -ForegroundColor Green
                    Write-CommandResponse "EXCHANGE_AUTH_SUCCESS"
                } catch {
                    Write-Status "ERROR: Exchange Online authentication failed - `$(`$_.Exception.Message)"
                    Write-Host "ERROR: Exchange Online authentication failed - `$(`$_.Exception.Message)" -ForegroundColor Red
                    Write-CommandResponse "EXCHANGE_AUTH_FAILED:`$(`$_.Exception.Message)"
                }
                
                Write-Host ""
                Write-Host "Waiting for Generate Reports command from GUI..." -ForegroundColor Green
                Write-Host ""
                
            } elseif (`$command -eq "GENERATE_REPORTS") {
                if (-not `$graphAuthenticated -or -not `$exchangeAuthenticated) {
                    Write-Host "ERROR: Both Graph and Exchange authentication must be completed first!" -ForegroundColor Red
                    Write-CommandResponse "GENERATE_REPORTS_FAILED:Authentication not completed"
                    continue
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
                try {
                    `$report = New-SecurityInvestigationReport -InvestigatorName `$InvestigatorName -CompanyName `$CompanyName -DaysBack `$DaysBack -StatusLabel `$null -MainForm `$null -IncludeMessageTrace `$reportSelections.IncludeMessageTrace -IncludeInboxRules `$reportSelections.IncludeInboxRules -IncludeTransportRules `$reportSelections.IncludeTransportRules -IncludeMailFlowConnectors `$reportSelections.IncludeMailFlowConnectors -IncludeMailboxForwarding `$reportSelections.IncludeMailboxForwarding -IncludeAuditLogs `$reportSelections.IncludeAuditLogs -IncludeConditionalAccessPolicies `$reportSelections.IncludeConditionalAccessPolicies -IncludeAppRegistrations `$reportSelections.IncludeAppRegistrations -IncludeSignInLogs `$reportSelections.IncludeSignInLogs -SignInLogsDaysBack `$reportSelections.SignInLogsDaysBack -SelectedUsers @()
                    Write-Status "Report generation function completed"
                    Write-Host "Report generation function completed successfully" -ForegroundColor Green
                } catch {
                    Write-Status "ERROR: Failed to generate report - `$(`$_.Exception.Message)"
                    Write-Host "ERROR: Failed to generate report - `$(`$_.Exception.Message)" -ForegroundColor Red
                    Write-Host "Error details: `$(`$_.Exception | Out-String)" -ForegroundColor Red
                    Write-CommandResponse "GENERATE_REPORTS_FAILED:`$(`$_.Exception.Message)"
                    continue
                }
                
                if (`$report -and `$report.OutputFolder) {
                    Write-Status "Report generation successful!"
                    Write-Host "`nReport generation successful!" -ForegroundColor Green
                    Write-Host "Reports saved to: `$(`$report.OutputFolder)" -ForegroundColor Green
                    "SUCCESS: `$(`$report.OutputFolder)" | Out-File -FilePath `$ResultFile -Encoding UTF8
                    Write-CommandResponse "GENERATE_REPORTS_SUCCESS:`$(`$report.OutputFolder)"
                } else {
                    Write-Status "Warning: Report generation returned no data"
                    Write-Host "Warning: Report generation returned no data" -ForegroundColor Yellow
                    `$defaultOutput = Join-Path ([Environment]::GetFolderPath('MyDocuments')) "ExchangeOnlineAnalyzer\SecurityInvestigation"
                    "NO_DATA: `$defaultOutput" | Out-File -FilePath `$ResultFile -Encoding UTF8
                    Write-CommandResponse "GENERATE_REPORTS_NO_DATA:`$defaultOutput"
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
                        `$runspace = [runspacefactory]::CreateRunspace()
                        `$runspace.Open()
                        `$ps = [PowerShell]::Create()
                        `$ps.Runspace = `$runspace
                        # Import module and disconnect
                        `$script = "Import-Module ExchangeOnlineManagement -ErrorAction SilentlyContinue; Disconnect-ExchangeOnline -Confirm:`$false -ErrorAction SilentlyContinue"
                        `$null = `$ps.AddScript(`$script)
                        `$handle = `$ps.BeginInvoke()
                        `$waited = `$handle.AsyncWaitHandle.WaitOne(5000)  # 5 second timeout
                        if (`$waited) {
                            try { `$ps.EndInvoke(`$handle) | Out-Null } catch {}
                        } else {
                            Write-Host "Exchange disconnect timed out (non-critical, continuing...)" -ForegroundColor Yellow
                            `$ps.Stop()
                        }
                        `$ps.Dispose()
                        `$runspace.Close()
                        `$runspace.Dispose()
                    }
                } catch {
                    Write-Host "Disconnect completed with warnings (non-critical)" -ForegroundColor Yellow
                }
                Write-Host ""
                Write-Host "==========================================" -ForegroundColor Green
                Write-Host "Client `$ClientNumber processing complete!" -ForegroundColor Green
                Write-Host "==========================================" -ForegroundColor Green
                Write-Host "This window will remain open. You may close it manually." -ForegroundColor Yellow
                Write-Host ""
                
                # Keep window open but stop polling
                break
            } elseif (`$command -eq "CANCEL_AUTH") {
                Write-Host "==========================================" -ForegroundColor Yellow
                Write-Host "CANCEL AUTHENTICATION COMMAND RECEIVED" -ForegroundColor Yellow
                Write-Host "==========================================" -ForegroundColor Yellow
                Write-Status "Cancelling authentication and resetting state..."
                
                # Reset authentication state
                `$graphAuthenticated = `$false
                `$exchangeAuthenticated = `$false
                `$tenantDisplayName = `$null
                
                # Disconnect any active sessions
                try {
                    Disconnect-MgGraph -ErrorAction SilentlyContinue
                } catch {}
                try {
                    if (Get-Command Disconnect-ExchangeOnline -ErrorAction SilentlyContinue) {
                        Disconnect-ExchangeOnline -Confirm:`$false -ErrorAction SilentlyContinue
                    }
                } catch {}
    
                Write-Status "Authentication cancelled and reset"
                Write-Host "Authentication cancelled and reset. Ready for new authentication attempt." -ForegroundColor Green
                Write-CommandResponse "CANCEL_AUTH_SUCCESS"
            } elseif (`$command -eq "EXIT") {
                Write-Host "Exit command received. Closing window..." -ForegroundColor Yellow
                break
            }
        }
        
        Start-Sleep -Milliseconds `$pollInterval
    }
    
} catch {
    `$errorMsg = `$_.Exception.Message
    Write-Status "ERROR: `$errorMsg"
    Write-Host "`nERROR: `$errorMsg" -ForegroundColor Red
    "ERROR: `$errorMsg" | Out-File -FilePath `$ResultFile -Encoding UTF8
    Write-Host "Press any key to exit..."
    `$null = `$Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 1
}
"@

    # Save the worker script
    $workerScriptFile = Join-Path $tempDir "BulkTenantWorker.ps1"
    try {
        $workerScriptContent | Out-File -FilePath $workerScriptFile -Encoding UTF8 -ErrorAction Stop
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Failed to create worker script: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                $bulkForm.ShowDialog() | Out-Null
                return
            }

    # Create command directory for inter-process communication
    $commandDir = Join-Path $tempDir "Commands"
    try {
        $null = New-Item -ItemType Directory -Path $commandDir -Force -ErrorAction Stop
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Failed to create command directory: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                $bulkForm.ShowDialog() | Out-Null
                return
            }
    
    # Store PowerShell processes for each client
    $script:clientProcesses = @{}
    $script:nextClientNumber = 1
    if (-not $script:readinessCheckCount) {
        $script:readinessCheckCount = @{}
    }
    
    # Create Authentication Console Form
    $authConsoleForm = New-Object System.Windows.Forms.Form
    $authConsoleForm.Text = "Bulk Tenant Authentication Console"
    $authConsoleForm.Size = New-Object System.Drawing.Size(1000, 700)
    $authConsoleForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $authConsoleForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
    $authConsoleForm.MaximizeBox = $true

    # Title
    $authTitleLabel = New-Object System.Windows.Forms.Label
    $authTitleLabel.Text = "Client Authentication Console"
    $authTitleLabel.Font = New-Object System.Drawing.Font('Segoe UI', 16, [System.Drawing.FontStyle]::Bold)
    $authTitleLabel.Location = New-Object System.Drawing.Point(15, 15)
    $authTitleLabel.Size = New-Object System.Drawing.Size(500, 35)

    # Instructions
    $authInstructionsLabel = New-Object System.Windows.Forms.Label
    $authInstructionsLabel.Text = "Click 'Add Tenant' to add a new tenant. Authenticate each client sequentially. Complete Graph authentication, then Exchange Online authentication for each client before proceeding to the next."
    $authInstructionsLabel.Font = New-Object System.Drawing.Font('Segoe UI', 9)
    $authInstructionsLabel.Location = New-Object System.Drawing.Point(15, 55)
    $authInstructionsLabel.Size = New-Object System.Drawing.Size(950, 40)
    $authInstructionsLabel.MaximumSize = New-Object System.Drawing.Size(950, 0)
    $authInstructionsLabel.AutoSize = $true

    # Add Tenant button
    $addTenantBtn = New-Object System.Windows.Forms.Button
    $addTenantBtn.Text = "Add Tenant"
    $addTenantBtn.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
    $addTenantBtn.Location = New-Object System.Drawing.Point(15, 100)
    $addTenantBtn.Size = New-Object System.Drawing.Size(150, 35)
    $addTenantBtn.BackColor = [System.Drawing.Color]::FromArgb(46, 125, 50)
    $addTenantBtn.ForeColor = [System.Drawing.Color]::White

    # Create Panel for client authentication rows
    $authPanel = New-Object System.Windows.Forms.Panel
    $authPanel.Location = New-Object System.Drawing.Point(15, 145)
    $authPanel.Size = New-Object System.Drawing.Size(960, 410)
    $authPanel.AutoScroll = $true
    $authPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle

    # Store client authentication state and controls
    $script:clientAuthStates = @{}
    $script:clientAuthControls = @{}
    $script:clientCacheDirs = @{}
    $clientRowHeight = 50
    $clientRowSpacing = 5

    # Add controls to form
    $authConsoleForm.Controls.AddRange(@($authTitleLabel, $authInstructionsLabel, $addTenantBtn, $authPanel))

    # Close button
    $authCloseBtn = New-Object System.Windows.Forms.Button
    $authCloseBtn.Text = "Close"
    $authCloseBtn.Location = New-Object System.Drawing.Point(880, 570)
    $authCloseBtn.Size = New-Object System.Drawing.Size(100, 40)
    $authCloseBtn.add_Click({
        # Send exit command to all active PowerShell processes
        foreach ($clientNum in $script:clientProcesses.Keys) {
            try {
                Send-CommandToSession -ClientNumber $clientNum -Command "EXIT" -TimeoutSeconds 5 | Out-Null
                Start-Sleep -Milliseconds 500
                $proc = $script:clientProcesses[$clientNum]
                if (-not $proc.HasExited) {
                    Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
                }
            } catch {}
        }
        $authConsoleForm.Close()
    })
    $authConsoleForm.Controls.Add($authCloseBtn)

    # Status text box
    $authStatusTextBox = New-Object System.Windows.Forms.TextBox
    $authStatusTextBox.Multiline = $true
    $authStatusTextBox.ReadOnly = $true
    $authStatusTextBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
    $authStatusTextBox.Location = New-Object System.Drawing.Point(15, 620)
    $authStatusTextBox.Size = New-Object System.Drawing.Size(965, 30)
    $authStatusTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $authConsoleForm.Controls.Add($authStatusTextBox)
    
    # Store in script scope for closure access
    $script:authStatusTextBox = $authStatusTextBox
    $script:authConsoleForm = $authConsoleForm
    $script:commandDir = $commandDir
    $script:tempDir = $tempDir
    $script:investigator = $investigator
    $script:company = $company
    $script:days = $days
    $script:reportSelections = $reportSelections
    $script:workerScriptFile = $workerScriptFile
    $script:reportSelectionsFile = $reportSelectionsFile
    $script:authPanel = $authPanel

    # Function to add a new tenant dynamically
    function Add-NewTenant {
        param([int]$ClientNumber)
        
        # Launch PowerShell process for this client
        $statusFile = Join-Path $script:tempDir "Client${ClientNumber}_Status.txt"
        $resultFile = Join-Path $script:tempDir "Client${ClientNumber}_Result.txt"
        
        # Build process arguments - use $script:scriptRoot instead of $PSScriptRoot
        $processArgs = "-NoProfile -ExecutionPolicy Bypass -File `"$script:workerScriptFile`" -ClientNumber $ClientNumber -ScriptRoot `"$script:scriptRoot`" -InvestigatorName `"$script:investigator`" -CompanyName `"$script:company`" -DaysBack $script:days -ReportSelectionsFile `"$script:reportSelectionsFile`" -StatusFile `"$statusFile`" -ResultFile `"$resultFile`" -CommandDir `"$script:commandDir`""

        try {
            # Try PowerShell 7 (pwsh.exe) first, fall back to Windows PowerShell (powershell.exe)
            $psExe = "pwsh.exe"
            if (-not (Get-Command $psExe -ErrorAction SilentlyContinue)) {
                $psExe = "powershell.exe"
            }

            # Use Normal window style so users can see progress
            $process = Start-Process -FilePath $psExe -ArgumentList $processArgs -PassThru -WindowStyle Normal
            $script:clientProcesses[$ClientNumber] = $process
            Write-Host "Launched Client $ClientNumber PowerShell window (PID: $($process.Id))" -ForegroundColor Green
            $script:authStatusTextBox.AppendText("Launched Client $ClientNumber PowerShell window (PID: $($process.Id))`r`n")
            $script:authStatusTextBox.ScrollToCaret()
            [System.Windows.Forms.Application]::DoEvents()
            
            # Wait a moment for PowerShell session to initialize
            Start-Sleep -Seconds 2
            
            # Verify process is still running
            try {
                $procCheck = Get-Process -Id $process.Id -ErrorAction Stop
                Write-Host "  Process verified running" -ForegroundColor Green
            } catch {
                Write-Host "  WARNING: Process may have exited immediately!" -ForegroundColor Yellow
                $script:authStatusTextBox.AppendText("WARNING: Client $ClientNumber process may have exited immediately!`r`n")
                $script:authStatusTextBox.ScrollToCaret()
                [System.Windows.Forms.Application]::DoEvents()
                return $false
            }
            
            # Start monitoring status file for readiness
            $statusFile = Join-Path $script:tempDir "Client${ClientNumber}_Status.txt"
            $readinessTimer = New-Object System.Windows.Forms.Timer
            $readinessTimer.Interval = 1000  # Check every second
            if (-not $script:readinessCheckCount) {
                $script:readinessCheckCount = @{}
            }
            $script:readinessCheckCount[$ClientNumber] = 0
            $maxReadinessChecks = 60  # Wait up to 60 seconds for readiness
            $capturedClientNum = $ClientNumber
            
            $readinessTimer.add_Tick({
                try {
                    $clientNum = $capturedClientNum
                    if (-not $clientNum) {
                        try { $readinessTimer.Stop(); $readinessTimer.Dispose() } catch {}
                        return
                    }
                    
                    # Ensure hashtable exists
                    if (-not $script:readinessCheckCount) {
                        $script:readinessCheckCount = @{}
                    }
                    
                    # Ensure key exists before accessing
                    if (-not $script:readinessCheckCount.ContainsKey($clientNum)) {
                        $script:readinessCheckCount[$clientNum] = 0
                    }
                    
                    $script:readinessCheckCount[$clientNum]++
                    $checkCount = $script:readinessCheckCount[$clientNum]
                    
                    if (-not $script:clientAuthControls -or -not $script:clientAuthControls.ContainsKey($clientNum)) {
                        try { $readinessTimer.Stop(); $readinessTimer.Dispose() } catch {}
                        return
                    }
                    
                    $controls = $script:clientAuthControls[$clientNum]
                    if (-not $controls) {
                        try { $readinessTimer.Stop(); $readinessTimer.Dispose() } catch {}
                        return
                    }
                    
                    $statusFilePath = Join-Path $script:tempDir "Client${clientNum}_Status.txt"
                
                    if (Test-Path $statusFilePath) {
                        try {
                            $statusLines = Get-Content $statusFilePath -Tail 5 -ErrorAction SilentlyContinue
                            $readyFound = $false
                            
                            foreach ($line in $statusLines) {
                                # Check for "Command polling loop started" - this means the loop is actually running
                                # Also check for "Ready!" as fallback
                                # Status file format: [timestamp] Message
                                if ($line -match "Command polling loop started|Ready!.*Waiting for Graph Auth|Modules imported successfully") {
                                    $readyFound = $true
                                    break
                                }
                            }
                            
                            if ($readyFound) {
                                # Wait an additional 2 seconds to ensure the polling loop is fully started and ready
                                Start-Sleep -Seconds 2
                                
                                # Double-check that the worker script is still running
                                if ($script:clientProcesses.ContainsKey($clientNum)) {
                                    $proc = $script:clientProcesses[$clientNum]
                                    try {
                                        $procInfo = Get-Process -Id $proc.Id -ErrorAction SilentlyContinue
                                        if (-not $procInfo -or $procInfo.HasExited) {
                                            if ($script:authStatusTextBox) {
                                                $script:authStatusTextBox.AppendText("WARNING: Client $clientNum PowerShell process has exited!`r`n")
                                                $script:authStatusTextBox.ScrollToCaret()
                                            }
                                            try {
                                                $readinessTimer.Stop()
                                                $readinessTimer.Dispose()
                                            } catch {}
                                            return
                                        }
                                    } catch {}
                                }
                                
                                # Worker script is ready - enable Graph Auth button
                                if ($controls -and $controls.GraphButton) {
                                    $controls.GraphButton.Enabled = $true
                                    $controls.GraphButton.Text = "Graph Auth"
                                }
                                if ($controls -and $controls.StatusLabel) {
                                    $controls.StatusLabel.Text = "Ready for Graph Auth"
                                    $controls.StatusLabel.ForeColor = [System.Drawing.Color]::Blue
                                }
                                if ($script:authStatusTextBox) {
                                    $script:authStatusTextBox.AppendText("Client $clientNum is ready for authentication (polling loop confirmed running).`r`n")
                                    $script:authStatusTextBox.ScrollToCaret()
                                }
                                [System.Windows.Forms.Application]::DoEvents()
                                try {
                                    $readinessTimer.Stop()
                                    $readinessTimer.Dispose()
                                } catch {}
                                return
                            }
                        } catch {
                            # Silently ignore errors reading status file
                        }
                    }
                    
                    # Update status to show we're waiting
                    if ($checkCount % 5 -eq 0) {
                        if ($controls -and $controls.StatusLabel) {
                            $controls.StatusLabel.Text = "Initializing... ($checkCount s)"
                            [System.Windows.Forms.Application]::DoEvents()
                        }
                    }
                    
                    # Timeout after max checks
                    if ($checkCount -ge $maxReadinessChecks) {
                        if ($controls -and $controls.GraphButton) {
                            $controls.GraphButton.Enabled = $true
                            $controls.GraphButton.Text = "Graph Auth"
                        }
                        if ($controls -and $controls.StatusLabel) {
                            $controls.StatusLabel.Text = "Ready (timeout)"
                            $controls.StatusLabel.ForeColor = [System.Drawing.Color]::Orange
                        }
                        if ($script:authStatusTextBox) {
                            $script:authStatusTextBox.AppendText("Client $clientNum readiness check timed out, but enabling Graph Auth button anyway.`r`n")
                            $script:authStatusTextBox.ScrollToCaret()
                        }
                        [System.Windows.Forms.Application]::DoEvents()
                        try {
                            $readinessTimer.Stop()
                            $readinessTimer.Dispose()
                        } catch {}
                    }
                    } catch {
                        # Silently handle any errors in the timer handler to prevent crashes
                        try {
                            if ($readinessTimer) {
                                $readinessTimer.Stop()
                                $readinessTimer.Dispose()
                            }
                        } catch {}
                    }
            })
            
            $readinessTimer.Start()
        } catch {
            $errorMsg = "Failed to launch Client $ClientNumber - $($_.Exception.Message)"
            Write-Host $errorMsg -ForegroundColor Red
            $script:authStatusTextBox.AppendText("ERROR: $errorMsg`r`n")
            $script:authStatusTextBox.ScrollToCaret()
            [System.Windows.Forms.Application]::DoEvents()
            return $false
        }
        
        # Create UI row for this client
        $clientRowHeight = 50
        $clientRowSpacing = 5
        $existingRows = ($script:clientAuthControls.Keys | Measure-Object).Count
        $yPos = $existingRows * ($clientRowHeight + $clientRowSpacing) + 10
        
        # Client label
        $clientLabel = New-Object System.Windows.Forms.Label
        $clientLabel.Text = "Client $ClientNumber"
        $clientLabel.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
        $clientLabel.Location = New-Object System.Drawing.Point(10, ($yPos + 15))
        $clientLabel.Size = New-Object System.Drawing.Size(250, 20)
        $clientLabel.AutoEllipsis = $true

        # Status label
        $statusLabel = New-Object System.Windows.Forms.Label
        $statusLabel.Text = "Initializing..."
        $statusLabel.Font = New-Object System.Drawing.Font('Segoe UI', 9)
        $statusLabel.Location = New-Object System.Drawing.Point(270, ($yPos + 15))
        $statusLabel.Size = New-Object System.Drawing.Size(200, 20)
        $statusLabel.ForeColor = [System.Drawing.Color]::Gray

        # Graph Auth button (disabled until worker script is ready)
        $graphAuthBtn = New-Object System.Windows.Forms.Button
        $graphAuthBtn.Text = "Graph Auth (Waiting...)"
        $graphAuthBtn.Location = New-Object System.Drawing.Point(480, ($yPos + 10))
        $graphAuthBtn.Size = New-Object System.Drawing.Size(120, 30)
        $graphAuthBtn.Enabled = $false  # Disabled until worker script is ready
        $graphAuthBtn.Tag = $ClientNumber

        # Exchange Online Auth button
        $exchangeAuthBtn = New-Object System.Windows.Forms.Button
        $exchangeAuthBtn.Text = "Exchange Online Auth"
        $exchangeAuthBtn.Location = New-Object System.Drawing.Point(610, ($yPos + 10))
        $exchangeAuthBtn.Size = New-Object System.Drawing.Size(150, 30)
        $exchangeAuthBtn.Enabled = $false
        $exchangeAuthBtn.Tag = $ClientNumber

        # Remove Tenant button
        $removeTenantBtn = New-Object System.Windows.Forms.Button
        $removeTenantBtn.Text = "Remove"
        $removeTenantBtn.Location = New-Object System.Drawing.Point(770, ($yPos + 10))
        $removeTenantBtn.Size = New-Object System.Drawing.Size(80, 30)
        $removeTenantBtn.Enabled = $true
        $removeTenantBtn.Tag = $ClientNumber
        $removeTenantBtn.ForeColor = [System.Drawing.Color]::DarkRed

        # Reset Auth button
        $resetAuthBtn = New-Object System.Windows.Forms.Button
        $resetAuthBtn.Text = "Reset Auth"
        $resetAuthBtn.Location = New-Object System.Drawing.Point(860, ($yPos + 10))
        $resetAuthBtn.Size = New-Object System.Drawing.Size(100, 30)
        $resetAuthBtn.Enabled = $true
        $resetAuthBtn.Tag = $ClientNumber
        $resetAuthBtn.ForeColor = [System.Drawing.Color]::DarkRed

        # Add controls to panel
        $script:authPanel.Controls.AddRange(@($clientLabel, $statusLabel, $graphAuthBtn, $exchangeAuthBtn, $removeTenantBtn, $resetAuthBtn))

        # Store controls and state
        $script:clientAuthStates[$ClientNumber] = @{
            GraphAuthenticated = $false
            ExchangeAuthenticated = $false
            GraphContext = $null
            TenantId = $null
            TenantName = $null
            Account = $null
        }
        $script:clientAuthControls[$ClientNumber] = @{
            ClientLabel = $clientLabel
            StatusLabel = $statusLabel
            GraphButton = $graphAuthBtn
            ExchangeButton = $exchangeAuthBtn
            RemoveButton = $removeTenantBtn
            ResetButton = $resetAuthBtn
        }
        
        # Update panel height to accommodate new row
        $newHeight = ($existingRows + 1) * ($clientRowHeight + $clientRowSpacing) + 20
        if ($newHeight -gt 410) {
            $script:authPanel.AutoScroll = $true
        }
        
        # Wire up button handlers
        $capturedClientNum = $ClientNumber
        
        # Graph Auth button handler
        $graphAuthBtn.add_Click({
            $clientNum = $this.Tag
            if (-not $clientNum) { $clientNum = $capturedClientNum }
            if ($script:authStatusTextBox) {
                $script:authStatusTextBox.AppendText("Sending Graph authentication command to Client $clientNum PowerShell session...`r`n")
                $script:authStatusTextBox.ScrollToCaret()
                [System.Windows.Forms.Application]::DoEvents()
            }
            $this.Enabled = $false
            $this.Text = "Sending Command..."
            
            if ($script:clientProcesses.ContainsKey($clientNum)) {
                $proc = $script:clientProcesses[$clientNum]
                try {
                    $procInfo = Get-Process -Id $proc.Id -ErrorAction SilentlyContinue
                    if (-not $procInfo -or $procInfo.HasExited) {
                        $script:authStatusTextBox.AppendText("ERROR: Client $clientNum PowerShell process is not running!`r`n")
                        $this.Enabled = $true
                        $this.Text = "Graph Auth"
                        return
                    }
                } catch {
                    $script:authStatusTextBox.AppendText("ERROR: Could not verify Client $clientNum PowerShell process!`r`n")
                    $this.Enabled = $true
                    $this.Text = "Graph Auth"
                    return
                }
            } else {
                $script:authStatusTextBox.AppendText("ERROR: Client $clientNum PowerShell process not found!`r`n")
                $this.Enabled = $true
                $this.Text = "Graph Auth"
                return
            }
            
            # Verify command directory exists
            if (-not (Test-Path $script:commandDir)) {
                $script:authStatusTextBox.AppendText("ERROR: Command directory does not exist: $script:commandDir`r`n")
                $this.Enabled = $true
                $this.Text = "Graph Auth"
                return
            }
            
            # Verify command file path
            $commandFile = Join-Path $script:commandDir "Client${clientNum}_Command.txt"
            $script:authStatusTextBox.AppendText("Client ${clientNum}: Command file will be: $commandFile`r`n")
            $script:authStatusTextBox.ScrollToCaret()
            [System.Windows.Forms.Application]::DoEvents()
            
            $response = Send-CommandToSession -ClientNumber $clientNum -Command "GRAPH_AUTH" -TimeoutSeconds 60
            
            # Check if Send-CommandToSession returned false (error writing command file)
            if ($response -eq $false) {
                $script:authStatusTextBox.AppendText("ERROR: Failed to send command to Client $clientNum. Check the status messages above.`r`n")
                $this.Enabled = $true
                $this.Text = "Graph Auth"
                return
            }
            
            # If response is null or empty, check the response file directly (might have been written after timeout)
            if (-not $response) {
                $responseFile = Join-Path $script:commandDir "Client${clientNum}_Response.txt"
                $script:authStatusTextBox.AppendText("Client ${clientNum}: No immediate response, checking response file: $responseFile`r`n")
                if (Test-Path $responseFile) {
                    try {
                        $response = Get-Content $responseFile -Raw -ErrorAction Stop | ForEach-Object { $_.Trim() }
                        $script:authStatusTextBox.AppendText("Client ${clientNum}: Read response from file: $response`r`n")
                        $script:authStatusTextBox.ScrollToCaret()
                        [System.Windows.Forms.Application]::DoEvents()
                    } catch {
                        $script:authStatusTextBox.AppendText("Client ${clientNum}: Could not read response file: $($_.Exception.Message)`r`n")
                    }
                } else {
                    $script:authStatusTextBox.AppendText("Client ${clientNum}: Response file does not exist. Checking if command file exists...`r`n")
                    if (Test-Path $commandFile) {
                        $cmdContent = Get-Content $commandFile -Raw -ErrorAction SilentlyContinue
                        $script:authStatusTextBox.AppendText("Client ${clientNum}: Command file still exists with content: '$cmdContent'`r`n")
                        $script:authStatusTextBox.AppendText("Client ${clientNum}: Worker script may not be polling. Check PowerShell window.`r`n")
                    } else {
                        $script:authStatusTextBox.AppendText("Client ${clientNum}: Command file was removed (worker script should have received it).`r`n")
                    }
                }
            }
            
            # If we got GRAPH_AUTH_STARTED, continue polling for the final result
            if ($response -eq "GRAPH_AUTH_STARTED") {
                $script:authStatusTextBox.AppendText("Client $clientNum Graph authentication started. Waiting for browser popup (may take 10-30 seconds)...`r`n")
                $script:authStatusTextBox.ScrollToCaret()
                [System.Windows.Forms.Application]::DoEvents()
                
                # Update status label
                $script:clientAuthControls[$clientNum].StatusLabel.Text = "Waiting for browser popup..."
                $script:clientAuthControls[$clientNum].StatusLabel.ForeColor = [System.Drawing.Color]::Orange
                
                # Continue polling the response file for the final result
                $responseFile = Join-Path $script:commandDir "Client${clientNum}_Response.txt"
                $startTime = Get-Date
                $finalResponse = $null
                $pollCount = 0
                
                while (((Get-Date) - $startTime).TotalSeconds -lt 300) {
                    $pollCount++
                    $elapsedSeconds = [int]((Get-Date) - $startTime).TotalSeconds
                    
                    # Update status every 10 seconds
                    if ($pollCount % 50 -eq 0) {
                        $statusMsg = "Waiting for browser popup... (${elapsedSeconds}s elapsed)"
                        $script:authStatusTextBox.AppendText("Client ${clientNum}: $statusMsg`r`n")
                        $script:authStatusTextBox.ScrollToCaret()
                        $script:clientAuthControls[$clientNum].StatusLabel.Text = $statusMsg
                        [System.Windows.Forms.Application]::DoEvents()
                    }
                    
                    if (Test-Path $responseFile) {
                        Start-Sleep -Milliseconds 200
                        try {
                            $finalResponse = Get-Content $responseFile -Raw -ErrorAction Stop | ForEach-Object { $_.Trim() }
                            # Check if we got a final response (not GRAPH_AUTH_STARTED)
                            if ($finalResponse -and $finalResponse -ne "GRAPH_AUTH_STARTED" -and $finalResponse -notmatch "^GRAPH_AUTH_STARTED") {
                                $script:authStatusTextBox.AppendText("Client ${clientNum}: Final response received: $finalResponse`r`n")
                                $script:authStatusTextBox.ScrollToCaret()
                                [System.Windows.Forms.Application]::DoEvents()
                                $response = $finalResponse
                                break
                            }
                        } catch {}
                    }
                    Start-Sleep -Milliseconds 200
                    [System.Windows.Forms.Application]::DoEvents()
                }
                
                if (-not $finalResponse -or $finalResponse -eq "GRAPH_AUTH_STARTED") {
                    $script:authStatusTextBox.AppendText("Client ${clientNum}: Timeout waiting for Graph authentication. The browser popup may not have appeared.`r`n")
                    $script:authStatusTextBox.ScrollToCaret()
                    $script:clientAuthControls[$clientNum].StatusLabel.Text = "Timeout - Use Reset Auth"
                    $script:clientAuthControls[$clientNum].StatusLabel.ForeColor = [System.Drawing.Color]::Red
                    $this.Enabled = $true
                    $this.Text = "Graph Auth"
                    [System.Windows.Forms.Application]::DoEvents()
                    return
                }
            }
            
            if ($response -like "GRAPH_AUTH_SUCCESS:*") {
                $tenantName = $response -replace "GRAPH_AUTH_SUCCESS:", ""
                $script:clientAuthStates[$clientNum].GraphAuthenticated = $true
                $script:clientAuthStates[$clientNum].TenantName = $tenantName
                $script:clientAuthControls[$clientNum].ClientLabel.Text = "Client $clientNum - $tenantName"
                $script:clientAuthControls[$clientNum].StatusLabel.Text = "Graph Auth Complete - Ready for Exchange"
                $script:clientAuthControls[$clientNum].StatusLabel.ForeColor = [System.Drawing.Color]::Orange
                $script:clientAuthControls[$clientNum].ExchangeButton.Enabled = $true
                $this.Text = "Graph Auth ✓"
                $script:authStatusTextBox.AppendText("Client $clientNum Graph authentication successful! Tenant: $tenantName`r`n")
                $script:authStatusTextBox.AppendText("Client $clientNum Exchange Online Auth button is now enabled. Click it to proceed.`r`n")
            } elseif ($response -like "GRAPH_AUTH_FAILED:*") {
                $errorMsg = $response -replace "GRAPH_AUTH_FAILED:", ""
                $this.Enabled = $true
                $this.Text = "Graph Auth"
                $script:authStatusTextBox.AppendText("Client $clientNum Graph authentication failed: $errorMsg`r`n")
            } else {
                $this.Enabled = $true
                $this.Text = "Graph Auth"
                $script:authStatusTextBox.AppendText("Client $clientNum Graph authentication failed or timeout. Response: $response`r`n")
                $script:authStatusTextBox.AppendText("Client $clientNum Check the PowerShell window for details.`r`n")
            }
            $script:authStatusTextBox.ScrollToCaret()
            [System.Windows.Forms.Application]::DoEvents()
        })
        
        # Exchange Auth button handler
        $exchangeAuthBtn.add_Click({
            $clientNum = $this.Tag
            if (-not $clientNum) { $clientNum = $capturedClientNum }
            if ($script:authStatusTextBox) {
                $script:authStatusTextBox.AppendText("Sending Exchange Online authentication command to Client $clientNum PowerShell session...`r`n")
                $script:authStatusTextBox.ScrollToCaret()
                [System.Windows.Forms.Application]::DoEvents()
            }
            $this.Enabled = $false
            $this.Text = "Sending Command..."
            
            $response = Send-CommandToSession -ClientNumber $clientNum -Command "EXCHANGE_AUTH" -TimeoutSeconds 30
            
            # If response is null or empty, check the response file directly
            if (-not $response) {
                $responseFile = Join-Path $script:commandDir "Client${clientNum}_Response.txt"
                if (Test-Path $responseFile) {
                    try {
                        $response = Get-Content $responseFile -Raw -ErrorAction Stop | ForEach-Object { $_.Trim() }
                        $script:authStatusTextBox.AppendText("Client ${clientNum}: Read Exchange auth response from file: $response`r`n")
                        $script:authStatusTextBox.ScrollToCaret()
                        [System.Windows.Forms.Application]::DoEvents()
                    } catch {
                        $script:authStatusTextBox.AppendText("Client ${clientNum}: Could not read response file: $($_.Exception.Message)`r`n")
                    }
                }
            }
            
            # If we got EXCHANGE_AUTH_STARTED, continue polling for the final result
            if ($response -eq "EXCHANGE_AUTH_STARTED") {
                $script:authStatusTextBox.AppendText("Client $clientNum Exchange Online authentication started. Waiting for browser popup...`r`n")
                $script:authStatusTextBox.ScrollToCaret()
                [System.Windows.Forms.Application]::DoEvents()
                
                # Update status label
                $script:clientAuthControls[$clientNum].StatusLabel.Text = "Waiting for browser popup..."
                $script:clientAuthControls[$clientNum].StatusLabel.ForeColor = [System.Drawing.Color]::Orange
                
                # Continue polling the response file for the final result
                $responseFile = Join-Path $script:commandDir "Client${clientNum}_Response.txt"
                $startTime = Get-Date
                $finalResponse = $null
                $pollCount = 0
                
                while (((Get-Date) - $startTime).TotalSeconds -lt 300) {
                    $pollCount++
                    $elapsedSeconds = [int]((Get-Date) - $startTime).TotalSeconds
                    
                    if ($pollCount % 50 -eq 0) {
                        $statusMsg = "Waiting for browser popup... (${elapsedSeconds}s elapsed)"
                        $script:authStatusTextBox.AppendText("Client ${clientNum}: $statusMsg`r`n")
                        $script:authStatusTextBox.ScrollToCaret()
                        $script:clientAuthControls[$clientNum].StatusLabel.Text = $statusMsg
                        [System.Windows.Forms.Application]::DoEvents()
                    }
                    
                    if (Test-Path $responseFile) {
                        Start-Sleep -Milliseconds 200
                        try {
                            $finalResponse = Get-Content $responseFile -Raw -ErrorAction Stop | ForEach-Object { $_.Trim() }
                            if ($finalResponse -and $finalResponse -ne "EXCHANGE_AUTH_STARTED" -and $finalResponse -notmatch "^EXCHANGE_AUTH_STARTED") {
                                $script:authStatusTextBox.AppendText("Client ${clientNum}: Final Exchange auth response: $finalResponse`r`n")
                                $script:authStatusTextBox.ScrollToCaret()
                                [System.Windows.Forms.Application]::DoEvents()
                                $response = $finalResponse
                                break
                            }
                        } catch {}
                    }
                    Start-Sleep -Milliseconds 200
                    [System.Windows.Forms.Application]::DoEvents()
                }
                
                if (-not $finalResponse -or $finalResponse -eq "EXCHANGE_AUTH_STARTED") {
                    $script:authStatusTextBox.AppendText("Client ${clientNum}: Timeout waiting for Exchange authentication.`r`n")
                    $script:authStatusTextBox.ScrollToCaret()
                    $script:clientAuthControls[$clientNum].StatusLabel.Text = "Timeout - Use Reset Auth"
                    $script:clientAuthControls[$clientNum].StatusLabel.ForeColor = [System.Drawing.Color]::Red
                    $this.Enabled = $true
                    $this.Text = "Exchange Online Auth"
                    [System.Windows.Forms.Application]::DoEvents()
                    return
                }
            }
            
            if ($response -like "EXCHANGE_AUTH_SUCCESS*") {
                $script:clientAuthStates[$clientNum].ExchangeAuthenticated = $true
                $script:clientAuthControls[$clientNum].StatusLabel.Text = "Exchange Auth Complete"
                $script:clientAuthControls[$clientNum].StatusLabel.ForeColor = [System.Drawing.Color]::Green
                $this.Text = "Exchange Auth ✓"
                $script:authStatusTextBox.AppendText("Client $clientNum Exchange Online authentication successful!`r`n")
                
                # Automatically trigger report generation
                $script:authStatusTextBox.AppendText("Starting report generation for Client $clientNum...`r`n")
                $script:authStatusTextBox.ScrollToCaret()
                [System.Windows.Forms.Application]::DoEvents()
                
                # Update status to show reports are starting
                $script:clientAuthControls[$clientNum].StatusLabel.Text = "Generating Reports..."
                $script:clientAuthControls[$clientNum].StatusLabel.ForeColor = [System.Drawing.Color]::Blue
                
                # Small delay before sending GENERATE_REPORTS
                Start-Sleep -Milliseconds 500
                
                $reportResponse = Send-CommandToSession -ClientNumber $clientNum -Command "GENERATE_REPORTS" -TimeoutSeconds 600
                if ($reportResponse -like "GENERATE_REPORTS_SUCCESS:*") {
                    $outputPath = $reportResponse -replace "GENERATE_REPORTS_SUCCESS:", ""
                    $script:clientAuthControls[$clientNum].StatusLabel.Text = "Complete"
                    $script:authStatusTextBox.AppendText("Client $clientNum report generation successful! Output: $outputPath`r`n")
                } elseif ($reportResponse -like "GENERATE_REPORTS_STARTED*") {
                    $script:authStatusTextBox.AppendText("Client $clientNum report generation started (running in background).`r`n")
                    $script:clientAuthControls[$clientNum].StatusLabel.Text = "Generating Reports..."
                } else {
                    $script:authStatusTextBox.AppendText("Client $clientNum report generation failed or timeout: $reportResponse`r`n")
                    $script:clientAuthControls[$clientNum].StatusLabel.Text = "Report Failed"
                    $script:clientAuthControls[$clientNum].StatusLabel.ForeColor = [System.Drawing.Color]::Red
                }
            } elseif ($response -like "EXCHANGE_AUTH_FAILED:*") {
                $errorMsg = $response -replace "EXCHANGE_AUTH_FAILED:", ""
                $this.Enabled = $true
                $this.Text = "Exchange Online Auth"
                $script:authStatusTextBox.AppendText("Client $clientNum Exchange Online authentication failed: $errorMsg`r`n")
            } else {
                $this.Enabled = $true
                $this.Text = "Exchange Online Auth"
                $script:authStatusTextBox.AppendText("Client $clientNum Exchange Online authentication failed or timeout. Response: $response`r`n")
                $script:authStatusTextBox.AppendText("Client $clientNum Check the PowerShell window for details.`r`n")
            }
            $script:authStatusTextBox.ScrollToCaret()
            [System.Windows.Forms.Application]::DoEvents()
        })
        
        # Reset Auth button handler
        $resetAuthBtn.add_Click({
            $clientNum = $this.Tag
            if (-not $clientNum) { $clientNum = $capturedClientNum }
            $script:authStatusTextBox.AppendText("Resetting authentication for Client $clientNum...`r`n")
            $script:authStatusTextBox.ScrollToCaret()
            [System.Windows.Forms.Application]::DoEvents()
            
            Send-CommandToSession -ClientNumber $clientNum -Command "CANCEL_AUTH" -TimeoutSeconds 30 | Out-Null
            
            $script:clientAuthStates[$clientNum].GraphAuthenticated = $false
            $script:clientAuthStates[$clientNum].ExchangeAuthenticated = $false
            $script:clientAuthStates[$clientNum].TenantName = $null
            $script:clientAuthControls[$clientNum].ClientLabel.Text = "Client $clientNum"
            $script:clientAuthControls[$clientNum].StatusLabel.Text = "Ready for Graph Auth"
            $script:clientAuthControls[$clientNum].StatusLabel.ForeColor = [System.Drawing.Color]::Blue
            $script:clientAuthControls[$clientNum].GraphButton.Enabled = $true
            $script:clientAuthControls[$clientNum].GraphButton.Text = "Graph Auth"
            $script:clientAuthControls[$clientNum].ExchangeButton.Enabled = $false
            $script:clientAuthControls[$clientNum].ExchangeButton.Text = "Exchange Online Auth"
            
            $script:authStatusTextBox.AppendText("Client $clientNum authentication reset complete.`r`n")
            $script:authStatusTextBox.ScrollToCaret()
            [System.Windows.Forms.Application]::DoEvents()
        })
        
        # Remove Tenant button handler
        $removeTenantBtn.add_Click({
            $clientNum = $this.Tag
            if (-not $clientNum) { $clientNum = $capturedClientNum }
            
            $result = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to remove Client $clientNum? This will close the PowerShell window and remove it from the list.", "Confirm Removal", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
            if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
                # Send exit command to PowerShell process
                if ($script:clientProcesses.ContainsKey($clientNum)) {
                    try {
                        Send-CommandToSession -ClientNumber $clientNum -Command "EXIT" -TimeoutSeconds 5 | Out-Null
                        Start-Sleep -Seconds 1
                        $proc = $script:clientProcesses[$clientNum]
                        if (-not $proc.HasExited) {
                            Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
                        }
                    } catch {}
                    $script:clientProcesses.Remove($clientNum)
                }
                
                # Remove controls from panel
                $controls = $script:clientAuthControls[$clientNum]
                $script:authPanel.Controls.Remove($controls.ClientLabel)
                $script:authPanel.Controls.Remove($controls.StatusLabel)
                $script:authPanel.Controls.Remove($controls.GraphButton)
                $script:authPanel.Controls.Remove($controls.ExchangeButton)
                $script:authPanel.Controls.Remove($controls.RemoveButton)
                $script:authPanel.Controls.Remove($controls.ResetButton)
                
                # Remove from state dictionaries
                $script:clientAuthStates.Remove($clientNum)
                $script:clientAuthControls.Remove($clientNum)
                
                # Recalculate positions for remaining rows
                $rowIndex = 0
                foreach ($key in ($script:clientAuthControls.Keys | Sort-Object)) {
                    $yPos = $rowIndex * ($clientRowHeight + $clientRowSpacing) + 10
                    $controls = $script:clientAuthControls[$key]
                    $controls.ClientLabel.Location = New-Object System.Drawing.Point(10, ($yPos + 15))
                    $controls.StatusLabel.Location = New-Object System.Drawing.Point(270, ($yPos + 15))
                    $controls.GraphButton.Location = New-Object System.Drawing.Point(480, ($yPos + 10))
                    $controls.ExchangeButton.Location = New-Object System.Drawing.Point(610, ($yPos + 10))
                    $controls.RemoveButton.Location = New-Object System.Drawing.Point(770, ($yPos + 10))
                    $controls.ResetButton.Location = New-Object System.Drawing.Point(860, ($yPos + 10))
                    $rowIndex++
                }
                
                $script:authStatusTextBox.AppendText("Client $clientNum removed.`r`n")
                $script:authStatusTextBox.ScrollToCaret()
                [System.Windows.Forms.Application]::DoEvents()
            }
        })
        
        return $true
    }

    # Add Tenant button click handler
    $addTenantBtn.add_Click({
        $newClientNum = $script:nextClientNumber
        if (Add-NewTenant -ClientNumber $newClientNum) {
            $script:nextClientNumber++
            $script:authStatusTextBox.AppendText("Added new tenant: Client $newClientNum`r`n")
            $script:authStatusTextBox.ScrollToCaret()
            [System.Windows.Forms.Application]::DoEvents()
        }
    })

    # Function to send command to PowerShell session and wait for response
    function Send-CommandToSession {
        param(
            [int]$ClientNumber,
            [string]$Command,
            [int]$TimeoutSeconds = 60
        )
        
        $commandFile = Join-Path $script:commandDir "Client${ClientNumber}_Command.txt"
        $responseFile = Join-Path $script:commandDir "Client${ClientNumber}_Response.txt"
        
        # Remove old response file if exists BEFORE writing command
        if (Test-Path $responseFile) {
            Write-Host "Send-CommandToSession: Removing old response file before sending command" -ForegroundColor Gray
            Remove-Item $responseFile -Force -ErrorAction SilentlyContinue
            Start-Sleep -Milliseconds 100  # Brief delay to ensure file is deleted
        }
        
        # Write command file
        Write-Host "Send-CommandToSession: Writing command file: $commandFile" -ForegroundColor Cyan
        Write-Host "Send-CommandToSession: Command to write: $Command" -ForegroundColor Cyan
        try {
            $Command | Out-File -FilePath $commandFile -Encoding UTF8 -Force
            Write-Host "Send-CommandToSession: Command file written successfully" -ForegroundColor Green
            
            # Verify file was written
            Start-Sleep -Milliseconds 100
            if (Test-Path $commandFile) {
                $fileContent = Get-Content $commandFile -Raw -ErrorAction SilentlyContinue
                Write-Host "Send-CommandToSession: Verified file exists, content: '$fileContent'" -ForegroundColor Gray
            } else {
                Write-Host "Send-CommandToSession: WARNING - File was written but doesn't exist!" -ForegroundColor Red
            }
            
            $script:authStatusTextBox.AppendText("Client ${ClientNumber}: Sent command '$Command'`r`n")
            $script:authStatusTextBox.AppendText("Client ${ClientNumber}: Command file: $commandFile`r`n")
            $script:authStatusTextBox.ScrollToCaret()
            [System.Windows.Forms.Application]::DoEvents()
        } catch {
            $errorMsg = "Failed to send command - $($_.Exception.Message)"
            Write-Host "Send-CommandToSession: ERROR - $errorMsg" -ForegroundColor Red
            $script:authStatusTextBox.AppendText("Client ${ClientNumber}: $errorMsg`r`n")
            return $false
        }
        
        # Wait for response
        Write-Host "Send-CommandToSession: Waiting for response file: $responseFile" -ForegroundColor Cyan
        $startTime = Get-Date
        $response = $null
        $pollCount = 0
        while (((Get-Date) - $startTime).TotalSeconds -lt $TimeoutSeconds) {
            $pollCount++
            if ($pollCount % 50 -eq 0) {
                Write-Host "Send-CommandToSession: Still waiting... ($pollCount polls, $(([int]((Get-Date) - $startTime).TotalSeconds))s elapsed)" -ForegroundColor Gray
            }
            
            if (Test-Path $responseFile) {
                Write-Host "Send-CommandToSession: Response file detected!" -ForegroundColor Yellow
                Start-Sleep -Milliseconds 200  # Brief delay to ensure file is fully written
                try {
                    $response = Get-Content $responseFile -Raw -ErrorAction Stop | ForEach-Object { $_.Trim() }
                    if ($response) {
                        Write-Host "Send-CommandToSession: Response received: $response" -ForegroundColor Green
                        return $response
                    } else {
                        Write-Host "Send-CommandToSession: Response file exists but is empty" -ForegroundColor Yellow
                    }
                } catch {
                    Write-Host "Send-CommandToSession: Error reading response file: $($_.Exception.Message)" -ForegroundColor Red
                }
            }
            Start-Sleep -Milliseconds 200
            [System.Windows.Forms.Application]::DoEvents()
        }
        
        Write-Host "Send-CommandToSession: Timeout waiting for response after $TimeoutSeconds seconds" -ForegroundColor Red
        $script:authStatusTextBox.AppendText("Client ${ClientNumber}: Timeout waiting for response to '$Command'`r`n")
        return $null
    }

    # Timer to periodically update status from status files
    $statusUpdateTimer = New-Object System.Windows.Forms.Timer
    $statusUpdateTimer.Interval = 2000  # Update every 2 seconds
    $statusUpdateTimer.add_Tick({
        foreach ($clientNum in $script:clientAuthControls.Keys) {
            $statusFile = Join-Path $script:tempDir "Client${clientNum}_Status.txt"
            if (Test-Path $statusFile) {
                try {
                    # Read last few lines of status file
                    $statusLines = Get-Content $statusFile -Tail 3 -ErrorAction SilentlyContinue
                    if ($statusLines) {
                        $latestStatus = $statusLines | Select-Object -Last 1
                        # Extract just the message part (after timestamp)
                        if ($latestStatus -match '\]\s+(.+)') {
                            $statusMessage = $matches[1]
                            $controls = $script:clientAuthControls[$clientNum]
                            if ($controls) {
                                # Check if worker script is ready and enable Graph Auth button if needed
                                # Wait for "Command polling loop started" to ensure the loop is actually running
                                if ($statusMessage -match 'Command polling loop started|Ready!.*Waiting for Graph Auth|Modules imported successfully') {
                                    if ($controls.GraphButton -and -not $controls.GraphButton.Enabled) {
                                        # Small delay to ensure polling loop has started
                                        Start-Sleep -Milliseconds 500
                                        $controls.GraphButton.Enabled = $true
                                        $controls.GraphButton.Text = "Graph Auth"
                                        if ($script:authStatusTextBox) {
                                            $script:authStatusTextBox.AppendText("Client $clientNum is ready for authentication (detected by status timer).`r`n")
                                            $script:authStatusTextBox.ScrollToCaret()
                                        }
                                    }
                                }
                                
                                # Only update if status has changed to avoid flickering
                                if ($controls.StatusLabel.Text -ne $statusMessage) {
                                    # Update status label with latest message
                                    $controls.StatusLabel.Text = $statusMessage
                                    
                                    # Color code based on status
                                    if ($statusMessage -match 'successful|complete|SUCCESS|authenticated') {
                                        $controls.StatusLabel.ForeColor = [System.Drawing.Color]::Green
                                    } elseif ($statusMessage -match 'error|failed|ERROR|FAILED') {
                                        $controls.StatusLabel.ForeColor = [System.Drawing.Color]::Red
                                    } elseif ($statusMessage -match 'generating|processing|running|starting') {
                                        $controls.StatusLabel.ForeColor = [System.Drawing.Color]::Blue
                                    } elseif ($statusMessage -match 'Ready!|Waiting for Graph Auth') {
                                        $controls.StatusLabel.ForeColor = [System.Drawing.Color]::Blue
                                    } elseif ($statusMessage -match 'waiting|polling') {
                                        $controls.StatusLabel.ForeColor = [System.Drawing.Color]::Gray
                                    }
                                }
                            }
                        }
                    }
                } catch {
                    # Silently ignore errors reading status file
                }
            }
        }
    })
    $statusUpdateTimer.Start()

    # Stop timer when form closes
    $authConsoleForm.add_FormClosed({
        $statusUpdateTimer.Stop()
        $statusUpdateTimer.Dispose()
    })

    # View Status Files button (for debugging)
    $viewStatusBtn = New-Object System.Windows.Forms.Button
    $viewStatusBtn.Text = "View Status Files"
    $viewStatusBtn.Location = New-Object System.Drawing.Point(15, 570)
    $viewStatusBtn.Size = New-Object System.Drawing.Size(150, 40)
    $viewStatusBtn.add_Click({
        if (Test-Path $script:tempDir) {
            Start-Process explorer.exe -ArgumentList $script:tempDir
        } else {
            [System.Windows.Forms.MessageBox]::Show("Temp directory not found: $script:tempDir", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
    })
    $authConsoleForm.Controls.Add($viewStatusBtn)

    # Show authentication console form
    $authConsoleForm.ShowDialog() | Out-Null
    
    # When authentication console closes, show the main form again
    $bulkForm.ShowDialog() | Out-Null
})

# Show the main form
[System.Windows.Forms.Application]::EnableVisualStyles()
$bulkForm.ShowDialog() | Out-Null

