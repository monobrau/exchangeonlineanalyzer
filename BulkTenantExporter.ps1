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

# Function to search and validate users from search terms
function Search-AndValidateUsers {
    param(
        [string]$SearchTerms,
        [object]$StatusLabel
    )
    
    if ([string]::IsNullOrWhiteSpace($SearchTerms)) {
        return @()
    }
    
    # Parse comma-separated search terms
    $searchTerms = $SearchTerms -split ',' | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    
    if ($searchTerms.Count -eq 0) {
        return @()
    }
    
    $allFoundUsers = @()
    
    # Check if Graph is connected
    try {
        $null = Get-MgContext -ErrorAction Stop
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Please connect to Microsoft Graph first to validate users.", "Not Connected", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return @()
    }
    
    if ($StatusLabel) {
        $StatusLabel.Text = "Searching for users..."
    }
    
    # Search for each term individually and combine results
    foreach ($searchTerm in $searchTerms) {
        Write-Host "Searching for users matching: '$searchTerm'"
        
        $users = @()
        try {
            # Try server-side filtering first (startsWith) - case-sensitive but we'll also try case variations
            # Microsoft Graph OData filters are case-sensitive, so try both original and lowercase
            $users1 = Get-MgUser -Filter "startsWith(DisplayName,'$searchTerm') or startsWith(UserPrincipalName,'$searchTerm')" -All -Property Id, UserPrincipalName, DisplayName -ErrorAction SilentlyContinue
            $searchTermLower = $searchTerm.ToLower()
            $searchTermUpper = $searchTerm.ToUpper()
            $searchTermTitle = (Get-Culture).TextInfo.ToTitleCase($searchTermLower)
            $users2 = Get-MgUser -Filter "startsWith(DisplayName,'$searchTermLower') or startsWith(UserPrincipalName,'$searchTermLower')" -All -Property Id, UserPrincipalName, DisplayName -ErrorAction SilentlyContinue
            $users3 = Get-MgUser -Filter "startsWith(DisplayName,'$searchTermUpper') or startsWith(UserPrincipalName,'$searchTermUpper')" -All -Property Id, UserPrincipalName, DisplayName -ErrorAction SilentlyContinue
            $users4 = Get-MgUser -Filter "startsWith(DisplayName,'$searchTermTitle') or startsWith(UserPrincipalName,'$searchTermTitle')" -All -Property Id, UserPrincipalName, DisplayName -ErrorAction SilentlyContinue
            $users = @($users1) + @($users2) + @($users3) + @($users4) | Sort-Object UserPrincipalName -Unique
            Write-Host "  Found $($users.Count) users with startsWith filter (tried multiple case variations)"
        } catch {
            Write-Host "  startsWith filter failed: $($_.Exception.Message), trying alternatives..."
        }
        
        if ($users.Count -eq 0) {
            # Try alternative search methods
            try {
                # Try exact match (case-sensitive first, then variations)
                $usersAlt1 = Get-MgUser -Filter "DisplayName eq '$searchTerm'" -All -Property Id, UserPrincipalName, DisplayName -ErrorAction SilentlyContinue
                $usersAlt1 += Get-MgUser -Filter "DisplayName eq '$searchTermLower'" -All -Property Id, UserPrincipalName, DisplayName -ErrorAction SilentlyContinue
                $usersAlt1 = $usersAlt1 | Sort-Object UserPrincipalName -Unique
                Write-Host "  Alternative search 1 (exact DisplayName match): Found $($usersAlt1.Count) users"
                
                $usersAlt2 = Get-MgUser -Filter "UserPrincipalName eq '$searchTerm'" -All -Property Id, UserPrincipalName, DisplayName -ErrorAction SilentlyContinue
                $usersAlt2 += Get-MgUser -Filter "UserPrincipalName eq '$searchTermLower'" -All -Property Id, UserPrincipalName, DisplayName -ErrorAction SilentlyContinue
                $usersAlt2 = $usersAlt2 | Sort-Object UserPrincipalName -Unique
                Write-Host "  Alternative search 2 (exact UserPrincipalName match): Found $($usersAlt2.Count) users"
                
                # Try case-insensitive search by getting all users and filtering client-side
                Write-Host "  Fetching all users for client-side filtering..."
                try {
                    $allUsers = Get-MgUser -All -Property Id, UserPrincipalName, DisplayName -ErrorAction Stop
                    Write-Host "  Retrieved $($allUsers.Count) total users from tenant"
                    
                    # Use case-insensitive matching with -ilike or -match
                    $searchTermPattern = "*$searchTerm*"
                    $usersAlt3 = $allUsers | Where-Object { 
                        ($_.DisplayName -and $_.DisplayName -ilike $searchTermPattern) -or 
                        ($_.UserPrincipalName -and $_.UserPrincipalName -ilike $searchTermPattern)
                    }
                    Write-Host "  Alternative search 3 (client-side filtering): Found $($usersAlt3.Count) users matching '$searchTerm'"
                    
                    # Show sample matches for debugging
                    if ($usersAlt3.Count -gt 0 -and $usersAlt3.Count -le 5) {
                        Write-Host "  Sample matches:" -ForegroundColor Gray
                        foreach ($u in $usersAlt3) {
                            Write-Host "    - $($u.DisplayName) ($($u.UserPrincipalName))" -ForegroundColor Gray
                        }
                    } elseif ($usersAlt3.Count -gt 5) {
                        Write-Host "  Sample matches (first 5):" -ForegroundColor Gray
                        foreach ($u in ($usersAlt3 | Select-Object -First 5)) {
                            Write-Host "    - $($u.DisplayName) ($($u.UserPrincipalName))" -ForegroundColor Gray
                        }
                    }
                } catch {
                    Write-Host "  Failed to retrieve all users for client-side filtering: $($_.Exception.Message)" -ForegroundColor Yellow
                    $usersAlt3 = @()
                }
                
                # Combine all results
                $users = @($usersAlt1) + @($usersAlt2) + @($usersAlt3) | Sort-Object UserPrincipalName -Unique
                Write-Host "  Combined alternative searches: Found $($users.Count) users"
            } catch {
                Write-Host "  Alternative searches also failed: $($_.Exception.Message)" -ForegroundColor Red
            }
        } else {
            # Show sample matches for startsWith results too
            if ($users.Count -gt 0 -and $users.Count -le 5) {
                Write-Host "  Sample matches:" -ForegroundColor Gray
                foreach ($u in $users) {
                    Write-Host "    - $($u.DisplayName) ($($u.UserPrincipalName))" -ForegroundColor Gray
                }
            } elseif ($users.Count -gt 5) {
                Write-Host "  Sample matches (first 5):" -ForegroundColor Gray
                foreach ($u in ($users | Select-Object -First 5)) {
                    Write-Host "    - $($u.DisplayName) ($($u.UserPrincipalName))" -ForegroundColor Gray
                }
            }
        }
        
        # Add found users to the collection (will deduplicate later)
        if ($users.Count -gt 0) {
            $allFoundUsers += $users
        }
    }
    
    # Remove duplicates based on UserPrincipalName
    $uniqueUsers = $allFoundUsers | Sort-Object UserPrincipalName -Unique
    
    Write-Host "Total unique users found: $($uniqueUsers.Count)"
    
    # Return array of UserPrincipalNames (strings)
    return $uniqueUsers | ForEach-Object { $_.UserPrincipalName }
}

# Import required modules
Write-Host "Loading required modules..." -ForegroundColor Cyan
Safe-ImportModule "$script:scriptRoot\Modules\ExportUtils.psm1"
Safe-ImportModule "$script:scriptRoot\Modules\GraphOnline.psm1"
Safe-ImportModule "$script:scriptRoot\Modules\BrowserIntegration.psm1"
Safe-ImportModule "$script:scriptRoot\Modules\Settings.psm1"
Write-Host "All modules loaded successfully." -ForegroundColor Green

# Load settings (shared with main application if it exists)
# Get-AppSettings will use custom location if configured, otherwise default location
$settings = $null
try {
    $settings = Get-AppSettings
    $actualSettingsPath = Get-SettingsPath
    Write-Host "Settings loaded from: $actualSettingsPath" -ForegroundColor Green
} catch {
    Write-Warning "Could not load settings: $($_.Exception.Message)"
    $settings = $null
}

# Initialize script-scope variables
$script:clientProcesses = @{}
$script:nextClientNumber = 1
$script:readinessCheckCount = @{}
$script:clientAuthStates = @{}
$script:clientAuthControls = @{}
$script:clientCacheDirs = @{}
$script:clientValidatedUsers = @{}  # Store validated UserPrincipalNames per tenant (keyed by ClientNumber)
$script:clientSearchTerms = @{}  # Store search terms per tenant when validation can't complete (keyed by ClientNumber)
$script:clientTickets = @{}  # Store ConnectWise ticket content per tenant (keyed by ClientNumber)
$script:clientReportFolders = @{}  # Store report output folder paths per tenant (keyed by ClientNumber)

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
$bulkConfigGroupBox.Size = New-Object System.Drawing.Size(400, 80)

# Days Back
$bulkDaysLabel = New-Object System.Windows.Forms.Label
$bulkDaysLabel.Text = "Days Back (Message Trace):"
$bulkDaysLabel.Location = New-Object System.Drawing.Point(20, 25)
$bulkDaysLabel.Size = New-Object System.Drawing.Size(150, 20)

$bulkDaysComboBox = New-Object System.Windows.Forms.ComboBox
$bulkDaysComboBox.Location = New-Object System.Drawing.Point(180, 23)
$bulkDaysComboBox.Size = New-Object System.Drawing.Size(100, 20)
$bulkDaysComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$bulkDaysComboBox.Items.AddRange(@("1", "3", "5", "7", "10", "14", "30", "45", "60", "90"))
$bulkDaysComboBox.SelectedIndex = 4  # Default to 10 days

$bulkConfigGroupBox.Controls.AddRange(@($bulkDaysLabel, $bulkDaysComboBox))

# Report Selection section
$bulkReportsGroupBox = New-Object System.Windows.Forms.GroupBox
$bulkReportsGroupBox.Text = "Select Reports to Export"
$bulkReportsGroupBox.Location = New-Object System.Drawing.Point(15, 280)
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
$bulkMessageTraceCheckBox.Text = "Message Trace"
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

$bulkMfaCoverageCheckBox = New-Object System.Windows.Forms.CheckBox
$bulkMfaCoverageCheckBox.Text = "MFA Coverage"
$bulkMfaCoverageCheckBox.Location = New-Object System.Drawing.Point(10, 265)
$bulkMfaCoverageCheckBox.Size = New-Object System.Drawing.Size(360, 20)
$bulkMfaCoverageCheckBox.Checked = $true

$bulkSignInLogsDaysLabel = New-Object System.Windows.Forms.Label
$bulkSignInLogsDaysLabel.Text = "Sign-In Logs Days:"
$bulkSignInLogsDaysLabel.Location = New-Object System.Drawing.Point(30, 290)
$bulkSignInLogsDaysLabel.Size = New-Object System.Drawing.Size(120, 20)

$bulkSignInLogsDaysComboBox = New-Object System.Windows.Forms.ComboBox
$bulkSignInLogsDaysComboBox.Location = New-Object System.Drawing.Point(160, 288)
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
    $bulkMfaCoverageCheckBox.Checked = $true
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
    $bulkMfaCoverageCheckBox.Checked = $false
})

# Add all controls to scrollable panel
$bulkReportsScrollPanel.Controls.AddRange(@(
    $bulkSelectAllBtn, $bulkDeselectAllBtn,
    $bulkMessageTraceCheckBox, $bulkInboxRulesCheckBox, $bulkTransportRulesCheckBox,
    $bulkMailFlowCheckBox, $bulkMailboxForwardingCheckBox, $bulkAuditLogsCheckBox,
    $bulkCaPoliciesCheckBox, $bulkAppRegistrationsCheckBox,
    $bulkSignInLogsCheckBox, $bulkMfaCoverageCheckBox,
    $bulkSignInLogsDaysLabel, $bulkSignInLogsDaysComboBox
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
    # Load Investigator Name and Company Name from settings
    try {
        $settings = Get-AppSettings
        $investigator = if ($settings -and $settings.InvestigatorName) { $settings.InvestigatorName } else { 'Security Administrator' }
        $company = if ($settings -and $settings.CompanyName) { $settings.CompanyName } else { 'Organization' }
    } catch {
        $investigator = 'Security Administrator'
        $company = 'Organization'
    }
    $days = [int]$bulkDaysComboBox.SelectedItem

    # Parse sign-in logs time range
    $signInLogsDays = 7
    $selectedRange = $bulkSignInLogsDaysComboBox.SelectedItem
    if ($selectedRange -eq "1 day") { $signInLogsDays = 1 }
    elseif ($selectedRange -eq "7 days") { $signInLogsDays = 7 }
    elseif ($selectedRange -eq "30 days") { $signInLogsDays = 30 }

    # Get report selections from checkboxes
    $days = [int]$bulkDaysComboBox.SelectedItem
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
        IncludeMfaCoverage = $bulkMfaCoverageCheckBox.Checked
        SignInLogsDaysBack = $signInLogsDays
        MessageTraceDaysBack = $days
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
    [string]`$CommandDir,
    [string[]]`$SelectedUsers = @()
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
    `$env:AZURE_IDENTITY_DISABLE_BROKER = "true"
    `$env:MSAL_DISABLE_BROKER = "1"
    `$env:MSAL_EXPERIMENTAL_DISABLE_BROKER = "1"
    Write-Status "Using isolated cache directory: `$cacheDir"
    Write-Host "Cache directory: `$cacheDir" -ForegroundColor Gray
    Write-Host ""
    
    # Import required modules
    Write-Status "Importing modules..."
    Write-Host "Importing modules..." -ForegroundColor Cyan
    Import-Module "`$ScriptRoot\Modules\ExportUtils.psm1" -Force -ErrorAction Stop
    Import-Module "`$ScriptRoot\Modules\GraphOnline.psm1" -Force -ErrorAction SilentlyContinue
    Import-Module "`$ScriptRoot\Modules\BrowserIntegration.psm1" -Force -ErrorAction SilentlyContinue
    # Import Settings module for memberberry integration and AI readme generation
    Import-Module "`$ScriptRoot\Modules\Settings.psm1" -Force -ErrorAction SilentlyContinue
    Write-Status "Modules imported successfully"
    Write-Host ""
    
    # CRITICAL: Disconnect any existing Graph session before starting
    # This ensures each tenant starts with a fresh authentication state
    # Even though each tenant has its own process, WAM might cache credentials globally
    try {
        `$existingContext = Get-MgContext -ErrorAction SilentlyContinue
        if (`$existingContext) {
            Write-Host "Found existing Graph session - disconnecting to ensure fresh authentication..." -ForegroundColor Yellow
            Disconnect-MgGraph -ErrorAction SilentlyContinue
            Start-Sleep -Milliseconds 500  # Brief pause to ensure disconnection completes
        }
    } catch {
        # Ignore errors - no session exists or module not loaded yet
    }
    
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
            IncludeMfaCoverage = if (`$null -ne `$jsonObj.IncludeMfaCoverage -and `$jsonObj.IncludeMfaCoverage -ne "") { [bool]`$jsonObj.IncludeMfaCoverage } else { `$false }
            SignInLogsDaysBack = if (`$null -ne `$jsonObj.SignInLogsDaysBack) { `$jsonObj.SignInLogsDaysBack } else { 7 }
            MessageTraceDaysBack = if (`$null -ne `$jsonObj.MessageTraceDaysBack) { `$jsonObj.MessageTraceDaysBack } else { 10 }
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
                
                # Clear any existing sessions and token caches
                # NOTE: Each tenant runs in its own isolated PowerShell process, so disconnecting only affects
                # this tenant's session, not other tenants' sessions running in separate processes.
                Write-Status "Clearing existing sessions and token caches..."
                Write-Host "Clearing existing sessions and token caches..." -ForegroundColor Cyan
                
                # Disconnect Graph session first (only if one exists in this process)
                # CRITICAL: This must happen BEFORE clearing cache to ensure session is fully released
                try { 
                    `$mgContext = Get-MgContext -ErrorAction SilentlyContinue
                    if (`$mgContext) {
                        Write-Host "Found existing Graph context - Tenant: `$(`$mgContext.TenantId), Account: `$(`$mgContext.Account)" -ForegroundColor Yellow
                        Disconnect-MgGraph -ErrorAction SilentlyContinue 
                        Write-Host "Disconnected existing Graph session for this tenant" -ForegroundColor Gray
                        # Wait a moment to ensure disconnection completes
                        Start-Sleep -Milliseconds 500
                    } else {
                        Write-Host "No existing Graph session to disconnect" -ForegroundColor Gray
                    }
                } catch {
                    # Ignore errors - session may not exist
                }
                
                # Clear Graph token cache and reset GraphSession singleton
                try {
                    `$graphSession = [Microsoft.Graph.PowerShell.Authentication.GraphSession]::Instance
                    if (`$graphSession) {
                        if (`$graphSession.AuthContext) {
                            `$graphSession.AuthContext.ClearTokenCache()
                            Write-Host "Cleared Graph token cache" -ForegroundColor Gray
                        }
                        # Try to reset the session instance to ensure fresh state
                        try {
                            `$graphSession.Reset() | Out-Null
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
                    if (`$env:MSAL_CACHE_DIR -and (Test-Path `$env:MSAL_CACHE_DIR)) {
                        `$allCacheFiles = Get-ChildItem -Path `$env:MSAL_CACHE_DIR -File -Recurse -ErrorAction SilentlyContinue
                        `$fileCount = `$allCacheFiles.Count
                        foreach (`$file in `$allCacheFiles) {
                            Remove-Item `$file.FullName -Force -ErrorAction SilentlyContinue
                        }
                        Write-Host "Cleared all files from MSAL cache directory (`$fileCount files)" -ForegroundColor Gray
                    }
                } catch {
                    # Ignore errors clearing MSAL cache
                }
                
                # Also clear IDENTITY_SERVICE_CACHE_DIR if it exists
                try {
                    if (`$env:IDENTITY_SERVICE_CACHE_DIR -and (Test-Path `$env:IDENTITY_SERVICE_CACHE_DIR)) {
                        `$allIdentityFiles = Get-ChildItem -Path `$env:IDENTITY_SERVICE_CACHE_DIR -File -Recurse -ErrorAction SilentlyContinue
                        `$identityFileCount = `$allIdentityFiles.Count
                        foreach (`$file in `$allIdentityFiles) {
                            Remove-Item `$file.FullName -Force -ErrorAction SilentlyContinue
                        }
                        Write-Host "Cleared all files from Identity cache directory (`$identityFileCount files)" -ForegroundColor Gray
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
                `$env:AZURE_IDENTITY_DISABLE_BROKER = "true"
                `$env:MSAL_DISABLE_BROKER = "1"
                `$env:MSAL_EXPERIMENTAL_DISABLE_BROKER = "1"

                # Ensure the per-client cache directory exists and is completely empty before authenticating
                # This is critical to prevent reusing tokens from previous tenants
                if (`$env:MSAL_CACHE_DIR) {
                    try {
                        if (-not (Test-Path `$env:MSAL_CACHE_DIR)) {
                            New-Item -ItemType Directory -Path `$env:MSAL_CACHE_DIR -Force -ErrorAction SilentlyContinue | Out-Null
                        } else {
                            # Remove ALL contents (files and subdirectories) to ensure fresh authentication
                            Get-ChildItem -Path `$env:MSAL_CACHE_DIR -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
                            Write-Host "Cleared MSAL cache directory contents before authentication" -ForegroundColor Gray
                        }
                    } catch {
                        # Ignore cache cleanup errors to avoid blocking auth
                    }
                }
                
                # Also clear IDENTITY_SERVICE_CACHE_DIR before authentication
                if (`$env:IDENTITY_SERVICE_CACHE_DIR) {
                    try {
                        if (-not (Test-Path `$env:IDENTITY_SERVICE_CACHE_DIR)) {
                            New-Item -ItemType Directory -Path `$env:IDENTITY_SERVICE_CACHE_DIR -Force -ErrorAction SilentlyContinue | Out-Null
                        } else {
                            Get-ChildItem -Path `$env:IDENTITY_SERVICE_CACHE_DIR -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
                            Write-Host "Cleared Identity cache directory contents before authentication" -ForegroundColor Gray
                        }
                    } catch {
                        # Ignore cache cleanup errors to avoid blocking auth
                    }
                }

                # Clear default MSAL cache location in user profile (in addition to custom cache dir)
                try {
                    `$defaultMsalCache = Join-Path `$env:LOCALAPPDATA ".IdentityService"
                    if (Test-Path `$defaultMsalCache) {
                        Get-ChildItem -Path `$defaultMsalCache -Recurse -File -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
                        Write-Host "Cleared default IdentityService cache in user profile" -ForegroundColor Gray
                    }
                } catch {
                    # Ignore errors
                }

                # Clear Microsoft.Graph module's own cache
                try {
                    `$graphModuleCache = Join-Path `$env:LOCALAPPDATA "Microsoft\Graph"
                    if (Test-Path `$graphModuleCache) {
                        Get-ChildItem -Path `$graphModuleCache -Recurse -File -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
                        Write-Host "Cleared Microsoft Graph module cache" -ForegroundColor Gray
                    }
                } catch {
                    # Ignore errors
                }

                # Clear Windows WAM (Web Account Manager) token cache
                # This helps prevent reusing cached credentials from previous sessions
                try {
                    `$wamCache = Join-Path `$env:LOCALAPPDATA "Packages\Microsoft.AAD.BrokerPlugin_cw5n1h2txyewy\AC\TokenBroker\Cache"
                    if (Test-Path `$wamCache) {
                        Get-ChildItem -Path `$wamCache -Recurse -File -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
                        Write-Host "Cleared WAM token broker cache" -ForegroundColor Gray
                    }
                } catch {
                    # Ignore errors - WAM cache may not exist or may be in use
                }

                # Also try alternative WAM cache location
                try {
                    `$wamCache2 = Join-Path `$env:LOCALAPPDATA "Packages\Microsoft.AAD.BrokerPlugin_cw5n1h2txyewy\LocalState"
                    if (Test-Path `$wamCache2) {
                        Get-ChildItem -Path `$wamCache2 -Recurse -File -Filter "*.dat" -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
                        Write-Host "Cleared WAM local state cache" -ForegroundColor Gray
                    }
                } catch {
                    # Ignore errors
                }

                `$scopes = @(
                    "AuditLog.Read.All",
                    "User.Read.All",
                    "Directory.Read.All",
                    "Policy.Read.All",
                    "Application.Read.All",
                    "Reports.Read.All"
                )

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
                    `$env:AZURE_IDENTITY_DISABLE_BROKER = "true"
                    `$env:MSAL_DISABLE_BROKER = "1"
                    `$env:MSAL_EXPERIMENTAL_DISABLE_BROKER = "1"
                    Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Yellow
                    Connect-MgGraph -Scopes `$scopes -ContextScope Process -NoWelcome -ErrorAction Stop
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

                    # Query all verified domains for the tenant
                    `$verifiedDomains = @()
                    try {
                        Write-Host "Querying tenant domains..." -ForegroundColor Gray
                        `$domainsResponse = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/domains" -ErrorAction Stop
                        if (`$domainsResponse -and `$domainsResponse.value) {
                            `$verifiedDomains = `$domainsResponse.value |
                                               Where-Object { `$_.isVerified -eq `$true } |
                                               ForEach-Object { `$_.id }
                            Write-Host "Found `$(`$verifiedDomains.Count) verified domain(s): `$(`$verifiedDomains -join ', ')" -ForegroundColor Cyan
                        }
                    } catch {
                        Write-Host "Warning: Failed to query tenant domains: `$(`$_.Exception.Message)" -ForegroundColor Yellow
                        Write-Host "Falling back to tenant name as primary domain" -ForegroundColor Yellow
                    }

                    # Build response with tenant name and domains
                    if (`$verifiedDomains -and `$verifiedDomains.Count -gt 0) {
                        `$domainsString = `$verifiedDomains -join ','
                        Write-CommandResponse "GRAPH_AUTH_SUCCESS:`$tenantDisplayName|DOMAINS:`$domainsString"
                    } else {
                        # Fallback: just return tenant name without domains
                        Write-CommandResponse "GRAPH_AUTH_SUCCESS:`$tenantDisplayName"
                    }
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
                
            } elseif (`$command -match "^VALIDATE_USERS") {
                if (-not `$graphAuthenticated) {
                    Write-Host "ERROR: Graph authentication must be completed first!" -ForegroundColor Red
                    Write-CommandResponse "VALIDATE_USERS_FAILED:Graph authentication not completed"
                    continue
                }
                
                Write-Host "==========================================" -ForegroundColor Yellow
                Write-Host "VALIDATE USERS COMMAND RECEIVED" -ForegroundColor Yellow
                Write-Host "==========================================" -ForegroundColor Yellow
                Write-Status "User validation command received"
                Write-CommandResponse "VALIDATE_USERS_STARTED"
                
                try {
                    # Parse search terms from command (format: VALIDATE_USERS|SEARCH_TERMS:term1,term2)
                    `$searchTerms = @()
                    if (`$command -match '\|SEARCH_TERMS:(.+)$') {
                        `$searchTermsJson = `$Matches[1]
                        try {
                            `$searchTermsArray = `$searchTermsJson | ConvertFrom-Json -ErrorAction Stop
                            if (`$searchTermsArray -is [array]) {
                                `$searchTerms = `$searchTermsArray
                            } elseif (`$searchTermsArray -is [string]) {
                                `$searchTerms = @(`$searchTermsArray)
                            } else {
                                `$searchTerms = @(`$searchTermsArray)
                            }
                        } catch {
                            # If JSON parsing fails, try splitting as comma-separated string
                            `$searchTerms = `$searchTermsJson -split ',' | ForEach-Object { `$_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace(`$_) }
                        }
                    } else {
                        Write-Warning "No search terms found in VALIDATE_USERS command"
                        Write-CommandResponse "VALIDATE_USERS_FAILED:No search terms provided"
                        continue
                    }
                    
                    Write-Host "Search terms received: `$(`$searchTerms -join ', ')" -ForegroundColor Cyan
                    Write-Status "Validating users for search terms: `$(`$searchTerms -join ', ')"
                    
                    # Perform user search using improved search logic
                    `$allFoundUsers = @()
                    foreach (`$searchTerm in `$searchTerms) {
                        Write-Host "  Searching for users matching: '`$searchTerm'" -ForegroundColor Gray
                        `$users = @()
                        try {
                            # Try server-side filtering first (startsWith) - try multiple case variations
                            `$users1 = Get-MgUser -Filter "startsWith(DisplayName,'`$searchTerm') or startsWith(UserPrincipalName,'`$searchTerm')" -All -Property Id, UserPrincipalName, DisplayName -ErrorAction SilentlyContinue
                            `$searchTermLower = `$searchTerm.ToLower()
                            `$searchTermUpper = `$searchTerm.ToUpper()
                            `$searchTermTitle = (Get-Culture).TextInfo.ToTitleCase(`$searchTermLower)
                            `$users2 = Get-MgUser -Filter "startsWith(DisplayName,'`$searchTermLower') or startsWith(UserPrincipalName,'`$searchTermLower')" -All -Property Id, UserPrincipalName, DisplayName -ErrorAction SilentlyContinue
                            `$users3 = Get-MgUser -Filter "startsWith(DisplayName,'`$searchTermUpper') or startsWith(UserPrincipalName,'`$searchTermUpper')" -All -Property Id, UserPrincipalName, DisplayName -ErrorAction SilentlyContinue
                            `$users4 = Get-MgUser -Filter "startsWith(DisplayName,'`$searchTermTitle') or startsWith(UserPrincipalName,'`$searchTermTitle')" -All -Property Id, UserPrincipalName, DisplayName -ErrorAction SilentlyContinue
                            `$users = @(`$users1) + @(`$users2) + @(`$users3) + @(`$users4) | Sort-Object UserPrincipalName -Unique
                            Write-Host "    Found `$(`$users.Count) users with startsWith filter (tried multiple case variations)" -ForegroundColor Gray
                        } catch {
                            Write-Host "    startsWith filter failed: `$(`$_.Exception.Message), trying alternatives..." -ForegroundColor Yellow
                        }
                        
                        if (`$users.Count -eq 0) {
                            # Try alternative search methods
                            try {
                                # Try exact match (case-sensitive first, then variations)
                                `$usersAlt1 = Get-MgUser -Filter "DisplayName eq '`$searchTerm'" -All -Property Id, UserPrincipalName, DisplayName -ErrorAction SilentlyContinue
                                `$usersAlt1 += Get-MgUser -Filter "DisplayName eq '`$searchTermLower'" -All -Property Id, UserPrincipalName, DisplayName -ErrorAction SilentlyContinue
                                `$usersAlt1 = `$usersAlt1 | Sort-Object UserPrincipalName -Unique
                                
                                `$usersAlt2 = Get-MgUser -Filter "UserPrincipalName eq '`$searchTerm'" -All -Property Id, UserPrincipalName, DisplayName -ErrorAction SilentlyContinue
                                `$usersAlt2 += Get-MgUser -Filter "UserPrincipalName eq '`$searchTermLower'" -All -Property Id, UserPrincipalName, DisplayName -ErrorAction SilentlyContinue
                                `$usersAlt2 = `$usersAlt2 | Sort-Object UserPrincipalName -Unique
                                
                                # Try case-insensitive search by getting all users and filtering client-side
                                Write-Host "    Fetching all users for client-side filtering..." -ForegroundColor Gray
                                try {
                                    `$allUsers = Get-MgUser -All -Property Id, UserPrincipalName, DisplayName -ErrorAction Stop
                                    Write-Host "    Retrieved `$(`$allUsers.Count) total users from tenant" -ForegroundColor Gray
                                    
                                    # Use case-insensitive matching with -ilike
                                    `$searchTermPattern = "*`$searchTerm*"
                                    `$usersAlt3 = `$allUsers | Where-Object { 
                                        (`$_.DisplayName -and `$_.DisplayName -ilike `$searchTermPattern) -or 
                                        (`$_.UserPrincipalName -and `$_.UserPrincipalName -ilike `$searchTermPattern)
                                    }
                                    Write-Host "    Client-side filtering: Found `$(`$usersAlt3.Count) users matching '`$searchTerm'" -ForegroundColor Gray
                                } catch {
                                    Write-Warning "Failed to retrieve all users for client-side filtering: `$(`$_.Exception.Message)"
                                    `$usersAlt3 = @()
                                }
                                
                                # Combine all results
                                `$users = @(`$usersAlt1) + @(`$usersAlt2) + @(`$usersAlt3) | Sort-Object UserPrincipalName -Unique
                                Write-Host "    Combined alternative searches: Found `$(`$users.Count) users" -ForegroundColor Gray
                            } catch {
                                Write-Warning "Could not search for users matching '`$searchTerm': `$(`$_.Exception.Message)"
                            }
                        }
                        if (`$users.Count -gt 0) {
                            `$allFoundUsers += `$users
                        }
                    }
                    
                    # Get unique UserPrincipalNames
                    `$validatedUsers = (`$allFoundUsers | Sort-Object UserPrincipalName -Unique | ForEach-Object { `$_.UserPrincipalName })
                    
                    if (`$validatedUsers.Count -gt 0) {
                        Write-Host "Validation successful: Found `$(`$validatedUsers.Count) user(s)" -ForegroundColor Green
                        Write-Status "Validation successful: Found `$(`$validatedUsers.Count) user(s)"
                        `$responseJson = @{
                            Success = `$true
                            UserCount = `$validatedUsers.Count
                            Users = `$validatedUsers
                        } | ConvertTo-Json -Compress
                        Write-CommandResponse "VALIDATE_USERS_SUCCESS:`$responseJson"
                    } else {
                        Write-Host "Validation completed: No users found matching search terms" -ForegroundColor Yellow
                        Write-Status "Validation completed: No users found matching search terms"
                        `$responseJson = @{
                            Success = `$false
                            UserCount = 0
                            Users = @()
                            Message = "No users found matching the search terms"
                        } | ConvertTo-Json -Compress
                        Write-CommandResponse "VALIDATE_USERS_SUCCESS:`$responseJson"
                    }
                } catch {
                    Write-Host "ERROR: User validation failed - `$(`$_.Exception.Message)" -ForegroundColor Red
                    Write-Status "ERROR: User validation failed - `$(`$_.Exception.Message)"
                    Write-CommandResponse "VALIDATE_USERS_FAILED:`$(`$_.Exception.Message)"
                }
                
                Write-Host ""
                Write-Host "Waiting for next command from GUI..." -ForegroundColor Green
                Write-Host ""
                
            } elseif (`$command -match "^GENERATE_REPORTS") {
                if (-not `$graphAuthenticated -or -not `$exchangeAuthenticated) {
                    Write-Host "ERROR: Both Graph and Exchange authentication must be completed first!" -ForegroundColor Red
                    Write-CommandResponse "GENERATE_REPORTS_FAILED:Authentication not completed"
                    continue
                }
                
                # Parse SelectedUsers and TicketData from command if provided
                `$selectedUsersForReport = @()
                `$ticketNumbers = @()
                `$ticketContent = ''
                
                # Parse ticket data from command (format: |TICKET_DATA:{"TicketNumbers":["12345"],"TicketContent":"..."})
                # Use a more robust regex that captures everything after TICKET_DATA: until end of string
                # This handles cases where ticket content might contain special characters
                Write-Host "Parsing ticket data from command. Command length: `$(`$command.Length)" -ForegroundColor Gray
                Write-Host "Command preview (first 500 chars): `$(`$command.Substring(0, [Math]::Min(500, `$command.Length)))" -ForegroundColor Gray
                if (`$command -match '\|TICKET_DATA:(.+)$') {
                    Write-Host "TICKET_DATA regex matched!" -ForegroundColor Green
                    try {
                        `$ticketDataJson = `$Matches[1]
                        Write-Host "Ticket data JSON extracted (length: `$(`$ticketDataJson.Length))" -ForegroundColor Gray
                        Write-Host "Ticket data JSON preview (first 300 chars): `$(`$ticketDataJson.Substring(0, [Math]::Min(300, `$ticketDataJson.Length)))" -ForegroundColor Gray
                        `$ticketData = `$ticketDataJson | ConvertFrom-Json -ErrorAction Stop
                        Write-Host "Ticket data JSON parsed successfully" -ForegroundColor Green
                        if (`$ticketData.TicketNumbers) {
                            Write-Host "TicketNumbers property found: `$(`$ticketData.TicketNumbers)" -ForegroundColor Gray
                            # Ensure TicketNumbers is always an array
                            if (`$ticketData.TicketNumbers -is [string]) {
                                `$ticketNumbers = @(`$ticketData.TicketNumbers)
                                Write-Host "TicketNumbers was string, converted to array: `$ticketNumbers" -ForegroundColor Gray
                            } elseif (`$ticketData.TicketNumbers -is [array]) {
                                `$ticketNumbers = `$ticketData.TicketNumbers
                                Write-Host "TicketNumbers was array: `$ticketNumbers" -ForegroundColor Gray
                            } else {
                                `$ticketNumbers = @(`$ticketData.TicketNumbers)
                                Write-Host "TicketNumbers was other type, converted to array: `$ticketNumbers" -ForegroundColor Gray
                            }
                        } else {
                            Write-Host "TicketNumbers property not found in parsed data" -ForegroundColor Yellow
                        }
                        if (`$ticketData.TicketContent) {
                            `$ticketContent = `$ticketData.TicketContent
                            Write-Host "TicketContent property found (length: `$(`$ticketContent.Length))" -ForegroundColor Gray
                        } else {
                            Write-Host "TicketContent property not found in parsed data" -ForegroundColor Yellow
                        }
                        Write-Host "Ticket data parsed: TicketNumbers=`$(`$ticketNumbers.Count) (`$(`$ticketNumbers -join ', ')), TicketContent length=`$(`$ticketContent.Length)" -ForegroundColor Cyan
                        Write-Host "Ticket data found: `$(`$ticketNumbers.Count) ticket number(s): `$(`$ticketNumbers -join ', ')" -ForegroundColor Cyan
                        Write-Status "Ticket data found: `$(`$ticketNumbers.Count) ticket number(s): `$(`$ticketNumbers -join ', ')"
                    } catch {
                        Write-Warning "Could not parse ticket data from command: `$(`$_.Exception.Message)"
                        Write-Host "Ticket data JSON that failed to parse: `$ticketDataJson" -ForegroundColor Yellow
                        Write-Host "Full command was: `$command" -ForegroundColor Yellow
                        Write-Host "Exception details: `$(`$_.Exception | Out-String)" -ForegroundColor Red
                    }
                } else {
                    Write-Host "No TICKET_DATA found in command. Command preview: `$(`$command.Substring(0, [Math]::Min(500, `$command.Length)))" -ForegroundColor Yellow
                    Write-Host "Checking if command contains 'TICKET_DATA': `$(`$command.Contains('TICKET_DATA'))" -ForegroundColor Yellow
                }
                
                # Check if this is a search terms command (GENERATE_REPORTS_SEARCH:["term1","term2"])
                # Extract search terms before TICKET_DATA if present
                if (`$command -match "^GENERATE_REPORTS_SEARCH:(.+?)(?:\|TICKET_DATA:|$)") {
                    try {
                        `$searchTermsJson = `$Matches[1]
                        `$searchTermsParsed = `$searchTermsJson | ConvertFrom-Json -ErrorAction Stop
                        # Ensure searchTerms is always an array (ConvertFrom-Json might return a string for single values)
                        if (`$searchTermsParsed -is [string]) {
                            `$searchTerms = @(`$searchTermsParsed)
                        } elseif (`$searchTermsParsed -is [array]) {
                            `$searchTerms = `$searchTermsParsed
                        } else {
                            `$searchTerms = @(`$searchTermsParsed)
                        }
                        Write-Host "User filtering enabled with search terms. Validating users..." -ForegroundColor Cyan
                        Write-Status "User filtering enabled with search terms. Validating users..."
                        
                        # Validate search terms using Graph API
                        `$allFoundUsers = @()
                        foreach (`$searchTerm in `$searchTerms) {
                            Write-Host "  Searching for users matching: '`$searchTerm'" -ForegroundColor Gray
                            `$users = @()
                            try {
                                # Try server-side filtering first (startsWith) - try multiple case variations
                                `$users1 = Get-MgUser -Filter "startsWith(DisplayName,'`$searchTerm') or startsWith(UserPrincipalName,'`$searchTerm')" -All -Property Id, UserPrincipalName, DisplayName -ErrorAction SilentlyContinue
                                `$searchTermLower = `$searchTerm.ToLower()
                                `$searchTermUpper = `$searchTerm.ToUpper()
                                `$searchTermTitle = (Get-Culture).TextInfo.ToTitleCase(`$searchTermLower)
                                `$users2 = Get-MgUser -Filter "startsWith(DisplayName,'`$searchTermLower') or startsWith(UserPrincipalName,'`$searchTermLower')" -All -Property Id, UserPrincipalName, DisplayName -ErrorAction SilentlyContinue
                                `$users3 = Get-MgUser -Filter "startsWith(DisplayName,'`$searchTermUpper') or startsWith(UserPrincipalName,'`$searchTermUpper')" -All -Property Id, UserPrincipalName, DisplayName -ErrorAction SilentlyContinue
                                `$users4 = Get-MgUser -Filter "startsWith(DisplayName,'`$searchTermTitle') or startsWith(UserPrincipalName,'`$searchTermTitle')" -All -Property Id, UserPrincipalName, DisplayName -ErrorAction SilentlyContinue
                                `$users = @(`$users1) + @(`$users2) + @(`$users3) + @(`$users4) | Sort-Object UserPrincipalName -Unique
                                Write-Host "    Found `$(`$users.Count) users with startsWith filter (tried multiple case variations)" -ForegroundColor Gray
                            } catch {
                                Write-Host "    startsWith filter failed: `$(`$_.Exception.Message), trying alternatives..." -ForegroundColor Yellow
                            }
                            
                            if (`$users.Count -eq 0) {
                                # Try alternative search methods
                                try {
                                    # Try exact match (case-sensitive first, then variations)
                                    `$usersAlt1 = Get-MgUser -Filter "DisplayName eq '`$searchTerm'" -All -Property Id, UserPrincipalName, DisplayName -ErrorAction SilentlyContinue
                                    `$usersAlt1 += Get-MgUser -Filter "DisplayName eq '`$searchTermLower'" -All -Property Id, UserPrincipalName, DisplayName -ErrorAction SilentlyContinue
                                    `$usersAlt1 = `$usersAlt1 | Sort-Object UserPrincipalName -Unique
                                    
                                    `$usersAlt2 = Get-MgUser -Filter "UserPrincipalName eq '`$searchTerm'" -All -Property Id, UserPrincipalName, DisplayName -ErrorAction SilentlyContinue
                                    `$usersAlt2 += Get-MgUser -Filter "UserPrincipalName eq '`$searchTermLower'" -All -Property Id, UserPrincipalName, DisplayName -ErrorAction SilentlyContinue
                                    `$usersAlt2 = `$usersAlt2 | Sort-Object UserPrincipalName -Unique
                                    
                                    # Try case-insensitive search by getting all users and filtering client-side
                                    Write-Host "    Fetching all users for client-side filtering..." -ForegroundColor Gray
                                    try {
                                        `$allUsers = Get-MgUser -All -Property Id, UserPrincipalName, DisplayName -ErrorAction Stop
                                        Write-Host "    Retrieved `$(`$allUsers.Count) total users from tenant" -ForegroundColor Gray
                                        
                                        # Use case-insensitive matching with -ilike
                                        `$searchTermPattern = "*`$searchTerm*"
                                        `$usersAlt3 = `$allUsers | Where-Object { 
                                            (`$_.DisplayName -and `$_.DisplayName -ilike `$searchTermPattern) -or 
                                            (`$_.UserPrincipalName -and `$_.UserPrincipalName -ilike `$searchTermPattern)
                                        }
                                        Write-Host "    Client-side filtering: Found `$(`$usersAlt3.Count) users matching '`$searchTerm'" -ForegroundColor Gray
                                    } catch {
                                        Write-Warning "Failed to retrieve all users for client-side filtering: `$(`$_.Exception.Message)"
                                        `$usersAlt3 = @()
                                    }
                                    
                                    # Combine all results
                                    `$users = @(`$usersAlt1) + @(`$usersAlt2) + @(`$usersAlt3) | Sort-Object UserPrincipalName -Unique
                                    Write-Host "    Combined alternative searches: Found `$(`$users.Count) users" -ForegroundColor Gray
                                } catch {
                                    Write-Warning "Could not search for users matching '`$searchTerm': `$(`$_.Exception.Message)"
                                }
                            }
                            if (`$users.Count -gt 0) {
                                `$allFoundUsers += `$users
                            }
                        }
                        
                        # Get unique UserPrincipalNames
                        `$selectedUsersForReport = (`$allFoundUsers | Sort-Object UserPrincipalName -Unique | ForEach-Object { `$_.UserPrincipalName })
                        Write-Host "User filtering enabled: Found `$(`$selectedUsersForReport.Count) user(s) from search terms" -ForegroundColor Cyan
                        Write-Status "User filtering enabled: Found `$(`$selectedUsersForReport.Count) user(s) from search terms"
                        
                        # Warn if search terms were provided but no users found
                        if (`$selectedUsersForReport.Count -eq 0) {
                            Write-Warning "No users found matching the search terms. Report will be generated without user filtering."
                            Write-Status "WARNING: No users found matching search terms - generating report without filtering"
                        }
                    } catch {
                        Write-Warning "Could not parse or validate search terms from command: `$(`$_.Exception.Message)"
                        Write-Status "ERROR: Failed to validate search terms - `$(`$_.Exception.Message)"
                        # Set to empty array so report continues without filtering
                        `$selectedUsersForReport = @()
                    }
                }
                # Check if this is a direct users command (GENERATE_REPORTS|SelectedUsers:["user1","user2"])
                elseif (`$command -match '\|SelectedUsers:(.+?)(?:\||$)') {
                    try {
                        `$usersJson = `$Matches[1]
                        `$selectedUsersForReport = `$usersJson | ConvertFrom-Json -ErrorAction Stop
                        Write-Host "User filtering enabled: `$(`$selectedUsersForReport.Count) user(s) selected" -ForegroundColor Cyan
                        Write-Status "User filtering enabled: `$(`$selectedUsersForReport.Count) user(s)"
                    } catch {
                        Write-Warning "Could not parse SelectedUsers from command: `$(`$_.Exception.Message)"
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
                if (`$ticketContent -and -not [string]::IsNullOrWhiteSpace(`$ticketContent)) {
                    try {
                        Import-Module "`$ScriptRoot\Modules\Settings.psm1" -Force -ErrorAction SilentlyContinue
                        if (Get-Command Filter-TicketContent -ErrorAction SilentlyContinue) {
                            `$originalLength = `$ticketContent.Length
                            `$ticketContent = Filter-TicketContent -TicketContent `$ticketContent
                            Write-Host "Ticket content filtered: `$originalLength -> `$(`$ticketContent.Length) characters" -ForegroundColor Gray
                        } else {
                            Write-Warning "Filter-TicketContent function not found, using raw ticket content"
                        }
                    } catch {
                        Write-Warning "Failed to filter ticket content: `$(`$_.Exception.Message). Using raw content."
                    }
                }
                
                Write-Host "Ticket data being passed: TicketNumbers=`$(`$ticketNumbers.Count) (`$(`$ticketNumbers -join ', ')), TicketContent length=`$(`$ticketContent.Length)" -ForegroundColor Cyan
                try {
                    `$messageTraceDays = if (`$reportSelections.MessageTraceDaysBack) { `$reportSelections.MessageTraceDaysBack } else { `$DaysBack }
                    `$report = New-SecurityInvestigationReport -InvestigatorName `$InvestigatorName -CompanyName `$CompanyName -DaysBack `$DaysBack -StatusLabel `$null -MainForm `$null -IncludeMessageTrace `$reportSelections.IncludeMessageTrace -IncludeInboxRules `$reportSelections.IncludeInboxRules -IncludeTransportRules `$reportSelections.IncludeTransportRules -IncludeMailFlowConnectors `$reportSelections.IncludeMailFlowConnectors -IncludeMailboxForwarding `$reportSelections.IncludeMailboxForwarding -IncludeAuditLogs `$reportSelections.IncludeAuditLogs -IncludeConditionalAccessPolicies `$reportSelections.IncludeConditionalAccessPolicies -IncludeAppRegistrations `$reportSelections.IncludeAppRegistrations -IncludeSignInLogs `$reportSelections.IncludeSignInLogs -IncludeMfaCoverage `$reportSelections.IncludeMfaCoverage -SignInLogsDaysBack `$reportSelections.SignInLogsDaysBack -MessageTraceDaysBack `$messageTraceDays -SelectedUsers `$selectedUsersForReport -TicketNumbers `$ticketNumbers -TicketContent `$ticketContent
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
                
                # Clear authentication context and token cache more thoroughly
                try {
                    `$graphSession = [Microsoft.Graph.PowerShell.Authentication.GraphSession]::Instance
                    if (`$graphSession -and `$graphSession.AuthContext) {
                        `$graphSession.AuthContext.ClearTokenCache()
                        Write-Host "Cleared Graph token cache" -ForegroundColor Cyan
                    }
                } catch {
                    # Ignore errors clearing token cache
                }
                
                # Also try to clear any MSAL cache
                try {
                    `$msalCache = [Microsoft.Identity.Client.TokenCacheHelper]::GetCacheFilePath()
                    if (`$msalCache -and (Test-Path `$msalCache)) {
                        Remove-Item `$msalCache -Force -ErrorAction SilentlyContinue
                        Write-Host "Cleared MSAL token cache" -ForegroundColor Cyan
                    }
                } catch {
                    # Ignore errors clearing MSAL cache - method may not be available
                }
                
                # Clear Exchange Online token cache
                try {
                    `$exoSession = Get-PSSession | Where-Object { `$_.ConfigurationName -eq "Microsoft.Exchange" }
                    if (`$exoSession) {
                        Remove-PSSession `$exoSession -ErrorAction SilentlyContinue
                        Write-Host "Cleared Exchange Online sessions" -ForegroundColor Cyan
                    }
                } catch {
                    # Ignore errors clearing Exchange sessions
                }
    
                Write-Status "Authentication cancelled and reset"
                Write-Host "Authentication cancelled and reset. All token caches cleared. Ready for new authentication attempt." -ForegroundColor Green
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

    # Expand All button
    $expandAllBtn = New-Object System.Windows.Forms.Button
    $expandAllBtn.Text = "Expand All"
    $expandAllBtn.Font = New-Object System.Drawing.Font('Segoe UI', 9)
    $expandAllBtn.Location = New-Object System.Drawing.Point(175, 100)
    $expandAllBtn.Size = New-Object System.Drawing.Size(100, 35)
    $expandAllBtn.BackColor = [System.Drawing.Color]::FromArgb(33, 150, 243)
    $expandAllBtn.ForeColor = [System.Drawing.Color]::White

    # Collapse All button
    $collapseAllBtn = New-Object System.Windows.Forms.Button
    $collapseAllBtn.Text = "Collapse All"
    $collapseAllBtn.Font = New-Object System.Drawing.Font('Segoe UI', 9)
    $collapseAllBtn.Location = New-Object System.Drawing.Point(285, 100)
    $collapseAllBtn.Size = New-Object System.Drawing.Size(100, 35)
    $collapseAllBtn.BackColor = [System.Drawing.Color]::FromArgb(156, 39, 176)
    $collapseAllBtn.ForeColor = [System.Drawing.Color]::White

    # Create Panel for client authentication rows
    $authPanel = New-Object System.Windows.Forms.Panel
    $authPanel.Location = New-Object System.Drawing.Point(15, 145)
    $authPanel.Size = New-Object System.Drawing.Size(970, 420)
    $authPanel.AutoScroll = $true
    $authPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle

    # Store client authentication state and controls
    $script:clientAuthStates = @{}
    $script:clientAuthControls = @{}
    $script:clientCacheDirs = @{}
    $clientRowHeight = 200  # Increased to accommodate all controls including ticket textbox (80px) and view reports button
    $clientRowSpacing = 10  # Increased spacing between rows

    # Add controls to form
    $authConsoleForm.Controls.AddRange(@($authTitleLabel, $authInstructionsLabel, $addTenantBtn, $expandAllBtn, $collapseAllBtn, $authPanel))

    # Close button
    $authCloseBtn = New-Object System.Windows.Forms.Button
    $authCloseBtn.Text = "Close"
    $authCloseBtn.Location = New-Object System.Drawing.Point(880, 570)
    $authCloseBtn.Size = New-Object System.Drawing.Size(100, 40)
    $authCloseBtn.add_Click({
        # Stop the status update timer first to prevent it from accessing disposed controls
        try {
            if ($statusUpdateTimer -and $statusUpdateTimer.Enabled) {
                $statusUpdateTimer.Stop()
            }
        } catch {}
        
        # Send exit command to all active PowerShell processes
        foreach ($clientNum in $script:clientProcesses.Keys) {
            try {
                Send-CommandToSession -ClientNumber $clientNum -Command "EXIT" -TimeoutSeconds 5 | Out-Null
                Start-Sleep -Milliseconds 500
                $proc = $script:clientProcesses[$clientNum]
                if ($proc -and -not $proc.HasExited) {
                    Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
                }
            } catch {}
        }
        
        # Close the form using DialogResult to properly close modal dialog
        try {
            $authConsoleForm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        } catch {
            # Fallback to Close() if DialogResult fails
            try {
                $authConsoleForm.Close()
            } catch {}
        }
    })
    $authConsoleForm.Controls.Add($authCloseBtn)

    # Status text box
    $authStatusTextBox = New-Object System.Windows.Forms.TextBox
    $authStatusTextBox.Multiline = $true
    $authStatusTextBox.ReadOnly = $true
    $authStatusTextBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
    $authStatusTextBox.Location = New-Object System.Drawing.Point(15, 610)
    $authStatusTextBox.Size = New-Object System.Drawing.Size(985, 80)
    $authStatusTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Top
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

    # Function to update tenant positions after minimize/expand
    function Update-TenantPositions {
        $clientRowSpacing = 10
        $minimizedHeight = 50
        $expandedHeight = 200
        $currentY = 10

        # Sort client numbers to maintain order
        $sortedClientNums = $script:clientAuthControls.Keys | Sort-Object

        foreach ($clientNum in $sortedClientNums) {
            $controls = $script:clientAuthControls[$clientNum]
            if (-not $controls) { continue }

            # Determine height based on expanded state
            $isExpanded = $script:clientAuthStates[$clientNum].IsExpanded
            $rowHeight = if ($isExpanded) { $expandedHeight } else { $minimizedHeight }

            # Update Y position for all controls in this tenant
            $controls.BorderPanel.Location = New-Object System.Drawing.Point(0, $currentY)
            $controls.BorderPanel.Height = $rowHeight

            $controls.ToggleButton.Location = New-Object System.Drawing.Point(10, ($currentY + 10))
            $controls.ClientLabel.Location = New-Object System.Drawing.Point(50, ($currentY + 15))
            $controls.StatusLabel.Location = New-Object System.Drawing.Point(270, ($currentY + 15))
            $controls.WarningLabel.Location = New-Object System.Drawing.Point(270, ($currentY + 35))

            # Minimized view controls
            $controls.GraphStatusLabel.Location = New-Object System.Drawing.Point(480, ($currentY + 15))
            $controls.ExchangeStatusLabel.Location = New-Object System.Drawing.Point(590, ($currentY + 15))
            $controls.OpenReportsButton.Location = New-Object System.Drawing.Point(720, ($currentY + 10))
            $controls.RemoveMinimizedButton.Location = New-Object System.Drawing.Point(850, ($currentY + 10))

            # Expanded view controls
            $controls.GraphButton.Location = New-Object System.Drawing.Point(480, ($currentY + 10))
            $controls.ExchangeButton.Location = New-Object System.Drawing.Point(610, ($currentY + 10))
            $controls.RemoveButton.Location = New-Object System.Drawing.Point(760, ($currentY + 10))
            $controls.ResetButton.Location = New-Object System.Drawing.Point(840, ($currentY + 10))

            $controls.UserFilterCheckBox.Location = New-Object System.Drawing.Point(10, ($currentY + 50))
            $controls.UserSearchTextBox.Location = New-Object System.Drawing.Point(120, ($currentY + 48))
            $controls.ValidateUsersButton.Location = New-Object System.Drawing.Point(330, ($currentY + 47))
            $controls.UserValidationLabel.Location = New-Object System.Drawing.Point(410, ($currentY + 50))

            $controls.TicketLabel.Location = New-Object System.Drawing.Point(10, ($currentY + 75))
            $controls.TicketTextBox.Location = New-Object System.Drawing.Point(170, ($currentY + 73))
            $controls.TicketNumbersLabel.Location = New-Object System.Drawing.Point(580, ($currentY + 73))

            $controls.GenerateReportsButton.Location = New-Object System.Drawing.Point(760, ($currentY + 47))
            $controls.ViewReportsButton.Location = New-Object System.Drawing.Point(760, ($currentY + 160))

            # Move to next position
            $currentY += $rowHeight + $clientRowSpacing
        }
    }

    # Function to attempt auto-populating email addresses from ticket content
    function Attempt-AutoPopulateEmails {
        param([int]$ClientNumber)

        $controls = $script:clientAuthControls[$ClientNumber]
        $state = $script:clientAuthStates[$ClientNumber]

        # Check prerequisites
        # 1. Both Graph AND Exchange auth must be complete
        if (-not $state.GraphAuthenticated -or -not $state.ExchangeAuthenticated) {
            return $false
        }

        # 2. User search textbox must be empty
        if (-not [string]::IsNullOrWhiteSpace($controls.UserSearchTextBox.Text)) {
            return $false
        }

        # 3. Must have ticket content
        if (-not $script:clientTickets.ContainsKey($ClientNumber)) {
            return $false
        }
        $ticketData = $script:clientTickets[$ClientNumber]
        if (-not $ticketData -or [string]::IsNullOrWhiteSpace($ticketData.Content)) {
            return $false
        }

        # 4. Must have tenant domains
        if (-not $state.TenantDomains -or $state.TenantDomains.Count -eq 0) {
            return $false
        }

        # Import Settings module to access Extract-EmailsFromTicket
        try {
            Import-Module "$script:scriptRoot\Modules\Settings.psm1" -Force -ErrorAction Stop
        } catch {
            Write-Host "Warning: Failed to load Settings module: $($_.Exception.Message)" -ForegroundColor Yellow
            return $false
        }

        # Extract emails from ticket content
        $emails = @()
        try {
            if (Get-Command Extract-EmailsFromTicket -ErrorAction SilentlyContinue) {
                $emails = Extract-EmailsFromTicket -TicketContent $ticketData.Content -TenantDomains $state.TenantDomains
            }
        } catch {
            Write-Host "Warning: Failed to extract emails from ticket: $($_.Exception.Message)" -ForegroundColor Yellow
            return $false
        }

        if (-not $emails -or $emails.Count -eq 0) {
            return $false
        }

        # Populate user search textbox
        $emailsText = $emails -join '; '
        $controls.UserSearchTextBox.Text = $emailsText

        # Show visual feedback
        $controls.UserValidationLabel.Text = "Auto-detected $($emails.Count) email(s) from ticket"
        $controls.UserValidationLabel.ForeColor = [System.Drawing.Color]::Blue
        $controls.UserValidationLabel.Visible = $true

        # Auto-validate (since field was empty and we populated it)
        try {
            $controls.ValidateUsersButton.PerformClick()
        } catch {
            Write-Host "Warning: Auto-validation failed: $($_.Exception.Message)" -ForegroundColor Yellow
        }

        return $true
    }

    # Function to add a new tenant dynamically
    function Add-NewTenant {
        param([int]$ClientNumber)
        
        # Launch PowerShell process for this client
        $statusFile = Join-Path $script:tempDir "Client${ClientNumber}_Status.txt"
        $resultFile = Join-Path $script:tempDir "Client${ClientNumber}_Result.txt"
        
        # Build process arguments - use $script:scriptRoot instead of $PSScriptRoot
        # Pass SelectedUsers as comma-separated string if provided
        $selectedUsersArg = ""
        if ($script:selectedUsers -and $script:selectedUsers.Count -gt 0) {
            # Escape single quotes in UPNs and build array argument
            $escapedUsers = $script:selectedUsers | ForEach-Object { $_.Replace("'", "''") }
            $selectedUsersArg = " -SelectedUsers @('$($escapedUsers -join "','")')"
        }
        $processArgs = "-NoProfile -ExecutionPolicy Bypass -File `"$script:workerScriptFile`" -ClientNumber $ClientNumber -ScriptRoot `"$script:scriptRoot`" -InvestigatorName `"$script:investigator`" -CompanyName `"$script:company`" -DaysBack $script:days -ReportSelectionsFile `"$script:reportSelectionsFile`" -StatusFile `"$statusFile`" -ResultFile `"$resultFile`" -CommandDir `"$script:commandDir`"$selectedUsersArg"

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
        # Row height must account for: buttons (30px) + user filter row (25px) + ticket row (80px) + view reports button (25px) + spacing
        $clientRowHeight = 200  # Increased to accommodate all controls including ticket textbox (80px) and view reports button
        $clientRowSpacing = 10  # Increased spacing between rows
        $existingRows = ($script:clientAuthControls.Keys | Measure-Object).Count
        $yPos = $existingRows * ($clientRowHeight + $clientRowSpacing) + 10
        
        # Client label
        $clientLabel = New-Object System.Windows.Forms.Label
        $clientLabel.Text = "Client $ClientNumber"
        $clientLabel.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
        $clientLabel.Location = New-Object System.Drawing.Point(50, ($yPos + 15))
        $clientLabel.Size = New-Object System.Drawing.Size(210, 20)
        $clientLabel.AutoEllipsis = $true

        # Status label
        $statusLabel = New-Object System.Windows.Forms.Label
        $statusLabel.Text = "Initializing..."
        $statusLabel.Font = New-Object System.Drawing.Font('Segoe UI', 9)
        $statusLabel.Location = New-Object System.Drawing.Point(270, ($yPos + 15))
        $statusLabel.Size = New-Object System.Drawing.Size(200, 20)
        $statusLabel.ForeColor = [System.Drawing.Color]::Gray

        # Warning label (for license issues, etc.)
        $warningLabel = New-Object System.Windows.Forms.Label
        $warningLabel.Text = ""
        $warningLabel.Font = New-Object System.Drawing.Font('Segoe UI', 8, [System.Drawing.FontStyle]::Bold)
        $warningLabel.Location = New-Object System.Drawing.Point(270, ($yPos + 35))
        $warningLabel.Size = New-Object System.Drawing.Size(600, 15)
        $warningLabel.ForeColor = [System.Drawing.Color]::Orange
        $warningLabel.Visible = $false
        $warningLabel.AutoEllipsis = $true

        # Border panel for status indication (color-coded left border)
        $borderPanel = New-Object System.Windows.Forms.Panel
        $borderPanel.Location = New-Object System.Drawing.Point(0, $yPos)
        $borderPanel.Size = New-Object System.Drawing.Size(5, $clientRowHeight)
        $borderPanel.BackColor = [System.Drawing.Color]::Gray  # Default: Not started

        # Toggle button (▼ for expanded, ▶ for minimized)
        $toggleBtn = New-Object System.Windows.Forms.Button
        $toggleBtn.Text = "▼"
        $toggleBtn.Location = New-Object System.Drawing.Point(10, ($yPos + 10))
        $toggleBtn.Size = New-Object System.Drawing.Size(30, 30)
        $toggleBtn.Tag = $ClientNumber
        $toggleBtn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $toggleBtn.Font = New-Object System.Drawing.Font('Segoe UI', 10)

        # Graph Status Indicator (for minimized view)
        $graphStatusLabel = New-Object System.Windows.Forms.Label
        $graphStatusLabel.Text = "Graph: ○"
        $graphStatusLabel.Font = New-Object System.Drawing.Font('Segoe UI', 9)
        $graphStatusLabel.Location = New-Object System.Drawing.Point(480, ($yPos + 15))
        $graphStatusLabel.Size = New-Object System.Drawing.Size(100, 20)
        $graphStatusLabel.ForeColor = [System.Drawing.Color]::Gray
        $graphStatusLabel.Visible = $false  # Only visible when minimized

        # Exchange Status Indicator (for minimized view)
        $exchangeStatusLabel = New-Object System.Windows.Forms.Label
        $exchangeStatusLabel.Text = "Exchange: ○"
        $exchangeStatusLabel.Font = New-Object System.Drawing.Font('Segoe UI', 9)
        $exchangeStatusLabel.Location = New-Object System.Drawing.Point(590, ($yPos + 15))
        $exchangeStatusLabel.Size = New-Object System.Drawing.Size(120, 20)
        $exchangeStatusLabel.ForeColor = [System.Drawing.Color]::Gray
        $exchangeStatusLabel.Visible = $false  # Only visible when minimized

        # Open Reports button (for minimized view)
        $openReportsBtn = New-Object System.Windows.Forms.Button
        $openReportsBtn.Text = "Open Reports"
        $openReportsBtn.Location = New-Object System.Drawing.Point(720, ($yPos + 10))
        $openReportsBtn.Size = New-Object System.Drawing.Size(120, 30)
        $openReportsBtn.Enabled = $false
        $openReportsBtn.Visible = $false  # Only visible when minimized and reports exist
        $openReportsBtn.Tag = $ClientNumber
        $openReportsBtn.BackColor = [System.Drawing.Color]::FromArgb(33, 150, 243)
        $openReportsBtn.ForeColor = [System.Drawing.Color]::White

        # Remove button (for minimized view)
        $removeMinimizedBtn = New-Object System.Windows.Forms.Button
        $removeMinimizedBtn.Text = "×"
        $removeMinimizedBtn.Location = New-Object System.Drawing.Point(850, ($yPos + 10))
        $removeMinimizedBtn.Size = New-Object System.Drawing.Size(30, 30)
        $removeMinimizedBtn.Enabled = $true
        $removeMinimizedBtn.Visible = $false  # Only visible when minimized
        $removeMinimizedBtn.Tag = $ClientNumber
        $removeMinimizedBtn.ForeColor = [System.Drawing.Color]::DarkRed
        $removeMinimizedBtn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat

        # Graph Auth button (disabled until worker script is ready)
        $graphAuthBtn = New-Object System.Windows.Forms.Button
        $graphAuthBtn.Text = "Graph Auth (Waiting...)"
        $graphAuthBtn.Location = New-Object System.Drawing.Point(480, ($yPos + 10))
        $graphAuthBtn.Size = New-Object System.Drawing.Size(120, 30)
        $graphAuthBtn.Enabled = $false  # Disabled until worker script is ready
        $graphAuthBtn.Tag = $ClientNumber

        # User Filtering Checkbox (shown after Graph Auth, on second row)
        $userFilterCheckBox = New-Object System.Windows.Forms.CheckBox
        $userFilterCheckBox.Text = "Filter by users"
        $userFilterCheckBox.Location = New-Object System.Drawing.Point(10, ($yPos + 50))
        $userFilterCheckBox.Size = New-Object System.Drawing.Size(100, 20)
        $userFilterCheckBox.Enabled = $false  # Enabled after Graph Auth
        $userFilterCheckBox.Visible = $false  # Shown after Graph Auth
        $userFilterCheckBox.Tag = $ClientNumber

        # User Search TextBox
        $userSearchTextBox = New-Object System.Windows.Forms.TextBox
        $userSearchTextBox.Location = New-Object System.Drawing.Point(120, ($yPos + 48))
        $userSearchTextBox.Size = New-Object System.Drawing.Size(200, 20)
        $userSearchTextBox.Enabled = $false
        $userSearchTextBox.Visible = $false
        $userSearchTextBox.Tag = $ClientNumber

        # Validate Users Button
        $validateUsersBtn = New-Object System.Windows.Forms.Button
        $validateUsersBtn.Text = "Validate"
        $validateUsersBtn.Location = New-Object System.Drawing.Point(330, ($yPos + 47))
        $validateUsersBtn.Size = New-Object System.Drawing.Size(70, 25)
        $validateUsersBtn.Enabled = $false
        $validateUsersBtn.Visible = $false
        $validateUsersBtn.Tag = $ClientNumber

        # User Validation Status Label
        $userValidationLabel = New-Object System.Windows.Forms.Label
        $userValidationLabel.Text = ""
        $userValidationLabel.Location = New-Object System.Drawing.Point(410, ($yPos + 50))
        $userValidationLabel.Size = New-Object System.Drawing.Size(160, 15)
        $userValidationLabel.ForeColor = [System.Drawing.Color]::Blue
        $userValidationLabel.Font = New-Object System.Drawing.Font('Segoe UI', 8)
        $userValidationLabel.Visible = $false

        # ConnectWise Ticket Label
        $ticketLabel = New-Object System.Windows.Forms.Label
        $ticketLabel.Text = "ConnectWise Ticket(s):"
        $ticketLabel.Location = New-Object System.Drawing.Point(10, ($yPos + 75))
        $ticketLabel.Size = New-Object System.Drawing.Size(150, 20)
        $ticketLabel.Enabled = $false
        $ticketLabel.Visible = $false  # Shown after Exchange Auth
        $ticketLabel.Font = New-Object System.Drawing.Font('Segoe UI', 9)

        # ConnectWise Ticket TextBox (multiline)
        $ticketTextBox = New-Object System.Windows.Forms.TextBox
        $ticketTextBox.Multiline = $true
        $ticketTextBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
        $ticketTextBox.Location = New-Object System.Drawing.Point(170, ($yPos + 73))
        $ticketTextBox.Size = New-Object System.Drawing.Size(400, 80)
        $ticketTextBox.Enabled = $false
        $ticketTextBox.Visible = $false  # Shown after Exchange Auth
        $ticketTextBox.Tag = $ClientNumber
        $ticketTextBox.ShortcutsEnabled = $true
        $ticketTextBox.AcceptsReturn = $true
        $ticketTextBox.AcceptsTab = $false

        # Ticket Numbers Detected Label
        $ticketNumbersLabel = New-Object System.Windows.Forms.Label
        $ticketNumbersLabel.Text = ""
        $ticketNumbersLabel.Location = New-Object System.Drawing.Point(580, ($yPos + 73))
        $ticketNumbersLabel.Size = New-Object System.Drawing.Size(200, 15)
        $ticketNumbersLabel.ForeColor = [System.Drawing.Color]::DarkGreen
        $ticketNumbersLabel.Font = New-Object System.Drawing.Font('Segoe UI', 8)
        $ticketNumbersLabel.Visible = $false

        # Extract Emails Button (to the left of Generate Reports button)
        $extractEmailsBtn = New-Object System.Windows.Forms.Button
        $extractEmailsBtn.Text = "Extract Emails from Ticket"
        $extractEmailsBtn.Location = New-Object System.Drawing.Point(580, ($yPos + 47))
        $extractEmailsBtn.Size = New-Object System.Drawing.Size(170, 25)
        $extractEmailsBtn.Enabled = $false
        $extractEmailsBtn.Visible = $false
        $extractEmailsBtn.Tag = $ClientNumber
        $extractEmailsBtn.BackColor = [System.Drawing.Color]::FromArgb(94, 53, 177)
        $extractEmailsBtn.ForeColor = [System.Drawing.Color]::White

        # Exchange Online Auth button
        $exchangeAuthBtn = New-Object System.Windows.Forms.Button
        $exchangeAuthBtn.Text = "Exchange Online Auth"
        $exchangeAuthBtn.Location = New-Object System.Drawing.Point(610, ($yPos + 10))
        $exchangeAuthBtn.Size = New-Object System.Drawing.Size(140, 30)
        $exchangeAuthBtn.Enabled = $false
        $exchangeAuthBtn.Tag = $ClientNumber

        # Remove Tenant button
        $removeTenantBtn = New-Object System.Windows.Forms.Button
        $removeTenantBtn.Text = "Remove"
        $removeTenantBtn.Location = New-Object System.Drawing.Point(760, ($yPos + 10))
        $removeTenantBtn.Size = New-Object System.Drawing.Size(70, 30)
        $removeTenantBtn.Enabled = $true
        $removeTenantBtn.Tag = $ClientNumber
        $removeTenantBtn.ForeColor = [System.Drawing.Color]::DarkRed

        # Reset Auth button
        $resetAuthBtn = New-Object System.Windows.Forms.Button
        $resetAuthBtn.Text = "Reset Auth"
        $resetAuthBtn.Location = New-Object System.Drawing.Point(840, ($yPos + 10))
        $resetAuthBtn.Size = New-Object System.Drawing.Size(90, 30)
        $resetAuthBtn.Enabled = $true
        $resetAuthBtn.Tag = $ClientNumber
        $resetAuthBtn.ForeColor = [System.Drawing.Color]::DarkRed

        # Generate Reports button (shown after Exchange Auth)
        $generateReportsBtn = New-Object System.Windows.Forms.Button
        $generateReportsBtn.Text = "Generate Reports"
        $generateReportsBtn.Location = New-Object System.Drawing.Point(760, ($yPos + 47))
        $generateReportsBtn.Size = New-Object System.Drawing.Size(140, 25)
        $generateReportsBtn.Enabled = $false
        $generateReportsBtn.Visible = $false
        $generateReportsBtn.Tag = $ClientNumber
        $generateReportsBtn.BackColor = [System.Drawing.Color]::FromArgb(46, 125, 50)
        $generateReportsBtn.ForeColor = [System.Drawing.Color]::White

        # View Reports button (shown after report generation completes)
        $viewReportsBtn = New-Object System.Windows.Forms.Button
        $viewReportsBtn.Text = "View Reports"
        $viewReportsBtn.Location = New-Object System.Drawing.Point(760, ($yPos + 160))
        $viewReportsBtn.Size = New-Object System.Drawing.Size(140, 25)
        $viewReportsBtn.Enabled = $false
        $viewReportsBtn.Visible = $false
        $viewReportsBtn.Tag = $ClientNumber
        $viewReportsBtn.BackColor = [System.Drawing.Color]::FromArgb(33, 150, 243)
        $viewReportsBtn.ForeColor = [System.Drawing.Color]::White

        # Add controls to panel
        $script:authPanel.Controls.AddRange(@($borderPanel, $toggleBtn, $clientLabel, $statusLabel, $warningLabel, $graphStatusLabel, $exchangeStatusLabel, $openReportsBtn, $removeMinimizedBtn, $graphAuthBtn, $exchangeAuthBtn, $removeTenantBtn, $resetAuthBtn, $userFilterCheckBox, $userSearchTextBox, $validateUsersBtn, $userValidationLabel, $generateReportsBtn, $ticketLabel, $ticketTextBox, $ticketNumbersLabel, $extractEmailsBtn, $viewReportsBtn))

        # Store controls and state
        $script:clientAuthStates[$ClientNumber] = @{
            GraphAuthenticated = $false
            ExchangeAuthenticated = $false
            GraphContext = $null
            TenantId = $null
            TenantName = $null
            TenantDomains = @()  # All verified domains for the tenant
            Account = $null
            IsExpanded = $true  # Start expanded so user can interact with fields
        }
        $script:clientAuthControls[$ClientNumber] = @{
            BorderPanel = $borderPanel
            ToggleButton = $toggleBtn
            ClientLabel = $clientLabel
            StatusLabel = $statusLabel
            WarningLabel = $warningLabel
            GraphStatusLabel = $graphStatusLabel
            ExchangeStatusLabel = $exchangeStatusLabel
            OpenReportsButton = $openReportsBtn
            RemoveMinimizedButton = $removeMinimizedBtn
            GraphButton = $graphAuthBtn
            ExchangeButton = $exchangeAuthBtn
            RemoveButton = $removeTenantBtn
            ResetButton = $resetAuthBtn
            UserFilterCheckBox = $userFilterCheckBox
            UserSearchTextBox = $userSearchTextBox
            ValidateUsersButton = $validateUsersBtn
            UserValidationLabel = $userValidationLabel
            GenerateReportsButton = $generateReportsBtn
            TicketLabel = $ticketLabel
            TicketTextBox = $ticketTextBox
            TicketNumbersLabel = $ticketNumbersLabel
            ExtractEmailsButton = $extractEmailsBtn
            ViewReportsButton = $viewReportsBtn
        }

        # View Reports button handler
        $capturedClientNumForView = $ClientNumber
        $viewReportsBtn.add_Click({
            $clientNum = $this.Tag
            if (-not $clientNum) { $clientNum = $capturedClientNumForView }

            if ($script:clientReportFolders.ContainsKey($clientNum)) {
                $reportFolder = $script:clientReportFolders[$clientNum]
                if ($reportFolder) {
                    $reportFolder = $reportFolder.Trim()
                    if (Test-Path $reportFolder) {
                        Start-Process explorer.exe -ArgumentList "`"$reportFolder`""
                    } else {
                        [System.Windows.Forms.MessageBox]::Show("Report folder not found: $reportFolder", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                    }
                } else {
                    [System.Windows.Forms.MessageBox]::Show("Report folder path is empty for Client $clientNum", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                }
            } else {
                [System.Windows.Forms.MessageBox]::Show("No report folder available for Client $clientNum", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            }
        })

        # Extract Emails button handler
        $capturedClientNumForExtract = $ClientNumber
        $extractEmailsBtn.add_Click({
            $clientNum = $this.Tag
            if (-not $clientNum) { $clientNum = $capturedClientNumForExtract }

            # Get controls and state
            $controls = $script:clientAuthControls[$clientNum]
            $state = $script:clientAuthStates[$clientNum]

            # Check prerequisites
            if (-not $state.GraphAuthenticated -or -not $state.ExchangeAuthenticated) {
                [System.Windows.Forms.MessageBox]::Show("Both Graph and Exchange authentication must be complete before extracting emails.", "Authentication Required", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                return
            }

            # Read ticket content directly from the textbox
            $ticketContent = $controls.TicketTextBox.Text
            if ([string]::IsNullOrWhiteSpace($ticketContent)) {
                [System.Windows.Forms.MessageBox]::Show("Please paste ticket content first.", "No Ticket Content", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                return
            }

            if (-not $state.TenantDomains -or $state.TenantDomains.Count -eq 0) {
                [System.Windows.Forms.MessageBox]::Show("No tenant domains found. Please ensure Graph authentication completed successfully.", "No Tenant Domains", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                return
            }

            # Import Settings module to access Extract-EmailsFromTicket
            try {
                Import-Module "$script:scriptRoot\Modules\Settings.psm1" -Force -ErrorAction Stop
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Failed to load Settings module: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                return
            }

            # Extract emails from ticket content
            $emails = @()
            try {
                if (Get-Command Extract-EmailsFromTicket -ErrorAction SilentlyContinue) {
                    $emails = Extract-EmailsFromTicket -TicketContent $ticketContent -TenantDomains $state.TenantDomains
                }
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Failed to extract emails from ticket: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                return
            }

            if (-not $emails -or $emails.Count -eq 0) {
                [System.Windows.Forms.MessageBox]::Show("No emails matching tenant domains found in ticket content.`n`nTenant domains: $($state.TenantDomains -join ', ')", "No Emails Found", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                return
            }

            # Populate user search textbox
            $emailsText = $emails -join '; '
            $controls.UserSearchTextBox.Text = $emailsText

            # Show visual feedback
            $controls.UserValidationLabel.Text = "Extracted $($emails.Count) email(s) from ticket"
            $controls.UserValidationLabel.ForeColor = [System.Drawing.Color]::Blue
            $controls.UserValidationLabel.Visible = $true

            # Auto-validate
            try {
                $controls.ValidateUsersButton.PerformClick()
            } catch {
                Write-Host "Warning: Auto-validation failed: $($_.Exception.Message)" -ForegroundColor Yellow
            }
        })

        # Add event handler to extract ticket numbers when text changes
        $capturedClientNum = $ClientNumber
        $ticketTextBox.add_TextChanged({
            try {
                $ticketContent = $this.Text

                # Store ticket content in clientTickets hashtable
                if (-not [string]::IsNullOrWhiteSpace($ticketContent)) {
                    if (-not $script:clientTickets) {
                        $script:clientTickets = @{}
                    }
                    $script:clientTickets[$capturedClientNum] = @{
                        Content = $ticketContent
                        Numbers = @()
                    }
                } else {
                    # Clear ticket data if content is empty
                    if ($script:clientTickets -and $script:clientTickets.ContainsKey($capturedClientNum)) {
                        $script:clientTickets.Remove($capturedClientNum)
                    }
                }

                Import-Module "$script:scriptRoot\Modules\Settings.psm1" -Force -ErrorAction SilentlyContinue
                if (Get-Command Extract-TicketNumbers -ErrorAction SilentlyContinue) {
                    if (-not [string]::IsNullOrWhiteSpace($ticketContent)) {
                        $ticketNums = Extract-TicketNumbers -TicketContent $ticketContent
                        if ($ticketNums -and $ticketNums.Count -gt 0) {
                            $ticketNumsStr = ($ticketNums | ForEach-Object { "#$_" }) -join ', '
                            $script:clientAuthControls[$capturedClientNum].TicketNumbersLabel.Text = "Detected: $ticketNumsStr"
                            $script:clientAuthControls[$capturedClientNum].TicketNumbersLabel.Visible = $true
                            # Store ticket numbers in hashtable
                            if ($script:clientTickets.ContainsKey($capturedClientNum)) {
                                $script:clientTickets[$capturedClientNum].Numbers = $ticketNums
                            }
                        } else {
                            $script:clientAuthControls[$capturedClientNum].TicketNumbersLabel.Text = ""
                            $script:clientAuthControls[$capturedClientNum].TicketNumbersLabel.Visible = $false
                        }
                    } else {
                        $script:clientAuthControls[$capturedClientNum].TicketNumbersLabel.Text = ""
                        $script:clientAuthControls[$capturedClientNum].TicketNumbersLabel.Visible = $false
                    }
                }

                # Enable Extract Emails button if both auths complete and ticket has content
                $ticketContent = $this.Text
                if ($script:clientAuthStates.ContainsKey($capturedClientNum)) {
                    $state = $script:clientAuthStates[$capturedClientNum]
                    if ($state.GraphAuthenticated -and $state.ExchangeAuthenticated -and
                        -not [string]::IsNullOrWhiteSpace($ticketContent) -and
                        $script:clientAuthControls[$capturedClientNum].ExtractEmailsButton) {
                        $script:clientAuthControls[$capturedClientNum].ExtractEmailsButton.Enabled = $true
                    }
                }

                # Attempt auto-population of emails from ticket (if conditions are met)
                Attempt-AutoPopulateEmails -ClientNumber $capturedClientNum
            } catch {
                # Ignore errors
            }
        })

        # Toggle button handler (minimize/expand tenant display)
        $toggleBtn.add_Click({
            $clientNum = $this.Tag
            if (-not $clientNum) { $clientNum = $capturedClientNum }

            # Toggle the expanded state
            $script:clientAuthStates[$clientNum].IsExpanded = -not $script:clientAuthStates[$clientNum].IsExpanded
            $isExpanded = $script:clientAuthStates[$clientNum].IsExpanded

            # Get controls
            $controls = $script:clientAuthControls[$clientNum]
            if (-not $controls) { return }

            # Update toggle button text
            $this.Text = if ($isExpanded) { "▼" } else { "▶" }

            # Calculate heights
            $minimizedHeight = 50
            $expandedHeight = 200

            # Show/hide controls based on state
            if ($isExpanded) {
                # Expanded view - hide minimized controls, show expanded controls
                $controls.GraphStatusLabel.Visible = $false
                $controls.ExchangeStatusLabel.Visible = $false
                $controls.OpenReportsButton.Visible = $false
                $controls.RemoveMinimizedButton.Visible = $false

                # Show expanded controls
                $controls.GraphButton.Visible = $true
                $controls.ExchangeButton.Visible = $true
                $controls.RemoveButton.Visible = $true
                $controls.ResetButton.Visible = $true

                # Show controls based on auth state
                if ($script:clientAuthStates[$clientNum].GraphAuthenticated) {
                    $controls.UserFilterCheckBox.Visible = $true
                    $controls.UserSearchTextBox.Visible = $true
                    $controls.ValidateUsersButton.Visible = $true
                }

                if ($script:clientAuthStates[$clientNum].ExchangeAuthenticated) {
                    $controls.TicketLabel.Visible = $true
                    $controls.TicketTextBox.Visible = $true
                    $controls.GenerateReportsButton.Visible = $true
                }

                # Show View Reports if available
                if ($script:clientReportFolders.ContainsKey($clientNum) -and $script:clientReportFolders[$clientNum]) {
                    $controls.ViewReportsButton.Visible = $true
                }
            } else {
                # Minimized view - show minimized controls, hide expanded controls
                $controls.GraphStatusLabel.Visible = $true
                $controls.ExchangeStatusLabel.Visible = $true
                $controls.RemoveMinimizedButton.Visible = $true

                # Show Open Reports if available
                if ($script:clientReportFolders.ContainsKey($clientNum) -and $script:clientReportFolders[$clientNum]) {
                    $controls.OpenReportsButton.Visible = $true
                    $controls.OpenReportsButton.Enabled = $true
                }

                # Hide expanded controls
                $controls.GraphButton.Visible = $false
                $controls.ExchangeButton.Visible = $false
                $controls.RemoveButton.Visible = $false
                $controls.ResetButton.Visible = $false
                $controls.UserFilterCheckBox.Visible = $false
                $controls.UserSearchTextBox.Visible = $false
                $controls.ValidateUsersButton.Visible = $false
                $controls.UserValidationLabel.Visible = $false
                $controls.TicketLabel.Visible = $false
                $controls.TicketTextBox.Visible = $false
                $controls.TicketNumbersLabel.Visible = $false
                $controls.GenerateReportsButton.Visible = $false
                $controls.ViewReportsButton.Visible = $false
                $controls.WarningLabel.Visible = $false
            }

            # Update border panel height
            $newHeight = if ($isExpanded) { $expandedHeight } else { $minimizedHeight }
            $controls.BorderPanel.Height = $newHeight

            # Recalculate positions of all tenants
            Update-TenantPositions
        })

        # Open Reports button handler (minimized view)
        $openReportsBtn.add_Click({
            $clientNum = $this.Tag
            if (-not $clientNum) { $clientNum = $capturedClientNumForView }

            if ($script:clientReportFolders.ContainsKey($clientNum)) {
                $reportFolder = $script:clientReportFolders[$clientNum]
                if ($reportFolder -and (Test-Path $reportFolder)) {
                    Start-Process explorer.exe -ArgumentList "`"$reportFolder`""
                }
            }
        })

        # Remove Minimized button handler
        $removeMinimizedBtn.add_Click({
            $clientNum = $this.Tag
            if (-not $clientNum) { $clientNum = $capturedClientNum }

            # Use the same logic as the regular remove button
            $controls = $script:clientAuthControls[$clientNum]
            if ($controls -and $controls.RemoveButton) {
                $controls.RemoveButton.PerformClick()
            }
        })

        # Update panel height to accommodate new row (accounting for user filtering row, warning label, and ticket controls)
        $newHeight = ($existingRows + 1) * ($clientRowHeight + $clientRowSpacing) + 100  # Extra space for user filtering row, warning label, and ticket controls
        if ($newHeight -gt 420) {
            $script:authPanel.AutoScroll = $true
        }

        # Wire up button handlers
        $capturedClientNum = $ClientNumber
        
        # User Filter Checkbox handler
        $userFilterCheckBox.add_CheckedChanged({
            $clientNum = $this.Tag
            if (-not $clientNum) { $clientNum = $capturedClientNum }
            $controls = $script:clientAuthControls[$clientNum]
            if ($controls) {
                $controls.UserSearchTextBox.Enabled = $this.Checked
                $controls.ValidateUsersButton.Enabled = $this.Checked
                if (-not $this.Checked) {
                    $controls.UserSearchTextBox.Text = ""
                    $controls.UserValidationLabel.Text = ""
                    if ($script:clientValidatedUsers.ContainsKey($clientNum)) {
                        $script:clientValidatedUsers.Remove($clientNum)
                    }
                }
            }
        })
        
        # Validate Users button handler (per tenant)
        $validateUsersBtn.add_Click({
            $clientNum = $this.Tag
            if (-not $clientNum) { $clientNum = $capturedClientNum }
            $controls = $script:clientAuthControls[$clientNum]
            
            if (-not $controls -or [string]::IsNullOrWhiteSpace($controls.UserSearchTextBox.Text)) {
                [System.Windows.Forms.MessageBox]::Show("Please enter user search terms.", "No Search Terms", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                return
            }
            
            # Check if Graph is connected for this tenant
            if (-not $script:clientAuthStates[$clientNum].GraphAuthenticated) {
                [System.Windows.Forms.MessageBox]::Show("Please complete Graph authentication first.", "Not Authenticated", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                return
            }
            
            try {
                $this.Enabled = $false
                $controls.UserValidationLabel.Text = "Validating..."
                $controls.UserValidationLabel.ForeColor = [System.Drawing.Color]::Blue
                [System.Windows.Forms.Application]::DoEvents()
                
                # Send VALIDATE_USERS command to worker script (which has the Graph context)
                $searchTerms = $controls.UserSearchTextBox.Text
                $searchTermsArray = ($searchTerms -split ',' | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
                $searchTermsJson = ($searchTermsArray | ConvertTo-Json -Compress)
                
                $command = "VALIDATE_USERS|SEARCH_TERMS:$searchTermsJson"
                Write-Host "Sending VALIDATE_USERS command to Client $clientNum with search terms: $searchTerms" -ForegroundColor Cyan
                $script:authStatusTextBox.AppendText("Client $clientNum : Validating users: $searchTerms`r`n")
                
                $response = Send-CommandToSession -ClientNumber $clientNum -Command $command -TimeoutSeconds 60
                
                # If we got VALIDATE_USERS_STARTED, continue polling for the final result
                if ($response -eq "VALIDATE_USERS_STARTED") {
                    $script:authStatusTextBox.AppendText("Client $clientNum : User validation started. Searching...`r`n")
                    $script:authStatusTextBox.ScrollToCaret()
                    [System.Windows.Forms.Application]::DoEvents()
                    
                    # Continue polling the response file for the final result
                    $responseFile = Join-Path $script:commandDir "Client${clientNum}_Response.txt"
                    $startTime = Get-Date
                    $finalResponse = $null
                    $pollCount = 0
                    
                    while (((Get-Date) - $startTime).TotalSeconds -lt 60) {
                        $pollCount++
                        $elapsedSeconds = [int]((Get-Date) - $startTime).TotalSeconds
                        
                        # Update status every 5 seconds
                        if ($pollCount % 25 -eq 0) {
                            $statusMsg = "Validating users... (${elapsedSeconds}s elapsed)"
                            $controls.UserValidationLabel.Text = $statusMsg
                            $script:authStatusTextBox.AppendText("Client ${clientNum}: $statusMsg`r`n")
                            $script:authStatusTextBox.ScrollToCaret()
                            [System.Windows.Forms.Application]::DoEvents()
                        }
                        
                        if (Test-Path $responseFile) {
                            Start-Sleep -Milliseconds 200
                            try {
                                $finalResponse = Get-Content $responseFile -Raw -ErrorAction Stop | ForEach-Object { $_.Trim() }
                                # Check if we got a final response (not VALIDATE_USERS_STARTED)
                                if ($finalResponse -and $finalResponse -ne "VALIDATE_USERS_STARTED" -and $finalResponse -notmatch "^VALIDATE_USERS_STARTED") {
                                    $script:authStatusTextBox.AppendText("Client ${clientNum}: Final validation response received`r`n")
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
                    
                    if (-not $finalResponse -or $finalResponse -eq "VALIDATE_USERS_STARTED") {
                        $script:authStatusTextBox.AppendText("Client ${clientNum}: Timeout waiting for user validation response.`r`n")
                        $script:authStatusTextBox.ScrollToCaret()
                        $controls.UserValidationLabel.Text = "Validation timeout"
                        $controls.UserValidationLabel.ForeColor = [System.Drawing.Color]::Red
                        [System.Windows.Forms.MessageBox]::Show("Timeout waiting for user validation response for Client $clientNum.", "Validation Timeout", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                        return
                    }
                }
                
                if ($response -match "^VALIDATE_USERS_SUCCESS:(.+)$") {
                    $responseJson = $Matches[1]
                    try {
                        $result = $responseJson | ConvertFrom-Json
                        
                        if ($result.Success -and $result.UserCount -gt 0) {
                            $validatedUsers = if ($result.Users -is [array]) { $result.Users } else { @($result.Users) }
                            $script:clientValidatedUsers[$clientNum] = $validatedUsers
                            $controls.UserValidationLabel.Text = "Validated: $($validatedUsers.Count) user(s)"
                            $controls.UserValidationLabel.ForeColor = [System.Drawing.Color]::Green
                            $script:authStatusTextBox.AppendText("Client $clientNum : Found $($validatedUsers.Count) user(s)`r`n")
                            [System.Windows.Forms.MessageBox]::Show("Found and validated $($validatedUsers.Count) user(s) for Client $clientNum :`n`n$($validatedUsers -join "`n")", "Validation Successful", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                        } else {
                            if ($script:clientValidatedUsers.ContainsKey($clientNum)) {
                                $script:clientValidatedUsers.Remove($clientNum)
                            }
                            $controls.UserValidationLabel.Text = "No users found"
                            $controls.UserValidationLabel.ForeColor = [System.Drawing.Color]::Red
                            $message = if ($result.Message) { $result.Message } else { "No users found matching the search terms." }
                            $script:authStatusTextBox.AppendText("Client $clientNum : $message`r`n")
                            [System.Windows.Forms.MessageBox]::Show("$message for Client $clientNum.", "No Users Found", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                        }
                    } catch {
                        Write-Host "Failed to parse validation response: $($_.Exception.Message)" -ForegroundColor Red
                        $controls.UserValidationLabel.Text = "Validation failed"
                        $controls.UserValidationLabel.ForeColor = [System.Drawing.Color]::Red
                        [System.Windows.Forms.MessageBox]::Show("Error parsing validation response for Client $clientNum : $($_.Exception.Message)", "Validation Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                    }
                } elseif ($response -match "^VALIDATE_USERS_FAILED:(.+)$") {
                    $errorMsg = $Matches[1]
                    Write-Host "Validation failed: $errorMsg" -ForegroundColor Red
                    if ($script:clientValidatedUsers.ContainsKey($clientNum)) {
                        $script:clientValidatedUsers.Remove($clientNum)
                    }
                    $controls.UserValidationLabel.Text = "Validation failed"
                    $controls.UserValidationLabel.ForeColor = [System.Drawing.Color]::Red
                    $script:authStatusTextBox.AppendText("Client $clientNum : Validation failed - $errorMsg`r`n")
                    [System.Windows.Forms.MessageBox]::Show("Validation failed for Client $clientNum : $errorMsg", "Validation Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                } else {
                    Write-Host "Unexpected response from validation command: $response" -ForegroundColor Yellow
                    if ($script:clientValidatedUsers.ContainsKey($clientNum)) {
                        $script:clientValidatedUsers.Remove($clientNum)
                    }
                    $controls.UserValidationLabel.Text = "Validation failed"
                    $controls.UserValidationLabel.ForeColor = [System.Drawing.Color]::Red
                    [System.Windows.Forms.MessageBox]::Show("Unexpected response from validation command for Client $clientNum.", "Validation Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                }
            } catch {
                Write-Host "Error validating users for Client $clientNum : $($_.Exception.Message)" -ForegroundColor Red
                if ($script:clientValidatedUsers.ContainsKey($clientNum)) {
                    $script:clientValidatedUsers.Remove($clientNum)
                }
                $controls.UserValidationLabel.Text = "Validation failed"
                $controls.UserValidationLabel.ForeColor = [System.Drawing.Color]::Red
                $script:authStatusTextBox.AppendText("Client $clientNum : Validation error - $($_.Exception.Message)`r`n")
                [System.Windows.Forms.MessageBox]::Show("Error validating users for Client $clientNum : $($_.Exception.Message)", "Validation Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            } finally {
                $this.Enabled = $userFilterCheckBox.Checked
            }
        })
        
        # Generate Reports button handler - REMOVED (duplicate, replaced by handler below with ticket extraction support)
        
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
                # Parse tenant name and domains from response
                # Format: "GRAPH_AUTH_SUCCESS:tenantName" or "GRAPH_AUTH_SUCCESS:tenantName|DOMAINS:domain1,domain2,domain3"
                $responseParts = ($response -replace "^GRAPH_AUTH_SUCCESS:", "") -split '\|'
                $tenantName = $responseParts[0]

                # Parse domains if present
                $tenantDomains = @()
                foreach ($part in $responseParts) {
                    if ($part -like "DOMAINS:*") {
                        $domainsStr = $part -replace "^DOMAINS:", ""
                        $tenantDomains = $domainsStr -split ',' | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
                    }
                }

                # Fallback: if no domains returned, use tenant name as domain
                if ($tenantDomains.Count -eq 0 -and -not [string]::IsNullOrWhiteSpace($tenantName)) {
                    $tenantDomains = @($tenantName)
                }

                # Store in state
                $script:clientAuthStates[$clientNum].GraphAuthenticated = $true
                $script:clientAuthStates[$clientNum].TenantName = $tenantName
                $script:clientAuthStates[$clientNum].TenantDomains = $tenantDomains
                $script:clientAuthControls[$clientNum].ClientLabel.Text = "Client $clientNum - $tenantName"
                $script:clientAuthControls[$clientNum].StatusLabel.Text = "Graph Auth Complete - Ready for Exchange"
                $script:clientAuthControls[$clientNum].StatusLabel.ForeColor = [System.Drawing.Color]::Orange
                $script:clientAuthControls[$clientNum].ExchangeButton.Enabled = $true
                $this.Text = "Graph Auth ✓"
                
                # Show user filtering controls after Graph Auth
                $script:clientAuthControls[$clientNum].UserFilterCheckBox.Visible = $true
                $script:clientAuthControls[$clientNum].UserFilterCheckBox.Enabled = $true
                $script:clientAuthControls[$clientNum].UserSearchTextBox.Visible = $true
                $script:clientAuthControls[$clientNum].ValidateUsersButton.Visible = $true
                $script:clientAuthControls[$clientNum].UserValidationLabel.Visible = $true
                
                $script:authStatusTextBox.AppendText("Client $clientNum Graph authentication successful! Tenant: $tenantName`r`n")
                $script:authStatusTextBox.AppendText("Client $clientNum Exchange Online Auth button is now enabled. Click it to proceed.`r`n")
                $script:authStatusTextBox.AppendText("Client $clientNum User filtering controls are now available.`r`n")
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
                $script:clientAuthControls[$clientNum].StatusLabel.Text = "Exchange Auth Complete - Ready to Generate Reports"
                $script:clientAuthControls[$clientNum].StatusLabel.ForeColor = [System.Drawing.Color]::Green
                $this.Text = "Exchange Auth ✓"
                $this.Enabled = $false
                $script:authStatusTextBox.AppendText("Client $clientNum Exchange Online authentication successful!`r`n")
                $script:authStatusTextBox.AppendText("Client $clientNum Ready to generate reports. Click 'Generate Reports' button when ready.`r`n")
                
                # Show Generate Reports button
                $script:clientAuthControls[$clientNum].GenerateReportsButton.Visible = $true
                $script:clientAuthControls[$clientNum].GenerateReportsButton.Enabled = $true
                
                # Show ticket controls
                $script:clientAuthControls[$clientNum].TicketLabel.Visible = $true
                $script:clientAuthControls[$clientNum].TicketLabel.Enabled = $true
                $script:clientAuthControls[$clientNum].TicketTextBox.Visible = $true
                $script:clientAuthControls[$clientNum].TicketTextBox.Enabled = $true

                # Show and enable Extract Emails button (both auths now complete)
                $script:clientAuthControls[$clientNum].ExtractEmailsButton.Visible = $true
                $script:clientAuthControls[$clientNum].ExtractEmailsButton.Enabled = $true

                # Attempt auto-population of emails from ticket (both auths now complete)
                Attempt-AutoPopulateEmails -ClientNumber $clientNum
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
        
        # Generate Reports button handler
        $generateReportsBtn.add_Click({
            $clientNum = $this.Tag
            if (-not $clientNum) { $clientNum = $capturedClientNum }
            
            # Check if both authentications are complete
            if (-not $script:clientAuthStates[$clientNum].GraphAuthenticated -or -not $script:clientAuthStates[$clientNum].ExchangeAuthenticated) {
                [System.Windows.Forms.MessageBox]::Show("Please complete both Graph and Exchange authentication first.", "Authentication Required", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                return
            }
            
            # Check if user filtering is enabled - do this BEFORE processing ticket data
            $controls = $script:clientAuthControls[$clientNum]
            if ($controls.UserFilterCheckBox.Checked) {
                # Check if users were validated OR if search terms are stored for validation during export
                $hasValidatedUsers = $script:clientValidatedUsers.ContainsKey($clientNum) -and $script:clientValidatedUsers[$clientNum].Count -gt 0
                $hasSearchTerms = $script:clientSearchTerms.ContainsKey($clientNum) -and -not [string]::IsNullOrWhiteSpace($script:clientSearchTerms[$clientNum])
                
                Write-Host "Generate Reports: Client $clientNum - HasValidatedUsers: $hasValidatedUsers, HasSearchTerms: $hasSearchTerms" -ForegroundColor Cyan
                if ($hasSearchTerms) {
                    Write-Host "Generate Reports: Search terms for Client $clientNum : $($script:clientSearchTerms[$clientNum])" -ForegroundColor Cyan
                }
                
                if (-not $hasValidatedUsers -and -not $hasSearchTerms) {
                    # No validation and no search terms - ask if they want to proceed
                    Write-Host "Generate Reports: No validated users and no search terms - showing warning dialog" -ForegroundColor Yellow
                    $result = [System.Windows.Forms.MessageBox]::Show("User filtering is enabled but no users have been validated. Do you want to proceed without filtering?", "No Users Validated", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
                    if ($result -eq [System.Windows.Forms.DialogResult]::No) {
                        Write-Host "Generate Reports: User clicked No - canceling report generation" -ForegroundColor Yellow
                        $script:authStatusTextBox.AppendText("Client $clientNum : Report generation canceled. Please validate users or disable filtering.`r`n")
                        $script:authStatusTextBox.ScrollToCaret()
                        [System.Windows.Forms.Application]::DoEvents()
                        return  # Exit the function - do not proceed with report generation
                    }
                    # User clicked Yes - proceed without filtering
                    Write-Host "Generate Reports: User clicked Yes - proceeding without user filtering" -ForegroundColor Green
                }
            }
            
            # Get ticket content and extract ticket numbers
            $ticketContent = $script:clientAuthControls[$clientNum].TicketTextBox.Text
            $ticketNumbers = @()
            $filteredTicketContent = ''
            
            Write-Host "Generate Reports: Processing ticket content (length: $($ticketContent.Length))" -ForegroundColor Cyan
            Write-Host "Generate Reports: Ticket textbox exists: $($null -ne $script:clientAuthControls[$clientNum].TicketTextBox)" -ForegroundColor Gray
            Write-Host "Generate Reports: Ticket textbox text length: $($script:clientAuthControls[$clientNum].TicketTextBox.Text.Length)" -ForegroundColor Gray
            if (-not [string]::IsNullOrWhiteSpace($ticketContent)) {
                Write-Host "Generate Reports: Ticket content is not empty, extracting..." -ForegroundColor Green
                try {
                    Import-Module "$script:scriptRoot\Modules\Settings.psm1" -Force -ErrorAction SilentlyContinue
                    if (Get-Command Extract-TicketNumbers -ErrorAction SilentlyContinue) {
                        $ticketNumbers = Extract-TicketNumbers -TicketContent $ticketContent
                        Write-Host "Generate Reports: Extracted $($ticketNumbers.Count) ticket number(s): $($ticketNumbers -join ', ')" -ForegroundColor Cyan
                    } else {
                        Write-Warning "Extract-TicketNumbers function not found"
                    }
                    if (Get-Command Filter-TicketContent -ErrorAction SilentlyContinue) {
                        $filteredTicketContent = Filter-TicketContent -TicketContent $ticketContent
                        Write-Host "Generate Reports: Filtered ticket content length: $($filteredTicketContent.Length)" -ForegroundColor Cyan
                    } else {
                        $filteredTicketContent = $ticketContent
                        Write-Warning "Filter-TicketContent function not found, using raw content"
                    }
                } catch {
                    Write-Warning "Failed to process ticket content: $($_.Exception.Message)"
                    Write-Host "Exception details: $($_.Exception | Out-String)" -ForegroundColor Red
                    $filteredTicketContent = $ticketContent
                }
            } else {
                Write-Host "Generate Reports: No ticket content provided (textbox is empty or whitespace)" -ForegroundColor Yellow
                Write-Host "Generate Reports: Ticket content check - IsNullOrWhiteSpace: $([string]::IsNullOrWhiteSpace($ticketContent))" -ForegroundColor Yellow
            }
            
            Write-Host "Generate Reports: After extraction - TicketNumbers=$($ticketNumbers.Count) ($($ticketNumbers -join ', ')), FilteredContent length=$($filteredTicketContent.Length)" -ForegroundColor Cyan
            
            # Store ticket data
            if ($ticketNumbers.Count -gt 0 -or -not [string]::IsNullOrWhiteSpace($filteredTicketContent)) {
                $script:clientTickets[$clientNum] = @{
                    Content = $filteredTicketContent
                    TicketNumbers = $ticketNumbers
                }
            }
            
            # Get validated users or search terms
            $selectedUsers = @()
            if ($script:clientValidatedUsers.ContainsKey($clientNum)) {
                $selectedUsers = $script:clientValidatedUsers[$clientNum]
            } elseif ($script:clientSearchTerms.ContainsKey($clientNum)) {
                # If search terms exist but not validated, send GENERATE_REPORTS_SEARCH command
                $searchTerms = $script:clientSearchTerms[$clientNum]
                if (-not [string]::IsNullOrWhiteSpace($searchTerms)) {
                    # Parse search terms (comma-separated) into array
                    $searchTermsArray = @()
                    if ($searchTerms -match ',') {
                        $searchTermsArray = ($searchTerms -split ',' | ForEach-Object { $_.Trim() }) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
                    } else {
                        $searchTermsArray = @($searchTerms.Trim())
                    }
                    # Convert to JSON array for proper parsing (ensure it's always an array, not a string)
                    $searchTermsJson = ($searchTermsArray | ConvertTo-Json -Compress)
                    # Ensure it's a JSON array (not a string) - if ConvertTo-Json returned a string, wrap it
                    if ($searchTermsJson -notmatch '^\[') {
                        $searchTermsJson = "[$searchTermsJson]"
                    }
                    $command = "GENERATE_REPORTS_SEARCH:$searchTermsJson"
                    # Include ticket data if we have ticket numbers OR ticket content
                    Write-Host "Generate Reports (SEARCH): Checking ticket data - TicketNumbers.Count=$($ticketNumbers.Count), FilteredContent length=$($filteredTicketContent.Length), IsNullOrWhiteSpace=$([string]::IsNullOrWhiteSpace($filteredTicketContent))" -ForegroundColor Cyan
                    if ($ticketNumbers.Count -gt 0 -or -not [string]::IsNullOrWhiteSpace($filteredTicketContent)) {
                        Write-Host "Generate Reports (SEARCH): Ticket data condition met, including in command" -ForegroundColor Green
                        # Ensure ticketNumbers is always an array for JSON serialization
                        $ticketNumsArray = if ($ticketNumbers -is [array]) { $ticketNumbers } else { @($ticketNumbers) }
                        Write-Host "Generate Reports (SEARCH): TicketNumbers array: $($ticketNumsArray -join ', ')" -ForegroundColor Gray
                        # Force TicketNumbers to be serialized as an array by ensuring it's always an array type
                        $ticketDataObj = [PSCustomObject]@{
                            TicketNumbers = [array]$ticketNumsArray
                            TicketContent = [string]$filteredTicketContent
                        }
                        $ticketDataJson = ($ticketDataObj | ConvertTo-Json -Compress -Depth 10)
                        Write-Host "Generate Reports (SEARCH): Ticket data JSON before verification: $($ticketDataJson.Substring(0, [Math]::Min(300, $ticketDataJson.Length)))..." -ForegroundColor Gray
                        # Verify TicketNumbers is an array in JSON (should be ["1811523"], not "1811523")
                        if ($ticketDataJson -notmatch '"TicketNumbers"\s*:\s*\[') {
                            Write-Warning "TicketNumbers was not serialized as an array, fixing..."
                            # Manually fix the JSON if needed
                            $ticketDataJson = $ticketDataJson -replace '"TicketNumbers"\s*:\s*"([^"]+)"', '"TicketNumbers":["$1"]'
                            Write-Host "Generate Reports (SEARCH): Ticket data JSON after fix: $($ticketDataJson.Substring(0, [Math]::Min(300, $ticketDataJson.Length)))..." -ForegroundColor Yellow
                        }
                        $command += "|TICKET_DATA:$ticketDataJson"
                        Write-Host "Generate Reports (SEARCH): Including ticket data - TicketNumbers=$($ticketNumsArray.Count) ($($ticketNumsArray -join ', ')), TicketContent length=$($filteredTicketContent.Length)" -ForegroundColor Cyan
                        Write-Host "Generate Reports (SEARCH): Ticket data JSON preview: $($ticketDataJson.Substring(0, [Math]::Min(200, $ticketDataJson.Length)))..." -ForegroundColor Gray
                    } else {
                        Write-Host "Generate Reports (SEARCH): No ticket data to include (TicketNumbers.Count=$($ticketNumbers.Count), FilteredContent empty=$([string]::IsNullOrWhiteSpace($filteredTicketContent)))" -ForegroundColor Yellow
                    }
                    Write-Host "Generate Reports (SEARCH): Final command being sent: $($command.Substring(0, [Math]::Min(500, $command.Length)))..." -ForegroundColor Cyan
                    $reportResponse = Send-CommandToSession -ClientNumber $clientNum -Command $command -TimeoutSeconds 300

                    # Auto-minimize when report generation starts
                    if ($script:clientAuthStates[$clientNum].IsExpanded) {
                        $script:clientAuthStates[$clientNum].IsExpanded = $false
                        $controls.ToggleButton.PerformClick()
                    }

                    if ($reportResponse -like "GENERATE_REPORTS_SUCCESS:*") {
                        $outputPath = ($reportResponse -replace "GENERATE_REPORTS_SUCCESS:", "").Trim()
                        if ($outputPath) {
                            $script:clientReportFolders[$clientNum] = $outputPath
                            $script:clientAuthControls[$clientNum].ViewReportsButton.Visible = $true
                            $script:clientAuthControls[$clientNum].ViewReportsButton.Enabled = $true
                        }
                    }
                    $script:authStatusTextBox.AppendText("Client $($clientNum): Generating reports with user search and ticket data...`r`n")
                    $script:authStatusTextBox.ScrollToCaret()
                    [System.Windows.Forms.Application]::DoEvents()
                    return
                }
            }
            
            # Build GENERATE_REPORTS command
            $command = "GENERATE_REPORTS"
            if ($selectedUsers.Count -gt 0) {
                $usersJson = ($selectedUsers | ConvertTo-Json -Compress)
                $command += "|SelectedUsers:$usersJson"
            }
            # Include ticket data if we have ticket numbers OR ticket content
            if ($ticketNumbers.Count -gt 0 -or -not [string]::IsNullOrWhiteSpace($filteredTicketContent)) {
                # Ensure ticketNumbers is always an array for JSON serialization
                $ticketNumsArray = if ($ticketNumbers -is [array]) { $ticketNumbers } else { @($ticketNumbers) }
                # Force TicketNumbers to be serialized as an array by ensuring it's always an array type
                $ticketDataObj = [PSCustomObject]@{
                    TicketNumbers = [array]$ticketNumsArray
                    TicketContent = [string]$filteredTicketContent
                }
                $ticketDataJson = ($ticketDataObj | ConvertTo-Json -Compress -Depth 10)
                # Verify TicketNumbers is an array in JSON (should be ["1811523"], not "1811523")
                if ($ticketDataJson -notmatch '"TicketNumbers"\s*:\s*\[') {
                    Write-Warning "TicketNumbers was not serialized as an array, fixing..."
                    # Manually fix the JSON if needed
                    $ticketDataJson = $ticketDataJson -replace '"TicketNumbers"\s*:\s*"([^"]+)"', '"TicketNumbers":["$1"]'
                }
                $command += "|TICKET_DATA:$ticketDataJson"
                Write-Host "Generate Reports: Including ticket data - TicketNumbers=$($ticketNumsArray.Count) ($($ticketNumsArray -join ', ')), TicketContent length=$($filteredTicketContent.Length)" -ForegroundColor Cyan
                Write-Host "Generate Reports: Ticket data JSON preview: $($ticketDataJson.Substring(0, [Math]::Min(200, $ticketDataJson.Length)))..." -ForegroundColor Gray
            }
            
            # Send command to worker script
            $this.Enabled = $false
            $this.Text = "Generating..."
            $script:authStatusTextBox.AppendText("Client $($clientNum): Sending generate reports command...`r`n")
            if ($ticketNumbers.Count -gt 0) {
                $script:authStatusTextBox.AppendText("Client $($clientNum): Ticket numbers detected: $(($ticketNumbers | ForEach-Object { "#$_" }) -join ', ')`r`n")
            }
            if (-not [string]::IsNullOrWhiteSpace($filteredTicketContent)) {
                $script:authStatusTextBox.AppendText("Client $($clientNum): Ticket content included ($($filteredTicketContent.Length) characters)`r`n")
            }
            $script:authStatusTextBox.ScrollToCaret()
            [System.Windows.Forms.Application]::DoEvents()
            
            # Use longer timeout for report generation (reports can take several minutes, but we just need GENERATE_REPORTS_STARTED response)
            Write-Host "Generate Reports: Final command being sent: $($command.Substring(0, [Math]::Min(500, $command.Length)))..." -ForegroundColor Cyan
            $reportResponse = Send-CommandToSession -ClientNumber $clientNum -Command $command -TimeoutSeconds 300

            # Auto-minimize when report generation starts
            if ($script:clientAuthStates[$clientNum].IsExpanded) {
                $script:clientAuthStates[$clientNum].IsExpanded = $false
                $controls.ToggleButton.PerformClick()
            }

            if ($reportResponse -like "GENERATE_REPORTS_SUCCESS:*") {
                $outputPath = ($reportResponse -replace "GENERATE_REPORTS_SUCCESS:", "").Trim()
                $script:clientReportFolders[$clientNum] = $outputPath
                if ($script:clientAuthControls[$clientNum].ViewReportsButton) {
                    $script:clientAuthControls[$clientNum].ViewReportsButton.Visible = $true
                    $script:clientAuthControls[$clientNum].ViewReportsButton.Enabled = $true
                }
                $script:authStatusTextBox.AppendText("Client $($clientNum): Report generation completed! Output: $outputPath`r`n")
            } else {
                $script:authStatusTextBox.AppendText("Client $($clientNum): Report generation started.`r`n")
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
            
            # Send CANCEL_AUTH command to worker script to clear sessions and token caches
            Send-CommandToSession -ClientNumber $clientNum -Command "CANCEL_AUTH" -TimeoutSeconds 30 | Out-Null
            
            # Clear all tenant information from state
            $script:clientAuthStates[$clientNum].GraphAuthenticated = $false
            $script:clientAuthStates[$clientNum].ExchangeAuthenticated = $false
            $script:clientAuthStates[$clientNum].TenantName = $null
            $script:clientAuthStates[$clientNum].TenantId = $null
            $script:clientAuthStates[$clientNum].Account = $null
            $script:clientAuthStates[$clientNum].GraphContext = $null
            
            # Clear cache directory for this tenant if it exists
            if ($script:clientCacheDirs -and $script:clientCacheDirs.ContainsKey($clientNum)) {
                $cacheDir = $script:clientCacheDirs[$clientNum]
                if ($cacheDir -and (Test-Path $cacheDir)) {
                    try {
                        Remove-Item -Path $cacheDir -Recurse -Force -ErrorAction SilentlyContinue
                        $script:authStatusTextBox.AppendText("Cleared cache directory for Client $clientNum`r`n")
                    } catch {
                        # Ignore errors clearing cache directory
                    }
                }
                $script:clientCacheDirs.Remove($clientNum)
            }
            
            # Reset UI controls
            $script:clientAuthControls[$clientNum].ClientLabel.Text = "Client $clientNum"
            $script:clientAuthControls[$clientNum].StatusLabel.Text = "Ready for Graph Auth"
            $script:clientAuthControls[$clientNum].StatusLabel.ForeColor = [System.Drawing.Color]::Blue
            $script:clientAuthControls[$clientNum].GraphButton.Enabled = $true
            $script:clientAuthControls[$clientNum].GraphButton.Text = "Graph Auth"
            $script:clientAuthControls[$clientNum].ExchangeButton.Enabled = $false
            $script:clientAuthControls[$clientNum].ExchangeButton.Text = "Exchange Online Auth"
            
            # Hide user filtering controls
            $script:clientAuthControls[$clientNum].UserFilterCheckBox.Visible = $false
            $script:clientAuthControls[$clientNum].UserFilterCheckBox.Enabled = $false
            $script:clientAuthControls[$clientNum].UserFilterCheckBox.Checked = $false
            $script:clientAuthControls[$clientNum].UserSearchTextBox.Visible = $false
            $script:clientAuthControls[$clientNum].UserSearchTextBox.Enabled = $false
            $script:clientAuthControls[$clientNum].UserSearchTextBox.Text = ""
            $script:clientAuthControls[$clientNum].ValidateUsersButton.Visible = $false
            $script:clientAuthControls[$clientNum].ValidateUsersButton.Enabled = $false
            $script:clientAuthControls[$clientNum].UserValidationLabel.Visible = $false
            $script:clientAuthControls[$clientNum].UserValidationLabel.Text = ""
            $script:clientAuthControls[$clientNum].GenerateReportsButton.Visible = $false
            $script:clientAuthControls[$clientNum].GenerateReportsButton.Enabled = $false
            $script:clientAuthControls[$clientNum].GenerateReportsButton.Text = "Generate Reports"
            
            # Hide ticket controls
            $script:clientAuthControls[$clientNum].TicketLabel.Visible = $false
            $script:clientAuthControls[$clientNum].TicketLabel.Enabled = $false
            $script:clientAuthControls[$clientNum].TicketTextBox.Visible = $false
            $script:clientAuthControls[$clientNum].TicketTextBox.Enabled = $false
            $script:clientAuthControls[$clientNum].TicketTextBox.Text = ""
            $script:clientAuthControls[$clientNum].TicketNumbersLabel.Visible = $false
            $script:clientAuthControls[$clientNum].TicketNumbersLabel.Text = ""
            
            # Hide View Reports button
            $script:clientAuthControls[$clientNum].ViewReportsButton.Visible = $false
            $script:clientAuthControls[$clientNum].ViewReportsButton.Enabled = $false
            
            # Clear report folder for this tenant
            if ($script:clientReportFolders.ContainsKey($clientNum)) {
                $script:clientReportFolders.Remove($clientNum)
            }
            
            # Clear ticket data for this tenant
            if ($script:clientTickets.ContainsKey($clientNum)) {
                $script:clientTickets.Remove($clientNum)
            }
            
            # Hide and reset warning label
            if ($script:clientAuthControls[$clientNum].WarningLabel) {
                $script:clientAuthControls[$clientNum].WarningLabel.Visible = $false
                $script:clientAuthControls[$clientNum].WarningLabel.Text = ""
            }
            
            # Clear validated users and search terms for this tenant
            if ($script:clientValidatedUsers.ContainsKey($clientNum)) {
                $script:clientValidatedUsers.Remove($clientNum)
            }
            if ($script:clientSearchTerms.ContainsKey($clientNum)) {
                $script:clientSearchTerms.Remove($clientNum)
            }
            
            $script:authStatusTextBox.AppendText("Client $clientNum authentication reset complete. Ready for full authentication.`r`n")
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
                $script:authPanel.Controls.Remove($controls.BorderPanel)
                $script:authPanel.Controls.Remove($controls.ToggleButton)
                $script:authPanel.Controls.Remove($controls.ClientLabel)
                $script:authPanel.Controls.Remove($controls.StatusLabel)
                $script:authPanel.Controls.Remove($controls.WarningLabel)
                $script:authPanel.Controls.Remove($controls.GraphStatusLabel)
                $script:authPanel.Controls.Remove($controls.ExchangeStatusLabel)
                $script:authPanel.Controls.Remove($controls.OpenReportsButton)
                $script:authPanel.Controls.Remove($controls.RemoveMinimizedButton)
                $script:authPanel.Controls.Remove($controls.GraphButton)
                $script:authPanel.Controls.Remove($controls.ExchangeButton)
                $script:authPanel.Controls.Remove($controls.RemoveButton)
                $script:authPanel.Controls.Remove($controls.ResetButton)
                $script:authPanel.Controls.Remove($controls.UserFilterCheckBox)
                $script:authPanel.Controls.Remove($controls.UserSearchTextBox)
                $script:authPanel.Controls.Remove($controls.ValidateUsersButton)
                $script:authPanel.Controls.Remove($controls.UserValidationLabel)
                $script:authPanel.Controls.Remove($controls.GenerateReportsButton)
                $script:authPanel.Controls.Remove($controls.TicketLabel)
                $script:authPanel.Controls.Remove($controls.TicketTextBox)
                $script:authPanel.Controls.Remove($controls.TicketNumbersLabel)
                $script:authPanel.Controls.Remove($controls.ViewReportsButton)
                
                # Remove from state dictionaries
                $script:clientAuthStates.Remove($clientNum)
                $script:clientAuthControls.Remove($clientNum)
                if ($script:clientTickets.ContainsKey($clientNum)) {
                    $script:clientTickets.Remove($clientNum)
                }
                if ($script:clientReportFolders.ContainsKey($clientNum)) {
                    $script:clientReportFolders.Remove($clientNum)
                }
                if ($script:clientReportFolders.ContainsKey($clientNum)) {
                    $script:clientReportFolders.Remove($clientNum)
                }
                
                # Recalculate positions for remaining tenants
                Update-TenantPositions
                
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

    # Expand All button click handler
    $expandAllBtn.add_Click({
        foreach ($clientNum in $script:clientAuthControls.Keys) {
            if (-not $script:clientAuthStates[$clientNum].IsExpanded) {
                $controls = $script:clientAuthControls[$clientNum]
                if ($controls -and $controls.ToggleButton) {
                    # Don't set IsExpanded first - let the toggle button handler toggle it
                    $controls.ToggleButton.PerformClick()
                }
            }
        }
    })

    # Collapse All button click handler
    $collapseAllBtn.add_Click({
        foreach ($clientNum in $script:clientAuthControls.Keys) {
            if ($script:clientAuthStates[$clientNum].IsExpanded) {
                $controls = $script:clientAuthControls[$clientNum]
                if ($controls -and $controls.ToggleButton) {
                    # Don't set IsExpanded first - let the toggle button handler toggle it
                    $controls.ToggleButton.PerformClick()
                }
            }
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
        try {
            # Check if form is still valid before processing
            if (-not $authConsoleForm -or $authConsoleForm.IsDisposed) {
                if ($statusUpdateTimer) {
                    $statusUpdateTimer.Stop()
                }
                return
            }
            
            foreach ($clientNum in $script:clientAuthControls.Keys) {
                try {
                    $statusFile = Join-Path $script:tempDir "Client${clientNum}_Status.txt"
                    if (Test-Path $statusFile) {
                        # Read last few lines of status file (read more to catch warnings)
                        $statusLines = Get-Content $statusFile -Tail 15 -ErrorAction SilentlyContinue
                        if ($statusLines) {
                            $latestStatus = $statusLines | Select-Object -Last 1
                            # Extract just the message part (after timestamp)
                            if ($latestStatus -match '\]\s+(.+)') {
                                $statusMessage = $matches[1]
                                $controls = $script:clientAuthControls[$clientNum]
                                if ($controls -and $controls.StatusLabel -and -not $controls.StatusLabel.IsDisposed) {
                                    # Check for sign-in logs license warning in status file
                                    $signInLogsWarning = $false
                                    $warningText = ""
                                    foreach ($line in $statusLines) {
                                        if ($line -match 'License required.*Sign-in logs|Azure AD Premium.*Sign-in logs|Sign-in logs require.*Premium|free tenants.*limited.*7 days|WARNING.*License required.*Sign-in') {
                                            $signInLogsWarning = $true
                                            # Extract the warning message
                                            if ($line -match '\]\s+(.+)') {
                                                $warningText = $matches[1]
                                            } else {
                                                $warningText = "Sign-in logs require Azure AD Premium license - pull manually"
                                            }
                                            break
                                        }
                                    }
                                    
                                    # Show/hide warning label based on license warning
                                    if ($signInLogsWarning -and $controls.WarningLabel -and -not $controls.WarningLabel.IsDisposed) {
                                        try {
                                            if (-not $controls.WarningLabel.Visible -or $controls.WarningLabel.Text -ne "⚠ WARNING: $warningText") {
                                                $controls.WarningLabel.Text = "⚠ WARNING: Sign-in logs require Azure AD Premium license - pull manually"
                                                $controls.WarningLabel.ForeColor = [System.Drawing.Color]::Orange
                                                $controls.WarningLabel.Visible = $true
                                            }
                                        } catch {}
                                    }
                                    
                                    # Check if worker script is ready and enable Graph Auth button if needed
                                    # Wait for "Command polling loop started" to ensure the loop is actually running
                                    if ($statusMessage -match 'Command polling loop started|Ready!.*Waiting for Graph Auth|Modules imported successfully') {
                                        if ($controls.GraphButton -and -not $controls.GraphButton.IsDisposed -and -not $controls.GraphButton.Enabled) {
                                            try {
                                                # Small delay to ensure polling loop has started
                                                Start-Sleep -Milliseconds 500
                                                $controls.GraphButton.Enabled = $true
                                                $controls.GraphButton.Text = "Graph Auth"
                                                if ($script:authStatusTextBox -and -not $script:authStatusTextBox.IsDisposed) {
                                                    $script:authStatusTextBox.AppendText("Client $clientNum is ready for authentication (detected by status timer).`r`n")
                                                    $script:authStatusTextBox.ScrollToCaret()
                                                }
                                            } catch {}
                                        }
                                    }
                                    
                                    # Check for report generation completion
                                    $responseFile = Join-Path $script:commandDir "Client${clientNum}_Response.txt"
                                    if (Test-Path $responseFile) {
                                        try {
                                            $responseContent = Get-Content $responseFile -Raw -ErrorAction SilentlyContinue | ForEach-Object { $_.Trim() }
                                            if ($responseContent -match '^GENERATE_REPORTS_SUCCESS:(.+)$') {
                                                $reportFolder = $matches[1].Trim()
                                                if (-not [string]::IsNullOrWhiteSpace($reportFolder) -and (Test-Path $reportFolder)) {
                                                    # Store report folder and show View Reports button
                                                    $script:clientReportFolders[$clientNum] = $reportFolder
                                                    if ($controls.ViewReportsButton -and -not $controls.ViewReportsButton.IsDisposed) {
                                                        $controls.ViewReportsButton.Visible = $true
                                                        $controls.ViewReportsButton.Enabled = $true
                                                    }
                                                    # Also enable Open Reports button in minimized view
                                                    if ($controls.OpenReportsButton -and -not $controls.OpenReportsButton.IsDisposed) {
                                                        $controls.OpenReportsButton.Enabled = $true
                                                    }
                                                }
                                            }
                                        } catch {
                                            # Ignore errors reading response file
                                        }
                                    }

                                    # Update Graph/Exchange status indicators for minimized view
                                    if ($controls.GraphStatusLabel -and -not $controls.GraphStatusLabel.IsDisposed) {
                                        if ($script:clientAuthStates[$clientNum].GraphAuthenticated) {
                                            $controls.GraphStatusLabel.Text = "Graph: ✓"
                                            $controls.GraphStatusLabel.ForeColor = [System.Drawing.Color]::Green
                                        } else {
                                            $controls.GraphStatusLabel.Text = "Graph: ○"
                                            $controls.GraphStatusLabel.ForeColor = [System.Drawing.Color]::Gray
                                        }
                                    }

                                    if ($controls.ExchangeStatusLabel -and -not $controls.ExchangeStatusLabel.IsDisposed) {
                                        if ($script:clientAuthStates[$clientNum].ExchangeAuthenticated) {
                                            $controls.ExchangeStatusLabel.Text = "Exchange: ✓"
                                            $controls.ExchangeStatusLabel.ForeColor = [System.Drawing.Color]::Green
                                        } else {
                                            $controls.ExchangeStatusLabel.Text = "Exchange: ○"
                                            $controls.ExchangeStatusLabel.ForeColor = [System.Drawing.Color]::Gray
                                        }
                                    }

                                    # Update border panel color based on overall status
                                    if ($controls.BorderPanel -and -not $controls.BorderPanel.IsDisposed) {
                                        $borderColor = [System.Drawing.Color]::Gray  # Default: Not started

                                        if ($script:clientAuthStates[$clientNum].GraphAuthenticated -and $script:clientAuthStates[$clientNum].ExchangeAuthenticated) {
                                            # Both auths complete
                                            if ($statusMessage -match 'error|failed|ERROR|FAILED') {
                                                $borderColor = [System.Drawing.Color]::Red  # Error state
                                            } elseif ($statusMessage -match 'generating|processing|running|starting') {
                                                $borderColor = [System.Drawing.Color]::Orange  # Processing
                                            } elseif ($statusMessage -match 'successful|complete|SUCCESS') {
                                                $borderColor = [System.Drawing.Color]::Green  # Complete
                                            } else {
                                                $borderColor = [System.Drawing.Color]::Green  # Both auths done
                                            }
                                        } elseif ($script:clientAuthStates[$clientNum].GraphAuthenticated -or $script:clientAuthStates[$clientNum].ExchangeAuthenticated) {
                                            # Partial auth
                                            if ($statusMessage -match 'error|failed|ERROR|FAILED') {
                                                $borderColor = [System.Drawing.Color]::Red  # Error state
                                            } else {
                                                $borderColor = [System.Drawing.Color]::Orange  # Partial auth or processing
                                            }
                                        } elseif ($statusMessage -match 'error|failed|ERROR|FAILED') {
                                            $borderColor = [System.Drawing.Color]::Red  # Error state
                                        }

                                        $controls.BorderPanel.BackColor = $borderColor
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
                    }
                } catch {
                    # Silently ignore errors reading status file for individual clients
                }
            }
        } catch {
            # If there's an error in the timer handler, stop the timer to prevent repeated errors
            try {
                if ($statusUpdateTimer) {
                    $statusUpdateTimer.Stop()
                }
            } catch {}
        }
    })
    $statusUpdateTimer.Start()

    # Stop timer when form closes
    $authConsoleForm.add_FormClosed({
        try {
            if ($statusUpdateTimer) {
                if ($statusUpdateTimer.Enabled) {
                    $statusUpdateTimer.Stop()
                }
                # Small delay to ensure timer stops processing
                Start-Sleep -Milliseconds 100
                if ($statusUpdateTimer) {
                    $statusUpdateTimer.Dispose()
                }
            }
        } catch {
            # Silently ignore disposal errors
        }
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
    # Use Show() instead of ShowDialog() since the form was already shown modally
    if (-not $bulkForm.Visible) {
        $bulkForm.Show()
    }
})

# Show the main form
[System.Windows.Forms.Application]::EnableVisualStyles()
$bulkForm.ShowDialog() | Out-Null


