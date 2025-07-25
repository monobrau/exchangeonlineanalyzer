<#
.SYNOPSIS
A PowerShell script with a GUI to analyze Exchange Online inbox rules. Allows selection of mailboxes,
exports to a formatted XLSX file, and includes enhanced mailbox-level forwarding, Inbox delegates, 
and Full Access permissions in the output. Now includes session revocation, transport rules review,
connector review capabilities, MS Graph integration for user sign-in blocking, and sending restrictions management.

.DESCRIPTION
This script provides a Windows Forms interface to:
- Connect to Exchange Online (uses existing session if available, loads mailboxes).
- Automatically attempts to connect to Microsoft Graph after successful Exchange Online connection.
- Auto-detect organization domains from loaded mailboxes and pre-populate the domains field.
- Manually input organization domains and suspicious keywords (with auto-detection assistance).
- Select individual or multiple mailboxes for rule analysis for export.
- Launch a separate window to view, select, and delete inbox rules for a single selected mailbox.
- Retrieve mailbox forwarding settings (with enhanced SmtpAddress extraction), Inbox delegate permissions, and Full Access mailbox permissions.
- Export analysis results to an XLSX file with specific formatting.
- Filename for export uses a prioritized approach for tenant domain.
- Includes disconnect and open last file buttons.
- NEW: Auto-detect organization domains from mailbox UPNs with manual override capability.
- NEW: Manual Microsoft Graph Connect/Disconnect button for better user control.
- NEW: Revoke user sessions for selected accounts.
- NEW: View and review transport rules.
- NEW: View and review Exchange Online connectors.
- NEW: Connects to Microsoft Graph to enable additional user management features.
- NEW: Block or Unblock user sign-in for selected accounts via MS Graph.
- NEW: View if a selected user is on the "Restricted Users" (blocked from sending) list and remove them via MS Graph.
- Does not automatically disconnect from Exchange Online when the script GUI is closed.

.NOTES

Version: 6.3-FIXED-AUTODOMAINS-GRAPHCONTROL (Added automatic domain detection and manual MS Graph connect/disconnect controls)
Requires:
    - PowerShell 5.1+
    - ExchangeOnlineManagement module
    - Microsoft.Graph.Authentication module
    - Microsoft.Graph.Users module
    - Microsoft.Graph.Identity.SignIns module
    - *** Microsoft Excel Installed *** (for XLSX conversion and formatting of the main report)
Permissions: Requires Exchange administrative privileges AND appropriate Azure AD/Microsoft Graph permissions
             (e.g., User.ReadWrite.All, SecurityEvents.ReadWrite.All) for the new features.

.LINK
Install Exchange Module: Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force
Install Graph Modules:
  Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force
  Install-Module Microsoft.Graph.Users -Scope CurrentUser -Force
  Install-Module Microsoft.Graph.Identity.SignIns -Scope CurrentUser -Force

.EXAMPLE
.\Enhanced_Exchange_Analyzer_GUI_v6_FIXED.ps1
#>

# Import all modules with error handling
function Safe-ImportModule($modulePath) {
    try {
        Import-Module $modulePath -Global -ErrorAction Stop
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to import module: $modulePath`nError: $($_.Exception.Message)", "Module Import Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        exit
    }
}
Safe-ImportModule "$PSScriptRoot\Modules\ExchangeOnline.psm1"
Safe-ImportModule "$PSScriptRoot\Modules\GraphOnline.psm1"
Safe-ImportModule "$PSScriptRoot\Modules\MailboxAnalysis.psm1"
Safe-ImportModule "$PSScriptRoot\Modules\TransportRules.psm1"
Safe-ImportModule "$PSScriptRoot\Modules\Connectors.psm1"
Safe-ImportModule "$PSScriptRoot\Modules\SessionRevocation.psm1"
Safe-ImportModule "$PSScriptRoot\Modules\SignInManagement.psm1"
Safe-ImportModule "$PSScriptRoot\Modules\RestrictedSender.psm1"
Safe-ImportModule "$PSScriptRoot\Modules\ExportUtils.psm1"
Safe-ImportModule "$PSScriptRoot\Modules\EntraInvestigator.psm1"

# Function to update Entra tab button states
function UpdateEntraButtonStates {
    $hasPath = -not [string]::IsNullOrWhiteSpace($entraOutputFolderTextBox.Text)
    $checkedCount = 0
    for ($i = 0; $i -lt $entraUserGrid.Rows.Count; $i++) {
        if ($entraUserGrid.Rows[$i].Cells["Select"].Value -eq $true) { $checkedCount++ }
    }
    # Only export buttons require export folder path and selection
    $entraExportSignInLogsButton.Enabled = $hasPath -and ($checkedCount -gt 0)
    $entraExportAuditLogsButton.Enabled = $hasPath -and ($checkedCount -eq 1)
    # View, User Details, and Analyze MFA buttons are always enabled
    $entraViewSignInLogsButton.Enabled = $true
    $entraViewAuditLogsButton.Enabled = $true
    $entraDetailsFetchButton.Enabled = $true
    $entraMfaFetchButton.Enabled = $true
}

# --- Configuration ---
$BaseSuspiciousKeywords = @("invoice", "payment", "password", "confidential", "urgent", "bank", "account", "auto forward", "external", "hidden")
$highlightColorIndexYellow = 6 # Excel ColorIndex for Yellow
$highlightColorIndexLightRed = 38 # Excel ColorIndex for Light Red (Rose)


# Script-level variables
$script:lastExportedXlsxPath = $null 
$script:currentExchangeConnection = $null
$script:allLoadedMailboxUPNs = @() 

# MS Graph related script-level variables
$script:graphConnection = $null
$script:graphConnectionAttempted = $false
$script:requiredGraphModules = @(
    @{Name="Microsoft.Graph.Authentication"; MinVersion="2.0"},
    @{Name="Microsoft.Graph.Users"; MinVersion="2.0"},
    @{Name="Microsoft.Graph.Identity.SignIns"; MinVersion="2.0"}
)
$script:graphScopes = @(
    "User.Read.All",
    "User.ReadWrite.All",
    "SecurityEvents.Read.All",
    "SecurityEvents.ReadWrite.All"
)

# --- GUI Setup ---
Add-Type -AssemblyName System.Windows.Forms; Add-Type -AssemblyName System.Drawing
$mainForm = New-Object System.Windows.Forms.Form; $mainForm.Text = "Exchange Online Analyzer (Enhanced v6.3-FIXED with Auto-Domain Detection & Manual Graph Control)"; $mainForm.Size = New-Object System.Drawing.Size(1100, 900); $mainForm.MinimumSize = New-Object System.Drawing.Size(900, 700); $mainForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen; $mainForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable; $mainForm.MaximizeBox = $true; $mainForm.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi
$statusStrip = New-Object System.Windows.Forms.StatusStrip; $statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel; $statusLabel.Name = "statusLabel"; $statusLabel.Text = "Ready. Connect to Exchange Online."; $statusStrip.Items.Add($statusLabel); $mainForm.Controls.Add($statusStrip)

# --- Main TabControl (fills the form) ---
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Dock = 'Fill'
$mainForm.Controls.Add($tabControl)

# --- Exchange Online Controls Instantiation ---
$connectButton = New-Object System.Windows.Forms.Button
$connectButton.Text = "Connect"
$connectButton.Width = 100

$disconnectButton = New-Object System.Windows.Forms.Button
$disconnectButton.Text = "Disconnect"
$disconnectButton.Width = 100

$userMailboxListLabel = New-Object System.Windows.Forms.Label
$userMailboxListLabel.Text = "Mailboxes:"

$selectAllButton = New-Object System.Windows.Forms.Button
$selectAllButton.Text = "Select All"
$selectAllButton.Width = 100

$deselectAllButton = New-Object System.Windows.Forms.Button
$deselectAllButton.Text = "Deselect All"
$deselectAllButton.Width = 100

$orgDomainsLabel = New-Object System.Windows.Forms.Label
$orgDomainsLabel.Text = "Org Domains:"

$orgDomainsTextBox = New-Object System.Windows.Forms.TextBox
$orgDomainsTextBox.Width = 200

$keywordsLabel = New-Object System.Windows.Forms.Label
$keywordsLabel.Text = "Keywords:"

$keywordsTextBox = New-Object System.Windows.Forms.TextBox
$keywordsTextBox.Width = 200

$outputFolderLabel = New-Object System.Windows.Forms.Label
$outputFolderLabel.Text = "Output Folder:"

$outputFolderTextBox = New-Object System.Windows.Forms.TextBox
$outputFolderTextBox.Width = 200

$browseFolderButton = New-Object System.Windows.Forms.Button
$browseFolderButton.Text = "Browse..."
$browseFolderButton.Width = 100

$getRulesButton = New-Object System.Windows.Forms.Button
$getRulesButton.Text = "Export Rules"
$getRulesButton.Width = 120

$manageRulesButton = New-Object System.Windows.Forms.Button
$manageRulesButton.Text = "Manage Rules"
$manageRulesButton.Width = 120
$manageRulesButton.Enabled = $true

$openFileButton = New-Object System.Windows.Forms.Button
$openFileButton.Text = "Open Last File"
$openFileButton.Width = 120

$blockUserButton = New-Object System.Windows.Forms.Button
$blockUserButton.Text = "Block User"
$blockUserButton.Width = 100
$blockUserButton.Enabled = $true

$unblockUserButton = New-Object System.Windows.Forms.Button
$unblockUserButton.Text = "Unblock User"
$unblockUserButton.Width = 100
$unblockUserButton.Enabled = $true

$revokeSessionsButton = New-Object System.Windows.Forms.Button
$revokeSessionsButton.Text = "Revoke Sessions"
$revokeSessionsButton.Width = 120

$manageRestrictedSendersButton = New-Object System.Windows.Forms.Button
$manageRestrictedSendersButton.Text = "Manage Restricted Senders"
$manageRestrictedSendersButton.Width = 180

# Replace CheckedListBox with DataGridView for mailbox list
$userMailboxGrid = New-Object System.Windows.Forms.DataGridView
$userMailboxGrid.Dock = 'Fill'
$userMailboxGrid.ReadOnly = $false
$userMailboxGrid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$userMailboxGrid.MultiSelect = $true
$userMailboxGrid.AllowUserToAddRows = $false
$userMailboxGrid.AutoGenerateColumns = $false
$userMailboxGrid.RowHeadersVisible = $false
$userMailboxGrid.AllowUserToOrderColumns = $true
$userMailboxGrid.AllowUserToResizeRows = $true
$userMailboxGrid.AllowUserToResizeColumns = $true
$userMailboxGrid.AutoSizeColumnsMode = 'Fill'

# Define columns
$colCheck = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
$colCheck.HeaderText = "Select"
$colCheck.Width = 40
$colCheck.Name = "Select"
$colCheck.ReadOnly = $false
$userMailboxGrid.Columns.Add($colCheck)

$colUPN = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colUPN.HeaderText = "UserPrincipalName"
$colUPN.DataPropertyName = "UserPrincipalName"
$colUPN.Width = 220
$colUPN.ReadOnly = $true
$userMailboxGrid.Columns.Add($colUPN)

$colDisplayName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colDisplayName.HeaderText = "DisplayName"
$colDisplayName.DataPropertyName = "DisplayName"
$colDisplayName.Width = 180
$colDisplayName.ReadOnly = $true
$userMailboxGrid.Columns.Add($colDisplayName)

$colBlocked = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colBlocked.HeaderText = "SignInBlocked"
$colBlocked.DataPropertyName = "SignInBlocked"
$colBlocked.Width = 100
$colBlocked.ReadOnly = $true
$userMailboxGrid.Columns.Add($colBlocked)

$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Width = 200

$exchangeGrid = New-Object System.Windows.Forms.DataGridView
$exchangeGrid.ReadOnly = $true
$exchangeGrid.AllowUserToAddRows = $false
$exchangeGrid.AutoGenerateColumns = $true

# --- Entra ID Investigator Controls Instantiation ---
$entraConnectGraphButton = New-Object System.Windows.Forms.Button
$entraConnectGraphButton.Text = "Connect Entra"
$entraConnectGraphButton.Width = 140

$entraDisconnectGraphButton = New-Object System.Windows.Forms.Button
$entraDisconnectGraphButton.Text = "Disconnect Entra"
$entraDisconnectGraphButton.Width = 140

$entraOutputFolderLabel = New-Object System.Windows.Forms.Label
$entraOutputFolderLabel.Text = "Export Folder:"
$entraOutputFolderTextBox = New-Object System.Windows.Forms.TextBox
$entraOutputFolderTextBox.Width = 300
$entraBrowseFolderButton = New-Object System.Windows.Forms.Button
$entraBrowseFolderButton.Text = "Browse..."
$entraBrowseFolderButton.Width = 100

$entraUserListLabel           = New-Object System.Windows.Forms.Label
$entraUserListLabel.Text      = "Users:"

$entraUserCheckedListBox      = New-Object System.Windows.Forms.CheckedListBox
$entraUserCheckedListBox.Width = 200
$entraUserCheckedListBox.Height = 80

$entraSignInDaysLabel         = New-Object System.Windows.Forms.Label
$entraSignInDaysLabel.Text    = "Sign-in Days:"

$entraSignInDaysUpDown        = New-Object System.Windows.Forms.NumericUpDown
$entraSignInDaysUpDown.Minimum = 1
$entraSignInDaysUpDown.Maximum = 90
$entraSignInDaysUpDown.Value   = 7

$entraSignInExportButton      = New-Object System.Windows.Forms.Button
$entraSignInExportButton.Text = "Fetch Sign-in Logs"
$entraSignInExportButton.Width = 140

$entraSignInExportXlsxButton  = New-Object System.Windows.Forms.Button
$entraSignInExportXlsxButton.Text = "Export Sign-in XLSX"
$entraSignInExportXlsxButton.Width = 140
$entraSignInExportXlsxButton.Enabled = $false

$entraDetailsFetchButton      = New-Object System.Windows.Forms.Button
$entraDetailsFetchButton.Text = "User Details && Roles"
$entraDetailsFetchButton.Width = 140

$entraAuditFetchButton        = New-Object System.Windows.Forms.Button
$entraAuditFetchButton.Text   = "Fetch Audit Logs"
$entraAuditFetchButton.Width = 140

$entraAuditExportXlsxButton   = New-Object System.Windows.Forms.Button
$entraAuditExportXlsxButton.Text = "Export Audit XLSX"
$entraAuditExportXlsxButton.Width = 140
$entraAuditExportXlsxButton.Enabled = $false

$entraMfaFetchButton          = New-Object System.Windows.Forms.Button
$entraMfaFetchButton.Text     = "Analyze MFA"
$entraMfaFetchButton.Width = 120

$entraSignInGrid              = New-Object System.Windows.Forms.DataGridView
$entraSignInGrid.ReadOnly     = $true
$entraSignInGrid.AllowUserToAddRows = $false
$entraSignInGrid.AutoGenerateColumns = $true

$entraAuditGrid               = New-Object System.Windows.Forms.DataGridView
$entraAuditGrid.ReadOnly      = $true
$entraAuditGrid.AllowUserToAddRows = $false
$entraAuditGrid.AutoGenerateColumns = $true

# Instantiate Entra ID Investigator tab buttons before layout
$entraViewSignInLogsButton = New-Object System.Windows.Forms.Button
$entraViewSignInLogsButton.Text = "View Sign-in Logs"
$entraViewSignInLogsButton.Width = 140

$entraViewAuditLogsButton = New-Object System.Windows.Forms.Button
$entraViewAuditLogsButton.Text = "View Audit Logs"
$entraViewAuditLogsButton.Width = 140

$entraExportSignInLogsButton = New-Object System.Windows.Forms.Button
$entraExportSignInLogsButton.Text = "Export Sign-in Logs"
$entraExportSignInLogsButton.Width = 160
$entraExportSignInLogsButton.Enabled = $false

$entraExportAuditLogsButton = New-Object System.Windows.Forms.Button
$entraExportAuditLogsButton.Text = "Export Audit Logs"
$entraExportAuditLogsButton.Width = 160
$entraExportAuditLogsButton.Enabled = $false

$entraOpenLastExportButton = New-Object System.Windows.Forms.Button
$entraOpenLastExportButton.Text = "Open Last Export"
$entraOpenLastExportButton.Width = 140
$entraOpenLastExportButton.Enabled = $true

# --- Exchange Online Tab Layout ---
$exchangeTab = New-Object System.Windows.Forms.TabPage; $exchangeTab.Text = "Exchange Online"

# Top action panel for Connect/Disconnect/Select All/Deselect All/Block/Unblock User/Revoke Sessions/Manage Rules
$topActionPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$topActionPanel.Dock = 'Top'
$topActionPanel.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
$topActionPanel.WrapContents = $false
$topActionPanel.AutoSize = $true
$topActionPanel.Controls.AddRange(@($connectButton, $disconnectButton, $selectAllButton, $deselectAllButton, $blockUserButton, $unblockUserButton, $revokeSessionsButton, $manageRulesButton, $manageRestrictedSendersButton))
$exchangeTab.Controls.Add($topActionPanel)

# Panel for mailbox label and CheckedListBox (fills space)
$mailboxPanel = New-Object System.Windows.Forms.Panel
$mailboxPanel.Dock = 'Fill'
$mailboxPanel.Padding = 0
$mailboxPanel.Margin = 0
$userMailboxListLabel.Dock = 'Top'
$userMailboxGrid.Dock = 'Fill'
$userMailboxGrid.ScrollBars = [System.Windows.Forms.ScrollBars]::Both  # Show both scrollbars
$mailboxPanel.Controls.Add($userMailboxGrid)
$mailboxPanel.Controls.Add($userMailboxListLabel)
$exchangeTab.Controls.Add($mailboxPanel)

# Action buttons panel at the very bottom (full width, 2 rows)
$actionPanel = New-Object System.Windows.Forms.Panel
$actionPanel.Dock = 'Bottom'
$actionPanel.MinimumSize = New-Object System.Drawing.Size(0, 80)
$actionPanel.Height = 80

# Row 1: Output Folder, Browse, Export Rules, Open Last File
$row1 = New-Object System.Windows.Forms.FlowLayoutPanel
$row1.Dock = 'Top'
$row1.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
$row1.WrapContents = $false
$row1.AutoSize = $true
$row1.Controls.AddRange(@($outputFolderLabel, $outputFolderTextBox, $browseFolderButton, $getRulesButton, $openFileButton))

# Row 2: Org Domains, Keywords, Manage Restricted Senders, ProgressBar
$row2 = New-Object System.Windows.Forms.FlowLayoutPanel
$row2.Dock = 'Top'
$row2.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
$row2.WrapContents = $false
$row2.AutoSize = $true
$row2.Controls.AddRange(@($orgDomainsLabel, $orgDomainsTextBox, $keywordsLabel, $keywordsTextBox, $progressBar))

$actionPanel.Controls.Add($row1)
$actionPanel.Controls.Add($row2)

# Remove old actionPanel and add new one
$exchangeTab.Controls.Remove($actionPanel)
$exchangeTab.Controls.Add($actionPanel)

# DataGridView for results (hidden by default, shown when results are present)
$exchangeGrid.Dock = 'Fill'
$exchangeGrid.Visible = $false
$exchangeTab.Controls.Add($exchangeGrid)

# --- Entra ID Investigator Tab Layout (full rewrite for clarity and order) ---
$entraTab = New-Object System.Windows.Forms.TabPage; $entraTab.Text = "Entra ID Investigator"

# User grid
$entraUserGrid = New-Object System.Windows.Forms.DataGridView
$entraUserGrid.Dock = 'Fill'
$entraUserGrid.ReadOnly = $false
$entraUserGrid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$entraUserGrid.MultiSelect = $true
$entraUserGrid.AllowUserToAddRows = $false
$entraUserGrid.AutoGenerateColumns = $false
$entraUserGrid.RowHeadersVisible = $false
$entraUserGrid.AllowUserToOrderColumns = $true
$entraUserGrid.AllowUserToResizeRows = $true
$entraUserGrid.AllowUserToResizeColumns = $true
$entraUserGrid.AutoSizeColumnsMode = 'Fill'

# Define columns
$colEntraCheck = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
$colEntraCheck.HeaderText = "Select"
$colEntraCheck.Name = "Select"
$colEntraCheck.ReadOnly = $false
$entraUserGrid.Columns.Add($colEntraCheck)

$colEntraUPN = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colEntraUPN.HeaderText = "UserPrincipalName"
$colEntraUPN.Name = "UserPrincipalName"
$colEntraUPN.ReadOnly = $true
$entraUserGrid.Columns.Add($colEntraUPN)

$colEntraDisplayName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colEntraDisplayName.HeaderText = "DisplayName"
$colEntraDisplayName.Name = "DisplayName"
$colEntraDisplayName.ReadOnly = $true
$entraUserGrid.Columns.Add($colEntraDisplayName)

$colEntraLicensed = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colEntraLicensed.HeaderText = "Licensed"
$colEntraLicensed.Name = "Licensed"
$colEntraLicensed.ReadOnly = $true
$entraUserGrid.Columns.Add($colEntraLicensed)

# Add MailboxType column to userMailboxGrid
$colMailboxType = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colMailboxType.HeaderText = "MailboxType"
$colMailboxType.DataPropertyName = "MailboxType"
$colMailboxType.Width = 120
$colMailboxType.ReadOnly = $true
$userMailboxGrid.Columns.Add($colMailboxType)

# Top action panel
$entraTopActionPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$entraTopActionPanel.Dock = 'Top'
$entraTopActionPanel.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
$entraTopActionPanel.WrapContents = $false
$entraTopActionPanel.AutoSize = $true
$entraTopActionPanel.Controls.Clear()
$entraTopActionPanel.Controls.AddRange(@(
    $entraConnectGraphButton,
    $entraDisconnectGraphButton,
    $entraViewSignInLogsButton,
    $entraViewAuditLogsButton,
    $entraDetailsFetchButton,
    $entraMfaFetchButton
))

# Bottom panel
$entraBottomPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$entraBottomPanel.Dock = 'Bottom'
$entraBottomPanel.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
$entraBottomPanel.WrapContents = $true
$entraBottomPanel.AutoSize = $true
$entraBottomPanel.MinimumSize = New-Object System.Drawing.Size(0, 80)
$entraBottomPanel.Height = 80
$entraBottomPanel.Controls.Clear()
$entraBottomPanel.Controls.AddRange(@(
    $entraOutputFolderLabel,
    $entraOutputFolderTextBox,
    $entraBrowseFolderButton,
    $entraSelectedPathTextBox,
    $entraExportSignInLogsButton,
    $entraExportAuditLogsButton,
    $entraOpenLastExportButton
))

# Add a read-only textbox to display the selected export path
$entraSelectedPathTextBox = New-Object System.Windows.Forms.TextBox
$entraSelectedPathTextBox.ReadOnly = $true
$entraSelectedPathTextBox.Width = 300
$entraSelectedPathTextBox.Text = ""

# Update the selected path textbox when the folder changes
$entraOutputFolderTextBox.add_TextChanged({
    $entraSelectedPathTextBox.Text = $entraOutputFolderTextBox.Text
    UpdateEntraButtonStates
})

# Update the selected path textbox when Browse is used
$entraBrowseFolderButton.add_Click({
    $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($folderDialog.ShowDialog() -eq 'OK') {
        $entraOutputFolderTextBox.Text = $folderDialog.SelectedPath
        $entraSelectedPathTextBox.Text = $folderDialog.SelectedPath
    }
})

# Add the selected path textbox to the bottom panel after the Browse button
$entraBottomPanel.Controls.Clear()
$entraBottomPanel.Controls.AddRange(@(
    $entraOutputFolderLabel,
    $entraOutputFolderTextBox,
    $entraBrowseFolderButton,
    $entraSelectedPathTextBox,
    $entraExportSignInLogsButton,
    $entraExportAuditLogsButton,
    $entraOpenLastExportButton
))

# Wire up the Browse button for the Entra ID Investigator tab
$entraBrowseFolderButton = New-Object System.Windows.Forms.Button
$entraBrowseFolderButton.Text = "Browse..."
$entraBrowseFolderButton.Width = 100
$entraBrowseFolderButton.add_Click({
    $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($folderDialog.ShowDialog() -eq 'OK') {
        $entraOutputFolderTextBox.Text = $folderDialog.SelectedPath
    }
})

# Main panel for user grid (fills space)
$entraUserPanel = New-Object System.Windows.Forms.Panel
$entraUserPanel.Dock = 'Fill'
$entraUserPanel.Controls.Clear()
$entraUserPanel.Controls.Add($entraUserGrid)

# Add all to the tab in order: top panel, user grid, bottom panel
$entraTab.Controls.Clear()
$entraTab.Controls.Add($entraTopActionPanel)
$entraTab.Controls.Add($entraUserPanel)
$entraTab.Controls.Add($entraBottomPanel)
$entraTab.Padding = 0
$entraTab.Margin = 0
$entraTab.Dock = 'Fill'

# Populate Entra user grid after Graph authentication
$entraConnectGraphButton.add_Click({
    $statusLabel.Text = "Connecting to Microsoft Graph..."
    $mainForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $entraConnectGraphButton.Enabled = $false
    $entraSignInExportButton.Enabled = $false; $entraDetailsFetchButton.Enabled = $false; $entraAuditFetchButton.Enabled = $false; $entraMfaFetchButton.Enabled = $false

    if (Connect-EntraGraph) {
        try {
            $users = Get-EntraUsers
            $entraUserGrid.Rows.Clear()
            foreach ($u in $users) {
                try {
                    $userDetails = Get-MgUser -UserId $u.UserPrincipalName -Property AssignedLicenses
                    $isLicensed = $userDetails.AssignedLicenses.Count -gt 0
                } catch {
                    $isLicensed = $false
                }
                $licensedText = if ($isLicensed) { "Licensed" } else { "Unlicensed" }
                $entraUserGrid.Rows.Add($false, $u.UserPrincipalName, $u.DisplayName, $licensedText)
                UpdateEntraButtonStates
            }
            $statusLabel.Text = "Connected to Microsoft Graph. Users loaded."
            $entraSignInExportButton.Enabled = $true; $entraDetailsFetchButton.Enabled = $true; $entraAuditFetchButton.Enabled = $true; $entraMfaFetchButton.Enabled = $true
            UpdateEntraButtonStates
        } catch {
            $statusLabel.Text = "Failed to load users: $($_.Exception.Message)"
            [System.Windows.Forms.MessageBox]::Show("Failed to load users: $($_.Exception.Message)", "Graph Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    } else {
        $statusLabel.Text = "Failed to connect to Microsoft Graph."
    }
    $mainForm.Cursor = [System.Windows.Forms.Cursors]::Default
    $entraConnectGraphButton.Enabled = $true
})
$entraSignInExportButton.add_Click({
    $entraUserGrid.EndEdit()
    Write-Host 'EntraUserGrid Columns:'
    foreach ($col in $entraUserGrid.Columns) { Write-Host $col.Name }
    $selectedUpns = @()
    for ($i = 0; $i -lt $entraUserGrid.Rows.Count; $i++) {
        if ($entraUserGrid.Rows[$i].Cells["Select"].Value -eq $true) {
            $upn = $entraUserGrid.Rows[$i].Cells["UserPrincipalName"].Value
            if (-not [string]::IsNullOrWhiteSpace($upn)) { $selectedUpns += $upn }
        }
    }
    Write-Host "Selected UPNs: $($selectedUpns -join ', ')"
    if ($selectedUpns.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Select at least one user with a valid UPN.", "No Users Selected", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning); return
    }
    $statusLabel.Text = "Fetching sign-in logs..."; $mainForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $days = $entraSignInDaysUpDown.Value
    $outputFolder = $entraOutputFolderTextBox.Text
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $csvFilePath = Join-Path $outputFolder "EntraSignInLogs_$timestamp.csv"
    $xlsxFilePath = Join-Path $outputFolder "EntraSignInLogs_$timestamp.xlsx"
    try {
        $allLogs = Get-EntraSignInLogs -UserPrincipalNames $selectedUpns -Days $days
        $entraSignInGrid.DataSource = $null
        if (-not $allLogs -or $allLogs.Count -eq 0) {
            $statusLabel.Text = "No sign-in logs found."; $mainForm.Cursor = [System.Windows.Forms.Cursors]::Default
            [System.Windows.Forms.MessageBox]::Show("No sign-in logs found for selected users.", "No Logs", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information); $entraSignInExportXlsxButton.Enabled = $false; return
        }
        $entraSignInGrid.DataSource = $allLogs
        $entraSignInExportXlsxButton.Tag = $allLogs
        $entraSignInExportXlsxButton.Enabled = $true
        $statusLabel.Text = "Sign-in logs loaded."
    } catch {
        $statusLabel.Text = "Error during sign-in log fetch: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Error during sign-in log fetch: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        $entraSignInExportXlsxButton.Enabled = $false
    } finally { $mainForm.Cursor = [System.Windows.Forms.Cursors]::Default }
})
$entraSignInExportXlsxButton.add_Click({
    $allLogs = $entraSignInExportXlsxButton.Tag
    if (-not $allLogs -or $allLogs.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("No sign-in logs to export.", "No Data", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information); return }
    $outputFolder = $entraOutputFolderTextBox.Text
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $csvFilePath = Join-Path $outputFolder "EntraSignInLogs_$timestamp.csv"
    $xlsxFilePath = Join-Path $outputFolder "EntraSignInLogs_$timestamp.xlsx"
    $allLogs | Export-Csv -Path $csvFilePath -NoTypeInformation -Encoding UTF8
    if (Format-InboxRuleXlsx -CsvPath $csvFilePath -XlsxPath $xlsxFilePath) {
        try { Remove-Item $csvFilePath -Force } catch {}
        $entraOpenFileButton.Tag = $xlsxFilePath
        $entraOpenFileButton.Enabled = $true
        $script:lastExportedXlsxPath = $xlsxFilePath # Update the script-level variable
        [System.Windows.Forms.MessageBox]::Show("Exported and formatted sign-in logs to:\n$xlsxFilePath", "XLSX Export OK", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    } else {
        [System.Windows.Forms.MessageBox]::Show("CSV Exported to:\n$csvFilePath\n\nXLSX/Format FAILED. Check Excel install & console.", "XLSX Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    }
})
$entraDetailsFetchButton.add_Click({
    $entraUserGrid.EndEdit()
    Write-Host 'EntraUserGrid Columns:'
    foreach ($col in $entraUserGrid.Columns) { Write-Host $col.Name }
    $selectedUpns = @()
    for ($i = 0; $i -lt $entraUserGrid.Rows.Count; $i++) {
        if ($entraUserGrid.Rows[$i].Cells["Select"].Value -eq $true) {
            $upn = $entraUserGrid.Rows[$i].Cells["UserPrincipalName"].Value
            if (-not [string]::IsNullOrWhiteSpace($upn)) { $selectedUpns += $upn }
        }
    }
    Write-Host "Selected UPNs: $($selectedUpns -join ', ')"
    if ($selectedUpns.Count -ne 1) {
        [System.Windows.Forms.MessageBox]::Show("Select exactly one user with a valid UPN.", "Select One User", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information); return
    }
    $upn = $selectedUpns[0]
    $statusLabel.Text = "Fetching user details..."; $mainForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        $result = Get-EntraUserDetailsAndRoles -UserPrincipalName $upn
        if ($result.User) {
            $details = "User Principal Name: $($result.User.UserPrincipalName)`r`nDisplay Name: $($result.User.DisplayName)`r`nAccount Enabled: $($result.User.AccountEnabled)`r`nLast Password Change: $($result.User.LastPasswordChangeDateTime)`r`n" +
                "-----------------------------`r`nRoles:`r`n" +
                ($result.Roles.Count -gt 0 ? ($result.Roles -join "`r`n") : "None") +
                "`r`n-----------------------------`r`nGroups:`r`n" +
                ($result.Groups.Count -gt 0 ? ($result.Groups -join "`r`n") : "None")
            $form = New-Object System.Windows.Forms.Form
            $form.Text = "User Details && Roles"
            $form.Size = New-Object System.Drawing.Size(600, 500)
            $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
            $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
            $form.MaximizeBox = $true
            $textbox = New-Object System.Windows.Forms.TextBox
            $textbox.Multiline = $true
            $textbox.ReadOnly = $true
            $textbox.ScrollBars = 'Both'
            $textbox.Dock = 'Fill'
            $textbox.Font = New-Object System.Drawing.Font('Consolas', 10)
            $textbox.Text = $details
            $form.Controls.Add($textbox)
            $form.ShowDialog($mainForm)
            $form.Dispose()
        } else {
            [System.Windows.Forms.MessageBox]::Show("Error: $($result.Error)", "User Details & Roles Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error fetching user details: $($_.Exception.Message)", "User Details & Roles Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    } finally { $mainForm.Cursor = [System.Windows.Forms.Cursors]::Default }
})
$entraAuditFetchButton.add_Click({
    $entraUserGrid.EndEdit()
    Write-Host 'EntraUserGrid Columns:'
    foreach ($col in $entraUserGrid.Columns) { Write-Host $col.Name }
    $selectedUpns = @()
    for ($i = 0; $i -lt $entraUserGrid.Rows.Count; $i++) {
        if ($entraUserGrid.Rows[$i].Cells["Select"].Value -eq $true) {
            $upn = $entraUserGrid.Rows[$i].Cells["UserPrincipalName"].Value
            if (-not [string]::IsNullOrWhiteSpace($upn)) { $selectedUpns += $upn }
        }
    }
    Write-Host "Selected UPNs: $($selectedUpns -join ', ')"
    if ($selectedUpns.Count -ne 1) {
        [System.Windows.Forms.MessageBox]::Show("Select exactly one user with a valid UPN.", "Select One User", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information); return
    }
    $upn = $selectedUpns[0]
    $days = $entraSignInDaysUpDown.Value
    $statusLabel.Text = "Fetching audit logs..."; $mainForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        $logs = Get-EntraUserAuditLogs -UserPrincipalName $upn -Days $days
        $entraAuditGrid.DataSource = $null
        if ($logs -and $logs.Count -gt 0) {
            $entraAuditGrid.DataSource = $logs
            $entraAuditExportXlsxButton.Tag = $logs
            $entraAuditExportXlsxButton.Enabled = $true
            $statusLabel.Text = "Audit logs loaded."
        } else {
            [System.Windows.Forms.MessageBox]::Show("No audit logs found for $upn.", "Audit Logs", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            $entraAuditExportXlsxButton.Enabled = $false
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error fetching audit logs: $($_.Exception.Message)", "Audit Logs Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        $entraAuditExportXlsxButton.Enabled = $false
    } finally { $mainForm.Cursor = [System.Windows.Forms.Cursors]::Default }
})
$entraAuditExportXlsxButton.add_Click({
    $logs = $entraAuditExportXlsxButton.Tag
    if (-not $logs -or $logs.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("No audit logs to export.", "No Data", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information); return }
    $outputFolder = $entraOutputFolderTextBox.Text
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $csvFilePath = Join-Path $outputFolder "EntraAuditLogs_$timestamp.csv"
    $xlsxFilePath = Join-Path $outputFolder "EntraAuditLogs_$timestamp.xlsx"
    $logs | Export-Csv -Path $csvFilePath -NoTypeInformation -Encoding UTF8
    if (Format-InboxRuleXlsx -CsvPath $csvFilePath -XlsxPath $xlsxFilePath) {
        try { Remove-Item $csvFilePath -Force } catch {}
        $entraOpenFileButton.Tag = $xlsxFilePath
        $entraOpenFileButton.Enabled = $true
        $script:lastExportedXlsxPath = $xlsxFilePath # Update the script-level variable
        [System.Windows.Forms.MessageBox]::Show("Exported and formatted audit logs to:\n$xlsxFilePath", "XLSX Export OK", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    } else {
        [System.Windows.Forms.MessageBox]::Show("CSV Exported to:\n$csvFilePath\n\nXLSX/Format FAILED. Check Excel install & console.", "XLSX Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    }
})
$entraMfaFetchButton.add_Click({
    $entraUserGrid.EndEdit()
    Write-Host 'EntraUserGrid Columns:'
    foreach ($col in $entraUserGrid.Columns) { Write-Host $col.Name }
    $selectedUpns = @()
    for ($i = 0; $i -lt $entraUserGrid.Rows.Count; $i++) {
        if ($entraUserGrid.Rows[$i].Cells["Select"].Value -eq $true) {
            $upn = $entraUserGrid.Rows[$i].Cells["UserPrincipalName"].Value
            if (-not [string]::IsNullOrWhiteSpace($upn)) { $selectedUpns += $upn }
        }
    }
    Write-Host "Selected UPNs: $($selectedUpns -join ', ')"
    if ($selectedUpns.Count -ne 1) {
        [System.Windows.Forms.MessageBox]::Show("Select exactly one user with a valid UPN.", "Select One User", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information); return
    }
    $upn = $selectedUpns[0]
    $statusLabel.Text = "Analyzing MFA..."; $mainForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        $result = Get-EntraUserMfaStatus -UserPrincipalName $upn
        $details = "MFA Analysis for: $upn`r`n" +
            "-----------------------------`r`nPer-User MFA: $($result.PerUserMfa.Details)`r`nSecurity Defaults: $($result.SecurityDefaults.Details)`r`nConditional Access: $($result.ConditionalAccess.Details)`r`n-----------------------------`r`nOverall Status: $($result.OverallStatus)"
        $form = New-Object System.Windows.Forms.Form
        $form.Text = "MFA Analysis"
        $form.Size = New-Object System.Drawing.Size(600, 400)
        $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
        $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
        $form.MaximizeBox = $true
        $textbox = New-Object System.Windows.Forms.TextBox
        $textbox.Multiline = $true
        $textbox.ReadOnly = $true
        $textbox.ScrollBars = 'Both'
        $textbox.Dock = 'Fill'
        $textbox.Font = New-Object System.Drawing.Font('Consolas', 10)
        $textbox.Text = $details
        $form.Controls.Add($textbox)
        $form.ShowDialog($mainForm)
        $form.Dispose()
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error analyzing MFA: $($_.Exception.Message)", "MFA Analysis Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    } finally { $mainForm.Cursor = [System.Windows.Forms.Cursors]::Default }
})

# --- Export Sign-in Logs button: fetch and export in one click ---
$entraExportSignInLogsButton.add_Click({
    $entraUserGrid.EndEdit()
    $selectedUpns = @()
    for ($i = 0; $i -lt $entraUserGrid.Rows.Count; $i++) {
        if ($entraUserGrid.Rows[$i].Cells["Select"].Value -eq $true) {
            $upn = $entraUserGrid.Rows[$i].Cells["UserPrincipalName"].Value
            if (-not [string]::IsNullOrWhiteSpace($upn)) { $selectedUpns += $upn }
        }
    }
    if ($selectedUpns.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Select at least one user with a valid UPN.", "No Users Selected", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning); return
    }
    $statusLabel.Text = "Fetching and exporting sign-in logs..."; $mainForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $days = $entraSignInDaysUpDown.Value
    $outputFolder = $entraOutputFolderTextBox.Text
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $csvFilePath = Join-Path $outputFolder "EntraSignInLogs_$timestamp.csv"
    $xlsxFilePath = Join-Path $outputFolder "EntraSignInLogs_$timestamp.xlsx"
    try {
        $allLogs = Get-EntraSignInLogs -UserPrincipalNames $selectedUpns -Days $days
        if (-not $allLogs -or $allLogs.Count -eq 0) {
            $statusLabel.Text = "No sign-in logs found."; $mainForm.Cursor = [System.Windows.Forms.Cursors]::Default
            [System.Windows.Forms.MessageBox]::Show("No sign-in logs found for selected users.", "No Logs", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information); return
        }
        $allLogs | Export-Csv -Path $csvFilePath -NoTypeInformation -Encoding UTF8
        if (Format-InboxRuleXlsx -CsvPath $csvFilePath -XlsxPath $xlsxFilePath) {
            try { Remove-Item $csvFilePath -Force } catch {}
            $script:lastExportedXlsxPath = $xlsxFilePath
            $statusLabel.Text = "Exported and formatted sign-in logs to $xlsxFilePath"
            [System.Windows.Forms.MessageBox]::Show("Exported and formatted sign-in logs to:\n$xlsxFilePath", "XLSX Export OK", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        } else {
            [System.Windows.Forms.MessageBox]::Show("CSV Exported to:\n$csvFilePath\n\nXLSX/Format FAILED. Check Excel install & console.", "XLSX Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
    } catch {
        $statusLabel.Text = "Error during sign-in log export: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Error during sign-in log export: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    } finally { $mainForm.Cursor = [System.Windows.Forms.Cursors]::Default }
})

# --- Export Audit Logs button: fetch and export in one click ---
$entraExportAuditLogsButton.add_Click({
    $entraUserGrid.EndEdit()
    $selectedUpns = @()
    for ($i = 0; $i -lt $entraUserGrid.Rows.Count; $i++) {
        if ($entraUserGrid.Rows[$i].Cells["Select"].Value -eq $true) {
            $upn = $entraUserGrid.Rows[$i].Cells["UserPrincipalName"].Value
            if (-not [string]::IsNullOrWhiteSpace($upn)) { $selectedUpns += $upn }
        }
    }
    if ($selectedUpns.Count -ne 1) {
        [System.Windows.Forms.MessageBox]::Show("Select exactly one user with a valid UPN.", "Select One User", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information); return
    }
    $upn = $selectedUpns[0]
    $days = $entraSignInDaysUpDown.Value
    $statusLabel.Text = "Fetching and exporting audit logs..."; $mainForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $outputFolder = $entraOutputFolderTextBox.Text
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $csvFilePath = Join-Path $outputFolder "EntraAuditLogs_$timestamp.csv"
    $xlsxFilePath = Join-Path $outputFolder "EntraAuditLogs_$timestamp.xlsx"
    try {
        $logs = Get-EntraUserAuditLogs -UserPrincipalName $upn -Days $days
        if (-not $logs -or $logs.Count -eq 0) {
            $statusLabel.Text = "No audit logs found."; $mainForm.Cursor = [System.Windows.Forms.Cursors]::Default
            [System.Windows.Forms.MessageBox]::Show("No audit logs found for $upn.", "Audit Logs", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information); return
        }
        $logs | Export-Csv -Path $csvFilePath -NoTypeInformation -Encoding UTF8
        if (Format-InboxRuleXlsx -CsvPath $csvFilePath -XlsxPath $xlsxFilePath) {
            try { Remove-Item $csvFilePath -Force } catch {}
            $script:lastExportedXlsxPath = $xlsxFilePath
            $statusLabel.Text = "Exported and formatted audit logs to $xlsxFilePath"
            [System.Windows.Forms.MessageBox]::Show("Exported and formatted audit logs to:\n$xlsxFilePath", "XLSX Export OK", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        } else {
            [System.Windows.Forms.MessageBox]::Show("CSV Exported to:\n$csvFilePath\n\nXLSX/Format FAILED. Check Excel install & console.", "XLSX Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
    } catch {
        $statusLabel.Text = "Error during audit log export: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Error during audit log export: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    } finally { $mainForm.Cursor = [System.Windows.Forms.Cursors]::Default }
})

# Remove or disable the intermediate Export XLSX buttons
$entraSignInExportXlsxButton.Visible = $false
$entraAuditExportXlsxButton.Visible = $false

# --- Exchange Online Tab Event Handlers ---
$connectButton.add_Click({
    $statusLabel.Text = "Connecting to Exchange Online..."; $mainForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        Connect-ExchangeOnline -ErrorAction Stop
        $statusLabel.Text = "Connected. Loading mailboxes..."
        $mailboxes = Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox,SharedMailbox,RoomMailbox,EquipmentMailbox | Select-Object UserPrincipalName, DisplayName, AccountDisabled, IsLicensed, RecipientTypeDetails | Sort-Object UserPrincipalName
        $userMailboxGrid.Rows.Clear()
        $script:allLoadedMailboxUPNs = @()  # Store for domain detection
        foreach ($mbx in $mailboxes) {
            $script:allLoadedMailboxUPNs += $mbx.UserPrincipalName
            try {
                $user = Get-User -Identity $mbx.UserPrincipalName -ErrorAction Stop
                if ($null -ne $user.AccountDisabled) {
                    $signInBlocked = if ($user.AccountDisabled) { "Blocked" } else { "Allowed" }
                } else {
                    $signInBlocked = "Unknown"
                }
            } catch {
                $signInBlocked = "Unknown"
            }
            $rowIdx = $userMailboxGrid.Rows.Add($false, $mbx.UserPrincipalName, $mbx.DisplayName, $signInBlocked, $mbx.RecipientTypeDetails)
        }
        # Auto-detect tenant/org domains from loaded mailboxes
        $detectedDomains = Get-AutoDetectedDomains -MailboxUPNs $script:allLoadedMailboxUPNs
        if ($detectedDomains -and $detectedDomains.Count -gt 0) {
            $orgDomainsTextBox.Text = ($detectedDomains -join ", ")
        } else {
            $orgDomainsTextBox.Text = ""
        }
        # Populate suspicious keywords from $BaseSuspiciousKeywords
        $keywordsTextBox.Text = ($BaseSuspiciousKeywords -join ", ")
        $selectAllButton.Enabled = $true; $deselectAllButton.Enabled = $true; $disconnectButton.Enabled = $true; $connectButton.Enabled = $false
        # Enable/disable action buttons based on selection
        $manageRulesButton.Enabled = $true
        $blockUserButton.Enabled = $false
        $unblockUserButton.Enabled = $false
        # No automatic connect to Entra Graph here, user must click their button
    } catch {
        $statusLabel.Text = "Connection failed: $($_.Exception.Message)"
    } finally { $mainForm.Cursor = [System.Windows.Forms.Cursors]::Default }
})
$disconnectButton.add_Click({
    $statusLabel.Text = "Disconnecting..."; $mainForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try { Disconnect-ExchangeOnline -Confirm:$false } catch {}
    $userMailboxGrid.Rows.Clear(); $selectAllButton.Enabled = $false; $deselectAllButton.Enabled = $false; $disconnectButton.Enabled = $false; $connectButton.Enabled = $true
    $statusLabel.Text = "Disconnected."
            $mainForm.Cursor = [System.Windows.Forms.Cursors]::Default
})
$selectAllButton.add_Click({ for ($i = 0; $i -lt $userMailboxGrid.Rows.Count; $i++) { $userMailboxGrid.Rows[$i].Cells["Select"].Value = $true } })
$deselectAllButton.add_Click({ for ($i = 0; $i -lt $userMailboxGrid.Rows.Count; $i++) { $userMailboxGrid.Rows[$i].Cells["Select"].Value = $false } })
$browseFolderButton.add_Click({ $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog; if ($folderDialog.ShowDialog() -eq 'OK') { $outputFolderTextBox.Text = $folderDialog.SelectedPath } })
$getRulesButton.add_Click({
    $selectedUpns = @()
    for ($i = 0; $i -lt $userMailboxGrid.Rows.Count; $i++) {
        if ($userMailboxGrid.Rows[$i].Cells["Select"].Value -eq $true) {
            $selectedUpns += $userMailboxGrid.Rows[$i].Cells["UserPrincipalName"].Value
        }
    }
    if ($selectedUpns.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Select at least one mailbox.", "No Mailbox Selected", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning); return
    }
    $statusLabel.Text = "Analyzing inbox rules..."; $mainForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $outputFolder = $outputFolderTextBox.Text
    if ([string]::IsNullOrWhiteSpace($outputFolder)) {
        [System.Windows.Forms.MessageBox]::Show("Please select an output folder before analyzing rules.", "Output Folder Required", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        $mainForm.Cursor = [System.Windows.Forms.Cursors]::Default
        return
    }
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $csvFilePath = Join-Path $outputFolder "InboxRules_$timestamp.csv"
    $xlsxFilePath = Join-Path $outputFolder "InboxRules_$timestamp.xlsx"
    $allRuleData = @()
    try {
        foreach ($upn in $selectedUpns) {
            $rules = Get-InboxRule -Mailbox $upn -IncludeHidden -ErrorAction SilentlyContinue
            if ($rules) {
                foreach ($rule in $rules) {
                    $matchedKeywords = @()
                    foreach ($kw in $BaseSuspiciousKeywords) {
                        if ($rule.Name -and $rule.Name -match [regex]::Escape($kw)) {
                            $matchedKeywords += $kw
                        }
                    }
                    $allRuleData += [PSCustomObject]@{
                        MailboxOwner                = $upn
                        RuleName                    = $rule.Name
                        IsEnabled                   = $rule.Enabled
                        Priority                    = $rule.Priority
                        IsHidden                    = $rule.IsHidden
                        IsForwardingExternal        = [bool]($rule.ForwardTo -match '@')
                        IsDeleting                  = $rule.DeleteMessage
                        IsMarkingAsRead             = $rule.MarkAsRead
                        IsMovingToFolder            = [bool]$rule.MoveToFolder
                        MoveToFolderName            = $rule.MoveToFolder
                        SuspiciousKeywordsInName    = ($matchedKeywords -join ', ')
                        Description                 = $rule.Description
                        StopProcessingRules         = $rule.StopProcessingRules
                        Conditions                  = $rule.Conditions
                        Actions                     = $rule.Actions
                        Exceptions                  = $rule.Exceptions
                        RuleID                      = "'$($rule.RuleIdentity)"
                    }
                }
            }
        }
        if ($allRuleData.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("No inbox rules found for selected mailboxes.", "No Rules", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information); return
        }
        $exchangeGrid.DataSource = $null
        $exchangeGrid.DataSource = $allRuleData
        $allRuleData | Export-Csv -Path $csvFilePath -NoTypeInformation -Encoding UTF8
        if (Format-InboxRuleXlsx -CsvPath $csvFilePath -XlsxPath $xlsxFilePath) {
            try { Remove-Item $csvFilePath -Force } catch {}
            $openFileButton.Tag = $xlsxFilePath
            $openFileButton.Enabled = $true
            $statusLabel.Text = "Exported and formatted inbox rules to $xlsxFilePath"
            [System.Windows.Forms.MessageBox]::Show("Exported and formatted inbox rules to:\n$xlsxFilePath", "XLSX Export OK", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        } else {
            $statusLabel.Text = "CSV OK, XLSX/Format Failed."; $openFileButton.Enabled = $false
            [System.Windows.Forms.MessageBox]::Show("CSV Exported to:\n$csvFilePath\n\nXLSX/Format FAILED. Check Excel install & console.", "XLSX Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
    } catch {
        $statusLabel.Text = "Error during analysis: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Error during analysis: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    } finally { $mainForm.Cursor = [System.Windows.Forms.Cursors]::Default }
})
$manageRulesButton.add_Click({
    $checkedRows = @()
    for ($i = 0; $i -lt $userMailboxGrid.Rows.Count; $i++) {
        if ($userMailboxGrid.Rows[$i].Cells["Select"].Value -eq $true) {
            $checkedRows += $userMailboxGrid.Rows[$i]
        }
    }
    if ($checkedRows.Count -ne 1) {
        [System.Windows.Forms.MessageBox]::Show("Select exactly one mailbox to manage rules.", "Select One Mailbox", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information); return
    }
    $upn = $checkedRows[0].Cells[1].Value
    $rulesForm = New-Object System.Windows.Forms.Form
    $rulesForm.Text = "Manage Inbox Rules for $upn"
    $rulesForm.Size = New-Object System.Drawing.Size(900, 500)
    $rulesForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
    $rulesForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
    $rulesForm.MaximizeBox = $true

    # Create a new DataGridView for rules each time
    $rulesGrid = New-Object System.Windows.Forms.DataGridView
    $rulesGrid.Dock = 'Fill'
    $rulesGrid.ReadOnly = $true
    $rulesGrid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $rulesGrid.AutoGenerateColumns = $true
    $rulesGrid.AllowUserToAddRows = $false
    $rulesGrid.AutoSizeColumnsMode = 'Fill'

    # Panel for buttons
    $buttonPanel = New-Object System.Windows.Forms.Panel
    $buttonPanel.Dock = 'Bottom'
    $buttonPanel.Height = 40

    $deleteButton = New-Object System.Windows.Forms.Button
    $deleteButton.Text = "Delete Selected Rule(s)"
    $deleteButton.Dock = 'Left'
    $closeButton = New-Object System.Windows.Forms.Button
    $closeButton.Text = "Close"
    $closeButton.Dock = 'Right'
    $buttonPanel.Controls.Add($deleteButton)
    $buttonPanel.Controls.Add($closeButton)

    $rulesForm.Controls.Add($rulesGrid)
    $rulesForm.Controls.Add($buttonPanel)

    # Load rules
    $rules = Get-InboxRule -Mailbox $upn -IncludeHidden -ErrorAction SilentlyContinue
    if ($rules -and $rules.Count -gt 0) {
        $displayRules = foreach ($rule in $rules) {
            [PSCustomObject]@{
                Name = $rule.Name
                Enabled = $rule.Enabled
                Priority = $rule.Priority
                RuleIdentity = "$($rule.RuleIdentity)"  # Force string to avoid scientific notation
            }
        }
        # Convert to DataTable
        $dt = New-Object System.Data.DataTable
        if ($displayRules.Count -gt 0) {
            $displayRules[0].psobject.Properties.Name | ForEach-Object { [void]$dt.Columns.Add($_) }
            foreach ($row in $displayRules) {
                $dt.Rows.Add($row.psobject.Properties.Value)
            }
        }
        $rulesGrid.DataSource = $dt
        $rulesGrid.DataSource = $dt
        $rulesGrid.AutoSizeColumnsMode = 'Fill'
        foreach ($col in $rulesGrid.Columns) { $col.AutoSizeMode = 'Fill' }
    } else {
        $rulesGrid.DataSource = $null
    }

    $deleteButton.add_Click({
        if (-not $rulesGrid.SelectedRows -or $rulesGrid.SelectedRows.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Select at least one rule to delete.", "No Rule Selected", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning); return
        }
        $selectedNames = @()
        foreach ($row in $rulesGrid.SelectedRows) {
            $selectedNames += $row.Cells["Name"].Value
        }
        $confirm = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to delete the selected rule(s)?\n" + ($selectedNames -join "\n"), "Confirm Delete", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        foreach ($ruleName in $selectedNames) {
            try {
                Remove-InboxRule -Mailbox $upn -Identity $ruleName -Confirm:$false -ErrorAction Stop
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Failed to delete rule: $ruleName`n$($_.Exception.Message)", "Delete Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }
        # Reload rules
        $rules = Get-InboxRule -Mailbox $upn -IncludeHidden -ErrorAction SilentlyContinue
        if ($rules -and $rules.Count -gt 0) {
            $displayRules = foreach ($rule in $rules) {
                [PSCustomObject]@{
                    Name = $rule.Name
                    Enabled = $rule.Enabled
                    Priority = $rule.Priority
                    RuleIdentity = "$($rule.RuleIdentity)"  # Force string to avoid scientific notation
                }
            }
            # Convert to DataTable
            $dt = New-Object System.Data.DataTable
            if ($displayRules.Count -gt 0) {
                $displayRules[0].psobject.Properties.Name | ForEach-Object { [void]$dt.Columns.Add($_) }
                foreach ($row in $displayRules) {
                    $dt.Rows.Add($row.psobject.Properties.Value)
                }
            }
            $rulesGrid.DataSource = $dt
            $rulesGrid.AutoSizeColumnsMode = 'Fill'
            foreach ($col in $rulesGrid.Columns) { $col.AutoSizeMode = 'Fill' }
        } else {
            $rulesGrid.DataSource = $null
        }
    })
    $closeButton.add_Click({ $rulesForm.Close() })
    [void]$rulesForm.ShowDialog($mainForm)
    $rulesForm.Dispose()
})
$openFileButton.add_Click({
    if ($openFileButton.Tag -and (Test-Path $openFileButton.Tag)) {
        try { Invoke-Item -Path $openFileButton.Tag -ErrorAction Stop } catch { [System.Windows.Forms.MessageBox]::Show("Could not open file: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) }
    } else {
        [System.Windows.Forms.MessageBox]::Show("No file exported or file not found.", "No File", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
})
$blockUserButton.add_Click({
    $selectedUPNs = @()
    for ($i = 0; $i -lt $userMailboxGrid.Rows.Count; $i++) {
        if ($userMailboxGrid.Rows[$i].Cells["Select"].Value -eq $true) {
            $selectedUPNs += $userMailboxGrid.Rows[$i].Cells["UserPrincipalName"].Value
        }
    }
    if ($selectedUPNs.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Select at least one mailbox to block sign-in.", "No Mailbox Selected", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning); return
    }
    $confirm = [System.Windows.Forms.MessageBox]::Show("Block sign-in for the following user(s)?\n" + ($selectedUPNs -join "\n"), "Confirm Block", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }
    try {
        Set-UserSignInBlockedState -UserPrincipalNames $selectedUPNs -Blocked $true -StatusLabel $statusLabel -MainForm $mainForm
        [System.Windows.Forms.MessageBox]::Show("Blocked sign-in for selected user(s).", "Block User", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        } catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to block user(s): $($_.Exception.Message)", "Block User Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})
$unblockUserButton.add_Click({
    $selectedUPNs = @()
    for ($i = 0; $i -lt $userMailboxGrid.Rows.Count; $i++) {
        if ($userMailboxGrid.Rows[$i].Cells["Select"].Value -eq $true) {
            $selectedUPNs += $userMailboxGrid.Rows[$i].Cells["UserPrincipalName"].Value
        }
    }
    if ($selectedUPNs.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Select at least one mailbox to unblock sign-in.", "No Mailbox Selected", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning); return
    }
    $confirm = [System.Windows.Forms.MessageBox]::Show("Unblock sign-in for the following user(s)?\n" + ($selectedUPNs -join "\n"), "Confirm Unblock", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }
    try {
        Set-UserSignInBlockedState -UserPrincipalNames $selectedUPNs -Blocked $false -StatusLabel $statusLabel -MainForm $mainForm
        [System.Windows.Forms.MessageBox]::Show("Unblocked sign-in for selected user(s).", "Unblock User", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        } catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to unblock user(s): $($_.Exception.Message)", "Unblock User Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})
$revokeSessionsButton.add_Click({
    $selectedUPNs = @()
    for ($i = 0; $i -lt $userMailboxGrid.Rows.Count; $i++) {
        if ($userMailboxGrid.Rows[$i].Cells["Select"].Value -eq $true) {
            $selectedUPNs += $userMailboxGrid.Rows[$i].Cells["UserPrincipalName"].Value
        }
    }
    if ($selectedUPNs.Count -eq 0) {
        $selectedUPNs = $script:allLoadedMailboxUPNs
    }
    Show-SessionRevocationTool -mainForm $mainForm -statusLabel $statusLabel -allLoadedMailboxUPNs $selectedUPNs
})
$manageRestrictedSendersButton.add_Click({
    $userMailboxGrid.EndEdit()
    $checkedRows = @()
    for ($i = 0; $i -lt $userMailboxGrid.Rows.Count; $i++) {
        if ($userMailboxGrid.Rows[$i].Cells["Select"].Value -eq $true) {
            $checkedRows += $userMailboxGrid.Rows[$i]
        }
    }
    if ($checkedRows.Count -eq 1) {
        $row = $checkedRows[0]
        $upnCell = $row.Cells[1].Value  # Use index 1 for UPN
        $upn = if ($upnCell -ne $null) { $upnCell.ToString().Trim() } else { "" }
    } else {
        $upn = ""
    }
    if ($checkedRows.Count -ne 1 -or [string]::IsNullOrWhiteSpace($upn)) {
        [System.Windows.Forms.MessageBox]::Show("Select exactly one mailbox to manage restricted senders.", "Select One Mailbox", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information); return
    }
    [System.Windows.Forms.MessageBox]::Show("DEBUG: About to call Show-RestrictedSenderManagementDialog for UPN: $upn")
    Show-RestrictedSenderManagementDialog -UserPrincipalName $upn -ParentForm $mainForm -StatusLabelGlobal $statusLabel
})

$userMailboxGrid.add_CellContentClick({
    $mainForm.BeginInvoke([System.Action]{
        $manageRulesButton.Enabled = $true
        $blockUserButton.Enabled = $true
        $unblockUserButton.Enabled = $true
        $checkedCount = 0
        for ($i = 0; $i -lt $userMailboxGrid.Rows.Count; $i++) {
            if ($userMailboxGrid.Rows[$i].Cells["Select"].Value -eq $true) { $checkedCount++ }
        }
    })
})

# --- After all Entra tab buttons and panels are created ---

# Clear and repopulate the top action panel
$entraTopActionPanel.Controls.Clear()
$entraTopActionPanel.Controls.AddRange(@(
    $entraConnectGraphButton,
    $entraDisconnectGraphButton,
    $entraViewSignInLogsButton,
    $entraViewAuditLogsButton,
    $entraDetailsFetchButton,
    $entraMfaFetchButton
))

# Activate View Sign-in Logs button
$entraViewSignInLogsButton.add_Click({
    $entraUserGrid.EndEdit()
    $selectedUpns = @()
    for ($i = 0; $i -lt $entraUserGrid.Rows.Count; $i++) {
        if ($entraUserGrid.Rows[$i].Cells["Select"].Value -eq $true) {
            $upn = $entraUserGrid.Rows[$i].Cells["UserPrincipalName"].Value
            if (-not [string]::IsNullOrWhiteSpace($upn)) { $selectedUpns += $upn }
        }
    }
    if ($selectedUpns.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Select at least one user with a valid UPN.", "No Users Selected", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning); return
    }
    $days = $entraSignInDaysUpDown.Value
    try {
        $allLogs = @(Get-EntraSignInLogs -UserPrincipalNames $selectedUpns -Days $days)
        if (-not $allLogs -or $allLogs.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("No sign-in logs found for selected users.", "No Logs", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information); return
        }
        # Flatten logs for DataGridView
        $data = foreach ($log in $allLogs) {
            [PSCustomObject]@{
                UserPrincipalName = $log.UserPrincipalName
                CreatedDateTime   = $log.CreatedDateTime
                AppDisplayName    = $log.AppDisplayName
                IPAddress         = $log.IPAddress
                Location          = if ($log.Location) { ($log.Location.City + ', ' + $log.Location.State + ', ' + $log.Location.CountryOrRegion) } else { '' }
                Status            = if ($log.Status) { $log.Status.AdditionalDetails } else { '' }
                Device            = if ($log.DeviceDetail) { ($log.DeviceDetail.Browser + ' / ' + $log.DeviceDetail.OperatingSystem) } else { '' }
                RiskLevelAggregated = $log.RiskLevelAggregated
                ConditionalAccessStatus = $log.ConditionalAccessStatus
            }
        }
        # Convert to DataTable for DataGridView
        $dt = New-Object System.Data.DataTable
        if ($data.Count -gt 0) {
            $data[0].psobject.Properties.Name | ForEach-Object { [void]$dt.Columns.Add($_) }
            foreach ($row in $data) {
                $dt.Rows.Add($row.psobject.Properties.Value)
            }
        }
        $popup = New-Object System.Windows.Forms.Form
        $popup.Text = "Sign-in Logs"
        $popup.Size = New-Object System.Drawing.Size(900, 600)
        $popup.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
        $popup.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
        $popup.MaximizeBox = $true
        $grid = New-Object System.Windows.Forms.DataGridView
        $grid.Dock = 'Fill'
        $grid.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
        $grid.ReadOnly = $true
        $grid.AllowUserToAddRows = $false
        $grid.AutoGenerateColumns = $true
        $grid.AutoSizeColumnsMode = 'Fill'
        $grid.MinimumSize = New-Object System.Drawing.Size(800, 400)
        $grid.DataSource = $dt
        $popup.Controls.Add($grid)
        $popup.ShowDialog($mainForm)
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error fetching sign-in logs: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

# Activate View Audit Logs button
$entraViewAuditLogsButton.add_Click({
    $entraUserGrid.EndEdit()
    $selectedUpns = @()
    for ($i = 0; $i -lt $entraUserGrid.Rows.Count; $i++) {
        if ($entraUserGrid.Rows[$i].Cells["Select"].Value -eq $true) {
            $upn = $entraUserGrid.Rows[$i].Cells["UserPrincipalName"].Value
            if (-not [string]::IsNullOrWhiteSpace($upn)) { $selectedUpns += $upn }
        }
    }
    if ($selectedUpns.Count -ne 1) {
        [System.Windows.Forms.MessageBox]::Show("Select exactly one user with a valid UPN.", "Select One User", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information); return
    }
    $upn = $selectedUpns[0]
    $days = $entraSignInDaysUpDown.Value
    try {
        $logs = @(Get-EntraUserAuditLogs -UserPrincipalName $upn -Days $days)
        if (-not $logs -or $logs.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("No audit logs found for $upn.", "Audit Logs", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information); return
        }
        # Flatten logs for DataGridView
        $data = foreach ($log in $logs) {
            [PSCustomObject]@{
                ActivityDisplayName = $log.ActivityDisplayName
                ActivityDateTime    = $log.ActivityDateTime
                InitiatedBy         = if ($log.InitiatedBy -and $log.InitiatedBy.User) { $log.InitiatedBy.User.UserPrincipalName } else { '' }
                TargetResources     = if ($log.TargetResources) { ($log.TargetResources | ForEach-Object { $_.UserPrincipalName }) -join ", " } else { '' }
                Category            = $log.Category
                Result              = $log.Result
                CorrelationId       = $log.CorrelationId
                LoggedByService     = $log.LoggedByService
                OperationType       = $log.OperationType
                UserPrincipalName   = $log.UserPrincipalName
                IPAddress           = $log.IPAddress
            }
        }
        # Convert to DataTable for DataGridView
        $dt = New-Object System.Data.DataTable
        if ($data.Count -gt 0) {
            $data[0].psobject.Properties.Name | ForEach-Object { [void]$dt.Columns.Add($_) }
            foreach ($row in $data) {
                $dt.Rows.Add($row.psobject.Properties.Value)
            }
        }
        $popup = New-Object System.Windows.Forms.Form
        $popup.Text = "Audit Logs"
        $popup.Size = New-Object System.Drawing.Size(900, 600)
        $popup.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
        $popup.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
        $popup.MaximizeBox = $true
        $grid = New-Object System.Windows.Forms.DataGridView
        $grid.Dock = 'Fill'
        $grid.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
        $grid.ReadOnly = $true
        $grid.AllowUserToAddRows = $false
        $grid.AutoGenerateColumns = $true
        $grid.AutoSizeColumnsMode = 'Fill'
        $grid.MinimumSize = New-Object System.Drawing.Size(800, 400)
        $grid.DataSource = $dt
        $popup.Controls.Add($grid)
        $popup.ShowDialog($mainForm)
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error fetching audit logs: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

# Clear and repopulate the bottom panel
$entraBottomPanel.Controls.Clear()
$entraBottomPanel.Controls.AddRange(@(
    $entraOutputFolderLabel,
    $entraOutputFolderTextBox,
    $entraBrowseFolderButton,
    $entraSelectedPathTextBox,
    $entraExportSignInLogsButton,
    $entraExportAuditLogsButton,
    $entraOpenLastExportButton
))

# Add tabs to the tab control
$tabControl.TabPages.Add($exchangeTab)
$tabControl.TabPages.Add($entraTab)

# Add Help tab after other tabs
$helpTab = New-Object System.Windows.Forms.TabPage
$helpTab.Text = "Help"
$helpTextBox = New-Object System.Windows.Forms.TextBox
$helpTextBox.Multiline = $true
$helpTextBox.ReadOnly = $true
$helpTextBox.ScrollBars = 'Both'
$helpTextBox.Dock = 'Fill'
$helpTextBox.Font = New-Object System.Drawing.Font('Consolas', 10)
$helpTextBox.Text = Get-Content "$PSScriptRoot\readme.md" -Raw
$helpTab.Controls.Add($helpTextBox)
$tabControl.TabPages.Add($helpTab)

# Set Entra user grid column read-only properties
$entraUserGrid.ReadOnly = $false
$colEntraCheck.ReadOnly = $false
$colEntraUPN.ReadOnly = $true
$colEntraDisplayName.ReadOnly = $true
$colEntraLicensed.ReadOnly = $true

# --- Show Form ---
# Remove all auto-connect logic from the form's Shown event
$mainForm.Add_Shown({ $mainForm.Activate() })
[void]$mainForm.ShowDialog()

# --- Script End ---
Write-Host "Script finished."
# No automatic disconnect on GUI close. User must use the "Disconnect" button.
# if ($script:currentExchangeConnection) { Write-Host "Disconnecting from Exchange Online..."; Disconnect-ExchangeOnline -Confirm:$false -EA SilentlyContinue }
# if ($script:graphConnection) { Write-Host "Disconnecting from Microsoft Graph..."; Disconnect-MgGraph -EA SilentlyContinue }

# --- Open Last Export button event handler ---
$entraOpenLastExportButton.add_Click({
    if ($script:lastExportedXlsxPath) {
        if (Test-Path $script:lastExportedXlsxPath) {
            try {
                $statusLabel.Text = "Opening: $script:lastExportedXlsxPath"
                Invoke-Item -Path $script:lastExportedXlsxPath -ErrorAction Stop
            } catch {
                $statusLabel.Text = "Failed to open: $script:lastExportedXlsxPath"
                [System.Windows.Forms.MessageBox]::Show("Could not open file: $($_.Exception.Message)\nPath: $script:lastExportedXlsxPath", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        } else {
            $statusLabel.Text = "File not found: $script:lastExportedXlsxPath"
            [System.Windows.Forms.MessageBox]::Show("No file exported or file not found.\nPath: $script:lastExportedXlsxPath", "No File", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
    } else {
        $statusLabel.Text = "No export path set."
        [System.Windows.Forms.MessageBox]::Show("No file exported yet.", "No File", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
})

# After every successful export, ensure the Open Last Export button is enabled
# (This is already handled by setting $script:lastExportedXlsxPath, but reinforce if needed)

# --- Disconnect Entra button event handler ---
$entraDisconnectGraphButton.add_Click({
    try {
        Disconnect-MgGraph -ErrorAction Stop
        $statusLabel.Text = "Disconnected from Microsoft Graph."
    } catch {
        $statusLabel.Text = "Error disconnecting from Microsoft Graph: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Error disconnecting from Microsoft Graph: $($_.Exception.Message)", "Disconnect Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

# Configure grids to auto-expand horizontally
$userMailboxGrid.AutoSizeColumnsMode = 'Fill'
$entraUserGrid.AutoSizeColumnsMode = 'Fill'

# Add a catch-all event to always enable the button after any grid change
$userMailboxGrid.add_SelectionChanged({ $manageRulesButton.Enabled = $true })
$userMailboxGrid.add_CellValueChanged({ $manageRulesButton.Enabled = $true })