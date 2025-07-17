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
Author: Gemini (Enhanced) - FIXED VERSION with Auto-Domain Detection & Manual Graph Control
Date: 2025-06-03
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

# Import all modules
Import-Module "$PSScriptRoot\Modules\ExchangeOnline.psm1" -Global
Import-Module "$PSScriptRoot\Modules\GraphOnline.psm1" -Global
Import-Module "$PSScriptRoot\Modules\MailboxAnalysis.psm1" -Global
Import-Module "$PSScriptRoot\Modules\TransportRules.psm1" -Global
Import-Module "$PSScriptRoot\Modules\Connectors.psm1" -Global
Import-Module "$PSScriptRoot\Modules\SessionRevocation.psm1" -Global
Import-Module "$PSScriptRoot\Modules\SignInManagement.psm1" -Global
Import-Module "$PSScriptRoot\Modules\RestrictedSender.psm1" -Global
Import-Module "$PSScriptRoot\Modules\ExportUtils.psm1" -Global

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
$mainForm = New-Object System.Windows.Forms.Form; $mainForm.Text = "Exchange Online Analyzer (Enhanced v6.3-FIXED with Auto-Domain Detection & Manual Graph Control)"; $mainForm.Size = New-Object System.Drawing.Size(700, 950); $mainForm.MinimumSize = New-Object System.Drawing.Size(680, 910); $mainForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen; $mainForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable; $mainForm.MaximizeBox = $true; $mainForm.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi
$statusStrip = New-Object System.Windows.Forms.StatusStrip; $statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel; $statusLabel.Name = "statusLabel"; $statusLabel.Text = "Ready. Connect to Exchange Online."; $statusStrip.Items.Add($statusLabel); $mainForm.Controls.Add($statusStrip)

# Connect Button
$connectButton = New-Object System.Windows.Forms.Button; $connectButton.Location = New-Object System.Drawing.Point(20, 20); $connectButton.Size = New-Object System.Drawing.Size(180, 30); $connectButton.Text = "Connect & Load Mailboxes"; $connectButton.Anchor = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left)
$connectButton.add_Click({
    $statusLabel.Text = "Checking existing connection..."; $mainForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $script:lastExportedXlsxPath = $null; if ($openFileButton) { $openFileButton.Enabled = $false }
    $userMailboxCheckedListBox.Items.Clear(); $getRulesButton.Enabled = $false; $selectAllButton.Enabled = $false; $deselectAllButton.Enabled = $false; $manageRulesButton.Enabled = $false
    $script:allLoadedMailboxUPNs = @() 
    # Disable Graph buttons initially
    $blockUserButton.Enabled = $false; $unblockUserButton.Enabled = $false; $manageRestrictedSendersButton.Enabled = $false


    try {
        $existingConnection = Get-ConnectionInformation -ErrorAction SilentlyContinue
        if ($existingConnection) {
            Write-Host "Using existing Exchange Online session for $($existingConnection.UserPrincipalName)." -ForegroundColor Cyan
            $script:currentExchangeConnection = $existingConnection
            $statusLabel.Text = "Using existing session: $($existingConnection.UserPrincipalName). Fetching mailboxes..."
        } else {
            $statusLabel.Text = "Connecting to Exchange Online..."; $mainForm.Refresh()
            Connect-ExchangeOnline -ErrorAction Stop
            $script:currentExchangeConnection = Get-ConnectionInformation
            Write-Host "Successfully connected to Exchange Online as $($script:currentExchangeConnection.UserPrincipalName)." -ForegroundColor Green
            $statusLabel.Text = "Connected as $($script:currentExchangeConnection.UserPrincipalName). Fetching mailboxes..."
        }
        $mainForm.Refresh()
        
        $statusLabel.Text = "Fetching mailboxes..."; $mainForm.Refresh(); Write-Host "Fetching mailboxes..."
        $allMailboxes = Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox -ErrorAction Stop | Select-Object UserPrincipalName, DisplayName | Sort-Object UserPrincipalName
        $script:allLoadedMailboxUPNs = $allMailboxes.UserPrincipalName 
        if ($allMailboxes) {
            foreach ($mbx in $allMailboxes) { $userMailboxCheckedListBox.Items.Add($mbx.UserPrincipalName, $false) }
            
            # Auto-detect and populate organization domains
            $statusLabel.Text = "Auto-detecting organization domains..."; $mainForm.Refresh()
            try {
                $autoDetectedDomains = Get-AutoDetectedDomains -MailboxUPNs $script:allLoadedMailboxUPNs
                if ($autoDetectedDomains.Count -gt 0) {
                    $orgDomainsTextBox.Text = ($autoDetectedDomains -join ', ')
                    $statusLabel.Text = "Connected. Loaded $($allMailboxes.Count) mailboxes. Auto-detected domains: $($autoDetectedDomains -join ', ')"
                    Write-Host "Auto-populated organization domains: $($autoDetectedDomains -join ', ')" -ForegroundColor Green
                } else {
                    $statusLabel.Text = "Connected. Loaded $($allMailboxes.Count) mailboxes. Please enter Organization Domains manually."
                    Write-Warning "Could not auto-detect domains. Please enter organization domains manually."
                }
            } catch {
                $ex = $_.Exception
                Write-Warning "Domain auto-detection failed: $($ex.Message)"
                $statusLabel.Text = "Connected. Loaded $($allMailboxes.Count) mailboxes. Auto-detection failed - please enter Organization Domains manually."
            }
            
            Write-Host "Loaded $($allMailboxes.Count) mailboxes." -FG Green
            $selectAllButton.Enabled = $true; $deselectAllButton.Enabled = $true
        } else { $statusLabel.Text = "Connected. No mailboxes found."; Write-Warning "No mailboxes found." }
        
        $disconnectButton.Enabled = $true; $connectButton.Enabled = $false; 
        $transportRulesButton.Enabled = $true; $connectorsButton.Enabled = $true; $sessionRevocationButton.Enabled = $true; $autoDetectDomainsButton.Enabled = $true
        $graphConnectButton.Enabled = $true; Update-GraphButtonText

        # --- Attempt MS Graph Connection ---
        if ($script:currentExchangeConnection) { 
            $currentStatus = $statusLabel.Text
            $statusLabel.Text = ($currentStatus + " Attempting MS Graph connection...")
            $mainForm.Refresh()
            if (Connect-GraphService -statusLabel $statusLabel -mainForm $mainForm) {
                # Graph buttons will be enabled/disabled by ItemCheck event based on selection
                $statusLabel.Text = ($currentStatus + " MS Graph Connected.")
                Update-GraphButtonText
            } else {
                $statusLabel.Text = ($currentStatus + " MS Graph connection failed/skipped.")
                Update-GraphButtonText
            }
        }

    } catch {
        $ex = $_.Exception
        $statusLabel.Text = "Connection/Load failed. See console."; Write-Error ("Connection/Load failed: {0}" -f $ex.Message)
        [System.Windows.Forms.MessageBox]::Show(("Failed to connect or load mailboxes.`nError: {0}" -f $ex.Message), "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        $disconnectButton.Enabled = $false; $connectButton.Enabled = $true; $getRulesButton.Enabled = $false; $selectAllButton.Enabled = $false; $deselectAllButton.Enabled = $false; $manageRulesButton.Enabled = $false
        $transportRulesButton.Enabled = $false; $connectorsButton.Enabled = $false; $sessionRevocationButton.Enabled = $false; $autoDetectDomainsButton.Enabled = $false
        $blockUserButton.Enabled = $false; $unblockUserButton.Enabled = $false; $manageRestrictedSendersButton.Enabled = $false
        $graphConnectButton.Enabled = $false; Update-GraphButtonText
    } finally { $mainForm.Cursor = [System.Windows.Forms.Cursors]::Default }
})
$mainForm.Controls.Add($connectButton)

# Disconnect Button
$disconnectButton = New-Object System.Windows.Forms.Button; $disconnectButton.Location = New-Object System.Drawing.Point(210, 20); $disconnectButton.Size = New-Object System.Drawing.Size(180, 30); $disconnectButton.Text = "Disconnect from Exchange"; $disconnectButton.Enabled = $false; $disconnectButton.Anchor = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left)
$disconnectButton.add_Click({
    $statusLabel.Text = "Disconnecting..."; $mainForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        if ($script:currentExchangeConnection) { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction Stop; $script:currentExchangeConnection = $null; Write-Host "Disconnected from Exchange Online." -FG Green; $statusLabel.Text = "Disconnected from Exchange." }
        else { Write-Host "Not connected to Exchange." -FG Yellow; $statusLabel.Text = "Not connected to Exchange." }
        
        $script:graphConnection = $null 
        $script:graphConnectionAttempted = $false 

        $disconnectButton.Enabled = $false; $connectButton.Enabled = $true; $getRulesButton.Enabled = $false; $script:lastExportedXlsxPath = $null; if ($openFileButton) { $openFileButton.Enabled = $false }; $userMailboxCheckedListBox.Items.Clear(); $selectAllButton.Enabled = $false; $deselectAllButton.Enabled = $false; $script:allLoadedMailboxUPNs = @(); $manageRulesButton.Enabled = $false
        $transportRulesButton.Enabled = $false; $connectorsButton.Enabled = $false; $sessionRevocationButton.Enabled = $false; $autoDetectDomainsButton.Enabled = $false
        $blockUserButton.Enabled = $false; $unblockUserButton.Enabled = $false; $manageRestrictedSendersButton.Enabled = $false
        $orgDomainsTextBox.Text = "" # Clear the domains when disconnecting
        $graphConnectButton.Enabled = $false; Update-GraphButtonText

        $statusLabel.Text = "Disconnected. Ready to connect."

    } catch { 
        $ex = $_.Exception
        $statusLabel.Text = "Disconnection error."; Write-Error ("Disconnection error: {0}" -f $ex.Message); 
        [System.Windows.Forms.MessageBox]::Show(("Disconnection error.`nError: {0}" -f $ex.Message), "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) 
    } finally { $mainForm.Cursor = [System.Windows.Forms.Cursors]::Default }
})
$mainForm.Controls.Add($disconnectButton)

# MS Graph Connect/Disconnect Button
$graphConnectButton = New-Object System.Windows.Forms.Button; $graphConnectButton.Location = New-Object System.Drawing.Point(400, 90); $graphConnectButton.Size = New-Object System.Drawing.Size(260, 30); $graphConnectButton.Text = "Connect to MS Graph"; $graphConnectButton.Enabled = $false; $graphConnectButton.Anchor = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right)
$graphConnectButton.add_Click({
    if ($script:graphConnection) {
        # Disconnect from Graph
        $statusLabel.Text = "Disconnecting from Microsoft Graph..."
        $mainForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            Disconnect-MgGraph -ErrorAction Stop
            $script:graphConnection = $null
            $script:graphConnectionAttempted = $false
            Write-Host "Disconnected from Microsoft Graph." -ForegroundColor Yellow
            $statusLabel.Text = "Disconnected from Microsoft Graph."
            
            # Disable Graph-dependent buttons
            $blockUserButton.Enabled = $false
            $unblockUserButton.Enabled = $false  
            $manageRestrictedSendersButton.Enabled = $false
            Update-GraphButtonText
            
        } catch {
            $ex = $_.Exception
            Write-Error "Error disconnecting from Microsoft Graph: $($ex.Message)"
            $statusLabel.Text = "Error disconnecting from MS Graph."
            [System.Windows.Forms.MessageBox]::Show("Error disconnecting from Microsoft Graph:`n$($ex.Message)", "Disconnect Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        } finally {
            $mainForm.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    } else {
        # Connect to Graph
        if (Connect-GraphService -statusLabel $statusLabel -mainForm $mainForm) {
            Update-GraphButtonText
            # Re-evaluate Graph button states based on current selections
            $selectedCount = $userMailboxCheckedListBox.CheckedItems.Count
            if ($script:graphConnection) {
                $blockUserButton.Enabled = ($selectedCount -gt 0)
                $unblockUserButton.Enabled = ($selectedCount -gt 0)
                $manageRestrictedSendersButton.Enabled = ($selectedCount -eq 1)
            }
        } else {
            Update-GraphButtonText
        }
    }
})
$mainForm.Controls.Add($graphConnectButton)

# Function to update Graph button text based on connection status
Function Update-GraphButtonText {
    if ($script:graphConnection) {
        $graphConnectButton.Text = "Disconnect from MS Graph"
        $graphConnectButton.ForeColor = [System.Drawing.Color]::Red
    } else {
        $graphConnectButton.Text = "Connect to MS Graph"  
        $graphConnectButton.ForeColor = [System.Drawing.Color]::Black
    }
}

# --- Exchange Feature Buttons ---
$transportRulesButton = New-Object System.Windows.Forms.Button; $transportRulesButton.Location = New-Object System.Drawing.Point(400, 20); $transportRulesButton.Size = New-Object System.Drawing.Size(130, 30); $transportRulesButton.Text = "Transport Rules"; $transportRulesButton.Enabled = $false; $transportRulesButton.Anchor = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left)
$transportRulesButton.add_Click({
    Write-Host "Transport Rules button clicked"
    Show-TransportRulesViewer -mainForm $mainForm -statusLabel $statusLabel
})
$mainForm.Controls.Add($transportRulesButton)

$connectorsButton = New-Object System.Windows.Forms.Button; $connectorsButton.Location = New-Object System.Drawing.Point(540, 20); $connectorsButton.Size = New-Object System.Drawing.Size(120, 30); $connectorsButton.Text = "Connectors"; $connectorsButton.Enabled = $false; $connectorsButton.Anchor = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right)
$connectorsButton.add_Click({
    Write-Host "Connectors button clicked"
    Show-ConnectorsViewer -mainForm $mainForm -statusLabel $statusLabel
})
$mainForm.Controls.Add($connectorsButton)

$sessionRevocationButton = New-Object System.Windows.Forms.Button; $sessionRevocationButton.Location = New-Object System.Drawing.Point(20, 90); $sessionRevocationButton.Size = New-Object System.Drawing.Size(260, 30); $sessionRevocationButton.Text = "Revoke User Sessions (Graph)"; $sessionRevocationButton.Enabled = $false; $sessionRevocationButton.Anchor = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left)
$sessionRevocationButton.add_Click({
    Write-Host "Revoke User Sessions button clicked"
    Show-SessionRevocationTool -mainForm $mainForm -statusLabel $statusLabel -allLoadedMailboxUPNs $script:allLoadedMailboxUPNs
})
$mainForm.Controls.Add($sessionRevocationButton)

# User Mailbox List Label & CheckedListBox
$userMailboxListLabel = New-Object System.Windows.Forms.Label; $userMailboxListLabel.Location = New-Object System.Drawing.Point(20, 130); $userMailboxListLabel.Size = New-Object System.Drawing.Size(200, 20); $userMailboxListLabel.Text = "Select Mailbox(es) to Analyze:"; $mainForm.Controls.Add($userMailboxListLabel)

$selectAllButton = New-Object System.Windows.Forms.Button; $selectAllButton.Location = New-Object System.Drawing.Point(230, 128); $selectAllButton.Size = New-Object System.Drawing.Size(100, 23); $selectAllButton.Text = "Select All"; $selectAllButton.Enabled = $false; $selectAllButton.Anchor = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left)
$selectAllButton.add_Click({ for ($i = 0; $i -lt $userMailboxCheckedListBox.Items.Count; $i++) { $userMailboxCheckedListBox.SetItemChecked($i, $true) } })
$mainForm.Controls.Add($selectAllButton)

$deselectAllButton = New-Object System.Windows.Forms.Button; $deselectAllButton.Location = New-Object System.Drawing.Point(340, 128); $deselectAllButton.Size = New-Object System.Drawing.Size(100, 23); $deselectAllButton.Text = "Deselect All"; $deselectAllButton.Enabled = $false; $deselectAllButton.Anchor = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left)
$deselectAllButton.add_Click({ for ($i = 0; $i -lt $userMailboxCheckedListBox.Items.Count; $i++) { $userMailboxCheckedListBox.SetItemChecked($i, $false) } })
$mainForm.Controls.Add($deselectAllButton)

$userMailboxCheckedListBox = New-Object System.Windows.Forms.CheckedListBox; $userMailboxCheckedListBox.Location = New-Object System.Drawing.Point(20, 155); $userMailboxCheckedListBox.Size = New-Object System.Drawing.Size(640, 150); $userMailboxCheckedListBox.CheckOnClick = $true; $userMailboxCheckedListBox.Anchor = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right); 
$userMailboxCheckedListBox.add_ItemCheck({
    $mainForm.BeginInvoke([System.Action]{
        $selectedCount = $userMailboxCheckedListBox.CheckedItems.Count 
        $getRulesButton.Enabled = ($selectedCount -gt 0 -and -not [string]::IsNullOrWhiteSpace($orgDomainsTextBox.Text) -and -not [string]::IsNullOrWhiteSpace($outputFolderTextBox.Text) -and $disconnectButton.Enabled)
        $manageRulesButton.Enabled = ($selectedCount -eq 1 -and $disconnectButton.Enabled) 
        
        if ($script:graphConnection) {
            $blockUserButton.Enabled = ($selectedCount -gt 0)
            $unblockUserButton.Enabled = ($selectedCount -gt 0)
            $manageRestrictedSendersButton.Enabled = ($selectedCount -eq 1)
            $sessionRevocationButton.Enabled = $true 
        } else {
            $blockUserButton.Enabled = $false
            $unblockUserButton.Enabled = $false
            $manageRestrictedSendersButton.Enabled = $false
            $sessionRevocationButton.Enabled = ($disconnectButton.Enabled -and $script:graphConnectionAttempted) 
        }
    })
})
$mainForm.Controls.Add($userMailboxCheckedListBox)

# Organization Domains Label and TextBox
$orgDomainsLabel = New-Object System.Windows.Forms.Label; $orgDomainsLabel.Location = New-Object System.Drawing.Point(20, 320); $orgDomainsLabel.Size = New-Object System.Drawing.Size(200, 20); $orgDomainsLabel.Text = "Organization Domains (comma-sep):"; $mainForm.Controls.Add($orgDomainsLabel)
$orgDomainsTextBox = New-Object System.Windows.Forms.TextBox; $orgDomainsTextBox.Location = New-Object System.Drawing.Point(230, 320); $orgDomainsTextBox.Size = New-Object System.Drawing.Size(350, 25); $orgDomainsTextBox.Anchor = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right);
$orgDomainsTextBox.add_TextChanged({ $getRulesButton.Enabled = ($userMailboxCheckedListBox.CheckedItems.Count -gt 0 -and -not [string]::IsNullOrWhiteSpace($orgDomainsTextBox.Text) -and -not [string]::IsNullOrWhiteSpace($outputFolderTextBox.Text) -and $disconnectButton.Enabled) })
$mainForm.Controls.Add($orgDomainsTextBox)

# Auto-Detect Domains Button
$autoDetectDomainsButton = New-Object System.Windows.Forms.Button; $autoDetectDomainsButton.Location = New-Object System.Drawing.Point(590, 318); $autoDetectDomainsButton.Size = New-Object System.Drawing.Size(70, 27); $autoDetectDomainsButton.Text = "Auto-Detect"; $autoDetectDomainsButton.Enabled = $false; $autoDetectDomainsButton.Anchor = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right)
$autoDetectDomainsButton.add_Click({
    if (-not $script:allLoadedMailboxUPNs -or $script:allLoadedMailboxUPNs.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No mailboxes loaded. Please connect to Exchange Online first.", "No Mailboxes", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    $statusLabel.Text = "Auto-detecting organization domains..."
    $autoDetectDomainsButton.Enabled = $false
    $mainForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $mainForm.Refresh()
    
    try {
        $autoDetectedDomains = Get-AutoDetectedDomains -MailboxUPNs $script:allLoadedMailboxUPNs
        if ($autoDetectedDomains.Count -gt 0) {
            $orgDomainsTextBox.Text = ($autoDetectedDomains -join ', ')
            $statusLabel.Text = "Auto-detected domains: $($autoDetectedDomains -join ', ')"
            [System.Windows.Forms.MessageBox]::Show("Auto-detected domains:`n`n$($autoDetectedDomains -join "`n")`n`nYou can edit these domains if needed.", "Auto-Detection Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        } else {
            $statusLabel.Text = "Could not auto-detect domains from loaded mailboxes."
            [System.Windows.Forms.MessageBox]::Show("Could not auto-detect organization domains from the loaded mailboxes.`n`nPlease enter the domains manually.", "Auto-Detection Failed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
    } catch {
        $ex = $_.Exception
        $statusLabel.Text = "Domain auto-detection failed."
        [System.Windows.Forms.MessageBox]::Show("Error during auto-detection:`n$($ex.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    } finally {
        $autoDetectDomainsButton.Enabled = $script:allLoadedMailboxUPNs.Count -gt 0
        $mainForm.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})
$mainForm.Controls.Add($autoDetectDomainsButton)

# Suspicious Keywords Label and TextBox
$keywordsLabel = New-Object System.Windows.Forms.Label; $keywordsLabel.Location = New-Object System.Drawing.Point(20, 355); $keywordsLabel.Size = New-Object System.Drawing.Size(200, 20); $keywordsLabel.Text = "Additional Keywords (comma-sep):"; $mainForm.Controls.Add($keywordsLabel)
$keywordsTextBox = New-Object System.Windows.Forms.TextBox; $keywordsTextBox.Location = New-Object System.Drawing.Point(230, 355); $keywordsTextBox.Size = New-Object System.Drawing.Size(430, 25); $keywordsTextBox.Anchor = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right); $keywordsTextBox.Text = ($BaseSuspiciousKeywords -join ', '); $mainForm.Controls.Add($keywordsTextBox)

# Output Folder Label and TextBox
$outputFolderLabel = New-Object System.Windows.Forms.Label; $outputFolderLabel.Location = New-Object System.Drawing.Point(20, 390); $outputFolderLabel.Size = New-Object System.Drawing.Size(100, 20); $outputFolderLabel.Text = "Output Folder:"; $mainForm.Controls.Add($outputFolderLabel)
$outputFolderTextBox = New-Object System.Windows.Forms.TextBox; $outputFolderTextBox.Location = New-Object System.Drawing.Point(120, 390); $outputFolderTextBox.Size = New-Object System.Drawing.Size(450, 25); $outputFolderTextBox.ReadOnly = $true; $outputFolderTextBox.Anchor = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right);
$outputFolderTextBox.add_TextChanged({ $getRulesButton.Enabled = ($userMailboxCheckedListBox.CheckedItems.Count -gt 0 -and -not [string]::IsNullOrWhiteSpace($orgDomainsTextBox.Text) -and -not [string]::IsNullOrWhiteSpace($outputFolderTextBox.Text) -and $disconnectButton.Enabled) })
$mainForm.Controls.Add($outputFolderTextBox)

# Browse Button for Output Folder
$browseFolderButton = New-Object System.Windows.Forms.Button; $browseFolderButton.Location = New-Object System.Drawing.Point(580, 388); $browseFolderButton.Size = New-Object System.Drawing.Size(80, 27); $browseFolderButton.Text = "Browse..."; $browseFolderButton.Anchor = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right)
$browseFolderButton.add_Click({
    $folderBrowserDialog = New-Object System.Windows.Forms.FolderBrowserDialog; $folderBrowserDialog.Description = "Select folder for XLSX file"; $folderBrowserDialog.ShowNewFolderButton = $true
    if ($folderBrowserDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $outputFolderTextBox.Text = $folderBrowserDialog.SelectedPath; $statusLabel.Text = "Output folder: $($outputFolderTextBox.Text)"; $getRulesButton.Enabled = ($userMailboxCheckedListBox.CheckedItems.Count -gt 0 -and -not [string]::IsNullOrWhiteSpace($orgDomainsTextBox.Text) -and -not [string]::IsNullOrWhiteSpace($outputFolderTextBox.Text) -and $disconnectButton.Enabled) }
})
$mainForm.Controls.Add($browseFolderButton)

# Progress Bar
$progressBar = New-Object System.Windows.Forms.ProgressBar; $progressBar.Location = New-Object System.Drawing.Point(20, 425); $progressBar.Size = New-Object System.Drawing.Size(640, 20); $progressBar.Anchor = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right); $progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous; $mainForm.Controls.Add($progressBar)

# Get Inbox Rules Button
$getRulesButton = New-Object System.Windows.Forms.Button; $getRulesButton.Location = New-Object System.Drawing.Point(20, 455); $getRulesButton.Size = New-Object System.Drawing.Size(315, 40); $getRulesButton.Text = "Get Rules for Selected (Export)"; $getRulesButton.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 10, [System.Drawing.FontStyle]::Bold); $getRulesButton.Enabled = $false; $getRulesButton.Anchor = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left); 
$getRulesButton.add_Click({
    param($sender, $e)
    if (-not $script:currentExchangeConnection) { [System.Windows.Forms.MessageBox]::Show("Connect to Exchange Online first.", "Not Connected", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning); return }
    if ([string]::IsNullOrWhiteSpace($orgDomainsTextBox.Text)) { [System.Windows.Forms.MessageBox]::Show("Enter organization domains.", "Config Missing", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning); $orgDomainsTextBox.Focus(); return }
    if ([string]::IsNullOrWhiteSpace($outputFolderTextBox.Text) -or (-not (Test-Path -Path $outputFolderTextBox.Text -PathType Container))) { [System.Windows.Forms.MessageBox]::Show("Select a valid output folder.", "Output Folder Invalid", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning); $browseFolderButton.PerformClick(); return }
    $selectedUpns = $userMailboxCheckedListBox.CheckedItems | ForEach-Object { $_ }
    if ($selectedUpns.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("Select at least one mailbox.", "No Mailbox Selected", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning); return }

    $statusLabel.Text = "Starting rule analysis for selected mailboxes..."; $mainForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor; $getRulesButton.Enabled = $false; $manageRulesButton.Enabled = $false; $connectButton.Enabled = $false; $disconnectButton.Enabled = $false; if ($openFileButton) { $openFileButton.Enabled = $false }; $progressBar.Value = 0
    $lcOrganizationDomains = $orgDomainsTextBox.Text -split ',' | ForEach-Object { $_.Trim().ToLowerInvariant() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    $SuspiciousKeywordsToUse = $keywordsTextBox.Text -split ',' | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }; if ($SuspiciousKeywordsToUse.Count -eq 0) { $SuspiciousKeywordsToUse = $BaseSuspiciousKeywords }
    Write-Host "Using Org Domains: $($lcOrganizationDomains -join ', ')" -FG Cyan; Write-Host "Using Keywords: $($SuspiciousKeywordsToUse -join ', ')" -FG Cyan
    $outputFolder = $outputFolderTextBox.Text; $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    
    # --- Tenant Identifier for Filename (Prioritized Logic) ---
    $tenantIdentifierForFile = $null
    $guiDomainsForFilename = ($orgDomainsTextBox.Text -split ',') | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

    if ($guiDomainsForFilename.Count -gt 0) {
        $firstGuiDomain = $guiDomainsForFilename[0]
        if ($firstGuiDomain -notlike "*.onmicrosoft.com") {
            $tenantIdentifierForFile = $firstGuiDomain
            Write-Host "Filename domain (Priority 1: GUI non-onmicrosoft): $tenantIdentifierForFile"
        } else {
            $tenantIdentifierForFile = $firstGuiDomain 
            Write-Host "Filename domain (Priority 1: GUI onmicrosoft - candidate): $tenantIdentifierForFile"
        }
    }
        
    if ((-not $tenantIdentifierForFile -or $tenantIdentifierForFile -like "*.onmicrosoft.com") -and $script:allLoadedMailboxUPNs.Count -gt 0) {
        Write-Host "Current tenant identifier '$tenantIdentifierForFile' is null or onmicrosoft.com. Checking loaded mailboxes for a better domain..."
        $domainFromLoadedUPNs = $null; $representativeDomainFound = $false
        $sampleCount = [System.Math]::Min(50, $script:allLoadedMailboxUPNs.Count) 
        for ($s_idx = 0; $s_idx -lt $sampleCount; $s_idx++) {
            $sampleUpn = $script:allLoadedMailboxUPNs[$s_idx]
            if ($sampleUpn -like "*@*") {
                $potentialDomain = ($sampleUpn.Split('@')[1])
                if (-not [string]::IsNullOrWhiteSpace($potentialDomain) -and $potentialDomain -notlike "*.onmicrosoft.com") {
                    $domainFromLoadedUPNs = $potentialDomain; $representativeDomainFound = $true; break 
                } elseif (-not $domainFromLoadedUPNs -and -not [string]::IsNullOrWhiteSpace($potentialDomain)) { $domainFromLoadedUPNs = $potentialDomain }
            }
        }
        if ($representativeDomainFound) { $tenantIdentifierForFile = $domainFromLoadedUPNs; Write-Host "Filename domain (Priority 2: Loaded mailbox UPN - vanity): $tenantIdentifierForFile" }
        elseif ($domainFromLoadedUPNs -and (-not $tenantIdentifierForFile -or $tenantIdentifierForFile -like "*.onmicrosoft.com" )) { $tenantIdentifierForFile = $domainFromLoadedUPNs; Write-Host "Filename domain (Priority 2: Loaded mailbox UPN - onmicrosoft): $tenantIdentifierForFile" }
        elseif (-not $domainFromLoadedUPNs) { Write-Warning "Could not derive any domain from sample of loaded mailboxes." }
    }

    if (-not $tenantIdentifierForFile -or ($tenantIdentifierForFile -like "*.onmicrosoft.com")) { 
        if ($script:currentExchangeConnection.UserPrincipalName -like "*@*") {
            $adminUpnDomain = ($script:currentExchangeConnection.UserPrincipalName.Split('@')[1])
            if (-not [string]::IsNullOrWhiteSpace($adminUpnDomain)) {
                if (($adminUpnDomain -notlike "*.onmicrosoft.com" -and (-not $tenantIdentifierForFile -or $tenantIdentifierForFile -like "*.onmicrosoft.com")) -or (-not $tenantIdentifierForFile)) { $tenantIdentifierForFile = $adminUpnDomain; Write-Host "Filename domain (Priority 3: Admin UPN - non-onmicrosoft): $tenantIdentifierForFile" }
                elseif(-not $tenantIdentifierForFile) {$tenantIdentifierForFile = $adminUpnDomain; Write-Host "Filename domain (Priority 3: Admin UPN - onmicrosoft as last resort): $tenantIdentifierForFile"}
            }
        }
    }
    if (-not $tenantIdentifierForFile) { $tenantIdentifierForFile = "UnknownTenant"; Write-Warning "Filename domain (Final Fallback): UnknownTenant" }
    $safeTenantDomain = $tenantIdentifierForFile -replace "[^a-zA-Z0-9_.-]", "" -replace "\.", "_"
    # --- End Tenant Identifier Logic ---

    $baseFileName = "ExchangeInboxRules_$($safeTenantDomain)_$timestamp"; $csvFilePath = Join-Path -Path $outputFolder -ChildPath "$($baseFileName)_temp.csv"; $xlsxFilePath = Join-Path -Path $outputFolder -ChildPath "$($baseFileName).xlsx"
    Write-Host "Final output file will be: $xlsxFilePath"

    $allRuleAnalysisData = @(); $errorOccurred = $false; $csvExported = $false
    try {
        $totalMailboxesToProcess = $selectedUpns.Count; $progressCount = 0
        $progressBar.Maximum = 100; $progressBar.Value = 0; $progressBar.Step = 1  # Ensure progress bar is properly initialized
        
        foreach ($userPrincipalName in $selectedUpns) {
            $progressCount++; $statusLabel.Text = "Processing $userPrincipalName ($progressCount/$totalMailboxesToProcess)"
            
            # Calculate progress as percentage and ensure it doesn't exceed maximum
            $progressPercentage = [Math]::Min(100, [Math]::Round(($progressCount / $totalMailboxesToProcess) * 100))
            $progressBar.Value = $progressPercentage
            $mainForm.Refresh()
            Write-Verbose "Getting rules for: $userPrincipalName"
            $mailboxForwardingSmtpAddress = $null; $mailboxForwardingAddressUPN = $null; $mailboxDeliverToAndForward = $null; $mailboxInboxDelegates = $null; $mailboxFullAccessPermissions = $null; $mbxSettings = $null
            try {
                $mbxSettings = Get-Mailbox -Identity $userPrincipalName | Select-Object UserPrincipalName, DisplayName, ForwardingSmtpAddress, ForwardingAddress, DeliverToMailboxAndForward, GrantSendOnBehalfTo
                if ($mbxSettings) { 
                    $fwdSmtpAddrDisplay = "N/A"
                    if ($mbxSettings.ForwardingSmtpAddress) {
                        if (-not [string]::IsNullOrWhiteSpace($mbxSettings.ForwardingSmtpAddress.SmtpAddress)) { $fwdSmtpAddrDisplay = $mbxSettings.ForwardingSmtpAddress.SmtpAddress } 
                        elseif (-not [string]::IsNullOrWhiteSpace($mbxSettings.ForwardingSmtpAddress.AddressString)) { $fwdSmtpAddrDisplay = $mbxSettings.ForwardingSmtpAddress.AddressString } 
                        elseif ($mbxSettings.ForwardingSmtpAddress.ToString() -like "*@*") { $fwdSmtpAddrDisplay = $mbxSettings.ForwardingSmtpAddress.ToString() }
                    }
                    $mailboxForwardingSmtpAddress = $fwdSmtpAddrDisplay

                    $mailboxForwardingAddressUPN = if ($mbxSettings.ForwardingAddress) {$mbxSettings.ForwardingAddress.Name} else {$null} 
                    $mailboxDeliverToAndForward = $mbxSettings.DeliverToMailboxAndForward
                }
                $inboxPermissions = Get-MailboxFolderPermission -Identity "$($userPrincipalName):\Inbox" -ErrorAction SilentlyContinue; $delegatesList = @()
                if ($inboxPermissions) { foreach ($perm in $inboxPermissions) { if ($perm.User.UserType -ne "Default" -and $perm.User.UserType -ne "Anonymous") { if ($perm.AccessRights -match "Editor" -or $perm.AccessRights -match "Owner" -or $perm.AccessRights -match "Reviewer" -or $perm.AccessRights -match "Author") { $delegatesList += "$($perm.User.DisplayName) ($($perm.AccessRights -join ','))" } } } }
                $mailboxInboxDelegates = if ($delegatesList.Count -gt 0) { $delegatesList -join "; " } else { $null } 

                $fullAccessPerms = Get-MailboxPermission -Identity $userPrincipalName -ErrorAction SilentlyContinue
                $fullAccessList = @()
                if ($fullAccessPerms) {
                    foreach($perm in $fullAccessPerms) {
                        if ($perm.AccessRights -contains "FullAccess" -and $perm.IsInherited -eq $false -and $perm.User -notlike "NT AUTHORITY\SELF") {
                            $fullAccessList += $perm.User.ToString() 
                        }
                    }
                }
                $mailboxFullAccessPermissions = if ($fullAccessList.Count -gt 0) { $fullAccessList -join "; " } else { $null }

            } catch { 
                $ex = $_.Exception
                Write-Warning ("Could not retrieve forwarding/delegate/full access info for '{0}': {1}" -f $userPrincipalName, $ex.Message) 
                $mailboxInboxDelegates = "Error fetching delegates" 
                $mailboxFullAccessPermissions = "Error fetching full access"
            }
            try {
                $rules = Get-InboxRule -Mailbox $userPrincipalName -IncludeHidden -ErrorAction SilentlyContinue
                if ($rules) {
                    foreach ($rule in $rules) {
                        # Debug: Output all rule properties for analysis
                        Write-Host "Rule: $($rule.Name) | Properties: $($rule | Format-List | Out-String)"
                        $isHiddenValue = $false
                        if ($rule.PSObject.Properties.Match('IsHidden').Count -gt 0) {
                            $isHiddenValue = $rule.IsHidden
                        }
                        # Improved fallback: check for system-generated, RuleId, or other known patterns
                        if (-not $isHiddenValue) {
                            if (
                                $rule.Name -like 'RuleId:*' -or
                                ($rule.Description -match 'system-generated' -or $rule.Description -match 'Generated by Microsoft Exchange') -or
                                ($rule.Description -match 'hidden' -or $rule.Name -match 'hidden')
                            ) {
                                $isHiddenValue = $true
                            }
                        }
                        $recipientsToCheck = @(); if ($rule.ForwardTo) {$recipientsToCheck += $rule.ForwardTo}; if ($rule.ForwardAsAttachmentTo) {$recipientsToCheck += $rule.ForwardAsAttachmentTo}; if ($rule.RedirectTo) {$recipientsToCheck += $rule.RedirectTo}
                        $isForwardingExt = $false; if ($recipientsToCheck.Count -gt 0) { $isForwardingExt = Test-ExternalForwarding -RecipientAddresses $recipientsToCheck -InternalDomains $lcOrganizationDomains }
                        $keywordsFoundInName = $false; foreach ($keyword in $SuspiciousKeywordsToUse) { if ($rule.Name -match $keyword) { $keywordsFoundInName = $true; break } }
                        
                        $ruleDetails = [PSCustomObject]@{ 
                            MailboxOwner = $userPrincipalName
                            MailboxForwardingSmtpAddress = $mailboxForwardingSmtpAddress
                            MailboxForwardingAddressUPN = $mailboxForwardingAddressUPN
                            MailboxDeliverToAndForward = $mailboxDeliverToAndForward
                            MailboxGrantSendOnBehalfTo = if($mbxSettings){$mbxSettings.GrantSendOnBehalfTo -join '; '}else{$null}
                            MailboxInboxDelegates = $mailboxInboxDelegates
                            MailboxFullAccessPermissions = $mailboxFullAccessPermissions 
                            RuleName = $rule.Name
                            IsEnabled = $rule.Enabled
                            Priority = $rule.Priority
                            IsHidden = $isHiddenValue
                            IsForwardingExternal = $isForwardingExt 
                            IsDeleting = $rule.DeleteMessage
                            IsMarkingAsRead = $rule.MarkAsRead
                            IsMovingToFolder = ($null -ne $rule.MoveToFolder)
                            MoveToFolderName = if($rule.MoveToFolder){$rule.MoveToFolder.ToString()}else{$null}
                            SuspiciousKeywordsInName= $keywordsFoundInName
                            Description = $rule.Description
                            StopProcessingRules = $rule.StopProcessingRules
                            Conditions = ($rule.Conditions | Out-String).Trim()
                            Actions = ($rule.Actions | Out-String).Trim()
                            Exceptions = ($rule.Exceptions | Out-String).Trim()
                            RuleID = $rule.Identity.ToString() 
                        }
                        $allRuleAnalysisData += $ruleDetails
                    }
                }
            } catch { 
                $ex = $_.Exception
                Write-Warning ("Error processing rules for mailbox '{0}': {1}" -f $userPrincipalName, $ex.Message); 
                $errorOccurred = $true 
            } 
        }
        $progressBar.Value = [Math]::Min(100, $progressBar.Maximum); $statusLabel.Text = "Rule processing complete. Exporting..."; $mainForm.Refresh()
        if ($allRuleAnalysisData.Count -eq 0) {
            if (-not $errorOccurred) { $statusLabel.Text = "No rules for selected mailboxes."; [System.Windows.Forms.MessageBox]::Show("No inbox rules found for selected mailboxes.", "No Rules", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) }
            else { $statusLabel.Text = "Errors & no rules found."; [System.Windows.Forms.MessageBox]::Show("Errors occurred & no rules found.", "Complete with Errors", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) }
        } else {
            Write-Host "Exporting $($allRuleAnalysisData.Count) rules to CSV: $csvFilePath"
            try { $allRuleAnalysisData | Sort-Object MailboxOwner, Priority | Export-Csv -Path $csvFilePath -NoTypeInformation -Encoding UTF8 -EA Stop; Write-Host "Temp CSV export OK." -FG Green; $csvExported = $true }
            catch { 
                $ex = $_.Exception
                Write-Error "CSV Export Failed: $($ex.Message)"; 
                $statusLabel.Text = "CSV Export Error."; 
                [System.Windows.Forms.MessageBox]::Show("CSV Export Failed.`nError: $($ex.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error); 
                $errorOccurred = $true 
            }
            if ($csvExported) {
                $statusLabel.Text = "Converting to XLSX & formatting..."; Write-Host "Converting to XLSX..."
                if (Format-InboxRuleXlsx -CsvPath $csvFilePath -XlsxPath $xlsxFilePath -RowHighlightColumnHeader "IsHidden" -RowHighlightValue $true -CellHighlightColor $highlightColorIndexLightRed -CellHighlightValue $true) { 
                    $script:lastExportedXlsxPath = $xlsxFilePath
                    if ($openFileButton -and (Test-Path $script:lastExportedXlsxPath)) { $openFileButton.Enabled = $true }
                    $statusLabel.Text = "Exported & formatted $($allRuleAnalysisData.Count) rules to $xlsxFilePath"; [System.Windows.Forms.MessageBox]::Show("Exported & formatted $($allRuleAnalysisData.Count) rules to:`n$xlsxFilePath", "XLSX Export OK", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                    try { Remove-Item -Path $csvFilePath -Force -EA SilentlyContinue } catch {}
                } else {
                    $openFileButton.Enabled = $false
                    $statusLabel.Text = "CSV OK, XLSX/Format Failed."; [System.Windows.Forms.MessageBox]::Show("CSV Exported to:`n$csvFilePath`n`nXLSX/Format FAILED. Check Excel install & console.", "XLSX Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning); $errorOccurred = $true }
            }
        }
    } catch { 
        $ex = $_.Exception
        $statusLabel.Text = "Unexpected error. See console."; Write-Error "Unexpected error: $($ex.Message)`n$($ex.ScriptStackTrace)"; 
        [System.Windows.Forms.MessageBox]::Show("Unexpected error.`nError: $($ex.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error); 
        $errorOccurred = $true 
    }
            finally {
        $mainForm.Cursor = [System.Windows.Forms.Cursors]::Default; $getRulesButton.Enabled = $disconnectButton.Enabled; $manageRulesButton.Enabled = ($userMailboxCheckedListBox.CheckedItems.Count -eq 1 -and $disconnectButton.Enabled); $connectButton.Enabled = (-not $disconnectButton.Enabled); $disconnectButton.Enabled = $disconnectButton.Enabled
        if ($script:lastExportedXlsxPath -and (Test-Path $script:lastExportedXlsxPath)) { if ($openFileButton) { $openFileButton.Enabled = $true } } else { if ($openFileButton) { $openFileButton.Enabled = $false } }
        if ($errorOccurred) { $statusLabel.Text = "Finished with errors. See console." } elseif ($allRuleAnalysisData.Count -gt 0) {  } else { } 
        if ($csvExported -and $errorOccurred -and (Test-Path $csvFilePath)) { Write-Host "Temp CSV ($csvFilePath) kept due to XLSX error." }
        $progressBar.Value = 0  # Safely reset progress bar
    }
})
$mainForm.Controls.Add($getRulesButton)

# Manage Mailbox Rules Button
$manageRulesButton = New-Object System.Windows.Forms.Button; $manageRulesButton.Location = New-Object System.Drawing.Point(345, 455); $manageRulesButton.Size = New-Object System.Drawing.Size(315, 40); $manageRulesButton.Text = "Manage Rules for Selected Mailbox..."; $manageRulesButton.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 10, [System.Drawing.FontStyle]::Regular); $manageRulesButton.Enabled = $false; $manageRulesButton.Anchor = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right)
$manageRulesButton.add_Click({
    param($sender, $e)
    if ($userMailboxCheckedListBox.CheckedItems.Count -ne 1) {
        [System.Windows.Forms.MessageBox]::Show("Please select exactly one mailbox from the list to manage its rules.", "Select One Mailbox", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }
    $selectedMailboxUpn = $userMailboxCheckedListBox.CheckedItems[0]
    
    # --- Create and Show Rule Management Form ---
    $ruleManagementForm = New-Object System.Windows.Forms.Form
    $ruleManagementForm.Text = "Manage Inbox Rules for: $selectedMailboxUpn"
    $ruleManagementForm.Size = New-Object System.Drawing.Size(600, 450)
    $ruleManagementForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
    $ruleManagementForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog

    $rulesLabel = New-Object System.Windows.Forms.Label; $rulesLabel.Location = New-Object System.Drawing.Point(15, 15); $rulesLabel.Size = New-Object System.Drawing.Size(550, 20); $rulesLabel.Text = ("Inbox Rules for {0}:" -f $selectedMailboxUpn); $ruleManagementForm.Controls.Add($rulesLabel) 
    
    $rulesDisplayCheckedListBox = New-Object System.Windows.Forms.CheckedListBox; $rulesDisplayCheckedListBox.Location = New-Object System.Drawing.Point(15, 40); $rulesDisplayCheckedListBox.Size = New-Object System.Drawing.Size(550, 280); $rulesDisplayCheckedListBox.CheckOnClick = $true; $rulesDisplayCheckedListBox.Anchor = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom); $ruleManagementForm.Controls.Add($rulesDisplayCheckedListBox)
    
    $ruleDataStore = @{} 

    $loadRulesFunc = {
        param($targetMailboxUpn)
        $rulesDisplayCheckedListBox.Items.Clear()
        $ruleDataStore.Clear()
        $statusLabel.Text = "Loading rules for $targetMailboxUpn..."
        $mainForm.Refresh() 
        try {
            $rules = Get-InboxRule -Mailbox $targetMailboxUpn -IncludeHidden -ErrorAction Stop
            if ($rules) {
                foreach ($rule in $rules) {
                    $displayString = "$($rule.Name) (ID: $($rule.RuleIdentity); Enabled: $($rule.Enabled); Priority: $($rule.Priority))"
                    $rulesDisplayCheckedListBox.Items.Add($displayString, $false)
                    $ruleDataStore[$displayString] = $rule 
                }
                $statusLabel.Text = "Loaded $($rules.Count) rules for $targetMailboxUpn."
            } else {
                $statusLabel.Text = "No rules found for $targetMailboxUpn."
            }
        } catch {
            $ex = $_.Exception
            [System.Windows.Forms.MessageBox]::Show(("Error loading rules for {0}`n{1}" -f $targetMailboxUpn, $ex.Message), "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            $statusLabel.Text = "Error loading rules for $targetMailboxUpn."
        }
        $mainForm.Refresh()
    }

    $refreshRulesButton = New-Object System.Windows.Forms.Button; $refreshRulesButton.Location = New-Object System.Drawing.Point(15, 330); $refreshRulesButton.Size = New-Object System.Drawing.Size(120, 30); $refreshRulesButton.Text = "Refresh Rules"; $refreshRulesButton.Anchor = ([System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left)
    $refreshRulesButton.add_Click({ $loadRulesFunc.Invoke($selectedMailboxUpn) })
    $ruleManagementForm.Controls.Add($refreshRulesButton)

    $deleteRulesButton = New-Object System.Windows.Forms.Button; $deleteRulesButton.Location = New-Object System.Drawing.Point(145, 330); $deleteRulesButton.Size = New-Object System.Drawing.Size(150, 30); $deleteRulesButton.Text = "Delete Selected Rules"; $deleteRulesButton.ForeColor = [System.Drawing.Color]::Red; $deleteRulesButton.Anchor = ([System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left)
    $deleteRulesButton.add_Click({
        $checkedRuleItems = $rulesDisplayCheckedListBox.CheckedItems
        if ($checkedRuleItems.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("No rules selected for deletion.", "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            return
        }
        $ruleNamesToDelete = $checkedRuleItems | ForEach-Object { $_.ToString().Split('(')[0].Trim() }
        $confirmResult = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to delete the following $($checkedRuleItems.Count) rule(s) for $selectedMailboxUpn?`n`n$($ruleNamesToDelete -join "`n")", "Confirm Deletion", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($confirmResult -eq [System.Windows.Forms.DialogResult]::Yes) {
            $statusLabel.Text = "Deleting selected rules for $selectedMailboxUpn..."
            $mainForm.Refresh()
            $ruleManagementForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            foreach ($itemString in $checkedRuleItems) {
                $ruleObject = $ruleDataStore[$itemString]
                if ($ruleObject) {
                    try {
                        Write-Host "Attempting to delete rule '$($ruleObject.Name)' (Identity: $($ruleObject.RuleIdentity)) for $selectedMailboxUpn"
                        Remove-InboxRule -Mailbox $selectedMailboxUpn -Identity $ruleObject.RuleIdentity -Confirm:$false -ErrorAction Stop
                        Write-Host "Successfully deleted rule '$($ruleObject.Name)'" -ForegroundColor Green
                    } catch {
                        $ex = $_.Exception
                        Write-Warning "Failed to delete rule '$($ruleObject.Name)' using RuleIdentity. Error: $($ex.Message). Attempting with Name..."
                        try {
                            Remove-InboxRule -Mailbox $selectedMailboxUpn -Identity $ruleObject.Name -Confirm:$false -ErrorAction Stop
                            Write-Host "Successfully deleted rule '$($ruleObject.Name)' using Name as Identity." -ForegroundColor Green
                        } catch {
                            $ex2 = $_.Exception
                            [System.Windows.Forms.MessageBox]::Show(("Failed to delete rule '{0}' for {1} using both RuleIdentity and Name.`nError (RuleIdentity): {2}`nError (Name): {3}" -f $ruleObject.Name, $selectedMailboxUpn, $ex.Message, $ex2.Message), "Deletion Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                            Write-Warning ("Failed to delete rule '{0}' using Name as Identity: {1}" -f $ruleObject.Name, $ex2.Message)
                        }
                    }
                }
            }
            $ruleManagementForm.Cursor = [System.Windows.Forms.Cursors]::Default
            $loadRulesFunc.Invoke($selectedMailboxUpn) # Refresh the list
            $statusLabel.Text = "Rule deletion process complete for $selectedMailboxUpn."
            $mainForm.Refresh()
        }
    })
    $ruleManagementForm.Controls.Add($deleteRulesButton)

    $closeRuleFormButton = New-Object System.Windows.Forms.Button; $closeRuleFormButton.Location = New-Object System.Drawing.Point(450, 330); $closeRuleFormButton.Size = New-Object System.Drawing.Size(120, 30); $closeRuleFormButton.Text = "Close"; $closeRuleFormButton.Anchor = ([System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right)
    $closeRuleFormButton.add_Click({ $ruleManagementForm.Close() })
    $ruleManagementForm.Controls.Add($closeRuleFormButton)

    # Load rules when form is shown
    $ruleManagementForm.Add_Shown({ $loadRulesFunc.Invoke($selectedMailboxUpn) })
    [void]$ruleManagementForm.ShowDialog($mainForm) # Show as modal to the main form
    $ruleManagementForm.Dispose()
})
$mainForm.Controls.Add($manageRulesButton)

# Open Last Exported File Button
$openFileButton = New-Object System.Windows.Forms.Button; $openFileButton.Location = New-Object System.Drawing.Point(20, 505); $openFileButton.Size = New-Object System.Drawing.Size(640, 30); $openFileButton.Text = "Open Last Exported XLSX File"; $openFileButton.Enabled = $false; $openFileButton.Anchor = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right)
$openFileButton.add_Click({
    if ($script:lastExportedXlsxPath -and (Test-Path $script:lastExportedXlsxPath)) {
        try { Invoke-Item -Path $script:lastExportedXlsxPath -EA Stop; $statusLabel.Text = "Opening: $($script:lastExportedXlsxPath)" }
        catch { 
            $ex = $_.Exception
            [System.Windows.Forms.MessageBox]::Show(("Could not open: {0}`nError: {1}" -f $script:lastExportedXlsxPath, $ex.Message), "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error); 
            $statusLabel.Text = "Error opening file." 
        }
    } else { [System.Windows.Forms.MessageBox]::Show("No file exported or file not found.", "No File", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information); $statusLabel.Text = "No recent file." }
})
$mainForm.Controls.Add($openFileButton)

# --- MS Graph Action Buttons ---
$blockUserButton = New-Object System.Windows.Forms.Button
$blockUserButton.Location = New-Object System.Drawing.Point(20, 545) 
$blockUserButton.Size = New-Object System.Drawing.Size(200, 30)
$blockUserButton.Text = "Block Sign-in for Selected"
$blockUserButton.Enabled = $false
$blockUserButton.Anchor = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left)
$blockUserButton.add_Click({
    $selectedUpns = $userMailboxCheckedListBox.CheckedItems | ForEach-Object { $_ }
    if ($selectedUpns.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select at least one user.", "No Users Selected", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning); return
    }
    
    # Pre-check: Verify which users exist in Azure AD
    $statusLabel.Text = "Checking which users exist in Azure AD..."
    $mainForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $script:validUsersBlock = @()
    $script:invalidUsersBlock = @()
    
    foreach ($upn in $selectedUpns) {
        try {
            $mgUser = Get-MgUser -UserId $upn -Property Id -ErrorAction Stop
            if ($mgUser) {
                $script:validUsersBlock += $upn
            }
        } catch {
            $script:invalidUsersBlock += $upn
        }
    }
    
    $mainForm.Cursor = [System.Windows.Forms.Cursors]::Default
    
    # Show summary of valid vs invalid users
    $message = "Ready to BLOCK SIGN-IN for users:`n`n"
    if ($script:validUsersBlock.Count -gt 0) {
        $message += "✓ Users found in Azure AD ($($script:validUsersBlock.Count)):`n"
        $message += ($script:validUsersBlock -join "`n") + "`n`n"
    }
    if ($script:invalidUsersBlock.Count -gt 0) {
        $message += "⚠ Users NOT found in Azure AD ($($script:invalidUsersBlock.Count)):`n"
        $message += ($script:invalidUsersBlock -join "`n") + "`n"
        $message += "(These users exist only in Exchange Online and cannot be blocked via Azure AD)`n`n"
    }
    
    if ($script:validUsersBlock.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("None of the selected users exist in Azure AD.`n`nThese users likely exist only in Exchange Online or are on-premises users.`nSign-in blocking requires Azure AD accounts.", "No Azure AD Users Found", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }
    
    $message += "Continue with blocking sign-in for the $($script:validUsersBlock.Count) Azure AD user(s)?"
    $confirmResult = [System.Windows.Forms.MessageBox]::Show($message, "Confirm Block Sign-in", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
    
    if ($confirmResult -eq [System.Windows.Forms.DialogResult]::Yes) {
        # Create a local copy to avoid any scope issues  
        $usersToBlock = @($script:validUsersBlock)
        Write-Host "DEBUG: Created local copy with $($usersToBlock.Count) users for blocking" -ForegroundColor Cyan
        Set-UserSignInBlockedState -UserPrincipalNames $usersToBlock -Blocked $true -StatusLabel $statusLabel -ProgressBar $progressBar -MainForm $mainForm
    }
})
$mainForm.Controls.Add($blockUserButton)

$unblockUserButton = New-Object System.Windows.Forms.Button
$unblockUserButton.Location = New-Object System.Drawing.Point(230, 545) 
$unblockUserButton.Size = New-Object System.Drawing.Size(200, 30)
$unblockUserButton.Text = "Unblock Sign-in for Selected"
$unblockUserButton.Enabled = $false
$unblockUserButton.Anchor = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left)
$unblockUserButton.add_Click({
    $selectedUpns = $userMailboxCheckedListBox.CheckedItems | ForEach-Object { $_ }
    if ($selectedUpns.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select at least one user.", "No Users Selected", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning); return
    }
    
    # Pre-check: Verify which users exist in Azure AD (same as block button)
    $statusLabel.Text = "Checking which users exist in Azure AD..."
    $mainForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $script:validUsersUnblock = @()
    $script:invalidUsersUnblock = @()
    
    foreach ($upn in $selectedUpns) {
        try {
            $mgUser = Get-MgUser -UserId $upn -Property Id -ErrorAction Stop
            if ($mgUser) {
                $script:validUsersUnblock += $upn
            }
        } catch {
            $script:invalidUsersUnblock += $upn
        }
    }
    
    $mainForm.Cursor = [System.Windows.Forms.Cursors]::Default
    
    # Show summary of valid vs invalid users
    $message = "Ready to UNBLOCK SIGN-IN for users:`n`n"
    if ($script:validUsersUnblock.Count -gt 0) {
        $message += "✓ Users found in Azure AD ($($script:validUsersUnblock.Count)):`n"
        $message += ($script:validUsersUnblock -join "`n") + "`n`n"
    }
    if ($script:invalidUsersUnblock.Count -gt 0) {
        $message += "⚠ Users NOT found in Azure AD ($($script:invalidUsersUnblock.Count)):`n"
        $message += ($script:invalidUsersUnblock -join "`n") + "`n"
        $message += "(These users exist only in Exchange Online and cannot be unblocked via Azure AD)`n`n"
    }
    
    if ($script:validUsersUnblock.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("None of the selected users exist in Azure AD.`n`nThese users likely exist only in Exchange Online or are on-premises users.`nSign-in unblocking requires Azure AD accounts.", "No Azure AD Users Found", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }
    
    $message += "Continue with unblocking sign-in for the $($script:validUsersUnblock.Count) Azure AD user(s)?"
    $confirmResult = [System.Windows.Forms.MessageBox]::Show($message, "Confirm Unblock Sign-in", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Information)
    
    if ($confirmResult -eq [System.Windows.Forms.DialogResult]::Yes) {
        # Create a local copy to avoid any scope issues
        $usersToUnblock = @($script:validUsersUnblock)
        Write-Host "DEBUG: Created local copy with $($usersToUnblock.Count) users for unblocking" -ForegroundColor Cyan
        Set-UserSignInBlockedState -UserPrincipalNames $usersToUnblock -Blocked $false -StatusLabel $statusLabel -ProgressBar $progressBar -MainForm $mainForm
    }
})

$mainForm.Controls.Add($unblockUserButton)

$manageRestrictedSendersButton = New-Object System.Windows.Forms.Button
$manageRestrictedSendersButton.Location = New-Object System.Drawing.Point(440, 545) 
$manageRestrictedSendersButton.Size = New-Object System.Drawing.Size(220, 30) 
$manageRestrictedSendersButton.Text = "Manage Sending Restrictions..."
$manageRestrictedSendersButton.Enabled = $false
$manageRestrictedSendersButton.Anchor = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right) 
$manageRestrictedSendersButton.add_Click({
    if ($userMailboxCheckedListBox.CheckedItems.Count -ne 1) {
        [System.Windows.Forms.MessageBox]::Show("Please select exactly one user to manage sending restrictions.", "Select One User", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information); return
    }
    $selectedUpn = $userMailboxCheckedListBox.CheckedItems[0]
    
    # Check if user exists in Azure AD first
    $statusLabel.Text = "Checking if user exists in Azure AD..."
    $mainForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $userExistsInAzureAD = $false
    
    try {
        $mgUser = Get-MgUser -UserId $selectedUpn -Property Id -ErrorAction Stop
        if ($mgUser) {
            $userExistsInAzureAD = $true
        }
    } catch {
        $userExistsInAzureAD = $false
    }
    
    $mainForm.Cursor = [System.Windows.Forms.Cursors]::Default
    $statusLabel.Text = "Opening sending restrictions management..."
    
    # Show the dialog regardless, but with appropriate context
    Show-RestrictedSenderManagementDialog -UserPrincipalName $selectedUpn -ParentForm $mainForm -StatusLabelGlobal $statusLabel
})
$mainForm.Controls.Add($manageRestrictedSendersButton)


# --- Show Form ---
$mainForm.Add_Shown({$mainForm.Activate(); Update-GraphButtonText}) 
[void]$mainForm.ShowDialog()

# --- Script End ---
Write-Host "Script finished."
# No automatic disconnect on GUI close. User must use the "Disconnect" button.
# if ($script:currentExchangeConnection) { Write-Host "Disconnecting from Exchange Online..."; Disconnect-ExchangeOnline -Confirm:$false -EA SilentlyContinue }
# if ($script:graphConnection) { Write-Host "Disconnecting from Microsoft Graph..."; Disconnect-MgGraph -EA SilentlyContinue }