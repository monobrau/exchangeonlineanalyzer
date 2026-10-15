# ThreatHunterApp.ps1
# AI-powered threat investigation GUI with ConnectWise integration

# Requires Windows PowerShell 5.1 or PowerShell 7+
#Requires -Version 5.1

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Import required modules
$modulesPath = Join-Path $PSScriptRoot "Modules"
Import-Module (Join-Path $modulesPath "ConnectWiseIntegration.psm1") -Force
Import-Module (Join-Path $modulesPath "AIInvestigator.psm1") -Force
Import-Module (Join-Path $modulesPath "Settings.psm1") -Force
Import-Module (Join-Path $modulesPath "MemberberryIntegration.psm1") -Force -ErrorAction SilentlyContinue

# App state
$script:CurrentInvestigation = $null
$script:CurrentTicket = $null
$script:Settings = $null

#region UI Helper Functions

function New-ResizablePanel {
    param(
        [string]$Title,
        [int]$Width,
        [int]$Height,
        [System.Windows.Forms.DockStyle]$Dock = [System.Windows.Forms.DockStyle]::Fill
    )
    
    $panel = New-Object System.Windows.Forms.Panel
    $panel.Size = New-Object System.Drawing.Size($Width, $Height)
    $panel.Dock = $Dock
    $panel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    
    if ($Title) {
        $label = New-Object System.Windows.Forms.Label
        $label.Text = $Title
        $label.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
        $label.AutoSize = $true
        $label.Location = New-Object System.Drawing.Point(5, 5)
        $panel.Controls.Add($label)
    }
    
    return $panel
}

function New-ScrollableTextBox {
    param(
        [int]$Top = 30,
        [int]$Left = 5,
        [int]$Width = 590,
        [int]$Height = 300,
        [bool]$ReadOnly = $false,
        [string]$Text = ""
    )
    
    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Multiline = $true
    $textBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Both
    $textBox.Location = New-Object System.Drawing.Point($Left, $Top)
    $textBox.Size = New-Object System.Drawing.Size($Width, $Height)
    $textBox.ReadOnly = $ReadOnly
    $textBox.Text = $Text
    $textBox.Font = New-Object System.Drawing.Font("Consolas", 9)
    $textBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor 
                      [System.Windows.Forms.AnchorStyles]::Bottom -bor 
                      [System.Windows.Forms.AnchorStyles]::Left -bor 
                      [System.Windows.Forms.AnchorStyles]::Right
    
    return $textBox
}

function Show-StatusMessage {
    param(
        [string]$Message,
        [System.Windows.Forms.Label]$StatusLabel
    )
    
    if ($StatusLabel) {
        $StatusLabel.Text = $Message
        $StatusLabel.Refresh()
    }
    Write-Host $Message -ForegroundColor Cyan
}

#endregion

#region Event Handlers

function OnLoadTicket {
    param($sender, $e)
    
    try {
        # Prompt for ticket number
        $ticketNum = [Microsoft.VisualBasic.Interaction]::InputBox("Enter ConnectWise ticket number:", "Load Ticket", "")
        if ([string]::IsNullOrWhiteSpace($ticketNum)) { return }
        
        Show-StatusMessage -Message "Loading ticket #$ticketNum..." -StatusLabel $script:StatusLabel
        
        # Connect to ConnectWise if not already connected
        $cwStatus = Get-ConnectWiseStatus
        if (-not $cwStatus.Connected) {
            $connected = ConnectToConnectWise
            if (-not $connected) { return }
        }
        
        # Get ticket
        $ticket = Get-CWTicket -TicketNumber ([int]$ticketNum)
        if (-not $ticket) {
            [System.Windows.Forms.MessageBox]::Show("Failed to load ticket #$ticketNum", "Error", 
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
        
        $script:CurrentTicket = $ticket
        
        # Display ticket in UI
        $ticketText = @"
Ticket #$($ticket.id): $($ticket.summary)
Company: $($ticket.company.name)
Status: $($ticket.status.name)
Priority: $($ticket.priority.name)
Created: $($ticket.dateEntered)
Owner: $($ticket.owner.name)

Description:
$($ticket.initialDescription)
"@
        
        $script:TicketTextBox.Text = $ticketText
        $script:TicketContentForAI = $ticket.initialDescription
        
        # Enable investigation button
        $script:InvestigateButton.Enabled = $true
        
        Show-StatusMessage -Message "✓ Loaded ticket #$ticketNum" -StatusLabel $script:StatusLabel
        
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error loading ticket: $($_.Exception.Message)", "Error", 
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

function OnNewTicket {
    param($sender, $e)
    
    try {
        # Show form to create new ticket
        $form = New-Object System.Windows.Forms.Form
        $form.Text = "Create New Ticket"
        $form.Size = New-Object System.Drawing.Size(600, 500)
        $form.StartPosition = "CenterParent"
        $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
        $form.MaximizeBox = $false
        
        # Company
        $labelComp = New-Object System.Windows.Forms.Label
        $labelComp.Text = "Company Identifier:"
        $labelComp.Location = New-Object System.Drawing.Point(10, 20)
        $labelComp.AutoSize = $true
        $form.Controls.Add($labelComp)
        
        $txtCompany = New-Object System.Windows.Forms.TextBox
        $txtCompany.Location = New-Object System.Drawing.Point(150, 18)
        $txtCompany.Size = New-Object System.Drawing.Size(400, 20)
        $form.Controls.Add($txtCompany)
        
        # Summary
        $labelSummary = New-Object System.Windows.Forms.Label
        $labelSummary.Text = "Summary:"
        $labelSummary.Location = New-Object System.Drawing.Point(10, 60)
        $labelSummary.AutoSize = $true
        $form.Controls.Add($labelSummary)
        
        $txtSummary = New-Object System.Windows.Forms.TextBox
        $txtSummary.Location = New-Object System.Drawing.Point(150, 58)
        $txtSummary.Size = New-Object System.Drawing.Size(400, 20)
        $txtSummary.Text = "Security Alert Investigation"
        $form.Controls.Add($txtSummary)
        
        # Description
        $labelDesc = New-Object System.Windows.Forms.Label
        $labelDesc.Text = "Description:"
        $labelDesc.Location = New-Object System.Drawing.Point(10, 100)
        $labelDesc.AutoSize = $true
        $form.Controls.Add($labelDesc)
        
        $txtDesc = New-Object System.Windows.Forms.TextBox
        $txtDesc.Multiline = $true
        $txtDesc.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
        $txtDesc.Location = New-Object System.Drawing.Point(10, 120)
        $txtDesc.Size = New-Object System.Drawing.Size(560, 250)
        $form.Controls.Add($txtDesc)
        
        # Owner
        $labelOwner = New-Object System.Windows.Forms.Label
        $labelOwner.Text = "Assign To (optional):"
        $labelOwner.Location = New-Object System.Drawing.Point(10, 385)
        $labelOwner.AutoSize = $true
        $form.Controls.Add($labelOwner)
        
        $txtOwner = New-Object System.Windows.Forms.TextBox
        $txtOwner.Location = New-Object System.Drawing.Point(150, 383)
        $txtOwner.Size = New-Object System.Drawing.Size(200, 20)
        $txtOwner.Text = $script:Settings.InvestigatorName
        $form.Controls.Add($txtOwner)
        
        # Buttons
        $btnCreate = New-Object System.Windows.Forms.Button
        $btnCreate.Text = "Create && Investigate"
        $btnCreate.Location = New-Object System.Drawing.Point(350, 420)
        $btnCreate.Size = New-Object System.Drawing.Size(150, 30)
        $btnCreate.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.Controls.Add($btnCreate)
        $form.AcceptButton = $btnCreate
        
        $btnCancel = New-Object System.Windows.Forms.Button
        $btnCancel.Text = "Cancel"
        $btnCancel.Location = New-Object System.Drawing.Point(510, 420)
        $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $form.Controls.Add($btnCancel)
        $form.CancelButton = $btnCancel
        
        $result = $form.ShowDialog()
        
        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
            Show-StatusMessage -Message "Creating ticket..." -StatusLabel $script:StatusLabel
            
            # Connect to ConnectWise if needed
            $cwStatus = Get-ConnectWiseStatus
            if (-not $cwStatus.Connected) {
                $connected = ConnectToConnectWise
                if (-not $connected) { return }
            }
            
            # Create ticket
            $newTicketParams = @{
                Summary = $txtSummary.Text
                Company = $txtCompany.Text
                Description = $txtDesc.Text
                Status = "In Progress"
            }
            
            if (-not [string]::IsNullOrWhiteSpace($txtOwner.Text)) {
                $newTicketParams['AssignTo'] = $txtOwner.Text
            }
            
            $ticket = New-CWTicket @newTicketParams
            
            if ($ticket) {
                $script:CurrentTicket = $ticket
                
                # Display in UI
                $ticketText = @"
Ticket #$($ticket.id): $($ticket.summary)
Company: $($ticket.company.name)
Status: $($ticket.status.name)
Created: $($ticket.dateEntered)
Owner: $($ticket.owner.name)

Description:
$($txtDesc.Text)
"@
                
                $script:TicketTextBox.Text = $ticketText
                $script:TicketContentForAI = $txtDesc.Text
                $script:InvestigateButton.Enabled = $true
                
                Show-StatusMessage -Message "✓ Created ticket #$($ticket.id)" -StatusLabel $script:StatusLabel
            }
        }
        
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error creating ticket: $($_.Exception.Message)", "Error", 
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

function OnInvestigate {
    param($sender, $e)
    
    if ([string]::IsNullOrWhiteSpace($script:TicketContentForAI)) {
        [System.Windows.Forms.MessageBox]::Show("No ticket loaded. Load or create a ticket first.", "No Ticket", 
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    # Validate API key
    if ([string]::IsNullOrWhiteSpace($script:Settings.ClaudeApiKey)) {
        [System.Windows.Forms.MessageBox]::Show("Claude API key not configured. Please configure in Settings.", "Missing API Key", 
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    try {
        # Disable button during investigation
        $script:InvestigateButton.Enabled = $false
        $script:InvestigateButton.Text = "🔍 Investigating..."
        
        # Clear previous results
        $script:EvidenceTextBox.Text = "Starting investigation...`r`n`r`n"
        $script:VerdictTextBox.Text = ""
        $script:DraftTextBox.Text = ""
        
        Show-StatusMessage -Message "🔍 AI investigation in progress..." -StatusLabel $script:StatusLabel
        
        # Get Memberberry instructions if available
        $memberberryInstructions = ""
        if ($script:CurrentTicket -and $script:CurrentTicket.company) {
            $companyName = $script:CurrentTicket.company.name
            
            # Try to compile Memberberry instructions
            if (Get-Command Invoke-MemberberryCompile -ErrorAction SilentlyContinue) {
                try {
                    $mbPath = Invoke-MemberberryCompile -Client $companyName
                    if ($mbPath -and (Test-Path $mbPath)) {
                        $memberberryInstructions = Get-Content -Path $mbPath -Raw
                    }
                } catch {
                    Write-Warning "Could not load Memberberry instructions: $_"
                }
            }
        }
        
        # Status update callback
        $statusCallback = {
            param($message)
            $script:EvidenceTextBox.AppendText("$message`r`n")
            $script:EvidenceTextBox.Refresh()
            [System.Windows.Forms.Application]::DoEvents()
        }
        
        # Stream chunk callback for draft reply
        $streamCallback = {
            param($chunk)
            $script:DraftTextBox.AppendText($chunk)
            $script:DraftTextBox.Refresh()
            [System.Windows.Forms.Application]::DoEvents()
        }
        
        # Start investigation
        $investigation = Start-AIInvestigation `
            -TicketContent $script:TicketContentForAI `
            -CompanyName ($script:CurrentTicket.company.name) `
            -MemberberryInstructions $memberberryInstructions `
            -ApiKey $script:Settings.ClaudeApiKey `
            -OnStatusUpdate $statusCallback `
            -OnStreamChunk $streamCallback `
            -MaxIterations 10
        
        $script:CurrentInvestigation = $investigation
        
        # Display verdict
        $verdictColor = switch ($investigation.Verdict) {
            { $_ -match "False Positive" } { [System.Drawing.Color]::Green }
            { $_ -match "True Positive" } { [System.Drawing.Color]::Red }
            { $_ -match "Authorized" } { [System.Drawing.Color]::Orange }
            default { [System.Drawing.Color]::Gray }
        }
        
        $script:VerdictTextBox.BackColor = $verdictColor
        $script:VerdictTextBox.Text = @"
VERDICT: $($investigation.Verdict)
CONFIDENCE: $($investigation.Confidence)

SUMMARY:
$($investigation.Summary)

STATISTICS:
- Tool Calls: $($investigation.Stats.ToolCalls)
- Cache Hits: $($investigation.Stats.CacheHits)
- Duration: $([math]::Round($investigation.Stats.TotalDuration, 1))s
"@
        
        # Draft is already streamed to DraftTextBox
        # If not streamed, set it here
        if ([string]::IsNullOrWhiteSpace($script:DraftTextBox.Text)) {
            $script:DraftTextBox.Text = $investigation.ClientReply
        }
        
        # Enable editing and approval
        $script:DraftTextBox.ReadOnly = $false
        $script:ReviseButton.Enabled = $true
        $script:ApproveButton.Enabled = $true
        $script:ExportButton.Enabled = $true
        
        Show-StatusMessage -Message "✅ Investigation complete: $($investigation.Verdict)" -StatusLabel $script:StatusLabel
        
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Investigation failed: $($_.Exception.Message)", "Error", 
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        
        $script:EvidenceTextBox.AppendText("`r`nERROR: $($_.Exception.Message)`r`n")
        
    } finally {
        # Re-enable button
        $script:InvestigateButton.Enabled = $true
        $script:InvestigateButton.Text = "🔍 Start Investigation"
    }
}

function OnRevise {
    param($sender, $e)
    
    # Prompt for additional context
    $context = [Microsoft.VisualBasic.Interaction]::InputBox(
        "Provide additional context or clarification for the AI to revise the draft:",
        "Revise Draft",
        ""
    )
    
    if ([string]::IsNullOrWhiteSpace($context)) { return }
    
    # TODO: Call AI again with the additional context
    # For now, just append a note
    $script:DraftTextBox.AppendText("`r`n`r`n[Revision requested: $context]`r`n`r`n")
    
    [System.Windows.Forms.MessageBox]::Show("AI draft revision not yet implemented. Please edit manually.", "Info", 
        [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
}

function OnApprove {
    param($sender, $e)
    
    if (-not $script:CurrentTicket) {
        [System.Windows.Forms.MessageBox]::Show("No ticket loaded.", "Error", 
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    # Show approval dialog with status selection
    $approvalForm = New-Object System.Windows.Forms.Form
    $approvalForm.Text = "Approve & Post Reply"
    $approvalForm.Size = New-Object System.Drawing.Size(500, 250)
    $approvalForm.StartPosition = "CenterParent"
    $approvalForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $approvalForm.MaximizeBox = $false
    
    $label = New-Object System.Windows.Forms.Label
    $label.Text = "This will post the reply to ticket #$($script:CurrentTicket.id) and update the status."
    $label.Location = New-Object System.Drawing.Point(10, 20)
    $label.Size = New-Object System.Drawing.Size(460, 40)
    $approvalForm.Controls.Add($label)
    
    $labelStatus = New-Object System.Windows.Forms.Label
    $labelStatus.Text = "Set ticket status to:"
    $labelStatus.Location = New-Object System.Drawing.Point(10, 70)
    $labelStatus.AutoSize = $true
    $approvalForm.Controls.Add($labelStatus)
    
    $comboStatus = New-Object System.Windows.Forms.ComboBox
    $comboStatus.Location = New-Object System.Drawing.Point(150, 68)
    $comboStatus.Size = New-Object System.Drawing.Size(300, 20)
    $comboStatus.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $comboStatus.Items.AddRange(@("Waiting on Client", "Review", "In Progress", "Resolved"))
    $comboStatus.SelectedIndex = 0
    $approvalForm.Controls.Add($comboStatus)
    
    $checkNotify = New-Object System.Windows.Forms.CheckBox
    $checkNotify.Text = "Send email notification to client"
    $checkNotify.Location = New-Object System.Drawing.Point(10, 110)
    $checkNotify.Size = New-Object System.Drawing.Size(400, 25)
    $checkNotify.Checked = $true
    $approvalForm.Controls.Add($checkNotify)
    
    $btnPost = New-Object System.Windows.Forms.Button
    $btnPost.Text = "Post Reply"
    $btnPost.Location = New-Object System.Drawing.Point(250, 160)
    $btnPost.Size = New-Object System.Drawing.Size(100, 35)
    $btnPost.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $approvalForm.Controls.Add($btnPost)
    $approvalForm.AcceptButton = $btnPost
    
    $btnCancelPost = New-Object System.Windows.Forms.Button
    $btnCancelPost.Text = "Cancel"
    $btnCancelPost.Location = New-Object System.Drawing.Point(360, 160)
    $btnCancelPost.Size = New-Object System.Drawing.Size(100, 35)
    $btnCancelPost.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $approvalForm.Controls.Add($btnCancelPost)
    $approvalForm.CancelButton = $btnCancelPost
    
    $result = $approvalForm.ShowDialog()
    
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        try {
            Show-StatusMessage -Message "Posting reply..." -StatusLabel $script:StatusLabel
            
            # Connect to ConnectWise if needed
            $cwStatus = Get-ConnectWiseStatus
            if (-not $cwStatus.Connected) {
                $connected = ConnectToConnectWise
                if (-not $connected) { return }
            }
            
            # Post note
            $draftText = $script:DraftTextBox.Text
            $noteResult = Add-CWTicketNote `
                -TicketNumber $script:CurrentTicket.id `
                -Text $draftText `
                -InternalNote $false `
                -ProcessNotifications $checkNotify.Checked
            
            if (-not $noteResult) {
                throw "Failed to post note"
            }
            
            # Update status
            $statusName = $comboStatus.SelectedItem
            $statusResult = Set-CWTicketStatus `
                -TicketNumber $script:CurrentTicket.id `
                -StatusName $statusName
            
            if (-not $statusResult) {
                throw "Failed to update status"
            }
            
            [System.Windows.Forms.MessageBox]::Show(
                "Reply posted successfully!`n`nTicket #$($script:CurrentTicket.id) is now: $statusName",
                "Success",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
            
            Show-StatusMessage -Message "✅ Reply posted and status updated" -StatusLabel $script:StatusLabel
            
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Failed to post reply: $($_.Exception.Message)", "Error", 
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
}

function OnExport {
    param($sender, $e)
    
    if (-not $script:CurrentInvestigation) {
        [System.Windows.Forms.MessageBox]::Show("No investigation to export.", "Warning", 
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    try {
        # Save dialog
        $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveDialog.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
        $saveDialog.Title = "Export Investigation Report"
        $saveDialog.FileName = "Investigation_Ticket_$($script:CurrentTicket.id)_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
        
        if ($saveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $report = @"
THREAT HUNTER INVESTIGATION REPORT
Generated: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
Ticket: #$($script:CurrentTicket.id) - $($script:CurrentTicket.summary)
Company: $($script:CurrentTicket.company.name)

================================================================================
VERDICT
================================================================================
$($script:VerdictTextBox.Text)

================================================================================
EVIDENCE GATHERED
================================================================================
$($script:EvidenceTextBox.Text)

================================================================================
DRAFTED CLIENT REPLY
================================================================================
$($script:DraftTextBox.Text)

================================================================================
TOOL CALLS EXECUTED
================================================================================
$(
    $script:CurrentInvestigation.ToolCalls | ForEach-Object {
        "Tool: $($_.Tool)`nParameters: $($_.Parameters | ConvertTo-Json -Compress)`n---`n"
    }
)
"@
            
            $report | Out-File -FilePath $saveDialog.FileName -Encoding UTF8
            
            [System.Windows.Forms.MessageBox]::Show("Investigation report exported to:`n$($saveDialog.FileName)", "Success", 
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
        
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Export failed: $($_.Exception.Message)", "Error", 
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

function OnSettings {
    param($sender, $e)
    
    # Show settings dialog
    $settingsForm = New-Object System.Windows.Forms.Form
    $settingsForm.Text = "Settings"
    $settingsForm.Size = New-Object System.Drawing.Size(700, 600)
    $settingsForm.StartPosition = "CenterParent"
    $settingsForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $settingsForm.MaximizeBox = $false
    
    # Tabs
    $tabControl = New-Object System.Windows.Forms.TabControl
    $tabControl.Dock = [System.Windows.Forms.DockStyle]::Fill
    $settingsForm.Controls.Add($tabControl)
    
    # ConnectWise tab
    $cwTab = New-Object System.Windows.Forms.TabPage
    $cwTab.Text = "ConnectWise"
    $tabControl.Controls.Add($cwTab)
    
    # AI tab
    $aiTab = New-Object System.Windows.Forms.TabPage
    $aiTab.Text = "AI Settings"
    $tabControl.Controls.Add($aiTab)
    
    # Memberberry tab
    $mbTab = New-Object System.Windows.Forms.TabPage
    $mbTab.Text = "Memberberry"
    $tabControl.Controls.Add($mbTab)
    
    # TODO: Build settings UI
    # For now, just show a message
    $lblPlaceholder = New-Object System.Windows.Forms.Label
    $lblPlaceholder.Text = "Settings UI: Configure CW credentials, AI API keys, Memberberry path, etc.`n`nUse existing Settings.psm1 functions to save/load."
    $lblPlaceholder.Location = New-Object System.Drawing.Point(20, 20)
    $lblPlaceholder.Size = New-Object System.Drawing.Size(600, 100)
    $cwTab.Controls.Add($lblPlaceholder)
    
    # Buttons
    $btnSave = New-Object System.Windows.Forms.Button
    $btnSave.Text = "Save"
    $btnSave.Location = New-Object System.Drawing.Point(500, 520)
    $btnSave.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $settingsForm.Controls.Add($btnSave)
    
    $btnCancelSettings = New-Object System.Windows.Forms.Button
    $btnCancelSettings.Text = "Cancel"
    $btnCancelSettings.Location = New-Object System.Drawing.Point(590, 520)
    $btnCancelSettings.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $settingsForm.Controls.Add($btnCancelSettings)
    
    $settingsForm.ShowDialog() | Out-Null
}

function ConnectToConnectWise {
    # Try to load saved credentials
    $cwCompany = [Microsoft.VisualBasic.Interaction]::InputBox("Enter ConnectWise Company ID:", "ConnectWise Login", "")
    if ([string]::IsNullOrWhiteSpace($cwCompany)) { return $false }
    
    $cwCred = Get-CWCredential -CompanyID $cwCompany
    
    if ($cwCred) {
        # Use saved credentials
        $connected = Connect-ConnectWise `
            -CompanyID $cwCred.CompanyID `
            -PublicKey $cwCred.PublicKey `
            -PrivateKey $cwCred.PrivateKey `
            -BaseUrl $cwCred.BaseUrl
        
        return $connected
    } else {
        # Prompt for credentials
        $cwPublic = [Microsoft.VisualBasic.Interaction]::InputBox("Enter API Public Key:", "ConnectWise Login", "")
        if ([string]::IsNullOrWhiteSpace($cwPublic)) { return $false }
        
        $cwPrivate = Read-Host "Enter API Private Key" -AsSecureString
        
        $connected = Connect-ConnectWise `
            -CompanyID $cwCompany `
            -PublicKey $cwPublic `
            -PrivateKey $cwPrivate
        
        if ($connected) {
            # Offer to save credentials
            $save = [System.Windows.Forms.MessageBox]::Show(
                "Save ConnectWise credentials to Windows Credential Manager?",
                "Save Credentials",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question
            )
            
            if ($save -eq [System.Windows.Forms.DialogResult]::Yes) {
                Save-CWCredential -CompanyID $cwCompany -PublicKey $cwPublic -PrivateKey $cwPrivate | Out-Null
            }
        }
        
        return $connected
    }
}

#endregion

#region Main Form

function Show-ThreatHunterApp {
    # Load settings
    $script:Settings = Get-AppSettings
    
    # Create main form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Threat Hunter - AI Investigation Platform"
    $form.Size = New-Object System.Drawing.Size(1400, 900)
    $form.StartPosition = "CenterScreen"
    $form.MinimumSize = New-Object System.Drawing.Size(1000, 700)
    
    # Add VisualBasic assembly for InputBox
    Add-Type -AssemblyName Microsoft.VisualBasic
    
    # Top toolbar
    $toolbar = New-Object System.Windows.Forms.Panel
    $toolbar.Dock = [System.Windows.Forms.DockStyle]::Top
    $toolbar.Height = 50
    $toolbar.BackColor = [System.Drawing.Color]::WhiteSmoke
    $form.Controls.Add($toolbar)
    
    $btnNewTicket = New-Object System.Windows.Forms.Button
    $btnNewTicket.Text = "📝 New Ticket"
    $btnNewTicket.Location = New-Object System.Drawing.Point(10, 10)
    $btnNewTicket.Size = New-Object System.Drawing.Size(120, 30)
    $btnNewTicket.Add_Click({ OnNewTicket $this $_ })
    $toolbar.Controls.Add($btnNewTicket)
    
    $btnLoadTicket = New-Object System.Windows.Forms.Button
    $btnLoadTicket.Text = "📂 Load Ticket"
    $btnLoadTicket.Location = New-Object System.Drawing.Point(140, 10)
    $btnLoadTicket.Size = New-Object System.Drawing.Size(120, 30)
    $btnLoadTicket.Add_Click({ OnLoadTicket $this $_ })
    $toolbar.Controls.Add($btnLoadTicket)
    
    $btnSettings = New-Object System.Windows.Forms.Button
    $btnSettings.Text = "⚙️ Settings"
    $btnSettings.Location = New-Object System.Drawing.Point(270, 10)
    $btnSettings.Size = New-Object System.Drawing.Size(120, 30)
    $btnSettings.Add_Click({ OnSettings $this $_ })
    $toolbar.Controls.Add($btnSettings)
    
    # Main split container (left/right)
    $mainSplit = New-Object System.Windows.Forms.SplitContainer
    $mainSplit.Dock = [System.Windows.Forms.DockStyle]::Fill
    $mainSplit.Orientation = [System.Windows.Forms.Orientation]::Vertical
    $mainSplit.SplitterDistance = 600
    $mainSplit.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $form.Controls.Add($mainSplit)
    
    # Left side split (ticket top, evidence bottom)
    $leftSplit = New-Object System.Windows.Forms.SplitContainer
    $leftSplit.Dock = [System.Windows.Forms.DockStyle]::Fill
    $leftSplit.Orientation = [System.Windows.Forms.Orientation]::Horizontal
    $leftSplit.SplitterDistance = 300
    $leftSplit.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $mainSplit.Panel1.Controls.Add($leftSplit)
    
    # Ticket Info Panel (top left)
    $ticketPanel = New-Object System.Windows.Forms.Panel
    $ticketPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
    $leftSplit.Panel1.Controls.Add($ticketPanel)
    
    $lblTicket = New-Object System.Windows.Forms.Label
    $lblTicket.Text = "📋 Ticket Information"
    $lblTicket.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $lblTicket.Location = New-Object System.Drawing.Point(10, 10)
    $lblTicket.AutoSize = $true
    $ticketPanel.Controls.Add($lblTicket)
    
    $script:TicketTextBox = New-ScrollableTextBox -Top 40 -Left 10 -Width 570 -Height 250 -ReadOnly $true
    $ticketPanel.Controls.Add($script:TicketTextBox)
    
    # Evidence Panel (bottom left)
    $evidencePanel = New-Object System.Windows.Forms.Panel
    $evidencePanel.Dock = [System.Windows.Forms.DockStyle]::Fill
    $leftSplit.Panel2.Controls.Add($evidencePanel)
    
    $lblEvidence = New-Object System.Windows.Forms.Label
    $lblEvidence.Text = "🔎 Investigation Evidence"
    $lblEvidence.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $lblEvidence.Location = New-Object System.Drawing.Point(10, 10)
    $lblEvidence.AutoSize = $true
    $evidencePanel.Controls.Add($lblEvidence)
    
    $script:EvidenceTextBox = New-ScrollableTextBox -Top 40 -Left 10 -Width 570 -Height 450 -ReadOnly $true
    $evidencePanel.Controls.Add($script:EvidenceTextBox)
    
    # Right side split (verdict top, draft bottom)
    $rightSplit = New-Object System.Windows.Forms.SplitContainer
    $rightSplit.Dock = [System.Windows.Forms.DockStyle]::Fill
    $rightSplit.Orientation = [System.Windows.Forms.Orientation]::Horizontal
    $rightSplit.SplitterDistance = 250
    $rightSplit.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $mainSplit.Panel2.Controls.Add($rightSplit)
    
    # Verdict Panel (top right)
    $verdictPanel = New-Object System.Windows.Forms.Panel
    $verdictPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
    $rightSplit.Panel1.Controls.Add($verdictPanel)
    
    $lblVerdict = New-Object System.Windows.Forms.Label
    $lblVerdict.Text = "⚖️ AI Verdict"
    $lblVerdict.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $lblVerdict.Location = New-Object System.Drawing.Point(10, 10)
    $lblVerdict.AutoSize = $true
    $verdictPanel.Controls.Add($lblVerdict)
    
    $script:VerdictTextBox = New-ScrollableTextBox -Top 40 -Left 10 -Width 740 -Height 190 -ReadOnly $true
    $script:VerdictTextBox.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $verdictPanel.Controls.Add($script:VerdictTextBox)
    
    # Draft Reply Panel (bottom right)
    $draftPanel = New-Object System.Windows.Forms.Panel
    $draftPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
    $rightSplit.Panel2.Controls.Add($draftPanel)
    
    $lblDraft = New-Object System.Windows.Forms.Label
    $lblDraft.Text = "✍️ Draft Client Reply (Editable)"
    $lblDraft.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $lblDraft.Location = New-Object System.Drawing.Point(10, 10)
    $lblDraft.AutoSize = $true
    $draftPanel.Controls.Add($lblDraft)
    
    $script:DraftTextBox = New-ScrollableTextBox -Top 40 -Left 10 -Width 740 -Height 390 -ReadOnly $false
    $script:DraftTextBox.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $draftPanel.Controls.Add($script:DraftTextBox)
    
    # Bottom control panel
    $controlPanel = New-Object System.Windows.Forms.Panel
    $controlPanel.Dock = [System.Windows.Forms.DockStyle]::Bottom
    $controlPanel.Height = 70
    $controlPanel.BackColor = [System.Drawing.Color]::WhiteSmoke
    $form.Controls.Add($controlPanel)
    
    $script:InvestigateButton = New-Object System.Windows.Forms.Button
    $script:InvestigateButton.Text = "🔍 Start Investigation"
    $script:InvestigateButton.Location = New-Object System.Drawing.Point(20, 15)
    $script:InvestigateButton.Size = New-Object System.Drawing.Size(180, 40)
    $script:InvestigateButton.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $script:InvestigateButton.BackColor = [System.Drawing.Color]::LightSkyBlue
    $script:InvestigateButton.Enabled = $false
    $script:InvestigateButton.Add_Click({ OnInvestigate $this $_ })
    $controlPanel.Controls.Add($script:InvestigateButton)
    
    $script:ReviseButton = New-Object System.Windows.Forms.Button
    $script:ReviseButton.Text = "🔄 Revise Draft"
    $script:ReviseButton.Location = New-Object System.Drawing.Point(220, 15)
    $script:ReviseButton.Size = New-Object System.Drawing.Size(150, 40)
    $script:ReviseButton.Enabled = $false
    $script:ReviseButton.Add_Click({ OnRevise $this $_ })
    $controlPanel.Controls.Add($script:ReviseButton)
    
    $script:ApproveButton = New-Object System.Windows.Forms.Button
    $script:ApproveButton.Text = "✅ Approve & Post"
    $script:ApproveButton.Location = New-Object System.Drawing.Point(390, 15)
    $script:ApproveButton.Size = New-Object System.Drawing.Size(170, 40)
    $script:ApproveButton.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $script:ApproveButton.BackColor = [System.Drawing.Color]::LightGreen
    $script:ApproveButton.Enabled = $false
    $script:ApproveButton.Add_Click({ OnApprove $this $_ })
    $controlPanel.Controls.Add($script:ApproveButton)
    
    $script:ExportButton = New-Object System.Windows.Forms.Button
    $script:ExportButton.Text = "💾 Export Report"
    $script:ExportButton.Location = New-Object System.Drawing.Point(580, 15)
    $script:ExportButton.Size = New-Object System.Drawing.Size(150, 40)
    $script:ExportButton.Enabled = $false
    $script:ExportButton.Add_Click({ OnExport $this $_ })
    $controlPanel.Controls.Add($script:ExportButton)
    
    # Status label
    $script:StatusLabel = New-Object System.Windows.Forms.Label
    $script:StatusLabel.Text = "Ready. Load or create a ticket to begin."
    $script:StatusLabel.Location = New-Object System.Drawing.Point(750, 25)
    $script:StatusLabel.Size = New-Object System.Drawing.Size(600, 30)
    $script:StatusLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)
    $script:StatusLabel.ForeColor = [System.Drawing.Color]::DarkSlateGray
    $controlPanel.Controls.Add($script:StatusLabel)
    
    # Module scope variables
    $script:TicketContentForAI = ""
    
    # Show form
    [void]$form.ShowDialog()
}

#endregion

# Entry point
try {
    Show-ThreatHunterApp
} catch {
    [System.Windows.Forms.MessageBox]::Show(
        "Fatal error: $($_.Exception.Message)`n`n$($_.ScriptStackTrace)",
        "Error",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    )
}
