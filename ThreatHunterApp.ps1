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
        [bool]$ReadOnly = $false
    )
    
    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Multiline = $true
    $textBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
    $textBox.WordWrap = $true
    $textBox.Dock = [System.Windows.Forms.DockStyle]::Fill
    $textBox.ReadOnly = $ReadOnly
    $textBox.Font = New-Object System.Drawing.Font("Consolas", 9)
    return $textBox
}

function Add-LabeledFillBox {
    param($Panel, $Title, $ReadOnly)
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $Title
    $label.Dock = [System.Windows.Forms.DockStyle]::Top
    $label.Height = 28
    $label.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $label.Padding = New-Object System.Windows.Forms.Padding(8, 6, 0, 0)
    $box = New-ScrollableTextBox -ReadOnly $ReadOnly
    # Top is added first so it claims its strip before Fill takes the rest.
    $Panel.Controls.Add($label)
    $Panel.Controls.Add($box)
    return $box
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
        $notes = @(Get-CWTicketNotes -TicketNumber $ticket.id)
        $noteText = New-Object System.Text.StringBuilder
        foreach ($note in $notes) {
            if ([string]::IsNullOrWhiteSpace($note.text)) { continue }
            $kind = if ($note.internalAnalysisFlag) { 'Internal' } else { 'Detail' }
            [void]$noteText.AppendLine("")
            [void]$noteText.AppendLine("[$kind note]")
            [void]$noteText.AppendLine([string]$note.text)
        }

        $memberId = [string]$script:Settings.ConnectWiseMemberIdentifier
        if (-not [string]::IsNullOrWhiteSpace($memberId)) {
            $ownerResult = Set-CWTicketOwner -TicketNumber $ticket.id -MemberIdentifier $memberId
            if ($ownerResult) { $ticket = $ownerResult }
        }
        $statusResult = Set-CWTicketStatus -TicketNumber $ticket.id -StatusName 'In Progress'
        if ($statusResult) { $ticket = $statusResult }
        $script:CurrentTicket = $ticket
        
        $script:TicketTextBox.Text = @"
Ticket #$($ticket.id): $($ticket.summary)
Company: $($ticket.company.name)
Status: $($ticket.status.name)
Priority: $($ticket.priority.name)
Created: $($ticket.dateEntered)
Owner: $($ticket.owner.name)

Description:
$($ticket.initialDescription)
$($noteText.ToString())
"@
        $script:TicketContentForAI = (($ticket.initialDescription + "`n" + $noteText.ToString())).Trim()
        if ([string]::IsNullOrWhiteSpace($script:TicketContentForAI)) {
            $script:TicketContentForAI = [string]$ticket.summary
        }
        
        if ($statusResult) {
            $ticket = $statusResult
            Show-StatusMessage -Message "Loaded ticket #$ticketNum and set it to In Progress" -StatusLabel $script:StatusLabel
        } else {
            Show-StatusMessage -Message "Loaded ticket #$ticketNum. Status was not changed; confirm the board has a status named In Progress." -StatusLabel $script:StatusLabel
        }
        $script:CurrentTicket = $ticket
        $script:InvestigateButton.Enabled = $true
        
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
        $txtOwner.Text = [string]$script:Settings.ConnectWiseMemberIdentifier
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
        
        $memberberryInstructions = ""
        $ticketForAi = $script:TicketContentForAI
        $companyName = ""
        if ($script:CurrentTicket -and $script:CurrentTicket.company) { $companyName = [string]$script:CurrentTicket.company.name }
        if (Get-Command Invoke-MemberberryCleanTicket -ErrorAction SilentlyContinue) {
            try { $ticketForAi = Invoke-MemberberryCleanTicket -TicketContent $ticketForAi } catch {}
        }
        if (Get-Command Invoke-MemberberryExtractCompany -ErrorAction SilentlyContinue -and [string]::IsNullOrWhiteSpace($companyName)) {
            try { $companyName = Invoke-MemberberryExtractCompany -TicketContent $ticketForAi } catch {}
        }
        $alertType = ""
        if (Get-Command Invoke-MemberberryDetectAlertType -ErrorAction SilentlyContinue) {
            try { $alertType = Invoke-MemberberryDetectAlertType -TicketContent $ticketForAi } catch {}
        }
        if (Get-Command Invoke-MemberberryCompile -ErrorAction SilentlyContinue) {
            try {
                $compileParams = @{}
                if ($companyName) { $compileParams.Client = $companyName }
                if ($alertType) { $compileParams.AlertType = $alertType }
                $mbPath = Invoke-MemberberryCompile @compileParams
                if ($mbPath -and (Test-Path $mbPath)) {
                    $memberberryInstructions = Get-Content -Path $mbPath -Raw
                    if ($memberberryInstructions.Length -gt 6000) {
                        $memberberryInstructions = $memberberryInstructions.Substring(0, 6000) + "`n[Memberberry instructions truncated.]"
                    }
                }
            } catch {
                Write-Warning "Could not load Memberberry instructions: $_"
            }
        }
        
        $statusCallback = {
            param($message)
            $script:EvidenceTextBox.AppendText("$message`r`n")
            $script:EvidenceTextBox.SelectionStart = $script:EvidenceTextBox.TextLength
            $script:EvidenceTextBox.ScrollToCaret()
            [System.Windows.Forms.Application]::DoEvents()
        }
        
        # Narration goes to the evidence pane. The draft is set from the parsed client reply after the run.
        $streamCallback = {
            param($chunk)
            $script:EvidenceTextBox.AppendText($chunk)
            [System.Windows.Forms.Application]::DoEvents()
        }
        
        $model = [string]$script:Settings.ClaudeModel
        $investigation = Start-AIInvestigation `
            -TicketContent $ticketForAi `
            -CompanyName $companyName `
            -MemberberryInstructions $memberberryInstructions `
            -ApiKey $script:Settings.ClaudeApiKey `
            -Model $model `
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
        
        $script:DraftTextBox.Text = [string]$investigation.ClientReply
        
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
    if ([string]::IsNullOrWhiteSpace($script:Settings.ClaudeApiKey)) {
        [System.Windows.Forms.MessageBox]::Show("Claude API key not configured.", "Missing API Key", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }

    try {
        $script:ReviseButton.Enabled = $false
        $script:EvidenceTextBox.AppendText("`r`nRevising draft...`r`n")
        $revised = Update-AIClientDraft `
            -ApiKey $script:Settings.ClaudeApiKey `
            -CurrentDraft $script:DraftTextBox.Text `
            -Clarification $context `
            -Model ([string]$script:Settings.ClaudeModel)
        if (-not [string]::IsNullOrWhiteSpace($revised)) {
            $script:DraftTextBox.Text = $revised
            Show-StatusMessage -Message "Draft revised. Nothing has been posted." -StatusLabel $script:StatusLabel
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Revision failed: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    } finally {
        $script:ReviseButton.Enabled = $true
    }
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

function Read-MaskedSecret {
    param([string]$Title, [string]$Prompt)
    $dialog = New-Object System.Windows.Forms.Form
    $dialog.Text = $Title
    $dialog.Size = New-Object System.Drawing.Size(460, 170)
    $dialog.StartPosition = 'CenterParent'
    $dialog.FormBorderStyle = 'FixedDialog'
    $dialog.MaximizeBox = $false
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $Prompt
    $label.Location = New-Object System.Drawing.Point(12, 15)
    $label.Size = New-Object System.Drawing.Size(420, 32)
    $box = New-Object System.Windows.Forms.TextBox
    $box.Location = New-Object System.Drawing.Point(12, 50)
    $box.Size = New-Object System.Drawing.Size(420, 24)
    $box.UseSystemPasswordChar = $true
    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = 'OK'
    $ok.DialogResult = 'OK'
    $ok.Location = New-Object System.Drawing.Point(250, 90)
    $cancel = New-Object System.Windows.Forms.Button
    $cancel.Text = 'Cancel'
    $cancel.DialogResult = 'Cancel'
    $cancel.Location = New-Object System.Drawing.Point(340, 90)
    $dialog.Controls.AddRange(@($label, $box, $ok, $cancel))
    $dialog.AcceptButton = $ok
    $dialog.CancelButton = $cancel
    if ($dialog.ShowDialog() -ne 'OK' -or [string]::IsNullOrEmpty($box.Text)) { return $null }
    return (ConvertTo-SecureString -String $box.Text -AsPlainText -Force)
}

function OnSettings {
    param($sender, $e)
    $settings = $script:Settings
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Threat Hunter Settings'
    $form.Size = New-Object System.Drawing.Size(640, 460)
    $form.StartPosition = 'CenterParent'
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false

    $fields = @(
        @{ Name = 'ClaudeApiKey'; Label = 'Claude API key'; Secret = $true },
        @{ Name = 'ClaudeModel'; Label = 'Claude model'; Secret = $false },
        @{ Name = 'ConnectWiseCompanyID'; Label = 'ConnectWise company id'; Secret = $false },
        @{ Name = 'ConnectWiseClientId'; Label = 'ConnectWise integrator client id'; Secret = $false },
        @{ Name = 'ConnectWiseMemberIdentifier'; Label = 'ConnectWise member identifier'; Secret = $false },
        @{ Name = 'ConnectWiseBaseUrl'; Label = 'ConnectWise base URL'; Secret = $false },
        @{ Name = 'MemberberryPath'; Label = 'Memberberry path'; Secret = $false }
    )
    $boxes = @{}
    $y = 16
    foreach ($field in $fields) {
        $label = New-Object System.Windows.Forms.Label
        $label.Text = $field.Label
        $label.Location = New-Object System.Drawing.Point(12, $y)
        $label.Size = New-Object System.Drawing.Size(250, 22)
        $box = New-Object System.Windows.Forms.TextBox
        $box.Location = New-Object System.Drawing.Point(270, ($y - 2))
        $box.Size = New-Object System.Drawing.Size(340, 24)
        $box.Text = [string]$settings.($field.Name)
        if ($field.Secret) { $box.UseSystemPasswordChar = $true }
        $form.Controls.Add($label)
        $form.Controls.Add($box)
        $boxes[$field.Name] = $box
        $y += 36
    }
    $publicLabel = New-Object System.Windows.Forms.Label
    $publicLabel.Text = 'CW API public key (saved to Credential Manager)'
    $publicLabel.Location = New-Object System.Drawing.Point(12, $y)
    $publicLabel.Size = New-Object System.Drawing.Size(250, 36)
    $publicBox = New-Object System.Windows.Forms.TextBox
    $publicBox.Location = New-Object System.Drawing.Point(270, $y)
    $publicBox.Size = New-Object System.Drawing.Size(340, 24)
    $y += 40
    $privateLabel = New-Object System.Windows.Forms.Label
    $privateLabel.Text = 'CW API private key (leave blank to keep saved key)'
    $privateLabel.Location = New-Object System.Drawing.Point(12, $y)
    $privateLabel.Size = New-Object System.Drawing.Size(250, 36)
    $privateBox = New-Object System.Windows.Forms.TextBox
    $privateBox.Location = New-Object System.Drawing.Point(270, $y)
    $privateBox.Size = New-Object System.Drawing.Size(340, 24)
    $privateBox.UseSystemPasswordChar = $true
    $form.Controls.AddRange(@($publicLabel, $publicBox, $privateLabel, $privateBox))

    $save = New-Object System.Windows.Forms.Button
    $save.Text = 'Save'
    $save.DialogResult = 'OK'
    $save.Location = New-Object System.Drawing.Point(420, 380)
    $cancel = New-Object System.Windows.Forms.Button
    $cancel.Text = 'Cancel'
    $cancel.DialogResult = 'Cancel'
    $cancel.Location = New-Object System.Drawing.Point(510, 380)
    $form.Controls.AddRange(@($save, $cancel))
    $form.AcceptButton = $save
    $form.CancelButton = $cancel
    if ($form.ShowDialog() -ne 'OK') { return }

    foreach ($field in $fields) {
        $settings.($field.Name) = $boxes[$field.Name].Text.Trim()
    }
    $script:Settings = $settings
    Save-AppSettings -Settings $settings | Out-Null

    if (-not [string]::IsNullOrWhiteSpace($publicBox.Text) -and -not [string]::IsNullOrWhiteSpace($privateBox.Text)) {
        $secret = ConvertTo-SecureString -String $privateBox.Text -AsPlainText -Force
        Save-CWCredential -CompanyID $settings.ConnectWiseCompanyID -PublicKey $publicBox.Text.Trim() -PrivateKey $secret -ClientId $settings.ConnectWiseClientId -BaseUrl $settings.ConnectWiseBaseUrl | Out-Null
    }
    Show-StatusMessage -Message 'Settings saved.' -StatusLabel $script:StatusLabel
}

function ConnectToConnectWise {
    $company = [string]$script:Settings.ConnectWiseCompanyID
    $clientId = [string]$script:Settings.ConnectWiseClientId
    $baseUrl = [string]$script:Settings.ConnectWiseBaseUrl
    if ([string]::IsNullOrWhiteSpace($baseUrl)) { $baseUrl = 'https://api-na.myconnectwise.net' }
    if ([string]::IsNullOrWhiteSpace($company)) {
        $company = [Microsoft.VisualBasic.Interaction]::InputBox('Enter ConnectWise company id:', 'ConnectWise Login', '')
    }
    if ([string]::IsNullOrWhiteSpace($company)) { return $false }
    if ([string]::IsNullOrWhiteSpace($clientId)) {
        [System.Windows.Forms.MessageBox]::Show('Set the ConnectWise integrator client id in Settings. That value comes from the ConnectWise Developer Network and is not the API public key.', 'ConnectWise', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return $false
    }

    $cwCred = Get-CWCredential -CompanyID $company
    if ($cwCred -and [string]::IsNullOrWhiteSpace($cwCred.ClientId)) { $cwCred.ClientId = $clientId }
    if ($cwCred -and $cwCred.PrivateKey -and $cwCred.PublicKey) {
        return [bool](Connect-ConnectWise -CompanyID $cwCred.CompanyID -PublicKey $cwCred.PublicKey -PrivateKey $cwCred.PrivateKey -ClientId $cwCred.ClientId -BaseUrl $(if ($cwCred.BaseUrl) { $cwCred.BaseUrl } else { $baseUrl }))
    }

    $cwPublic = [Microsoft.VisualBasic.Interaction]::InputBox('Enter API public key:', 'ConnectWise Login', '')
    if ([string]::IsNullOrWhiteSpace($cwPublic)) { return $false }
    $cwPrivate = Read-MaskedSecret -Title 'ConnectWise' -Prompt 'Enter the API private key'
    if (-not $cwPrivate) { return $false }
    $connected = Connect-ConnectWise -CompanyID $company -PublicKey $cwPublic -PrivateKey $cwPrivate -ClientId $clientId -BaseUrl $baseUrl
    if ($connected) {
        $save = [System.Windows.Forms.MessageBox]::Show('Save ConnectWise credentials to Windows Credential Manager?', 'Save Credentials', [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($save -eq [System.Windows.Forms.DialogResult]::Yes) {
            Save-CWCredential -CompanyID $company -PublicKey $cwPublic -PrivateKey $cwPrivate -ClientId $clientId -BaseUrl $baseUrl | Out-Null
        }
    }
    return [bool]$connected
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
    
    # Edge bars are added before the Fill split so they keep their strips.
    $controlPanel = New-Object System.Windows.Forms.Panel
    $controlPanel.Dock = [System.Windows.Forms.DockStyle]::Bottom
    $controlPanel.Height = 70
    $controlPanel.BackColor = [System.Drawing.Color]::WhiteSmoke
    $form.Controls.Add($controlPanel)

    $mainSplit = New-Object System.Windows.Forms.SplitContainer
    $mainSplit.Dock = [System.Windows.Forms.DockStyle]::Fill
    $mainSplit.Orientation = [System.Windows.Forms.Orientation]::Vertical
    $mainSplit.Panel1MinSize = 120
    $mainSplit.Panel2MinSize = 120
    $mainSplit.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    
    $leftSplit = New-Object System.Windows.Forms.SplitContainer
    $leftSplit.Dock = [System.Windows.Forms.DockStyle]::Fill
    $leftSplit.Orientation = [System.Windows.Forms.Orientation]::Horizontal
    $leftSplit.Panel1MinSize = 80
    $leftSplit.Panel2MinSize = 80
    $leftSplit.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $mainSplit.Panel1.Controls.Add($leftSplit)
    
    $ticketPanel = New-Object System.Windows.Forms.Panel
    $ticketPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
    $leftSplit.Panel1.Controls.Add($ticketPanel)
    $script:TicketTextBox = Add-LabeledFillBox -Panel $ticketPanel -Title "Ticket Information" -ReadOnly $true
    
    $evidencePanel = New-Object System.Windows.Forms.Panel
    $evidencePanel.Dock = [System.Windows.Forms.DockStyle]::Fill
    $leftSplit.Panel2.Controls.Add($evidencePanel)
    $script:EvidenceTextBox = Add-LabeledFillBox -Panel $evidencePanel -Title "Investigation Evidence" -ReadOnly $true
    
    $rightSplit = New-Object System.Windows.Forms.SplitContainer
    $rightSplit.Dock = [System.Windows.Forms.DockStyle]::Fill
    $rightSplit.Orientation = [System.Windows.Forms.Orientation]::Horizontal
    $rightSplit.Panel1MinSize = 80
    $rightSplit.Panel2MinSize = 80
    $rightSplit.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $mainSplit.Panel2.Controls.Add($rightSplit)
    
    $verdictPanel = New-Object System.Windows.Forms.Panel
    $verdictPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
    $rightSplit.Panel1.Controls.Add($verdictPanel)
    $script:VerdictTextBox = Add-LabeledFillBox -Panel $verdictPanel -Title "AI Verdict" -ReadOnly $true
    $script:VerdictTextBox.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    
    $draftPanel = New-Object System.Windows.Forms.Panel
    $draftPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
    $rightSplit.Panel2.Controls.Add($draftPanel)
    $script:DraftTextBox = Add-LabeledFillBox -Panel $draftPanel -Title "Draft Client Reply (editable, not posted)" -ReadOnly $false
    $script:DraftTextBox.Font = New-Object System.Drawing.Font("Segoe UI", 10)

    $form.Controls.Add($mainSplit)
    $script:MainSplit = $mainSplit
    $script:LeftSplit = $leftSplit
    $script:RightSplit = $rightSplit
    $form.Add_Shown({
        foreach ($pair in @(
            @{ Split = $script:MainSplit; Desired = 640 },
            @{ Split = $script:LeftSplit; Desired = 260 },
            @{ Split = $script:RightSplit; Desired = 200 }
        )) {
            $split = $pair.Split
            $limit = if ($split.Orientation -eq [System.Windows.Forms.Orientation]::Vertical) { $split.Width } else { $split.Height }
            $max = $limit - $split.Panel2MinSize - $split.SplitterWidth
            $min = $split.Panel1MinSize
            if ($max -gt $min) {
                $distance = [Math]::Max($min, [Math]::Min([int]$pair.Desired, $max))
                $split.SplitterDistance = $distance
            }
        }
    })
    
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
