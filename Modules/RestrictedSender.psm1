function Show-RestrictedSenderManagementDialog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [System.Windows.Forms.Form]$ParentForm,
        
        [Parameter(Mandatory=$true)]
        [System.Windows.Forms.ToolStripStatusLabel]$StatusLabelGlobal
    )
    
    # Create the dialog form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Manage Blocked Sender Addresses"
    $form.Size = New-Object System.Drawing.Size(900, 600)
    $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
    $form.MaximizeBox = $true
    $form.MinimizeBox = $true
    
    # Create list view for blocked senders
    $listView = New-Object System.Windows.Forms.ListView
    $listView.Location = New-Object System.Drawing.Point(10, 50)
    $listView.Size = New-Object System.Drawing.Size(860, 450)
    $listView.View = 'Details'
    $listView.FullRowSelect = $true
    $listView.GridLines = $true
    $listView.MultiSelect = $false
    
    # Add columns
    $listView.Columns.Add("Email Address", 250)
    $listView.Columns.Add("Reason", 400)
    $listView.Columns.Add("Created", 120)
    $listView.Columns.Add("Status", 100)
    
    # Create refresh button
    $refreshButton = New-Object System.Windows.Forms.Button
    $refreshButton.Text = "Refresh List"
    $refreshButton.Location = New-Object System.Drawing.Point(10, 10)
    $refreshButton.Size = New-Object System.Drawing.Size(100, 30)
    
    # Create remove button
    $removeButton = New-Object System.Windows.Forms.Button
    $removeButton.Text = "Remove Selected"
    $removeButton.Location = New-Object System.Drawing.Point(120, 10)
    $removeButton.Size = New-Object System.Drawing.Size(120, 30)
    $removeButton.BackColor = [System.Drawing.Color]::LightCoral
    
    # Create close button
    $closeButton = New-Object System.Windows.Forms.Button
    $closeButton.Text = "Close"
    $closeButton.Location = New-Object System.Drawing.Point(250, 10)
    $closeButton.Size = New-Object System.Drawing.Size(100, 30)
    
    # Create web link button for Microsoft Defender Restricted Users
    $defenderButton = New-Object System.Windows.Forms.Button
    $defenderButton.Text = "Open Defender Restricted Users"
    $defenderButton.Location = New-Object System.Drawing.Point(360, 10)
    $defenderButton.Size = New-Object System.Drawing.Size(180, 30)
    $defenderButton.BackColor = [System.Drawing.Color]::LightBlue
    
    # Add controls to form
    $form.Controls.AddRange(@($listView, $refreshButton, $removeButton, $closeButton, $defenderButton))
    
    # Function to load blocked senders
    $loadBlockedSenders = {
        $listView.Items.Clear()
        $StatusLabelGlobal.Text = "Loading blocked sender addresses..."
        
        try {
            # Check if connected to Exchange Online - try multiple methods
            $exchangeSession = Get-PSSession | Where-Object { $_.ConfigurationName -eq "Microsoft.Exchange" -and $_.State -eq "Opened" }
            $isConnected = $false
            
            if ($exchangeSession) {
                $isConnected = $true
            } else {
                # Try to test if we can run Exchange Online cmdlets
                try {
                    $testConnection = Get-Mailbox -ResultSize 1 -ErrorAction Stop
                    $isConnected = $true
                } catch {
                    $isConnected = $false
                }
            }
            
            if (-not $isConnected) {
                [System.Windows.Forms.MessageBox]::Show("Please connect to Exchange Online first.`n`nThese cmdlets require Exchange Online connection and specific admin roles.", "Not Connected", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                $StatusLabelGlobal.Text = "Not connected to Exchange Online"
                return
            }
            
            # Check if user has required roles
            try {
                $roleCheck = Get-ManagementRoleAssignment -Role "Transport Hygiene" -ErrorAction Stop
                if (-not $roleCheck) {
                    throw "Transport Hygiene role not found"
                }
            } catch {
                $roleError = "`n`nRequired Admin Roles:`n- Security Administrator`n- Global Administrator`n- Organization Management (Exchange role)`n- Compliance Administrator`n`nTo check your roles, run:`nGet-ManagementRoleAssignment -Role 'Transport Hygiene'`n`nOr check via Microsoft 365 admin center under Roles > Security Administrator"
                [System.Windows.Forms.MessageBox]::Show("Permission denied. You need specific admin roles to manage blocked senders.$roleError", "Insufficient Permissions", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                $StatusLabelGlobal.Text = "Insufficient permissions for blocked sender management"
                return
            }
            
            # Get blocked sender addresses
            $blockedSenders = Get-BlockedSenderAddress -ErrorAction Stop
            
            if ($blockedSenders -and $blockedSenders.Count -gt 0) {
                foreach ($sender in $blockedSenders) {
                    $item = $listView.Items.Add($sender.SenderAddress)
                    $item.SubItems.Add($sender.Reason)
                    $item.SubItems.Add($sender.DateTimeCreated.ToString("yyyy-MM-dd"))
                    $item.SubItems.Add("Blocked")
                }
                
                $StatusLabelGlobal.Text = "Loaded $($blockedSenders.Count) blocked sender addresses"
            } else {
                $StatusLabelGlobal.Text = "No blocked sender addresses found"
                [System.Windows.Forms.MessageBox]::Show("No blocked sender addresses found.", "No Results", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            }
        } catch {
            $errorMsg = $_.Exception.Message
            $helpText = "`n`nRequired Admin Roles:`n- Security Administrator`n- Global Administrator`n- Organization Management (Exchange role)`n- Compliance Administrator`n`nTo check your roles, run:`nGet-ManagementRoleAssignment -Role 'Transport Hygiene'`n`nOr check via Microsoft 365 admin center under Roles > Security Administrator"
            [System.Windows.Forms.MessageBox]::Show("Error loading blocked sender addresses: $errorMsg$helpText", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            $StatusLabelGlobal.Text = "Error loading blocked sender addresses"
        }
    }
    
    # Refresh button click event
    $refreshButton.add_Click({
        & $loadBlockedSenders
    })
    
    # Remove button click event
    $removeButton.add_Click({
        if ($listView.SelectedItems.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Please select a blocked sender to remove.", "No Selection", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        
        $selectedSender = $listView.SelectedItems[0].Text
        $confirm = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to remove '$selectedSender' from the blocked senders list?`n`nThis will allow this email address to send mail again.", "Confirm Removal", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
        
        if ($confirm -eq [System.Windows.Forms.DialogResult]::Yes) {
            $StatusLabelGlobal.Text = "Removing $selectedSender from blocked senders..."
            
            try {
                # Check if connected to Exchange Online - try multiple methods
                $exchangeSession = Get-PSSession | Where-Object { $_.ConfigurationName -eq "Microsoft.Exchange" -and $_.State -eq "Opened" }
                $isConnected = $false
                
                if ($exchangeSession) {
                    $isConnected = $true
                } else {
                    # Try to test if we can run Exchange Online cmdlets
                    try {
                        $testConnection = Get-Mailbox -ResultSize 1 -ErrorAction Stop
                        $isConnected = $true
                    } catch {
                        $isConnected = $false
                    }
                }
                
                if (-not $isConnected) {
                    [System.Windows.Forms.MessageBox]::Show("Please connect to Exchange Online first.`n`nThese cmdlets require Exchange Online connection and specific admin roles.", "Not Connected", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                    $StatusLabelGlobal.Text = "Not connected to Exchange Online"
                    return
                }
                
                # Check if user has required roles
                try {
                    $roleCheck = Get-ManagementRoleAssignment -Role "Transport Hygiene" -ErrorAction Stop
                    if (-not $roleCheck) {
                        throw "Transport Hygiene role not found"
                    }
                } catch {
                    $roleError = "`n`nRequired Admin Roles:`n- Security Administrator`n- Global Administrator`n- Organization Management (Exchange role)`n- Compliance Administrator`n`nTo check your roles, run:`nGet-ManagementRoleAssignment -Role 'Transport Hygiene'`n`nOr check via Microsoft 365 admin center under Roles > Security Administrator"
                    [System.Windows.Forms.MessageBox]::Show("Permission denied. You need specific admin roles to manage blocked senders.$roleError", "Insufficient Permissions", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                    $StatusLabelGlobal.Text = "Insufficient permissions for blocked sender management"
                    return
                }
                
                # Remove the blocked sender
                Remove-BlockedSenderAddress -SenderAddress $selectedSender -ErrorAction Stop
                
                # Remove from list view
                $listView.SelectedItems[0].Remove()
                
                $StatusLabelGlobal.Text = "Successfully removed $selectedSender from blocked senders"
                [System.Windows.Forms.MessageBox]::Show("Successfully removed '$selectedSender' from blocked senders.", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                
            } catch {
                $errorMsg = $_.Exception.Message
                [System.Windows.Forms.MessageBox]::Show("Error removing blocked sender: $errorMsg`n`nPlease ensure you have the required permissions (SecurityActions.ReadWrite.All)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                $StatusLabelGlobal.Text = "Error removing blocked sender"
            }
        }
    })
    
    # Close button click event
    $closeButton.add_Click({
        $form.Close()
    })
    
    # Defender button click event
    $defenderButton.add_Click({
        Start-Process "https://security.microsoft.com/restrictedentities"
    })
    
    # Load data initially
    & $loadBlockedSenders
    
    # Show the dialog
    $form.ShowDialog($ParentForm)
}

function Show-EmailInputDialog {
    $inputForm = New-Object System.Windows.Forms.Form
    $inputForm.Text = "Enter Microsoft 365 Email"
    $inputForm.Size = New-Object System.Drawing.Size(400, 150)
    $inputForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
    
    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Enter your Microsoft 365 email address:"
    $label.Location = New-Object System.Drawing.Point(10, 20)
    $label.Size = New-Object System.Drawing.Size(350, 20)
    
    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = New-Object System.Drawing.Point(10, 50)
    $textBox.Size = New-Object System.Drawing.Size(350, 20)
    
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "OK"
    $okButton.Location = New-Object System.Drawing.Point(200, 80)
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "Cancel"
    $cancelButton.Location = New-Object System.Drawing.Point(280, 80)
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    
    $inputForm.Controls.AddRange(@($label, $textBox, $okButton, $cancelButton))
    $inputForm.AcceptButton = $okButton
    $inputForm.CancelButton = $cancelButton
    
    $result = $inputForm.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        return $textBox.Text
    } else {
        throw "User cancelled role assignment"
    }
}

function Show-SignInLogsDialog {
    param(
        [Parameter(Mandatory=$true)]
        [array]$Logs,
        [Parameter(Mandatory=$true)]
        [string]$UserName,
        [Parameter(Mandatory=$true)]
        [System.Windows.Forms.Form]$ParentForm
    )
    
    $logForm = New-Object System.Windows.Forms.Form
    $logForm.Text = "Sign-in Logs - $UserName"
    $logForm.Size = New-Object System.Drawing.Size(900, 600)
    $logForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
    $logForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
    $logForm.MaximizeBox = $true
    
    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Dock = 'Fill'
    $grid.ReadOnly = $true
    $grid.AllowUserToAddRows = $false
    $grid.AutoGenerateColumns = $true
    $grid.AutoSizeColumnsMode = 'Fill'
    
    # Convert logs to displayable format
    $displayData = foreach ($log in $logs) {
        [PSCustomObject]@{
            CreatedDateTime = $log.CreatedDateTime
            AppDisplayName = $log.AppDisplayName
            IPAddress = $log.IPAddress
            Location = if ($log.Location) { "$($log.Location.City), $($log.Location.State), $($log.Location.CountryOrRegion)" } else { "Unknown" }
            Status = if ($log.Status) { $log.Status.AdditionalDetails } else { "Unknown" }
            RiskLevel = $log.RiskLevelAggregated
        }
    }
    
    $grid.DataSource = $displayData
    $logForm.Controls.Add($grid)
    [void]$logForm.ShowDialog($ParentForm)
}

function Show-AuditLogsDialog {
    param(
        [Parameter(Mandatory=$true)]
        [array]$Logs,
        [Parameter(Mandatory=$true)]
        [string]$UserName,
        [Parameter(Mandatory=$true)]
        [System.Windows.Forms.Form]$ParentForm
    )
    
    $logForm = New-Object System.Windows.Forms.Form
    $logForm.Text = "Audit Logs - $UserName"
    $logForm.Size = New-Object System.Drawing.Size(900, 600)
    $logForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
    $logForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
    $logForm.MaximizeBox = $true
    
    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Dock = 'Fill'
    $grid.ReadOnly = $true
    $grid.AllowUserToAddRows = $false
    $grid.AutoGenerateColumns = $true
    $grid.AutoSizeColumnsMode = 'Fill'
    
    # Convert logs to displayable format
    $displayData = foreach ($log in $logs) {
        [PSCustomObject]@{
            ActivityDateTime = $log.ActivityDateTime
            Activity = $log.Activity
            InitiatedBy = if ($log.InitiatedBy -and $log.InitiatedBy.User) { $log.InitiatedBy.User.UserPrincipalName } else { "System" }
            TargetResource = if ($log.TargetResources) { $log.TargetResources[0].DisplayName } else { "Unknown" }
            Result = $log.Result
        }
    }
    
    $grid.DataSource = $displayData
    $logForm.Controls.Add($grid)
    [void]$logForm.ShowDialog($ParentForm)
}

# Add Microsoft Graph Security API functions for restricted senders management
function Test-SecurityApiPermissions {
    [CmdletBinding()]
    param()
    try {
        # Test if we can access the Security API
        # Note: The Security API might not be available through standard Graph modules
        # We'll test with a simple Graph call first
        $context = Get-MgContext -ErrorAction Stop
        if (-not $context) {
            return $false
        }
        
        # Try to access the Security API
        try {
            $testResult = Get-MgSecurityThreatIntelligenceHost -Top 1 -ErrorAction Stop
            return $true
        } catch {
            # If Security API is not available, we'll use alternative methods
            return $false
        }
    } catch {
        return $false
    }
}



function Get-RestrictedSendersList {
    [CmdletBinding()]
    param()
    try {
        # Use Exchange Online anti-spam policies (most reliable for restricted senders)
        if (Get-Command -Name "Get-HostedContentFilterPolicy" -ErrorAction SilentlyContinue) {
            try {
                $blockedSenders = @()
                
                # Get the default policy first
                $defaultPolicy = Get-HostedContentFilterPolicy -Identity Default -ErrorAction Stop
                if ($defaultPolicy.BlockedSenders) {
                    foreach ($sender in $defaultPolicy.BlockedSenders) {
                        $blockedSenders += [PSCustomObject]@{
                            Host = $sender
                            FirstSeenDateTime = Get-Date
                            LastSeenDateTime = Get-Date
                            Description = "Blocked via Default Anti-spam Policy"
                            Source = "Exchange Online Anti-spam"
                        }
                    }
                }
                
                # Get all policies and their blocked senders
                $allPolicies = Get-HostedContentFilterPolicy -ErrorAction Stop
                foreach ($policy in $allPolicies) {
                    if ($policy.BlockedSenders -and $policy.Name -ne "Default") {
                        foreach ($sender in $policy.BlockedSenders) {
                            $blockedSenders += [PSCustomObject]@{
                                Host = $sender
                                FirstSeenDateTime = Get-Date
                                LastSeenDateTime = Get-Date
                                Description = "Blocked via $($policy.Name) Anti-spam Policy"
                                Source = "Exchange Online Anti-spam"
                            }
                        }
                    }
                }
                
                return $blockedSenders
            } catch {
                Write-Warning "Exchange Online anti-spam policies not available. Restricted senders list cannot be retrieved."
                return @()
            }
        } else {
            Write-Warning "Get-HostedContentFilterPolicy cmdlet not available. Restricted senders list cannot be retrieved."
            return @()
        }
    } catch {
        Write-Error "Failed to retrieve restricted senders list: $_"
        return @()
    }
}

function Test-UserInRestrictedSenders {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$UserPrincipalName
    )
    try {
        # Get user's email addresses
        $user = Get-MgUser -UserId $UserPrincipalName -Property Mail,ProxyAddresses -ErrorAction Stop
        $emailAddresses = @()
        
        if ($user.Mail) { $emailAddresses += $user.Mail }
        if ($user.ProxyAddresses) {
            $emailAddresses += $user.ProxyAddresses | Where-Object { $_ -like "smtp:*" } | ForEach-Object { $_ -replace "smtp:", "" }
        }
        
        $isRestricted = $false
        $restrictedEmails = @()
        
        # Try Microsoft Defender for Office 365 API first (this is what security.microsoft.com uses)
        try {
            # Check if we have the required Microsoft Graph Security module
            if (-not (Get-Module -Name "Microsoft.Graph.Security" -ListAvailable)) {
                Write-Warning "Microsoft.Graph.Security module not available. Installing..."
                Install-Module -Name "Microsoft.Graph.Security" -Force -AllowClobber
            }
            
            Import-Module Microsoft.Graph.Security -Force
            
            # Get current restricted entities - use correct API endpoint
            $restrictedEntities = Get-MgSecurityThreatIntelligenceHost -ErrorAction Stop
            
            # Check each email address against restricted entities
            foreach ($email in $emailAddresses) {
                $hostEntity = $restrictedEntities | Where-Object { $_.Host -eq $email -and $_.ThreatIntelligence -and $_.ThreatIntelligence.IsBlocked -eq $true }
                if ($hostEntity) {
                    $isRestricted = $true
                    $restrictedEmails += $email
                }
            }
            
            if ($isRestricted) {
                return @{
                    IsRestricted = $isRestricted
                    RestrictedEmails = $restrictedEmails
                    AllUserEmails = $emailAddresses
                    Source = "Microsoft Defender API"
                }
            }
        } catch {
            Write-Warning "Microsoft Defender API not available or failed: $($_.Exception.Message)"
        }
        
        # Fallback to Exchange Online anti-spam policies (legacy method)
        if (Get-Command -Name "Get-HostedContentFilterPolicy" -ErrorAction SilentlyContinue) {
            try {
                # Check default policy first
                $defaultPolicy = Get-HostedContentFilterPolicy -Identity Default -ErrorAction Stop
                if ($defaultPolicy.BlockedSenders) {
                    foreach ($email in $emailAddresses) {
                        if ($defaultPolicy.BlockedSenders -contains $email) {
                            $isRestricted = $true
                            $restrictedEmails += $email
                        }
                    }
                }
                
                # Check all policies
                $allPolicies = Get-HostedContentFilterPolicy -ErrorAction Stop
                foreach ($policy in $allPolicies) {
                    if ($policy.BlockedSenders -and $policy.Name -ne "Default") {
                        foreach ($email in $emailAddresses) {
                            if ($policy.BlockedSenders -contains $email) {
                                $isRestricted = $true
                                $restrictedEmails += $email
                            }
                        }
                    }
                }
            } catch {
                Write-Warning "Exchange Online anti-spam policies not available. Cannot check restricted senders status."
            }
        } else {
            Write-Warning "Get-HostedContentFilterPolicy cmdlet not available. Cannot check restricted senders status."
        }
        
        return @{
            IsRestricted = $isRestricted
            RestrictedEmails = $restrictedEmails
            AllUserEmails = $emailAddresses
            Source = "Exchange Online Anti-spam Policies"
        }
    } catch {
        Write-Error "Failed to check restricted senders status: $_"
        return @{
            IsRestricted = $false
            RestrictedEmails = @()
            AllUserEmails = @()
            Source = "Error"
        }
    }
}

function Add-UserToRestrictedSenders {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$UserPrincipalName,
        [Parameter(Mandatory=$false)]
        [string]$Reason = "Added via management tool"
    )
    try {
        # Get user's primary email
        $user = Get-MgUser -UserId $UserPrincipalName -Property Mail -ErrorAction Stop
        if (-not $user.Mail) {
            throw "User does not have a primary email address"
        }
        
        # Try Microsoft Defender for Office 365 API first (this is what security.microsoft.com uses)
        try {
            # Check if we have the required Microsoft Graph Security module
            if (-not (Get-Module -Name "Microsoft.Graph.Security" -ListAvailable)) {
                Write-Warning "Microsoft.Graph.Security module not available. Installing..."
                Install-Module -Name "Microsoft.Graph.Security" -Force -AllowClobber
            }
            
            Import-Module Microsoft.Graph.Security -Force
            
            # Check if the host is already restricted
            $existingHost = Get-MgSecurityThreatIntelligenceHost -ErrorAction SilentlyContinue | Where-Object { $_.Host -eq $user.Mail }
            
            if ($existingHost) {
                # Update existing host to be blocked
                $updateParams = @{
                    ThreatIntelligence = @{
                        IsBlocked = $true
                        Confidence = "High"
                        Source = "Manual"
                        Description = $Reason
                    }
                }
                
                Update-MgSecurityThreatIntelligenceHost -HostId $existingHost.Id -BodyParameter $updateParams -ErrorAction Stop
                return @{ Host = $user.Mail; Status = "Updated via Microsoft Defender API" }
            } else {
                # Create new blocked host
                $newHost = @{
                    Host = $user.Mail
                    ThreatIntelligence = @{
                        IsBlocked = $true
                        Confidence = "High"
                        Source = "Manual"
                        Description = $Reason
                    }
                }
                
                New-MgSecurityThreatIntelligenceHost -BodyParameter $newHost -ErrorAction Stop
                return @{ Host = $user.Mail; Status = "Added via Microsoft Defender API" }
            }
        } catch {
            Write-Warning "Microsoft Defender API not available or failed: $($_.Exception.Message)"
        }
        
        # Fallback to Exchange Online anti-spam policies (legacy method)
        if (Get-Command -Name "Set-HostedContentFilterPolicy" -ErrorAction SilentlyContinue) {
            Write-Host "Falling back to Exchange Online anti-spam policies..." -ForegroundColor Yellow
            
            # Get current blocked senders from default policy
            $defaultPolicy = Get-HostedContentFilterPolicy -Identity Default -ErrorAction Stop
            $currentBlockedSenders = @()
            if ($defaultPolicy.BlockedSenders) {
                $currentBlockedSenders = $defaultPolicy.BlockedSenders
            }
            
            # Add the new email if not already present
            if ($currentBlockedSenders -notcontains $user.Mail) {
                $currentBlockedSenders += $user.Mail
                Set-HostedContentFilterPolicy -Identity Default -BlockedSenders $currentBlockedSenders -ErrorAction Stop
                return @{ Host = $user.Mail; Status = "Added via Exchange Online Anti-spam Policy" }
            } else {
                return @{ Host = $user.Mail; Status = "Already in blocked senders list" }
            }
        } else {
            throw "Neither Microsoft Defender API nor Exchange Online cmdlets are available. Cannot add user to restricted senders."
        }
    } catch {
        $errorMessage = $_.Exception.Message
        if ($errorMessage -match "401.*Unauthorized") {
            Write-Error "Failed to add user to restricted senders: Permission denied. The 'SecurityEvents.ReadWrite.All' scope is required. Please reconnect to Microsoft Graph with the correct permissions."
        } else {
            Write-Error "Failed to add user to restricted senders: $_"
        }
        throw
    }
}

function Remove-UserFromRestrictedSenders {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$UserPrincipalName
    )
    try {
        # Get user's email addresses
        $user = Get-MgUser -UserId $UserPrincipalName -Property Mail,ProxyAddresses -ErrorAction Stop
        $emailAddresses = @()
        
        if ($user.Mail) { 
            $cleanMail = $user.Mail.Trim()
            if ($cleanMail -match '^[^@]+@[^@]+\.[^@]+$') {
                $emailAddresses += $cleanMail
            }
        }
        if ($user.ProxyAddresses) {
            $emailAddresses += $user.ProxyAddresses | Where-Object { $_ -like "smtp:*" } | ForEach-Object { 
                $cleanEmail = ($_ -replace "smtp:", "").Trim()
                if ($cleanEmail -match '^[^@]+@[^@]+\.[^@]+$') {
                    $cleanEmail
                }
            }
        }
        
        # Ensure all entries are strings
        $emailAddresses = $emailAddresses | ForEach-Object { [string]$_ }
        $removedCount = 0
        
        # Try Microsoft Defender for Office 365 API first (this is what security.microsoft.com uses)
        try {
            # Check if we have the required Microsoft Graph Security module
            if (-not (Get-Module -Name "Microsoft.Graph.Security" -ListAvailable)) {
                Write-Warning "Microsoft.Graph.Security module not available. Installing..."
                Install-Module -Name "Microsoft.Graph.Security" -Force -AllowClobber
            }
            
            Import-Module Microsoft.Graph.Security -Force
            
            # Get current restricted entities - use correct API endpoint
            $restrictedEntities = Get-MgSecurityThreatIntelligenceHost -ErrorAction Stop
            
            # Remove each email address from restricted entities
            foreach ($email in $emailAddresses) {
                try {
                    # Find the host entity for this email that is blocked
                    $hostEntity = $restrictedEntities | Where-Object { $_.Host -eq $email -and $_.ThreatIntelligence -and $_.ThreatIntelligence.IsBlocked -eq $true }
                    if ($hostEntity) {
                        # Update the host to unblock it
                        $updateParams = @{
                            ThreatIntelligence = @{
                                IsBlocked = $false
                            }
                        }
                        
                        Update-MgSecurityThreatIntelligenceHost -HostId $hostEntity.Id -BodyParameter $updateParams -ErrorAction Stop
                        $removedCount++
                        Write-Host "Removed $email from restricted entities via Microsoft Defender API" -ForegroundColor Green
                    } else {
                        Write-Host "Email $email not found in restricted entities or not blocked" -ForegroundColor Yellow
                    }
                } catch {
                    Write-Warning "Failed to remove $email via Microsoft Defender API: $($_.Exception.Message)"
                }
            }
            
            if ($removedCount -gt 0) {
                return $removedCount
            }
        } catch {
            Write-Warning "Microsoft Defender API not available or failed: $($_.Exception.Message)"
        }
        
        # Try REST API approach as backup for Global Administrator
        try {
            $token = Get-MgContext | Select-Object -ExpandProperty TokenCache | Select-Object -ExpandProperty AccessToken
            if ($token) {
                Write-Host "Trying REST API approach with Global Administrator privileges..." -ForegroundColor Yellow
                
                foreach ($email in $emailAddresses) {
                    try {
                        $apiUrl = "https://api.security.microsoft.com/api/threatintelligence/hosts/$email"
                        $headers = @{
                            'Authorization' = "Bearer $token"
                            'Content-Type' = 'application/json'
                        }
                        
                        # First check if the entity exists and is blocked
                        $response = Invoke-RestMethod -Uri $apiUrl -Headers $headers -Method Get -ErrorAction SilentlyContinue
                        if ($response -and $response.isBlocked -eq $true) {
                            $body = @{ isBlocked = $false } | ConvertTo-Json
                            Invoke-RestMethod -Uri $apiUrl -Headers $headers -Method Patch -Body $body -ErrorAction Stop
                            $removedCount++
                            Write-Host "Removed $email from restricted entities via REST API" -ForegroundColor Green
                        } else {
                            Write-Host "Email $email not found in restricted entities or not blocked via REST API" -ForegroundColor Yellow
                        }
                    } catch {
                        Write-Warning "Failed to remove $email via REST API: $($_.Exception.Message)"
                    }
                }
                
                if ($removedCount -gt 0) {
                    return $removedCount
                }
            }
        } catch {
            Write-Warning "REST API approach failed: $($_.Exception.Message)"
        }
        
        # Fallback to Exchange Online anti-spam policies (legacy method)
        if (Get-Command -Name "Set-HostedContentFilterPolicy" -ErrorAction SilentlyContinue) {
            Write-Host "Falling back to Exchange Online anti-spam policies..." -ForegroundColor Yellow
            
            # Get current blocked senders from default policy
            $defaultPolicy = Get-HostedContentFilterPolicy -Identity Default -ErrorAction Stop
            $currentBlockedSenders = @()
            if ($defaultPolicy.BlockedSenders) {
                $currentBlockedSenders = $defaultPolicy.BlockedSenders | ForEach-Object { [string]$_ }
            }
            
            # Remove each email address
            foreach ($email in $emailAddresses) {
                if ($currentBlockedSenders -contains $email) {
                    $currentBlockedSenders = $currentBlockedSenders | Where-Object { $_ -ne $email } | ForEach-Object { [string]$_ }
                    $removedCount++
                }
            }
            
            # Update the policy - ensure proper array format
            if ($currentBlockedSenders.Count -eq 0) {
                Set-HostedContentFilterPolicy -Identity Default -BlockedSenders @() -ErrorAction Stop
            } else {
                Set-HostedContentFilterPolicy -Identity Default -BlockedSenders $currentBlockedSenders -ErrorAction Stop
            }
            
            Write-Host "Removed $removedCount email addresses via Exchange Online anti-spam policies" -ForegroundColor Green
        } else {
            throw "Neither Microsoft Defender API nor Exchange Online cmdlets are available. Cannot remove user from restricted senders."
        }
        
        return $removedCount
    } catch {
        $errorMessage = $_.Exception.Message
        if ($errorMessage -match "Invalid sender address") {
            Write-Error "Failed to remove user from restricted senders: Invalid email address format. Please check the user's email addresses."
        } else {
            Write-Error "Failed to remove user from restricted senders: $_"
        }
        throw
    }
}

function Remove-UserFromRestrictedSendersViaRestApi {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$UserPrincipalName
    )
    try {
        # Get user's email addresses
        $user = Get-MgUser -UserId $UserPrincipalName -Property Mail,ProxyAddresses -ErrorAction Stop
        $emailAddresses = @()
        
        if ($user.Mail) { 
            $cleanMail = $user.Mail.Trim()
            if ($cleanMail -match '^[^@]+@[^@]+\.[^@]+$') {
                $emailAddresses += $cleanMail
            }
        }
        if ($user.ProxyAddresses) {
            $emailAddresses += $user.ProxyAddresses | Where-Object { $_ -like "smtp:*" } | ForEach-Object { 
                $cleanEmail = ($_ -replace "smtp:", "").Trim()
                if ($cleanEmail -match '^[^@]+@[^@]+\.[^@]+$') {
                    $cleanEmail
                }
            }
        }
        
        $removedCount = 0
        
        # Use Microsoft Defender for Office 365 REST API
        try {
            $context = Get-MgContext -ErrorAction Stop
            if (-not $context) {
                throw "Not connected to Microsoft Graph"
            }
            
            # Get access token for Microsoft Defender API
            $token = Get-MgContext -ErrorAction Stop | Select-Object -ExpandProperty TokenCache | Select-Object -ExpandProperty AccessToken
            
            foreach ($email in $emailAddresses) {
                try {
                    # Microsoft Defender for Office 365 API endpoint for restricted entities
                    $apiUrl = "https://api.security.microsoft.com/api/threatintelligence/hosts/$email"
                    
                    $headers = @{
                        'Authorization' = "Bearer $token"
                        'Content-Type' = 'application/json'
                    }
                    
                    # Get current status
                    $response = Invoke-RestMethod -Uri $apiUrl -Headers $headers -Method Get -ErrorAction SilentlyContinue
                    
                    if ($response -and $response.isBlocked -eq $true) {
                        # Remove the restriction by updating to unblocked
                        $body = @{
                            isBlocked = $false
                        } | ConvertTo-Json
                        
                        $updateResponse = Invoke-RestMethod -Uri $apiUrl -Headers $headers -Method Patch -Body $body -ErrorAction Stop
                        $removedCount++
                        Write-Host "Removed $email from restricted entities via Microsoft Defender REST API" -ForegroundColor Green
                    } else {
                        Write-Host "Email $email not found in restricted entities or not blocked" -ForegroundColor Yellow
                    }
                } catch {
                    Write-Warning "Failed to remove $email via Microsoft Defender REST API: $($_.Exception.Message)"
                }
            }
            
            if ($removedCount -gt 0) {
                return $removedCount
            }
        } catch {
            Write-Warning "Microsoft Defender REST API not available or failed: $($_.Exception.Message)"
        }
        
        # Method 2: Try Remove-BlockedSenderAddress (simpler Exchange Online approach)
        try {
            if (Get-Command -Name "Remove-BlockedSenderAddress" -ErrorAction SilentlyContinue) {
                Write-Host "Trying Remove-BlockedSenderAddress method..." -ForegroundColor Yellow
                
                foreach ($email in $emailAddresses) {
                    try {
                        Remove-BlockedSenderAddress -SenderAddress $email -Confirm:$false -ErrorAction Stop
                        $removedCount++
                        Write-Host "Removed $email using Remove-BlockedSenderAddress" -ForegroundColor Green
                    } catch {
                        Write-Warning "Could not remove $email via Remove-BlockedSenderAddress: $($_.Exception.Message)"
                    }
                }
                
                if ($removedCount -gt 0) {
                    return $removedCount
                }
            }
        } catch {
            Write-Warning "Remove-BlockedSenderAddress not available: $($_.Exception.Message)"
        }
        
        # Method 3: Fallback to Exchange Online anti-spam policies
        if (Get-Command -Name "Set-HostedContentFilterPolicy" -ErrorAction SilentlyContinue) {
            Write-Host "Falling back to Exchange Online anti-spam policies..." -ForegroundColor Yellow
            
            # Get current blocked senders from default policy
            $defaultPolicy = Get-HostedContentFilterPolicy -Identity Default -ErrorAction Stop
            $currentBlockedSenders = @()
            if ($defaultPolicy.BlockedSenders) {
                $currentBlockedSenders = $defaultPolicy.BlockedSenders | ForEach-Object { [string]$_ }
            }
            
            # Remove each email address
            foreach ($email in $emailAddresses) {
                if ($currentBlockedSenders -contains $email) {
                    $currentBlockedSenders = $currentBlockedSenders | Where-Object { $_ -ne $email } | ForEach-Object { [string]$_ }
                    $removedCount++
                }
            }
            
            # Update the policy - ensure proper array format
            if ($currentBlockedSenders.Count -eq 0) {
                Set-HostedContentFilterPolicy -Identity Default -BlockedSenders @() -ErrorAction Stop
            } else {
                Set-HostedContentFilterPolicy -Identity Default -BlockedSenders $currentBlockedSenders -ErrorAction Stop
            }
            
            Write-Host "Removed $removedCount email addresses via Exchange Online anti-spam policies" -ForegroundColor Green
        } else {
            throw "Neither Microsoft Defender REST API, Remove-BlockedSenderAddress, nor Exchange Online cmdlets are available. Cannot remove user from restricted senders."
        }
        
        return $removedCount
    } catch {
        $errorMessage = $_.Exception.Message
        Write-Error "Failed to remove user from restricted senders: $_"
        throw
    }
}

function Remove-RestrictedSender {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$SenderAddress,
        
        [Parameter(Mandatory=$false)]
        [string]$PolicyName = "Default"
    )
    
    try {
        $removed = $false
        
        # Method 1: Remove from Microsoft Defender restricted entities
        try {
            $context = Get-MgContext -ErrorAction SilentlyContinue
            if ($context) {
                Write-Host "Checking Microsoft Defender restricted entities..." -ForegroundColor Yellow
                
                # Try to remove from security alerts
                $securityAlerts = Get-MgSecurityAlert -Filter "status eq 'newAlert' or status eq 'inProgress'" -ErrorAction SilentlyContinue
                foreach ($alert in $securityAlerts) {
                    if ($alert.UserStates -and $alert.UserStates.UserPrincipalName -eq $SenderAddress) {
                        # Try to resolve the alert
                        try {
                            $updateBody = @{
                                Status = "resolved"
                                AssignedTo = "automated"
                            }
                            Update-MgSecurityAlert -AlertId $alert.Id -BodyParameter $updateBody -ErrorAction Stop
                            $removed = $true
                            Write-Host "Resolved security alert for $SenderAddress" -ForegroundColor Green
                        } catch {
                            Write-Warning "Could not resolve security alert: $($_.Exception.Message)"
                        }
                    }
                }
                
                # Try to remove from security incidents
                $securityIncidents = Get-MgSecurityIncident -Filter "status eq 'active'" -ErrorAction SilentlyContinue
                foreach ($incident in $securityIncidents) {
                    if ($incident.UserStates -and $incident.UserStates.UserPrincipalName -eq $SenderAddress) {
                        # Try to resolve the incident
                        try {
                            $updateBody = @{
                                Status = "resolved"
                                AssignedTo = "automated"
                            }
                            Update-MgSecurityIncident -SecurityIncidentId $incident.Id -BodyParameter $updateBody -ErrorAction Stop
                            $removed = $true
                            Write-Host "Resolved security incident for $SenderAddress" -ForegroundColor Green
                        } catch {
                            Write-Warning "Could not resolve security incident: $($_.Exception.Message)"
                        }
                    }
                }
            }
        } catch {
            Write-Warning "Could not access Microsoft Defender data: $($_.Exception.Message)"
        }
        
        # Method 2: Remove from Anti-Spam policy
        $policy = Get-HostedContentFilterPolicy -Identity $PolicyName -ErrorAction SilentlyContinue
        if ($policy -and $policy.BlockedSenders -contains $SenderAddress) {
            $newBlockedList = $policy.BlockedSenders | Where-Object { $_ -ne $SenderAddress }
            Set-HostedContentFilterPolicy -Identity $PolicyName -BlockedSenders $newBlockedList
            $removed = $true
            Write-Host "Removed $SenderAddress from Anti-Spam policy blocked senders"
        }
        
        # Method 2: Remove from Transport Rules
        $transportRules = Get-TransportRule | Where-Object { 
            ($_.SenderDomainIs -contains $SenderAddress) -or 
            ($_.From -contains $SenderAddress) 
        }
        
        foreach ($rule in $transportRules) {
            if ($rule.SenderDomainIs -contains $SenderAddress) {
                $newSenderDomains = $rule.SenderDomainIs | Where-Object { $_ -ne $SenderAddress }
                Set-TransportRule -Identity $rule.Identity -SenderDomainIs $newSenderDomains
                $removed = $true
                Write-Host "Removed $SenderAddress from Transport Rule: $($rule.Name)"
            }
            
            if ($rule.From -contains $SenderAddress) {
                $newFromList = $rule.From | Where-Object { $_ -ne $SenderAddress }
                Set-TransportRule -Identity $rule.Identity -From $newFromList
                $removed = $true
                Write-Host "Removed $SenderAddress from Transport Rule: $($rule.Name)"
            }
        }
        
        # Method 3: Remove from quarantine
        $quarantineMessages = Get-QuarantineMessage -SenderAddress $SenderAddress -ErrorAction SilentlyContinue
        foreach ($message in $quarantineMessages) {
            Release-QuarantineMessage -Identity $message.Identity -ReleaseToAll
            $removed = $true
            Write-Host "Released quarantined message from $SenderAddress"
        }
        
        return $removed
    } catch {
        Write-Error "Failed to remove restricted sender $SenderAddress : $($_.Exception.Message)"
        return $false
    }
}

function Add-RestrictedSender {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$SenderAddress,
        
        [Parameter(Mandatory=$false)]
        [string]$PolicyName = "Default"
    )
    
    try {
        # Add to Anti-Spam policy
        $policy = Get-HostedContentFilterPolicy -Identity $PolicyName -ErrorAction SilentlyContinue
        if ($policy) {
            $currentBlocked = $policy.BlockedSenders
            if ($currentBlocked -notcontains $SenderAddress) {
                $newBlockedList = $currentBlocked + $SenderAddress
                Set-HostedContentFilterPolicy -Identity $PolicyName -BlockedSenders $newBlockedList
                Write-Host "Added $SenderAddress to Anti-Spam policy blocked senders"
                return $true
            } else {
                Write-Host "$SenderAddress is already in the blocked senders list"
                return $true
            }
        }
        return $false
    } catch {
        Write-Error "Failed to add restricted sender $SenderAddress : $($_.Exception.Message)"
        return $false
    }
}

function Get-RestrictedSenders {
    [CmdletBinding()]
    param()
    
    try {
        Write-Host "Checking Microsoft Defender restricted entities..." -ForegroundColor Cyan
        
        # Initialize results
        $results = @{
            DefenderRestrictedEntities = @()
            BlockedSenders = @()
            QuarantinedSenders = @()
            TransportRuleBlocked = @()
        }
        
        # Method 1: Get Microsoft Defender Restricted Entities via Graph API
        try {
            # Check if connected to Microsoft Graph
            $context = Get-MgContext -ErrorAction SilentlyContinue
            if ($context) {
                Write-Host "Getting restricted entities from Microsoft Graph Security API..." -ForegroundColor Yellow
                
                # Method 1a: Try security incidents endpoint
                try {
                    $securityIncidents = Get-MgSecurityIncident -Filter "status eq 'active'" -ErrorAction SilentlyContinue
                    foreach ($incident in $securityIncidents) {
                        if ($incident.Title -like "*restricted*" -or $incident.Title -like "*sending*") {
                            Write-Host "Found security incident: $($incident.Title)" -ForegroundColor Red
                            # Extract user information from incident
                            if ($incident.UserStates) {
                                foreach ($userState in $incident.UserStates) {
                                    $results.DefenderRestrictedEntities += [PSCustomObject]@{
                                        UserPrincipalName = $userState.UserPrincipalName
                                        Type = "Security Incident"
                                        Reason = $incident.Title
                                        Category = "Security"
                                        Severity = $incident.Severity
                                        CreatedDateTime = $incident.CreatedDateTime
                                        Status = $incident.Status
                                        IncidentId = $incident.Id
                                    }
                                }
                            }
                        }
                    }
                } catch {
                    Write-Warning "Could not access security incidents: $($_.Exception.Message)"
                }
                
                # Method 1b: Try security alerts endpoint  
                try {
                    $securityAlerts = Get-MgSecurityAlert -Filter "status eq 'newAlert' or status eq 'inProgress'" -ErrorAction SilentlyContinue
                    foreach ($alert in $securityAlerts) {
                        if ($alert.Title -like "*restricted*" -or $alert.Title -like "*sending*" -or $alert.Category -eq "Email") {
                            Write-Host "Found security alert: $($alert.Title)" -ForegroundColor Red
                            # Extract user information from alert
                            if ($alert.UserStates) {
                                foreach ($userState in $alert.UserStates) {
                                    $results.DefenderRestrictedEntities += [PSCustomObject]@{
                                        UserPrincipalName = $userState.UserPrincipalName
                                        Type = "Security Alert"
                                        Reason = $alert.Title
                                        Category = $alert.Category
                                        Severity = $alert.Severity
                                        CreatedDateTime = $alert.CreatedDateTime
                                        Status = $alert.Status
                                        AlertId = $alert.Id
                                    }
                                }
                            }
                        }
                    }
                } catch {
                    Write-Warning "Could not access security alerts: $($_.Exception.Message)"
                }
                
                # Method 1c: Direct REST API call to security.microsoft.com endpoints
                try {
                    Write-Host "Attempting direct REST API call for restricted entities..." -ForegroundColor Yellow
                    
                    # Get access token for Graph API - multiple methods
                    $accessToken = $null
                    
                    # Method 1: Try Get-MgContext
                    try {
                        $context = Get-MgContext -ErrorAction SilentlyContinue
                        if ($context -and $context.TokenCache) {
                            $accessToken = $context.TokenCache.AccessToken
                        }
                    } catch {
                        Write-Warning "Could not get token via Get-MgContext"
                    }
                    
                    # Method 2: Try GraphSession if Method 1 failed
                    if (-not $accessToken) {
                        try {
                            $accessToken = [Microsoft.Graph.PowerShell.Authentication.GraphSession]::Instance.AuthContext.TokenCache.ReadItems() | 
                                           Where-Object { $_.Resource -eq "https://graph.microsoft.com" } | 
                                           Select-Object -First 1 -ExpandProperty AccessToken
                        } catch {
                            Write-Warning "Could not get token via GraphSession"
                        }
                    }
                    
                    if ($accessToken) {
                        $headers = @{
                            'Authorization' = "Bearer $accessToken"
                            'Content-Type' = 'application/json'
                        }
                        
                        # Try Microsoft 365 Defender API endpoint
                        $defenderUri = "https://graph.microsoft.com/v1.0/security/alerts?`$filter=category eq 'Email'"
                        $response = Invoke-RestMethod -Uri $defenderUri -Headers $headers -Method Get -ErrorAction SilentlyContinue
                        
                        foreach ($alert in $response.value) {
                            if ($alert.title -like "*restricted*" -or $alert.title -like "*sending*") {
                                Write-Host "Found restricted entity via REST API: $($alert.title)" -ForegroundColor Red
                                
                                # Extract user information
                                if ($alert.userStates) {
                                    foreach ($userState in $alert.userStates) {
                                        $results.DefenderRestrictedEntities += [PSCustomObject]@{
                                            UserPrincipalName = $userState.userPrincipalName
                                            Type = "Defender Restricted Entity"
                                            Reason = $alert.title
                                            Category = $alert.category
                                            Severity = $alert.severity
                                            CreatedDateTime = $alert.createdDateTime
                                            Status = $alert.status
                                            AlertId = $alert.id
                                        }
                                    }
                                }
                            }
                        }
                    }
                } catch {
                    Write-Warning "Could not access Defender API via REST: $($_.Exception.Message)"
                }
            } else {
                Write-Warning "Not connected to Microsoft Graph. Cannot check Defender restricted entities."
            }
        } catch {
            Write-Warning "Error accessing Microsoft Defender data: $($_.Exception.Message)"
        }
        
        # Method 2: Alternative - Check via Exchange Online for restricted sending
        try {
            Write-Host "Checking Exchange Online for restricted senders..." -ForegroundColor Yellow
            
            # Get users who cannot send mail (this might catch Defender restrictions)
            $mailboxes = Get-Mailbox -ResultSize 1000 -ErrorAction SilentlyContinue
            foreach ($mailbox in $mailboxes) {
                try {
                    # Check if user has send restrictions
                    $user = Get-User -Identity $mailbox.UserPrincipalName -ErrorAction SilentlyContinue
                    if ($user -and $user.BlockCredential -eq $true) {
                        $results.DefenderRestrictedEntities += [PSCustomObject]@{
                            UserPrincipalName = $mailbox.UserPrincipalName
                            Type = "Exchange Blocked Credential"
                            Reason = "User credentials blocked"
                            Category = "Authentication"
                            Severity = "High"
                            CreatedDateTime = $null
                            Status = "Blocked"
                        }
                    }
                } catch {
                    # Continue processing other mailboxes
                }
            }
        } catch {
            Write-Warning "Error checking Exchange Online restrictions: $($_.Exception.Message)"
        }
        
        # Method 3: Get blocked senders from Anti-Spam policy (existing code)
        try {
            $antiSpamPolicies = Get-HostedContentFilterPolicy -ErrorAction SilentlyContinue
            foreach ($policy in $antiSpamPolicies) {
                if ($policy.BlockedSenders) {
                    $results.BlockedSenders += $policy.BlockedSenders
                }
            }
        } catch {
            Write-Warning "Error getting anti-spam policies: $($_.Exception.Message)"
        }
        
        # Method 4: Get quarantined senders (existing code)
        try {
            $quarantineMessages = Get-QuarantineMessage -StartReceivedDate (Get-Date).AddDays(-30) -ErrorAction SilentlyContinue
            $results.QuarantinedSenders = $quarantineMessages | Group-Object SenderAddress | Select-Object Name, Count
        } catch {
            Write-Warning "Error getting quarantine messages: $($_.Exception.Message)"
        }
        
        Write-Host "Found $($results.DefenderRestrictedEntities.Count) Defender restricted entities" -ForegroundColor Green
        Write-Host "Found $($results.BlockedSenders.Count) blocked senders" -ForegroundColor Green
        Write-Host "Found $($results.QuarantinedSenders.Count) quarantined senders" -ForegroundColor Green
        
        return $results
    } catch {
        Write-Error "Failed to get restricted senders: $($_.Exception.Message)"
        return $null
    }
}



Export-ModuleMember -Function Show-RestrictedSenderManagementDialog, Get-RestrictedSendersList, Test-UserInRestrictedSenders, Add-UserToRestrictedSenders, Remove-UserFromRestrictedSenders, Remove-UserFromRestrictedSendersViaRestApi, Test-SecurityApiPermissions, Remove-RestrictedSender, Add-RestrictedSender, Get-RestrictedSenders 