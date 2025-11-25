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
    
    # Create test cmdlet button (for diagnostics)
    $testCmdletButton = New-Object System.Windows.Forms.Button
    $testCmdletButton.Text = "Test Cmdlet"
    $testCmdletButton.Location = New-Object System.Drawing.Point(10, 10)
    $testCmdletButton.Size = New-Object System.Drawing.Size(100, 30)
    $testCmdletButton.BackColor = [System.Drawing.Color]::LightYellow
    $testCmdletButtonTooltip = New-Object System.Windows.Forms.ToolTip
    $testCmdletButtonTooltip.SetToolTip($testCmdletButton, "Test if Get-BlockedSenderAddress works directly in PowerShell")
    $testCmdletButton.add_Click({
        try {
            $StatusLabelGlobal.Text = "Testing Get-BlockedSenderAddress cmdlet..."
            $form.Refresh()
            
            # Try to run the cmdlet directly
            $testResult = Get-BlockedSenderAddress -ErrorAction Stop
            $count = if ($testResult) { $testResult.Count } else { 0 }
            
            [System.Windows.Forms.MessageBox]::Show("SUCCESS!`n`nThe cmdlet works and returned $count restricted entity(ies).`n`nThe issue may be with how the app is calling it. Try clicking 'Refresh List' again.", "Cmdlet Test - Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            $StatusLabelGlobal.Text = "Cmdlet test successful - $count entities found"
        } catch {
            $errorMsg = $_.Exception.Message
            $errorMsg = $errorMsg -replace '\|\|', ''
            
            $diagnostic = "The cmdlet FAILED with this error:`n`n$errorMsg`n`n"
            $diagnostic += "This confirms the cmdlet is not accessible even with direct PowerShell execution.`n`n"
            $diagnostic += "Possible reasons:`n"
            $diagnostic += "1. The cmdlet may not be available in your Exchange Online environment`n"
            $diagnostic += "2. Your tenant may require different permissions or licensing`n"
            $diagnostic += "3. The cmdlet may require direct role assignment (not role group membership)`n"
            $diagnostic += "4. There may be a tenant-specific limitation`n`n"
            $diagnostic += "RECOMMENDATION: Use the Microsoft 365 Defender portal instead.`n"
            $diagnostic += "Click 'Open Defender Portal' button to access restricted entities there."
            
            [System.Windows.Forms.MessageBox]::Show($diagnostic, "Cmdlet Test - Failed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            $StatusLabelGlobal.Text = "Cmdlet test failed - use Defender portal instead"
        }
    })
    
    # Create refresh button
    $refreshButton = New-Object System.Windows.Forms.Button
    $refreshButton.Text = "Refresh List"
    $refreshButton.Location = New-Object System.Drawing.Point(120, 10)
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
    $defenderButton.Text = "Open Defender Portal"
    $defenderButton.Location = New-Object System.Drawing.Point(360, 10)
    $defenderButton.Size = New-Object System.Drawing.Size(150, 30)
    $defenderButton.BackColor = [System.Drawing.Color]::LightBlue
    
    # Create button to manage Exchange role groups
    $manageRolesButton = New-Object System.Windows.Forms.Button
    $manageRolesButton.Text = "Manage Exchange Roles"
    $manageRolesButton.Location = New-Object System.Drawing.Point(520, 10)
    $manageRolesButton.Size = New-Object System.Drawing.Size(160, 30)
    $manageRolesButton.BackColor = [System.Drawing.Color]::LightGreen
    $manageRolesButtonTooltip = New-Object System.Windows.Forms.ToolTip
    $manageRolesButtonTooltip.SetToolTip($manageRolesButton, "Add yourself or others to Exchange Online role groups (e.g., Organization Management)")
    
    # Add controls to form
    $form.Controls.AddRange(@($listView, $testCmdletButton, $refreshButton, $removeButton, $closeButton, $defenderButton, $manageRolesButton))
    
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
            
            # Check Exchange Online role assignments first
            $currentUser = $env:USERNAME
            $userPrincipalName = $null
            try {
                $currentSession = Get-PSSession | Where-Object { $_.ConfigurationName -eq "Microsoft.Exchange" -and $_.State -eq "Opened" } | Select-Object -First 1
                if ($currentSession) {
                    $userPrincipalName = $currentSession.Name -replace '.*@', ''
                    if ($userPrincipalName) {
                        $userPrincipalName = "$currentUser@$userPrincipalName"
                    }
                }
            } catch {}
            
            # First, let's check what roles the user actually has
            Write-Host "Checking Exchange Online role group memberships..." -ForegroundColor Cyan
            $actualRoleGroups = @()
            try {
                $allRoleGroups = Get-RoleGroup -ErrorAction SilentlyContinue
                $currentUserEmail = $null
                
                # Try to get current user email
                try {
                    $session = Get-PSSession | Where-Object { $_.ConfigurationName -eq "Microsoft.Exchange" -and $_.State -eq "Opened" } | Select-Object -First 1
                    if ($session -and $session.Name -match '@') {
                        $currentUserEmail = $session.Name
                    } else {
                        # Try Get-Mailbox
                        $mbx = Get-Mailbox -Identity $env:USERNAME -ErrorAction SilentlyContinue
                        if ($mbx) { $currentUserEmail = $mbx.PrimarySmtpAddress }
                    }
                } catch {}
                
                if ($currentUserEmail) {
                    Write-Host "Checking role groups for: $currentUserEmail" -ForegroundColor Yellow
                    foreach ($group in $allRoleGroups) {
                        try {
                            $members = Get-RoleGroupMember -Identity $group.Name -ErrorAction SilentlyContinue
                            foreach ($member in $members) {
                                if ($member.WindowsLiveID -eq $currentUserEmail -or $member.PrimarySmtpAddress -eq $currentUserEmail -or $member.Name -eq $currentUserEmail) {
                                    $actualRoleGroups += $group.Name
                                    Write-Host "  Found in: $($group.Name)" -ForegroundColor Green
                                    break
                                }
                            }
                        } catch {
                            Write-Warning "Could not check $($group.Name): $($_.Exception.Message)"
                        }
                    }
                }
            } catch {
                Write-Warning "Could not check role groups: $($_.Exception.Message)"
            }
            
            # First, check if the cmdlet exists
            $cmdletExists = $false
            try {
                $cmdletCheck = Get-Command Get-BlockedSenderAddress -ErrorAction Stop
                $cmdletExists = $true
                Write-Host "Get-BlockedSenderAddress cmdlet found" -ForegroundColor Green
                
                # Also check what roles are actually assigned to the current user for this cmdlet
                Write-Host "Checking role assignments for Get-BlockedSenderAddress..." -ForegroundColor Cyan
                $transportHygieneFound = $false
                try {
                    $currentUserForCheck = $userPrincipalName
                    if (-not $currentUserForCheck) {
                        # Try to get current user
                        $session = Get-PSSession | Where-Object { $_.ConfigurationName -eq "Microsoft.Exchange" -and $_.State -eq "Opened" } | Select-Object -First 1
                        if ($session -and $session.Name -match '@') {
                            $currentUserForCheck = $session.Name
                        }
                    }
                    
                    if ($currentUserForCheck) {
                        Write-Host "Checking Transport Hygiene role for: $currentUserForCheck" -ForegroundColor Yellow
                        # Check if user has Transport Hygiene role (required for Get-BlockedSenderAddress)
                        $transportHygieneRoles = Get-ManagementRoleAssignment -Role "Transport Hygiene" -ErrorAction SilentlyContinue
                        if ($transportHygieneRoles) {
                            # Check if user is in a role group that has Transport Hygiene
                            $userRoleGroups = @()
                            try {
                                $allRoleGroups = Get-RoleGroup -ErrorAction SilentlyContinue
                                foreach ($rg in $allRoleGroups) {
                                    $members = Get-RoleGroupMember -Identity $rg.Name -ErrorAction SilentlyContinue
                                    if ($members | Where-Object { 
                                        $_.WindowsLiveID -eq $currentUserForCheck -or 
                                        $_.PrimarySmtpAddress -eq $currentUserForCheck -or 
                                        $_.Name -eq $currentUserForCheck 
                                    }) {
                                        $userRoleGroups += $rg.Name
                                    }
                                }
                            } catch {}
                            
                            # Check if any of the user's role groups have Transport Hygiene
                            foreach ($roleAssignment in $transportHygieneRoles) {
                                if ($roleAssignment.RoleAssigneeType -eq 'RoleGroup' -and 
                                    $userRoleGroups -contains $roleAssignment.RoleAssigneeName) {
                                    $transportHygieneFound = $true
                                    Write-Host "✓ Transport Hygiene role found via role group: $($roleAssignment.RoleAssigneeName)" -ForegroundColor Green
                                    break
                                }
                            }
                            
                            if (-not $transportHygieneFound) {
                                Write-Host "⚠ Transport Hygiene role exists but may not be accessible via your role groups" -ForegroundColor Yellow
                                Write-Host "  Your role groups: $($userRoleGroups -join ', ')" -ForegroundColor Yellow
                                Write-Host "  Role groups with Transport Hygiene: $($transportHygieneRoles | Where-Object { $_.RoleAssigneeType -eq 'RoleGroup' } | Select-Object -ExpandProperty RoleAssigneeName -Unique | Sort-Object | Join-String -Separator ', ')" -ForegroundColor Yellow
                            }
                        } else {
                            Write-Host "WARNING: No Transport Hygiene role assignments found in tenant" -ForegroundColor Red
                        }
                    }
                } catch {
                    Write-Warning "Could not check role assignments: $($_.Exception.Message)"
                }
            } catch {
                Write-Host "Get-BlockedSenderAddress cmdlet NOT found - may not be available in your Exchange Online environment" -ForegroundColor Yellow
                [System.Windows.Forms.MessageBox]::Show("The Get-BlockedSenderAddress cmdlet is not available in your Exchange Online environment.`n`nThis cmdlet may:`n- Not be available in your tenant`n- Require a newer version of ExchangeOnlineManagement module`n- Not be available in your region`n`nAlternative: Use the Microsoft 365 Defender portal to view restricted entities:`nhttps://security.microsoft.com/restrictedentities", "Cmdlet Not Available", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                $StatusLabelGlobal.Text = "Get-BlockedSenderAddress cmdlet not available"
                return
            }
            
            # Try to get blocked sender addresses (this will fail with a clear error if permissions are missing)
            # We'll catch permission errors specifically and provide helpful guidance
            $blockedSenders = $null
            try {
                Write-Host "Attempting to call Get-BlockedSenderAddress..." -ForegroundColor Cyan
                $blockedSenders = Get-BlockedSenderAddress -ErrorAction Stop
                Write-Host "Successfully retrieved blocked sender addresses: $($blockedSenders.Count) found" -ForegroundColor Green
            } catch {
                # Get the full error details and clean up formatting
                $errorMessage = $_.Exception.Message
                $errorMessage = $errorMessage -replace '\|\|', ''  # Remove double pipes
                $errorMessage = $errorMessage.Trim()
                $errorCategory = $_.CategoryInfo.Category
                $errorDetails = $_.Exception.ToString()
                $errorDetails = $errorDetails -replace '\|\|', ''  # Clean up error details too
                
                Write-Host "Error calling Get-BlockedSenderAddress:" -ForegroundColor Red
                Write-Host "  Message: $errorMessage" -ForegroundColor Red
                Write-Host "  Category: $errorCategory" -ForegroundColor Red
                Write-Host "  Full Error: $errorDetails" -ForegroundColor Red
                
                if ($errorMessage -match "role definition" -or $errorMessage -match "aren't present" -or $errorMessage -match "permission" -or $errorMessage -match "access denied") {
                    # Try to check Exchange role assignments
                    $exchangeRoles = @()
                    $roleGroupMembership = @()
                    try {
                        # Get current user's Exchange role assignments
                        if ($userPrincipalName) {
                            $roleAssignments = Get-ManagementRoleAssignment -RoleAssignee $userPrincipalName -ErrorAction SilentlyContinue
                            if ($roleAssignments) {
                                foreach ($assignment in $roleAssignments) {
                                    $exchangeRoles += $assignment.Role.Name
                                }
                            }
                            
                            # Check role group memberships
                            $roleGroups = Get-RoleGroup -ErrorAction SilentlyContinue
                            foreach ($group in $roleGroups) {
                                $members = Get-RoleGroupMember -Identity $group.Name -ErrorAction SilentlyContinue
                                if ($members | Where-Object { $_.Name -eq $userPrincipalName -or $_.WindowsLiveID -eq $userPrincipalName }) {
                                    $roleGroupMembership += $group.Name
                                }
                            }
                        }
                    } catch {
                        # Silently fail - we'll provide generic guidance
                    }
                    
                    $roleInfo = ""
                    if ($exchangeRoles.Count -gt 0) {
                        $roleInfo = "`n`nYour Exchange Roles: $($exchangeRoles -join ', ')"
                    }
                    if ($roleGroupMembership.Count -gt 0) {
                        $roleInfo += "`nYour Role Groups: $($roleGroupMembership -join ', ')"
                    }
                    if ($roleInfo -eq "") {
                        $roleInfo = "`n`nNote: Even as a Global Administrator, you need Exchange Online role group membership."
                    }
                    
                    $roleInfoText = ""
                    if ($actualRoleGroups.Count -gt 0) {
                        $roleInfoText = "`n`nYour Detected Role Groups:`n$($actualRoleGroups -join "`n")`n`n"
                    } else {
                        # Try to detect role groups using the user principal name we have
                        $detectedGroups = @()
                        try {
                            if ($userPrincipalName) {
                                # Try Get-ManagementRoleAssignment method
                                $assignments = Get-ManagementRoleAssignment -RoleAssignee $userPrincipalName -ErrorAction SilentlyContinue
                                if ($assignments) {
                                    $detectedGroups = $assignments | Where-Object { $_.RoleAssigneeType -eq 'RoleGroup' } | 
                                                     Select-Object -ExpandProperty RoleAssigneeName -Unique
                                }
                                
                                # If that didn't work, try checking common groups directly
                                if ($detectedGroups.Count -eq 0) {
                                    $commonGroups = @("Organization Management", "Compliance Management", "Records Management", "View-Only Organization Management")
                                    foreach ($groupName in $commonGroups) {
                                        try {
                                            $members = Get-RoleGroupMember -Identity $groupName -ErrorAction SilentlyContinue
                                            if ($members | Where-Object { 
                                                $_.WindowsLiveID -eq $userPrincipalName -or 
                                                $_.PrimarySmtpAddress -eq $userPrincipalName -or 
                                                $_.Name -eq $userPrincipalName 
                                            }) {
                                                $detectedGroups += $groupName
                                            }
                                        } catch {}
                                    }
                                }
                            }
                        } catch {}
                        
                        if ($detectedGroups.Count -gt 0) {
                            $roleInfoText = "`n`nYour Detected Role Groups:`n$($detectedGroups -join "`n")`n`n"
                        } else {
                            $roleInfoText = "`n`nCould not automatically detect your role group memberships.`n"
                            $roleInfoText += "You may need to manually verify your roles or disconnect/reconnect.`n`n"
                        }
                    }
                    
                    $roleError = "`n`nYou have the required role groups (Organization Management and Compliance Management), but still getting a permission error.`n`n$roleInfoText"
                    
                    if ($transportHygieneFound) {
                        $roleError += "✓ Transport Hygiene role IS assigned to your role groups.`n"
                        $roleError += "⚠ Your Exchange Online session has CACHED OLD PERMISSIONS.`n"
                        $roleError += "You MUST disconnect and reconnect to refresh your session.`n`n"
                    } else {
                        $roleError += "CRITICAL: The Get-BlockedSenderAddress cmdlet requires the 'Transport Hygiene' role.`n"
                        $roleError += "Your session may have cached old permissions. You MUST disconnect and reconnect.`n`n"
                    }
                    $roleError += "REQUIRED ACTION:`n"
                    $roleError += "1. Close this Restricted Entities dialog`n"
                    $roleError += "2. Click 'Disconnect' in the Exchange Online tab`n"
                    $roleError += "3. Wait 10 seconds`n"
                    $roleError += "4. Click 'Connect' to reconnect to Exchange Online`n"
                    $roleError += "5. Open Restricted Entities again and try refreshing`n`n"
                    $roleError += "If still failing after disconnect/reconnect:`n"
                    $roleError += "6. Verify Transport Hygiene role: Get-ManagementRoleAssignment -Role 'Transport Hygiene' -RoleAssignee 'rradmin@naviant.com'`n"
                    $roleError += "7. Wait 5-10 minutes for role changes to fully propagate`n"
                    $roleError += "8. Try running directly in PowerShell: Get-BlockedSenderAddress`n`n"
                    $roleError += "Full Error Details:`n$errorDetails`n`n"
                    $roleError += "Note: If this persists after disconnecting/reconnecting, you may need:`n"
                    $roleError += "- Direct role assignment (not just role group membership)`n"
                    $roleError += "- Or the cmdlet may not be available in your Exchange Online environment`n`n"
                    $roleError += "Alternative: Use Microsoft 365 Defender portal: https://security.microsoft.com/restrictedentities"
                    
                    [System.Windows.Forms.MessageBox]::Show("Permission Error Despite Having Required Roles`n`nError: $errorMessage`n`n$roleError", "Permission Issue", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                    $StatusLabelGlobal.Text = "Permission error - try disconnecting and reconnecting"
                    return
                } else {
                    # Re-throw if it's not a permission error
                    throw
                }
            }
            
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
            # Clean up any formatting issues (remove double pipes, etc.)
            $errorMsg = $errorMsg -replace '\|\|', ''
            $errorMsg = $errorMsg.Trim()
            
            # Check if this is a permission error
            if ($errorMsg -match "role definition" -or $errorMsg -match "aren't present" -or $errorMsg -match "permission" -or $errorMsg -match "access denied") {
                $helpText = "`n`nRequired Admin Roles:`n- Security Administrator`n- Global Administrator`n- Organization Management (Exchange role)`n- Compliance Administrator`n`nTo check your roles:`n1. PowerShell: Get-ManagementRoleAssignment -Role 'Transport Hygiene'`n2. Microsoft 365 Admin Center: Roles > Security Administrator`n`nNote: You may need to request these roles from your administrator."
                [System.Windows.Forms.MessageBox]::Show("Permission Denied`n`n$errorMsg$helpText", "Insufficient Permissions", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                $StatusLabelGlobal.Text = "Insufficient permissions for restricted entities"
            } else {
                # Other errors
                $helpText = "`n`nIf this is a permission error, you may need one of these roles:`n- Security Administrator`n- Global Administrator`n- Organization Management (Exchange role)`n- Compliance Administrator"
                [System.Windows.Forms.MessageBox]::Show("Error loading blocked sender addresses:`n`n$errorMsg$helpText", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                $StatusLabelGlobal.Text = "Error loading blocked sender addresses"
            }
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
                
                # Try to remove the blocked sender (this will fail with a clear error if permissions are missing)
                try {
                    Remove-BlockedSenderAddress -SenderAddress $selectedSender -ErrorAction Stop
                } catch {
                    $errorMessage = $_.Exception.Message
                    # Clean up any formatting issues
                    $errorMessage = $errorMessage -replace '\|\|', ''
                    $errorMessage = $errorMessage.Trim()
                    
                    # Check if this is a permission error
                    if ($errorMessage -match "role definition" -or $errorMessage -match "aren't present" -or $errorMessage -match "permission" -or $errorMessage -match "access denied") {
                        $roleError = "`n`nRequired Admin Roles:`n- Security Administrator`n- Global Administrator`n- Organization Management (Exchange role)`n- Compliance Administrator`n`nTo check your roles:`n1. PowerShell: Get-ManagementRoleAssignment -Role 'Transport Hygiene'`n2. Microsoft 365 Admin Center: Roles > Security Administrator`n`nNote: You may need to request these roles from your administrator."
                        [System.Windows.Forms.MessageBox]::Show("Permission Denied`n`n$errorMessage$roleError", "Insufficient Permissions", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                        $StatusLabelGlobal.Text = "Insufficient permissions for restricted entities"
                        return
                    } else {
                        # Re-throw if it's not a permission error
                        throw
                    }
                }
                
                # Remove from list view
                $listView.SelectedItems[0].Remove()
                
                $StatusLabelGlobal.Text = "Successfully removed $selectedSender from blocked senders"
                [System.Windows.Forms.MessageBox]::Show("Successfully removed '$selectedSender' from blocked senders.", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                
            } catch {
                $errorMsg = $_.Exception.Message
                # Clean up any formatting issues
                $errorMsg = $errorMsg -replace '\|\|', ''
                $errorMsg = $errorMsg.Trim()
                
                # Check if this is a permission error
                if ($errorMsg -match "role definition" -or $errorMsg -match "aren't present" -or $errorMsg -match "permission" -or $errorMsg -match "access denied") {
                    $roleError = "`n`nRequired Admin Roles:`n- Security Administrator`n- Global Administrator`n- Organization Management (Exchange role)`n- Compliance Administrator`n`nTo check your roles:`n1. PowerShell: Get-ManagementRoleAssignment -Role 'Transport Hygiene'`n2. Microsoft 365 Admin Center: Roles > Security Administrator"
                    [System.Windows.Forms.MessageBox]::Show("Permission Denied`n`n$errorMsg$roleError", "Insufficient Permissions", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                    $StatusLabelGlobal.Text = "Insufficient permissions for restricted entities"
                } else {
                    [System.Windows.Forms.MessageBox]::Show("Error removing blocked sender:`n`n$errorMsg`n`nPlease ensure you have the required permissions.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                    $StatusLabelGlobal.Text = "Error removing blocked sender"
                }
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
    
    # Manage Roles button click event
    $manageRolesButton.add_Click({
        Show-ExchangeRoleGroupManager -ParentForm $form -StatusLabelGlobal $StatusLabelGlobal
    })
    
    # Load data initially
    & $loadBlockedSenders
    
    # Show the dialog
    $form.ShowDialog($ParentForm)
}

function Show-ExchangeRoleGroupManager {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [System.Windows.Forms.Form]$ParentForm,
        
        [Parameter(Mandatory=$true)]
        [System.Windows.Forms.ToolStripStatusLabel]$StatusLabelGlobal
    )
    
    # Create the role group manager dialog
    $roleForm = New-Object System.Windows.Forms.Form
    $roleForm.Text = "Exchange Online Role Group Manager"
    $roleForm.Size = New-Object System.Drawing.Size(700, 550)
    $roleForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
    $roleForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $roleForm.MaximizeBox = $false
    $roleForm.MinimizeBox = $false
    
    # Get current user info - try multiple methods (prioritize actual authenticated user)
    $currentUserUpn = $null
    try {
        # Method 1: Try to get authenticated user from Exchange session
        try {
            $exchangeSession = Get-PSSession | Where-Object { $_.ConfigurationName -eq "Microsoft.Exchange" -and $_.State -eq "Opened" } | Select-Object -First 1
            if ($exchangeSession) {
                # Try to invoke command in session to get current authenticated user
                try {
                    $sessionUser = Invoke-Command -Session $exchangeSession -ScriptBlock { 
                        try {
                            $currentUser = Get-User -Identity $env:USERNAME -ErrorAction SilentlyContinue
                            if ($currentUser -and $currentUser.PrimarySmtpAddress) {
                                return $currentUser.PrimarySmtpAddress
                            }
                        } catch {}
                        return $null
                    } -ErrorAction SilentlyContinue
                    if ($sessionUser) {
                        $currentUserUpn = $sessionUser
                    }
                } catch {}
                
                # Fallback: Extract from session name if it contains email
                if (-not $currentUserUpn) {
                    $sessionName = $exchangeSession.Name
                    if ($sessionName -match '@([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z]{2,})') {
                        $currentUserUpn = $matches[1]
                    }
                }
            }
        } catch {}
        
        # Method 2: Try Get-Mailbox with username across all accepted domains (prioritize primary domain)
        if (-not $currentUserUpn) {
            try {
                $domains = Get-AcceptedDomain -ErrorAction SilentlyContinue | Select-Object -ExpandProperty DomainName
                if ($domains) {
                    # Sort domains - prefer domains without "inc" or "corp" suffixes, and shorter domains
                    $sortedDomains = $domains | Sort-Object { 
                        $score = 0
                        if ($_ -notmatch 'inc|corp') { $score += 10 }
                        $score += (20 - $_.Length)  # Prefer shorter domains
                        return $score
                    } -Descending
                    
                    foreach ($domain in $sortedDomains) {
                        $testUpn = "$env:USERNAME@$domain"
                        try {
                            $testMailbox = Get-Mailbox -Identity $testUpn -ErrorAction SilentlyContinue
                            if ($testMailbox -and $testMailbox.PrimarySmtpAddress) {
                                $currentUserUpn = $testMailbox.PrimarySmtpAddress
                                Write-Host "Found user in domain: $domain" -ForegroundColor Green
                                break
                            }
                        } catch {}
                    }
                }
            } catch {}
        }
        
        # Method 3: Try Get-Recipient
        if (-not $currentUserUpn) {
            try {
                $recipient = Get-Recipient -Identity $env:USERNAME -ErrorAction SilentlyContinue
                if ($recipient -and $recipient.PrimarySmtpAddress) {
                    $currentUserUpn = $recipient.PrimarySmtpAddress
                }
            } catch {}
        }
        
            # Method 4: Try Get-Mailbox with filter (but prefer current domain, avoid "inc" domains)
            if (-not $currentUserUpn) {
                try {
                    $mailboxes = Get-Mailbox -Filter "Alias -eq '$env:USERNAME'" -ResultSize 10 -ErrorAction SilentlyContinue
                    if ($mailboxes) {
                        $domains = Get-AcceptedDomain -ErrorAction SilentlyContinue | Select-Object -ExpandProperty DomainName
                        # First pass: prefer domains in accepted domains list that don't have "inc" in them
                        foreach ($mbx in $mailboxes) {
                            if ($mbx.PrimarySmtpAddress) {
                                $mbxDomain = ($mbx.PrimarySmtpAddress -split '@')[1]
                                if ($domains -contains $mbxDomain -and $mbxDomain -notmatch 'inc') {
                                    $currentUserUpn = $mbx.PrimarySmtpAddress
                                    Write-Host "Found user in accepted domain (non-inc): $mbxDomain" -ForegroundColor Green
                                    break
                                }
                            }
                        }
                        # Second pass: if still not found, use any accepted domain
                        if (-not $currentUserUpn) {
                            foreach ($mbx in $mailboxes) {
                                if ($mbx.PrimarySmtpAddress) {
                                    $mbxDomain = ($mbx.PrimarySmtpAddress -split '@')[1]
                                    if ($domains -contains $mbxDomain) {
                                        $currentUserUpn = $mbx.PrimarySmtpAddress
                                        Write-Host "Found user in accepted domain: $mbxDomain" -ForegroundColor Yellow
                                        break
                                    }
                                }
                            }
                        }
                    }
                } catch {}
            }
    } catch {
        # Silently fail - user will need to enter manually
    }
    
    # Title label
    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = "Exchange Online Role Group Management"
    $titleLabel.Font = New-Object System.Drawing.Font('Segoe UI', 12, [System.Drawing.FontStyle]::Bold)
    $titleLabel.Location = New-Object System.Drawing.Point(10, 10)
    $titleLabel.Size = New-Object System.Drawing.Size(650, 25)
    $roleForm.Controls.Add($titleLabel)
    
    # Info label
    $infoLabel = New-Object System.Windows.Forms.Label
    $infoLabel.Text = "Note: Global Administrator does not automatically grant Exchange Online permissions. Add yourself to 'Organization Management' to access restricted entities."
    $infoLabel.Location = New-Object System.Drawing.Point(10, 40)
    $infoLabel.Size = New-Object System.Drawing.Size(650, 40)
    $infoLabel.ForeColor = [System.Drawing.Color]::DarkBlue
    $roleForm.Controls.Add($infoLabel)
    
    # User input section
    $userLabel = New-Object System.Windows.Forms.Label
    $userLabel.Text = "User Principal Name (email):"
    $userLabel.Location = New-Object System.Drawing.Point(10, 90)
    $userLabel.Size = New-Object System.Drawing.Size(200, 20)
    $roleForm.Controls.Add($userLabel)
    
    $userTextBox = New-Object System.Windows.Forms.TextBox
    $userTextBox.Location = New-Object System.Drawing.Point(10, 110)
    $userTextBox.Size = New-Object System.Drawing.Size(400, 20)
    if ($currentUserUpn) {
        $userTextBox.Text = $currentUserUpn
    }
    $roleForm.Controls.Add($userTextBox)
    
    $useCurrentButton = New-Object System.Windows.Forms.Button
    $useCurrentButton.Text = "Use Current User"
    $useCurrentButton.Location = New-Object System.Drawing.Point(420, 108)
    $useCurrentButton.Size = New-Object System.Drawing.Size(120, 25)
    $useCurrentButton.add_Click({
        # Re-detect current user when button is clicked (in case connection changed)
        $detectedUser = $null
        try {
            # Method 1: Get authenticated user from Exchange session using Get-OrganizationConfig
            # This shows who is actually authenticated
            try {
                $orgConfig = Get-OrganizationConfig -ErrorAction Stop
                if ($orgConfig -and $orgConfig.Name) {
                    # Try to get current user from the session's authenticated identity
                    # The session's Runspace might have the user info
                    $exchangeSession = Get-PSSession | Where-Object { $_.ConfigurationName -eq "Microsoft.Exchange" -and $_.State -eq "Opened" } | Select-Object -First 1
                    if ($exchangeSession) {
                        # Try to invoke a command in the session to get the current user
                        try {
                            $sessionUser = Invoke-Command -Session $exchangeSession -ScriptBlock { 
                                try {
                                    # Try to get whoami equivalent in Exchange
                                    $currentUser = Get-User -Identity $env:USERNAME -ErrorAction SilentlyContinue
                                    if ($currentUser -and $currentUser.PrimarySmtpAddress) {
                                        return $currentUser.PrimarySmtpAddress
                                    }
                                } catch {}
                                return $null
                            } -ErrorAction SilentlyContinue
                            if ($sessionUser) {
                                $detectedUser = $sessionUser
                            }
                        } catch {}
                    }
                }
            } catch {}
            
            # Method 2: Try to get from Exchange session's Runspace or properties
            if (-not $detectedUser) {
                try {
                    $exchangeSession = Get-PSSession | Where-Object { $_.ConfigurationName -eq "Microsoft.Exchange" -and $_.State -eq "Opened" } | Select-Object -First 1
                    if ($exchangeSession) {
                        # Check if session has UserPrincipalName property
                        if ($exchangeSession.ComputerName -match '@') {
                            $detectedUser = $exchangeSession.ComputerName
                        }
                        # Try to extract from session ID or other properties
                        elseif ($exchangeSession.Id) {
                            # Session name format is usually "ExchangeOnlineInternalSession#0-<email>"
                            $sessionName = $exchangeSession.Name
                            if ($sessionName -match '@([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z]{2,})') {
                                $detectedUser = $matches[1]
                            }
                        }
                    }
                } catch {}
            }
            
            # Method 3: Use Get-Mailbox to find mailbox matching current Windows user (prioritize naviant.com over naviantinc.com)
            if (-not $detectedUser) {
                try {
                    # Get accepted domains to find the right domain
                    $domains = Get-AcceptedDomain -ErrorAction SilentlyContinue | Select-Object -ExpandProperty DomainName
                    if ($domains) {
                        # Sort domains - prefer naviant.com over naviantinc.com (avoid "inc" domains, prefer shorter)
                        $sortedDomains = $domains | Sort-Object { 
                            $score = 0
                            if ($_ -notmatch 'inc|corp') { $score += 10 }
                            $score += (20 - $_.Length)  # Prefer shorter domains (naviant.com wins over naviantinc.com)
                            return $score
                        } -Descending
                        
                        # Try each domain in priority order
                        foreach ($domain in $sortedDomains) {
                            $testUpn = "$env:USERNAME@$domain"
                            try {
                                $testMailbox = Get-Mailbox -Identity $testUpn -ErrorAction SilentlyContinue
                                if ($testMailbox -and $testMailbox.PrimarySmtpAddress) {
                                    $detectedUser = $testMailbox.PrimarySmtpAddress
                                    Write-Host "Found user in prioritized domain: $domain" -ForegroundColor Green
                                    break
                                }
                            } catch {
                                # Try next domain
                            }
                        }
                    }
                } catch {}
            }
            
            # Method 4: Try Get-Recipient with current username
            if (-not $detectedUser) {
                try {
                    $recipient = Get-Recipient -Identity $env:USERNAME -ErrorAction SilentlyContinue
                    if ($recipient -and $recipient.PrimarySmtpAddress) {
                        $detectedUser = $recipient.PrimarySmtpAddress
                    }
                } catch {}
            }
            
            # Method 5: Try Get-Mailbox with filter on alias
            if (-not $detectedUser) {
                try {
                    $mailboxes = Get-Mailbox -Filter "Alias -eq '$env:USERNAME'" -ResultSize 10 -ErrorAction SilentlyContinue
                    if ($mailboxes) {
                        # Prefer mailboxes in the current tenant's domains
                        $domains = Get-AcceptedDomain -ErrorAction SilentlyContinue | Select-Object -ExpandProperty DomainName
                        foreach ($mbx in $mailboxes) {
                            if ($mbx.PrimarySmtpAddress) {
                                $mbxDomain = ($mbx.PrimarySmtpAddress -split '@')[1]
                                if ($domains -contains $mbxDomain) {
                                    $detectedUser = $mbx.PrimarySmtpAddress
                                    break
                                }
                            }
                        }
                        # If no domain match, use first result
                        if (-not $detectedUser -and $mailboxes[0].PrimarySmtpAddress) {
                            $detectedUser = $mailboxes[0].PrimarySmtpAddress
                        }
                    }
                } catch {}
            }
        } catch {
            # Silently fail
        }
        
        if ($detectedUser) {
            $userTextBox.Text = $detectedUser
            $StatusLabelGlobal.Text = "Current user detected: $detectedUser"
            Write-Host "Detected current user: $detectedUser" -ForegroundColor Green
        } else {
            [System.Windows.Forms.MessageBox]::Show("Could not automatically detect your email address.`n`nPlease enter your User Principal Name (email) manually.`n`nYou are logged in as: rradmin@naviant.com`n`nTip: Enter the email address you used to connect to Exchange Online.", "User Detection Failed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            $StatusLabelGlobal.Text = "Please enter email address manually"
        }
    })
    $roleForm.Controls.Add($useCurrentButton)
    
    # Role group selection
    $roleGroupLabel = New-Object System.Windows.Forms.Label
    $roleGroupLabel.Text = "Role Group:"
    $roleGroupLabel.Location = New-Object System.Drawing.Point(10, 145)
    $roleGroupLabel.Size = New-Object System.Drawing.Size(200, 20)
    $roleForm.Controls.Add($roleGroupLabel)
    
    $roleGroupComboBox = New-Object System.Windows.Forms.ComboBox
    $roleGroupComboBox.Location = New-Object System.Drawing.Point(10, 165)
    $roleGroupComboBox.Size = New-Object System.Drawing.Size(400, 20)
    $roleGroupComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $roleGroupComboBox.Items.AddRange(@("Organization Management", "Compliance Management", "Records Management", "View-Only Organization Management"))
    $roleGroupComboBox.SelectedIndex = 0
    $roleForm.Controls.Add($roleGroupComboBox)
    
    # Current memberships list
    $membershipLabel = New-Object System.Windows.Forms.Label
    $membershipLabel.Text = "Current Role Group Memberships:"
    $membershipLabel.Location = New-Object System.Drawing.Point(10, 200)
    $membershipLabel.Size = New-Object System.Drawing.Size(300, 20)
    $roleForm.Controls.Add($membershipLabel)
    
    $membershipListBox = New-Object System.Windows.Forms.ListBox
    $membershipListBox.Location = New-Object System.Drawing.Point(10, 220)
    $membershipListBox.Size = New-Object System.Drawing.Size(650, 200)
    $membershipListBox.SelectionMode = [System.Windows.Forms.SelectionMode]::None
    $roleForm.Controls.Add($membershipListBox)
    
    # Function to refresh memberships - optimized version
    $refreshMemberships = {
        $membershipListBox.Items.Clear()
        $membershipListBox.Items.Add("Loading...")
        $StatusLabelGlobal.Text = "Checking role group memberships..."
        $refreshMembershipsButton.Enabled = $false
        $roleForm.Refresh()
        
        try {
            $userToCheck = $userTextBox.Text.Trim()
            if ([string]::IsNullOrWhiteSpace($userToCheck)) {
                $membershipListBox.Items.Clear()
                $membershipListBox.Items.Add("Please enter a user principal name")
                $StatusLabelGlobal.Text = "Ready"
                $refreshMembershipsButton.Enabled = $true
                return
            }
            
            $foundMemberships = @()
            
            # Method 1: Use Get-ManagementRoleAssignment (much faster - gets assignments directly)
            try {
                $StatusLabelGlobal.Text = "Checking role assignments (fast method)..."
                $roleForm.Refresh()
                
                $assignments = Get-ManagementRoleAssignment -RoleAssignee $userToCheck -ErrorAction SilentlyContinue
                if ($assignments) {
                    # Get unique role group names from assignments
                    $roleGroupNames = $assignments | Where-Object { $_.RoleAssigneeType -eq 'RoleGroup' } | 
                                     Select-Object -ExpandProperty RoleAssigneeName -Unique
                    
                    foreach ($rgName in $roleGroupNames) {
                        $foundMemberships += $rgName
                    }
                }
            } catch {
                Write-Warning "Fast method failed, using fallback: $($_.Exception.Message)"
            }
            
            # Method 2: If fast method didn't work or found nothing, check common role groups only
            if ($foundMemberships.Count -eq 0) {
                $StatusLabelGlobal.Text = "Checking common role groups..."
                $roleForm.Refresh()
                
                # Only check the most common/relevant role groups instead of all
                $commonRoleGroups = @(
                    "Organization Management",
                    "Compliance Management", 
                    "Records Management",
                    "View-Only Organization Management",
                    "Security Administrator",
                    "Global Administrator",
                    "Recipient Management",
                    "Help Desk",
                    "Discovery Management"
                )
                
                foreach ($groupName in $commonRoleGroups) {
                    try {
                        $members = Get-RoleGroupMember -Identity $groupName -ErrorAction SilentlyContinue
                        if ($members) {
                            foreach ($member in $members) {
                                if ($member.WindowsLiveID -eq $userToCheck -or 
                                    $member.Name -eq $userToCheck -or 
                                    $member.PrimarySmtpAddress -eq $userToCheck) {
                                    if ($foundMemberships -notcontains $groupName) {
                                        $foundMemberships += $groupName
                                    }
                                    break
                                }
                            }
                        }
                    } catch {
                        # Group might not exist, continue
                    }
                }
            }
            
            # Update UI
            $membershipListBox.Items.Clear()
            
            if ($foundMemberships.Count -gt 0) {
                foreach ($membership in $foundMemberships | Sort-Object) {
                    $membershipListBox.Items.Add("✓ $membership")
                }
                $StatusLabelGlobal.Text = "Found $($foundMemberships.Count) role group membership(s)"
            } else {
                $membershipListBox.Items.Add("No role group memberships found for this user")
                $membershipListBox.Items.Add("")
                $membershipListBox.Items.Add("Note: This checks common role groups only.")
                $membershipListBox.Items.Add("If you know you're in a specific group, try adding yourself.")
                $StatusLabelGlobal.Text = "No role group memberships found"
            }
        } catch {
            $membershipListBox.Items.Clear()
            $membershipListBox.Items.Add("Error: $($_.Exception.Message)")
            $StatusLabelGlobal.Text = "Error checking memberships"
        } finally {
            $refreshMembershipsButton.Enabled = $true
        }
    }
    
    # Refresh button
    $refreshMembershipsButton = New-Object System.Windows.Forms.Button
    $refreshMembershipsButton.Text = "Refresh Memberships"
    $refreshMembershipsButton.Location = New-Object System.Drawing.Point(10, 430)
    $refreshMembershipsButton.Size = New-Object System.Drawing.Size(140, 30)
    $refreshMembershipsButton.add_Click($refreshMemberships)
    $roleForm.Controls.Add($refreshMembershipsButton)
    
    # Add to role group button
    $addToRoleButton = New-Object System.Windows.Forms.Button
    $addToRoleButton.Text = "Add to Selected Role Group"
    $addToRoleButton.Location = New-Object System.Drawing.Point(160, 430)
    $addToRoleButton.Size = New-Object System.Drawing.Size(180, 30)
    $addToRoleButton.BackColor = [System.Drawing.Color]::LightGreen
    $addToRoleButton.add_Click({
        $userToAdd = $userTextBox.Text.Trim()
        $roleGroupName = $roleGroupComboBox.SelectedItem.ToString()
        
        if ([string]::IsNullOrWhiteSpace($userToAdd)) {
            [System.Windows.Forms.MessageBox]::Show("Please enter a user principal name.", "Input Required", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        
        $confirm = [System.Windows.Forms.MessageBox]::Show("Add '$userToAdd' to '$roleGroupName' role group?`n`nThis will grant Exchange Online permissions.", "Confirm", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
        
        if ($confirm -eq [System.Windows.Forms.DialogResult]::Yes) {
            $StatusLabelGlobal.Text = "Adding user to role group..."
            $roleForm.Refresh()
            
            try {
                # Check if user is already a member
                $members = Get-RoleGroupMember -Identity $roleGroupName -ErrorAction Stop
                $alreadyMember = $false
                foreach ($member in $members) {
                    if ($member.WindowsLiveID -eq $userToAdd -or $member.Name -eq $userToAdd -or $member.PrimarySmtpAddress -eq $userToAdd) {
                        $alreadyMember = $true
                        break
                    }
                }
                
                if ($alreadyMember) {
                    [System.Windows.Forms.MessageBox]::Show("User '$userToAdd' is already a member of '$roleGroupName'.", "Already Member", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                    $StatusLabelGlobal.Text = "User already a member"
                    & $refreshMemberships
                    return
                }
                
                # Add the user
                Add-RoleGroupMember -Identity $roleGroupName -Member $userToAdd -ErrorAction Stop
                
                [System.Windows.Forms.MessageBox]::Show("Successfully added '$userToAdd' to '$roleGroupName' role group.`n`nYou may need to disconnect and reconnect to Exchange Online for permissions to take effect.", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                $StatusLabelGlobal.Text = "User added to role group successfully"
                
                # Refresh the memberships list
                & $refreshMemberships
                
            } catch {
                $errorMsg = $_.Exception.Message
                [System.Windows.Forms.MessageBox]::Show("Error adding user to role group:`n`n$errorMsg`n`nMake sure:`n- You have permission to modify role groups`n- The user exists in Exchange Online`n- You are connected to Exchange Online", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                $StatusLabelGlobal.Text = "Error adding user to role group"
            }
        }
    })
    $roleForm.Controls.Add($addToRoleButton)
    
    # Remove from role group button
    $removeFromRoleButton = New-Object System.Windows.Forms.Button
    $removeFromRoleButton.Text = "Remove from Selected Role Group"
    $removeFromRoleButton.Location = New-Object System.Drawing.Point(350, 430)
    $removeFromRoleButton.Size = New-Object System.Drawing.Size(200, 30)
    $removeFromRoleButton.BackColor = [System.Drawing.Color]::LightCoral
    $removeFromRoleButton.add_Click({
        $userToRemove = $userTextBox.Text.Trim()
        $roleGroupName = $roleGroupComboBox.SelectedItem.ToString()
        
        if ([string]::IsNullOrWhiteSpace($userToRemove)) {
            [System.Windows.Forms.MessageBox]::Show("Please enter a user principal name.", "Input Required", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        
        $confirm = [System.Windows.Forms.MessageBox]::Show("Remove '$userToRemove' from '$roleGroupName' role group?`n`nThis will revoke Exchange Online permissions.", "Confirm", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
        
        if ($confirm -eq [System.Windows.Forms.DialogResult]::Yes) {
            $StatusLabelGlobal.Text = "Removing user from role group..."
            $roleForm.Refresh()
            
            try {
                Remove-RoleGroupMember -Identity $roleGroupName -Member $userToRemove -ErrorAction Stop
                
                [System.Windows.Forms.MessageBox]::Show("Successfully removed '$userToRemove' from '$roleGroupName' role group.", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                $StatusLabelGlobal.Text = "User removed from role group successfully"
                
                # Refresh the memberships list
                & $refreshMemberships
                
            } catch {
                $errorMsg = $_.Exception.Message
                [System.Windows.Forms.MessageBox]::Show("Error removing user from role group:`n`n$errorMsg", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                $StatusLabelGlobal.Text = "Error removing user from role group"
            }
        }
    })
    $roleForm.Controls.Add($removeFromRoleButton)
    
    # Close button
    $closeRoleButton = New-Object System.Windows.Forms.Button
    $closeRoleButton.Text = "Close"
    $closeRoleButton.Location = New-Object System.Drawing.Point(560, 430)
    $closeRoleButton.Size = New-Object System.Drawing.Size(100, 30)
    $closeRoleButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $roleForm.Controls.Add($closeRoleButton)
    
    # Load initial memberships if user is specified
    if ($currentUserUpn) {
        & $refreshMemberships
    }
    
    # Show the dialog
    $roleForm.ShowDialog($ParentForm)
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