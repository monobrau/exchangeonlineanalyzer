function Show-SessionRevocationTool {
    param($mainForm, $statusLabel, $allLoadedMailboxUPNs)
    # --- Create and Show Session Revocation Tool Form ---
    $sessionForm = New-Object System.Windows.Forms.Form
    $sessionForm.Text = "Revoke User Sessions (Graph)"
    $sessionForm.Size = New-Object System.Drawing.Size(600, 400)
    $sessionForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
    $sessionForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $sessionForm.MaximizeBox = $true

    $userListBox = New-Object System.Windows.Forms.CheckedListBox
    $userListBox.Location = New-Object System.Drawing.Point(20, 20)
    $userListBox.Size = New-Object System.Drawing.Size(540, 200)
    $userListBox.CheckOnClick = $true
    $userListBox.Anchor = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right)
    foreach ($upn in $allLoadedMailboxUPNs) { $userListBox.Items.Add($upn, $false) }
    $sessionForm.Controls.Add($userListBox)

    $selectAllButton = New-Object System.Windows.Forms.Button
    $selectAllButton.Location = New-Object System.Drawing.Point(20, 230)
    $selectAllButton.Size = New-Object System.Drawing.Size(120, 30)
    $selectAllButton.Text = "Select All"
    $selectAllButton.Anchor = ([System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left)
    $selectAllButton.add_Click({ for ($i = 0; $i -lt $userListBox.Items.Count; $i++) { $userListBox.SetItemChecked($i, $true) } })
    $sessionForm.Controls.Add($selectAllButton)

    $deselectAllButton = New-Object System.Windows.Forms.Button
    $deselectAllButton.Location = New-Object System.Drawing.Point(150, 230)
    $deselectAllButton.Size = New-Object System.Drawing.Size(120, 30)
    $deselectAllButton.Text = "Deselect All"
    $deselectAllButton.Anchor = ([System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left)
    $deselectAllButton.add_Click({ for ($i = 0; $i -lt $userListBox.Items.Count; $i++) { $userListBox.SetItemChecked($i, $false) } })
    $sessionForm.Controls.Add($deselectAllButton)

    $revokeButton = New-Object System.Windows.Forms.Button
    $revokeButton.Location = New-Object System.Drawing.Point(20, 280)
    $revokeButton.Size = New-Object System.Drawing.Size(250, 40)
    $revokeButton.Text = "Revoke Sessions for Selected"
    $revokeButton.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 10, [System.Drawing.FontStyle]::Bold)
    $revokeButton.Anchor = ([System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left)
    $revokeButton.add_Click({
        $selectedUpns = $userListBox.CheckedItems | ForEach-Object { $_ }
        if ($selectedUpns.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Select at least one user to revoke sessions.", "No Users Selected", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        
        # Check if we're connected to Microsoft Graph
        try {
            $context = Get-MgContext -ErrorAction Stop
            if (-not $context) {
                [System.Windows.Forms.MessageBox]::Show("Not connected to Microsoft Graph. Please connect first.", "Connection Required", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                return
            }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Microsoft Graph connection required. Please connect first.", "Connection Required", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        
        # Ensure required Graph modules are imported (on-demand import)
        $statusLabel.Text = "Loading required modules..."
        $sessionForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        [System.Windows.Forms.Application]::DoEvents()
        
        try {
            # Check if Revoke-MgUserSignInSession cmdlet is available
            if (-not (Get-Command Revoke-MgUserSignInSession -ErrorAction SilentlyContinue)) {
                Write-Host "Revoke-MgUserSignInSession not found. Importing required module..." -ForegroundColor Yellow
                
                # Try to use Ensure-GraphCmdletsAvailable if available (from GraphOnline module)
                if (Get-Command Ensure-GraphCmdletsAvailable -ErrorAction SilentlyContinue) {
                    $cmdletCheck = Ensure-GraphCmdletsAvailable -CmdletNames @("Revoke-MgUserSignInSession")
                    if (-not $cmdletCheck.AllAvailable) {
                        $missingMsg = "Required Graph modules are not available. Missing cmdlets: $($cmdletCheck.Missing -join ', ')`n`nPlease ensure Microsoft Graph modules are installed.`n`nYou may need to install: Install-Module Microsoft.Graph.Users.Actions -Scope CurrentUser"
                        [System.Windows.Forms.MessageBox]::Show($missingMsg, "Module Required", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                        $sessionForm.Cursor = [System.Windows.Forms.Cursors]::Default
                        return
                    }
                } else {
                    # Fallback: Try to import the module directly
                    try {
                        if (Get-Module -ListAvailable -Name Microsoft.Graph.Users.Actions -ErrorAction SilentlyContinue) {
                            Import-Module Microsoft.Graph.Users.Actions -ErrorAction Stop -Force
                            Write-Host "Imported Microsoft.Graph.Users.Actions module" -ForegroundColor Green
                        } else {
                            throw "Microsoft.Graph.Users.Actions module is not installed"
                        }
                    } catch {
                        $errorMsg = "Failed to import required module Microsoft.Graph.Users.Actions.`n`nError: $($_.Exception.Message)`n`nPlease install it with:`nInstall-Module Microsoft.Graph.Users.Actions -Scope CurrentUser"
                        [System.Windows.Forms.MessageBox]::Show($errorMsg, "Module Import Failed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                        $sessionForm.Cursor = [System.Windows.Forms.Cursors]::Default
                        return
                    }
                }
                
                # Verify cmdlet is now available
                if (-not (Get-Command Revoke-MgUserSignInSession -ErrorAction SilentlyContinue)) {
                    [System.Windows.Forms.MessageBox]::Show("Revoke-MgUserSignInSession cmdlet is still not available after module import. Please ensure Microsoft.Graph.Users.Actions module is properly installed.", "Cmdlet Not Available", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                    $sessionForm.Cursor = [System.Windows.Forms.Cursors]::Default
                    return
                }
            }
        } catch {
            $errorMsg = "Failed to prepare Graph modules for session revocation.`n`nError: $($_.Exception.Message)"
            [System.Windows.Forms.MessageBox]::Show($errorMsg, "Module Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            $sessionForm.Cursor = [System.Windows.Forms.Cursors]::Default
            return
        }
        
        $statusLabel.Text = "Revoking sessions for $($selectedUpns.Count) user(s)..."
        [System.Windows.Forms.Application]::DoEvents()
        
        $successCount = 0
        $failCount = 0
        $errorDetails = @()
        
        foreach ($upn in $selectedUpns) {
            try {
                Revoke-MgUserSignInSession -UserId $upn -ErrorAction Stop
                $successCount++
                Write-Host "Successfully revoked sessions for: $upn" -ForegroundColor Green
            } catch {
                $failCount++
                $errorMsg = "Failed to revoke sessions for $upn`: $($_.Exception.Message)"
                Write-Error $errorMsg
                $errorDetails += $errorMsg
            }
        }
        $sessionForm.Cursor = [System.Windows.Forms.Cursors]::Default
        
        # Build result message
        $resultMsg = "Revoked sessions for $successCount user(s)."
        if ($failCount -gt 0) {
            $resultMsg += "`nFailed for $failCount user(s)."
            if ($errorDetails.Count -gt 0) {
                $resultMsg += "`n`nErrors:`n" + ($errorDetails -join "`n")
            }
        }
        
        $statusLabel.Text = $resultMsg
        $icon = if ($failCount -eq 0) { [System.Windows.Forms.MessageBoxIcon]::Information } else { [System.Windows.Forms.MessageBoxIcon]::Warning }
        [System.Windows.Forms.MessageBox]::Show($resultMsg, "Session Revocation Result", [System.Windows.Forms.MessageBoxButtons]::OK, $icon)
    })
    $sessionForm.Controls.Add($revokeButton)

    $closeButton = New-Object System.Windows.Forms.Button
    $closeButton.Location = New-Object System.Drawing.Point(400, 280)
    $closeButton.Size = New-Object System.Drawing.Size(120, 40)
    $closeButton.Text = "Close"
    $closeButton.Anchor = ([System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right)
    $closeButton.add_Click({ $sessionForm.Close() })
    $sessionForm.Controls.Add($closeButton)

    [void]$sessionForm.ShowDialog($mainForm)
    $sessionForm.Dispose()
}
Export-ModuleMember -Function Show-SessionRevocationTool 