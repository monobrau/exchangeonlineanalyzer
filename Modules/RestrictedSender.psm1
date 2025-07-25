function Show-RestrictedSenderManagementDialog {
    param(
        [Parameter(Mandatory=$true)]
        [string]$UserPrincipalName,
        [Parameter(Mandatory=$true)]
        [System.Windows.Forms.Form]$ParentForm,
        [Parameter(Mandatory=$true)]
        [System.Windows.Forms.ToolStripStatusLabel]$StatusLabelGlobal
    )
    try {
        $StatusLabelGlobal.Text = "Loading restricted senders..."
        $restrictedUsers = $null
        try {
            $restrictedUsers = Get-BlockedSenderAddress -ErrorAction Stop
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error loading restricted senders: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
        $userEntry = $restrictedUsers | Where-Object { $_.User -eq $UserPrincipalName }
        $form = New-Object System.Windows.Forms.Form
        $form.Text = "Manage Restricted Senders for $UserPrincipalName"
        $form.Size = New-Object System.Drawing.Size(500, 300)
        $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
        $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
        $form.MaximizeBox = $false
        $form.MinimizeBox = $false
        $form.TopMost = $true

        $label = New-Object System.Windows.Forms.Label
        $label.Text = "Restricted senders for this mailbox:"
        $label.AutoSize = $true
        $label.Location = New-Object System.Drawing.Point(10, 10)
        $form.Controls.Add($label)

        # Use DataGridView for consistency
        $grid = New-Object System.Windows.Forms.DataGridView
        $grid.Location = New-Object System.Drawing.Point(10, 35)
        $grid.Size = New-Object System.Drawing.Size(460, 150)
        $grid.ReadOnly = $true
        $grid.AllowUserToAddRows = $false
        $grid.AutoGenerateColumns = $true
        $form.Controls.Add($grid)

        # Prepare DataTable
        $dt = New-Object System.Data.DataTable
        $dt.Columns.Add('User') | Out-Null
        $dt.Columns.Add('BlockedDate') | Out-Null
        if ($userEntry) {
            $dt.Rows.Add($userEntry.User, $userEntry.BlockedDate)
        }
        $grid.DataSource = $dt
        $grid.AutoSizeColumnsMode = 'Fill'
        foreach ($col in $grid.Columns) { $col.AutoSizeMode = 'Fill' }

        $removeButton = New-Object System.Windows.Forms.Button
        $removeButton.Text = "Remove Sending Block"
        $removeButton.Location = New-Object System.Drawing.Point(10, 200)
        $removeButton.Size = New-Object System.Drawing.Size(200, 30)
        $removeButton.Enabled = $userEntry -ne $null
        $form.Controls.Add($removeButton)

        $closeButton = New-Object System.Windows.Forms.Button
        $closeButton.Text = "Close"
        $closeButton.Location = New-Object System.Drawing.Point(270, 200)
        $closeButton.Size = New-Object System.Drawing.Size(200, 30)
        $form.Controls.Add($closeButton)

        $removeButton.Add_Click({
            try {
                Remove-BlockedSenderAddress -User $UserPrincipalName -Confirm:$false -ErrorAction Stop
                $StatusLabelGlobal.Text = "Removed sending block for $UserPrincipalName."
                $dt.Rows.Clear()
                $dt.Rows.Add("(No sending block for this mailbox)", "")
                $removeButton.Enabled = $false
                [System.Windows.Forms.MessageBox]::Show("Sending block removed for $UserPrincipalName.", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            } catch {
                $StatusLabelGlobal.Text = "Failed to remove sending block: $($_.Exception.Message)"
                [System.Windows.Forms.MessageBox]::Show("Failed to remove sending block: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        })
        $closeButton.Add_Click({ $form.Close() })
        $StatusLabelGlobal.Text = ""
        [void]$form.ShowDialog()
    } catch {
        $StatusLabelGlobal.Text = "Error loading restricted senders: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Error loading restricted senders: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}
Export-ModuleMember -Function Show-RestrictedSenderManagementDialog 