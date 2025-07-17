function Show-TransportRulesViewer {
    param($mainForm, $statusLabel)
    Write-Host "Show-TransportRulesViewer function entered"
    # --- Create and Show Transport Rules Viewer Form ---
    $transportRulesForm = New-Object System.Windows.Forms.Form
    $transportRulesForm.Text = "Transport Rules Viewer"
    $transportRulesForm.Size = New-Object System.Drawing.Size(800, 500)
    $transportRulesForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
    $transportRulesForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $transportRulesForm.MaximizeBox = $true

    $rulesListView = New-Object System.Windows.Forms.ListView
    $rulesListView.Location = New-Object System.Drawing.Point(10, 10)
    $rulesListView.Size = New-Object System.Drawing.Size(760, 380)
    $rulesListView.View = [System.Windows.Forms.View]::Details
    $rulesListView.FullRowSelect = $true
    $rulesListView.GridLines = $true
    $rulesListView.Anchor = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom)
    $rulesListView.Columns.Clear()
    
    # Fix: Create an array of ColumnHeader objects
    $columns = New-Object 'System.Windows.Forms.ColumnHeader[]' 6
    $columns[0] = New-Object System.Windows.Forms.ColumnHeader; $columns[0].Text = "Name"; $columns[0].Width = 200
    $columns[1] = New-Object System.Windows.Forms.ColumnHeader; $columns[1].Text = "Priority"; $columns[1].Width = 60
    $columns[2] = New-Object System.Windows.Forms.ColumnHeader; $columns[2].Text = "Enabled"; $columns[2].Width = 60
    $columns[3] = New-Object System.Windows.Forms.ColumnHeader; $columns[3].Text = "Mode"; $columns[3].Width = 80
    $columns[4] = New-Object System.Windows.Forms.ColumnHeader; $columns[4].Text = "Comments"; $columns[4].Width = 200
    $columns[5] = New-Object System.Windows.Forms.ColumnHeader; $columns[5].Text = "Actions"; $columns[5].Width = 200
    foreach ($c in $columns) { Write-Host "Column type: $($c.GetType().FullName)" }
    $rulesListView.Columns.AddRange($columns)
    $transportRulesForm.Controls.Add($rulesListView)

    $refreshButton = New-Object System.Windows.Forms.Button
    $refreshButton.Location = New-Object System.Drawing.Point(10, 400)
    $refreshButton.Size = New-Object System.Drawing.Size(120, 30)
    $refreshButton.Text = "Refresh"
    $refreshButton.Anchor = ([System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left)
    $refreshButton.add_Click({
        $rulesListView.Items.Clear()
        try {
            $rules = Get-TransportRule -ErrorAction Stop | Sort-Object Priority
            foreach ($rule in $rules) {
                $actions = ($rule.Actions | Out-String).Trim()
                $item = New-Object System.Windows.Forms.ListViewItem([string]$rule.Name)
                $item.SubItems.Add([string]$rule.Priority)
                $item.SubItems.Add([string]$rule.Enabled)
                $item.SubItems.Add([string]$rule.Mode)
                $item.SubItems.Add([string]$rule.Comments)
                $item.SubItems.Add([string]$actions)
                $rulesListView.Items.Add($item)
            }
            $statusLabel.Text = "Loaded $($rules.Count) transport rules."
        } catch {
            $ex = $_.Exception
            [System.Windows.Forms.MessageBox]::Show(("Error loading transport rules:`n{0}" -f $ex.Message), "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            $statusLabel.Text = "Error loading transport rules."
        }
    })
    $transportRulesForm.Controls.Add($refreshButton)

    $closeButton = New-Object System.Windows.Forms.Button
    $closeButton.Location = New-Object System.Drawing.Point(650, 400)
    $closeButton.Size = New-Object System.Drawing.Size(120, 30)
    $closeButton.Text = "Close"
    $closeButton.Anchor = ([System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right)
    $closeButton.add_Click({ $transportRulesForm.Close() })
    $transportRulesForm.Controls.Add($closeButton)

    # Load rules on form show
    $transportRulesForm.Add_Shown({ $refreshButton.PerformClick() })
    [void]$transportRulesForm.ShowDialog($mainForm)
    $transportRulesForm.Dispose()
}
Export-ModuleMember -Function Show-TransportRulesViewer 