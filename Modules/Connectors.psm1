function Show-ConnectorsViewer {
    param($mainForm, $statusLabel)
    # --- Create and Show Connectors Viewer Form ---
    $connectorsForm = New-Object System.Windows.Forms.Form
    $connectorsForm.Text = "Connectors Viewer"
    $connectorsForm.Size = New-Object System.Drawing.Size(800, 500)
    $connectorsForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
    $connectorsForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $connectorsForm.MaximizeBox = $true

    $connectorsListView = New-Object System.Windows.Forms.ListView
    $connectorsListView.Location = New-Object System.Drawing.Point(10, 10)
    $connectorsListView.Size = New-Object System.Drawing.Size(760, 380)
    $connectorsListView.View = [System.Windows.Forms.View]::Details
    $connectorsListView.FullRowSelect = $true
    $connectorsListView.GridLines = $true
    $connectorsListView.Anchor = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom)
    $connectorsListView.Columns.Clear()
    
    # Fix: Create an array of ColumnHeader objects
    $columns = New-Object 'System.Windows.Forms.ColumnHeader[]' 6
    $columns[0] = New-Object System.Windows.Forms.ColumnHeader; $columns[0].Text = "Name"; $columns[0].Width = 200
    $columns[1] = New-Object System.Windows.Forms.ColumnHeader; $columns[1].Text = "Connector Type"; $columns[1].Width = 120
    $columns[2] = New-Object System.Windows.Forms.ColumnHeader; $columns[2].Text = "Enabled"; $columns[2].Width = 60
    $columns[3] = New-Object System.Windows.Forms.ColumnHeader; $columns[3].Text = "Sender Domains"; $columns[3].Width = 200
    $columns[4] = New-Object System.Windows.Forms.ColumnHeader; $columns[4].Text = "Recipient Domains"; $columns[4].Width = 200
    $columns[5] = New-Object System.Windows.Forms.ColumnHeader; $columns[5].Text = "SmartHosts"; $columns[5].Width = 200
    foreach ($c in $columns) { Write-Host "Column type: $($c.GetType().FullName)" }
    $connectorsListView.Columns.AddRange($columns)
    $connectorsForm.Controls.Add($connectorsListView)

    $refreshButton = New-Object System.Windows.Forms.Button
    $refreshButton.Location = New-Object System.Drawing.Point(10, 400)
    $refreshButton.Size = New-Object System.Drawing.Size(120, 30)
    $refreshButton.Text = "Refresh"
    $refreshButton.Anchor = ([System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left)
    $refreshButton.add_Click({
        $connectorsListView.Items.Clear()
        try {
            $connectors = Get-InboundConnector -ErrorAction SilentlyContinue
            foreach ($connector in $connectors) {
                $item = New-Object System.Windows.Forms.ListViewItem([string]$connector.Name)
                $item.SubItems.Add("Inbound")
                $item.SubItems.Add([string]$connector.Enabled)
                $item.SubItems.Add(($connector.SenderDomains -join ", "))
                $item.SubItems.Add(($connector.RecipientDomains -join ", "))
                $item.SubItems.Add(($connector.SmartHosts -join ", "))
                $connectorsListView.Items.Add($item)
            }
            $connectors = Get-OutboundConnector -ErrorAction SilentlyContinue
            foreach ($connector in $connectors) {
                $item = New-Object System.Windows.Forms.ListViewItem([string]$connector.Name)
                $item.SubItems.Add("Outbound")
                $item.SubItems.Add([string]$connector.Enabled)
                $item.SubItems.Add(($connector.SenderDomains -join ", "))
                $item.SubItems.Add(($connector.RecipientDomains -join ", "))
                $item.SubItems.Add(($connector.SmartHosts -join ", "))
                $connectorsListView.Items.Add($item)
            }
            $statusLabel.Text = "Loaded connectors."
        } catch {
            $ex = $_.Exception
            [System.Windows.Forms.MessageBox]::Show(("Error loading connectors:`n{0}" -f $ex.Message), "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            $statusLabel.Text = "Error loading connectors."
        }
    })
    $connectorsForm.Controls.Add($refreshButton)

    $closeButton = New-Object System.Windows.Forms.Button
    $closeButton.Location = New-Object System.Drawing.Point(650, 400)
    $closeButton.Size = New-Object System.Drawing.Size(120, 30)
    $closeButton.Text = "Close"
    $closeButton.Anchor = ([System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right)
    $closeButton.add_Click({ $connectorsForm.Close() })
    $connectorsForm.Controls.Add($closeButton)

    # Load connectors on form show
    $connectorsForm.Add_Shown({ $refreshButton.PerformClick() })
    [void]$connectorsForm.ShowDialog($mainForm)
    $connectorsForm.Dispose()
}
Export-ModuleMember -Function Show-ConnectorsViewer 