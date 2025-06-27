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

#Requires -Modules ExchangeOnlineManagement

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

# --- Function Definitions ---

Function Test-ExchangeModule {
    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        return $false
    }
    return $true
}

Function Install-ExchangeModule {
    Write-Host "Attempting to install ExchangeOnlineManagement module..." -ForegroundColor Yellow
    try {
        Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Repository PSGallery -Force -AllowClobber -ErrorAction Stop
        Write-Host "ExchangeOnlineManagement module installed successfully. Please restart the script." -ForegroundColor Green
        return $true
    } catch {
        $ex = $_.Exception 
        Write-Error ("Failed to install ExchangeOnlineManagement module. Please install it manually: Install-Module ExchangeOnlineManagement -Scope CurrentUser. Error: {0}" -f $ex.Message)
        if ($statusLabel) {
            $statusLabel.Text = "Error installing Exchange module. See console."
        }
        return $false
    }
}

Function Test-GraphModules {
    foreach ($moduleInfo in $script:requiredGraphModules) {
        if (-not (Get-Module -ListAvailable -Name $moduleInfo.Name)) {
            Write-Warning "Required Graph module $($moduleInfo.Name) is missing."
            return $false
        }
    }
    Write-Host "All required Microsoft Graph modules are available." -ForegroundColor Green
    return $true
}

Function Install-GraphModules {
    param($statusLabel)
    Write-Host "Attempting to install required Microsoft Graph modules..." -ForegroundColor Yellow
    if ($statusLabel) { $statusLabel.Text = "Installing Graph modules..." }

    $allInstalled = $true
    foreach ($moduleInfo in $script:requiredGraphModules) {
        if (-not (Get-Module -ListAvailable -Name $moduleInfo.Name)) {
            Write-Host "Installing module: $($moduleInfo.Name)..."
            try {
                Install-Module -Name $moduleInfo.Name -Scope CurrentUser -Repository PSGallery -Force -AllowClobber -ErrorAction Stop
                Write-Host "Module $($moduleInfo.Name) installed successfully." -ForegroundColor Green
            } catch {
                $ex = $_.Exception
                Write-Error ("Failed to install module $($moduleInfo.Name). Please install it manually: Install-Module $($moduleInfo.Name) -Scope CurrentUser. Error: {0}" -f $ex.Message)
                if ($statusLabel) { $statusLabel.Text = "Error installing $($moduleInfo.Name). See console." }
                $allInstalled = $false
            }
        }
    }
    if ($allInstalled) {
        Write-Host "All required Graph modules checked/installed. Please restart the script if prompted or if new modules were installed." -ForegroundColor Green
        if ($statusLabel) { $statusLabel.Text = "Graph modules installed/checked. Restart script if needed." }
        return $true
    } else {
        return $false
    }
}

Function Connect-GraphService {
    param($statusLabel, $mainForm)

    if ($script:graphConnection) {
        Write-Host "Already connected to Microsoft Graph as $($script:graphConnection.Account)." -ForegroundColor Cyan
        if ($statusLabel) { $statusLabel.Text = "Already connected to MS Graph: $($script:graphConnection.Account)." }
        return $true
    }

    if (-not (Test-GraphModules)) {
        $choice = [System.Windows.Forms.MessageBox]::Show("Required Microsoft Graph modules are missing. Install now? (Requires script restart after install)", "Missing Graph Modules", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($choice -eq [System.Windows.Forms.DialogResult]::Yes) {
            if (-not (Install-GraphModules -statusLabel $statusLabel)) {
                [System.Windows.Forms.MessageBox]::Show("Graph module installation failed. Please install them manually and restart the script.", "Install Failed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error);
                return $false
            } else {
                 [System.Windows.Forms.MessageBox]::Show("Graph modules installed. Please restart the script to use Graph features.", "Restart Required", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information);
                 Exit 
            }
        } else {
            [System.Windows.Forms.MessageBox]::Show("Microsoft Graph features will be unavailable without the required modules.", "Graph Modules Missing", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning);
            return $false
        }
    }

    Write-Host "Attempting to connect to Microsoft Graph with scopes: $($script:graphScopes -join ', ')" -ForegroundColor Yellow
    if ($statusLabel) { $statusLabel.Text = "Connecting to Microsoft Graph..." }
    if ($mainForm) { $mainForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor }

    try {
        # Ensure modules are imported for the session
        Import-Module Microsoft.Graph.Authentication -ErrorAction SilentlyContinue -Force
        Import-Module Microsoft.Graph.Users -ErrorAction SilentlyContinue -Force
        Import-Module Microsoft.Graph.Identity.SignIns -ErrorAction SilentlyContinue -Force 
        Write-Host "Attempted to import core Graph modules."

        $existingMgConnection = Get-MgContext -ErrorAction SilentlyContinue
        if ($existingMgConnection -and $existingMgConnection.Account) {
            $missingScopes = $script:graphScopes | Where-Object {$existingMgConnection.Scopes -notcontains $_}
            if ($missingScopes.Count -eq 0) {
                $script:graphConnection = $existingMgConnection
                Write-Host "Using existing Microsoft Graph session for $($script:graphConnection.Account) with sufficient scopes." -ForegroundColor Cyan
                if ($statusLabel) { $statusLabel.Text = "Using existing MS Graph session: $($script:graphConnection.Account)." }
                return $true
            } else {
                Write-Warning "Existing Graph session for $($existingMgConnection.Account) is missing required scopes: $($missingScopes -join ', '). Attempting to reconnect with all required scopes."
                Disconnect-MgGraph -ErrorAction SilentlyContinue 
            }
        }

        Connect-MgGraph -Scopes $script:graphScopes -ErrorAction Stop
        $script:graphConnection = Get-MgContext
        Write-Host "Successfully connected to Microsoft Graph as $($script:graphConnection.Account)." -ForegroundColor Green
        if ($statusLabel) { $statusLabel.Text = "MS Graph Connected: $($script:graphConnection.Account)." }
        return $true
    } catch {
        $ex = $_.Exception
        Write-Error ("Microsoft Graph connection failed: {0}" -f $ex.Message)
        if ($statusLabel) { $statusLabel.Text = "MS Graph connection failed. See console." }
        [System.Windows.Forms.MessageBox]::Show(("Failed to connect to Microsoft Graph.`nMake sure you have the necessary permissions and modules installed.`nError: {0}" -f $ex.Message), "MS Graph Connection Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        $script:graphConnection = $null
        return $false
    } finally {
        if ($mainForm) { $mainForm.Cursor = [System.Windows.Forms.Cursors]::Default }
        $script:graphConnectionAttempted = $true
    }
}


Function Format-InboxRuleXlsx {
    param(
        [Parameter(Mandatory=$true)]
        [string]$CsvPath,
        [Parameter(Mandatory=$true)]
        [string]$XlsxPath,
        [Parameter(Mandatory=$false)]
        [int]$RowHighlightColor = $highlightColorIndexYellow,
        [Parameter(Mandatory=$false)]
        [string]$RowHighlightColumnHeader = "IsHidden", 
        [Parameter(Mandatory=$false)]
        [object]$RowHighlightValue = $true,
        [Parameter(Mandatory=$false)]
        [int]$CellHighlightColor = $highlightColorIndexLightRed,
        [Parameter(Mandatory=$false)]
        [object]$CellHighlightValue = $true
    )

    $excel = $null; $workbook = $null; $worksheet = $null; $usedRange = $null; $columns = $null; $rows = $null; $headerRange = $null; $targetColumnIndex = $null
    $xlOpenXMLWorkbook = 51
    $xlExpression = 2 
    $xlCellValue = 1  
    $xlEqual = 3      
    $missing = [System.Reflection.Missing]::Value 

    try { $excel = New-Object -ComObject Excel.Application -ErrorAction Stop } 
    catch { 
        $ex = $_.Exception
        Write-Error ("Excel COM object creation failed: {0}" -f $ex.Message); 
        if ($statusLabel) { $statusLabel.Text = "Error: Excel not found." }; 
        return $false 
    }

    try {
        $excel.Visible = $false; $excel.DisplayAlerts = $false    
        Write-Host "Converting '$CsvPath' to '$XlsxPath'..."
        $workbook = $excel.Workbooks.Open($CsvPath); $workbook.SaveAs($XlsxPath, $xlOpenXMLWorkbook); $workbook.Close($false) 
        Write-Host "Initial conversion successful. Formatting..." -ForegroundColor Green
        $workbook = $excel.Workbooks.Open($XlsxPath); $worksheet = $workbook.Worksheets.Item(1); $usedRange = $worksheet.UsedRange; $columns = $usedRange.Columns; $rows = $usedRange.Rows

        if ($usedRange.Rows.Count -gt 0) {
            Write-Host " - Autofitting columns..."; $columns.AutoFit() | Out-Null
            Write-Host " - Bolding header row..."; $headerRange = $worksheet.Rows.Item(1); $headerRange.Font.Bold = $true

            if ($usedRange.Rows.Count -gt 1) {
                $dataRange = $usedRange.Offset(1,0).Resize($usedRange.Rows.Count -1) 
                Write-Host " - Clearing existing conditional formats from data range..."
                $dataRange.FormatConditions.Delete() | Out-Null

                Write-Host "   - Applying Rule 1: Light Red for any TRUE cell..."
                $formatCondition1 = $dataRange.FormatConditions.Add($xlCellValue, $xlEqual, "TRUE") 
                $formatCondition1.Interior.ColorIndex = $CellHighlightColor
                
                Write-Host "   - Applying Rule 2: Manually searching for '$RowHighlightColumnHeader' column for row highlighting..."
                for ($colIdx = 1; $colIdx -le $headerRange.Columns.Count; $colIdx++) {
                    $cell = $null
                    try {
                        $cell = $headerRange.Cells.Item(1, $colIdx)
                        if ($cell.Value2 -is [string] -and $cell.Value2.Equals($RowHighlightColumnHeader, [System.StringComparison]::OrdinalIgnoreCase)) {
                            $targetColumnIndex = $colIdx
                            Write-Host "     - Found '$RowHighlightColumnHeader' at column index $targetColumnIndex."
                            break
                        }
                    } finally {
                        if ($cell) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($cell) | Out-Null }
                    }
                }

                if ($targetColumnIndex) {
                    Write-Host "     - Highlighting rows where '$RowHighlightColumnHeader' is '$($RowHighlightValue)'..."
                    $columnLetter = $worksheet.Columns.Item($targetColumnIndex).Address($false, $false) -replace '\d','' 
                    $formulaForRowHighlight = "`$${columnLetter}$($dataRange.Row)=$($RowHighlightValue.ToString().ToUpper())" 
                    
                    Write-Host "     - Using formula for row highlight: $formulaForRowHighlight"
                    $formatCondition2 = $dataRange.FormatConditions.Add($xlExpression, $missing, $formulaForRowHighlight) 
                    $formatCondition2.Interior.ColorIndex = $RowHighlightColor
                    Write-Host "     - Row highlighting rule for '$RowHighlightColumnHeader' applied." -ForegroundColor Green
                } else { Write-Warning "   - '$RowHighlightColumnHeader' column not found. Skipping row highlighting." }
            } else { Write-Host " - Only header row found, skipping conditional formatting." }
        } else { Write-Host " - Worksheet appears empty, skipping formatting." }
        
        Write-Host "Saving formatted XLSX file..."; $workbook.Save(); $workbook.Close()
        Write-Host "XLSX formatting complete." -ForegroundColor Green
        $script:lastExportedXlsxPath = $XlsxPath; if ($openFileButton) { $openFileButton.Enabled = $true } 
    } catch {
        $ex = $_.Exception
        Write-Error ("Excel formatting/conversion error: {0}`n{1}" -f $ex.Message, $ex.ScriptStackTrace)
        if ($statusLabel) { $statusLabel.Text = "Error: XLSX formatting failed." }
        try { if ($workbook -ne $null) { $workbook.Close($false) } } catch {}
        return $false 
    } finally {
        if ($formatCondition1) {[System.Runtime.InteropServices.Marshal]::ReleaseComObject($formatCondition1) | Out-Null}
        if ($formatCondition2) {[System.Runtime.InteropServices.Marshal]::ReleaseComObject($formatCondition2) | Out-Null}
        if ($dataRange) {[System.Runtime.InteropServices.Marshal]::ReleaseComObject($dataRange) | Out-Null}
        if ($headerRange) {[System.Runtime.InteropServices.Marshal]::ReleaseComObject($headerRange) | Out-Null}
        if ($columns) {[System.Runtime.InteropServices.Marshal]::ReleaseComObject($columns) | Out-Null}
        if ($rows) {[System.Runtime.InteropServices.Marshal]::ReleaseComObject($rows) | Out-Null}
        if ($usedRange) {[System.Runtime.InteropServices.Marshal]::ReleaseComObject($usedRange) | Out-Null}
        if ($worksheet) {[System.Runtime.InteropServices.Marshal]::ReleaseComObject($worksheet) | Out-Null}
        if ($workbook) {[System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null}
        if ($excel) {$excel.Quit();[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null}
        [gc]::Collect(); [gc]::WaitForPendingFinalizers(); Write-Host "COM cleanup finished."
    }
    return $true 
}

Function Test-ExternalForwarding {
    param(
        [Parameter(Mandatory=$true)]
        [array]$RecipientAddresses,
        [Parameter(Mandatory=$true)]
        [array]$InternalDomains
    )
    if ($null -eq $RecipientAddresses -or $RecipientAddresses.Count -eq 0) { return $false }
    foreach ($recipient in $RecipientAddresses) {
        $emailAddress = $null
        if ($recipient -is [string]) { $emailAddress = $recipient }
        elseif ($recipient -is [Microsoft.Exchange.Data.SmtpAddress]) { $emailAddress = $recipient.ToString() }
        elseif ($recipient.PropertyNames -contains 'Address') { $emailAddress = $recipient.Address } 
        elseif ($recipient.PropertyNames -contains 'EmailAddress') { $emailAddress = $recipient.EmailAddress } 
        if (-not [string]::IsNullOrWhiteSpace($emailAddress)) {
            try {
                $domain = ($emailAddress -split '@')[1].ToLowerInvariant()
                if ($domain -notin $InternalDomains) { Write-Verbose "External domain: $domain"; return $true }
            } catch { Write-Warning "Could not parse domain: $emailAddress" }
        }
    }
    return $false
}

Function Show-TransportRulesViewer {
    param($mainForm, $statusLabel)
    
    $transportRulesForm = New-Object System.Windows.Forms.Form
    $transportRulesForm.Text = "Exchange Online Transport Rules Viewer"
    $transportRulesForm.Size = New-Object System.Drawing.Size(1000, 600)
    $transportRulesForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
    $transportRulesForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
    $transportRulesForm.MinimumSize = New-Object System.Drawing.Size(800, 500)

    # Transport Rules ListView
    $transportRulesListView = New-Object System.Windows.Forms.ListView
    $transportRulesListView.Location = New-Object System.Drawing.Point(15, 45)
    $transportRulesListView.Size = New-Object System.Drawing.Size(950, 400)
    $transportRulesListView.View = [System.Windows.Forms.View]::Details
    $transportRulesListView.FullRowSelect = $true
    $transportRulesListView.GridLines = $true
    $transportRulesListView.Anchor = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom)
    
    # Add columns
    $transportRulesListView.Columns.Add("Name", 200) | Out-Null
    $transportRulesListView.Columns.Add("Enabled", 80) | Out-Null
    $transportRulesListView.Columns.Add("Priority", 70) | Out-Null
    $transportRulesListView.Columns.Add("Mode", 100) | Out-Null
    $transportRulesListView.Columns.Add("State", 100) | Out-Null 
    $transportRulesListView.Columns.Add("Description", 300) | Out-Null
    $transportRulesListView.Columns.Add("Identity", 200) | Out-Null
    
    $transportRulesForm.Controls.Add($transportRulesListView)

    # Load Rules Button
    $loadTransportRulesButton = New-Object System.Windows.Forms.Button
    $loadTransportRulesButton.Location = New-Object System.Drawing.Point(15, 15)
    $loadTransportRulesButton.Size = New-Object System.Drawing.Size(150, 25)
    $loadTransportRulesButton.Text = "Load Transport Rules"
    $loadTransportRulesButton.add_Click({
        $statusLabel.Text = "Loading transport rules..."
        $mainForm.Refresh()
        $transportRulesForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        $transportRulesListView.Items.Clear()
        
        try {
            $transportRules = Get-TransportRule -ErrorAction Stop
            if ($transportRules) {
                foreach ($rule in $transportRules) {
                    $item = New-Object System.Windows.Forms.ListViewItem($rule.Name)
                    $item.SubItems.Add($rule.State.ToString()) | Out-Null 
                    $item.SubItems.Add($rule.Priority.ToString()) | Out-Null
                    $item.SubItems.Add($rule.Mode.ToString()) | Out-Null
                    $item.SubItems.Add($rule.RuleVersion.ToString()) | Out-Null 
                    $descriptionText = if($rule.Description){$rule.Description}else{"N/A"}
                    $item.SubItems.Add($descriptionText) | Out-Null
                    $item.SubItems.Add($rule.Identity.ToString()) | Out-Null
                    $item.Tag = $rule
                    $transportRulesListView.Items.Add($item) | Out-Null
                }
                $statusLabel.Text = "Loaded $($transportRules.Count) transport rules."
            } else {
                $statusLabel.Text = "No transport rules found."
            }
        } catch {
            $ex = $_.Exception
            [System.Windows.Forms.MessageBox]::Show("Error loading transport rules:`n$($ex.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            $statusLabel.Text = "Error loading transport rules."
        } finally {
            $transportRulesForm.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    })
    $transportRulesForm.Controls.Add($loadTransportRulesButton)

    # View Details Button
    $viewDetailsButton = New-Object System.Windows.Forms.Button
    $viewDetailsButton.Location = New-Object System.Drawing.Point(175, 15)
    $viewDetailsButton.Size = New-Object System.Drawing.Size(120, 25)
    $viewDetailsButton.Text = "View Details"
    $viewDetailsButton.add_Click({
        if ($transportRulesListView.SelectedItems.Count -eq 1) {
            $selectedRule = $transportRulesListView.SelectedItems[0].Tag
            $details = $selectedRule | Format-List | Out-String
            
            $detailsForm = New-Object System.Windows.Forms.Form
            $detailsForm.Text = "Transport Rule Details: $($selectedRule.Name)"
            $detailsForm.Size = New-Object System.Drawing.Size(800, 600)
            $detailsForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
            
            $detailsTextBox = New-Object System.Windows.Forms.TextBox
            $detailsTextBox.Location = New-Object System.Drawing.Point(15, 15)
            $detailsTextBox.Size = New-Object System.Drawing.Size(750, 520)
            $detailsTextBox.Multiline = $true
            $detailsTextBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Both
            $detailsTextBox.ReadOnly = $true
            $detailsTextBox.Text = $details
            $detailsTextBox.Anchor = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom)
            $detailsForm.Controls.Add($detailsTextBox)
            
            $closeDetailsButton = New-Object System.Windows.Forms.Button
            $closeDetailsButton.Location = New-Object System.Drawing.Point(350, 550)
            $closeDetailsButton.Size = New-Object System.Drawing.Size(100, 25)
            $closeDetailsButton.Text = "Close"
            $closeDetailsButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom
            $closeDetailsButton.add_Click({ $detailsForm.Close() })
            $detailsForm.Controls.Add($closeDetailsButton)
            
            [void]$detailsForm.ShowDialog($transportRulesForm)
        } else {
            [System.Windows.Forms.MessageBox]::Show("Please select a transport rule to view details.", "No Selection", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
    })
    $transportRulesForm.Controls.Add($viewDetailsButton)

    # Export Rules Button
    $exportTransportRulesButton = New-Object System.Windows.Forms.Button
    $exportTransportRulesButton.Location = New-Object System.Drawing.Point(305, 15)
    $exportTransportRulesButton.Size = New-Object System.Drawing.Size(120, 25)
    $exportTransportRulesButton.Text = "Export to CSV"
    $exportTransportRulesButton.add_Click({
        if ($transportRulesListView.Items.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("No transport rules loaded. Click 'Load Transport Rules' first.", "No Data", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            return
        }
        
        $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveFileDialog.Filter = "CSV files (*.csv)|*.csv"
        $saveFileDialog.DefaultExt = "csv"
        $saveFileDialog.FileName = "TransportRules_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        
        if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            try {
                $exportData = @()
                foreach ($item in $transportRulesListView.Items) {
                    $rule = $item.Tag
                    $exportData += [PSCustomObject]@{
                        Name = $rule.Name
                        Enabled = $rule.State
                        Priority = $rule.Priority
                        Mode = $rule.Mode
                        RuleVersion = $rule.RuleVersion
                        Description = $rule.Description
                        Identity = $rule.Identity
                        Conditions = ($rule.Conditions | Out-String).Trim()
                        Actions = ($rule.Actions | Out-String).Trim()
                        Exceptions = ($rule.Exceptions | Out-String).Trim()
                    }
                }
                $exportData | Export-Csv -Path $saveFileDialog.FileName -NoTypeInformation -Encoding UTF8
                [System.Windows.Forms.MessageBox]::Show("Transport rules exported to:`n$($saveFileDialog.FileName)", "Export Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                $statusLabel.Text = "Transport rules exported to $($saveFileDialog.FileName)"
            } catch {
                $ex = $_.Exception
                [System.Windows.Forms.MessageBox]::Show("Error exporting transport rules:`n$($ex.Message)", "Export Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }
    })
    $transportRulesForm.Controls.Add($exportTransportRulesButton)

    # Close Button
    $closeTransportButton = New-Object System.Windows.Forms.Button
    $closeTransportButton.Location = New-Object System.Drawing.Point(850, 15)
    $closeTransportButton.Size = New-Object System.Drawing.Size(100, 25)
    $closeTransportButton.Text = "Close"
    $closeTransportButton.Anchor = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right)
    $closeTransportButton.add_Click({ $transportRulesForm.Close() })
    $transportRulesForm.Controls.Add($closeTransportButton)

    [void]$transportRulesForm.ShowDialog($mainForm)
}

Function Show-ConnectorsViewer {
    param($mainForm, $statusLabel)
    
    $connectorsForm = New-Object System.Windows.Forms.Form
    $connectorsForm.Text = "Exchange Online Connectors Viewer"
    $connectorsForm.Size = New-Object System.Drawing.Size(1000, 600)
    $connectorsForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
    $connectorsForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
    $connectorsForm.MinimumSize = New-Object System.Drawing.Size(800, 500)

    # Tab Control for Inbound and Outbound Connectors
    $tabControl = New-Object System.Windows.Forms.TabControl
    $tabControl.Location = New-Object System.Drawing.Point(15, 45)
    $tabControl.Size = New-Object System.Drawing.Size(950, 480)
    $tabControl.Anchor = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom)
    
    # Inbound Connectors Tab
    $inboundTab = New-Object System.Windows.Forms.TabPage
    $inboundTab.Text = "Inbound Connectors"
    $inboundTab.UseVisualStyleBackColor = $true
    
    $inboundListView = New-Object System.Windows.Forms.ListView
    $inboundListView.Location = New-Object System.Drawing.Point(10, 10)
    $inboundListView.Size = New-Object System.Drawing.Size(920, 420)
    $inboundListView.View = [System.Windows.Forms.View]::Details
    $inboundListView.FullRowSelect = $true
    $inboundListView.GridLines = $true
    $inboundListView.Anchor = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom)
    
    $inboundListView.Columns.Add("Name", 200) | Out-Null
    $inboundListView.Columns.Add("Enabled", 80) | Out-Null
    $inboundListView.Columns.Add("Connector Type", 150) | Out-Null
    $inboundListView.Columns.Add("Sender Domains", 200) | Out-Null
    $inboundListView.Columns.Add("TLS Domain", 150) | Out-Null 
    $inboundListView.Columns.Add("Identity", 200) | Out-Null
    
    $inboundTab.Controls.Add($inboundListView)
    
    # Outbound Connectors Tab
    $outboundTab = New-Object System.Windows.Forms.TabPage
    $outboundTab.Text = "Outbound Connectors"
    $outboundTab.UseVisualStyleBackColor = $true
    
    $outboundListView = New-Object System.Windows.Forms.ListView
    $outboundListView.Location = New-Object System.Drawing.Point(10, 10)
    $outboundListView.Size = New-Object System.Drawing.Size(920, 420)
    $outboundListView.View = [System.Windows.Forms.View]::Details
    $outboundListView.FullRowSelect = $true
    $outboundListView.GridLines = $true
    $outboundListView.Anchor = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom)
    
    $outboundListView.Columns.Add("Name", 200) | Out-Null
    $outboundListView.Columns.Add("Enabled", 80) | Out-Null
    $outboundListView.Columns.Add("Connector Type", 150) | Out-Null
    $outboundListView.Columns.Add("Recipient Domains", 200) | Out-Null
    $outboundListView.Columns.Add("Smart Hosts", 150) | Out-Null
    $outboundListView.Columns.Add("Identity", 200) | Out-Null
    
    $outboundTab.Controls.Add($outboundListView)
    
    $tabControl.TabPages.Add($inboundTab)
    $tabControl.TabPages.Add($outboundTab)
    $connectorsForm.Controls.Add($tabControl)

    # Load Connectors Button
    $loadConnectorsButton = New-Object System.Windows.Forms.Button
    $loadConnectorsButton.Location = New-Object System.Drawing.Point(15, 15)
    $loadConnectorsButton.Size = New-Object System.Drawing.Size(150, 25)
    $loadConnectorsButton.Text = "Load All Connectors"
    $loadConnectorsButton.add_Click({
        $statusLabel.Text = "Loading connectors..."
        $mainForm.Refresh()
        $connectorsForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        $inboundListView.Items.Clear()
        $outboundListView.Items.Clear()
        
        try {
            # Load Inbound Connectors
            $inboundConnectors = Get-InboundConnector -ErrorAction Stop
            if ($inboundConnectors) {
                foreach ($connector in $inboundConnectors) {
                    $item = New-Object System.Windows.Forms.ListViewItem($connector.Name)
                    $item.SubItems.Add($connector.Enabled.ToString()) | Out-Null
                    $item.SubItems.Add($connector.ConnectorType.ToString()) | Out-Null
                    $item.SubItems.Add(($connector.SenderDomains -join "; ")) | Out-Null
                    $item.SubItems.Add($(if($connector.TlsSenderCertificateName){$connector.TlsSenderCertificateName}else{"N/A"})) | Out-Null 
                    $item.SubItems.Add($connector.Identity.ToString()) | Out-Null
                    $item.Tag = $connector
                    $inboundListView.Items.Add($item) | Out-Null
                }
            }
            
            # Load Outbound Connectors
            $outboundConnectors = Get-OutboundConnector -ErrorAction Stop
            if ($outboundConnectors) {
                foreach ($connector in $outboundConnectors) {
                    $item = New-Object System.Windows.Forms.ListViewItem($connector.Name)
                    $item.SubItems.Add($connector.Enabled.ToString()) | Out-Null
                    $item.SubItems.Add($connector.ConnectorType.ToString()) | Out-Null
                    $item.SubItems.Add(($connector.RecipientDomains -join "; ")) | Out-Null
                    $item.SubItems.Add(($connector.SmartHosts -join "; ")) | Out-Null
                    $item.SubItems.Add($connector.Identity.ToString()) | Out-Null
                    $item.Tag = $connector
                    $outboundListView.Items.Add($item) | Out-Null
                }
            }
            
            $statusLabel.Text = "Loaded $($inboundConnectors.Count) inbound and $($outboundConnectors.Count) outbound connectors."
        } catch {
            $ex = $_.Exception
            [System.Windows.Forms.MessageBox]::Show("Error loading connectors:`n$($ex.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            $statusLabel.Text = "Error loading connectors."
        } finally {
            $connectorsForm.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    })
    $connectorsForm.Controls.Add($loadConnectorsButton)

    # View Details Button
    $viewConnectorDetailsButton = New-Object System.Windows.Forms.Button
    $viewConnectorDetailsButton.Location = New-Object System.Drawing.Point(175, 15)
    $viewConnectorDetailsButton.Size = New-Object System.Drawing.Size(120, 25)
    $viewConnectorDetailsButton.Text = "View Details"
    $viewConnectorDetailsButton.add_Click({
        $selectedConnector = $null
        if ($tabControl.SelectedTab -eq $inboundTab -and $inboundListView.SelectedItems.Count -eq 1) {
            $selectedConnector = $inboundListView.SelectedItems[0].Tag
        } elseif ($tabControl.SelectedTab -eq $outboundTab -and $outboundListView.SelectedItems.Count -eq 1) {
            $selectedConnector = $outboundListView.SelectedItems[0].Tag
        }
        
        if ($selectedConnector) {
            $details = $selectedConnector | Format-List | Out-String
            
            $detailsForm = New-Object System.Windows.Forms.Form
            $detailsForm.Text = "Connector Details: $($selectedConnector.Name)"
            $detailsForm.Size = New-Object System.Drawing.Size(800, 600)
            $detailsForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
            
            $detailsTextBox = New-Object System.Windows.Forms.TextBox
            $detailsTextBox.Location = New-Object System.Drawing.Point(15, 15)
            $detailsTextBox.Size = New-Object System.Drawing.Size(750, 520)
            $detailsTextBox.Multiline = $true
            $detailsTextBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Both
            $detailsTextBox.ReadOnly = $true
            $detailsTextBox.Text = $details
            $detailsTextBox.Anchor = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom)
            $detailsForm.Controls.Add($detailsTextBox)
            
            $closeDetailsButton = New-Object System.Windows.Forms.Button
            $closeDetailsButton.Location = New-Object System.Drawing.Point(350, 550)
            $closeDetailsButton.Size = New-Object System.Drawing.Size(100, 25)
            $closeDetailsButton.Text = "Close"
            $closeDetailsButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom
            $closeDetailsButton.add_Click({ $detailsForm.Close() })
            $detailsForm.Controls.Add($closeDetailsButton)
            
            [void]$detailsForm.ShowDialog($connectorsForm)
        } else {
            [System.Windows.Forms.MessageBox]::Show("Please select a connector to view details.", "No Selection", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
    })
    $connectorsForm.Controls.Add($viewConnectorDetailsButton)

    # Export Connectors Button
    $exportConnectorsButton = New-Object System.Windows.Forms.Button
    $exportConnectorsButton.Location = New-Object System.Drawing.Point(305, 15)
    $exportConnectorsButton.Size = New-Object System.Drawing.Size(120, 25)
    $exportConnectorsButton.Text = "Export to CSV"
    $exportConnectorsButton.add_Click({
        if ($inboundListView.Items.Count -eq 0 -and $outboundListView.Items.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("No connectors loaded. Click 'Load All Connectors' first.", "No Data", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            return
        }
        
        $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveFileDialog.Filter = "CSV files (*.csv)|*.csv"
        $saveFileDialog.DefaultExt = "csv"
        $saveFileDialog.FileName = "ExchangeConnectors_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        
        if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            try {
                $exportData = @()
                
                # Export Inbound Connectors
                foreach ($item in $inboundListView.Items) {
                    $connector = $item.Tag
                    $exportData += [PSCustomObject]@{
                        ConnectorType = "Inbound"
                        Name = $connector.Name
                        Enabled = $connector.Enabled
                        ConnectorSource = $connector.ConnectorSource 
                        SenderDomains = ($connector.SenderDomains -join "; ")
                        SenderIPAddresses = ($connector.SenderIPAddresses -join "; ")
                        TlsSenderCertificateName = if($connector.TlsSenderCertificateName){$connector.TlsSenderCertificateName}else{"N/A"}
                        RequireTls = $connector.RequireTls
                        Identity = $connector.Identity
                        Description = if($connector.Comment){$connector.Comment}else{"N/A"}
                    }
                }
                
                # Export Outbound Connectors
                foreach ($item in $outboundListView.Items) {
                    $connector = $item.Tag
                    $exportData += [PSCustomObject]@{
                        ConnectorType = "Outbound"
                        Name = $connector.Name
                        Enabled = $connector.Enabled
                        ConnectorSource = $connector.ConnectorSource 
                        RecipientDomains = ($connector.RecipientDomains -join "; ")
                        SmartHosts = ($connector.SmartHosts -join "; ")
                        TlsDomain = $connector.TlsDomain 
                        RequireTls = $connector.RequireTls
                        Identity = $connector.Identity
                        Description = if($connector.Comment){$connector.Comment}else{"N/A"}
                    }
                }
                
                $exportData | Export-Csv -Path $saveFileDialog.FileName -NoTypeInformation -Encoding UTF8
                [System.Windows.Forms.MessageBox]::Show("Connectors exported to:`n$($saveFileDialog.FileName)", "Export Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                $statusLabel.Text = "Connectors exported to $($saveFileDialog.FileName)"
            } catch {
                $ex = $_.Exception
                [System.Windows.Forms.MessageBox]::Show("Error exporting connectors:`n$($ex.Message)", "Export Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }
    })
    $connectorsForm.Controls.Add($exportConnectorsButton)

    # Close Button
    $closeConnectorsButton = New-Object System.Windows.Forms.Button
    $closeConnectorsButton.Location = New-Object System.Drawing.Point(850, 15)
    $closeConnectorsButton.Size = New-Object System.Drawing.Size(100, 25)
    $closeConnectorsButton.Text = "Close"
    $closeConnectorsButton.Anchor = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right)
    $closeConnectorsButton.add_Click({ $connectorsForm.Close() })
    $connectorsForm.Controls.Add($closeConnectorsButton)

    [void]$connectorsForm.ShowDialog($mainForm)
}

Function Show-SessionRevocationTool {
    param($mainForm, $statusLabel, $allLoadedMailboxUPNs)
    
    $sessionForm = New-Object System.Windows.Forms.Form
    $sessionForm.Text = "User Session Revocation Tool"
    $sessionForm.Size = New-Object System.Drawing.Size(700, 500)
    $sessionForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
    $sessionForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $sessionForm.MaximizeBox = $false

    # Instructions Label
    $instructionsLabel = New-Object System.Windows.Forms.Label
    $instructionsLabel.Location = New-Object System.Drawing.Point(15, 15)
    $instructionsLabel.Size = New-Object System.Drawing.Size(650, 60)
    $instructionsLabel.Text = "Select users to revoke all active sessions. This will force them to re-authenticate.`nRequires Azure AD administrative privileges (e.g., User.ReadWrite.All scope for MS Graph).`n`nNote: Users that exist only in Exchange Online (not Azure AD) will show as 'not found'."
    $instructionsLabel.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Regular)
    $sessionForm.Controls.Add($instructionsLabel)

    # User Selection Controls
    $userSelectionLabel = New-Object System.Windows.Forms.Label
    $userSelectionLabel.Location = New-Object System.Drawing.Point(15, 85)
    $userSelectionLabel.Size = New-Object System.Drawing.Size(200, 20)
    $userSelectionLabel.Text = "Select Users for Session Revocation:"
    $sessionForm.Controls.Add($userSelectionLabel)

    $selectAllSessionsButton = New-Object System.Windows.Forms.Button
    $selectAllSessionsButton.Location = New-Object System.Drawing.Point(220, 83)
    $selectAllSessionsButton.Size = New-Object System.Drawing.Size(100, 23)
    $selectAllSessionsButton.Text = "Select All"
    $sessionForm.Controls.Add($selectAllSessionsButton)

    $deselectAllSessionsButton = New-Object System.Windows.Forms.Button
    $deselectAllSessionsButton.Location = New-Object System.Drawing.Point(330, 83)
    $deselectAllSessionsButton.Size = New-Object System.Drawing.Size(100, 23)
    $deselectAllSessionsButton.Text = "Deselect All"
    $sessionForm.Controls.Add($deselectAllSessionsButton)

    # User List
    $userSessionCheckedListBox = New-Object System.Windows.Forms.CheckedListBox
    $userSessionCheckedListBox.Location = New-Object System.Drawing.Point(15, 110)
    $userSessionCheckedListBox.Size = New-Object System.Drawing.Size(650, 230)
    $userSessionCheckedListBox.CheckOnClick = $true
    $sessionForm.Controls.Add($userSessionCheckedListBox)

    # Load users from the main form's loaded mailboxes
    if ($allLoadedMailboxUPNs -and $allLoadedMailboxUPNs.Count -gt 0) {
        foreach ($upn in $allLoadedMailboxUPNs | Sort-Object) {
            [void]$userSessionCheckedListBox.Items.Add($upn, $false)
        }
    }

    # Select/Deselect All button events
    $selectAllSessionsButton.add_Click({
        for ($i = 0; $i -lt $userSessionCheckedListBox.Items.Count; $i++) {
            $userSessionCheckedListBox.SetItemChecked($i, $true)
        }
    })

    $deselectAllSessionsButton.add_Click({
        for ($i = 0; $i -lt $userSessionCheckedListBox.Items.Count; $i++) {
            $userSessionCheckedListBox.SetItemChecked($i, $false)
        }
    })

    # Progress Bar
    $sessionProgressBar = New-Object System.Windows.Forms.ProgressBar
    $sessionProgressBar.Location = New-Object System.Drawing.Point(15, 350)
    $sessionProgressBar.Size = New-Object System.Drawing.Size(650, 20)
    $sessionProgressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
    $sessionForm.Controls.Add($sessionProgressBar)

    # Revoke Sessions Button
    $revokeSessionsButton = New-Object System.Windows.Forms.Button
    $revokeSessionsButton.Location = New-Object System.Drawing.Point(15, 380)
    $revokeSessionsButton.Size = New-Object System.Drawing.Size(200, 35)
    $revokeSessionsButton.Text = "Revoke Selected Sessions"
    $revokeSessionsButton.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 10, [System.Drawing.FontStyle]::Bold)
    $revokeSessionsButton.ForeColor = [System.Drawing.Color]::Red
    $revokeSessionsButton.add_Click({
        $selectedUsers = $userSessionCheckedListBox.CheckedItems | ForEach-Object { $_ }
        if ($selectedUsers.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Please select at least one user.", "No Users Selected", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        
        if (-not $script:graphConnection) {
             [System.Windows.Forms.MessageBox]::Show("Not connected to Microsoft Graph. Please connect from the main window first.", "MS Graph Not Connected", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning); return
        }

        $confirmResult = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to revoke all active sessions for the following $($selectedUsers.Count) user(s)?`n`nThis will force them to re-authenticate.`n`n$($selectedUsers -join "`n")", "Confirm Session Revocation", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
        
        if ($confirmResult -eq [System.Windows.Forms.DialogResult]::Yes) {
            $statusLabel.Text = "Revoking user sessions..."
            $sessionForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            $revokeSessionsButton.Enabled = $false
            $sessionProgressBar.Value = 0
            $sessionProgressBar.Maximum = $selectedUsers.Count
            
            $successCount = 0
            $errorCount = 0
            $errors = @()
            
            foreach ($user in $selectedUsers) {
                try {
                    Write-Host "Attempting to revoke sessions for user: $user via MS Graph" -ForegroundColor Yellow
                    
                    # First, verify the user exists in Azure AD
                    $mgUser = Get-MgUser -UserId $user -ErrorAction Stop
                    if (-not $mgUser) {
                        throw "User not found in Azure AD"
                    }
                    
                    # Now revoke sessions
                    Revoke-MgUserSignInSession -UserId $user -ErrorAction Stop
                    Write-Host "Successfully revoked sessions for $user using Microsoft Graph" -ForegroundColor Green
                    $successCount++
                } catch {
                    $ex = $_.Exception
                    $errorMessage = $ex.Message
                    
                    # Handle specific error cases
                    if ($errorMessage -like "*ResourceNotFound*" -or $errorMessage -like "*does not exist*" -or $errorMessage -like "*User not found*") {
                        Write-Warning "User $user not found in Azure AD or not synchronized. This is common for on-premises users or accounts that don't exist in Azure AD."
                        $errors += "User $user not found in Azure AD (may be on-premises only or not synchronized)"
                    } elseif ($errorMessage -like "*Forbidden*" -or $errorMessage -like "*Access*denied*") {
                        Write-Warning "Access denied for $user. May require additional permissions or user may be protected."
                        $errors += "Access denied for $user (insufficient permissions or protected account)"
                    } else {
                        Write-Warning "Session revocation failed for $user using MS Graph. Error: $($ex.Message)"
                        $errors += "Failed to revoke sessions for $user : $($ex.Message)"
                    }
                    $errorCount++
                }
                
                $sessionProgressBar.Value++
                $sessionForm.Refresh()
            }
            
            $sessionForm.Cursor = [System.Windows.Forms.Cursors]::Default
            $revokeSessionsButton.Enabled = $true
            $sessionProgressBar.Value = 0
            
            $resultMessage = "Session Revocation Complete:`n`nSuccessful: $successCount`nFailed: $errorCount"
            if ($errors.Count -gt 0) {
                $resultMessage += "`n`nCommon reasons for failures:"
                $resultMessage += "`n• User exists only in Exchange Online (not Azure AD)"
                $resultMessage += "`n• User is on-premises and not synchronized to Azure AD"
                $resultMessage += "`n• Insufficient permissions to revoke sessions for that user"
                $resultMessage += "`n• User account is a special/protected account`n"
                $resultMessage += "`nDetailed Errors:`n" + ($errors -join "`n")
            }
            
            if ($errorCount -eq 0) {
                [System.Windows.Forms.MessageBox]::Show($resultMessage, "Revocation Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                $statusLabel.Text = "Session revocation completed successfully for $successCount users."
            } else {
                [System.Windows.Forms.MessageBox]::Show($resultMessage, "Revocation Complete with Expected Failures", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                $statusLabel.Text = "Session revocation completed. $successCount successful, $errorCount failed (see details)."
            }
        }
    })
    $sessionForm.Controls.Add($revokeSessionsButton)

    # Module Requirements Button (Simplified as Graph connection is now primary)
    $moduleReqButton = New-Object System.Windows.Forms.Button
    $moduleReqButton.Location = New-Object System.Drawing.Point(230, 380)
    $moduleReqButton.Size = New-Object System.Drawing.Size(150, 35)
    $moduleReqButton.Text = "Check MS Graph Modules"
    $moduleReqButton.add_Click({
        $requirements = @("Required Microsoft Graph PowerShell Modules for Session Revocation:", "")
        $moduleOK = $true
        foreach($modInfo in $script:requiredGraphModules){
            if(Get-Module -ListAvailable -Name $modInfo.Name){
                $requirements += "✓ $($modInfo.Name): INSTALLED"
                 if($modInfo.Name -eq "Microsoft.Graph.Users"){ # Specifically check for Revoke-MgUserSignInSession
                    if(Get-Command "Revoke-MgUserSignInSession" -ErrorAction SilentlyContinue){
                         $requirements += "  - Revoke-MgUserSignInSession: Available"
                    } else {
                         $requirements += "  - Revoke-MgUserSignInSession: NOT Available (Command missing)"
                         $moduleOK = $false
                    }
                 }
            } else {
                $requirements += "✗ $($modInfo.Name): NOT INSTALLED"
                $requirements += "  Install with: Install-Module $($modInfo.Name) -Scope CurrentUser"
                $moduleOK = $false
            }
        }
        $requirements += ""
        if($moduleOK){
            $requirements += "All primary Graph modules seem to be available."
        } else {
            $requirements += "One or more Graph modules/commands might be missing."
        }
        $requirements += "Ensure you are connected to MS Graph via the main window."
        [System.Windows.Forms.MessageBox]::Show(($requirements -join "`n"), "MS Graph Module Requirements", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    })
    $sessionForm.Controls.Add($moduleReqButton)

    # Close Button
    $closeSessionButton = New-Object System.Windows.Forms.Button
    $closeSessionButton.Location = New-Object System.Drawing.Point(550, 380)
    $closeSessionButton.Size = New-Object System.Drawing.Size(100, 35)
    $closeSessionButton.Text = "Close"
    $closeSessionButton.add_Click({ $sessionForm.Close() })
    $sessionForm.Controls.Add($closeSessionButton)

    [void]$sessionForm.ShowDialog($mainForm)
}

Function Set-UserSignInBlockedState {
    param(
        [Parameter(Mandatory=$true)]
        [array]$UserPrincipalNames,
        [Parameter(Mandatory=$true)]
        [bool]$Blocked,
        [Parameter(Mandatory=$false)]
        [System.Windows.Forms.ToolStripStatusLabel]$StatusLabel,
        [Parameter(Mandatory=$false)]
        [System.Windows.Forms.ProgressBar]$ProgressBar,
        [Parameter(Mandatory=$false)]
        [System.Windows.Forms.Form]$MainForm
    )

    if (-not $script:graphConnection) {
        [System.Windows.Forms.MessageBox]::Show("Not connected to Microsoft Graph. Please connect first.", "MS Graph Not Connected", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }

    $action = if ($Blocked) { "block sign-in for" } else { "unblock sign-in for" }
    if ($StatusLabel) { $StatusLabel.Text = "Attempting to $action selected users..." }
    if ($MainForm) { $MainForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor }
    if ($ProgressBar) { 
        $ProgressBar.Maximum = $UserPrincipalNames.Count
        $ProgressBar.Value = 0
        $ProgressBar.Step = 1 
    }

    $successCount = 0
    $errorCount = 0
    $errorMessages = @()

    foreach ($upn in $UserPrincipalNames) {
        Write-Host "Attempting to $action $upn..."
        try {
            # First verify the user exists in Azure AD
            $mgUser = Get-MgUser -UserId $upn -Property Id,AccountEnabled -ErrorAction Stop
            if (-not $mgUser) {
                throw "User not found in Azure AD"
            }
            
            # Fixed: Use -AccountEnabled directly without casting
            $accountEnabled = -not $Blocked
            Update-MgUser -UserId $upn -AccountEnabled:$accountEnabled -ErrorAction Stop 
            Write-Host "Successfully $($action.Split(' ')[0] + 'ed') sign-in for $upn." -ForegroundColor Green
            $successCount++
        } catch {
            $ex = $_.Exception
            $errorMessage = $ex.Message
            
            # Handle specific error cases
            if ($errorMessage -like "*ResourceNotFound*" -or $errorMessage -like "*does not exist*" -or $errorMessage -like "*User not found*") {
                $errMsg = "User $upn not found in Azure AD (may be on-premises only or not synchronized)"
                Write-Warning $errMsg
            } elseif ($errorMessage -like "*Forbidden*" -or $errorMessage -like "*Access*denied*") {
                $errMsg = "Access denied for $upn (insufficient permissions or protected account)"
                Write-Warning $errMsg
            } else {
                $errMsg = "Failed to $action ${upn}: $($ex.Message)"
                Write-Warning $errMsg
            }
            
            $errorMessages += $errMsg
            $errorCount++
        }
        if ($ProgressBar) { $ProgressBar.PerformStep() }
    }

    if ($StatusLabel) { $StatusLabel.Text = "Sign-in status update complete. Success: $successCount, Failed: $errorCount." }
    if ($MainForm) { $MainForm.Cursor = [System.Windows.Forms.Cursors]::Default }
    if ($ProgressBar) { $ProgressBar.Value = 0 }

    $summaryMessage = "Sign-in Update Summary:`n`nSuccessfully $($action.Split(' ')[0] + 'ed'): $successCount user(s).`nFailed: $errorCount user(s)."
    if ($errorMessages.Count -gt 0) {
        $summaryMessage += "`n`nCommon reasons for failures:"
        $summaryMessage += "`n• User exists only in Exchange Online (not Azure AD)"
        $summaryMessage += "`n• User is on-premises and not synchronized to Azure AD"
        $summaryMessage += "`n• Insufficient permissions to modify that user"
        $summaryMessage += "`n• User account is a special/protected account`n"
        $summaryMessage += "`nDetailed Errors:`n$($errorMessages -join "`n")"
    }
    
    $iconType = if($errorCount -gt 0) { [System.Windows.Forms.MessageBoxIcon]::Information } else { [System.Windows.Forms.MessageBoxIcon]::Information }
    $titleText = if($errorCount -gt 0) { "Operation Complete with Expected Failures" } else { "Operation Complete" }
    
    [System.Windows.Forms.MessageBox]::Show($summaryMessage, $titleText, [System.Windows.Forms.MessageBoxButtons]::OK, $iconType)
}

Function Show-RestrictedSenderManagementDialog {
    param(
        [Parameter(Mandatory=$true)]
        [string]$UserPrincipalName,
        [Parameter(Mandatory=$true)]
        [System.Windows.Forms.Form]$ParentForm,
        [Parameter(Mandatory=$true)]
        [System.Windows.Forms.ToolStripStatusLabel]$StatusLabelGlobal 
    )

    if (-not $script:graphConnection) {
        [System.Windows.Forms.MessageBox]::Show("Not connected to Microsoft Graph. Please connect first.", "MS Graph Not Connected", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    # Try to import the module but don't fail if it's not available
    try {
        Import-Module Microsoft.Graph.Identity.SignIns -ErrorAction SilentlyContinue -Force
        Write-Host "Attempted to import Microsoft.Graph.Identity.SignIns module for RestrictedSenderDialog."
    } catch {
        Write-Warning "Could not import Microsoft.Graph.Identity.SignIns module. Some features may not be available."
    }

    # Check if the required cmdlets are available - using alternative approaches
    $hasRestrictedUserCmdlets = $false
    
    # Check for various possible cmdlets that might help with restricted users
    $possibleCmdlets = @("Get-MgRiskyUser", "Get-MgUserSignInActivity", "Get-MgUser")
    foreach ($cmdlet in $possibleCmdlets) {
        if (Get-Command $cmdlet -ErrorAction SilentlyContinue) {
            $hasRestrictedUserCmdlets = $true
            break
        }
    }

    if (-not $hasRestrictedUserCmdlets) {
        [System.Windows.Forms.MessageBox]::Show("The required Microsoft Graph cmdlets for managing restricted users are not available.`n`nThis feature requires:`n- Microsoft.Graph.Identity.SignIns module`n- Appropriate permissions in Azure AD`n`nYou may need to install additional modules or use alternative methods.", "Feature Not Available", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    $dialog = New-Object System.Windows.Forms.Form
    $dialog.Text = "Manage Sending Restrictions: $UserPrincipalName"
    $dialog.Size = New-Object System.Drawing.Size(500, 350)
    $dialog.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
    $dialog.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $dialog.MaximizeBox = $false
    $dialog.MinimizeBox = $false

    $statusLabelLocal = New-Object System.Windows.Forms.Label
    $statusLabelLocal.Location = New-Object System.Drawing.Point(15, 20)
    $statusLabelLocal.Size = New-Object System.Drawing.Size(450, 80) 
    $statusLabelLocal.Text = "Loading user management options for $UserPrincipalName...`n`nNote: For comprehensive sending restriction management, use Exchange Online PowerShell cmdlets or the Exchange Admin Center.`n`nClick 'Check User Info' to verify if this user exists in Azure AD."
    $dialog.Controls.Add($statusLabelLocal)

    $infoButton = New-Object System.Windows.Forms.Button
    $infoButton.Location = New-Object System.Drawing.Point(15, 110)
    $infoButton.Size = New-Object System.Drawing.Size(200, 30)
    $infoButton.Text = "Check User Info"
    $dialog.Controls.Add($infoButton)

    $exchangeInfoButton = New-Object System.Windows.Forms.Button
    $exchangeInfoButton.Location = New-Object System.Drawing.Point(230, 110)
    $exchangeInfoButton.Size = New-Object System.Drawing.Size(200, 30)
    $exchangeInfoButton.Text = "Check Exchange Restrictions"
    $dialog.Controls.Add($exchangeInfoButton)

    $closeButton = New-Object System.Windows.Forms.Button
    $closeButton.Location = New-Object System.Drawing.Point(370, 270) 
    $closeButton.Size = New-Object System.Drawing.Size(100, 30)
    $closeButton.Text = "Close"
    $closeButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel 
    $dialog.CancelButton = $closeButton 
    $dialog.Controls.Add($closeButton)

    $infoButton.add_Click({
        $dialog.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        $statusLabelLocal.Text = "Checking user information for $UserPrincipalName..."
        try {
            $mgUser = Get-MgUser -UserId $UserPrincipalName -Property Id,AccountEnabled,UserPrincipalName,DisplayName -ErrorAction Stop
            if ($mgUser) {
                $userInfo = "Azure AD User Information:`n"
                $userInfo += "Display Name: $($mgUser.DisplayName)`n"
                $userInfo += "UPN: $($mgUser.UserPrincipalName)`n"
                $userInfo += "Account Enabled: $($mgUser.AccountEnabled)`n"
                $userInfo += "Object ID: $($mgUser.Id)`n`n"
                $userInfo += "✓ This user exists in Azure AD and can be managed via Microsoft Graph."
                $statusLabelLocal.Text = $userInfo
            }
        } catch {
            $ex = $_.Exception
            if ($ex.Message -like "*ResourceNotFound*" -or $ex.Message -like "*does not exist*") {
                $userAnalysis = "User Analysis for $UserPrincipalName" + ":`n`n⚠ This user was NOT FOUND in Azure AD.`n`nThis typically means:`n• User exists only in Exchange Online`n• User is an on-premises account not synchronized to Azure AD`n• User is a mail-enabled contact or external user`n`nFor Azure AD management (sign-in blocking, session revocation), this user cannot be managed via Microsoft Graph.`n`nFor Exchange Online sending restrictions, use Exchange Online PowerShell cmdlets or the Exchange Admin Center."
                $statusLabelLocal.Text = $userAnalysis
            } else {
                $errorAnalysis = "Error checking user information:`n$($ex.Message)`n`nThis could indicate:`n• Insufficient permissions to read user details`n• Temporary connectivity issues`n• User exists but some properties are restricted"
                $statusLabelLocal.Text = $errorAnalysis
            }
        } finally {
            $dialog.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    })

    $exchangeInfoButton.add_Click({
        $dialog.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        $statusLabelLocal.Text = "Checking Exchange Online restrictions for $UserPrincipalName..."
        try {
            # Use the Exchange Online cmdlets to check for sending restrictions
            $restrictions = Get-ExchangeOnlineSendingRestrictions -UserPrincipalName $UserPrincipalName
            if ($restrictions) {
                $restrictionInfo = "Exchange Online Sending Restrictions:`n"
                $restrictionInfo += "Require Sender Auth: $($restrictions.RequireSenderAuthenticationEnabled)`n"
                $restrictionInfo += "Accept Messages Only From: $($restrictions.AcceptMessagesOnlyFrom -join '; ')`n"
                $restrictionInfo += "Reject Messages From: $($restrictions.RejectMessagesFrom -join '; ')`n"
                $restrictionInfo += "`nFor full management, use Exchange Admin Center or Exchange Online PowerShell."
                $statusLabelLocal.Text = $restrictionInfo
            } else {
                $statusLabelLocal.Text = "Could not retrieve Exchange Online restrictions. User may not exist or permissions may be insufficient."
            }
        } catch {
            $ex = $_.Exception
            $statusLabelLocal.Text = "Error checking Exchange restrictions: $($ex.Message)"
        } finally {
            $dialog.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    })

    [void]$dialog.ShowDialog($ParentForm)
    $dialog.Dispose()
}

# Alternative function for checking Exchange Online sending restrictions
Function Get-ExchangeOnlineSendingRestrictions {
    param(
        [Parameter(Mandatory=$true)]
        [string]$UserPrincipalName
    )
    
    try {
        # Check if user has any send restrictions at the mailbox level
        $mailbox = Get-Mailbox -Identity $UserPrincipalName -ErrorAction Stop
        $restrictions = @{
            RequireSenderAuthenticationEnabled = $mailbox.RequireSenderAuthenticationEnabled
            AcceptMessagesOnlyFrom = $mailbox.AcceptMessagesOnlyFrom
            AcceptMessagesOnlyFromDLMembers = $mailbox.AcceptMessagesOnlyFromDLMembers
            RejectMessagesFrom = $mailbox.RejectMessagesFrom
            RejectMessagesFromDLMembers = $mailbox.RejectMessagesFromDLMembers
        }
        
        # Check organization-level restrictions
        try {
            $orgConfig = Get-OrganizationConfig
            $restrictions.OutboundSpamFilteringEnabled = $orgConfig.OutboundSpamFilteringEnabled
        } catch {
            Write-Warning "Could not retrieve organization configuration"
        }
        
        return $restrictions
    } catch {
        Write-Error "Could not retrieve sending restrictions for $UserPrincipalName : $($_.Exception.Message)"
        return $null
    }
}

Function Get-AutoDetectedDomains {
    param(
        [Parameter(Mandatory=$true)]
        [array]$MailboxUPNs,
        [Parameter(Mandatory=$false)]
        [int]$MaxSampleSize = 100
    )
    
    Write-Host "Auto-detecting organization domains from loaded mailboxes..." -ForegroundColor Cyan
    
    if (-not $MailboxUPNs -or $MailboxUPNs.Count -eq 0) {
        Write-Warning "No mailbox UPNs provided for domain detection"
        return @()
    }
    
    # Sample the mailboxes if there are too many
    $sampleSize = [Math]::Min($MaxSampleSize, $MailboxUPNs.Count)
    $samplesToAnalyze = if ($MailboxUPNs.Count -gt $MaxSampleSize) {
        # Take a representative sample
        $step = [Math]::Floor($MailboxUPNs.Count / $MaxSampleSize)
        $samples = @()
        for ($i = 0; $i -lt $MailboxUPNs.Count; $i += $step) {
            $samples += $MailboxUPNs[$i]
            if ($samples.Count -ge $MaxSampleSize) { break }
        }
        $samples
    } else {
        $MailboxUPNs
    }
    
    Write-Host "Analyzing $($samplesToAnalyze.Count) mailbox UPNs for domain patterns..." -ForegroundColor Yellow
    
    # Extract domains and count occurrences
    $domainCounts = @{}
    $onMicrosoftDomains = @{}
    
    foreach ($upn in $samplesToAnalyze) {
        if ($upn -like "*@*") {
            try {
                $domain = ($upn.Split('@')[1]).ToLowerInvariant()
                if (-not [string]::IsNullOrWhiteSpace($domain)) {
                    if ($domain -like "*.onmicrosoft.com") {
                        if (-not $onMicrosoftDomains.ContainsKey($domain)) {
                            $onMicrosoftDomains[$domain] = 0
                        }
                        $onMicrosoftDomains[$domain]++
                    } else {
                        if (-not $domainCounts.ContainsKey($domain)) {
                            $domainCounts[$domain] = 0
                        }
                        $domainCounts[$domain]++
                    }
                }
            } catch {
                Write-Warning "Could not parse domain from UPN: $upn"
            }
        }
    }
    
    # Prioritize non-onmicrosoft.com domains
    $detectedDomains = @()
    
    # First, add non-onmicrosoft.com domains sorted by frequency
    if ($domainCounts.Count -gt 0) {
        $sortedDomains = $domainCounts.GetEnumerator() | Sort-Object Value -Descending
        foreach ($domainEntry in $sortedDomains) {
            $detectedDomains += $domainEntry.Key
            Write-Host "  Found primary domain: $($domainEntry.Key) (used by $($domainEntry.Value) mailboxes)" -ForegroundColor Green
        }
    }
    
    # If no primary domains found, add onmicrosoft.com domains
    if ($detectedDomains.Count -eq 0 -and $onMicrosoftDomains.Count -gt 0) {
        $sortedOnMicrosoftDomains = $onMicrosoftDomains.GetEnumerator() | Sort-Object Value -Descending
        foreach ($domainEntry in $sortedOnMicrosoftDomains) {
            $detectedDomains += $domainEntry.Key
            Write-Host "  Found onmicrosoft.com domain: $($domainEntry.Key) (used by $($domainEntry.Value) mailboxes)" -ForegroundColor Yellow
        }
    }
    
    # Limit to top 5 domains to avoid overwhelming the UI
    if ($detectedDomains.Count -gt 5) {
        Write-Host "  Limiting to top 5 most common domains..." -ForegroundColor Cyan
        $detectedDomains = $detectedDomains[0..4]
    }
    
    if ($detectedDomains.Count -gt 0) {
        Write-Host "Auto-detected domains: $($detectedDomains -join ', ')" -ForegroundColor Green
    } else {
        Write-Warning "No domains could be auto-detected from mailbox UPNs"
    }
    
    return $detectedDomains
}


# --- Check Prerequisites ---
if (-not (Test-ExchangeModule)) {
    $choice = [System.Windows.Forms.MessageBox]::Show("ExchangeOnlineManagement module is missing.`n`nInstall now?", "Missing Module", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
    if ($choice -eq [System.Windows.Forms.DialogResult]::Yes) {
        if (-not (Install-ExchangeModule)) { [System.Windows.Forms.MessageBox]::Show("Install failed. Do it manually & restart.", "Install Failed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error); Exit }
    } else { [System.Windows.Forms.MessageBox]::Show("Script needs ExchangeOnlineManagement module.", "Prerequisites Missing", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error); Exit }
    [System.Windows.Forms.MessageBox]::Show("Restart script after module install.", "Restart Required", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information); Exit
}


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
$transportRulesButton.add_Click({ Show-TransportRulesViewer -mainForm $mainForm -statusLabel $statusLabel })
$mainForm.Controls.Add($transportRulesButton)

$connectorsButton = New-Object System.Windows.Forms.Button; $connectorsButton.Location = New-Object System.Drawing.Point(540, 20); $connectorsButton.Size = New-Object System.Drawing.Size(120, 30); $connectorsButton.Text = "Connectors"; $connectorsButton.Enabled = $false; $connectorsButton.Anchor = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right)
$connectorsButton.add_Click({ Show-ConnectorsViewer -mainForm $mainForm -statusLabel $statusLabel })
$mainForm.Controls.Add($connectorsButton)

$sessionRevocationButton = New-Object System.Windows.Forms.Button; $sessionRevocationButton.Location = New-Object System.Drawing.Point(20, 90); $sessionRevocationButton.Size = New-Object System.Drawing.Size(260, 30); $sessionRevocationButton.Text = "Revoke User Sessions (Graph)"; $sessionRevocationButton.Enabled = $false; $sessionRevocationButton.Anchor = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left)
$sessionRevocationButton.add_Click({ Show-SessionRevocationTool -mainForm $mainForm -statusLabel $statusLabel -allLoadedMailboxUPNs $script:allLoadedMailboxUPNs })
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
                        $isHiddenValue = $false; if ($rule.PSObject.Properties.Match('IsHidden').Count -gt 0) { $isHiddenValue = $rule.IsHidden } else { if ($rule.Name -like 'RuleId:*' -or ($rule.Description -match 'system-generated' -or $rule.Description -match 'Generated by Microsoft Exchange')) { $isHiddenValue = $true } }
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
                    $statusLabel.Text = "Exported & formatted $($allRuleAnalysisData.Count) rules to $xlsxFilePath"; [System.Windows.Forms.MessageBox]::Show("Exported & formatted $($allRuleAnalysisData.Count) rules to:`n$xlsxFilePath", "XLSX Export OK", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                    try { Remove-Item -Path $csvFilePath -Force -EA SilentlyContinue } catch {}
                } else { $statusLabel.Text = "CSV OK, XLSX/Format Failed."; [System.Windows.Forms.MessageBox]::Show("CSV Exported to:`n$csvFilePath`n`nXLSX/Format FAILED. Check Excel install & console.", "XLSX Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning); $errorOccurred = $true }
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

    $deleteRulesButton = New-Object System.Windows.Forms.Button; $deleteRulesButton.Location = New-Object System.Drawing.Point(145, 330); $deleteRulesButton.Size = New-Object System.Drawing.Size(150, 30); $deleteRulesButton.Text = "Delete Selected Rules"; $deleteRulesButton.ForegroundColor = [System.Drawing.Color]::Red; $deleteRulesButton.Anchor = ([System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left)
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