<#
.SYNOPSIS
Microsoft 365 Management Tool - Exchange Online and Entra ID Analysis

.DESCRIPTION
Comprehensive PowerShell GUI tool for analyzing Exchange Online inbox rules, managing user accounts,
monitoring security configurations, and investigating Entra ID accounts.

Features:
- Exchange Online inbox rules analysis and management
- Entra ID user management and security analysis
- Microsoft Graph integration for user operations
- Transport rules and connectors management
- Sign-in logs and audit analysis
- XLSX report generation with advanced formatting

.NOTES
Version: 8.0
Requires: PowerShell 5.1+, ExchangeOnlineManagement, Microsoft.Graph modules, Microsoft Excel
Permissions: Exchange administrative privileges and Microsoft Graph permissions

.LINK
Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force
Install-Module Microsoft.Graph.Users -Scope CurrentUser -Force
Install-Module Microsoft.Graph.Users.Actions -Scope CurrentUser -Force
Install-Module Microsoft.Graph.Identity.SignIns -Scope CurrentUser -Force
Install-Module Microsoft.Graph.Reports -Scope CurrentUser -Force
#>

# Load Windows Forms assembly
Add-Type -AssemblyName System.Windows.Forms

# Function to create tooltips
function Add-ToolTip {
    param(
        [Parameter(Mandatory=$true)]
        [System.Windows.Forms.Control]$Control,
        [Parameter(Mandatory=$true)]
        [string]$Text
    )
    $tooltip = New-Object System.Windows.Forms.ToolTip
    $tooltip.AutoPopDelay = 5000
    $tooltip.InitialDelay = 1000
    $tooltip.ReshowDelay = 500
    $tooltip.ShowAlways = $true
    $tooltip.SetToolTip($Control, $Text)
}

# Import all modules with error handling
function Safe-ImportModule($modulePath) {
    try {
        # Get the module name from the path
        $moduleName = [System.IO.Path]::GetFileNameWithoutExtension($modulePath)
        
        # Remove the module if it's already loaded to force reload
        if (Get-Module -Name $moduleName -ErrorAction SilentlyContinue) {
            Remove-Module -Name $moduleName -Force -ErrorAction SilentlyContinue
        }
        
        Import-Module $modulePath -Global -ErrorAction Stop
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to import module: $modulePath`nError: $($_.Exception.Message)", "Module Import Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        exit
    }
}
Safe-ImportModule "$PSScriptRoot\Modules\ExchangeOnline.psm1"
Safe-ImportModule "$PSScriptRoot\Modules\GraphOnline.psm1"
Safe-ImportModule "$PSScriptRoot\Modules\MailboxAnalysis.psm1"
Safe-ImportModule "$PSScriptRoot\Modules\TransportRules.psm1"
Safe-ImportModule "$PSScriptRoot\Modules\Connectors.psm1"
Safe-ImportModule "$PSScriptRoot\Modules\SessionRevocation.psm1"
Safe-ImportModule "$PSScriptRoot\Modules\SignInManagement.psm1"
Safe-ImportModule "$PSScriptRoot\Modules\RestrictedSender.psm1"
Safe-ImportModule "$PSScriptRoot\Modules\ExportUtils.psm1"
Safe-ImportModule "$PSScriptRoot\Modules\EntraInvestigator.psm1"

# Function to show/hide progress bar
function Show-Progress {
    param($message, $progress = -1)
    $statusLabel.Text = $message
    if ($progress -ge 0) {
        $progressBar.Visible = $true
        $progressBar.Value = $progress
    } else {
        $progressBar.Visible = $false
    }
    [System.Windows.Forms.Application]::DoEvents()
}

# Function to load mailboxes with performance optimizations
function Load-MailboxesOptimized {
    param(
        [int]$MaxMailboxes = 1000,
        [switch]$LoadAll,
        [switch]$QuickLoad,
        [switch]$FullAnalysis
    )
    
    try {
        Show-Progress -message "Loading mailboxes..." -progress 10
        
        # Server-side filtering: Get mailboxes with enhanced filtering
        if ($LoadAll) {
            $mailboxes = Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox,SharedMailbox,RoomMailbox,EquipmentMailbox | 
                        Select-Object UserPrincipalName, DisplayName, AccountDisabled, IsLicensed, RecipientTypeDetails | 
                        Sort-Object UserPrincipalName
        } else {
            # Load only first batch for faster initial load
            $mailboxes = Get-Mailbox -ResultSize $MaxMailboxes -RecipientTypeDetails UserMailbox,SharedMailbox,RoomMailbox,EquipmentMailbox | 
                        Select-Object UserPrincipalName, DisplayName, AccountDisabled, IsLicensed, RecipientTypeDetails | 
                        Sort-Object UserPrincipalName
        }
        

        

        
        Show-Progress -message "Processing $($mailboxes.Count) mailboxes..." -progress 30
        
        # Batch get user details to reduce API calls
        $userPrincipalNames = $mailboxes | ForEach-Object { $_.UserPrincipalName }
        $userDetails = @{}
        
        # Process users in batches of 50 to avoid throttling
        $batchSize = 50
        for ($i = 0; $i -lt $userPrincipalNames.Count; $i += $batchSize) {
            $batch = $userPrincipalNames[$i..([Math]::Min($i + $batchSize - 1, $userPrincipalNames.Count - 1))]
            foreach ($upn in $batch) {
                try {
                    $user = Get-User -Identity $upn -ErrorAction SilentlyContinue
                    $userDetails[$upn] = $user
                } catch {
                    $userDetails[$upn] = $null
                }
            }
            Show-Progress -message "Processing users... ($([Math]::Min($i + $batchSize, $userPrincipalNames.Count))/$($userPrincipalNames.Count))" -progress (30 + ($i / $userPrincipalNames.Count * 20))
        }
        
        Show-Progress -message "Analyzing mailboxes..." -progress 60
        
        $userMailboxGrid.Rows.Clear()
        $script:allLoadedMailboxUPNs = @()
        $script:allLoadedMailboxes = $mailboxes  # Store the full mailbox objects for filtering
        $totalMailboxes = $mailboxes.Count
        $processedCount = 0
        
        # Smart loading strategy based on mailbox count and user preference
        $shouldAnalyzeRules = $true
        $shouldAnalyzePermissions = $true
        
        if ($QuickLoad -or $totalMailboxes -gt 200) {
            # For large tenants or quick load mode, skip detailed analysis initially
            $shouldAnalyzeRules = $false
            $shouldAnalyzePermissions = $false
            Show-Progress -message "Large tenant detected ($totalMailboxes mailboxes). Loading basic data only. Use 'Analyze Selected' for detailed analysis." -progress 65
        } elseif ($FullAnalysis) {
            # Force full analysis regardless of size
            $shouldAnalyzeRules = $true
            $shouldAnalyzePermissions = $true
            Show-Progress -message "Performing full analysis for all mailboxes..." -progress 65
        } else {
            # Default behavior for medium-sized tenants
            $shouldAnalyzeRules = $true
            $shouldAnalyzePermissions = $true
            Show-Progress -message "Performing standard analysis..." -progress 65
        }
        
        foreach ($mbx in $mailboxes) {
            $script:allLoadedMailboxUPNs += $mbx.UserPrincipalName
            

            
            # Use cached user details
            $user = $userDetails[$mbx.UserPrincipalName]
            if ($null -ne $user -and $null -ne $user.AccountDisabled) {
                $signInBlocked = if ($user.AccountDisabled) { "Blocked" } else { "Allowed" }
            } else {
                $signInBlocked = "Unknown"
            }
            
            # Initialize default values
            $rulesCount = "0"
            $hiddenRules = "0"
            $suspiciousRules = "0"
            $externalForwarding = "Unknown"
            $delegates = "Unknown"
            $fullAccess = "Unknown"
            
            # Set N/A for shared mailboxes (rules not applicable)
            if ($mbx.RecipientTypeDetails -eq "SharedMailbox") {
                $rulesCount = "N/A"
                $hiddenRules = "N/A"
                $suspiciousRules = "N/A"
                $externalForwarding = "N/A"
            }
            
            # Only analyze rules for user mailboxes (shared mailboxes don't have user-created inbox rules)
            if ($mbx.RecipientTypeDetails -eq "UserMailbox" -and $shouldAnalyzeRules) {
                try {
                    $rules = Get-InboxRule -Mailbox $mbx.UserPrincipalName -IncludeHidden -ErrorAction SilentlyContinue
                    if ($rules) {
                        $analysis = Analyze-MailboxRulesEnhanced -Rules $rules -BaseSuspiciousKeywords $BaseSuspiciousKeywords
                        $rulesCount = $analysis.TotalRules.ToString()
                        $hiddenRules = $analysis.SuspiciousHidden.ToString()
                        $suspiciousRules = $analysis.SuspiciousVisible.ToString()
                        $externalForwarding = if ($analysis.HasExternalForwarding) { "Yes" } else { "No" }
                    }
                } catch {
                    # Keep default values if analysis fails
                }
            } elseif ($mbx.RecipientTypeDetails -eq "SharedMailbox") {
                # Shared mailboxes can't have user-created inbox rules or external forwarding
                $rulesCount = "N/A"
                $hiddenRules = "N/A"
                $suspiciousRules = "N/A"
                $externalForwarding = "N/A"
            }
            
            # Only analyze permissions for user mailboxes
            if ($mbx.RecipientTypeDetails -eq "UserMailbox" -and $shouldAnalyzePermissions) {
                try {
                    $delegates = Analyze-MailboxDelegates -UserPrincipalName $mbx.UserPrincipalName
                    $fullAccess = Analyze-MailboxPermissions -UserPrincipalName $mbx.UserPrincipalName
                } catch {
                    $delegates = "Error"
                    $fullAccess = "Error"
                }
            }
            
            $rowIdx = $userMailboxGrid.Rows.Add()
            $userMailboxGrid.Rows[$rowIdx].Cells["Select"].Value = $false
            $userMailboxGrid.Rows[$rowIdx].Cells["UserPrincipalName"].Value = $mbx.UserPrincipalName
            $userMailboxGrid.Rows[$rowIdx].Cells["DisplayName"].Value = $mbx.DisplayName
            $userMailboxGrid.Rows[$rowIdx].Cells["SignInBlocked"].Value = $signInBlocked
            $userMailboxGrid.Rows[$rowIdx].Cells["RecipientType"].Value = $mbx.RecipientTypeDetails
            $userMailboxGrid.Rows[$rowIdx].Cells["TotalRules"].Value = $rulesCount
            $userMailboxGrid.Rows[$rowIdx].Cells["HiddenRules"].Value = $hiddenRules
            $userMailboxGrid.Rows[$rowIdx].Cells["SuspiciousRules"].Value = $suspiciousRules
            $userMailboxGrid.Rows[$rowIdx].Cells["ExternalForwarding"].Value = $externalForwarding
            $userMailboxGrid.Rows[$rowIdx].Cells["Delegates"].Value = $delegates
            $userMailboxGrid.Rows[$rowIdx].Cells["FullAccess"].Value = $fullAccess
            $processedCount++
            
            if ($processedCount % 20 -eq 0) {
                Show-Progress -message "Processing mailboxes... ($processedCount/$totalMailboxes)" -progress (60 + ($processedCount / $totalMailboxes * 30))
            }
        }
        
        Show-Progress -message "Finalizing..." -progress 90
        

        
        # Auto-detect tenant/org domains from loaded mailboxes
        $detectedDomains = Get-AutoDetectedDomains -MailboxUPNs $script:allLoadedMailboxUPNs
        if ($detectedDomains -and $detectedDomains.Count -gt 0) {
            $orgDomainsTextBox.Text = ($detectedDomains -join ", ")
        } else {
            $orgDomainsTextBox.Text = ""
        }
        
        # Populate suspicious keywords from $BaseSuspiciousKeywords
        $keywordsTextBox.Text = ($BaseSuspiciousKeywords -join ", ")
        
        # Enable/disable buttons
        $selectAllButton.Enabled = $true
        $deselectAllButton.Enabled = $true
        $disconnectButton.Enabled = $true
        $connectButton.Enabled = $false
        $loadAllMailboxesButton.Enabled = $true
        $searchMailboxesButton.Enabled = $true
        $manageRulesButton.Enabled = $true
        $analyzeSelectedButton.Enabled = $true
        $manageConnectorsButton.Enabled = $true
        $manageTransportRulesButton.Enabled = $true
        $blockUserButton.Enabled = $false
        $unblockUserButton.Enabled = $false
        
        # Debug: Log button state
        Write-Host "Analyze Selected Button Enabled: $($analyzeSelectedButton.Enabled)"
        Write-Host "Analyze Selected Button Visible: $($analyzeSelectedButton.Visible)"
        Write-Host "Analyze Selected Button Text: $($analyzeSelectedButton.Text)"
        
        Show-Progress -message "Ready. Connected to Exchange Online. Loaded $($mailboxes.Count) mailboxes." -progress -1
        

        
        return $mailboxes.Count
    } catch {
        throw $_
    }
}

# Note: Grid event handler removed due to timing issue

# Function to get mailboxes with server-side filtering
function Get-MailboxesWithFilters {
    param(
        [int]$MaxMailboxes = 1000,
        [switch]$LoadAll,
        [switch]$OnlyWithRules,
        [switch]$OnlyWithPermissions
    )
    
    try {
        $filter = @()
        
        # Build server-side filters where possible
        if ($OnlyWithRules) {
            # Note: Exchange Online doesn't have direct server-side filtering for rules
            # We'll use client-side filtering but optimize the process
            $filter += "RecipientTypeDetails -eq 'UserMailbox'"
        }
        
        if ($OnlyWithPermissions) {
            # Note: Exchange Online doesn't have direct server-side filtering for permissions
            # We'll use client-side filtering but optimize the process
            $filter += "RecipientTypeDetails -eq 'UserMailbox'"
        }
        
        $filterString = if ($filter.Count -gt 0) { $filter -join " -and " } else { $null }
        
        if ($LoadAll) {
            if ($filterString) {
                $mailboxes = Get-Mailbox -Filter $filterString -ResultSize Unlimited -RecipientTypeDetails UserMailbox,SharedMailbox,RoomMailbox,EquipmentMailbox | 
                            Select-Object UserPrincipalName, DisplayName, AccountDisabled, IsLicensed, RecipientTypeDetails | 
                            Sort-Object UserPrincipalName
            } else {
                $mailboxes = Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox,SharedMailbox,RoomMailbox,EquipmentMailbox | 
                            Select-Object UserPrincipalName, DisplayName, AccountDisabled, IsLicensed, RecipientTypeDetails | 
                            Sort-Object UserPrincipalName
            }
        } else {
            if ($filterString) {
                $mailboxes = Get-Mailbox -Filter $filterString -ResultSize $MaxMailboxes -RecipientTypeDetails UserMailbox,SharedMailbox,RoomMailbox,EquipmentMailbox | 
                            Select-Object UserPrincipalName, DisplayName, AccountDisabled, IsLicensed, RecipientTypeDetails | 
                            Sort-Object UserPrincipalName
            } else {
                $mailboxes = Get-Mailbox -ResultSize $MaxMailboxes -RecipientTypeDetails UserMailbox,SharedMailbox,RoomMailbox,EquipmentMailbox | 
                            Select-Object UserPrincipalName, DisplayName, AccountDisabled, IsLicensed, RecipientTypeDetails | 
                            Sort-Object UserPrincipalName
            }
        }
        
        return $mailboxes
    } catch {
        Write-Error "Failed to get mailboxes with filters: $_"
        return @()
    }
}

# Function to batch analyze mailbox rules and permissions with server-side optimization
function Analyze-MailboxBatch {
    param(
        [Parameter(Mandatory=$true)]
        [array]$Mailboxes,
        [Parameter(Mandatory=$true)]
        [array]$BaseSuspiciousKeywords,
        [int]$BatchSize = 50
    )
    
    $results = @{}
    
    # Process mailboxes in batches for better performance
    for ($i = 0; $i -lt $Mailboxes.Count; $i += $BatchSize) {
        $batch = $Mailboxes[$i..([Math]::Min($i + $BatchSize - 1, $Mailboxes.Count - 1))]
        
        foreach ($mbx in $batch) {
            if ($mbx.RecipientTypeDetails -eq "UserMailbox") {
                $upn = $mbx.UserPrincipalName
                $result = @{
                    RulesCount = "0"
                    HiddenRules = "0"
                    SuspiciousRules = "0"
                    ExternalForwarding = "No"
                    Delegates = "None"
                    FullAccess = "None"
                }
                
                # Check rules (only if likely to have them)
                try {
                    $rules = Get-InboxRule -Mailbox $upn -IncludeHidden -ErrorAction SilentlyContinue
                    if ($rules -and $rules.Count -gt 0) {
                        $analysis = Analyze-MailboxRulesEnhanced -Rules $rules -BaseSuspiciousKeywords $BaseSuspiciousKeywords
                        $result.RulesCount = $analysis.TotalRules.ToString()
                        $result.HiddenRules = $analysis.SuspiciousHidden.ToString()
                        $result.SuspiciousRules = $analysis.SuspiciousVisible.ToString()
                        $result.ExternalForwarding = if ($analysis.HasExternalForwarding) { "Yes" } else { "No" }
                    }
                } catch {
                    # Keep default values if analysis fails
                }
                
                # Check permissions (only if likely to have them)
                try {
                    $delegates = Analyze-MailboxDelegates -UserPrincipalName $upn
                    $fullAccess = Analyze-MailboxPermissions -UserPrincipalName $upn
                    $result.Delegates = $delegates
                    $result.FullAccess = $fullAccess
                } catch {
                    # Keep default values if analysis fails
                }
                
                $results[$upn] = $result
            }
        }
    }
    
    return $results
}

# Function to analyze mailbox rules with improved hidden rule detection
function Analyze-MailboxRulesEnhanced {
    param(
        [Parameter(Mandatory=$true)]
        [array]$Rules,
        [Parameter(Mandatory=$true)]
        [array]$BaseSuspiciousKeywords
    )
    
    $totalRules = $Rules.Count
    $suspiciousHiddenCount = 0
    $suspiciousVisibleCount = 0
    $hasExternalForwarding = $false
    
    foreach ($rule in $Rules) {
        # Enhanced hidden rule detection - only count truly suspicious hidden rules
        $isSuspiciousHidden = $false
        if ($rule.IsHidden) {
            # Check if this is a legitimate hidden rule or potentially malicious
            $ruleName = if ($rule.Name) { $rule.Name.ToLower() } else { "" }
            
            # Legitimate hidden rule patterns (system-generated, shared mailbox rules, etc.)
            $legitimatePatterns = @(
                "system", "default", "outlook", "microsoft", "office", "exchange",
                "shared", "team", "group", "distribution", "dl", "mailbox",
                "automatic", "auto", "sync", "migration", "upgrade",
                "clutter", "focused", "junk", "spam", "archive", "retention",
                "compliance", "legal", "hold", "litigation", "ediscovery"
            )
            
            $isLegitimate = $false
            foreach ($pattern in $legitimatePatterns) {
                if ($ruleName -like "*$pattern*") {
                    $isLegitimate = $true
                    break
                }
            }
            
            # Additional checks for suspicious hidden rules
            if (-not $isLegitimate) {
                # Check for suspicious keywords in hidden rules
                foreach ($kw in $BaseSuspiciousKeywords) {
                    if ($ruleName -match [regex]::Escape($kw)) {
                        $isSuspiciousHidden = $true
                        break
                    }
                }
                
                # Check for symbols-only names in hidden rules
                if (-not $isSuspiciousHidden -and $ruleName.Length -gt 0) {
                    $textCharacters = $ruleName -replace '[^\p{L}\p{N}\s]', ''
                    if ([string]::IsNullOrWhiteSpace($textCharacters)) {
                        $isSuspiciousHidden = $true
                    }
                }
                
                # Check for external forwarding in hidden rules
                if (-not $isSuspiciousHidden -and $rule.ForwardTo -and $rule.ForwardTo -match '@') {
                    $isSuspiciousHidden = $true
                }
            }
            
            # Only count as suspicious hidden if it meets suspicious criteria
            if ($isSuspiciousHidden) {
                $suspiciousHiddenCount++
            }
        }
        
        # Check for suspicious keywords in visible rules
        $isSuspiciousVisible = $false
        if (-not $rule.IsHidden) {  # Only check visible rules for suspicious keywords
            foreach ($kw in $BaseSuspiciousKeywords) {
                if ($rule.Name -and $rule.Name -match [regex]::Escape($kw)) {
                    $isSuspiciousVisible = $true
                    break
                }
            }
            
            # Check for symbols-only names in visible rules
            if (-not $isSuspiciousVisible -and $rule.Name -and $rule.Name.Length -gt 0) {
                $textCharacters = $rule.Name -replace '[^\p{L}\p{N}\s]', ''
                if ([string]::IsNullOrWhiteSpace($textCharacters)) {
                    $isSuspiciousVisible = $true
                }
            }
        }
        
        # Count suspicious rules (visible rules with suspicious characteristics)
        if ($isSuspiciousVisible) {
            $suspiciousVisibleCount++
        }
        
        # Check for external forwarding
        if ($rule.ForwardTo -and $rule.ForwardTo -match '@') {
            $hasExternalForwarding = $true
        }
    }
    
    return @{
        TotalRules = $totalRules
        SuspiciousHidden = $suspiciousHiddenCount
        SuspiciousVisible = $suspiciousVisibleCount
        HasExternalForwarding = $hasExternalForwarding
    }
}

# Function to show error messages in user-friendly format
function Show-UserFriendlyError {
    param($errorObject, $operation = "Operation")
    
    # Check if this is a user cancellation
    $errorMessage = $errorObject.Exception.Message
    $isUserCancellation = $errorMessage -match "User cancelled|Operation cancelled|User canceled|Authentication cancelled|Authentication canceled" -or 
                         $errorMessage -match "AADSTS50020|AADSTS50076|AADSTS50079" -or
                         $errorMessage -match "The user cancelled the authentication"
    
    if ($isUserCancellation) {
        # User cancelled - just update status without showing error popup
        $statusLabel.Text = "$operation cancelled by user."
        return
    }
    
    # Handle other error types
    $userFriendlyMessage = switch -Wildcard ($errorMessage) {
        "*Access is denied*" { "Access denied. Please check your permissions and try again." }
        "*Could not connect*" { "Connection failed. Please check your internet connection and credentials." }
        "*The remote server returned an error*" { "Server error. Please try again later." }
        "*Object reference not set*" { "Data not found. Please refresh and try again." }
        "*User cancelled*" { "Operation cancelled by user." }
        "*Operation cancelled*" { "Operation cancelled by user." }
        default { "An error occurred during $operation`: $errorMessage" }
    }
    
    [System.Windows.Forms.MessageBox]::Show($userFriendlyMessage, "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    $statusLabel.Text = "Error: $operation failed"
}

# Function to update Entra tab button states
function UpdateEntraButtonStates {
    $hasPath = -not [string]::IsNullOrWhiteSpace($entraOutputFolderTextBox.Text)
    $checkedCount = 0
    for ($i = 0; $i -lt $entraUserGrid.Rows.Count; $i++) {
        if ($entraUserGrid.Rows[$i].Cells["Select"].Value -eq $true) { $checkedCount++ }
    }
    # Only export buttons require export folder path and selection
    $entraExportSignInLogsButton.Enabled = $hasPath -and ($checkedCount -gt 0)
    $entraExportAuditLogsButton.Enabled = $hasPath -and ($checkedCount -eq 1)
    # View, User Details, and Analyze MFA buttons are always enabled
    $entraViewSignInLogsButton.Enabled = $true
    $entraViewAuditLogsButton.Enabled = $true
    $entraDetailsFetchButton.Enabled = $true
    $entraMfaFetchButton.Enabled = $true
    # User management buttons are always enabled when connected to Graph
    # Select All/Deselect All buttons enabled when users are loaded
    $entraSelectAllButton.Enabled = ($entraUserGrid.Rows.Count -gt 0)
    $entraDeselectAllButton.Enabled = ($entraUserGrid.Rows.Count -gt 0)
    
    # User management buttons are always enabled when connected to Graph
    $entraBlockUserButton.Enabled = $true
    $entraUnblockUserButton.Enabled = $true
    $entraRevokeSessionsButton.Enabled = $true
    $entraResetPasswordButton.Enabled = $true
    $entraOpenDefenderRestrictedUsersButton.Enabled = $true
    $entraRequirePwdChangeButton.Enabled = $true
    $entraRefreshRolesButton.Enabled = $true
    $entraViewAdminsButton.Enabled = $true
    
    # Load buttons are enabled when connected to Graph
    $loadAllUsersButton.Enabled = $script:graphConnection
    $searchUsersButton.Enabled = $script:graphConnection
}

# Function to generate professional report
function Generate-ProfessionalReport {
    $report = @"
# Microsoft 365 Environment Analysis Report
**Generated:** $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
**Tool:** Microsoft 365 Management Tool v8.0

## Executive Summary
This report provides a comprehensive analysis of the Microsoft 365 environment, including Exchange Online configuration and Entra ID (Azure AD) security posture.

---

## Exchange Online Analysis

### Connection Status
- **Status:** $(if ($script:currentExchangeConnection) { "Connected" } else { "Not Connected" })
- **Mailboxes Loaded:** $(if ($script:allLoadedMailboxUPNs) { $script:allLoadedMailboxUPNs.Count } else { "0" })

### Mailbox Analysis
$(if ($script:allLoadedMailboxUPNs -and $script:allLoadedMailboxUPNs.Count -gt 0) {
    $mailboxStats = @"
- **Total Mailboxes:** $($script:allLoadedMailboxUPNs.Count)
- **Sample Mailboxes:** $($script:allLoadedMailboxUPNs[0..4] -join ", ")
$(if ($script:allLoadedMailboxUPNs.Count -gt 5) { "- **Additional:** +$($script:allLoadedMailboxUPNs.Count - 5) more mailboxes" })
"@
    $mailboxStats
} else {
    "- No mailboxes loaded"
})

### Inbox Rules Analysis
$(if ($userMailboxGrid.Rows.Count -gt 0) {
    $selectedCount = 0
    for ($i = 0; $i -lt $userMailboxGrid.Rows.Count; $i++) {
        if ($userMailboxGrid.Rows[$i].Cells["Select"].Value -eq $true) { $selectedCount++ }
    }
    "- **Mailboxes Selected for Analysis:** $selectedCount"
} else {
    "- No mailboxes selected"
})

### Transport Rules & Connectors
- **Transport Rules:** Available for review via Manage Transport Rules
- **Connectors:** Available for review via Manage Connectors
- **Restricted Senders:** Available for management

---

## Entra ID (Azure AD) Analysis

### Connection Status
- **Status:** $(if ($script:graphConnection) { "Connected" } else { "Not Connected" })

### User Management
$(if ($entraUserGrid.Rows.Count -gt 0) {
    $userStats = @"
- **Total Users Loaded:** $($entraUserGrid.Rows.Count)
- **User Management Features:** Available
  - Block/Unblock User Sign-in
  - Revoke User Sessions
  - View User Details & Roles
  - MFA Analysis
"@
    $userStats
} else {
    "- No users loaded"
})

### Security Features
- **Sign-in Logs:** Available for export and analysis
- **Audit Logs:** Available for export and analysis
- **MFA Analysis:** Available for individual users
- **User Role Analysis:** Available

---

## Security Posture Assessment

### Exchange Online Security
- **Inbox Rules Review:** $(if ($userMailboxGrid.Rows.Count -gt 0) { "Available" } else { "Not Available" })
- **Forwarding Analysis:** Available
- **External Access:** Monitored via rules analysis
- **Suspicious Keywords:** Configured for detection

### Entra ID Security
- **User Account Status:** $(if ($entraUserGrid.Rows.Count -gt 0) { "Available for review" } else { "Not available" })
- **Sign-in Monitoring:** Available
- **Session Management:** Available
- **MFA Status:** Available for analysis

---

## Recommendations

### Immediate Actions
1. Review any suspicious inbox rules identified
2. Check for unauthorized external forwarding
3. Verify user account status and permissions
4. Review sign-in logs for suspicious activity

### Ongoing Monitoring
1. Regular inbox rules audits
2. Monitor user sign-in patterns
3. Review transport rules and connectors
4. Maintain MFA compliance

---

## Technical Details

### Environment Information
- **Tool Version:** 8.0
- **Report Generated:** $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
- **Exchange Connection:** $(if ($script:currentExchangeConnection) { "Active" } else { "Inactive" })
- **Graph Connection:** $(if ($script:graphConnection) { "Active" } else { "Inactive" })

### Data Sources
- Exchange Online PowerShell
- Microsoft Graph API
- User mailbox analysis
- Sign-in and audit logs

---

*This report was generated automatically by the Microsoft 365 Management Tool. For detailed analysis, use the individual tabs for specific data exports.*
"@

    return $report
}

# Function to generate Obsidian note format
function Generate-ObsidianNote {
    $note = "Microsoft 365 Environment Analysis`n"
    $note += "Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`n`n"
    $note += "## Environment Overview`n"
    $note += "- Exchange Online: $(if ($script:currentExchangeConnection) { 'Connected' } else { 'Not Connected' })`n"
    $note += "- Entra ID: $(if ($script:graphConnection) { 'Connected' } else { 'Not Connected' })`n"
    $note += "- Mailboxes: $(if ($script:allLoadedMailboxUPNs) { $script:allLoadedMailboxUPNs.Count } else { '0' })`n"
    $note += "- Users: $(if ($entraUserGrid.Rows.Count -gt 0) { $entraUserGrid.Rows.Count } else { '0' })`n`n"
    
    $note += "## Exchange Online Analysis`n`n"
    $note += "### Mailbox Status`n"
    if ($script:allLoadedMailboxUPNs -and $script:allLoadedMailboxUPNs.Count -gt 0) {
        $note += "- Total mailboxes: $($script:allLoadedMailboxUPNs.Count)`n"
    } else {
        $note += "- No mailboxes loaded`n"
    }
    $note += "`n### Selected for Analysis`n"
    if ($userMailboxGrid.Rows.Count -gt 0) {
        $selectedCount = 0
        for ($i = 0; $i -lt $userMailboxGrid.Rows.Count; $i++) {
            if ($userMailboxGrid.Rows[$i].Cells["Select"].Value -eq $true) { $selectedCount++ }
        }
        $note += "- Selected mailboxes: $selectedCount`n"
    } else {
        $note += "- No mailboxes selected`n"
    }
    
    $note += "`n## Entra ID Security`n`n"
    $note += "### User Management`n"
    if ($entraUserGrid.Rows.Count -gt 0) {
        $note += "- Loaded users: $($entraUserGrid.Rows.Count)`n"
        $note += "- User management features available`n"
    } else {
        $note += "- No users loaded`n"
    }
    
    $note += "`n### Available Features`n"
    $note += "- Sign-in logs export`n"
    $note += "- Audit logs export`n"
    $note += "- MFA analysis`n"
    $note += "- User role analysis`n"
    $note += "- Session revocation`n"
    $note += "- User blocking/unblocking`n"
    
    $note += "`n## Security Assessment`n`n"
    $note += "### Exchange Security`n"
    $note += "- Inbox rules analysis: $(if ($userMailboxGrid.Rows.Count -gt 0) { 'Available' } else { 'Not available' })`n"
    $note += "- Forwarding analysis: Available`n"
    $note += "- Transport rules: Available`n"
    $note += "- Connectors review: Available`n"
    
    $note += "`n### Entra ID Security`n"
    $note += "- User account monitoring: $(if ($entraUserGrid.Rows.Count -gt 0) { 'Available' } else { 'Not available' })`n"
    $note += "- Sign-in monitoring: Available`n"
    $note += "- Session management: Available`n"
    $note += "- MFA compliance: Available`n"
    
    $note += "`n## Action Items`n`n"
    $note += "### Immediate`n"
    $note += "- [ ] Review suspicious inbox rules`n"
    $note += "- [ ] Check external forwarding`n"
    $note += "- [ ] Verify user permissions`n"
    $note += "- [ ] Review sign-in logs`n"
    
    $note += "`n### Ongoing`n"
    $note += "- [ ] Regular inbox rules audits`n"
    $note += "- [ ] Monitor sign-in patterns`n"
    $note += "- [ ] Review transport rules`n"
    $note += "- [ ] Maintain MFA compliance`n"
    
    $note += "`n## Technical Notes`n`n"
    $note += "Tool: Microsoft 365 Management Tool v8.0`n"
    $note += "Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`n"
    $note += "Exchange: $(if ($script:currentExchangeConnection) { 'Active' } else { 'Inactive' })`n"
    $note += "Graph: $(if ($script:graphConnection) { 'Active' } else { 'Inactive' })`n`n"
    $note += "---`n"
    $note += "Tags: #microsoft365 #security #exchange #entra #analysis"

    return $note
}

# Function to populate unified account grid
function Update-UnifiedAccountGrid {
    $unifiedAccountGrid.Rows.Clear()
    
    # Create a combined list of accounts from both Exchange and Entra ID
    $allAccounts = @{}
    
    # Add Exchange Online accounts with detailed data
    if ($script:allLoadedMailboxUPNs -and $script:allLoadedMailboxUPNs.Count -gt 0) {
        foreach ($mailbox in $script:allLoadedMailboxUPNs) {
            # Get detailed mailbox data from the Exchange grid
            $mailboxData = $null
            for ($i = 0; $i -lt $userMailboxGrid.Rows.Count; $i++) {
                if ($userMailboxGrid.Rows[$i].Cells["UserPrincipalName"].Value -eq $mailbox) {
                    $mailboxData = @{
                        RulesCount = $userMailboxGrid.Rows[$i].Cells["RulesCount"].Value
                        SuspiciousRules = $userMailboxGrid.Rows[$i].Cells["SuspiciousRules"].Value
                        ExternalForwarding = $userMailboxGrid.Rows[$i].Cells["ExternalForwarding"].Value
                        Delegates = $userMailboxGrid.Rows[$i].Cells["Delegates"].Value
                        FullAccess = $userMailboxGrid.Rows[$i].Cells["FullAccess"].Value
                    }
                    break
                }
            }
            
            $allAccounts[$mailbox] = @{
                UPN = $mailbox
                DisplayName = $mailbox
                ExchangeStatus = "Available"
                EntraStatus = "Unknown"
                RulesCount = if ($mailboxData) { $mailboxData.RulesCount } else { "0" }
                SuspiciousRules = if ($mailboxData) { $mailboxData.SuspiciousRules } else { "0" }
                ExternalForwarding = if ($mailboxData) { $mailboxData.ExternalForwarding } else { "Unknown" }
                Delegates = if ($mailboxData) { $mailboxData.Delegates } else { "Unknown" }
                FullAccess = if ($mailboxData) { $mailboxData.FullAccess } else { "Unknown" }
            }
        }
    }
    
    # Add Entra ID accounts with detailed data
    if ($entraUserGrid.Rows.Count -gt 0) {
        for ($i = 0; $i -lt $entraUserGrid.Rows.Count; $i++) {
            $upn = $entraUserGrid.Rows[$i].Cells["UserPrincipalName"].Value
            $displayName = $entraUserGrid.Rows[$i].Cells["DisplayName"].Value
            $licensed = $entraUserGrid.Rows[$i].Cells["Licensed"].Value
            
            if ($allAccounts.ContainsKey($upn)) {
                $allAccounts[$upn].EntraStatus = "Available"
                $allAccounts[$upn].DisplayName = $displayName
                $allAccounts[$upn].Licensed = $licensed
            } else {
                $allAccounts[$upn] = @{
                    UPN = $upn
                    DisplayName = $displayName
                    ExchangeStatus = "Unknown"
                    EntraStatus = "Available"
                    RulesCount = "0"
                    SuspiciousRules = "0"
                    ExternalForwarding = "Unknown"
                    Delegates = "Unknown"
                    FullAccess = "Unknown"
                    Licensed = $licensed
                }
            }
        }
    }
    
    # Populate the grid with enhanced data
    foreach ($account in $allAccounts.Values) {
        $rowIdx = $unifiedAccountGrid.Rows.Add()
        $unifiedAccountGrid.Rows[$rowIdx].Cells["Select"].Value = $false
        $unifiedAccountGrid.Rows[$rowIdx].Cells["UserPrincipalName"].Value = $account.UPN
        $unifiedAccountGrid.Rows[$rowIdx].Cells["DisplayName"].Value = $account.DisplayName
        $unifiedAccountGrid.Rows[$rowIdx].Cells["ExchangeStatus"].Value = $account.ExchangeStatus
        $unifiedAccountGrid.Rows[$rowIdx].Cells["EntraStatus"].Value = $account.EntraStatus
        
        # Store additional data in the row for report generation
        $unifiedAccountGrid.Rows[$rowIdx].Tag = $account
    }
}

# Function to get selected accounts for unified reporting
function Get-SelectedUnifiedAccounts {
    $selectedAccounts = @()
    
    for ($i = 0; $i -lt $unifiedAccountGrid.Rows.Count; $i++) {
        if ($unifiedAccountGrid.Rows[$i].Cells["Select"].Value -eq $true) {
            $upn = $unifiedAccountGrid.Rows[$i].Cells["UserPrincipalName"].Value
            $displayName = $unifiedAccountGrid.Rows[$i].Cells["DisplayName"].Value
            $exchangeStatus = $unifiedAccountGrid.Rows[$i].Cells["ExchangeStatus"].Value
            $entraStatus = $unifiedAccountGrid.Rows[$i].Cells["EntraStatus"].Value
            
            # Get detailed data from the row's Tag property
            $detailedData = $unifiedAccountGrid.Rows[$i].Tag
            
            $selectedAccounts += [PSCustomObject]@{
                UserPrincipalName = $upn
                DisplayName = $displayName
                ExchangeStatus = $exchangeStatus
                EntraStatus = $entraStatus
                RulesCount = if ($detailedData) { $detailedData.RulesCount } else { "0" }
                SuspiciousRules = if ($detailedData) { $detailedData.SuspiciousRules } else { "0" }
                ExternalForwarding = if ($detailedData) { $detailedData.ExternalForwarding } else { "Unknown" }
                Delegates = if ($detailedData) { $detailedData.Delegates } else { "Unknown" }
                FullAccess = if ($detailedData) { $detailedData.FullAccess } else { "Unknown" }
                Licensed = if ($detailedData) { $detailedData.Licensed } else { "Unknown" }
            }
        }
    }
    
    return $selectedAccounts
}

# Function to generate unified professional report
function Generate-UnifiedProfessionalReport {
    param($selectedAccounts)
    
    # Build report content dynamically to avoid here-string issues
    $report = "Microsoft 365 Comprehensive Management Report`n"
    $report += "Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`n"
    $report += "Tool: Microsoft 365 Management Tool v8.0`n`n"
    
    $report += "Executive Summary`n"
    
    # Get the first selected user for single-user focus
    $firstSelectedUser = $selectedAccounts | Where-Object { $_.EntraStatus -eq "Available" } | Select-Object -First 1
    
    if ($firstSelectedUser) {
        $report += "User Account: $($firstSelectedUser.DisplayName)`n"
        $report += "User Principal Name: $($firstSelectedUser.UserPrincipalName)`n`n"
        
        $report += "This security analysis focuses on the above user account across Exchange Online and Entra ID configurations.`n`n"
    } else {
        $report += "This comprehensive report consolidates all available data from Exchange Online and Entra ID management functions, providing a complete overview of the Microsoft 365 environment configuration and security posture.`n`n"
    }
    
    $report += "Exchange Online Configuration`n`n"
    $report += "Connection Status`n"
    $report += "- Status: $(if ($script:currentExchangeConnection) { 'Connected' } else { 'Not Connected' })`n"
    $report += "- Mailboxes Loaded: $(if ($script:allLoadedMailboxUPNs) { $script:allLoadedMailboxUPNs.Count } else { '0' })`n`n"
    
    # Mailbox Analysis
    if ($selectedAccounts.Count -gt 0) {
        $selectedCount = 0
        $totalRules = 0
        $suspiciousRules = 0
        $externalForwarding = 0
        
        foreach ($account in $selectedAccounts) {
            if ($account.ExchangeStatus -eq "Available") {
                $selectedCount++ 
                $rulesCount = [int]$account.RulesCount
                $totalRules += $rulesCount
                if ($rulesCount -gt 0) {
                    $suspiciousRules += [int]$account.SuspiciousRules
                    if ($account.ExternalForwarding -eq "Yes") {
                        $externalForwarding++
                    }
                }
            }
        }
        
        $report += "Mailbox Inbox Rules Analysis`n"
        $report += "- Mailboxes Selected for Analysis: $selectedCount`n"
        $report += "- Total Inbox Rules Found: $totalRules`n"
        $report += "- Suspicious Rules Detected: $suspiciousRules`n"
        $report += "- Mailboxes with External Forwarding: $externalForwarding`n`n"
        
        $report += "Detailed Mailbox Analysis`n"
        foreach ($account in $selectedAccounts) {
            if ($account.ExchangeStatus -eq "Available") {
                $report += "- $($account.UserPrincipalName)`n"
                $report += "  - Total Rules: $($account.RulesCount)`n"
                $report += "  - Suspicious Rules: $($account.SuspiciousRules)`n"
                $report += "  - External Forwarding: $($account.ExternalForwarding)`n"
                $report += "  - Delegates: $($account.Delegates)`n"
                $report += "  - Full Access Users: $($account.FullAccess)`n"
                
                # Add detailed suspicious rule analysis
                if ([int]$account.RulesCount -gt 0) {
                    $report += "  - Suspicious Rule Analysis:`n"
                    $report += "    * Rules with symbols-only names (no text characters) are flagged as suspicious`n"
                    $report += "    * Hidden rules are flagged as suspicious`n"
                    $report += "    * Rules with suspicious keywords are flagged`n"
                    $report += "    * Rules with external forwarding are flagged`n"
                }
                $report += "`n"
            }
        }
    } else {
        $report += "Mailbox Inbox Rules Analysis`n"
        $report += "- No mailboxes selected for analysis`n`n"
    }
    
    # Transport Rules
    $report += "Transport Rules Configuration`n"
    try {
        $transportRules = Get-TransportRule -ErrorAction SilentlyContinue | Select-Object Name, State, Priority, Enabled
        if ($transportRules) {
            $report += "- Total Transport Rules: $($transportRules.Count)`n"
            $report += "- Active Rules: $(($transportRules | Where-Object { $_.State -eq 'Enabled' }).Count)`n"
            $report += "- Inactive Rules: $(($transportRules | Where-Object { $_.State -eq 'Disabled' }).Count)`n`n"
            
            $report += "Transport Rules Details`n"
            foreach ($rule in $transportRules | Select-Object -First 10) {
                $report += "- $($rule.Name) (Priority: $($rule.Priority), State: $($rule.State))`n"
            }
            if ($transportRules.Count -gt 10) {
                $report += "- ... and $($transportRules.Count - 10) more rules`n"
            }
            $report += "`n"
        } else {
            $report += "- No transport rules found or access denied`n`n"
        }
    } catch {
        $report += "- Transport rules data unavailable: $($_.Exception.Message)`n`n"
    }
    
    # Connectors
    $report += "Connectors Configuration`n"
    try {
        # Try different connector cmdlets that might be available
        $connectors = $null
        
        # First try Get-Connector (Exchange Online)
        try {
            $connectors = Get-Connector -ErrorAction Stop | Select-Object Name, ConnectorType, Enabled
        } catch {
            # Try Get-InboundConnector (Exchange Online)
            try {
                $inboundConnectors = Get-InboundConnector -ErrorAction Stop | Select-Object Name, ConnectorType, Enabled
                $outboundConnectors = Get-OutboundConnector -ErrorAction Stop | Select-Object Name, ConnectorType, Enabled
                $connectors = @($inboundConnectors) + @($outboundConnectors)
            } catch {
                # Try Get-HostedConnector (Exchange Online)
                try {
                    $connectors = Get-HostedConnector -ErrorAction Stop | Select-Object Name, ConnectorType, Enabled
                } catch {
                    $connectors = $null
                }
            }
        }
        
        if ($connectors -and $connectors.Count -gt 0) {
            $report += "- Total Connectors: $($connectors.Count)`n"
            $report += "- Enabled Connectors: $(($connectors | Where-Object { $_.Enabled -eq $true }).Count)`n"
            $report += "- Disabled Connectors: $(($connectors | Where-Object { $_.Enabled -eq $false }).Count)`n`n"
            
            $report += "Connectors Details`n"
            foreach ($connector in $connectors | Select-Object -First 10) {
                $report += "- $($connector.Name) (Type: $($connector.ConnectorType), Enabled: $($connector.Enabled))`n"
            }
            if ($connectors.Count -gt 10) {
                $report += "- ... and $($connectors.Count - 10) more connectors`n"
            }
            $report += "`n"
        } else {
            $report += "- No connectors found or access denied`n`n"
        }
    } catch {
        $report += "- Connectors data unavailable: $($_.Exception.Message)`n`n"
    }
    
    # Entra ID Section
    $report += "Entra ID (Azure AD) Configuration`n`n"
    $report += "Connection Status`n"
    $report += "- Status: $(if ($script:graphConnection) { 'Connected' } else { 'Not Connected' })`n"
    $report += "- Users Loaded: $(if ($entraUserGrid.Rows.Count -gt 0) { $entraUserGrid.Rows.Count } else { '0' })`n`n"
    
    # User Analysis
    if ($selectedAccounts.Count -gt 0) {
        $selectedCount = 0
        $licensedUsers = 0
        $unlicensedUsers = 0
        
        foreach ($account in $selectedAccounts) {
            if ($account.EntraStatus -eq "Available") {
                $selectedCount++ 
                if ($account.Licensed -eq "Yes") {
                    $licensedUsers++
                } else {
                    $unlicensedUsers++
                }
            }
        }
        
        $report += "User Account Analysis`n"
        $report += "- Users Selected for Analysis: $selectedCount`n"
        $report += "- Licensed Users: $licensedUsers`n"
        $report += "- Unlicensed Users: $unlicensedUsers`n`n"
        
        $report += "Selected User Details`n"
        foreach ($account in $selectedAccounts) {
            if ($account.EntraStatus -eq "Available") {
                $report += "- $($account.DisplayName) ($($account.UserPrincipalName))`n"
                $report += "  - Licensed: $($account.Licensed)`n"
                
                # Get MFA status for this user
                try {
                    $mfaStatus = Get-EntraUserMfaStatus -UserPrincipalName $account.UserPrincipalName -ErrorAction SilentlyContinue
                    if ($mfaStatus) {
                        $report += "  - MFA Status: $($mfaStatus.OverallStatus)`n"
                        $report += "  - MFA Summary: $($mfaStatus.Summary)`n"
                        if ($mfaStatus.PerUserMfa.Enabled) {
                            $report += "  - MFA Methods: $($mfaStatus.PerUserMfa.Details)`n"
                        }
                    } else {
                        $report += "  - MFA Status: Unable to retrieve`n"
                    }
                } catch {
                    $report += "  - MFA Status: Error retrieving MFA data`n"
                }
                $report += "`n"
            }
        }
    } else {
        $report += "User Account Analysis`n"
        $report += "- No users selected for analysis`n`n"
    }
    
    # Sign-in Logs
    $report += "Sign-in Logs Summary`n"
    try {
        # Get selected users for sign-in logs
        $selectedUsers = @()
        foreach ($account in $selectedAccounts) {
            if ($account.EntraStatus -eq "Available") {
                if (-not [string]::IsNullOrWhiteSpace($account.UserPrincipalName)) {
                    $selectedUsers += $account.UserPrincipalName
                }
            }
        }
        
        if ($selectedUsers.Count -gt 0) {
            $signInLogs = Get-EntraSignInLogs -UserPrincipalNames $selectedUsers -Days 7 -ErrorAction SilentlyContinue
            if ($signInLogs -and $signInLogs.Count -gt 0) {
                $recentLogs = $signInLogs | Select-Object -First 50
                $successfulLogins = ($recentLogs | Where-Object { $_.Status -eq "Success" }).Count
                $failedLogins = ($recentLogs | Where-Object { $_.Status -eq "Failure" }).Count
                $suspiciousLogins = ($recentLogs | Where-Object { $_.RiskLevel -eq "High" -or $_.RiskLevel -eq "Medium" }).Count
                
                # Analyze non-US sign-ins
                $nonUSSignIns = @()
                $usSignIns = @()
                foreach ($log in $recentLogs) {
                    if ($log.Location -and $log.Location.CountryOrRegion) {
                        if ($log.Location.CountryOrRegion -ne "US" -and $log.Location.CountryOrRegion -ne "United States") {
                            $nonUSSignIns += $log
                        } else {
                            $usSignIns += $log
                        }
                    }
                }
                
                $report += "- Recent Sign-in Activity (Last 50 events)`n"
                $report += "- Total Events: $($recentLogs.Count)`n"
                $report += "- Successful Logins: $successfulLogins`n"
                $report += "- Failed Logins: $failedLogins`n"
                $report += "- Suspicious Logins: $suspiciousLogins`n"
                $report += "- US Sign-ins: $($usSignIns.Count)`n"
                $report += "- Non-US Sign-ins: $($nonUSSignIns.Count)`n`n"
                
                $report += "Recent Sign-in Events`n"
                foreach ($log in $recentLogs | Select-Object -First 10) {
                    $location = if ($log.Location -and $log.Location.CountryOrRegion) { $log.Location.CountryOrRegion } else { "Unknown" }
                    $report += "- $($log.UserPrincipalName) - $($log.CreatedDateTime) - Status: $($log.Status) - Risk: $($log.RiskLevel) - Location: $location`n"
                }
                if ($recentLogs.Count -gt 10) {
                    $report += "- ... and $($recentLogs.Count - 10) more events`n"
                }
                $report += "`n"
                
                # Show non-US sign-ins if any found
                if ($nonUSSignIns.Count -gt 0) {
                    $report += "Non-US Sign-in Events (Security Alert)`n"
                    foreach ($log in $nonUSSignIns | Select-Object -First 5) {
                        $location = if ($log.Location -and $log.Location.CountryOrRegion) { $log.Location.CountryOrRegion } else { "Unknown" }
                        $city = if ($log.Location -and $log.Location.City) { $log.Location.City } else { "Unknown" }
                        $report += "- $($log.UserPrincipalName) - $($log.CreatedDateTime) - Status: $($log.Status) - Risk: $($log.RiskLevel) - Location: $city, $location`n"
                    }
                    if ($nonUSSignIns.Count -gt 5) {
                        $report += "- ... and $($nonUSSignIns.Count - 5) more non-US events`n"
                    }
                    $report += "`n"
                }
            } else {
                $report += "- No sign-in logs available for selected users`n`n"
            }
        } else {
            $report += "- No users selected for sign-in log analysis`n`n"
        }
    } catch {
        $report += "- Sign-in logs data unavailable: $($_.Exception.Message)`n`n"
    }
    
    # Audit Logs
    $report += "Audit Logs Summary`n"
    try {
        # Get selected users for audit logs
        $selectedUsers = @()
        foreach ($account in $selectedAccounts) {
            if ($account.EntraStatus -eq "Available") {
                if (-not [string]::IsNullOrWhiteSpace($account.UserPrincipalName)) {
                    $selectedUsers += $account.UserPrincipalName
                }
            }
        }
        
        if ($selectedUsers.Count -gt 0) {
            $auditLogs = Get-EntraUserAuditLogs -UserPrincipalName $selectedUsers[0] -Days 7 -ErrorAction SilentlyContinue
            if ($auditLogs -and $auditLogs.Count -gt 0) {
                $recentAudits = $auditLogs | Select-Object -First 50
                $adminActions = ($recentAudits | Where-Object { $_.Category -eq "AdministrativeUnit" }).Count
                $userManagement = ($recentAudits | Where-Object { $_.Category -eq "UserManagement" }).Count
                $applicationActivity = ($recentAudits | Where-Object { $_.Category -eq "Application" }).Count
                
                $report += "- Recent Audit Activity (Last 50 events)`n"
                $report += "- Total Events: $($recentAudits.Count)`n"
                $report += "- Administrative Actions: $adminActions`n"
                $report += "- User Management Events: $userManagement`n"
                $report += "- Application Activity: $applicationActivity`n`n"
                
                $report += "Recent Audit Events`n"
                foreach ($log in $recentAudits | Select-Object -First 10) {
                    $report += "- $($log.UserPrincipalName) - $($log.CreatedDateTime) - Category: $($log.Category) - Activity: $($log.Activity)`n"
                }
                if ($recentAudits.Count -gt 10) {
                    $report += "- ... and $($recentAudits.Count - 10) more events`n"
                }
                $report += "`n"
            } else {
                $report += "- No audit logs available for selected users`n`n"
            }
        } else {
            $report += "- No users selected for audit log analysis`n`n"
        }
    } catch {
        $report += "- Audit logs data unavailable: $($_.Exception.Message)`n`n"
    }
    
    # Security Assessment
    $report += "Security Posture Assessment`n`n"
    
    $report += "Exchange Online Security Findings`n"
    if ($selectedAccounts.Count -gt 0) {
        $selectedCount = 0
        $totalSuspiciousRules = 0
        $externalForwardingCount = 0
        
        foreach ($account in $selectedAccounts) {
            if ($account.ExchangeStatus -eq "Available") {
                $selectedCount++ 
                $totalSuspiciousRules += [int]$account.SuspiciousRules
                if ($account.ExternalForwarding -eq "Yes") {
                    $externalForwardingCount++
                }
            }
        }
        
        $report += "- Mailboxes Analyzed: $selectedCount`n"
        $report += "- Total Suspicious Rules Found: $totalSuspiciousRules`n"
        $report += "- Mailboxes with External Forwarding: $externalForwardingCount`n"
        $riskLevel = if ($totalSuspiciousRules -gt 0 -or $externalForwardingCount -gt 0) { "HIGH - Immediate attention required" } else { "LOW - No immediate concerns detected" }
        $report += "- Risk Level: $riskLevel`n`n"
    } else {
        $report += "- No mailboxes analyzed`n`n"
    }
    
    $report += "Entra ID Security Findings`n"
    if ($selectedAccounts.Count -gt 0) {
        $selectedCount = 0
        $unlicensedUsers = 0
        
        foreach ($account in $selectedAccounts) {
            if ($account.EntraStatus -eq "Available") {
                $selectedCount++ 
                if ($account.Licensed -ne "Yes") {
                    $unlicensedUsers++
                }
            }
        }
        
        $report += "- Users Analyzed: $selectedCount`n"
        $report += "- Unlicensed Users: $unlicensedUsers`n"
        $report += "- MFA Status: Available for individual analysis`n"
        $report += "- Session Management: Available for revocation`n`n"
    } else {
        $report += "- No users analyzed`n`n"
    }
    

    
    # Technical Details
    $report += "Technical Details`n`n"
    $report += "Environment Information`n"
    $report += "- Tool Version: 8.0`n"
    $report += "- Report Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`n"
    $report += "- Exchange Connection: $(if ($script:currentExchangeConnection) { 'Active' } else { 'Inactive' })`n"
    $report += "- Graph Connection: $(if ($script:graphConnection) { 'Active' } else { 'Inactive' })`n`n"
    
    $report += "Data Sources`n"
    $report += "- Exchange Online PowerShell (Inbox Rules, Transport Rules, Connectors)`n"
    $report += "- Microsoft Graph API (Users, Sign-in Logs, Audit Logs)`n"
    $report += "- Real-time mailbox analysis`n"
    $report += "- Security posture assessment`n`n"
    
    $report += "This comprehensive report was generated automatically by the Microsoft 365 Management Tool, consolidating all available management data for complete environment analysis."

    return $report
}

# Function to generate unified Obsidian note format
function Generate-UnifiedObsidianNote {
    param($selectedAccounts)
    
    # Build note content dynamically to avoid here-string issues
    $note = "Microsoft 365 Comprehensive Management Report`n"
    $note += "Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`n`n"
    
    $note += "## Executive Summary`n"
    
    # Get the first selected user for single-user focus
    $firstSelectedUser = $selectedAccounts | Where-Object { $_.EntraStatus -eq "Available" } | Select-Object -First 1
    
    if ($firstSelectedUser) {
        $note += "**User Account:** $($firstSelectedUser.DisplayName)`n"
        $note += "**User Principal Name:** $($firstSelectedUser.UserPrincipalName)`n`n"
        
        $note += "This security analysis focuses on the above user account across Exchange Online and Entra ID configurations.`n`n"
    } else {
        $note += "This comprehensive report consolidates all available data from Exchange Online and Entra ID management functions, providing a complete overview of the Microsoft 365 environment configuration and security posture.`n`n"
    }
    
    $note += "## Exchange Online Configuration`n`n"
    $note += "### Connection Status`n"
    $note += "- Exchange Online: $(if ($script:currentExchangeConnection) { 'Connected' } else { 'Not Connected' })`n"
    $note += "- Mailboxes Loaded: $(if ($script:allLoadedMailboxUPNs) { $script:allLoadedMailboxUPNs.Count } else { '0' })`n`n"
    
    # Mailbox Analysis
    if ($userMailboxGrid.Rows.Count -gt 0) {
        $selectedCount = 0
        $totalRules = 0
        $suspiciousRules = 0
        $externalForwarding = 0
        
        for ($i = 0; $i -lt $userMailboxGrid.Rows.Count; $i++) {
            if ($userMailboxGrid.Rows[$i].Cells["Select"].Value -eq $true) { 
                $selectedCount++ 
                $rulesCount = [int]$userMailboxGrid.Rows[$i].Cells["RulesCount"].Value
                $totalRules += $rulesCount
                if ($rulesCount -gt 0) {
                    $suspiciousRules += [int]$userMailboxGrid.Rows[$i].Cells["SuspiciousRules"].Value
                    if ($userMailboxGrid.Rows[$i].Cells["ExternalForwarding"].Value -eq "Yes") {
                        $externalForwarding++
                    }
                }
            }
        }
        
        $note += "### Mailbox Inbox Rules Analysis`n"
        $note += "- Mailboxes Selected for Analysis: $selectedCount`n"
        $note += "- Total Inbox Rules Found: $totalRules`n"
        $note += "- Suspicious Rules Detected: $suspiciousRules`n"
        $note += "- Mailboxes with External Forwarding: $externalForwarding`n`n"
        
        $note += "### Detailed Mailbox Analysis`n"
        for ($i = 0; $i -lt $userMailboxGrid.Rows.Count; $i++) {
            if ($userMailboxGrid.Rows[$i].Cells["Select"].Value -eq $true) {
                $upn = $userMailboxGrid.Rows[$i].Cells["UserPrincipalName"].Value
                $rulesCount = $userMailboxGrid.Rows[$i].Cells["RulesCount"].Value
                $suspiciousRules = $userMailboxGrid.Rows[$i].Cells["SuspiciousRules"].Value
                $externalForwarding = $userMailboxGrid.Rows[$i].Cells["ExternalForwarding"].Value
                $delegates = $userMailboxGrid.Rows[$i].Cells["Delegates"].Value
                $fullAccess = $userMailboxGrid.Rows[$i].Cells["FullAccess"].Value
                
                $note += "- **$upn**`n"
                $note += "  - Total Rules: $rulesCount`n"
                $note += "  - Suspicious Rules: $suspiciousRules`n"
                $note += "  - External Forwarding: $externalForwarding`n"
                $note += "  - Delegates: $delegates`n"
                $note += "  - Full Access Users: $fullAccess`n`n"
            }
        }
    } else {
        $note += "### Mailbox Inbox Rules Analysis`n"
        $note += "- No mailboxes selected for analysis`n`n"
    }
    
    # Transport Rules
    $note += "### Transport Rules Configuration`n"
    try {
        $transportRules = Get-TransportRule -ErrorAction SilentlyContinue | Select-Object Name, State, Priority, Enabled
        if ($transportRules) {
            $note += "- Total Transport Rules: $($transportRules.Count)`n"
            $note += "- Active Rules: $(($transportRules | Where-Object { $_.State -eq 'Enabled' }).Count)`n"
            $note += "- Inactive Rules: $(($transportRules | Where-Object { $_.State -eq 'Disabled' }).Count)`n`n"
            
            $note += "#### Transport Rules Details`n"
            foreach ($rule in $transportRules | Select-Object -First 10) {
                $note += "- **$($rule.Name)** (Priority: $($rule.Priority), State: $($rule.State))`n"
            }
            if ($transportRules.Count -gt 10) {
                $note += "- ... and $($transportRules.Count - 10) more rules`n"
            }
            $note += "`n"
        } else {
            $note += "- No transport rules found or access denied`n`n"
        }
    } catch {
        $note += "- Transport rules data unavailable: $($_.Exception.Message)`n`n"
    }
    
    # Connectors
    $note += "### Connectors Configuration`n"
    try {
        # Try different connector cmdlets that might be available
        $connectors = $null
        
        # First try Get-Connector (Exchange Online)
        try {
            $connectors = Get-Connector -ErrorAction Stop | Select-Object Name, ConnectorType, Enabled
        } catch {
            # Try Get-InboundConnector (Exchange Online)
            try {
                $inboundConnectors = Get-InboundConnector -ErrorAction Stop | Select-Object Name, ConnectorType, Enabled
                $outboundConnectors = Get-OutboundConnector -ErrorAction Stop | Select-Object Name, ConnectorType, Enabled
                $connectors = @($inboundConnectors) + @($outboundConnectors)
            } catch {
                # Try Get-HostedConnector (Exchange Online)
                try {
                    $connectors = Get-HostedConnector -ErrorAction Stop | Select-Object Name, ConnectorType, Enabled
                } catch {
                    $connectors = $null
                }
            }
        }
        
        if ($connectors -and $connectors.Count -gt 0) {
            $note += "- Total Connectors: $($connectors.Count)`n"
            $note += "- Enabled Connectors: $(($connectors | Where-Object { $_.Enabled -eq $true }).Count)`n"
            $note += "- Disabled Connectors: $(($connectors | Where-Object { $_.Enabled -eq $false }).Count)`n`n"
            
            $note += "#### Connectors Details`n"
            foreach ($connector in $connectors | Select-Object -First 10) {
                $note += "- **$($connector.Name)** (Type: $($connector.ConnectorType), Enabled: $($connector.Enabled))`n"
            }
            if ($connectors.Count -gt 10) {
                $note += "- ... and $($connectors.Count - 10) more connectors`n"
            }
            $note += "`n"
        } else {
            $note += "- No connectors found or access denied`n`n"
        }
    } catch {
        $note += "- Connectors data unavailable: $($_.Exception.Message)`n`n"
    }
    
    # Entra ID Section
    $note += "## Entra ID (Azure AD) Configuration`n`n"
    $note += "### Connection Status`n"
    $note += "- Entra ID: $(if ($script:graphConnection) { 'Connected' } else { 'Not Connected' })`n"
    $note += "- Users Loaded: $(if ($entraUserGrid.Rows.Count -gt 0) { $entraUserGrid.Rows.Count } else { '0' })`n`n"
    
    # User Analysis
    if ($entraUserGrid.Rows.Count -gt 0) {
        $selectedCount = 0
        $licensedUsers = 0
        $unlicensedUsers = 0
        
        for ($i = 0; $i -lt $entraUserGrid.Rows.Count; $i++) {
            if ($entraUserGrid.Rows[$i].Cells["Select"].Value -eq $true) { 
                $selectedCount++ 
                if ($entraUserGrid.Rows[$i].Cells["Licensed"].Value -eq "Yes") {
                    $licensedUsers++
                } else {
                    $unlicensedUsers++
                }
            }
        }
        
        $note += "### User Account Analysis`n"
        $note += "- Users Selected for Analysis: $selectedCount`n"
        $note += "- Licensed Users: $licensedUsers`n"
        $note += "- Unlicensed Users: $unlicensedUsers`n`n"
        
        $note += "### Selected User Details`n"
        for ($i = 0; $i -lt $entraUserGrid.Rows.Count; $i++) {
            if ($entraUserGrid.Rows[$i].Cells["Select"].Value -eq $true) {
                $upn = $entraUserGrid.Rows[$i].Cells["UserPrincipalName"].Value
                $displayName = $entraUserGrid.Rows[$i].Cells["DisplayName"].Value
                $licensed = $entraUserGrid.Rows[$i].Cells["Licensed"].Value
                
                $note += "- **$displayName** ($upn)`n"
                $note += "  - Licensed: $licensed`n`n"
            }
        }
    } else {
        $note += "### User Account Analysis`n"
        $note += "- No users selected for analysis`n`n"
    }
    
    # Sign-in Logs
    $note += "### Sign-in Logs Summary`n"
    try {
        # Get selected users for sign-in logs
        $selectedUsers = @()
        foreach ($account in $selectedAccounts) {
            if ($account.EntraStatus -eq "Available") {
                if (-not [string]::IsNullOrWhiteSpace($account.UserPrincipalName)) {
                    $selectedUsers += $account.UserPrincipalName
                }
            }
        }
        
        if ($selectedUsers.Count -gt 0) {
            $signInLogs = Get-EntraSignInLogs -UserPrincipalNames $selectedUsers -Days 7 -ErrorAction SilentlyContinue
            if ($signInLogs -and $signInLogs.Count -gt 0) {
                $recentLogs = $signInLogs | Select-Object -First 50
                $successfulLogins = ($recentLogs | Where-Object { $_.Status -eq "Success" }).Count
                $failedLogins = ($recentLogs | Where-Object { $_.Status -eq "Failure" }).Count
                $suspiciousLogins = ($recentLogs | Where-Object { $_.RiskLevel -eq "High" -or $_.RiskLevel -eq "Medium" }).Count
                
                # Analyze non-US sign-ins
                $nonUSSignIns = @()
                $usSignIns = @()
                foreach ($log in $recentLogs) {
                    if ($log.Location -and $log.Location.CountryOrRegion) {
                        if ($log.Location.CountryOrRegion -ne "US" -and $log.Location.CountryOrRegion -ne "United States") {
                            $nonUSSignIns += $log
                        } else {
                            $usSignIns += $log
                        }
                    }
                }
                
                $note += "- Recent Sign-in Activity (Last 50 events)`n"
                $note += "- Total Events: $($recentLogs.Count)`n"
                $note += "- Successful Logins: $successfulLogins`n"
                $note += "- Failed Logins: $failedLogins`n"
                $note += "- Suspicious Logins: $suspiciousLogins`n"
                $note += "- US Sign-ins: $($usSignIns.Count)`n"
                $note += "- Non-US Sign-ins: $($nonUSSignIns.Count)`n`n"
                
                $note += "#### Recent Sign-in Events`n"
                foreach ($log in $recentLogs | Select-Object -First 10) {
                    $location = if ($log.Location -and $log.Location.CountryOrRegion) { $log.Location.CountryOrRegion } else { "Unknown" }
                    $note += "- **$($log.UserPrincipalName)** - $($log.CreatedDateTime) - Status: $($log.Status) - Risk: $($log.RiskLevel) - Location: $location`n"
                }
                if ($recentLogs.Count -gt 10) {
                    $note += "- ... and $($recentLogs.Count - 10) more events`n"
                }
                $note += "`n"
                
                # Show non-US sign-ins if any found
                if ($nonUSSignIns.Count -gt 0) {
                    $note += "#### Non-US Sign-in Events (Security Alert)`n"
                    foreach ($log in $nonUSSignIns | Select-Object -First 5) {
                        $location = if ($log.Location -and $log.Location.CountryOrRegion) { $log.Location.CountryOrRegion } else { "Unknown" }
                        $city = if ($log.Location -and $log.Location.City) { $log.Location.City } else { "Unknown" }
                        $note += "- **$($log.UserPrincipalName)** - $($log.CreatedDateTime) - Status: $($log.Status) - Risk: $($log.RiskLevel) - Location: $city, $location`n"
                    }
                    if ($nonUSSignIns.Count -gt 5) {
                        $note += "- ... and $($nonUSSignIns.Count - 5) more non-US events`n"
                    }
                    $note += "`n"
                }
            } else {
                $note += "- No sign-in logs available for selected users`n`n"
            }
        } else {
            $note += "- No users selected for sign-in log analysis`n`n"
        }
    } catch {
        $note += "- Sign-in logs data unavailable: $($_.Exception.Message)`n`n"
    }
    
    # Audit Logs
    $note += "### Audit Logs Summary`n"
    try {
        # Get selected users for audit logs
        $selectedUsers = @()
        foreach ($account in $selectedAccounts) {
            if ($account.EntraStatus -eq "Available") {
                if (-not [string]::IsNullOrWhiteSpace($account.UserPrincipalName)) {
                    $selectedUsers += $account.UserPrincipalName
                }
            }
        }
        
        if ($selectedUsers.Count -gt 0) {
            $auditLogs = Get-EntraUserAuditLogs -UserPrincipalName $selectedUsers[0] -Days 7 -ErrorAction SilentlyContinue
            if ($auditLogs -and $auditLogs.Count -gt 0) {
                $recentAudits = $auditLogs | Select-Object -First 50
                $adminActions = ($recentAudits | Where-Object { $_.Category -eq "AdministrativeUnit" }).Count
                $userManagement = ($recentAudits | Where-Object { $_.Category -eq "UserManagement" }).Count
                $applicationActivity = ($recentAudits | Where-Object { $_.Category -eq "Application" }).Count
                
                $note += "- Recent Audit Activity (Last 50 events)`n"
                $note += "- Total Events: $($recentAudits.Count)`n"
                $note += "- Administrative Actions: $adminActions`n"
                $note += "- User Management Events: $userManagement`n"
                $note += "- Application Activity: $applicationActivity`n`n"
                
                $note += "#### Recent Audit Events`n"
                foreach ($log in $recentAudits | Select-Object -First 10) {
                    $note += "- **$($log.UserPrincipalName)** - $($log.CreatedDateTime) - Category: $($log.Category) - Activity: $($log.Activity)`n"
                }
                if ($recentAudits.Count -gt 10) {
                    $note += "- ... and $($recentAudits.Count - 10) more events`n"
                }
                $note += "`n"
            } else {
                $note += "- No audit logs available for selected users`n`n"
            }
        } else {
            $note += "- No users selected for audit log analysis`n`n"
        }
    } catch {
        $note += "- Audit logs data unavailable: $($_.Exception.Message)`n`n"
    }
    
    # Security Assessment
    $note += "## Security Posture Assessment`n`n"
    
    $note += "### Exchange Online Security Findings`n"
    if ($userMailboxGrid.Rows.Count -gt 0) {
        $selectedCount = 0
        $totalSuspiciousRules = 0
        $externalForwardingCount = 0
        
        for ($i = 0; $i -lt $userMailboxGrid.Rows.Count; $i++) {
            if ($userMailboxGrid.Rows[$i].Cells["Select"].Value -eq $true) { 
                $selectedCount++ 
                $totalSuspiciousRules += [int]$userMailboxGrid.Rows[$i].Cells["SuspiciousRules"].Value
                if ($userMailboxGrid.Rows[$i].Cells["ExternalForwarding"].Value -eq "Yes") {
                    $externalForwardingCount++
                }
            }
        }
        
        $note += "- Mailboxes Analyzed: $selectedCount`n"
        $note += "- Total Suspicious Rules Found: $totalSuspiciousRules`n"
        $note += "- Mailboxes with External Forwarding: $externalForwardingCount`n"
        $riskLevel = if ($totalSuspiciousRules -gt 0 -or $externalForwardingCount -gt 0) { "HIGH - Immediate attention required" } else { "LOW - No immediate concerns detected" }
        $note += "- Risk Level: $riskLevel`n`n"
    } else {
        $note += "- No mailboxes analyzed`n`n"
    }
    
    $note += "### Entra ID Security Findings`n"
    if ($entraUserGrid.Rows.Count -gt 0) {
        $selectedCount = 0
        $unlicensedUsers = 0
        
        for ($i = 0; $i -lt $entraUserGrid.Rows.Count; $i++) {
            if ($entraUserGrid.Rows[$i].Cells["Select"].Value -eq $true) { 
                $selectedCount++ 
                if ($entraUserGrid.Rows[$i].Cells["Licensed"].Value -ne "Yes") {
                    $unlicensedUsers++
                }
            }
        }
        
        $note += "- Users Analyzed: $selectedCount`n"
        $note += "- Unlicensed Users: $unlicensedUsers`n"
        $note += "- MFA Status: Available for individual analysis`n"
        $note += "- Session Management: Available for revocation`n`n"
    } else {
        $note += "- No users analyzed`n`n"
    }
    

    
    # Technical Details
    $note += "## Technical Details`n`n"
    $note += "### Environment Information`n"
    $note += "- Tool Version: 8.0`n"
    $note += "- Report Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`n"
    $note += "- Exchange Connection: $(if ($script:currentExchangeConnection) { 'Active' } else { 'Inactive' })`n"
    $note += "- Graph Connection: $(if ($script:graphConnection) { 'Active' } else { 'Inactive' })`n`n"
    
    $note += "### Data Sources`n"
    $note += "- Exchange Online PowerShell (Inbox Rules, Transport Rules, Connectors)`n"
    $note += "- Microsoft Graph API (Users, Sign-in Logs, Audit Logs)`n"
    $note += "- Real-time mailbox analysis`n"
    $note += "- Security posture assessment`n`n"
    
    $note += "---`n"
    $note += "Tags: #microsoft365 #security #exchange #entra #comprehensive-analysis"

    return $note
}

# Function to generate incident remediation checklist with enhanced data
function Generate-IncidentRemediationChecklist {
    param($selectedAccounts)
    
    # Get the first selected user for single-user focus
    $firstSelectedUser = $selectedAccounts | Where-Object { $_.EntraStatus -eq "Available" } | Select-Object -First 1
    
    if (-not $firstSelectedUser) {
        return "No user account selected for incident remediation analysis."
    }
    
    # Get additional data from script functions
    $transportRules = $null
    $connectors = $null
    $signInLogs = $null
    $auditLogs = $null
    
    try {
        # Get transport rules data
        $transportRules = Get-TransportRule -ErrorAction SilentlyContinue | Select-Object Name, State, Priority, Enabled
    } catch { }
    
    try {
        # Get connectors data
        $connectors = Get-Connector -ErrorAction SilentlyContinue | Select-Object Name, ConnectorType, Enabled
    } catch { }
    
    try {
        # Get sign-in logs for the user
        $signInLogs = Get-EntraSignInLogs -UserPrincipalNames @($firstSelectedUser.UserPrincipalName) -Days 7 -ErrorAction SilentlyContinue
    } catch { }
    
    try {
        # Get audit logs for the user
        $auditLogs = Get-EntraUserAuditLogs -UserPrincipalName $firstSelectedUser.UserPrincipalName -Days 7 -ErrorAction SilentlyContinue
    } catch { }
    
    $checklist = "The Essential Office 365 Account Incident Remediation Checklist`n"
    $checklist += "Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`n"
    $checklist += "User Account: $($firstSelectedUser.DisplayName)`n"
    $checklist += "User Principal Name: $($firstSelectedUser.UserPrincipalName)`n`n"
    
    # Checklist items with enhanced data analysis
    $checklist += "☐ Reset the Users Password in Active Directory or Office 365 if the account is a cloud-only account.`n"
    $checklist += "   Current Status: $(if ($firstSelectedUser.Licensed -eq "Yes") { "Licensed cloud account - Password reset required" } else { "Unlicensed account - Verify account status" })`n"
    $checklist += "   Date completed:			Technician: `n`n"
    
    $checklist += "☐ Recommend Multi-Factor Authentication (MFA) to the client`n"
    $checklist += "   Current Status: MFA status available for individual analysis`n"
    $checklist += "   Date completed:			Technician: `n`n"
    
    $checklist += "☐ Apply the Require user to sign in again via Cloud App Security (if available)`n"
    $checklist += "   Current Status: Session revocation available in Entra ID tab`n"
    $checklist += "   Date completed:			Technician: `n`n"
    
    $checklist += "☐ Force User Sign-out from Microsoft 365 Admin Panel`n"
    $checklist += "   Current Status: Session management available in Entra ID tab`n"
    $checklist += "   Date completed:			Technician: `n`n"
    
    $checklist += "☐ Review the mailbox for any mailbox delegates and remove from the compromised account`n"
    $checklist += "   Current Status: Delegates found: $($firstSelectedUser.Delegates)`n"
    $checklist += "   Date completed:			Technician: `n`n"
    
    $checklist += "☐ Review the mailbox for any mail forwarding rules that may have been created`n"
    $checklist += "   Current Status: External forwarding: $($firstSelectedUser.ExternalForwarding)`n"
    $checklist += "   Date completed:			Technician: `n`n"
    
    $checklist += "☐ Review the mailbox inbox rules and delete any suspicious ones.`n"
    $checklist += "   Current Status: Total rules: $($firstSelectedUser.RulesCount), Suspicious rules: $($firstSelectedUser.SuspiciousRules)`n"
    $checklist += "   Date completed:			Technician: `n`n"
    
    $checklist += "☐ Educate the user about security threats and methods used to gain access to users' credentials`n"
    $checklist += "   Current Status: User education required`n"
    $checklist += "   Date completed:			Technician: `n`n"
    
    $checklist += "☐ Run a mail trace to identify suspicious messages sent or received by this account`n"
    $checklist += "   Current Status: Mail trace analysis required`n"
    $checklist += "   Date completed:			Technician: `n`n"
    
    $checklist += "☐ Search the audit log to identify suspicious logins, attempt to identify the earliest date and time the account was compromised, and confirm no suspicious logins occur after password reset`n"
    $recentSignIns = if ($signInLogs -and $signInLogs.Count -gt 0) { $signInLogs | Select-Object -First 5 } else { $null }
    $suspiciousSignIns = if ($recentSignIns) { ($recentSignIns | Where-Object { $_.RiskLevel -eq "High" -or $_.RiskLevel -eq "Medium" }).Count } else { 0 }
    $checklist += "   Current Status: Sign-in logs available for selected user`n"
    $checklist += "   Recent Sign-ins: $(if ($recentSignIns) { $recentSignIns.Count } else { "None available" })`n"
    $checklist += "   Suspicious Sign-ins: $suspiciousSignIns`n"
    $checklist += "   Date completed:			Technician: `n`n"
    
    $checklist += "☐ Advise the user that if the password that was in use is also used on any other accounts, those passwords should also be changed immediately`n"
    $checklist += "   Current Status: Password security advisory required`n"
    $checklist += "   Date completed:			Technician: `n`n"
    
    $checklist += "☐ Review the list of Administrators/Global Administrators in the Administration console. Check this against the users who SHOULD be Admins/Global Admins`n"
    $checklist += "   Current Status: Admin review required`n"
    $checklist += "   Date completed:			Technician: `n`n"
    
    $checklist += "☐ Review the Global/Domain Transport rules to ensure no rules have been set up.`n"
    $activeTransportRules = if ($transportRules) { ($transportRules | Where-Object { $_.State -eq "Enabled" }).Count } else { "Unknown" }
    $checklist += "   Current Status: Transport rules analysis available`n"
    $checklist += "   Active Transport Rules: $activeTransportRules`n"
    $checklist += "   Date completed:			Technician: `n`n"
    
    $checklist += "☐ Review the list of licensed O365 Users. Check this against the list of users who SHOULD be in O365. Ensure that no disabled users or terminated users have a valid license assigned.`n"
    $checklist += "   Current Status: User licensing review required`n"
    $checklist += "   Date completed:			Technician: `n`n"
    
    $checklist += "`nAdditional notes (If Needed):`n"
    $checklist += "`n"
    $checklist += "Technical Analysis Summary:`n"
    $checklist += "- Exchange Status: $($firstSelectedUser.ExchangeStatus)`n"
    $checklist += "- Entra ID Status: $($firstSelectedUser.EntraStatus)`n"
    $checklist += "- Full Access Users: $($firstSelectedUser.FullAccess)`n"
    $checklist += "- Account Licensed: $($firstSelectedUser.Licensed)`n"
    $checklist += "- Total Transport Rules: $(if ($transportRules) { $transportRules.Count } else { "Unknown" })`n"
    $checklist += "- Active Connectors: $(if ($connectors) { ($connectors | Where-Object { $_.Enabled -eq $true }).Count } else { "Unknown" })`n"
    $checklist += "- Recent Sign-in Events: $(if ($signInLogs) { $signInLogs.Count } else { "None available" })`n"
    $checklist += "- Recent Audit Events: $(if ($auditLogs) { $auditLogs.Count } else { "None available" })`n"
    
    return $checklist
}

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
    @{Name="Microsoft.Graph.Users.Actions"; MinVersion="2.0"},
    @{Name="Microsoft.Graph.Identity.SignIns"; MinVersion="2.0"},
    @{Name="Microsoft.Graph.Reports"; MinVersion="2.0"}
)
$script:graphScopes = @(
    "User.Read.All",
    "User.ReadWrite.All",
    "SecurityEvents.Read.All",
    "SecurityEvents.ReadWrite.All",
    "SecurityAlert.Read.All",
    "SecurityAlert.ReadWrite.All",
    "SecurityIncident.Read.All",
    "SecurityIncident.ReadWrite.All",
    "ThreatIntelligence.Read.All",
    "ThreatIntelligence.ReadWrite.All",
    "AuditLog.Read.All",
    "Directory.Read.All"
)

# --- GUI Setup ---
Add-Type -AssemblyName System.Windows.Forms; Add-Type -AssemblyName System.Drawing
$mainForm = New-Object System.Windows.Forms.Form; $mainForm.Text = "Microsoft 365 Management Tool"; $mainForm.Size = New-Object System.Drawing.Size(1400, 900); $mainForm.MinimumSize = New-Object System.Drawing.Size(1200, 700); $mainForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen; $mainForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable; $mainForm.MaximizeBox = $true; $mainForm.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi
$statusStrip = New-Object System.Windows.Forms.StatusStrip
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Name = "statusLabel"
$statusLabel.Text = "Ready. Connect to Exchange Online."
$statusStrip.Items.Add($statusLabel)

# Add progress bar to status strip
$progressBar = New-Object System.Windows.Forms.ToolStripProgressBar
$progressBar.Name = "progressBar"
$progressBar.Visible = $false
$progressBar.Width = 200
$statusStrip.Items.Add($progressBar)

$mainForm.Controls.Add($statusStrip)

# --- Main TabControl (fills the form) ---
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Dock = 'Fill'
$mainForm.Controls.Add($tabControl)

# --- AI Analysis Tab ---
$aiTab = New-Object System.Windows.Forms.TabPage
$aiTab.Text = "AI Analysis"
$aiPanel = New-Object System.Windows.Forms.Panel
$aiPanel.Dock = 'Fill'
$aiPanel.Padding = New-Object System.Windows.Forms.Padding(10)

# Title and description
$aiTitle = New-Object System.Windows.Forms.Label
$aiTitle.Text = "AI Analysis"
$aiTitle.Font = New-Object System.Drawing.Font('Segoe UI', 12, [System.Drawing.FontStyle]::Bold)
$aiTitle.Location = New-Object System.Drawing.Point(10,10)
$aiTitle.AutoSize = $true

$aiDesc = New-Object System.Windows.Forms.Label
$aiDesc.Text = "Send the latest or selected investigation dataset to Gemini or Claude for analysis. Configure API keys in Settings."
$aiDesc.Location = New-Object System.Drawing.Point(10,35)
$aiDesc.Size = New-Object System.Drawing.Size(740, 30)

# Folder selection
$aiProviderLabel = New-Object System.Windows.Forms.Label
$aiProviderLabel.Text = "Provider:"
$aiProviderLabel.Location = New-Object System.Drawing.Point(10,65)
$aiProviderLabel.AutoSize = $true

$aiProviderCombo = New-Object System.Windows.Forms.ComboBox
$aiProviderCombo.Location = New-Object System.Drawing.Point(100, 62)
$aiProviderCombo.Width = 140
$aiProviderCombo.DropDownStyle = 'DropDownList'
$aiProviderCombo.Items.AddRange(@('Gemini','Claude'))
$aiProviderCombo.SelectedIndex = 0

$aiFolderLabel = New-Object System.Windows.Forms.Label
$aiFolderLabel.Text = "Report Folder:"
$aiFolderLabel.Location = New-Object System.Drawing.Point(250,65)
$aiFolderLabel.AutoSize = $true

$aiFolderText = New-Object System.Windows.Forms.TextBox
$aiFolderText.Location = New-Object System.Drawing.Point(340, 62)
$aiFolderText.Width = 380

$aiBrowseBtn = New-Object System.Windows.Forms.Button
$aiBrowseBtn.Text = "Browse..."
$aiBrowseBtn.Location = New-Object System.Drawing.Point(730, 60)
$aiBrowseBtn.Size = New-Object System.Drawing.Size(85, 24)
$aiBrowseBtn.add_Click({
    $fbd = New-Object System.Windows.Forms.FolderBrowserDialog
    $fbd.Description = "Select the report folder that contains LLM_Instructions.txt and CSV files"
    if ($aiFolderText.Text -and (Test-Path $aiFolderText.Text)) { $fbd.SelectedPath = $aiFolderText.Text }
    if ($fbd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $aiFolderText.Text = $fbd.SelectedPath }
})

# Extra files list
$aiExtraLabel = New-Object System.Windows.Forms.Label
$aiExtraLabel.Text = "Extra Files (optional):"
$aiExtraLabel.Location = New-Object System.Drawing.Point(10,110)
$aiExtraLabel.AutoSize = $true

$aiExtraList = New-Object System.Windows.Forms.ListBox
$aiExtraList.Location = New-Object System.Drawing.Point(10,130)
$aiExtraList.Size = New-Object System.Drawing.Size(610, 120)

$aiAddExtraBtn = New-Object System.Windows.Forms.Button
$aiAddExtraBtn.Text = "Add..."
$aiAddExtraBtn.Location = New-Object System.Drawing.Point(630, 130)
$aiAddExtraBtn.Size = New-Object System.Drawing.Size(85, 24)
$aiAddExtraBtn.add_Click({
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Title = "Select additional file(s) to include"
    $ofd.Filter = "CSV/Text|*.csv;*.txt|All Files|*.*"
    $ofd.Multiselect = $true
    if ($ofd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        foreach ($p in $ofd.FileNames) {
            if (-not ($aiExtraList.Items -contains $p)) { [void]$aiExtraList.Items.Add($p) }
        }
    }
})

$aiRemoveExtraBtn = New-Object System.Windows.Forms.Button
$aiRemoveExtraBtn.Text = "Remove"
$aiRemoveExtraBtn.Location = New-Object System.Drawing.Point(630, 160)
$aiRemoveExtraBtn.Size = New-Object System.Drawing.Size(85, 24)
$aiRemoveExtraBtn.add_Click({
    $sel = @($aiExtraList.SelectedItems)
    foreach ($it in $sel) { $aiExtraList.Items.Remove($it) }
})

# Send button and status
$aiSendBtn = New-Object System.Windows.Forms.Button
$aiSendBtn.Text = "Send to AI"
$aiSendBtn.Location = New-Object System.Drawing.Point(10, 265)
$aiSendBtn.Size = New-Object System.Drawing.Size(140, 30)

$aiStatus = New-Object System.Windows.Forms.Label
$aiStatus.Location = New-Object System.Drawing.Point(160, 270)
$aiStatus.Size = New-Object System.Drawing.Size(555, 20)
$aiStatus.ForeColor = [System.Drawing.Color]::FromArgb(80,80,80)

$aiPanel.Controls.AddRange(@($aiTitle,$aiDesc,$aiProviderLabel,$aiProviderCombo,$aiFolderLabel,$aiFolderText,$aiBrowseBtn,$aiExtraLabel,$aiExtraList,$aiAddExtraBtn,$aiRemoveExtraBtn,$aiSendBtn,$aiStatus))
$aiTab.Controls.Add($aiPanel)
$tabControl.TabPages.Add($aiTab)

# Helper: get latest report folder (nested or legacy)
$getLatestReportFolder = {
    try {
        $base = Join-Path $env:USERPROFILE "Documents\ExchangeOnlineAnalyzer\SecurityInvestigation"
        if (-not (Test-Path $base)) { return $null }
        $candidates = @()
        $tenants = Get-ChildItem -Path $base -Directory -ErrorAction SilentlyContinue
        foreach ($t in $tenants) {
            $runs = Get-ChildItem -Path $t.FullName -Directory -ErrorAction SilentlyContinue
            if ($runs) { $candidates += $runs }
        }
        $legacy = Get-ChildItem -Path $base -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -match '^\d{8}_\d{6}$' }
        if ($legacy) { $candidates += $legacy }
        if ($candidates -and $candidates.Count -gt 0) { return ($candidates | Sort-Object LastWriteTime -Descending | Select-Object -First 1).FullName }
    } catch {}
    return $null
}

# Prefill latest folder when the tab is entered
$aiTab.add_Enter({
    try {
        if (-not $aiFolderText.Text -or -not (Test-Path $aiFolderText.Text)) {
            $latest = & $getLatestReportFolder
            if ($latest) { $aiFolderText.Text = $latest }
        }
    } catch {}
})

# Send to Gemini handler
$aiSendBtn.add_Click({
    try {
        $folder = $aiFolderText.Text
        if (-not $folder -or -not (Test-Path $folder)) { $aiStatus.Text = "Select a valid report folder."; $aiStatus.ForeColor = [System.Drawing.Color]::Red; return }
        $provider = $aiProviderCombo.SelectedItem
        if ($provider -eq 'Gemini') {
            $scriptPath = Join-Path $PSScriptRoot "Scripts\Send-To-Gemini.ps1"
            if (-not (Test-Path $scriptPath)) { $aiStatus.Text = "Gemini sender script not found."; $aiStatus.ForeColor = [System.Drawing.Color]::Red; return }
        } else {
            $scriptPath = Join-Path $PSScriptRoot "Scripts\Send-To-Claude.ps1"
            if (-not (Test-Path $scriptPath)) { $aiStatus.Text = "Claude sender script not found."; $aiStatus.ForeColor = [System.Drawing.Color]::Red; return }
        }

        $extras = @(); foreach ($it in $aiExtraList.Items) { $extras += $it }

        $aiSendBtn.Enabled = $false; $aiStatus.Text = ("Submitting to {0}..." -f $provider); $aiStatus.ForeColor = [System.Drawing.Color]::FromArgb(80,80,80)
        $mainForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        $output = $null
        try {
            if ($provider -eq 'Gemini') {
                if ($extras.Count -gt 0) {
                    $ps = { param($sp,$of,$ef) & $sp -OutputFolder $of -ExtraFiles $ef -Verbose 4>&1 }
                    $output = & $ps $scriptPath $folder $extras
                } else {
                    $ps = { param($sp,$of) & $sp -OutputFolder $of -Verbose 4>&1 }
                    $output = & $ps $scriptPath $folder
                }
            } else {
                # Ensure Claude API key exists
                try {
                    Import-Module "$PSScriptRoot\Modules\Settings.psm1" -Force -ErrorAction SilentlyContinue
                    $s = Get-AppSettings
                    if (-not $s -or -not $s.ClaudeApiKey -or $s.ClaudeApiKey.Trim().Length -eq 0) {
                        [System.Windows.Forms.MessageBox]::Show("Please add your Claude API key in the Settings tab first.", "Claude API Key Required", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                        return
                    }
                } catch {}
                if ($extras.Count -gt 0) {
                    $ps = { param($sp,$of,$ef) & $sp -OutputFolder $of -ExtraFiles $ef -MaxCsvRows 2000 -VerboseOutput 4>&1 }
                    $output = & $ps $scriptPath $folder $extras
                } else {
                    $ps = { param($sp,$of) & $sp -OutputFolder $of -MaxCsvRows 2000 -VerboseOutput 4>&1 }
                    $output = & $ps $scriptPath $folder
                }
            }
        } catch {
            $output = $_.Exception.Message
        } finally {
            $mainForm.Cursor = [System.Windows.Forms.Cursors]::Default
            $aiSendBtn.Enabled = $true
        }

        $respFile = if ($provider -eq 'Gemini') { Join-Path $folder "Gemini_Response.md" } else { Join-Path $folder "Claude_Response.md" }
        $errFile  = if ($provider -eq 'Gemini') { Join-Path $folder "Gemini_Error.txt" } else { Join-Path $folder "Claude_Error.txt" }
        if (Test-Path $respFile) {
            $aiStatus.Text = ("Saved: {0}" -f $respFile); $aiStatus.ForeColor = [System.Drawing.Color]::Green
            try { Start-Process $respFile } catch {}
        } elseif (Test-Path $errFile) {
            $aiStatus.Text = ("{0} error. See: {1}" -f $provider, $errFile); $aiStatus.ForeColor = [System.Drawing.Color]::Red
            try { Start-Process $errFile } catch {}
        } else {
            $aiStatus.Text = ("Completed. Check folder for {0} (see console for details)." -f [System.IO.Path]::GetFileName($respFile));
            $aiStatus.ForeColor = [System.Drawing.Color]::FromArgb(80,80,80)
            if ($output) { [System.Windows.Forms.MessageBox]::Show(("Output:`n{0}" -f ($output | Out-String)), "Send to AI", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) }
        }
    } catch {
        $aiStatus.Text = $_.Exception.Message; $aiStatus.ForeColor = [System.Drawing.Color]::Red
    }
})
# --- Settings Tab ---
try { Import-Module "$PSScriptRoot\Modules\Settings.psm1" -Force -ErrorAction SilentlyContinue } catch {}
$settingsTab = New-Object System.Windows.Forms.TabPage
$settingsTab.Text = "Settings"
$settingsPanel = New-Object System.Windows.Forms.Panel
$settingsPanel.Dock = 'Fill'
$settingsPanel.Padding = New-Object System.Windows.Forms.Padding(10)

$sTitle = New-Object System.Windows.Forms.Label
$sTitle.Text = "Application Settings"
$sTitle.Font = New-Object System.Drawing.Font('Segoe UI', 12, [System.Drawing.FontStyle]::Bold)
$sTitle.Location = New-Object System.Drawing.Point(10,10)
$sTitle.AutoSize = $true

$lblInv = New-Object System.Windows.Forms.Label
$lblInv.Text = "Investigator Name:"
$lblInv.Location = New-Object System.Drawing.Point(10,45)
$lblInv.AutoSize = $true

$txtInv = New-Object System.Windows.Forms.TextBox
$txtInv.Location = New-Object System.Drawing.Point(150, 42)
$txtInv.Width = 300

$lblCo = New-Object System.Windows.Forms.Label
$lblCo.Text = "Company Name:"
$lblCo.Location = New-Object System.Drawing.Point(10,75)
$lblCo.AutoSize = $true

$txtCo = New-Object System.Windows.Forms.TextBox
$txtCo.Location = New-Object System.Drawing.Point(150, 72)
$txtCo.Width = 300

$lblGem = New-Object System.Windows.Forms.Label
$lblGem.Text = "Gemini API Key:"
$lblGem.Location = New-Object System.Drawing.Point(10,105)
$lblGem.AutoSize = $true

$txtGem = New-Object System.Windows.Forms.TextBox
$txtGem.Location = New-Object System.Drawing.Point(150, 102)
$txtGem.Width = 300
$txtGem.UseSystemPasswordChar = $true

$lblClaude = New-Object System.Windows.Forms.Label
$lblClaude.Text = "Claude API Key:"
$lblClaude.Location = New-Object System.Drawing.Point(10,135)
$lblClaude.AutoSize = $true

$txtClaude = New-Object System.Windows.Forms.TextBox
$txtClaude.Location = New-Object System.Drawing.Point(150, 132)
$txtClaude.Width = 300
$txtClaude.UseSystemPasswordChar = $true

$btnSave = New-Object System.Windows.Forms.Button
$btnSave.Text = "Save"
$btnSave.Location = New-Object System.Drawing.Point(150, 165)
$btnSave.Width = 100

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Location = New-Object System.Drawing.Point(10,200)
$lblStatus.AutoSize = $true
$lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(80,80,80)

$settingsPanel.Controls.AddRange(@($sTitle,$lblInv,$txtInv,$lblCo,$txtCo,$lblGem,$txtGem,$lblClaude,$txtClaude,$btnSave,$lblStatus))
$settingsTab.Controls.Add($settingsPanel)
$tabControl.TabPages.Add($settingsTab)

$settingsTab.add_Enter({
    try {
        Import-Module "$PSScriptRoot\Modules\Settings.psm1" -Force -ErrorAction SilentlyContinue
        $s = Get-AppSettings
        if ($s) { $txtInv.Text = $s.InvestigatorName; $txtCo.Text = $s.CompanyName; $txtGem.Text = $s.GeminiApiKey; $txtClaude.Text = $s.ClaudeApiKey }
        $lblStatus.Text = ""
    } catch {}
})

$btnSave.add_Click({
    try {
        Import-Module "$PSScriptRoot\Modules\Settings.psm1" -Force -ErrorAction SilentlyContinue
        $s = [pscustomobject]@{ InvestigatorName=$txtInv.Text; CompanyName=$txtCo.Text; GeminiApiKey=$txtGem.Text; ClaudeApiKey=$txtClaude.Text }
        if (Save-AppSettings -Settings $s) { $lblStatus.Text = "Saved."; $lblStatus.ForeColor = [System.Drawing.Color]::Green } else { $lblStatus.Text = "Failed to save."; $lblStatus.ForeColor = [System.Drawing.Color]::Red }
    } catch { $lblStatus.Text = $_.Exception.Message; $lblStatus.ForeColor = [System.Drawing.Color]::Red }
})

# Ensure Settings tab is the rightmost tab (last position)
try {
    $mainForm.add_Shown({
        try {
            if ($tabControl.TabPages.Contains($settingsTab)) {
                $tabControl.TabPages.Remove($settingsTab)
                $tabControl.TabPages.Add($settingsTab)
            }
        } catch {}
    })
} catch {}

# --- Exchange Online Controls Instantiation ---
$connectButton = New-Object System.Windows.Forms.Button
$connectButton.Text = "Connect"
$connectButton.Width = 100
$connectButtonTooltip = New-Object System.Windows.Forms.ToolTip
$connectButtonTooltip.SetToolTip($connectButton, "Connect to Exchange Online (Ctrl+O)")

$disconnectButton = New-Object System.Windows.Forms.Button
$disconnectButton.Text = "Disconnect"
$disconnectButton.Width = 100
$disconnectButtonTooltip = New-Object System.Windows.Forms.ToolTip
$disconnectButtonTooltip.SetToolTip($disconnectButton, "Disconnect from Exchange Online (Ctrl+D)")

$userMailboxListLabel = New-Object System.Windows.Forms.Label
$userMailboxListLabel.Text = "Mailboxes:"

$selectAllButton = New-Object System.Windows.Forms.Button
$selectAllButton.Text = "Select All"
$selectAllButton.Width = 100
$selectAllButtonTooltip = New-Object System.Windows.Forms.ToolTip
$selectAllButtonTooltip.SetToolTip($selectAllButton, "Select all mailboxes (Ctrl+A)")

$deselectAllButton = New-Object System.Windows.Forms.Button
$deselectAllButton.Text = "Deselect All"
$deselectAllButton.Width = 100
$deselectAllButtonTooltip = New-Object System.Windows.Forms.ToolTip
$deselectAllButtonTooltip.SetToolTip($deselectAllButton, "Deselect all mailboxes")

$orgDomainsLabel = New-Object System.Windows.Forms.Label
$orgDomainsLabel.Text = "Org Domains:"

$orgDomainsTextBox = New-Object System.Windows.Forms.TextBox
$orgDomainsTextBox.Width = 200
$orgDomainsTextBoxTooltip = New-Object System.Windows.Forms.ToolTip
$orgDomainsTextBoxTooltip.SetToolTip($orgDomainsTextBox, "Enter your organization domains (comma-separated) to identify external forwarding")

$keywordsLabel = New-Object System.Windows.Forms.Label
$keywordsLabel.Text = "Keywords:"

$keywordsTextBox = New-Object System.Windows.Forms.TextBox
$keywordsTextBox.Width = 200
$keywordsTextBoxTooltip = New-Object System.Windows.Forms.ToolTip
$keywordsTextBoxTooltip.SetToolTip($keywordsTextBox, "Enter suspicious keywords (comma-separated) to identify suspicious inbox rules")

$outputFolderLabel = New-Object System.Windows.Forms.Label
$outputFolderLabel.Text = "Output Folder:"

$outputFolderTextBox = New-Object System.Windows.Forms.TextBox
$outputFolderTextBox.Width = 200
$outputFolderTextBoxTooltip = New-Object System.Windows.Forms.ToolTip
$outputFolderTextBoxTooltip.SetToolTip($outputFolderTextBox, "Select folder where exported XLSX files will be saved")

$browseFolderButton = New-Object System.Windows.Forms.Button
$browseFolderButton.Text = "Browse..."
$browseFolderButton.Width = 100
$browseFolderButtonTooltip = New-Object System.Windows.Forms.ToolTip
$browseFolderButtonTooltip.SetToolTip($browseFolderButton, "Select folder for exporting XLSX reports")

$getRulesButton = New-Object System.Windows.Forms.Button
$getRulesButton.Text = "Export Rules"
$getRulesButton.Width = 120
$getRulesButtonTooltip = New-Object System.Windows.Forms.ToolTip
$getRulesButtonTooltip.SetToolTip($getRulesButton, "Export inbox rules for selected mailboxes (Ctrl+S)")

$manageRulesButton = New-Object System.Windows.Forms.Button
$manageRulesButton.Text = "Manage Rules"
$manageRulesButton.Width = 120
$manageRulesButton.Enabled = $true
$manageRulesButtonTooltip = New-Object System.Windows.Forms.ToolTip
$manageRulesButtonTooltip.SetToolTip($manageRulesButton, "View and manage inbox rules for selected mailbox")

$openFileButton = New-Object System.Windows.Forms.Button
$openFileButton.Text = "Open Last File"
$openFileButton.Width = 120
$openFileButtonTooltip = New-Object System.Windows.Forms.ToolTip
$openFileButtonTooltip.SetToolTip($openFileButton, "Open the last exported XLSX file")

$blockUserButton = New-Object System.Windows.Forms.Button
$blockUserButton.Text = "Block User"
$blockUserButton.Width = 100
$blockUserButton.Enabled = $true
$blockUserButtonTooltip = New-Object System.Windows.Forms.ToolTip
$blockUserButtonTooltip.SetToolTip($blockUserButton, "Block selected user from signing in (requires Graph connection)")

$unblockUserButton = New-Object System.Windows.Forms.Button
$unblockUserButton.Text = "Unblock User"
$unblockUserButton.Width = 100
$unblockUserButton.Enabled = $true
$unblockUserButtonTooltip = New-Object System.Windows.Forms.ToolTip
$unblockUserButtonTooltip.SetToolTip($unblockUserButton, "Unblock selected user from signing in (requires Graph connection)")

$revokeSessionsButton = New-Object System.Windows.Forms.Button
$revokeSessionsButton.Text = "Revoke Sessions"
$revokeSessionsButton.Width = 120
$revokeSessionsButtonTooltip = New-Object System.Windows.Forms.ToolTip
$revokeSessionsButtonTooltip.SetToolTip($revokeSessionsButton, "Revoke all active sessions for selected user (requires Graph connection)")

# Add load options for Exchange Online
$loadAllMailboxesButton = New-Object System.Windows.Forms.Button
$loadAllMailboxesButton.Text = "Load All Mailboxes"
$loadAllMailboxesButton.Width = 150
$loadAllMailboxesButton.Enabled = $false
$loadAllMailboxesButtonTooltip = New-Object System.Windows.Forms.ToolTip
$loadAllMailboxesButtonTooltip.SetToolTip($loadAllMailboxesButton, "Load all mailboxes (may take time for large tenants)")

$searchMailboxesButton = New-Object System.Windows.Forms.Button
$searchMailboxesButton.Text = "Search Mailboxes"
$searchMailboxesButton.Width = 150
$searchMailboxesButton.Enabled = $false
$searchMailboxesButtonTooltip = New-Object System.Windows.Forms.ToolTip
$searchMailboxesButtonTooltip.SetToolTip($searchMailboxesButton, "Search for specific mailboxes by name or email")

$manageConnectorsButton = New-Object System.Windows.Forms.Button
$manageConnectorsButton.Text = "Manage Connectors"
$manageConnectorsButton.Width = 140
$manageConnectorsButtonTooltip = New-Object System.Windows.Forms.ToolTip
$manageConnectorsButtonTooltip.SetToolTip($manageConnectorsButton, "View and manage Exchange Online connectors")

$manageTransportRulesButton = New-Object System.Windows.Forms.Button
$manageTransportRulesButton.Text = "Manage Transport Rules"
$manageTransportRulesButton.Width = 160
$manageTransportRulesButtonTooltip = New-Object System.Windows.Forms.ToolTip
$manageTransportRulesButtonTooltip.SetToolTip($manageTransportRulesButton, "View and manage Exchange Online transport rules")



# Add analyze selected button for Exchange Online tab
$analyzeSelectedButton = New-Object System.Windows.Forms.Button
$analyzeSelectedButton.Text = "Analyze Selected"
$analyzeSelectedButton.Width = 120
$analyzeSelectedButton.Enabled = $false
$analyzeSelectedButton.Visible = $true
$analyzeSelectedButtonTooltip = New-Object System.Windows.Forms.ToolTip
$analyzeSelectedButtonTooltip.SetToolTip($analyzeSelectedButton, "Perform detailed analysis (rules & permissions) for selected mailboxes")

# Debug: Log button creation
Write-Host "Analyze Selected Button created: $($analyzeSelectedButton.Text)"
Write-Host "Analyze Selected Button Enabled: $($analyzeSelectedButton.Enabled)"
Write-Host "Analyze Selected Button Visible: $($analyzeSelectedButton.Visible)"

# Mailbox list grid
$userMailboxGrid = New-Object System.Windows.Forms.DataGridView
$userMailboxGrid.Dock = 'Fill'
$userMailboxGrid.ReadOnly = $false
$userMailboxGrid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$userMailboxGrid.MultiSelect = $true
$userMailboxGrid.AllowUserToAddRows = $false
$userMailboxGrid.AutoGenerateColumns = $false
$userMailboxGrid.RowHeadersVisible = $false
$userMailboxGrid.AllowUserToOrderColumns = $true
$userMailboxGrid.AllowUserToResizeRows = $true
$userMailboxGrid.AllowUserToResizeColumns = $true
$userMailboxGrid.AutoSizeColumnsMode = 'Fill'
$userMailboxGrid.ColumnHeadersHeight = 25
$userMailboxGrid.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::DisableResizing
$userMailboxGrid.ColumnHeadersVisible = $true
$userMailboxGrid.EnableHeadersVisualStyles = $true

# Define columns with optimized widths
$colCheck = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
$colCheck.HeaderText = "Select"
$colCheck.DataPropertyName = "Select"
$colCheck.Width = 50
$colCheck.MinimumWidth = 50
$colCheck.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::AllCells
$colCheck.Name = "Select"
$colCheck.ReadOnly = $false
$userMailboxGrid.Columns.Add($colCheck)
$colCheckTooltip = New-Object System.Windows.Forms.ToolTip
$colCheckTooltip.SetToolTip($userMailboxGrid, "Check boxes to select mailboxes for analysis and export")

$colUPN = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colUPN.HeaderText = "UserPrincipalName"
$colUPN.Name = "UserPrincipalName"
$colUPN.DataPropertyName = "UserPrincipalName"
$colUPN.Width = 200
$colUPN.MinimumWidth = 150
$colUPN.ReadOnly = $true
$userMailboxGrid.Columns.Add($colUPN)

$colDisplayName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colDisplayName.HeaderText = "DisplayName"
$colDisplayName.Name = "DisplayName"
$colDisplayName.DataPropertyName = "DisplayName"
$colDisplayName.Width = 150
$colDisplayName.MinimumWidth = 100
$colDisplayName.ReadOnly = $true
$userMailboxGrid.Columns.Add($colDisplayName)

$colBlocked = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colBlocked.HeaderText = "SignInBlocked"
$colBlocked.Name = "SignInBlocked"
$colBlocked.DataPropertyName = "SignInBlocked"
$colBlocked.Width = 90
$colBlocked.MinimumWidth = 80
$colBlocked.ReadOnly = $true
$userMailboxGrid.Columns.Add($colBlocked)

$colRecipientType = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colRecipientType.HeaderText = "RecipientType"
$colRecipientType.Name = "RecipientType"
$colRecipientType.DataPropertyName = "RecipientType"
$colRecipientType.Width = 110
$colRecipientType.MinimumWidth = 100
$colRecipientType.ReadOnly = $true
$userMailboxGrid.Columns.Add($colRecipientType)

# Rule analysis columns
$colRulesCount = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colRulesCount.HeaderText = "TotalRules"
$colRulesCount.Name = "TotalRules"
$colRulesCount.DataPropertyName = "TotalRules"
$colRulesCount.Width = 70
$colRulesCount.MinimumWidth = 60
$colRulesCount.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::AllCells
$colRulesCount.ReadOnly = $true
$userMailboxGrid.Columns.Add($colRulesCount)

$colHiddenRules = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colHiddenRules.HeaderText = "Hidden Rules"
$colHiddenRules.Name = "HiddenRules"
$colHiddenRules.DataPropertyName = "HiddenRules"
$colHiddenRules.Width = 110
$colHiddenRules.MinimumWidth = 100
$colHiddenRules.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::AllCells
$colHiddenRules.ReadOnly = $true
$userMailboxGrid.Columns.Add($colHiddenRules)

$colSuspiciousRules = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colSuspiciousRules.HeaderText = "SuspiciousRules"
$colSuspiciousRules.Name = "SuspiciousRules"
$colSuspiciousRules.DataPropertyName = "SuspiciousRules"
$colSuspiciousRules.Width = 100
$colSuspiciousRules.MinimumWidth = 90
$colSuspiciousRules.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::AllCells
$colSuspiciousRules.ReadOnly = $true
$userMailboxGrid.Columns.Add($colSuspiciousRules)

$colExternalForwarding = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colExternalForwarding.HeaderText = "ExternalForwarding"
$colExternalForwarding.Name = "ExternalForwarding"
$colExternalForwarding.DataPropertyName = "ExternalForwarding"
$colExternalForwarding.Width = 110
$colExternalForwarding.MinimumWidth = 100
$colExternalForwarding.ReadOnly = $true
$userMailboxGrid.Columns.Add($colExternalForwarding)

$colDelegates = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colDelegates.HeaderText = "Delegates"
$colDelegates.Name = "Delegates"
$colDelegates.DataPropertyName = "Delegates"
$colDelegates.Width = 80
$colDelegates.MinimumWidth = 70
$colDelegates.ReadOnly = $true
$userMailboxGrid.Columns.Add($colDelegates)

$colFullAccess = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colFullAccess.HeaderText = "FullAccess"
$colFullAccess.Name = "FullAccess"
$colFullAccess.DataPropertyName = "FullAccess"
$colFullAccess.Width = 100
$colFullAccess.MinimumWidth = 80
$colFullAccess.ReadOnly = $true
$userMailboxGrid.Columns.Add($colFullAccess)

# Add search functionality for Exchange tab
$exchangeSearchLabel = New-Object System.Windows.Forms.Label
$exchangeSearchLabel.Text = "Search:"
$exchangeSearchLabel.Width = 50
$exchangeSearchLabel.Height = 20
$exchangeSearchLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft

$exchangeSearchTextBox = New-Object System.Windows.Forms.TextBox
$exchangeSearchTextBox.Width = 200
$exchangeSearchTextBox.Height = 20
$exchangeSearchTextBox.PlaceholderText = "Type to filter mailboxes..."


# Function to filter Exchange grid
function Filter-ExchangeGrid {
    param($searchText)
    $userMailboxGrid.Rows.Clear()
    
    # Get the original mailbox data from the script variable
    if (-not $script:allLoadedMailboxes) {
        return
    }
    
    foreach ($mbx in $script:allLoadedMailboxes) {
        if ([string]::IsNullOrWhiteSpace($searchText) -or 
            $mbx.UserPrincipalName -like "*$searchText*" -or 
            $mbx.DisplayName -like "*$searchText*") {
            # Get rule analysis for this mailbox
            $rulesCount = "0"
            $hiddenRules = "0"
            $suspiciousRules = "0"
            $externalForwarding = "Unknown"
            $delegates = "Unknown"
            $fullAccess = "Unknown"
            
            try {
                $rules = Get-InboxRule -Mailbox $mbx.UserPrincipalName -IncludeHidden -ErrorAction SilentlyContinue
                if ($rules) {
                    $analysis = Analyze-MailboxRulesEnhanced -Rules $rules -BaseSuspiciousKeywords $BaseSuspiciousKeywords
                    $rulesCount = $analysis.TotalRules.ToString()
                    $hiddenRules = $analysis.SuspiciousHidden.ToString()
                    $suspiciousRules = $analysis.SuspiciousVisible.ToString()
                    $externalForwarding = if ($analysis.HasExternalForwarding) { "Yes" } else { "No" }
                }
                
                # Analyze mailbox delegates and permissions
                try {
                    $delegates = Analyze-MailboxDelegates -UserPrincipalName $mbx.UserPrincipalName
                    $fullAccess = Analyze-MailboxPermissions -UserPrincipalName $mbx.UserPrincipalName
                } catch {
                    $delegates = "Error"
                    $fullAccess = "Error"
                }
            } catch {
                # Keep default values if analysis fails
            }
            
            $rowIdx = $userMailboxGrid.Rows.Add()
            $userMailboxGrid.Rows[$rowIdx].Cells["Select"].Value = $false
            $userMailboxGrid.Rows[$rowIdx].Cells["UserPrincipalName"].Value = $mbx.UserPrincipalName
            $userMailboxGrid.Rows[$rowIdx].Cells["DisplayName"].Value = $mbx.DisplayName
            $userMailboxGrid.Rows[$rowIdx].Cells["SignInBlocked"].Value = $mbx.SignInBlocked
            $userMailboxGrid.Rows[$rowIdx].Cells["RecipientType"].Value = $mbx.RecipientTypeDetails
            $userMailboxGrid.Rows[$rowIdx].Cells["TotalRules"].Value = $rulesCount
            $userMailboxGrid.Rows[$rowIdx].Cells["HiddenRules"].Value = $hiddenRules
            $userMailboxGrid.Rows[$rowIdx].Cells["SuspiciousRules"].Value = $suspiciousRules
            $userMailboxGrid.Rows[$rowIdx].Cells["ExternalForwarding"].Value = $externalForwarding
            $userMailboxGrid.Rows[$rowIdx].Cells["Delegates"].Value = $delegates
            $userMailboxGrid.Rows[$rowIdx].Cells["FullAccess"].Value = $fullAccess
        }
    }
}

# Function to filter Entra ID grid based on search text
function Filter-EntraGrid {
    param([string]$searchText)
    
    if ([string]::IsNullOrWhiteSpace($searchText)) {
        # Show all rows if search is empty
        for ($i = 0; $i -lt $entraUserGrid.Rows.Count; $i++) {
            $entraUserGrid.Rows[$i].Visible = $true
        }
        return
    }
    
    $searchText = $searchText.ToLower()
    
    # Filter rows based on search text
    for ($i = 0; $i -lt $entraUserGrid.Rows.Count; $i++) {
        $row = $entraUserGrid.Rows[$i]
        $upn = $row.Cells["UserPrincipalName"].Value
        $displayName = $row.Cells["DisplayName"].Value
        
        $visible = $false
        if ($upn -and $upn.ToLower().Contains($searchText)) { $visible = $true }
        if ($displayName -and $displayName.ToLower().Contains($searchText)) { $visible = $true }
        
        $row.Visible = $visible
    }
}

# Function to show a simple input dialog
function Show-InputDialog {
    param(
        [string]$Title = "Input",
        [string]$Prompt = "Enter value:",
        [string]$DefaultValue = ""
    )
    
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $Title
    $form.Size = New-Object System.Drawing.Size(400, 150)
    $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10, 20)
    $label.Size = New-Object System.Drawing.Size(360, 20)
    $label.Text = $Prompt
    $form.Controls.Add($label)
    
    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = New-Object System.Drawing.Point(10, 45)
    $textBox.Size = New-Object System.Drawing.Size(360, 20)
    $textBox.Text = $DefaultValue
    $form.Controls.Add($textBox)
    
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(200, 75)
    $okButton.Size = New-Object System.Drawing.Size(75, 23)
    $okButton.Text = "OK"
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $okButton
    $form.Controls.Add($okButton)
    
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(285, 75)
    $cancelButton.Size = New-Object System.Drawing.Size(75, 23)
    $cancelButton.Text = "Cancel"
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $cancelButton
    $form.Controls.Add($cancelButton)
    
    $textBox.Select()
    $result = $form.ShowDialog($mainForm)
    
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        return $textBox.Text
    } else {
        return $null
    }
}

$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Width = 200

$exchangeGrid = New-Object System.Windows.Forms.DataGridView
$exchangeGrid.ReadOnly = $true
$exchangeGrid.AllowUserToAddRows = $false
$exchangeGrid.AutoGenerateColumns = $true
$exchangeGrid.ColumnHeadersHeight = 25
$exchangeGrid.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::DisableResizing
$exchangeGrid.ColumnHeadersVisible = $true
$exchangeGrid.EnableHeadersVisualStyles = $true

# --- Entra ID Investigator Controls Instantiation ---
$entraConnectGraphButton = New-Object System.Windows.Forms.Button
$entraConnectGraphButton.Text = "Connect Entra"
$entraConnectGraphButton.Width = 140
$entraConnectGraphButtonTooltip = New-Object System.Windows.Forms.ToolTip
$entraConnectGraphButtonTooltip.SetToolTip($entraConnectGraphButton, "Connect to Microsoft Graph to load users and enable Entra ID features")

# Add load options for Entra ID
$loadAllUsersButton = New-Object System.Windows.Forms.Button
$loadAllUsersButton.Text = "Load All Users"
$loadAllUsersButton.Width = 150
$loadAllUsersButton.Enabled = $false
$loadAllUsersButtonTooltip = New-Object System.Windows.Forms.ToolTip
$loadAllUsersButtonTooltip.SetToolTip($loadAllUsersButton, "Load all users (may take time for large tenants)")

$searchUsersButton = New-Object System.Windows.Forms.Button
$searchUsersButton.Text = "Search Users"
$searchUsersButton.Width = 150
$searchUsersButton.Enabled = $false
$searchUsersButtonTooltip = New-Object System.Windows.Forms.ToolTip
$searchUsersButtonTooltip.SetToolTip($searchUsersButton, "Search for specific users by name or email")

$entraDisconnectGraphButton = New-Object System.Windows.Forms.Button
$entraDisconnectGraphButton.Text = "Disconnect Entra"
$entraDisconnectGraphButton.Width = 140
$entraDisconnectGraphButtonTooltip = New-Object System.Windows.Forms.ToolTip
$entraDisconnectGraphButtonTooltip.SetToolTip($entraDisconnectGraphButton, "Disconnect from Microsoft Graph")

$entraFixModulesButton = New-Object System.Windows.Forms.Button
$entraFixModulesButton.Text = "Fix Module Conflicts"
$entraFixModulesButton.Width = 160
$entraFixModulesButton.Enabled = $true
$entraFixModulesButton.BackColor = [System.Drawing.Color]::FromArgb(255, 193, 7) # Yellow/Orange color
$entraFixModulesButtonTooltip = New-Object System.Windows.Forms.ToolTip
$entraFixModulesButtonTooltip.SetToolTip($entraFixModulesButton, "Fix Microsoft Graph module version conflicts that prevent connection")

$entraOutputFolderLabel = New-Object System.Windows.Forms.Label
$entraOutputFolderLabel.Text = "Export Folder:"
$entraOutputFolderTextBox = New-Object System.Windows.Forms.TextBox
$entraOutputFolderTextBox.Width = 300
$entraBrowseFolderButton = New-Object System.Windows.Forms.Button
$entraBrowseFolderButton.Text = "Browse..."
$entraBrowseFolderButton.Width = 100
$entraBrowseFolderButtonTooltip = New-Object System.Windows.Forms.ToolTip
$entraBrowseFolderButtonTooltip.SetToolTip($entraBrowseFolderButton, "Select folder for exporting logs and reports")

$entraUserListLabel           = New-Object System.Windows.Forms.Label
$entraUserListLabel.Text      = "Users:"

$entraUserCheckedListBox      = New-Object System.Windows.Forms.CheckedListBox
$entraUserCheckedListBox.Width = 200
$entraUserCheckedListBox.Height = 80

$entraSignInDaysLabel         = New-Object System.Windows.Forms.Label
$entraSignInDaysLabel.Text    = "Sign-in Days:"

$entraSignInDaysUpDown        = New-Object System.Windows.Forms.NumericUpDown
$entraSignInDaysUpDown.Minimum = 1
$entraSignInDaysUpDown.Maximum = 90
$entraSignInDaysUpDown.Value   = 7

$entraSignInExportButton      = New-Object System.Windows.Forms.Button
$entraSignInExportButton.Text = "Fetch Sign-in Logs"
$entraSignInExportButton.Width = 140
$entraSignInExportButtonTooltip = New-Object System.Windows.Forms.ToolTip
$entraSignInExportButtonTooltip.SetToolTip($entraSignInExportButton, "Fetch sign-in logs for selected users (requires Graph connection)")

$entraSignInExportXlsxButton  = New-Object System.Windows.Forms.Button
$entraSignInExportXlsxButton.Text = "Export Sign-in XLSX"
$entraSignInExportXlsxButton.Width = 140
$entraSignInExportXlsxButton.Enabled = $false
$entraSignInExportXlsxButtonTooltip = New-Object System.Windows.Forms.ToolTip
$entraSignInExportXlsxButtonTooltip.SetToolTip($entraSignInExportXlsxButton, "Export sign-in logs to XLSX format")

$entraDetailsFetchButton      = New-Object System.Windows.Forms.Button
$entraDetailsFetchButton.Text = "User Details && Roles"
$entraDetailsFetchButton.Width = 140
$entraDetailsFetchButtonTooltip = New-Object System.Windows.Forms.ToolTip
$entraDetailsFetchButtonTooltip.SetToolTip($entraDetailsFetchButton, "View user details, roles, and group memberships (select one user)")

$entraAuditFetchButton        = New-Object System.Windows.Forms.Button
$entraAuditFetchButton.Text   = "Fetch Audit Logs"
$entraAuditFetchButton.Width = 140
$entraAuditFetchButtonTooltip = New-Object System.Windows.Forms.ToolTip
$entraAuditFetchButtonTooltip.SetToolTip($entraAuditFetchButton, "Fetch audit logs for selected user (select one user)")

$entraAuditExportXlsxButton   = New-Object System.Windows.Forms.Button
$entraAuditExportXlsxButton.Text = "Export Audit XLSX"
$entraAuditExportXlsxButton.Width = 140
$entraAuditExportXlsxButton.Enabled = $false
$entraAuditExportXlsxButtonTooltip = New-Object System.Windows.Forms.ToolTip
$entraAuditExportXlsxButtonTooltip.SetToolTip($entraAuditExportXlsxButton, "Export audit logs to XLSX format")

$entraMfaFetchButton          = New-Object System.Windows.Forms.Button
$entraMfaFetchButton.Text     = "Analyze MFA"
$entraMfaFetchButton.Width = 120
$entraMfaFetchButtonTooltip = New-Object System.Windows.Forms.ToolTip
$entraMfaFetchButtonTooltip.SetToolTip($entraMfaFetchButton, "Analyze MFA status for selected user (select one user)")

# Add user management buttons for Entra ID tab
$entraBlockUserButton = New-Object System.Windows.Forms.Button
$entraBlockUserButton.Text = "Block User"
$entraBlockUserButton.Width = 100
$entraBlockUserButton.Enabled = $false
$entraBlockUserButtonTooltip = New-Object System.Windows.Forms.ToolTip
$entraBlockUserButtonTooltip.SetToolTip($entraBlockUserButton, "Block selected user from signing in (requires Graph connection)")

$entraUnblockUserButton = New-Object System.Windows.Forms.Button
$entraUnblockUserButton.Text = "Unblock User"
$entraUnblockUserButton.Width = 100
$entraUnblockUserButton.Enabled = $false
$entraUnblockUserButtonTooltip = New-Object System.Windows.Forms.ToolTip
$entraUnblockUserButtonTooltip.SetToolTip($entraUnblockUserButton, "Unblock selected user from signing in (requires Graph connection)")

$entraRevokeSessionsButton = New-Object System.Windows.Forms.Button
$entraRevokeSessionsButton.Text = "Revoke Sessions"
$entraRevokeSessionsButton.Width = 120
$entraRevokeSessionsButton.Enabled = $false
$entraRevokeSessionsButtonTooltip = New-Object System.Windows.Forms.ToolTip
$entraRevokeSessionsButtonTooltip.SetToolTip($entraRevokeSessionsButton, "Revoke all active sessions for selected user (requires Graph connection)")

# Add password reset button for Entra ID tab
$entraResetPasswordButton = New-Object System.Windows.Forms.Button
$entraResetPasswordButton.Text = "Reset Password"
$entraResetPasswordButton.Width = 120
$entraResetPasswordButton.Enabled = $false
$entraResetPasswordButtonTooltip = New-Object System.Windows.Forms.ToolTip
$entraResetPasswordButtonTooltip.SetToolTip($entraResetPasswordButton, "Reset user password with memorable strong password (select one user)")

# Add restricted senders management button for Entra ID tab
$entraOpenDefenderRestrictedUsersButton = New-Object System.Windows.Forms.Button
$entraOpenDefenderRestrictedUsersButton.Text = "Open Defender Restricted Users"
$entraOpenDefenderRestrictedUsersButton.Width = 200
$entraOpenDefenderRestrictedUsersButton.Enabled = $true
$entraOpenDefenderRestrictedUsersButton.BackColor = [System.Drawing.Color]::LightBlue
$entraOpenDefenderRestrictedUsersButtonTooltip = New-Object System.Windows.Forms.ToolTip
$entraOpenDefenderRestrictedUsersButtonTooltip.SetToolTip($entraOpenDefenderRestrictedUsersButton, "Open Microsoft Defender Restricted Users page")

# Add Select All/Deselect All buttons for Entra ID tab
$entraSelectAllButton = New-Object System.Windows.Forms.Button
$entraSelectAllButton.Text = "Select All"
$entraSelectAllButton.Width = 80
$entraSelectAllButton.Enabled = $false
$entraSelectAllButtonTooltip = New-Object System.Windows.Forms.ToolTip
$entraSelectAllButtonTooltip.SetToolTip($entraSelectAllButton, "Select all users in the grid")

$entraDeselectAllButton = New-Object System.Windows.Forms.Button
$entraDeselectAllButton.Text = "Deselect All"
$entraDeselectAllButton.Width = 80
$entraDeselectAllButton.Enabled = $false
$entraDeselectAllButtonTooltip = New-Object System.Windows.Forms.ToolTip
$entraDeselectAllButtonTooltip.SetToolTip($entraDeselectAllButton, "Deselect all users in the grid")

# Add refresh roles button for Entra ID tab
$entraRefreshRolesButton = New-Object System.Windows.Forms.Button
$entraRefreshRolesButton.Text = "Refresh Roles"
$entraRefreshRolesButton.Width = 100
$entraRefreshRolesButton.Enabled = $false
$entraRefreshRolesButtonTooltip = New-Object System.Windows.Forms.ToolTip
$entraRefreshRolesButtonTooltip.SetToolTip($entraRefreshRolesButton, "Refresh role information for selected users")

# Add view admins button for Entra ID tab
$entraViewAdminsButton = New-Object System.Windows.Forms.Button
$entraViewAdminsButton.Text = "View Admins"
$entraViewAdminsButton.Width = 100
$entraViewAdminsButton.Enabled = $false
$entraViewAdminsButtonTooltip = New-Object System.Windows.Forms.ToolTip
$entraViewAdminsButtonTooltip.SetToolTip($entraViewAdminsButton, "Generate a report of all users with elevated roles")

# Add require password change button for Entra ID tab
$entraRequirePwdChangeButton = New-Object System.Windows.Forms.Button
$entraRequirePwdChangeButton.Text = "Require Password Change"
$entraRequirePwdChangeButton.Width = 170
$entraRequirePwdChangeButton.Enabled = $false
$entraRequirePwdChangeButtonTooltip = New-Object System.Windows.Forms.ToolTip
$entraRequirePwdChangeButtonTooltip.SetToolTip($entraRequirePwdChangeButton, "Require selected user(s) to change password at next sign-in (no password change)")

$entraSignInGrid              = New-Object System.Windows.Forms.DataGridView
$entraSignInGrid.ReadOnly     = $true
$entraSignInGrid.AllowUserToAddRows = $false
$entraSignInGrid.AutoGenerateColumns = $true
$entraSignInGrid.ColumnHeadersHeight = 25
$entraSignInGrid.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::DisableResizing
$entraSignInGrid.ColumnHeadersVisible = $true
$entraSignInGrid.EnableHeadersVisualStyles = $true

$entraAuditGrid               = New-Object System.Windows.Forms.DataGridView
$entraAuditGrid.ReadOnly      = $true
$entraAuditGrid.AllowUserToAddRows = $false
$entraAuditGrid.AutoGenerateColumns = $true
$entraAuditGrid.ColumnHeadersHeight = 25
$entraAuditGrid.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::DisableResizing
$entraAuditGrid.ColumnHeadersVisible = $true
$entraAuditGrid.EnableHeadersVisualStyles = $true

# Instantiate Entra ID Investigator tab buttons before layout
$entraViewSignInLogsButton = New-Object System.Windows.Forms.Button
$entraViewSignInLogsButton.Text = "View Sign-in Logs"
$entraViewSignInLogsButton.Width = 140

$entraViewAuditLogsButton = New-Object System.Windows.Forms.Button
$entraViewAuditLogsButton.Text = "View Audit Logs"
$entraViewAuditLogsButton.Width = 140

$entraExportSignInLogsButton = New-Object System.Windows.Forms.Button
$entraExportSignInLogsButton.Text = "Export Sign-in Logs"
$entraExportSignInLogsButton.Width = 160
$entraExportSignInLogsButton.Enabled = $false

$entraExportAuditLogsButton = New-Object System.Windows.Forms.Button
$entraExportAuditLogsButton.Text = "Export Audit Logs"
$entraExportAuditLogsButton.Width = 160
$entraExportAuditLogsButton.Enabled = $false

$entraOpenLastExportButton = New-Object System.Windows.Forms.Button
$entraOpenLastExportButton.Text = "Open Last Export"
$entraOpenLastExportButton.Width = 140
$entraOpenLastExportButton.Enabled = $true

# --- Exchange Online Tab Layout ---
$exchangeTab = New-Object System.Windows.Forms.TabPage; $exchangeTab.Text = "Exchange Online"

# Top action panel with two rows for better organization
$topActionPanel = New-Object System.Windows.Forms.Panel
$topActionPanel.Dock = 'Top'
$topActionPanel.Height = 80
$topActionPanel.AutoSize = $true

# First row - Connection, Loading, and Selection
$exchangeTopRow1 = New-Object System.Windows.Forms.FlowLayoutPanel
$exchangeTopRow1.Location = New-Object System.Drawing.Point(5, 5)
$exchangeTopRow1.Size = New-Object System.Drawing.Size(1200, 35)
$exchangeTopRow1.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
$exchangeTopRow1.WrapContents = $true
$exchangeTopRow1.AutoSize = $true
$exchangeTopRow1.Controls.AddRange(@($connectButton, $disconnectButton, $loadAllMailboxesButton, $searchMailboxesButton, $selectAllButton, $deselectAllButton))

# Second row - Analysis and Management
$exchangeTopRow2 = New-Object System.Windows.Forms.FlowLayoutPanel
$exchangeTopRow2.Location = New-Object System.Drawing.Point(5, 40)
$exchangeTopRow2.Size = New-Object System.Drawing.Size(1200, 35)
$exchangeTopRow2.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
$exchangeTopRow2.WrapContents = $true
$exchangeTopRow2.AutoSize = $true
$exchangeTopRow2.Controls.AddRange(@($manageRulesButton, $analyzeSelectedButton, $manageConnectorsButton, $manageTransportRulesButton))

# Add search controls to the first row
$exchangeTopRow1.Controls.Add($exchangeSearchLabel)
$exchangeTopRow1.Controls.Add($exchangeSearchTextBox)

# Add both rows to the top panel
$topActionPanel.Controls.Add($exchangeTopRow1)
$topActionPanel.Controls.Add($exchangeTopRow2)

# Debug: Log button addition to panel
Write-Host "Exchange Online buttons organized into two rows"
Write-Host "Row 1 Controls Count: $($exchangeTopRow1.Controls.Count)"
Write-Host "Row 2 Controls Count: $($exchangeTopRow2.Controls.Count)"

$exchangeTab.Controls.Add($topActionPanel)

# Panel for mailbox label and grid (fills remaining space)
$mailboxPanel = New-Object System.Windows.Forms.Panel
$mailboxPanel.Dock = 'Fill'
$mailboxPanel.Padding = New-Object System.Windows.Forms.Padding(5, 85, 5, 5)
$mailboxPanel.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 0)

# Add label and grid to mailbox panel
$userMailboxListLabel.Dock = 'Top'
$userMailboxListLabel.Height = 25
$userMailboxListLabel.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 5)

$userMailboxGrid.Dock = 'Fill'
$userMailboxGrid.ScrollBars = [System.Windows.Forms.ScrollBars]::Both  # Show both scrollbars

$mailboxPanel.Controls.Add($userMailboxGrid)
$mailboxPanel.Controls.Add($userMailboxListLabel)
$exchangeTab.Controls.Add($mailboxPanel)

# Action buttons panel at the very bottom (full width, 2 rows)
$actionPanel = New-Object System.Windows.Forms.Panel
$actionPanel.Dock = 'Bottom'
$actionPanel.MinimumSize = New-Object System.Drawing.Size(0, 80)
$actionPanel.Height = 80

# Row 1: Output Folder, Browse, Export Rules, Open Last File
$row1 = New-Object System.Windows.Forms.FlowLayoutPanel
$row1.Dock = 'Top'
$row1.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
$row1.WrapContents = $false
$row1.AutoSize = $true
$row1.Controls.AddRange(@($outputFolderLabel, $outputFolderTextBox, $browseFolderButton, $getRulesButton, $openFileButton))

# Row 2: Org Domains, Keywords, Manage Restricted Senders, ProgressBar
$row2 = New-Object System.Windows.Forms.FlowLayoutPanel
$row2.Dock = 'Top'
$row2.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
$row2.WrapContents = $false
$row2.AutoSize = $true
$row2.Controls.AddRange(@($orgDomainsLabel, $orgDomainsTextBox, $keywordsLabel, $keywordsTextBox, $progressBar))

$actionPanel.Controls.Add($row1)
$actionPanel.Controls.Add($row2)

# Remove old actionPanel and add new one
$exchangeTab.Controls.Remove($actionPanel)
$exchangeTab.Controls.Add($actionPanel)

# DataGridView for results (hidden by default, shown when results are present)
$exchangeGrid.Dock = 'Fill'
$exchangeGrid.Visible = $false
$exchangeTab.Controls.Add($exchangeGrid)



# --- Entra ID Investigator Tab Layout ---
$entraTab = New-Object System.Windows.Forms.TabPage; $entraTab.Text = "Entra ID Investigator"

# Top action panel with two rows for better organization
$entraTopPanel = New-Object System.Windows.Forms.Panel
$entraTopPanel.Dock = 'Top'
$entraTopPanel.Height = 100
$entraTopPanel.AutoSize = $true

# First row - Connection and basic functions
$entraTopRow1 = New-Object System.Windows.Forms.FlowLayoutPanel
$entraTopRow1.Location = New-Object System.Drawing.Point(5, 5)
$entraTopRow1.Size = New-Object System.Drawing.Size(1200, 35)
$entraTopRow1.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
$entraTopRow1.WrapContents = $true
$entraTopRow1.AutoSize = $true
$entraTopRow1.Controls.AddRange(@($entraConnectGraphButton, $entraDisconnectGraphButton, $entraFixModulesButton, $loadAllUsersButton, $searchUsersButton, $entraSelectAllButton, $entraDeselectAllButton, $entraViewSignInLogsButton, $entraViewAuditLogsButton, $entraDetailsFetchButton, $entraMfaFetchButton))

# Second row - User management functions
$entraTopRow2 = New-Object System.Windows.Forms.FlowLayoutPanel
$entraTopRow2.Location = New-Object System.Drawing.Point(5, 40)
$entraTopRow2.Size = New-Object System.Drawing.Size(1200, 35)
$entraTopRow2.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
$entraTopRow2.WrapContents = $false
$entraTopRow2.AutoSize = $false
$entraTopRow2.Controls.AddRange(@($entraBlockUserButton, $entraUnblockUserButton, $entraRevokeSessionsButton, $entraResetPasswordButton, $entraRequirePwdChangeButton, $entraRefreshRolesButton, $entraViewAdminsButton, $entraOpenDefenderRestrictedUsersButton))

# Add search controls to the top panel
$entraSearchLabel = New-Object System.Windows.Forms.Label
$entraSearchLabel.Text = "Search:"
$entraSearchLabel.Width = 50
$entraSearchLabel.Height = 20
$entraSearchLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft

$entraSearchTextBox = New-Object System.Windows.Forms.TextBox
$entraSearchTextBox.Width = 200
$entraSearchTextBox.Height = 20
$entraSearchTextBox.PlaceholderText = "Type to filter users..."

# Add search controls to the second row
$entraTopRow2.Controls.Add($entraSearchLabel)
$entraTopRow2.Controls.Add($entraSearchTextBox)

# Add both rows to the top panel
$entraTopPanel.Controls.Add($entraTopRow1)
$entraTopPanel.Controls.Add($entraTopRow2)

# Panel for user grid
$entraGridPanel = New-Object System.Windows.Forms.Panel
$entraGridPanel.Dock = 'Fill'
$entraGridPanel.Padding = New-Object System.Windows.Forms.Padding(5, 105, 5, 15)
$entraGridPanel.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 0)

# User grid
$entraUserGrid = New-Object System.Windows.Forms.DataGridView
$entraUserGrid.Dock = 'Fill'
$entraUserGrid.ReadOnly = $false
$entraUserGrid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$entraUserGrid.MultiSelect = $true
$entraUserGrid.AllowUserToAddRows = $false
$entraUserGrid.AutoGenerateColumns = $false
$entraUserGrid.RowHeadersVisible = $false
$entraUserGrid.ColumnHeadersVisible = $true
$entraUserGrid.EnableHeadersVisualStyles = $true
$entraUserGrid.ColumnHeadersHeight = 25
$entraUserGrid.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::DisableResizing
$entraUserGrid.AutoSizeColumnsMode = 'Fill'

# Define columns with optimized widths
$colEntraCheck = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
$colEntraCheck.HeaderText = "Select"
$colEntraCheck.Width = 50
$colEntraCheck.MinimumWidth = 50
$colEntraCheck.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::AllCells
$colEntraCheck.Name = "Select"
$colEntraCheck.ReadOnly = $false
$entraUserGrid.Columns.Add($colEntraCheck)

$colEntraUPN = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colEntraUPN.HeaderText = "UserPrincipalName"
$colEntraUPN.Width = 200
$colEntraUPN.MinimumWidth = 150
$colEntraUPN.Name = "UserPrincipalName"
$colEntraUPN.ReadOnly = $true
$entraUserGrid.Columns.Add($colEntraUPN)

$colEntraDisplayName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colEntraDisplayName.HeaderText = "DisplayName"
$colEntraDisplayName.Width = 150
$colEntraDisplayName.MinimumWidth = 100
$colEntraDisplayName.Name = "DisplayName"
$colEntraDisplayName.ReadOnly = $true
$entraUserGrid.Columns.Add($colEntraDisplayName)

$colEntraLicensed = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colEntraLicensed.HeaderText = "Licensed"
$colEntraLicensed.Width = 80
$colEntraLicensed.MinimumWidth = 70
$colEntraLicensed.Name = "Licensed"
$colEntraLicensed.ReadOnly = $true
$entraUserGrid.Columns.Add($colEntraLicensed)

$colEntraRoles = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colEntraRoles.HeaderText = "Roles"
$colEntraRoles.Width = 120
$colEntraRoles.MinimumWidth = 100
$colEntraRoles.Name = "Roles"
$colEntraRoles.ReadOnly = $true
$entraUserGrid.Columns.Add($colEntraRoles)

$entraGridPanel.Controls.Add($entraUserGrid)

# Bottom panel with buttons
$entraBottomPanel = New-Object System.Windows.Forms.Panel
$entraBottomPanel.Dock = 'Bottom'
$entraBottomPanel.Height = 70

# Add buttons to bottom panel
$entraBrowseFolderButton.Location = New-Object System.Drawing.Point(10, 15)
$entraBrowseFolderButton.Size = New-Object System.Drawing.Size(120, 30)
$entraExportSignInLogsButton.Location = New-Object System.Drawing.Point(140, 15)
$entraExportSignInLogsButton.Size = New-Object System.Drawing.Size(140, 30)
$entraExportAuditLogsButton.Location = New-Object System.Drawing.Point(290, 15)
$entraExportAuditLogsButton.Size = New-Object System.Drawing.Size(140, 30)
$entraOpenLastExportButton.Location = New-Object System.Drawing.Point(440, 15)
$entraOpenLastExportButton.Size = New-Object System.Drawing.Size(120, 30)

# Add export path controls
$entraOutputFolderLabel.Location = New-Object System.Drawing.Point(580, 18)
$entraOutputFolderTextBox.Location = New-Object System.Drawing.Point(680, 15)
$entraOutputFolderTextBox.Width = 200
$entraOutputFolderTextBox.Height = 25

$entraBottomPanel.Controls.AddRange(@($entraBrowseFolderButton, $entraExportSignInLogsButton, $entraExportAuditLogsButton, $entraOpenLastExportButton, $entraOutputFolderLabel, $entraOutputFolderTextBox))

# Add panels to tab in order
$entraTab.Controls.Add($entraTopPanel)
$entraTab.Controls.Add($entraGridPanel)
$entraTab.Controls.Add($entraBottomPanel)





# Add a read-only textbox to display the selected export path
$entraSelectedPathTextBox = New-Object System.Windows.Forms.TextBox
$entraSelectedPathTextBox.ReadOnly = $true
$entraSelectedPathTextBox.Width = 300
$entraSelectedPathTextBox.Text = ""

# Update the selected path textbox when the folder changes
$entraOutputFolderTextBox.add_TextChanged({
    $entraSelectedPathTextBox.Text = $entraOutputFolderTextBox.Text
    UpdateEntraButtonStates
})

# Update the selected path textbox when Browse is used
$entraBrowseFolderButton.add_Click({
    $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($folderDialog.ShowDialog() -eq 'OK') {
        $entraOutputFolderTextBox.Text = $folderDialog.SelectedPath
        $entraSelectedPathTextBox.Text = $folderDialog.SelectedPath
    }
})
$entraTab.Padding = 0
$entraTab.Margin = 0
$entraTab.Dock = 'Fill'

# Populate Entra user grid after Graph authentication
$entraConnectGraphButton.add_Click({
    $statusLabel.Text = "Connecting to Microsoft Graph..."
    $mainForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $entraConnectGraphButton.Enabled = $false
    $entraSignInExportButton.Enabled = $false; $entraDetailsFetchButton.Enabled = $false; $entraAuditFetchButton.Enabled = $false; $entraMfaFetchButton.Enabled = $false

    # Bring main form to front and focus it to ensure auth dialog appears on top
    $mainForm.BringToFront()
    $mainForm.Focus()
    [System.Windows.Forms.Application]::DoEvents()
    
    # Show user-friendly message about authentication
    $statusLabel.Text = "Authentication window should appear. If not visible, check your taskbar or Alt+Tab to find it."
    Show-Progress -message "Authentication window should appear. If not visible, check your taskbar or Alt+Tab to find it." -progress 10

    try {
        if (Connect-EntraGraph) {
            $script:graphConnection = $true
            
            # Enable load buttons and disable connect button
            $loadAllUsersButton.Enabled = $true
            $searchUsersButton.Enabled = $true
            $entraDisconnectGraphButton.Enabled = $true
            $entraConnectGraphButton.Enabled = $false
            
            Write-Host "Microsoft Graph connected. Load buttons enabled: LoadAll=$($loadAllUsersButton.Enabled), Search=$($searchUsersButton.Enabled)"
            
            $statusLabel.Text = "Connected to Microsoft Graph. Use 'Load All Users' or 'Search Users' to load data."
            Show-Progress -message "Connected to Microsoft Graph successfully." -progress 100
        } else {
            # Check if this is a user cancellation
            $errorMessage = $_.Exception.Message
            $isUserCancellation = $errorMessage -match "User cancelled|Operation cancelled|User canceled|Authentication cancelled|Authentication canceled" -or 
                                 $errorMessage -match "AADSTS50020|AADSTS50076|AADSTS50079" -or
                                 $errorMessage -match "The user cancelled the authentication"
            
            if ($isUserCancellation) {
                # User cancelled - just update status without showing error popup
                $statusLabel.Text = "Microsoft Graph connection cancelled by user."
            } else {
                # Real error - show user-friendly error message
                $statusLabel.Text = "Failed to connect to Microsoft Graph."
                [System.Windows.Forms.MessageBox]::Show("Failed to connect to Microsoft Graph: $($_.Exception.Message)", "Connection Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }
    } catch {
        # Check if this is a user cancellation
        $errorMessage = $_.Exception.Message
        $isUserCancellation = $errorMessage -match "User cancelled|Operation cancelled|User canceled|Authentication cancelled|Authentication canceled" -or 
                             $errorMessage -match "AADSTS50020|AADSTS50076|AADSTS50079" -or
                             $errorMessage -match "The user cancelled the authentication"
        
        if ($isUserCancellation) {
            # User cancelled - just update status without showing error popup
            $statusLabel.Text = "Microsoft Graph connection cancelled by user."
        } else {
            # Real error - show user-friendly error message
            $statusLabel.Text = "Failed to connect to Microsoft Graph."
            [System.Windows.Forms.MessageBox]::Show("Failed to connect to Microsoft Graph: $($_.Exception.Message)", "Connection Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
    $mainForm.Cursor = [System.Windows.Forms.Cursors]::Default
    $entraConnectGraphButton.Enabled = $true
})

# Load All Users button handler
$loadAllUsersButton.add_Click({
    if (-not $script:graphConnection) {
        [System.Windows.Forms.MessageBox]::Show("Please connect to Microsoft Graph first.", "Not Connected", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    try {
        $statusLabel.Text = "Loading all users..."
        $mainForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        
        # Get all users with full details
        $users = Get-MgUser -All -Property Id, UserPrincipalName, DisplayName, AssignedLicenses -ErrorAction Stop
        $entraUserGrid.Rows.Clear()
        
        $totalUsers = $users.Count
        $processedCount = 0
        
        foreach ($u in $users) {
            try {
                # Check licensing
                $isLicensed = $u.AssignedLicenses -and $u.AssignedLicenses.Count -gt 0
                $licensedText = if ($isLicensed) { "Licensed" } else { "Unlicensed" }
                
                # Get user roles
                $userRoles = @()
                try {
                    $userRoleMemberships = Get-MgUserMemberOf -UserId $u.UserPrincipalName -ErrorAction SilentlyContinue
                    if ($userRoleMemberships) {
                        foreach ($role in $userRoleMemberships) {
                            if ($role.'@odata.type' -eq '#microsoft.graph.directoryRole') {
                                $userRoles += $role.DisplayName
                            }
                        }
                    }
                } catch {
                    # Role lookup failed, continue without roles
                }
                
                $rolesText = if ($userRoles.Count -gt 0) { ($userRoles -join ", ") } else { "Click 'Refresh Roles' to view" }
                
                # Add row
                $rowIndex = $entraUserGrid.Rows.Add()
                $entraUserGrid.Rows[$rowIndex].Cells["Select"].Value = $false
                $entraUserGrid.Rows[$rowIndex].Cells["UserPrincipalName"].Value = $u.UserPrincipalName
                $entraUserGrid.Rows[$rowIndex].Cells["DisplayName"].Value = $u.DisplayName
                $entraUserGrid.Rows[$rowIndex].Cells["Licensed"].Value = $licensedText
                $entraUserGrid.Rows[$rowIndex].Cells["Roles"].Value = $rolesText
                
                $processedCount++
                
                # Update progress every 50 users
                if ($processedCount % 50 -eq 0) {
                    $statusLabel.Text = "Loading users... ($processedCount/$totalUsers)"
                    [System.Windows.Forms.Application]::DoEvents()
                }
            } catch {
                # Skip users that cause errors
                continue
            }
        }
        
        UpdateEntraButtonStates
        $statusLabel.Text = "Loaded $processedCount users"
        [System.Windows.Forms.MessageBox]::Show("Successfully loaded $processedCount users.", "Load Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        
    } catch {
        $statusLabel.Text = "Error loading users: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Error loading users: $($_.Exception.Message)", "Load Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    } finally {
        $mainForm.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})

# Search Users button handler
$searchUsersButton.add_Click({
    if (-not $script:graphConnection) {
        [System.Windows.Forms.MessageBox]::Show("Please connect to Microsoft Graph first.", "Not Connected", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    $searchTerm = Show-InputDialog -Title "Search Users" -Prompt "Enter search term (name or email):"
    if ([string]::IsNullOrWhiteSpace($searchTerm)) { return }
    
    Write-Host "Search term entered: '$searchTerm'"
    
    try {
        $statusLabel.Text = "Searching for users..."
        $mainForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        
        # Search for users using the search term (Microsoft Graph supports startsWith and eq)
        Write-Host "Searching for users with filter: startsWith(DisplayName,'$searchTerm') or startsWith(UserPrincipalName,'$searchTerm')"
        $users = Get-MgUser -Filter "startsWith(DisplayName,'$searchTerm') or startsWith(UserPrincipalName,'$searchTerm')" -All -Property Id, UserPrincipalName, DisplayName, AssignedLicenses -ErrorAction Stop
        
        Write-Host "Found $($users.Count) users"
        
        if ($users.Count -eq 0) {
            Write-Host "No users found with the current filter. Trying alternative search methods..."
            
            # Try alternative search methods using supported operators
            try {
                # Try exact match first
                $usersAlt1 = Get-MgUser -Filter "DisplayName eq '$searchTerm'" -All -Property Id, UserPrincipalName, DisplayName, AssignedLicenses -ErrorAction SilentlyContinue
                Write-Host "Alternative search 1 (exact DisplayName match): Found $($usersAlt1.Count) users"
                
                $usersAlt2 = Get-MgUser -Filter "UserPrincipalName eq '$searchTerm'" -All -Property Id, UserPrincipalName, DisplayName, AssignedLicenses -ErrorAction SilentlyContinue
                Write-Host "Alternative search 2 (exact UserPrincipalName match): Found $($usersAlt2.Count) users"
                
                # Try case-insensitive search by getting all users and filtering client-side
                $allUsers = Get-MgUser -All -Property Id, UserPrincipalName, DisplayName, AssignedLicenses -ErrorAction SilentlyContinue
                $usersAlt3 = $allUsers | Where-Object { 
                    $_.DisplayName -like "*$searchTerm*" -or $_.UserPrincipalName -like "*$searchTerm*" 
                }
                Write-Host "Alternative search 3 (client-side filtering): Found $($usersAlt3.Count) users"
                
                # Combine all results
                $users = @($usersAlt1) + @($usersAlt2) + @($usersAlt3) | Sort-Object UserPrincipalName -Unique
                Write-Host "Combined alternative searches: Found $($users.Count) users"
            } catch {
                Write-Host "Alternative searches also failed: $($_.Exception.Message)"
            }
            
            if ($users.Count -eq 0) {
                [System.Windows.Forms.MessageBox]::Show("No users found matching '$searchTerm'.", "No Results", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                return
            }
        }
        
        $entraUserGrid.Rows.Clear()
        $processedCount = 0
        
        foreach ($u in $users) {
            try {
                # Check licensing
                $isLicensed = $u.AssignedLicenses -and $u.AssignedLicenses.Count -gt 0
                $licensedText = if ($isLicensed) { "Licensed" } else { "Unlicensed" }
                
                # Get user roles
                $userRoles = @()
                try {
                    $userRoleMemberships = Get-MgUserMemberOf -UserId $u.UserPrincipalName -ErrorAction SilentlyContinue
                    if ($userRoleMemberships) {
                        foreach ($role in $userRoleMemberships) {
                            if ($role.'@odata.type' -eq '#microsoft.graph.directoryRole') {
                                $userRoles += $role.DisplayName
                            }
                        }
                    }
                } catch {
                    # Role lookup failed, continue without roles
                }
                
                $rolesText = if ($userRoles.Count -gt 0) { ($userRoles -join ", ") } else { "Click 'Refresh Roles' to view" }
                
                # Add row
                $rowIndex = $entraUserGrid.Rows.Add()
                $entraUserGrid.Rows[$rowIndex].Cells["Select"].Value = $false
                $entraUserGrid.Rows[$rowIndex].Cells["UserPrincipalName"].Value = $u.UserPrincipalName
                $entraUserGrid.Rows[$rowIndex].Cells["DisplayName"].Value = $u.DisplayName
                $entraUserGrid.Rows[$rowIndex].Cells["Licensed"].Value = $licensedText
                $entraUserGrid.Rows[$rowIndex].Cells["Roles"].Value = $rolesText
                
                $processedCount++
            } catch {
                # Skip users that cause errors
                continue
            }
        }
        
        UpdateEntraButtonStates
        $statusLabel.Text = "Loaded $processedCount users matching '$searchTerm'"
        [System.Windows.Forms.MessageBox]::Show("Found and loaded $processedCount users matching '$searchTerm'.", "Search Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        
    } catch {
        $statusLabel.Text = "Error searching users: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Error searching users: $($_.Exception.Message)", "Search Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    } finally {
        $mainForm.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})
$entraSignInExportButton.add_Click({
    $entraUserGrid.EndEdit()
    Write-Host 'EntraUserGrid Columns:'
    foreach ($col in $entraUserGrid.Columns) { Write-Host $col.Name }
    $selectedUpns = @()
    for ($i = 0; $i -lt $entraUserGrid.Rows.Count; $i++) {
        if ($entraUserGrid.Rows[$i].Cells["Select"].Value -eq $true) {
            $upn = $entraUserGrid.Rows[$i].Cells["UserPrincipalName"].Value
            if (-not [string]::IsNullOrWhiteSpace($upn)) { $selectedUpns += $upn }
        }
    }
    Write-Host "Selected UPNs: $($selectedUpns -join ', ')"
    if ($selectedUpns.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Select at least one user with a valid UPN.", "No Users Selected", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning); return
    }
    $statusLabel.Text = "Fetching sign-in logs..."; $mainForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $days = $entraSignInDaysUpDown.Value
    $outputFolder = $entraOutputFolderTextBox.Text
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $csvFilePath = Join-Path $outputFolder "EntraSignInLogs_$timestamp.csv"
    $xlsxFilePath = Join-Path $outputFolder "EntraSignInLogs_$timestamp.xlsx"
    try {
        $allLogs = Get-EntraSignInLogs -UserPrincipalNames $selectedUpns -Days $days
        $entraSignInGrid.DataSource = $null
        if (-not $allLogs -or $allLogs.Count -eq 0) {
            $statusLabel.Text = "No sign-in logs found."; $mainForm.Cursor = [System.Windows.Forms.Cursors]::Default
            [System.Windows.Forms.MessageBox]::Show("No sign-in logs found for selected users.", "No Logs", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information); $entraSignInExportXlsxButton.Enabled = $false; return
        }
        $entraSignInGrid.DataSource = $allLogs
        $entraSignInExportXlsxButton.Tag = $allLogs
        $entraSignInExportXlsxButton.Enabled = $true
        $statusLabel.Text = "Sign-in logs loaded."
    } catch {
        $statusLabel.Text = "Error during sign-in log fetch: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Error during sign-in log fetch: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        $entraSignInExportXlsxButton.Enabled = $false
    } finally { $mainForm.Cursor = [System.Windows.Forms.Cursors]::Default }
})
$entraSignInExportXlsxButton.add_Click({
    $allLogs = $entraSignInExportXlsxButton.Tag
    if (-not $allLogs -or $allLogs.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("No sign-in logs to export.", "No Data", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information); return }
    $outputFolder = $entraOutputFolderTextBox.Text
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $csvFilePath = Join-Path $outputFolder "EntraSignInLogs_$timestamp.csv"
    $xlsxFilePath = Join-Path $outputFolder "EntraSignInLogs_$timestamp.xlsx"
    $allLogs | Export-Csv -Path $csvFilePath -NoTypeInformation -Encoding UTF8
    if (Format-InboxRuleXlsx -CsvPath $csvFilePath -XlsxPath $xlsxFilePath) {
        try { Remove-Item $csvFilePath -Force } catch {}
        $entraOpenFileButton.Tag = $xlsxFilePath
        $entraOpenFileButton.Enabled = $true
        $script:lastExportedXlsxPath = $xlsxFilePath # Update the script-level variable
        [System.Windows.Forms.MessageBox]::Show("Exported and formatted sign-in logs to:\n$xlsxFilePath", "XLSX Export OK", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    } else {
        [System.Windows.Forms.MessageBox]::Show("CSV Exported to:\n$csvFilePath\n\nXLSX/Format FAILED. Check Excel install & console.", "XLSX Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    }
})
$entraDetailsFetchButton.add_Click({
    $entraUserGrid.EndEdit()
    Write-Host 'EntraUserGrid Columns:'
    foreach ($col in $entraUserGrid.Columns) { Write-Host $col.Name }
    $selectedUpns = @()
    for ($i = 0; $i -lt $entraUserGrid.Rows.Count; $i++) {
        if ($entraUserGrid.Rows[$i].Cells["Select"].Value -eq $true) {
            $upn = $entraUserGrid.Rows[$i].Cells["UserPrincipalName"].Value
            if (-not [string]::IsNullOrWhiteSpace($upn)) { $selectedUpns += $upn }
        }
    }
    Write-Host "Selected UPNs: $($selectedUpns -join ', ')"
    if ($selectedUpns.Count -ne 1) {
        [System.Windows.Forms.MessageBox]::Show("Select exactly one user with a valid UPN.", "Select One User", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information); return
    }
    $upn = $selectedUpns[0]
    $statusLabel.Text = "Fetching user details..."; $mainForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        $result = Get-EntraUserDetailsAndRoles -UserPrincipalName $upn
        if ($result.User) {
            # Build roles and groups strings separately to avoid syntax issues
            $rolesText = if ($result.Roles.Count -gt 0) { $result.Roles -join "`r`n" } else { "None" }
            $groupsText = if ($result.Groups.Count -gt 0) { $result.Groups -join "`r`n" } else { "None" }
            
            $details = "User Principal Name: $($result.User.UserPrincipalName)`r`nDisplay Name: $($result.User.DisplayName)`r`nAccount Enabled: $($result.User.AccountEnabled)`r`nLast Password Change: $($result.User.LastPasswordChangeDateTime)`r`n" +
                "-----------------------------`r`nRoles:`r`n$rolesText`r`n" +
                "-----------------------------`r`nGroups:`r`n$groupsText"
            $form = New-Object System.Windows.Forms.Form
            $form.Text = "User Details && Roles"
            $form.Size = New-Object System.Drawing.Size(600, 500)
            $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
            $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
            $form.MaximizeBox = $true
            $textbox = New-Object System.Windows.Forms.TextBox
            $textbox.Multiline = $true
            $textbox.ReadOnly = $true
            $textbox.ScrollBars = 'Both'
            $textbox.Dock = 'Fill'
            $textbox.Font = New-Object System.Drawing.Font('Consolas', 10)
            $textbox.Text = $details
            $form.Controls.Add($textbox)
            $form.ShowDialog($mainForm)
            $form.Dispose()
        } else {
            [System.Windows.Forms.MessageBox]::Show("Error: $($result.Error)", "User Details & Roles Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error fetching user details: $($_.Exception.Message)", "User Details & Roles Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    } finally { $mainForm.Cursor = [System.Windows.Forms.Cursors]::Default }
})
$entraAuditFetchButton.add_Click({
    $entraUserGrid.EndEdit()
    Write-Host 'EntraUserGrid Columns:'
    foreach ($col in $entraUserGrid.Columns) { Write-Host $col.Name }
    $selectedUpns = @()
    for ($i = 0; $i -lt $entraUserGrid.Rows.Count; $i++) {
        if ($entraUserGrid.Rows[$i].Cells["Select"].Value -eq $true) {
            $upn = $entraUserGrid.Rows[$i].Cells["UserPrincipalName"].Value
            if (-not [string]::IsNullOrWhiteSpace($upn)) { $selectedUpns += $upn }
        }
    }
    Write-Host "Selected UPNs: $($selectedUpns -join ', ')"
    if ($selectedUpns.Count -ne 1) {
        [System.Windows.Forms.MessageBox]::Show("Select exactly one user with a valid UPN.", "Select One User", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information); return
    }
    $upn = $selectedUpns[0]
    $days = $entraSignInDaysUpDown.Value
    $statusLabel.Text = "Fetching audit logs..."; $mainForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        $logs = Get-EntraUserAuditLogs -UserPrincipalName $upn -Days $days
        $entraAuditGrid.DataSource = $null
        if ($logs -and $logs.Count -gt 0) {
            $entraAuditGrid.DataSource = $logs
            $entraAuditExportXlsxButton.Tag = $logs
            $entraAuditExportXlsxButton.Enabled = $true
            $statusLabel.Text = "Audit logs loaded."
        } else {
            [System.Windows.Forms.MessageBox]::Show("No audit logs found for $upn.", "Audit Logs", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            $entraAuditExportXlsxButton.Enabled = $false
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error fetching audit logs: $($_.Exception.Message)", "Audit Logs Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        $entraAuditExportXlsxButton.Enabled = $false
    } finally { $mainForm.Cursor = [System.Windows.Forms.Cursors]::Default }
})
$entraAuditExportXlsxButton.add_Click({
    $logs = $entraAuditExportXlsxButton.Tag
    if (-not $logs -or $logs.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("No audit logs to export.", "No Data", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information); return }
    $outputFolder = $entraOutputFolderTextBox.Text
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $csvFilePath = Join-Path $outputFolder "EntraAuditLogs_$timestamp.csv"
    $xlsxFilePath = Join-Path $outputFolder "EntraAuditLogs_$timestamp.xlsx"
    $logs | Export-Csv -Path $csvFilePath -NoTypeInformation -Encoding UTF8
    if (Format-InboxRuleXlsx -CsvPath $csvFilePath -XlsxPath $xlsxFilePath) {
        try { Remove-Item $csvFilePath -Force } catch {}
        $entraOpenFileButton.Tag = $xlsxFilePath
        $entraOpenFileButton.Enabled = $true
        $script:lastExportedXlsxPath = $xlsxFilePath # Update the script-level variable
        [System.Windows.Forms.MessageBox]::Show("Exported and formatted audit logs to:\n$xlsxFilePath", "XLSX Export OK", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    } else {
        [System.Windows.Forms.MessageBox]::Show("CSV Exported to:\n$csvFilePath\n\nXLSX/Format FAILED. Check Excel install & console.", "XLSX Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    }
})
$entraMfaFetchButton.add_Click({
    $entraUserGrid.EndEdit()
    Write-Host 'EntraUserGrid Columns:'
    foreach ($col in $entraUserGrid.Columns) { Write-Host $col.Name }
    $selectedUpns = @()
    for ($i = 0; $i -lt $entraUserGrid.Rows.Count; $i++) {
        if ($entraUserGrid.Rows[$i].Cells["Select"].Value -eq $true) {
            $upn = $entraUserGrid.Rows[$i].Cells["UserPrincipalName"].Value
            if (-not [string]::IsNullOrWhiteSpace($upn)) { $selectedUpns += $upn }
        }
    }
    Write-Host "Selected UPNs: $($selectedUpns -join ', ')"
    if ($selectedUpns.Count -ne 1) {
        [System.Windows.Forms.MessageBox]::Show("Select exactly one user with a valid UPN.", "Select One User", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information); return
    }
    $upn = $selectedUpns[0]
    $statusLabel.Text = "Analyzing MFA..."; $mainForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        $result = Get-EntraUserMfaStatus -UserPrincipalName $upn
        $details = "MFA Analysis for: $upn`r`n" +
            "-----------------------------`r`nPer-User MFA: $($result.PerUserMfa.Details)`r`nSecurity Defaults: $($result.SecurityDefaults.Details)`r`nConditional Access: $($result.ConditionalAccess.Details)`r`n-----------------------------`r`nOverall Status: $($result.OverallStatus)"
        $form = New-Object System.Windows.Forms.Form
        $form.Text = "MFA Analysis"
        $form.Size = New-Object System.Drawing.Size(600, 400)
        $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
        $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
        $form.MaximizeBox = $true
        $textbox = New-Object System.Windows.Forms.TextBox
        $textbox.Multiline = $true
        $textbox.ReadOnly = $true
        $textbox.ScrollBars = 'Both'
        $textbox.Dock = 'Fill'
        $textbox.Font = New-Object System.Drawing.Font('Consolas', 10)
        $textbox.Text = $details
        $form.Controls.Add($textbox)
        $form.ShowDialog($mainForm)
        $form.Dispose()
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error analyzing MFA: $($_.Exception.Message)", "MFA Analysis Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    } finally { $mainForm.Cursor = [System.Windows.Forms.Cursors]::Default }
})

# --- Export Sign-in Logs button: fetch and export in one click ---
$entraExportSignInLogsButton.add_Click({
    $entraUserGrid.EndEdit()
    $selectedUpns = @()
    for ($i = 0; $i -lt $entraUserGrid.Rows.Count; $i++) {
        if ($entraUserGrid.Rows[$i].Cells["Select"].Value -eq $true) {
            $upn = $entraUserGrid.Rows[$i].Cells["UserPrincipalName"].Value
            if (-not [string]::IsNullOrWhiteSpace($upn)) { $selectedUpns += $upn }
        }
    }
    if ($selectedUpns.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Select at least one user with a valid UPN.", "No Users Selected", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning); return
    }
    $statusLabel.Text = "Fetching and exporting sign-in logs..."; $mainForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $days = $entraSignInDaysUpDown.Value
    $outputFolder = $entraOutputFolderTextBox.Text
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $csvFilePath = Join-Path $outputFolder "EntraSignInLogs_$timestamp.csv"
    $xlsxFilePath = Join-Path $outputFolder "EntraSignInLogs_$timestamp.xlsx"
    try {
        $allLogs = Get-EntraSignInLogs -UserPrincipalNames $selectedUpns -Days $days
        if (-not $allLogs -or $allLogs.Count -eq 0) {
            $statusLabel.Text = "No sign-in logs found."; $mainForm.Cursor = [System.Windows.Forms.Cursors]::Default
            [System.Windows.Forms.MessageBox]::Show("No sign-in logs found for selected users.", "No Logs", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information); return
        }
        $allLogs | Export-Csv -Path $csvFilePath -NoTypeInformation -Encoding UTF8
        if (Format-InboxRuleXlsx -CsvPath $csvFilePath -XlsxPath $xlsxFilePath) {
            try { Remove-Item $csvFilePath -Force } catch {}
            $script:lastExportedXlsxPath = $xlsxFilePath
            $statusLabel.Text = "Exported and formatted sign-in logs to $xlsxFilePath"
            [System.Windows.Forms.MessageBox]::Show("Exported and formatted sign-in logs to:\n$xlsxFilePath", "XLSX Export OK", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        } else {
            [System.Windows.Forms.MessageBox]::Show("CSV Exported to:\n$csvFilePath\n\nXLSX/Format FAILED. Check Excel install & console.", "XLSX Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
    } catch {
        $statusLabel.Text = "Error during sign-in log export: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Error during sign-in log export: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    } finally { $mainForm.Cursor = [System.Windows.Forms.Cursors]::Default }
})

# --- Export Audit Logs button: fetch and export in one click ---
$entraExportAuditLogsButton.add_Click({
    $entraUserGrid.EndEdit()
    $selectedUpns = @()
    for ($i = 0; $i -lt $entraUserGrid.Rows.Count; $i++) {
        if ($entraUserGrid.Rows[$i].Cells["Select"].Value -eq $true) {
            $upn = $entraUserGrid.Rows[$i].Cells["UserPrincipalName"].Value
            if (-not [string]::IsNullOrWhiteSpace($upn)) { $selectedUpns += $upn }
        }
    }
    if ($selectedUpns.Count -ne 1) {
        [System.Windows.Forms.MessageBox]::Show("Select exactly one user with a valid UPN.", "Select One User", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information); return
    }
    $upn = $selectedUpns[0]
    $days = $entraSignInDaysUpDown.Value
    $statusLabel.Text = "Fetching and exporting audit logs..."; $mainForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $outputFolder = $entraOutputFolderTextBox.Text
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $csvFilePath = Join-Path $outputFolder "EntraAuditLogs_$timestamp.csv"
    $xlsxFilePath = Join-Path $outputFolder "EntraAuditLogs_$timestamp.xlsx"
    try {
        $logs = Get-EntraUserAuditLogs -UserPrincipalName $upn -Days $days
        if (-not $logs -or $logs.Count -eq 0) {
            $statusLabel.Text = "No audit logs found."; $mainForm.Cursor = [System.Windows.Forms.Cursors]::Default
            [System.Windows.Forms.MessageBox]::Show("No audit logs found for $upn.", "Audit Logs", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information); return
        }
        $logs | Export-Csv -Path $csvFilePath -NoTypeInformation -Encoding UTF8
        if (Format-InboxRuleXlsx -CsvPath $csvFilePath -XlsxPath $xlsxFilePath) {
            try { Remove-Item $csvFilePath -Force } catch {}
            $script:lastExportedXlsxPath = $xlsxFilePath
            $statusLabel.Text = "Exported and formatted audit logs to $xlsxFilePath"
            [System.Windows.Forms.MessageBox]::Show("Exported and formatted audit logs to:\n$xlsxFilePath", "XLSX Export OK", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        } else {
            [System.Windows.Forms.MessageBox]::Show("CSV Exported to:\n$csvFilePath\n\nXLSX/Format FAILED. Check Excel install & console.", "XLSX Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
    } catch {
        $statusLabel.Text = "Error during audit log export: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Error during audit log export: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    } finally { $mainForm.Cursor = [System.Windows.Forms.Cursors]::Default }
})

# Remove or disable the intermediate Export XLSX buttons
$entraSignInExportXlsxButton.Visible = $false
$entraAuditExportXlsxButton.Visible = $false

# --- Exchange Online Tab Event Handlers ---
$connectButton.add_Click({
    Show-Progress -message "Connecting to Exchange Online..." -progress 0
    $mainForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    
    # Bring main form to front and focus it to ensure auth dialog appears on top
    $mainForm.BringToFront()
    $mainForm.Focus()
    [System.Windows.Forms.Application]::DoEvents()
    
    # Show user-friendly message about authentication
    $statusLabel.Text = "Authentication window should appear. If not visible, check your taskbar or Alt+Tab to find it."
    Show-Progress -message "Authentication window should appear. If not visible, check your taskbar or Alt+Tab to find it." -progress 10
    
    try {
        # Set up authentication with better window focus handling
        $authParams = @{
            ErrorAction = 'Stop'
            ShowBanner = $false  # Reduce visual clutter
        }
        
        # Connect with authentication
        Connect-ExchangeOnline @authParams
        
        $script:currentExchangeConnection = $true
        Show-Progress -message "Connected to Exchange Online successfully." -progress 100
        
        # Enable load buttons and disable connect button
        $loadAllMailboxesButton.Enabled = $true
        $searchMailboxesButton.Enabled = $true
        $disconnectButton.Enabled = $true
        $connectButton.Enabled = $false
        
        Write-Host "Exchange Online connected. Load buttons enabled: LoadAll=$($loadAllMailboxesButton.Enabled), Search=$($searchMailboxesButton.Enabled)"
        
        $statusLabel.Text = "Connected to Exchange Online. Use 'Load All Mailboxes' or 'Search Mailboxes' to load data."
    } catch {
        # Check if this is a user cancellation (common error messages when user cancels auth)
        $errorMessage = $_.Exception.Message
        $isUserCancellation = $errorMessage -match "User cancelled|Operation cancelled|User canceled|Authentication cancelled|Authentication canceled" -or 
                             $errorMessage -match "AADSTS50020|AADSTS50076|AADSTS50079" -or
                             $errorMessage -match "The user cancelled the authentication"
        
        if ($isUserCancellation) {
            # User cancelled - just update status without showing error popup
            $statusLabel.Text = "Exchange Online connection cancelled by user."
            Show-Progress -message "Connection cancelled." -progress -1
        } else {
            # Real error - show user-friendly error message
            Show-UserFriendlyError -errorObject $_ -operation "Exchange Online connection"
        }
    } finally { 
        $mainForm.Cursor = [System.Windows.Forms.Cursors]::Default 
    }
})

# Load All Mailboxes button handler
$loadAllMailboxesButton.add_Click({
    if (-not $script:currentExchangeConnection) {
        [System.Windows.Forms.MessageBox]::Show("Please connect to Exchange Online first.", "Not Connected", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    try {
        $statusLabel.Text = "Loading all mailboxes..."
        $mainForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        
        # Load all mailboxes with full analysis
        $mailboxCount = Load-MailboxesOptimized -MaxMailboxes 10000 -LoadAll
        
        $statusLabel.Text = "Loaded $mailboxCount mailboxes"
        [System.Windows.Forms.MessageBox]::Show("Successfully loaded $mailboxCount mailboxes.", "Load Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        
    } catch {
        $statusLabel.Text = "Error loading mailboxes: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Error loading mailboxes: $($_.Exception.Message)", "Load Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    } finally {
        $mainForm.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})

# Search Mailboxes button handler
$searchMailboxesButton.add_Click({
    if (-not $script:currentExchangeConnection) {
        [System.Windows.Forms.MessageBox]::Show("Please connect to Exchange Online first.", "Not Connected", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    $searchTerm = Show-InputDialog -Title "Search Mailboxes" -Prompt "Enter search term (name or email):"
    if ([string]::IsNullOrWhiteSpace($searchTerm)) { return }
    
    Write-Host "Search term entered: '$searchTerm'"
    
    try {
        $statusLabel.Text = "Searching for mailboxes..."
        $mainForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        
        # Search for mailboxes using the search term
        Write-Host "Searching for mailboxes with filter: DisplayName -like '*$searchTerm*' -or UserPrincipalName -like '*$searchTerm*'"
        $mailboxes = Get-Mailbox -Filter "DisplayName -like '*$searchTerm*' -or UserPrincipalName -like '*$searchTerm*'" -ResultSize 100
        
        Write-Host "Found $($mailboxes.Count) mailboxes"
        
        if ($mailboxes.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("No mailboxes found matching '$searchTerm'.", "No Results", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            return
        }
        
        # Load the found mailboxes by updating the script variables and grid
        $userMailboxGrid.Rows.Clear()
        $script:allLoadedMailboxUPNs = @()
        $script:allLoadedMailboxes = $mailboxes
        
        $mailboxCount = 0
        foreach ($mbx in $mailboxes) {
            $script:allLoadedMailboxUPNs += $mbx.UserPrincipalName
            
            # Get user details for sign-in status
            try {
                $user = Get-User -Identity $mbx.UserPrincipalName -ErrorAction SilentlyContinue
                $signInBlocked = if ($user -and $user.AccountDisabled) { "Blocked" } else { "Allowed" }
            } catch {
                $signInBlocked = "Unknown"
            }
            
            # Add row to grid
            $rowIdx = $userMailboxGrid.Rows.Add()
            $userMailboxGrid.Rows[$rowIdx].Cells["Select"].Value = $false
            $userMailboxGrid.Rows[$rowIdx].Cells["UserPrincipalName"].Value = $mbx.UserPrincipalName
            $userMailboxGrid.Rows[$rowIdx].Cells["DisplayName"].Value = $mbx.DisplayName
            $userMailboxGrid.Rows[$rowIdx].Cells["SignInBlocked"].Value = $signInBlocked
            $userMailboxGrid.Rows[$rowIdx].Cells["RecipientType"].Value = $mbx.RecipientTypeDetails
            
            # Initialize analysis values
            $rulesCount = "0"
            $hiddenRules = "0"
            $suspiciousRules = "0"
            $externalForwarding = "Unknown"
            $delegates = "Unknown"
            $fullAccess = "Unknown"
            
            # Perform analysis for UserMailbox type
            if ($mbx.RecipientTypeDetails -eq "UserMailbox") {
                try {
                    $rules = Get-InboxRule -Mailbox $mbx.UserPrincipalName -IncludeHidden -ErrorAction SilentlyContinue
                    if ($rules) {
                        $analysis = Analyze-MailboxRulesEnhanced -Rules $rules -BaseSuspiciousKeywords $BaseSuspiciousKeywords
                        $rulesCount = $analysis.TotalRules.ToString()
                        $hiddenRules = $analysis.SuspiciousHidden.ToString()
                        $suspiciousRules = $analysis.SuspiciousVisible.ToString()
                        $externalForwarding = if ($analysis.HasExternalForwarding) { "Yes" } else { "No" }
                    }
                } catch {
                    # Keep default values if analysis fails
                }
                
                # Analyze permissions
                try {
                    $delegates = Analyze-MailboxDelegates -UserPrincipalName $mbx.UserPrincipalName
                    $fullAccess = Analyze-MailboxPermissions -UserPrincipalName $mbx.UserPrincipalName
                } catch {
                    # Keep default values if analysis fails
                }
            } elseif ($mbx.RecipientTypeDetails -eq "SharedMailbox") {
                $rulesCount = "N/A"
                $hiddenRules = "N/A"
                $suspiciousRules = "N/A"
                $externalForwarding = "N/A"
                # Still analyze permissions for shared mailboxes
                try {
                    $delegates = Analyze-MailboxDelegates -UserPrincipalName $mbx.UserPrincipalName
                    $fullAccess = Analyze-MailboxPermissions -UserPrincipalName $mbx.UserPrincipalName
                } catch {
                    # Keep default values if analysis fails
                }
            }
            
            $userMailboxGrid.Rows[$rowIdx].Cells["TotalRules"].Value = $rulesCount
            $userMailboxGrid.Rows[$rowIdx].Cells["HiddenRules"].Value = $hiddenRules
            $userMailboxGrid.Rows[$rowIdx].Cells["SuspiciousRules"].Value = $suspiciousRules
            $userMailboxGrid.Rows[$rowIdx].Cells["ExternalForwarding"].Value = $externalForwarding
            $userMailboxGrid.Rows[$rowIdx].Cells["Delegates"].Value = $delegates
            $userMailboxGrid.Rows[$rowIdx].Cells["FullAccess"].Value = $fullAccess
            
            $mailboxCount++
        }
        
        $statusLabel.Text = "Loaded $mailboxCount mailboxes matching '$searchTerm'"
        [System.Windows.Forms.MessageBox]::Show("Found and loaded $mailboxCount mailboxes matching '$searchTerm'.", "Search Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        
    } catch {
        $statusLabel.Text = "Error searching mailboxes: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Error searching mailboxes: $($_.Exception.Message)", "Search Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    } finally {
        $mainForm.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})
$disconnectButton.add_Click({
    $statusLabel.Text = "Disconnecting..."; $mainForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try { Disconnect-ExchangeOnline -Confirm:$false } catch {}
    $script:currentExchangeConnection = $null
    $userMailboxGrid.Rows.Clear(); $selectAllButton.Enabled = $false; $deselectAllButton.Enabled = $false; $disconnectButton.Enabled = $false; $connectButton.Enabled = $true
    $loadAllMailboxesButton.Enabled = $false; $searchMailboxesButton.Enabled = $false
    $manageRulesButton.Enabled = $false; $manageConnectorsButton.Enabled = $false; $manageTransportRulesButton.Enabled = $false
    $statusLabel.Text = "Disconnected."
    $mainForm.Cursor = [System.Windows.Forms.Cursors]::Default
})
$selectAllButton.add_Click({ for ($i = 0; $i -lt $userMailboxGrid.Rows.Count; $i++) { $userMailboxGrid.Rows[$i].Cells["Select"].Value = $true } })
$deselectAllButton.add_Click({ for ($i = 0; $i -lt $userMailboxGrid.Rows.Count; $i++) { $userMailboxGrid.Rows[$i].Cells["Select"].Value = $false } })
$browseFolderButton.add_Click({ 
    $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog; 
    if ($folderDialog.ShowDialog() -eq 'OK') { 
        $outputFolderTextBox.Text = $folderDialog.SelectedPath 
    } 
})

# Search functionality
$exchangeSearchTextBox.add_TextChanged({
    Filter-ExchangeGrid -searchText $exchangeSearchTextBox.Text
})

# Add search functionality for Entra ID
$entraSearchTextBox.add_TextChanged({
    Filter-EntraGrid -searchText $entraSearchTextBox.Text
})
$getRulesButton.add_Click({
    $selectedUpns = @()
    for ($i = 0; $i -lt $userMailboxGrid.Rows.Count; $i++) {
        if ($userMailboxGrid.Rows[$i].Cells["Select"].Value -eq $true) {
            $selectedUpns += $userMailboxGrid.Rows[$i].Cells["UserPrincipalName"].Value
        }
    }
    if ($selectedUpns.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Select at least one mailbox.", "No Mailbox Selected", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning); return
    }
    $statusLabel.Text = "Analyzing inbox rules..."; $mainForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $outputFolder = $outputFolderTextBox.Text
    if ([string]::IsNullOrWhiteSpace($outputFolder)) {
        [System.Windows.Forms.MessageBox]::Show("Please select an output folder before analyzing rules.", "Output Folder Required", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        $mainForm.Cursor = [System.Windows.Forms.Cursors]::Default
        return
    }
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $csvFilePath = Join-Path $outputFolder "InboxRules_$timestamp.csv"
    $xlsxFilePath = Join-Path $outputFolder "InboxRules_$timestamp.xlsx"
    $allRuleData = @()
    try {
        foreach ($upn in $selectedUpns) {
            $rules = Get-InboxRule -Mailbox $upn -IncludeHidden -ErrorAction SilentlyContinue
            if ($rules) {
                foreach ($rule in $rules) {
                    $matchedKeywords = @()
                    foreach ($kw in $BaseSuspiciousKeywords) {
                        if ($rule.Name -and $rule.Name -match [regex]::Escape($kw)) {
                            $matchedKeywords += $kw
                        }
                    }
                    
                    # Check for symbols-only rule names (no text characters)
                    $isSymbolsOnly = $false
                    if ($rule.Name -and $rule.Name.Length -gt 0) {
                        $textCharacters = $rule.Name -replace '[^\p{L}\p{N}\s]', ''  # Remove all non-text characters
                        $isSymbolsOnly = [string]::IsNullOrWhiteSpace($textCharacters)
                    }
                    
                    # Check if rule is hidden
                    $isHidden = $rule.IsHidden
                    
                    # Determine if rule is suspicious based on new criteria
                    $isSuspicious = $false
                    $suspiciousReasons = @()
                    
                    if ($matchedKeywords.Count -gt 0) {
                        $isSuspicious = $true
                        $suspiciousReasons += "Contains suspicious keywords: $($matchedKeywords -join ', ')"
                    }
                    
                    if ($isSymbolsOnly) {
                        $isSuspicious = $true
                        $suspiciousReasons += "Symbols-only name (no text characters)"
                    }
                    
                    if ($isHidden) {
                        $isSuspicious = $true
                        $suspiciousReasons += "Hidden rule"
                    }
                    
                    $allRuleData += [PSCustomObject]@{
                        MailboxOwner                = $upn
                        RuleName                    = $rule.Name
                        IsEnabled                   = $rule.Enabled
                        Priority                    = $rule.Priority
                        IsHidden                    = $rule.IsHidden
                        IsSymbolsOnly               = $isSymbolsOnly
                        IsSuspicious                = $isSuspicious
                        SuspiciousReasons           = ($suspiciousReasons -join '; ')
                        IsForwardingExternal        = [bool]($rule.ForwardTo -match '@')
                        IsDeleting                  = $rule.DeleteMessage
                        IsMarkingAsRead             = $rule.MarkAsRead
                        IsMovingToFolder            = [bool]$rule.MoveToFolder
                        MoveToFolderName            = $rule.MoveToFolder
                        SuspiciousKeywordsInName    = ($matchedKeywords -join ', ')
                        Description                 = $rule.Description
                        StopProcessingRules         = $rule.StopProcessingRules
                        Conditions                  = $rule.Conditions
                        Actions                     = $rule.Actions
                        Exceptions                  = $rule.Exceptions
                        RuleID                      = "'$($rule.RuleIdentity)"
                    }
                }
            }
        }
        if ($allRuleData.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("No inbox rules found for selected mailboxes.", "No Rules", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information); return
        }
        $exchangeGrid.DataSource = $null
        $exchangeGrid.DataSource = $allRuleData
        $allRuleData | Export-Csv -Path $csvFilePath -NoTypeInformation -Encoding UTF8
        if (Format-InboxRuleXlsx -CsvPath $csvFilePath -XlsxPath $xlsxFilePath) {
            try { Remove-Item $csvFilePath -Force } catch {}
            $openFileButton.Tag = $xlsxFilePath
            $openFileButton.Enabled = $true
            $statusLabel.Text = "Exported and formatted inbox rules to $xlsxFilePath"
            [System.Windows.Forms.MessageBox]::Show("Exported and formatted inbox rules to:\n$xlsxFilePath", "XLSX Export OK", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        } else {
            $statusLabel.Text = "CSV OK, XLSX/Format Failed."; $openFileButton.Enabled = $false
            [System.Windows.Forms.MessageBox]::Show("CSV Exported to:\n$csvFilePath\n\nXLSX/Format FAILED. Check Excel install & console.", "XLSX Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
    } catch {
        $statusLabel.Text = "Error during analysis: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Error during analysis: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    } finally { $mainForm.Cursor = [System.Windows.Forms.Cursors]::Default }
})
$manageRulesButton.add_Click({
    $checkedRows = @()
    for ($i = 0; $i -lt $userMailboxGrid.Rows.Count; $i++) {
        if ($userMailboxGrid.Rows[$i].Cells["Select"].Value -eq $true) {
            $checkedRows += $userMailboxGrid.Rows[$i]
        }
    }
    if ($checkedRows.Count -ne 1) {
        [System.Windows.Forms.MessageBox]::Show("Select exactly one mailbox to manage rules.", "Select One Mailbox", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information); return
    }
    $upn = $checkedRows[0].Cells[1].Value
    $rulesForm = New-Object System.Windows.Forms.Form
    $rulesForm.Text = "Manage Inbox Rules for $upn"
    $rulesForm.Size = New-Object System.Drawing.Size(900, 500)
    $rulesForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
    $rulesForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
    $rulesForm.MaximizeBox = $true

    # Create a new DataGridView for rules each time
    $rulesGrid = New-Object System.Windows.Forms.DataGridView
    $rulesGrid.Dock = 'Fill'
    $rulesGrid.ReadOnly = $true
    $rulesGrid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $rulesGrid.AutoGenerateColumns = $true
    $rulesGrid.AllowUserToAddRows = $false
    $rulesGrid.AutoSizeColumnsMode = 'Fill'

    # Panel for buttons
    $buttonPanel = New-Object System.Windows.Forms.Panel
    $buttonPanel.Dock = 'Bottom'
    $buttonPanel.Height = 40

    $deleteButton = New-Object System.Windows.Forms.Button
    $deleteButton.Text = "Delete Selected Rule(s)"
    $deleteButton.Dock = 'Left'
    $closeButton = New-Object System.Windows.Forms.Button
    $closeButton.Text = "Close"
    $closeButton.Dock = 'Right'
    $buttonPanel.Controls.Add($deleteButton)
    $buttonPanel.Controls.Add($closeButton)

    $rulesForm.Controls.Add($rulesGrid)
    $rulesForm.Controls.Add($buttonPanel)

    # Load rules
    $rules = Get-InboxRule -Mailbox $upn -IncludeHidden -ErrorAction SilentlyContinue
    if ($rules -and $rules.Count -gt 0) {
        $displayRules = foreach ($rule in $rules) {
            [PSCustomObject]@{
                Name = $rule.Name
                Enabled = $rule.Enabled
                Priority = $rule.Priority
                RuleIdentity = "$($rule.RuleIdentity)"  # Force string to avoid scientific notation
            }
        }
        # Convert to DataTable
        $dt = New-Object System.Data.DataTable
        if ($displayRules.Count -gt 0) {
            $displayRules[0].psobject.Properties.Name | ForEach-Object { [void]$dt.Columns.Add($_) }
            foreach ($row in $displayRules) {
                $dt.Rows.Add($row.psobject.Properties.Value)
            }
        }
        $rulesGrid.DataSource = $dt
        $rulesGrid.DataSource = $dt
        $rulesGrid.AutoSizeColumnsMode = 'Fill'
        foreach ($col in $rulesGrid.Columns) { $col.AutoSizeMode = 'Fill' }
    } else {
        $rulesGrid.DataSource = $null
    }

    $deleteButton.add_Click({
        if (-not $rulesGrid.SelectedRows -or $rulesGrid.SelectedRows.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Select at least one rule to delete.", "No Rule Selected", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning); return
        }
        $selectedNames = @()
        foreach ($row in $rulesGrid.SelectedRows) {
            $selectedNames += $row.Cells["Name"].Value
        }
        $confirm = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to delete the selected rule(s)?\n" + ($selectedNames -join "\n"), "Confirm Delete", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        foreach ($ruleName in $selectedNames) {
            try {
                Remove-InboxRule -Mailbox $upn -Identity $ruleName -Confirm:$false -ErrorAction Stop
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Failed to delete rule: $ruleName`n$($_.Exception.Message)", "Delete Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }
        # Reload rules
        $rules = Get-InboxRule -Mailbox $upn -IncludeHidden -ErrorAction SilentlyContinue
        if ($rules -and $rules.Count -gt 0) {
            $displayRules = foreach ($rule in $rules) {
                [PSCustomObject]@{
                    Name = $rule.Name
                    Enabled = $rule.Enabled
                    Priority = $rule.Priority
                    RuleIdentity = "$($rule.RuleIdentity)"  # Force string to avoid scientific notation
                }
            }
            # Convert to DataTable
            $dt = New-Object System.Data.DataTable
            if ($displayRules.Count -gt 0) {
                $displayRules[0].psobject.Properties.Name | ForEach-Object { [void]$dt.Columns.Add($_) }
                foreach ($row in $displayRules) {
                    $dt.Rows.Add($row.psobject.Properties.Value)
                }
            }
            $rulesGrid.DataSource = $dt
            $rulesGrid.AutoSizeColumnsMode = 'Fill'
            foreach ($col in $rulesGrid.Columns) { $col.AutoSizeMode = 'Fill' }
        } else {
            $rulesGrid.DataSource = $null
        }
    })
    $closeButton.add_Click({ $rulesForm.Close() })
    [void]$rulesForm.ShowDialog($mainForm)
    $rulesForm.Dispose()
})

# Add click handler for analyze selected button
$analyzeSelectedButton.add_Click({
    $selectedUpns = @()
    $selectedRows = @()
    
    for ($i = 0; $i -lt $userMailboxGrid.Rows.Count; $i++) {
        if ($userMailboxGrid.Rows[$i].Cells["Select"].Value -eq $true) {
            $upn = $userMailboxGrid.Rows[$i].Cells["UserPrincipalName"].Value
            $displayName = $userMailboxGrid.Rows[$i].Cells["DisplayName"].Value
            $mailboxType = $userMailboxGrid.Rows[$i].Cells["RecipientType"].Value
            
            if (-not [string]::IsNullOrWhiteSpace($upn)) {
                $selectedUpns += $upn
                $selectedRows += @{
                    Index = $i
                    UPN = $upn
                    DisplayName = $displayName
                    Type = $mailboxType
                }
            }
        }
    }
    
    if ($selectedUpns.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Select at least one mailbox for detailed analysis.", "No Mailbox Selected", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    $statusLabel.Text = "Performing detailed analysis for selected mailboxes..."
    $mainForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    
    try {
        $processedCount = 0
        foreach ($selectedRow in $selectedRows) {
            $processedCount++
            $upn = $selectedRow.UPN
            $rowIndex = $selectedRow.Index
            $mailboxType = $selectedRow.Type
            
            $statusLabel.Text = "Analyzing mailbox $processedCount of $($selectedRows.Count): $upn"
            $mainForm.Refresh()
            
            # Analyze rules only for user mailboxes (shared mailboxes don't have user-created inbox rules)
            if ($mailboxType -eq "UserMailbox") {
                try {
                    $rules = Get-InboxRule -Mailbox $upn -IncludeHidden -ErrorAction SilentlyContinue
                    if ($rules) {
                        $analysis = Analyze-MailboxRulesEnhanced -Rules $rules -BaseSuspiciousKeywords $BaseSuspiciousKeywords
                        $userMailboxGrid.Rows[$rowIndex].Cells["TotalRules"].Value = $analysis.TotalRules.ToString()
                        $userMailboxGrid.Rows[$rowIndex].Cells["HiddenRules"].Value = $analysis.SuspiciousHidden.ToString()
                        $userMailboxGrid.Rows[$rowIndex].Cells["SuspiciousRules"].Value = $analysis.SuspiciousVisible.ToString()
                        $userMailboxGrid.Rows[$rowIndex].Cells["ExternalForwarding"].Value = if ($analysis.HasExternalForwarding) { "Yes" } else { "No" }
                    }
                } catch {
                    # Keep existing values if analysis fails
                }
            } elseif ($mailboxType -eq "SharedMailbox") {
                # Shared mailboxes can't have user-created inbox rules or external forwarding
                $userMailboxGrid.Rows[$rowIndex].Cells["TotalRules"].Value = "N/A"
                $userMailboxGrid.Rows[$rowIndex].Cells["HiddenRules"].Value = "N/A"
                $userMailboxGrid.Rows[$rowIndex].Cells["SuspiciousRules"].Value = "N/A"
                $userMailboxGrid.Rows[$rowIndex].Cells["ExternalForwarding"].Value = "N/A"
            }
            
            # Analyze permissions for all mailbox types (shared mailboxes can have permissions)
            try {
                $delegates = Analyze-MailboxDelegates -UserPrincipalName $upn
                $fullAccess = Analyze-MailboxPermissions -UserPrincipalName $upn
                $userMailboxGrid.Rows[$rowIndex].Cells["Delegates"].Value = $delegates
                $userMailboxGrid.Rows[$rowIndex].Cells["FullAccess"].Value = $fullAccess
            } catch {
                # Keep existing values if analysis fails
            }
        }
        
        $statusLabel.Text = "Detailed analysis completed for $($selectedRows.Count) mailboxes"
        
    } catch {
        $statusLabel.Text = "Error during detailed analysis: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Error during detailed analysis: $($_.Exception.Message)", "Analysis Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    } finally {
        $mainForm.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})
$openFileButton.add_Click({
    if ($openFileButton.Tag -and (Test-Path $openFileButton.Tag)) {
        try { Invoke-Item -Path $openFileButton.Tag -ErrorAction Stop } catch { [System.Windows.Forms.MessageBox]::Show("Could not open file: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) }
    } else {
        [System.Windows.Forms.MessageBox]::Show("No file exported or file not found.", "No File", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
})

$manageConnectorsButton.add_Click({
    Show-ConnectorsViewer -mainForm $mainForm -statusLabel $statusLabel
})

$manageTransportRulesButton.add_Click({
    Show-TransportRulesViewer -mainForm $mainForm -statusLabel $statusLabel
})



$userMailboxGrid.add_CellContentClick({
    $mainForm.BeginInvoke([System.Action]{
        $manageRulesButton.Enabled = $true
        $manageConnectorsButton.Enabled = $true
        $manageTransportRulesButton.Enabled = $true
        $checkedCount = 0
        for ($i = 0; $i -lt $userMailboxGrid.Rows.Count; $i++) {
            if ($userMailboxGrid.Rows[$i].Cells["Select"].Value -eq $true) { $checkedCount++ }
        }
    })
})

# --- After all Entra tab buttons and panels are created ---



# Activate View Sign-in Logs button
$entraViewSignInLogsButton.add_Click({
    $entraUserGrid.EndEdit()
    $selectedUpns = @()
    for ($i = 0; $i -lt $entraUserGrid.Rows.Count; $i++) {
        if ($entraUserGrid.Rows[$i].Cells["Select"].Value -eq $true) {
            $upn = $entraUserGrid.Rows[$i].Cells["UserPrincipalName"].Value
            if (-not [string]::IsNullOrWhiteSpace($upn)) { $selectedUpns += $upn }
        }
    }
    if ($selectedUpns.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Select at least one user with a valid UPN.", "No Users Selected", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning); return
    }
    $days = $entraSignInDaysUpDown.Value
    try {
        $allLogs = @(Get-EntraSignInLogs -UserPrincipalNames $selectedUpns -Days $days)
        if (-not $allLogs -or $allLogs.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("No sign-in logs found for selected users.", "No Logs", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information); return
        }
        # Flatten logs for DataGridView
        $data = foreach ($log in $allLogs) {
            [PSCustomObject]@{
                UserPrincipalName = $log.UserPrincipalName
                CreatedDateTime   = $log.CreatedDateTime
                AppDisplayName    = $log.AppDisplayName
                IPAddress         = $log.IPAddress
                Location          = if ($log.Location) { ($log.Location.City + ', ' + $log.Location.State + ', ' + $log.Location.CountryOrRegion) } else { '' }
                Status            = if ($log.Status) { $log.Status.AdditionalDetails } else { '' }
                Device            = if ($log.DeviceDetail) { ($log.DeviceDetail.Browser + ' / ' + $log.DeviceDetail.OperatingSystem) } else { '' }
                RiskLevelAggregated = $log.RiskLevelAggregated
                ConditionalAccessStatus = $log.ConditionalAccessStatus
            }
        }
        # Convert to DataTable for DataGridView
        $dt = New-Object System.Data.DataTable
        if ($data.Count -gt 0) {
            $data[0].psobject.Properties.Name | ForEach-Object { [void]$dt.Columns.Add($_) }
            foreach ($row in $data) {
                $dt.Rows.Add($row.psobject.Properties.Value)
            }
        }
        $popup = New-Object System.Windows.Forms.Form
        $popup.Text = "Sign-in Logs"
        $popup.Size = New-Object System.Drawing.Size(900, 600)
        $popup.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
        $popup.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
        $popup.MaximizeBox = $true
        $grid = New-Object System.Windows.Forms.DataGridView
        $grid.Dock = 'Fill'
        $grid.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
        $grid.ReadOnly = $true
        $grid.AllowUserToAddRows = $false
        $grid.AutoGenerateColumns = $true
        $grid.AutoSizeColumnsMode = 'Fill'
        $grid.MinimumSize = New-Object System.Drawing.Size(800, 400)
        $grid.DataSource = $dt
        $popup.Controls.Add($grid)
        $popup.ShowDialog($mainForm)
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error fetching sign-in logs: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

# Activate View Audit Logs button
$entraViewAuditLogsButton.add_Click({
    $entraUserGrid.EndEdit()
    $selectedUpns = @()
    for ($i = 0; $i -lt $entraUserGrid.Rows.Count; $i++) {
        if ($entraUserGrid.Rows[$i].Cells["Select"].Value -eq $true) {
            $upn = $entraUserGrid.Rows[$i].Cells["UserPrincipalName"].Value
            if (-not [string]::IsNullOrWhiteSpace($upn)) { $selectedUpns += $upn }
        }
    }
    if ($selectedUpns.Count -ne 1) {
        [System.Windows.Forms.MessageBox]::Show("Select exactly one user with a valid UPN.", "Select One User", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information); return
    }
    $upn = $selectedUpns[0]
    $days = $entraSignInDaysUpDown.Value
    try {
        $logs = @(Get-EntraUserAuditLogs -UserPrincipalName $upn -Days $days)
        if (-not $logs -or $logs.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("No audit logs found for $upn.", "Audit Logs", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information); return
        }
        # Flatten logs for DataGridView
        $data = foreach ($log in $logs) {
            [PSCustomObject]@{
                ActivityDisplayName = $log.ActivityDisplayName
                ActivityDateTime    = $log.ActivityDateTime
                InitiatedBy         = if ($log.InitiatedBy -and $log.InitiatedBy.User) { $log.InitiatedBy.User.UserPrincipalName } else { '' }
                TargetResources     = if ($log.TargetResources) { ($log.TargetResources | ForEach-Object { $_.UserPrincipalName }) -join ", " } else { '' }
                Category            = $log.Category
                Result              = $log.Result
                CorrelationId       = $log.CorrelationId
                LoggedByService     = $log.LoggedByService
                OperationType       = $log.OperationType
                UserPrincipalName   = $log.UserPrincipalName
                IPAddress           = $log.IPAddress
            }
        }
        # Convert to DataTable for DataGridView
        $dt = New-Object System.Data.DataTable
        if ($data.Count -gt 0) {
            $data[0].psobject.Properties.Name | ForEach-Object { [void]$dt.Columns.Add($_) }
            foreach ($row in $data) {
                $dt.Rows.Add($row.psobject.Properties.Value)
            }
        }
        $popup = New-Object System.Windows.Forms.Form
        $popup.Text = "Audit Logs"
        $popup.Size = New-Object System.Drawing.Size(900, 600)
        $popup.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
        $popup.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
        $popup.MaximizeBox = $true
        $grid = New-Object System.Windows.Forms.DataGridView
        $grid.Dock = 'Fill'
        $grid.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
        $grid.ReadOnly = $true
        $grid.AllowUserToAddRows = $false
        $grid.AutoGenerateColumns = $true
        $grid.AutoSizeColumnsMode = 'Fill'
        $grid.MinimumSize = New-Object System.Drawing.Size(800, 400)
        $grid.DataSource = $dt
        $popup.Controls.Add($grid)
        $popup.ShowDialog($mainForm)
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error fetching audit logs: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

# Bottom panel controls are already set during layout creation

# Add tabs to the tab control
$tabControl.TabPages.Add($exchangeTab)
$tabControl.TabPages.Add($entraTab)

# --- Report Generator Tab ---
$reportGeneratorTab = New-Object System.Windows.Forms.TabPage
$reportGeneratorTab.Text = "Report Generator"

# Create Report Generator tab layout
$reportGeneratorPanel = New-Object System.Windows.Forms.Panel
$reportGeneratorPanel.Dock = 'Fill'
$reportGeneratorPanel.Padding = New-Object System.Windows.Forms.Padding(10)
$reportGeneratorPanel.AutoScroll = $true

# Title label
$reportGeneratorTitleLabel = New-Object System.Windows.Forms.Label
$reportGeneratorTitleLabel.Text = "Professional Report Generator"
$reportGeneratorTitleLabel.Font = New-Object System.Drawing.Font('Segoe UI', 14, [System.Drawing.FontStyle]::Bold)
$reportGeneratorTitleLabel.Location = New-Object System.Drawing.Point(10, 10)
$reportGeneratorTitleLabel.Size = New-Object System.Drawing.Size(400, 30)
$reportGeneratorPanel.Controls.Add($reportGeneratorTitleLabel)

# Description label
$reportGeneratorDescLabel = New-Object System.Windows.Forms.Label
$reportGeneratorDescLabel.Text = "Generate professional reports combining Exchange Online and Entra ID data for support tickets or documentation."
$reportGeneratorDescLabel.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$reportGeneratorDescLabel.Location = New-Object System.Drawing.Point(10, 45)
$reportGeneratorDescLabel.Size = New-Object System.Drawing.Size(600, 40)
$reportGeneratorDescLabel.ForeColor = [System.Drawing.Color]::DarkGray
$reportGeneratorPanel.Controls.Add($reportGeneratorDescLabel)

# Account Selector Group
$accountSelectorGroup = New-Object System.Windows.Forms.GroupBox
$accountSelectorGroup.Text = "Account Selection"
$accountSelectorGroup.Location = New-Object System.Drawing.Point(10, 90)
$accountSelectorGroup.Size = New-Object System.Drawing.Size(800, 330)
$accountSelectorGroup.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$accountSelectorGroup.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)

# Account selector description
$accountSelectorDescLabel = New-Object System.Windows.Forms.Label
$accountSelectorDescLabel.Text = "Select accounts for unified reporting (combines Exchange Online and Entra ID data):"
$accountSelectorDescLabel.Location = New-Object System.Drawing.Point(10, 25)
$accountSelectorDescLabel.Size = New-Object System.Drawing.Size(760, 20)
$accountSelectorDescLabel.Font = New-Object System.Drawing.Font('Segoe UI', 8)
$accountSelectorGroup.Controls.Add($accountSelectorDescLabel)

# Unified account grid
$unifiedAccountGrid = New-Object System.Windows.Forms.DataGridView
$unifiedAccountGrid.Location = New-Object System.Drawing.Point(10, 50)
$unifiedAccountGrid.Size = New-Object System.Drawing.Size(760, 280)
$unifiedAccountGrid.ReadOnly = $false
$unifiedAccountGrid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$unifiedAccountGrid.MultiSelect = $true
$unifiedAccountGrid.AllowUserToAddRows = $false
$unifiedAccountGrid.AutoGenerateColumns = $false
$unifiedAccountGrid.RowHeadersVisible = $false
$unifiedAccountGrid.ColumnHeadersVisible = $true
$unifiedAccountGrid.EnableHeadersVisualStyles = $true
$unifiedAccountGrid.ColumnHeadersHeight = 25
$unifiedAccountGrid.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::DisableResizing
$unifiedAccountGrid.AutoSizeColumnsMode = 'Fill'

# Define columns for unified account grid
$colUnifiedCheck = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
$colUnifiedCheck.HeaderText = "Select"
$colUnifiedCheck.Name = "Select"
$colUnifiedCheck.ReadOnly = $false
$unifiedAccountGrid.Columns.Add($colUnifiedCheck)

$colUnifiedUPN = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colUnifiedUPN.HeaderText = "UserPrincipalName"
$colUnifiedUPN.Name = "UserPrincipalName"
$colUnifiedUPN.ReadOnly = $true
$unifiedAccountGrid.Columns.Add($colUnifiedUPN)

$colUnifiedDisplayName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colUnifiedDisplayName.HeaderText = "DisplayName"
$colUnifiedDisplayName.Name = "DisplayName"
$colUnifiedDisplayName.ReadOnly = $true
$unifiedAccountGrid.Columns.Add($colUnifiedDisplayName)

$colUnifiedExchange = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colUnifiedExchange.HeaderText = "Exchange Status"
$colUnifiedExchange.Name = "ExchangeStatus"
$colUnifiedExchange.ReadOnly = $true
$unifiedAccountGrid.Columns.Add($colUnifiedExchange)

$colUnifiedEntra = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colUnifiedEntra.HeaderText = "Entra Status"
$colUnifiedEntra.Name = "EntraStatus"
$colUnifiedEntra.ReadOnly = $true
$unifiedAccountGrid.Columns.Add($colUnifiedEntra)

$accountSelectorGroup.Controls.Add($unifiedAccountGrid)

# Account selector buttons
$refreshAccountsButton = New-Object System.Windows.Forms.Button
$refreshAccountsButton.Text = "🔄 Refresh Account List"
$refreshAccountsButton.Location = New-Object System.Drawing.Point(10, 370)
$refreshAccountsButton.Size = New-Object System.Drawing.Size(180, 35)
$refreshAccountsButton.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
$refreshAccountsButton.BackColor = [System.Drawing.Color]::LightBlue
$refreshAccountsButton.ForeColor = [System.Drawing.Color]::DarkBlue

$selectAllAccountsButton = New-Object System.Windows.Forms.Button
$selectAllAccountsButton.Text = "Select All"
$selectAllAccountsButton.Location = New-Object System.Drawing.Point(200, 370)
$selectAllAccountsButton.Size = New-Object System.Drawing.Size(100, 35)
$selectAllAccountsButton.Font = New-Object System.Drawing.Font('Segoe UI', 9)

$deselectAllAccountsButton = New-Object System.Windows.Forms.Button
$deselectAllAccountsButton.Text = "Deselect All"
$deselectAllAccountsButton.Location = New-Object System.Drawing.Point(310, 370)
$deselectAllAccountsButton.Size = New-Object System.Drawing.Size(100, 35)
$deselectAllAccountsButton.Font = New-Object System.Drawing.Font('Segoe UI', 9)

$accountSelectorGroup.Controls.AddRange(@($refreshAccountsButton, $selectAllAccountsButton, $deselectAllAccountsButton))

# Connection status indicator
$connectionStatusLabel = New-Object System.Windows.Forms.Label
$connectionStatusLabel.Text = "Connection Status: Checking..."
$connectionStatusLabel.Location = New-Object System.Drawing.Point(420, 460)
$connectionStatusLabel.Size = New-Object System.Drawing.Size(350, 35)
$connectionStatusLabel.Font = New-Object System.Drawing.Font('Segoe UI', 8)
$connectionStatusLabel.ForeColor = [System.Drawing.Color]::DarkGray
$accountSelectorGroup.Controls.Add($connectionStatusLabel)

# Function to update connection status
function Update-ConnectionStatus {
    # Robust checks
    $exoConnected = $false
    try { Get-OrganizationConfig -ErrorAction Stop | Out-Null; $exoConnected = $true } catch { $exoConnected = ($script:currentExchangeConnection -eq $true) }
    $mgConnected = $false
    try { $ctx = Get-MgContext -ErrorAction Stop; if ($ctx -and $ctx.Account) { $mgConnected = $true } } catch { $mgConnected = ($script:graphConnection -ne $null) }

    $exchangeStatus = if ($exoConnected) { "✅ Exchange Online" } else { "❌ Exchange Online" }
    $entraStatus = if ($mgConnected) { "✅ Entra ID" } else { "❌ Entra ID" }
    $connectionStatusLabel.Text = "Connection Status: $exchangeStatus | $entraStatus"
    
    if ($exoConnected -and $mgConnected) {
        $connectionStatusLabel.ForeColor = [System.Drawing.Color]::Green
    } elseif ($exoConnected -or $mgConnected) {
        $connectionStatusLabel.ForeColor = [System.Drawing.Color]::Orange
    } else {
        $connectionStatusLabel.ForeColor = [System.Drawing.Color]::Red
    }
}

# Generate Report button (moved down)
$generateReportButton = New-Object System.Windows.Forms.Button
$generateReportButton.Text = "Generate Professional Report"
$generateReportButton.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
$generateReportButton.Location = New-Object System.Drawing.Point(10, 620)
$generateReportButton.Size = New-Object System.Drawing.Size(250, 40)
$generateReportButton.BackColor = [System.Drawing.Color]::LightBlue
$reportGeneratorPanel.Controls.Add($generateReportButton)

# Incident Checklist Button
$incidentChecklistButton = New-Object System.Windows.Forms.Button
$incidentChecklistButton.Text = "Generate Incident Remediation Checklist"
$incidentChecklistButton.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
$incidentChecklistButton.Location = New-Object System.Drawing.Point(270, 620)
$incidentChecklistButton.Size = New-Object System.Drawing.Size(250, 40)
$incidentChecklistButton.BackColor = [System.Drawing.Color]::LightCoral
$reportGeneratorPanel.Controls.Add($incidentChecklistButton)

# Security Investigation Button
$securityInvestigationButton = New-Object System.Windows.Forms.Button
$securityInvestigationButton.Text = "🔍 Security Investigation Report"
$securityInvestigationButton.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
$securityInvestigationButton.Location = New-Object System.Drawing.Point(530, 620)
$securityInvestigationButton.Size = New-Object System.Drawing.Size(250, 40)
$securityInvestigationButton.BackColor = [System.Drawing.Color]::FromArgb(0, 122, 204) # Blue color
$securityInvestigationButton.ForeColor = [System.Drawing.Color]::White
$securityInvestigationButton.add_Click({
    $statusLabel.Text = "🔍 Opening Security Investigation Report..."
    $mainForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor

    try {
        # Create Security Investigation form
        $securityForm = New-Object System.Windows.Forms.Form
        $securityForm.Text = "Security Investigation Report Generator"
        $securityForm.Size = New-Object System.Drawing.Size(900, 700)
        $securityForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
        $securityForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
        $securityForm.MaximizeBox = $true

        # Create main panel
        $securityMainPanel = New-Object System.Windows.Forms.Panel
        $securityMainPanel.Dock = 'Fill'
        $securityMainPanel.Padding = New-Object System.Windows.Forms.Padding(15)

        # Title
        $securityTitleLabel = New-Object System.Windows.Forms.Label
        $securityTitleLabel.Text = "🔍 Comprehensive Security Investigation Report"
        $securityTitleLabel.Font = New-Object System.Drawing.Font('Segoe UI', 16, [System.Drawing.FontStyle]::Bold)
        $securityTitleLabel.Location = New-Object System.Drawing.Point(15, 15)
        $securityTitleLabel.Size = New-Object System.Drawing.Size(500, 35)

        # Description
        $securityDescLabel = New-Object System.Windows.Forms.Label
        $securityDescLabel.Text = "Generate comprehensive security analysis combining Exchange Online and Microsoft Graph data.`nThis report includes audit logs, sign-in patterns, email traces, and inbox rules for complete investigation."
        $securityDescLabel.Font = New-Object System.Drawing.Font('Segoe UI', 9)
        $securityDescLabel.Location = New-Object System.Drawing.Point(15, 55)
        $securityDescLabel.Size = New-Object System.Drawing.Size(600, 40)
        $securityDescLabel.MaximumSize = New-Object System.Drawing.Size(600, 0)
        $securityDescLabel.AutoSize = $true

        # Configuration section
        $configGroupBox = New-Object System.Windows.Forms.GroupBox
        $configGroupBox.Text = "Investigation Configuration"
        $configGroupBox.Location = New-Object System.Drawing.Point(15, 110)
        $configGroupBox.Size = New-Object System.Drawing.Size(400, 140)

        # Investigator Name
        $investigatorNameLabel = New-Object System.Windows.Forms.Label
        $investigatorNameLabel.Text = "Investigator Name:"
        $investigatorNameLabel.Location = New-Object System.Drawing.Point(20, 30)
        $investigatorNameLabel.Size = New-Object System.Drawing.Size(120, 20)

        $investigatorNameTextBox = New-Object System.Windows.Forms.TextBox
        $investigatorNameTextBox.Text = "Security Administrator"
        $investigatorNameTextBox.Location = New-Object System.Drawing.Point(145, 27)
        $investigatorNameTextBox.Size = New-Object System.Drawing.Size(230, 20)

        # Company Name
        $companyNameLabel = New-Object System.Windows.Forms.Label
        $companyNameLabel.Text = "Company Name:"
        $companyNameLabel.Location = New-Object System.Drawing.Point(20, 60)
        $companyNameLabel.Size = New-Object System.Drawing.Size(120, 20)

        $companyNameTextBox = New-Object System.Windows.Forms.TextBox
        $companyNameTextBox.Text = "Organization"
        $companyNameTextBox.Location = New-Object System.Drawing.Point(145, 57)
        $companyNameTextBox.Size = New-Object System.Drawing.Size(230, 20)

        # Prefill from saved settings if available
        try {
            Import-Module "$PSScriptRoot\Modules\Settings.psm1" -Force -ErrorAction SilentlyContinue
            $s = Get-AppSettings
            if ($s) {
                if ($s.InvestigatorName -and $s.InvestigatorName.Trim().Length -gt 0) { $investigatorNameTextBox.Text = $s.InvestigatorName }
                if ($s.CompanyName -and $s.CompanyName.Trim().Length -gt 0) { $companyNameTextBox.Text = $s.CompanyName }
            }
        } catch {}

        # Days to Analyze
        $daysLabel = New-Object System.Windows.Forms.Label
        $daysLabel.Text = "Days to Analyze:"
        $daysLabel.Location = New-Object System.Drawing.Point(20, 90)
        $daysLabel.Size = New-Object System.Drawing.Size(120, 20)

        $daysComboBox = New-Object System.Windows.Forms.ComboBox
        $daysComboBox.Items.AddRange(@("1", "3", "7", "10", "30"))
        $daysComboBox.SelectedItem = "10"
        $daysComboBox.Location = New-Object System.Drawing.Point(145, 87)
        $daysComboBox.Size = New-Object System.Drawing.Size(80, 20)

        # Connection Status
        $connectionStatusLabel = New-Object System.Windows.Forms.Label
        $connectionStatusLabel.Text = "Checking connections..."
        $connectionStatusLabel.Font = New-Object System.Drawing.Font('Segoe UI', 8, [System.Drawing.FontStyle]::Italic)
        $connectionStatusLabel.Location = New-Object System.Drawing.Point(20, 115)
        $connectionStatusLabel.Size = New-Object System.Drawing.Size(350, 20)

        $configGroupBox.Controls.AddRange(@($investigatorNameLabel, $investigatorNameTextBox, $companyNameLabel, $companyNameTextBox, $daysLabel, $daysComboBox, $connectionStatusLabel))

        # Generate Button
        $generateButton = New-Object System.Windows.Forms.Button
        $generateButton.Text = "🚀 Generate Security Investigation"
        $generateButton.Font = New-Object System.Drawing.Font('Segoe UI', 12, [System.Drawing.FontStyle]::Bold)
        $generateButton.Location = New-Object System.Drawing.Point(430, 140)
        $generateButton.Size = New-Object System.Drawing.Size(280, 50)
        $generateButton.BackColor = [System.Drawing.Color]::FromArgb(0, 122, 204)
        $generateButton.ForeColor = [System.Drawing.Color]::White

        # Progress label
        $progressLabel = New-Object System.Windows.Forms.Label
        $progressLabel.Text = "Ready to generate security investigation report."
        $progressLabel.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Italic)
        $progressLabel.Location = New-Object System.Drawing.Point(430, 200)
        $progressLabel.Size = New-Object System.Drawing.Size(400, 20)
        $progressLabel.ForeColor = [System.Drawing.Color]::Green

        # Update connection status
        $exchangeConnected = $script:currentExchangeConnection -ne $null
        $graphConnected = $script:graphConnection -ne $null

        if ($exchangeConnected -and $graphConnected) {
            $connectionStatusLabel.Text = "✅ Both Exchange Online and Microsoft Graph connected"
            $connectionStatusLabel.ForeColor = [System.Drawing.Color]::Green
            $generateButton.Enabled = $true
        } elseif ($exchangeConnected -or $graphConnected) {
            $connectionStatusLabel.Text = "⚠️ Partial connection - some data may be unavailable"
            $connectionStatusLabel.ForeColor = [System.Drawing.Color]::Orange
            $generateButton.Enabled = $true
        } else {
            $connectionStatusLabel.Text = "❌ No connections available - cannot generate report"
            $connectionStatusLabel.ForeColor = [System.Drawing.Color]::Red
            $generateButton.Enabled = $false
        }

        # Generate button click handler
        $generateButton.add_Click({
            try { Import-Module "$PSScriptRoot\Modules\Settings.psm1" -Force -ErrorAction SilentlyContinue } catch {}
            $settings = $null; try { $settings = Get-AppSettings } catch {}
            $investigator = if ($investigatorNameTextBox.Text -and $investigatorNameTextBox.Text.Trim().Length -gt 0) { $investigatorNameTextBox.Text } elseif ($settings -and $settings.InvestigatorName) { $settings.InvestigatorName } else { 'Security Administrator' }
            $company = if ($companyNameTextBox.Text -and $companyNameTextBox.Text.Trim().Length -gt 0) { $companyNameTextBox.Text } elseif ($settings -and $settings.CompanyName) { $settings.CompanyName } else { 'Organization' }
            $days = [int]$daysComboBox.SelectedItem

            $progressLabel.Text = "🔍 Starting comprehensive security investigation..."
            $progressLabel.ForeColor = [System.Drawing.Color]::Blue
            $securityForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            $generateButton.Enabled = $false

            try {
                # Import the security investigation module
                Import-Module "$PSScriptRoot\Modules\ExportUtils.psm1" -Force -ErrorAction Stop

                # Resolve output folder for this run
                # Determine tenant-scoped folder root to match ExportUtils behavior
                $defaultRoot = Join-Path $env:USERPROFILE "Documents\\ExchangeOnlineAnalyzer\\SecurityInvestigation"
                $tenantName = $null
                try {
                    Import-Module "$PSScriptRoot\Modules\BrowserIntegration.psm1" -Force -ErrorAction SilentlyContinue
                    $ti = $null; try { $ti = Get-TenantIdentity } catch {}
                    if ($ti) { if ($ti.TenantDisplayName) { $tenantName = $ti.TenantDisplayName } elseif ($ti.PrimaryDomain) { $tenantName = $ti.PrimaryDomain } }
                    if (-not $tenantName) { try { $org = Get-OrganizationConfig -ErrorAction Stop; if ($org.DisplayName) { $tenantName = $org.DisplayName } elseif ($org.Name) { $tenantName = $org.Name } } catch {} }
                } catch {}
                if (-not $tenantName -or [string]::IsNullOrWhiteSpace($tenantName)) { $tenantName = 'Tenant' }
                $invalid = [System.IO.Path]::GetInvalidFileNameChars()
                $safeName = ($tenantName.ToCharArray() | ForEach-Object { if ($invalid -contains $_) { '-' } else { $_ } }) -join ''
                $safeName = ($safeName -replace '\s+', ' ').Trim()
                if ($safeName.Length -gt 80) { $safeName = $safeName.Substring(0,80) }
                $tenantRoot = Join-Path $defaultRoot $safeName
                if (-not (Test-Path $tenantRoot)) { New-Item -ItemType Directory -Path $tenantRoot -Force | Out-Null }
                $timestampFolder = Join-Path $tenantRoot (Get-Date -Format "yyyyMMdd_HHmmss")

                # Generate the security investigation report with export paths
                $securityReport = New-SecurityInvestigationReport -InvestigatorName $investigator -CompanyName $company -DaysBack $days -StatusLabel $progressLabel -MainForm $securityForm -OutputFolder $timestampFolder

                if ($securityReport) {
                    $progressLabel.Text = "✅ Security investigation completed successfully!"
                    $progressLabel.ForeColor = [System.Drawing.Color]::Green

                    # Show results in a new form
                    $resultsForm = New-Object System.Windows.Forms.Form
                    $resultsForm.Text = "Security Investigation Results"
                    $resultsForm.Size = New-Object System.Drawing.Size(1000, 800)
                    $resultsForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent

                    # Create tab control for results
                    $resultsTabControl = New-Object System.Windows.Forms.TabControl
                    $resultsTabControl.Dock = 'Fill'

                    # Summary tab
                    $summaryTab = New-Object System.Windows.Forms.TabPage
                    $summaryTab.Text = "📋 Investigation Summary"
                    $summaryTextBox = New-Object System.Windows.Forms.RichTextBox
                    $summaryTextBox.Dock = 'Fill'
                    $summaryTextBox.ReadOnly = $true
                    $summaryTextBox.Font = New-Object System.Drawing.Font('Consolas', 9)
                    $summaryTextBox.Text = $securityReport.Summary
                    $summaryTab.Controls.Add($summaryTextBox)

                    # AI Prompt tab
                    $aiPromptTab = New-Object System.Windows.Forms.TabPage
                    $aiPromptTab.Text = "🤖 AI Investigation Prompt"
                    $aiPromptTextBox = New-Object System.Windows.Forms.RichTextBox
                    $aiPromptTextBox.Dock = 'Fill'
                    $aiPromptTextBox.ReadOnly = $true
                    $aiPromptTextBox.Font = New-Object System.Drawing.Font('Consolas', 9)
                    $aiPromptTextBox.Text = $securityReport.AIPrompt
                    $aiPromptTab.Controls.Add($aiPromptTextBox)

                    # Ticket Message tab
                    $ticketTab = New-Object System.Windows.Forms.TabPage
                    $ticketTab.Text = "📝 Non-Technical Summary"
                    $ticketTextBox = New-Object System.Windows.Forms.RichTextBox
                    $ticketTextBox.Dock = 'Fill'
                    $ticketTextBox.ReadOnly = $true
                    $ticketTextBox.Font = New-Object System.Drawing.Font('Segoe UI', 10)
                    $ticketTextBox.Text = $securityReport.TicketMessage
                    $ticketTab.Controls.Add($ticketTextBox)

                    # Add tabs
                    $resultsTabControl.TabPages.Add($summaryTab)
                    $resultsTabControl.TabPages.Add($aiPromptTab)
                    $resultsTabControl.TabPages.Add($ticketTab)

                    # Copy buttons and instructions panel
                    $copyPanel = New-Object System.Windows.Forms.Panel
                    $copyPanel.Dock = 'Bottom'
                    $copyPanel.Height = 85
                    $copyPanel.BackColor = [System.Drawing.Color]::FromArgb(240, 242, 244)

                    # Simple instructions
                    $instructionsLabel = New-Object System.Windows.Forms.Label
                    $instructionsLabel.AutoSize = $true
                    $instructionsLabel.Location = New-Object System.Drawing.Point(15, 10)
                    $instructionsLabel.ForeColor = [System.Drawing.Color]::FromArgb(80,80,80)
                    $instructionsLabel.Text = "Instructions: Review exported CSVs in the folder (use 'Open Export Folder'). Upload MessageTrace.csv, InboxRules.csv, AuditLogs.csv, MFAStatus.csv, and UserSecurityGroups.csv to your analysis workspace or LLM. Reminder: Download Entra sign-in logs from the Entra portal (Sign-in logs → Download CSV) and include them alongside these files for full analysis."

                    $copySummaryBtn = New-Object System.Windows.Forms.Button
                    $copySummaryBtn.Text = "📋 Copy Summary"
                    $copySummaryBtn.Location = New-Object System.Drawing.Point(15, 35)
                    $copySummaryBtn.Size = New-Object System.Drawing.Size(130, 30)
                    $copySummaryBtn.add_Click({
                        [System.Windows.Forms.Clipboard]::SetText($summaryTextBox.Text)
                        [System.Windows.Forms.MessageBox]::Show("Summary copied to clipboard!", "Copied", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                    })

                    $copyAIBtn = New-Object System.Windows.Forms.Button
                    $copyAIBtn.Text = "🤖 Copy AI Prompt"
                    $copyAIBtn.Location = New-Object System.Drawing.Point(155, 35)
                    $copyAIBtn.Size = New-Object System.Drawing.Size(130, 30)
                    $copyAIBtn.add_Click({
                        [System.Windows.Forms.Clipboard]::SetText($aiPromptTextBox.Text)
                        [System.Windows.Forms.MessageBox]::Show("AI prompt copied to clipboard!", "Copied", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                    })

                    $copyTicketBtn = New-Object System.Windows.Forms.Button
                    $copyTicketBtn.Text = "📝 Copy Ticket Message"
                    $copyTicketBtn.Location = New-Object System.Drawing.Point(295, 35)
                    $copyTicketBtn.Size = New-Object System.Drawing.Size(150, 30)
                    $copyTicketBtn.add_Click({
                        [System.Windows.Forms.Clipboard]::SetText($ticketTextBox.Text)
                        [System.Windows.Forms.MessageBox]::Show("Ticket message copied to clipboard!", "Copied", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                    })

                    # Open folder button if files were exported
                    $openFolderBtn = New-Object System.Windows.Forms.Button
                    $openFolderBtn.Text = "Open Export Folder"
                    $openFolderBtn.Location = New-Object System.Drawing.Point(465, 35)
                    $openFolderBtn.Size = New-Object System.Drawing.Size(150, 30)
                    $openFolderBtn.add_Click({ if ($securityReport.OutputFolder) { Start-Process $securityReport.OutputFolder } })

                    # Provider selector + Send to AI
                    $providerLabel = New-Object System.Windows.Forms.Label
                    $providerLabel.Text = "Provider:"
                    $providerLabel.Location = New-Object System.Drawing.Point(625, 12)
                    $providerLabel.AutoSize = $true

                    $providerCombo = New-Object System.Windows.Forms.ComboBox
                    $providerCombo.Location = New-Object System.Drawing.Point(685, 10)
                    $providerCombo.Width = 100
                    $providerCombo.DropDownStyle = 'DropDownList'
                    $providerCombo.Items.AddRange(@('Gemini','Claude'))
                    $providerCombo.SelectedIndex = 0

                    $sendAIBtn = New-Object System.Windows.Forms.Button
                    $sendAIBtn.Text = "Send to AI"
                    $sendAIBtn.Location = New-Object System.Drawing.Point(625, 35)
                    $sendAIBtn.Size = New-Object System.Drawing.Size(160, 30)
                    $sendAIBtn.add_Click({
                        try {
                            # Determine default folder
                            $folder = $null
                            if ($securityReport.OutputFolder -and (Test-Path $securityReport.OutputFolder)) { $folder = $securityReport.OutputFolder }
                            if (-not $folder) {
                                $base = Join-Path $env:USERPROFILE "Documents\ExchangeOnlineAnalyzer\SecurityInvestigation"
                                $candidates = @()
                                try {
                                    $tenants = Get-ChildItem -Path $base -Directory -ErrorAction SilentlyContinue
                                    foreach ($t in $tenants) {
                                        $runs = Get-ChildItem -Path $t.FullName -Directory -ErrorAction SilentlyContinue
                                        if ($runs) { $candidates += $runs }
                                    }
                                    $legacy = Get-ChildItem -Path $base -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -match '^\d{8}_\d{6}$' }
                                    if ($legacy) { $candidates += $legacy }
                                    if ($candidates -and $candidates.Count -gt 0) { $folder = ($candidates | Sort-Object LastWriteTime -Descending | Select-Object -First 1).FullName }
                                } catch {}
                            }

                            # Ask to use latest or choose
                            if ($folder -and (Test-Path $folder)) {
                                $resp = [System.Windows.Forms.MessageBox]::Show("Use last generated folder?`n`n$folder`n`nYes = use this folder, No = choose another.", "Send to AI", [System.Windows.Forms.MessageBoxButtons]::YesNoCancel, [System.Windows.Forms.MessageBoxIcon]::Question)
                                if ($resp -eq [System.Windows.Forms.DialogResult]::No) {
                                    $fbd = New-Object System.Windows.Forms.FolderBrowserDialog
                                    $fbd.Description = "Select the report folder that contains LLM_Instructions.txt and CSV files"
                                    $fbd.SelectedPath = $folder
                                    if ($fbd.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
                                    $folder = $fbd.SelectedPath
                                } elseif ($resp -eq [System.Windows.Forms.DialogResult]::Cancel) { return }
                            } else {
                                $fbd = New-Object System.Windows.Forms.FolderBrowserDialog
                                $fbd.Description = "Select the report folder that contains LLM_Instructions.txt and CSV files"
                                if ($fbd.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
                                $folder = $fbd.SelectedPath
                            }

                            if (-not (Test-Path $folder)) {
                                [System.Windows.Forms.MessageBox]::Show("Selected folder not found:", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                                return
                            }

                            $provider = $providerCombo.SelectedItem
                            if ($provider -eq 'Claude') {
                                $scriptPath = Join-Path $PSScriptRoot "Scripts\Send-To-Claude.ps1"
                                if (-not (Test-Path $scriptPath)) {
                                    [System.Windows.Forms.MessageBox]::Show("Claude sender script not found at:`n$scriptPath", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                                    return
                                }
                            } else {
                                $scriptPath = Join-Path $PSScriptRoot "Scripts\Send-To-Gemini.ps1"
                                if (-not (Test-Path $scriptPath)) {
                                    [System.Windows.Forms.MessageBox]::Show("Gemini sender script not found at:`n$scriptPath", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                                    return
                                }
                            }

                            $resultsForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
                            $sendAIBtn.Enabled = $false
                            $outMsg = $null
                            try {
                                if ($provider -eq 'Claude') {
                                    $ps = { param($sp,$of) & $sp -OutputFolder $of -VerboseOutput 4>&1 }
                                    $output = & $ps $scriptPath $folder
                                } else {
                                    $ps = { param($sp,$of) & $sp -OutputFolder $of -Verbose 4>&1 }
                                    $output = & $ps $scriptPath $folder
                                }
                                $outMsg = ($output | Out-String)
                            } catch {
                                $outMsg = $_.Exception.Message
                            } finally {
                                $resultsForm.Cursor = [System.Windows.Forms.Cursors]::Default
                                $sendAIBtn.Enabled = $true
                            }

                            $respFile = if ($provider -eq 'Claude') { Join-Path $folder "Claude_Response.md" } else { Join-Path $folder "Gemini_Response.md" }
                            if (Test-Path $respFile) {
                                [System.Windows.Forms.MessageBox]::Show(("{0} response saved:`n{1}" -f $provider,$respFile), $provider, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                                try { Start-Process $respFile } catch {}
                            } else {
                                [System.Windows.Forms.MessageBox]::Show(("Send-To-{0} completed. Check the folder for the response file.`n`nOutput:`n{1}" -f $provider,$outMsg), $provider, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                            }
                        } catch {
                            [System.Windows.Forms.MessageBox]::Show(("Failed to send to {0}:`n{1}" -f $provider, $_.Exception.Message), "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                        }
                    })

                    $copyPanel.Controls.AddRange(@($instructionsLabel, $copySummaryBtn, $copyAIBtn, $copyTicketBtn, $openFolderBtn, $providerLabel, $providerCombo, $sendAIBtn))

                    $resultsForm.Controls.Add($resultsTabControl)
                    $resultsForm.Controls.Add($copyPanel)

                    $resultsForm.ShowDialog()

                } else {
                    $progressLabel.Text = "❌ Failed to generate security investigation report"
                    $progressLabel.ForeColor = [System.Drawing.Color]::Red
                    [System.Windows.Forms.MessageBox]::Show("Failed to generate security investigation report. Please check connections and permissions.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                }
            } catch {
                $progressLabel.Text = "❌ Error: $($_.Exception.Message)"
                $progressLabel.ForeColor = [System.Drawing.Color]::Red
                [System.Windows.Forms.MessageBox]::Show("Error generating security investigation report: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            } finally {
                $generateButton.Enabled = $true
                $securityForm.Cursor = [System.Windows.Forms.Cursors]::Default
            }
        })

        # Close button
        $closeButton = New-Object System.Windows.Forms.Button
        $closeButton.Text = "Close"
        $closeButton.Location = New-Object System.Drawing.Point(730, 140)
        $closeButton.Size = New-Object System.Drawing.Size(100, 50)
        $closeButton.add_Click({ $securityForm.Close() })

        # Add all controls to main panel
        $securityMainPanel.Controls.AddRange(@($securityTitleLabel, $securityDescLabel, $configGroupBox, $generateButton, $progressLabel, $closeButton))

        $securityForm.Controls.Add($securityMainPanel)

        $securityForm.ShowDialog()

        $mainForm.Cursor = [System.Windows.Forms.Cursors]::Default
        $statusLabel.Text = "Security investigation interface closed"

    } catch {
        $mainForm.Cursor = [System.Windows.Forms.Cursors]::Default
        $statusLabel.Text = "❌ Error opening security investigation: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Error opening security investigation interface: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})
$reportGeneratorPanel.Controls.Add($securityInvestigationButton)

# Add account selector group to panel
$reportGeneratorPanel.Controls.Add($accountSelectorGroup)

# Add Report Generator tab to tab control
$tabControl.TabPages.Add($reportGeneratorTab)

# Reposition AI Analysis tab to the right of Report Generator
try {
    if ($tabControl.TabPages.Contains($aiTab)) { $tabControl.TabPages.Remove($aiTab) }
    $tabControl.TabPages.Add($aiTab)
} catch {}

# Initialize unified account grid when Report Generator tab is first shown
$reportGeneratorTab.add_Enter({
    try {
        # Update connection status first
        Update-ConnectionStatus
        
        # Check if we have any connection/data and only show popup when neither is connected/loaded
        $exoConnected = $false; try { Get-OrganizationConfig -ErrorAction Stop | Out-Null; $exoConnected = $true } catch {}
        $mgConnected = $false; try { $ctx = Get-MgContext -ErrorAction Stop; if ($ctx -and $ctx.Account) { $mgConnected = $true } } catch {}

        $hasExchangeData = $script:allLoadedMailboxUPNs -and $script:allLoadedMailboxUPNs.Count -gt 0
        $hasEntraData = $entraUserGrid.Rows.Count -gt 0
        
        if (-not $exoConnected -and -not $mgConnected -and -not $hasExchangeData -and -not $hasEntraData) {
            $statusLabel.Text = "⚠️ No data available - please connect to Exchange Online and/or Entra ID first"
            [System.Windows.Forms.MessageBox]::Show(
                "No account data available for reports.`n`nPlease connect to Exchange Online and/or Entra ID first, then refresh the account list.",
                "No Data Available",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            return
        }
        
        # Auto-refresh if grid is empty or if we have new data
        if ($unifiedAccountGrid.Rows.Count -eq 0) {
            $statusLabel.Text = "🔄 Auto-refreshing account list..."
            Update-UnifiedAccountGrid
            $accountCount = $unifiedAccountGrid.Rows.Count
            $statusLabel.Text = "✅ Reports tab ready - $accountCount accounts loaded"
        } else {
            $statusLabel.Text = "📊 Reports tab ready - $($unifiedAccountGrid.Rows.Count) accounts available"
        }
        
    } catch {
        $statusLabel.Text = "❌ Error initializing reports tab: $($_.Exception.Message)"
    }
})

# Add panel to tab
$reportGeneratorTab.Controls.Add($reportGeneratorPanel)

# Account selector button event handlers
$refreshAccountsButton.add_Click({
    try {
        $statusLabel.Text = "🔄 Refreshing unified account list..."
        $mainForm.Refresh()
        
        # Update connection status first
        Update-ConnectionStatus
        
        # Clear the grid first
        $unifiedAccountGrid.Rows.Clear()
        
        # Update the unified account grid with fresh data
        Update-UnifiedAccountGrid
        
        # Show success message with count
        $accountCount = $unifiedAccountGrid.Rows.Count
        $statusLabel.Text = "✅ Account list refreshed - $accountCount accounts loaded"
        
        # Show a brief success message
        [System.Windows.Forms.MessageBox]::Show(
            "Account list refreshed successfully!`n`nTotal accounts: $accountCount`n`nThis includes accounts from both Exchange Online and Entra ID.",
            "Refresh Complete",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
        
    } catch {
        $statusLabel.Text = "❌ Error refreshing account list: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show(
            "Error refreshing account list: $($_.Exception.Message)`n`nPlease ensure you are connected to both Exchange Online and Entra ID.",
            "Refresh Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
})

$selectAllAccountsButton.add_Click({
    for ($i = 0; $i -lt $unifiedAccountGrid.Rows.Count; $i++) {
        $unifiedAccountGrid.Rows[$i].Cells["Select"].Value = $true
    }
})

$deselectAllAccountsButton.add_Click({
    for ($i = 0; $i -lt $unifiedAccountGrid.Rows.Count; $i++) {
        $unifiedAccountGrid.Rows[$i].Cells["Select"].Value = $false
    }
})

# Generate Report button event handler
$generateReportButton.add_Click({
    try {
        $statusLabel.Text = "Generating unified professional report..."
        $mainForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        
        # Get selected accounts
        $selectedAccounts = Get-SelectedUnifiedAccounts
        
        if ($selectedAccounts.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Please select at least one account for unified reporting.", "No Accounts Selected", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            $mainForm.Cursor = [System.Windows.Forms.Cursors]::Default
            return
        }
        
        # Generate both report formats with selected accounts
        $professionalReport = Generate-UnifiedProfessionalReport -selectedAccounts $selectedAccounts
        $obsidianNote = Generate-UnifiedObsidianNote -selectedAccounts $selectedAccounts
        
        # Create popup form
        $reportForm = New-Object System.Windows.Forms.Form
        $reportForm.Text = "Unified Professional Report Generator"
        $reportForm.Size = New-Object System.Drawing.Size(900, 700)
        $reportForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
        $reportForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
        $reportForm.MaximizeBox = $true
        
        # Create tab control for different formats
        $reportTabControl = New-Object System.Windows.Forms.TabControl
        $reportTabControl.Dock = 'Fill'
        
        # Professional Report Tab
        $professionalTab = New-Object System.Windows.Forms.TabPage
        $professionalTab.Text = "Professional Report"
        
        $professionalTextBox = New-Object System.Windows.Forms.RichTextBox
        $professionalTextBox.Dock = 'Fill'
        $professionalTextBox.ReadOnly = $true
        $professionalTextBox.Font = New-Object System.Drawing.Font('Consolas', 10)
        $professionalTextBox.Text = $professionalReport
        $professionalTab.Controls.Add($professionalTextBox)
        
        # Copy button for professional report
        $copyProfessionalButton = New-Object System.Windows.Forms.Button
        $copyProfessionalButton.Text = "Copy Professional Report"
        $copyProfessionalButton.Location = New-Object System.Drawing.Point(10, 10)
        $copyProfessionalButton.Size = New-Object System.Drawing.Size(200, 30)
        $copyProfessionalButton.add_Click({
            [System.Windows.Forms.Clipboard]::SetText($professionalReport)
            [System.Windows.Forms.MessageBox]::Show("Professional report copied to clipboard!", "Copied", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        })
        $professionalTab.Controls.Add($copyProfessionalButton)
        
        # Obsidian Note Tab
        $obsidianTab = New-Object System.Windows.Forms.TabPage
        $obsidianTab.Text = "Obsidian Note"
        
        $obsidianTextBox = New-Object System.Windows.Forms.RichTextBox
        $obsidianTextBox.Dock = 'Fill'
        $obsidianTextBox.ReadOnly = $true
        $obsidianTextBox.Font = New-Object System.Drawing.Font('Consolas', 10)
        $obsidianTextBox.Text = $obsidianNote
        $obsidianTab.Controls.Add($obsidianTextBox)
        
        # Copy button for Obsidian note
        $copyObsidianButton = New-Object System.Windows.Forms.Button
        $copyObsidianButton.Text = "Copy Obsidian Note"
        $copyObsidianButton.Location = New-Object System.Drawing.Point(10, 10)
        $copyObsidianButton.Size = New-Object System.Drawing.Size(200, 30)
        $copyObsidianButton.add_Click({
            [System.Windows.Forms.Clipboard]::SetText($obsidianNote)
            [System.Windows.Forms.MessageBox]::Show("Obsidian note copied to clipboard!", "Copied", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        })
        $obsidianTab.Controls.Add($copyObsidianButton)

        
        # Add tabs to tab control
        $reportTabControl.TabPages.Add($professionalTab)
        $reportTabControl.TabPages.Add($obsidianTab)
        
        # Add tab control to form
        $reportForm.Controls.Add($reportTabControl)
        
        # Show the form
        $reportForm.ShowDialog()
        
        $mainForm.Cursor = [System.Windows.Forms.Cursors]::Default
        $statusLabel.Text = "Unified professional report generated successfully"
        
    } catch {
        $mainForm.Cursor = [System.Windows.Forms.Cursors]::Default
        $statusLabel.Text = "Error generating unified professional report: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Error generating unified professional report: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

# --- Entra Portal Shortcuts (v8.1b) ---
$entraPortalGroup = New-Object System.Windows.Forms.GroupBox
$entraPortalGroup.Text = "Entra Portal Shortcuts (Preview)"
$entraPortalGroup.Location = New-Object System.Drawing.Point(10, 465)
$entraPortalGroup.Size = New-Object System.Drawing.Size(780, 140)
$entraPortalGroup.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$reportGeneratorPanel.Controls.Add($entraPortalGroup)
$entraPortalGroup.BringToFront()

$profileLabel = New-Object System.Windows.Forms.Label
$profileLabel.Text = "Firefox Profile:"
$profileLabel.Location = New-Object System.Drawing.Point(15, 25)
$profileCombo = New-Object System.Windows.Forms.ComboBox
$profileCombo.Location = New-Object System.Drawing.Point(115, 22)
$profileCombo.Width = 100

$containerLabel = New-Object System.Windows.Forms.Label
$containerLabel.Text = "Container:"
$containerLabel.Location = New-Object System.Drawing.Point(250, 25)
$containerCombo = New-Object System.Windows.Forms.ComboBox
$containerCombo.Location = New-Object System.Drawing.Point(320, 22)
$containerCombo.Width = 200

$openSignInsBtn = New-Object System.Windows.Forms.Button
$openSignInsBtn.Text = "Open Sign-in Logs"
$openSignInsBtn.Location = New-Object System.Drawing.Point(15, 50)
$openSignInsBtn.Size = New-Object System.Drawing.Size(130, 25)

$openRestrictedBtn = New-Object System.Windows.Forms.Button
$openRestrictedBtn.Text = "Restricted Entities"
$openRestrictedBtn.Location = New-Object System.Drawing.Point(290, 50)
$openRestrictedBtn.Size = New-Object System.Drawing.Size(130, 25)

$openCABtn = New-Object System.Windows.Forms.Button
$openCABtn.Text = "Conditional Access"
$openCABtn.Location = New-Object System.Drawing.Point(150, 50)
$openCABtn.Size = New-Object System.Drawing.Size(130, 25)

$entraPortalGroup.Controls.AddRange(@($profileLabel,$profileCombo,$containerLabel,$containerCombo,$openSignInsBtn,$openRestrictedBtn,$openCABtn))
$profileCombo.BringToFront()
$containerCombo.BringToFront()

# Helper note about required extension
$extNote = New-Object System.Windows.Forms.Label
$extNote.AutoSize = $true
$extNote.Location = New-Object System.Drawing.Point(15, 80)
$extNote.ForeColor = [System.Drawing.Color]::FromArgb(120,120,120)
$extNote.Text = "Requires Firefox add-on 'Open external links in a container'. If not installed, links open in a normal tab."
$entraPortalGroup.Controls.Add($extNote)

$loadFirefoxUi = {
    try {
        $ffStatusLabel.Text = "Loading Firefox profiles..."
        Import-Module "$PSScriptRoot\Modules\BrowserIntegration.psm1" -Force -ErrorAction Stop

        $profilesIniPath = Join-Path $env:APPDATA 'Mozilla\Firefox\profiles.ini'
        $basePath = Join-Path $env:APPDATA 'Mozilla\Firefox'

        $profiles = @()
        try { $profiles = Get-FirefoxProfiles } catch { $ffStatusLabel.Text = "Error: Get-FirefoxProfiles failed: $($_.Exception.Message)"; return }

        $profileCombo.Items.Clear()
        foreach ($p in $profiles) { if ($p -and $p.Name) { [void]$profileCombo.Items.Add($p.Name) } }

        # Prefer the most recently used/updated profile based on containers.json timestamp; fall back to 'Default' then first
        $latestProfile = $null
        $latestTime = [datetime]::MinValue
        foreach ($p in $profiles) {
            try {
                if (-not $p -or -not $p.Path) { continue }
                $pp = if ($p.Path -like '*:*') { $p.Path } else { Join-Path $basePath $p.Path }
                $cp = Join-Path $pp 'containers.json'
                $t = $null
                if (Test-Path $cp) { $t = (Get-Item $cp).LastWriteTime }
                elseif (Test-Path $pp) { $t = (Get-Item $pp).LastWriteTime }
                if ($t -and ($t -gt $latestTime)) { $latestTime = $t; $latestProfile = $p }
            } catch {}
        }
        $default = $null
        try { $default = ($profiles | Where-Object { $_.Default -eq $true } | Select-Object -First 1) } catch {}
        if ($latestProfile -and $latestProfile.Name) { $profileCombo.SelectedItem = $latestProfile.Name }
        elseif ($default -and $default.Name) { $profileCombo.SelectedItem = $default.Name }
        elseif ($profileCombo.Items.Count -gt 0) { $profileCombo.SelectedIndex = 0 }

        $containerCombo.Items.Clear()
        if ($profileCombo.SelectedItem) {
            $prof = ($profiles | Where-Object { $_.Name -eq $profileCombo.SelectedItem } | Select-Object -First 1)
            if ($prof -and $prof.Path) {
                $ppath = if ($prof.Path -like '*:*') { $prof.Path } else { Join-Path $basePath $prof.Path }
                if (Test-Path $ppath) {
                    try {
                        $containers = Get-FirefoxContainers -ProfilePath $ppath
                        # Filter out internal/synthetic names and sort alphabetically
                        $visible = $containers | Where-Object { $_ -and $_.name -and ($_.name.Trim().Length -gt 0) -and ($_.name -notmatch '^userContextIdInternal') } | Sort-Object name
                        $containerCombo.Items.Clear(); foreach ($c in $visible) { [void]$containerCombo.Items.Add($c.name) }
                        if ($containerCombo.Items.Count -gt 0) {
                            $tenant = Get-TenantIdentity
                            $bestName = $null; $bestScore = 0.0
                            try {
                                Import-Module "$PSScriptRoot\Modules\BrowserIntegration.psm1" -Force -ErrorAction SilentlyContinue
                                $names = ($visible | Select-Object -ExpandProperty name)
                                $best = Get-BestContainerName -ContainerNames $names -TenantIdentity $tenant
                                if ($best) { $bestName = $best.Name; $bestScore = $best.Score }
                            } catch {}

                            # Fallback heuristics if low score
                            if (-not $bestName -or -not ($containerCombo.Items -contains $bestName) -or $bestScore -lt 0.5) {
                                $norm = @{}
                                foreach ($n in $containerCombo.Items) {
                                    $key = ($n.ToString().ToLower() -replace '[^a-z0-9 ]',' ' -replace '\s+',' ').Trim()
                                    if (-not $norm.ContainsKey($key)) { $norm[$key] = $n }
                                }
                                $targets = @()
                                if ($tenant.TenantDisplayName) { $targets += $tenant.TenantDisplayName }
                                if ($tenant.PrimaryDomain) { $targets += $tenant.PrimaryDomain; $targets += ($tenant.PrimaryDomain -split '\.')[0] }
                                if ($tenant.Domains) { $targets += $tenant.Domains }
                                $picked = $null
                                foreach ($t in $targets) {
                                    if (-not $t) { continue }
                                    $tk = ($t.ToLower() -replace '[^a-z0-9 ]',' ' -replace '\s+',' ').Trim()
                                    # direct contains/prefix tests across normalized keys
                                    foreach ($k in $norm.Keys) {
                                        if ($k.StartsWith($tk) -or $tk.StartsWith($k) -or ($k.Contains($tk)) -or ($tk.Contains($k))) { $picked = $norm[$k]; break }
                                    }
                                    if ($picked) { break }
                                }
                                if ($picked) { $bestName = $picked }
                            }

                            if ($bestName -and ($containerCombo.Items -contains $bestName)) { $containerCombo.SelectedItem = $bestName } else { $containerCombo.SelectedIndex = 0 }
                            if ($ffStatusLabel) {
                                $bn = if ($bestName) { $bestName } else { '(none)' }
                                $ffStatusLabel.Text = ("Loaded {0} profile(s); {1} container(s) | Auto-match: {2} (score {3:N2})" -f ($profileCombo.Items.Count), ($containerCombo.Items.Count), $bn, $bestScore)
                            }
                        }
                        $ffStatusLabel.Text = ("Loaded {0} profile(s); {1} container(s)" -f ($profileCombo.Items.Count), ($containerCombo.Items.Count))
                    } catch {
                        $ffStatusLabel.Text = "Error loading containers: $($_.Exception.Message)"
                    }
                } else {
                    $ffStatusLabel.Text = "Profile path not found: $ppath"
                }
            } else {
                $ffStatusLabel.Text = "Selected profile has no path"
            }
        } else {
            if ($profiles.Count -eq 0) { $ffStatusLabel.Text = "No Firefox profiles found at: $profilesIniPath" } else { $ffStatusLabel.Text = "Select a Firefox profile" }
        }
    } catch {
        $ffStatusLabel.Text = "Refresh error: $($_.Exception.Message)"
    }
}

$refreshContainersBtn = New-Object System.Windows.Forms.Button
$refreshContainersBtn.Text = "Refresh"
$refreshContainersBtn.Location = New-Object System.Drawing.Point(525, 20)
$refreshContainersBtn.Size = New-Object System.Drawing.Size(75, 24)
$refreshContainersBtn.add_Click({ & $loadFirefoxUi })
$entraPortalGroup.Controls.Add($refreshContainersBtn)
$refreshContainersBtn.BringToFront()

$reloadDiskBtn = New-Object System.Windows.Forms.Button
$reloadDiskBtn.Text = "Reload Disk"
$reloadDiskBtn.Location = New-Object System.Drawing.Point(605, 20)
$reloadDiskBtn.Size = New-Object System.Drawing.Size(90, 24)
$reloadDiskBtn.add_Click({
    try {
        $ffStatusLabel.Text = "Reloading from disk..."
        Import-Module "$PSScriptRoot\Modules\BrowserIntegration.psm1" -Force -ErrorAction Stop
        $basePath = Join-Path $env:APPDATA 'Mozilla\Firefox'
        $profiles = Get-FirefoxProfiles
        $prof = ($profiles | Where-Object { $_.Name -eq $profileCombo.SelectedItem } | Select-Object -First 1)
        if (-not $prof) { $ffStatusLabel.Text = "Select a Firefox profile"; return }
        $ppath = if ($prof.Path -like '*:*') { $prof.Path } else { Join-Path $basePath $prof.Path }
        $cpath = Join-Path $ppath 'containers.json'
        if (-not (Test-Path $cpath)) { $ffStatusLabel.Text = "containers.json not found: $cpath"; return }
        $ts = (Get-Item $cpath).LastWriteTime
        $containers = Get-FirefoxContainers -ProfilePath $ppath
        $visible = $containers | Where-Object { $_ -and $_.name -and ($_.name.Trim().Length -gt 0) -and ($_.name -notmatch '^userContextIdInternal') } | Sort-Object name
        $containerCombo.Items.Clear(); foreach ($c in $visible) { [void]$containerCombo.Items.Add($c.name) }
        if ($containerCombo.Items.Count -gt 0) { $containerCombo.SelectedIndex = 0 }
        $ffStatusLabel.Text = ("Disk reload OK ({0}); {1} container(s)" -f $ts, $containerCombo.Items.Count)
    } catch { $ffStatusLabel.Text = "Reload error: $($_.Exception.Message)" }
})
$entraPortalGroup.Controls.Add($reloadDiskBtn)

# status label for diagnostics
$ffStatusLabel = New-Object System.Windows.Forms.Label
$ffStatusLabel.AutoSize = $true
$ffStatusLabel.Location = New-Object System.Drawing.Point(15, 100)
$ffStatusLabel.ForeColor = [System.Drawing.Color]::FromArgb(120,120,120)
$ffStatusLabel.Text = ""
$entraPortalGroup.Controls.Add($ffStatusLabel)

$entraPortalGroup.add_Enter({
    # Initial populate immediately
    & $loadFirefoxUi
    # Then attempt an eager auto-select using current Graph context
    try {
        Start-Sleep -Milliseconds 150
        Import-Module "$PSScriptRoot\Modules\BrowserIntegration.psm1" -Force -ErrorAction SilentlyContinue
        $profiles = Get-FirefoxProfiles
        $basePath = Join-Path $env:APPDATA 'Mozilla\Firefox'
        $prof = ($profiles | Where-Object { $_.Name -eq $profileCombo.SelectedItem } | Select-Object -First 1)
        if ($prof -and $prof.Path) {
            $ppath = if ($prof.Path -like '*:*') { $prof.Path } else { Join-Path $basePath $prof.Path }
            if (Test-Path $ppath) {
                $containers = Get-FirefoxContainers -ProfilePath $ppath
                $visible = $containers | Where-Object { $_ -and $_.name -and ($_.name.Trim().Length -gt 0) -and ($_.name -notmatch '^userContextIdInternal') } | Sort-Object name
                if ($visible.Count -gt 0) {
                    $tenant = $null
                    try { $tenant = Get-TenantIdentity } catch {}
                    if ($tenant) {
                        $best = Get-BestContainerName -ContainerNames ($visible | Select-Object -ExpandProperty name) -TenantIdentity $tenant
                        if ($best -and $best.Name -and ($containerCombo.Items -contains $best.Name)) { $containerCombo.SelectedItem = $best.Name }
                    }
                }
            }
        }
    } catch {}
})

# Initial populate when building the panel (in case Enter doesn't fire yet)
& $loadFirefoxUi

$profileCombo.add_SelectedIndexChanged({
    try {
        Import-Module "$PSScriptRoot\Modules\BrowserIntegration.psm1" -Force -ErrorAction SilentlyContinue
        $profiles = Get-FirefoxProfiles
        $prof = ($profiles | Where-Object { $_.Name -eq $profileCombo.SelectedItem } | Select-Object -First 1)
        if ($prof -and $prof.Path) {
            $ppath = if ($prof.Path -like '*:*') { $prof.Path } else { Join-Path (Join-Path $env:APPDATA 'Mozilla\Firefox') $prof.Path }
            $containers = Get-FirefoxContainers -ProfilePath $ppath
            $visible = $containers | Where-Object { $_ -and $_.name -and ($_.name.Trim().Length -gt 0) -and ($_.name -notmatch '^userContextIdInternal') } | Sort-Object name
            $containerCombo.Items.Clear(); foreach ($c in $visible) { [void]$containerCombo.Items.Add($c.name) }
            if ($containerCombo.Items.Count -gt 0) {
                $tenant = $null
                try { $tenant = Get-TenantIdentity } catch {}
                if ($tenant) {
                    $best = Get-BestContainerName -ContainerNames ($visible | Select-Object -ExpandProperty name) -TenantIdentity $tenant
                    if ($best -and $best.Name -and ($containerCombo.Items -contains $best.Name)) { $containerCombo.SelectedItem = $best.Name }
                    elseif ($containerCombo.Items.Count -gt 0) { $containerCombo.SelectedIndex = 0 }
                } else {
                    $containerCombo.SelectedIndex = 0
                }
            }
        }
    } catch {}
})

$openSignInsBtn.add_Click({ try { Import-Module "$PSScriptRoot\Modules\BrowserIntegration.psm1" -Force; Open-EntraDeepLink -ProfileName $profileCombo.SelectedItem -ContainerName $containerCombo.SelectedItem -Target 'SignIns' } catch {} })
$openRestrictedBtn.add_Click({ try { Import-Module "$PSScriptRoot\Modules\BrowserIntegration.psm1" -Force; Open-EntraDeepLink -ProfileName $profileCombo.SelectedItem -ContainerName $containerCombo.SelectedItem -Target 'RestrictedEntities' } catch {} })
$openCABtn.add_Click({ try { Import-Module "$PSScriptRoot\Modules\BrowserIntegration.psm1" -Force; Open-EntraDeepLink -ProfileName $profileCombo.SelectedItem -ContainerName $containerCombo.SelectedItem -Target 'ConditionalAccess' } catch {} })

# Incident Checklist button event handler
$incidentChecklistButton.add_Click({
    try {
        $statusLabel.Text = "Generating interactive incident remediation checklist..."
        $mainForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        
        # Get selected accounts
        $selectedAccounts = Get-SelectedUnifiedAccounts
        
        if ($selectedAccounts.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Please select at least one account for incident remediation analysis.", "No Accounts Selected", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            $mainForm.Cursor = [System.Windows.Forms.Cursors]::Default
            return
        }
        
        # Generate initial incident checklist
        $initialChecklist = Generate-IncidentRemediationChecklist -selectedAccounts $selectedAccounts
        
        # Create interactive popup form for incident checklist
        $checklistForm = New-Object System.Windows.Forms.Form
        $checklistForm.Text = "Interactive Incident Remediation Checklist"
        $checklistForm.Size = New-Object System.Drawing.Size(1000, 700)
        $checklistForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
        $checklistForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
        $checklistForm.MaximizeBox = $true
        
        # Create main panel
        $mainPanel = New-Object System.Windows.Forms.Panel
        $mainPanel.Dock = 'Fill'
        $checklistForm.Controls.Add($mainPanel)
        
        # Create header panel
        $headerPanel = New-Object System.Windows.Forms.Panel
        $headerPanel.Dock = 'Top'
        $headerPanel.Height = 50
        $mainPanel.Controls.Add($headerPanel)
        
        # Technician name input
        $technicianLabel = New-Object System.Windows.Forms.Label
        $technicianLabel.Text = "Technician Name:"
        $technicianLabel.Location = New-Object System.Drawing.Point(10, 15)
        $technicianLabel.Size = New-Object System.Drawing.Size(100, 20)
        $headerPanel.Controls.Add($technicianLabel)
        
        $technicianTextBox = New-Object System.Windows.Forms.TextBox
        $technicianTextBox.Location = New-Object System.Drawing.Point(120, 12)
        $technicianTextBox.Size = New-Object System.Drawing.Size(150, 25)
        $headerPanel.Controls.Add($technicianTextBox)
        
        # Create scrollable panel for checklist items
        $scrollPanel = New-Object System.Windows.Forms.Panel
        $scrollPanel.Dock = 'Fill'
        $scrollPanel.AutoScroll = $true
        $mainPanel.Controls.Add($scrollPanel)
        
        # Create checklist items with checkboxes
        $checklistItems = @(
            "Reset the Users Password in Active Directory or Office 365 if the account is a cloud-only account.",
            "Recommend Multi-Factor Authentication (MFA) to the client",
            "Apply the Require user to sign in again via Cloud App Security (if available)",
            "Force User Sign-out from Microsoft 365 Admin Panel",
            "Review the mailbox for any mailbox delegates and remove from the compromised account",
            "Review the mailbox for any mail forwarding rules that may have been created",
            "Review the mailbox inbox rules and delete any suspicious ones.",
            "Educate the user about security threats and methods used to gain access to users' credentials",
            "Run a mail trace to identify suspicious messages sent or received by this account",
            "Search the audit log to identify suspicious logins, attempt to identify the earliest date and time the account was compromised, and confirm no suspicious logins occur after password reset",
            "Advise the user that if the password that was in use is also used on any other accounts, those passwords should also be changed immediately",
            "Review the list of Administrators/Global Administrators in the Administration console. Check this against the users who SHOULD be Admins/Global Admins",
            "Review the Global/Domain Transport rules to ensure no rules have been set up.",
            "Review the list of licensed O365 Users. Check this against the list of users who SHOULD be in O365. Ensure that no disabled users or terminated users have a valid license assigned."
        )
        
                 $checkboxes = @()
         $yPosition = 30
        
        foreach ($item in $checklistItems) {
            # Create checkbox
            $checkbox = New-Object System.Windows.Forms.CheckBox
            $checkbox.Text = $item
            $checkbox.Location = New-Object System.Drawing.Point(10, $yPosition)
            $checkbox.Size = New-Object System.Drawing.Size(950, 20)
            $checkbox.AutoSize = $false
            $checkbox.Font = New-Object System.Drawing.Font('Segoe UI', 9)
            $scrollPanel.Controls.Add($checkbox)
            $checkboxes += $checkbox
            
            $yPosition += 30
        }
        
        # Create button panel
        $buttonPanel = New-Object System.Windows.Forms.Panel
        $buttonPanel.Dock = 'Bottom'
        $buttonPanel.Height = 50
        $mainPanel.Controls.Add($buttonPanel)
        
        # Mark all as completed button
        $markAllButton = New-Object System.Windows.Forms.Button
        $markAllButton.Text = "Mark All as Completed"
        $markAllButton.Location = New-Object System.Drawing.Point(10, 10)
        $markAllButton.Size = New-Object System.Drawing.Size(150, 30)
        $markAllButton.add_Click({
            $technicianName = $technicianTextBox.Text
            if ([string]::IsNullOrWhiteSpace($technicianName)) {
                [System.Windows.Forms.MessageBox]::Show("Please enter a technician name first.", "Technician Name Required", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                return
            }
            
            $currentDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            foreach ($checkbox in $checkboxes) {
                if (-not $checkbox.Checked) {
                    $checkbox.Checked = $true
                    $checkbox.Text += " [Completed: $currentDate by $technicianName]"
                }
            }
        })
        $buttonPanel.Controls.Add($markAllButton)
        
        # Generate completed checklist button
        $generateCompletedButton = New-Object System.Windows.Forms.Button
        $generateCompletedButton.Text = "Generate Completed Checklist"
        $generateCompletedButton.Location = New-Object System.Drawing.Point(170, 10)
        $generateCompletedButton.Size = New-Object System.Drawing.Size(180, 30)
        $generateCompletedButton.add_Click({
            $technicianName = $technicianTextBox.Text
            if ([string]::IsNullOrWhiteSpace($technicianName)) {
                [System.Windows.Forms.MessageBox]::Show("Please enter a technician name first.", "Technician Name Required", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                return
            }
            
                         $completedChecklist = "The Essential Office 365 Account Incident Remediation Checklist`n"
             $completedChecklist += "Technician: $technicianName`n"
            $completedChecklist += "User Account: $($selectedAccounts[0].DisplayName)`n"
            $completedChecklist += "User Principal Name: $($selectedAccounts[0].UserPrincipalName)`n`n"
            
            $completedChecklist += "COMPLETED ITEMS:`n"
            $completedChecklist += "================`n`n"
            
            $completedItems = 0
                         foreach ($checkbox in $checkboxes) {
                 if ($checkbox.Checked) {
                     $completedItems++
                     $completedChecklist += "☑ $($checkbox.Text)`n`n"
                 }
             }
            
            
            
            # Create popup for completed checklist
            $completedForm = New-Object System.Windows.Forms.Form
            $completedForm.Text = "Completed Incident Checklist"
            $completedForm.Size = New-Object System.Drawing.Size(900, 600)
            $completedForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
            
            $completedTextBox = New-Object System.Windows.Forms.RichTextBox
            $completedTextBox.Dock = 'Fill'
            $completedTextBox.ReadOnly = $true
            $completedTextBox.Font = New-Object System.Drawing.Font('Consolas', 10)
            $completedTextBox.Text = $completedChecklist
            $completedForm.Controls.Add($completedTextBox)
            
            # Copy button for completed checklist
            $copyCompletedButton = New-Object System.Windows.Forms.Button
            $copyCompletedButton.Text = "Copy Completed Checklist"
            $copyCompletedButton.Location = New-Object System.Drawing.Point(10, 10)
            $copyCompletedButton.Size = New-Object System.Drawing.Size(200, 30)
            $copyCompletedButton.add_Click({
                [System.Windows.Forms.Clipboard]::SetText($completedChecklist)
                [System.Windows.Forms.MessageBox]::Show("Completed checklist copied to clipboard!", "Copied", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            })
            $completedForm.Controls.Add($copyCompletedButton)
            
            $completedForm.ShowDialog()
        })
        $buttonPanel.Controls.Add($generateCompletedButton)
        
        # Show the form
        $checklistForm.ShowDialog()
        
        $mainForm.Cursor = [System.Windows.Forms.Cursors]::Default
        $statusLabel.Text = "Interactive incident remediation checklist generated successfully"
        
    } catch {
        $mainForm.Cursor = [System.Windows.Forms.Cursors]::Default
        $statusLabel.Text = "Error generating interactive incident remediation checklist: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Error generating interactive incident remediation checklist: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})
        




# Add Help tab after other tabs
$helpTab = New-Object System.Windows.Forms.TabPage
$helpTab.Text = "Help"

# Create a RichTextBox for better formatting
$helpRichTextBox = New-Object System.Windows.Forms.RichTextBox
$helpRichTextBox.ReadOnly = $true
$helpRichTextBox.ScrollBars = 'Both'
$helpRichTextBox.Dock = 'Fill'
$helpRichTextBox.Font = New-Object System.Drawing.Font('Segoe UI', 10)
$helpRichTextBox.BackColor = [System.Drawing.Color]::White
$helpRichTextBox.ForeColor = [System.Drawing.Color]::Black
$helpRichTextBox.WordWrap = $true

# Create clean, formatted help content
$helpText = @"

MICROSOFT 365 MANAGEMENT TOOL - HELP

OVERVIEW
This tool provides comprehensive management capabilities for Microsoft 365 environments, including Exchange Online and Entra ID (Azure AD) administration.

EXCHANGE ONLINE TAB
• Connect to Exchange Online using modern authentication
• View and manage user mailboxes with detailed information
• Export inbox rules for analysis and backup
• Manage connectors (inbound/outbound) with delete capability
• Manage transport rules with delete capability
• Search and filter mailbox data
• Export data to CSV/Excel formats

ENTRA ID INVESTIGATOR TAB
• Connect to Microsoft Graph API
• View and manage user accounts
• Block/unblock user sign-in access
• Revoke user sessions for security
• Export sign-in and audit logs
• Analyze MFA status and user details
• View user roles and permissions

KEYBOARD SHORTCUTS
• Ctrl+O: Connect to services
• Ctrl+D: Disconnect from services
• Ctrl+S: Export rules/data
• F5: Refresh data
• Ctrl+A: Select all items
• Escape: Close dialogs

CONNECTION REQUIREMENTS
• Exchange Online PowerShell module
• Microsoft Graph PowerShell module
• Appropriate admin permissions
• Modern authentication enabled

TROUBLESHOOTING
• Ensure you have the required PowerShell modules installed
• Verify you have appropriate admin permissions
• Check your internet connection
• Ensure modern authentication is enabled for your tenant

For detailed documentation, please refer to the readme.md file in the application directory.

"@

$helpRichTextBox.Text = $helpText

$helpTab.Controls.Add($helpRichTextBox)
$tabControl.TabPages.Add($helpTab)

# Set Entra user grid column read-only properties
$entraUserGrid.ReadOnly = $false
$colEntraCheck.ReadOnly = $false
$colEntraUPN.ReadOnly = $true
$colEntraDisplayName.ReadOnly = $true
$colEntraLicensed.ReadOnly = $true

# --- Entra ID User Management Button Event Handlers ---
$entraBlockUserButton.add_Click({
    $entraUserGrid.EndEdit()
    $selectedUpns = @()
    for ($i = 0; $i -lt $entraUserGrid.Rows.Count; $i++) {
        if ($entraUserGrid.Rows[$i].Cells["Select"].Value -eq $true) {
            $upn = $entraUserGrid.Rows[$i].Cells["UserPrincipalName"].Value
            if (-not [string]::IsNullOrWhiteSpace($upn)) { $selectedUpns += $upn }
        }
    }
    if ($selectedUpns.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Select at least one user to block sign-in, or the operation will be performed on all loaded users.", "No Users Selected", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        # If no users selected, use all loaded users
        for ($i = 0; $i -lt $entraUserGrid.Rows.Count; $i++) {
            $upn = $entraUserGrid.Rows[$i].Cells["UserPrincipalName"].Value
            if (-not [string]::IsNullOrWhiteSpace($upn)) { $selectedUpns += $upn }
        }
        if ($selectedUpns.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("No users available to block.", "No Users Available", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
    }
    $confirm = [System.Windows.Forms.MessageBox]::Show("Block sign-in for the following user(s)?\n" + ($selectedUpns -join "\n"), "Confirm Block", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }
    try {
        Set-UserSignInBlockedState -UserPrincipalNames $selectedUpns -Blocked $true -StatusLabel $statusLabel -MainForm $mainForm
        [System.Windows.Forms.MessageBox]::Show("Blocked sign-in for selected user(s).", "Block User", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to block user(s): $($_.Exception.Message)", "Block User Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

$entraUnblockUserButton.add_Click({
    $entraUserGrid.EndEdit()
    $selectedUpns = @()
    for ($i = 0; $i -lt $entraUserGrid.Rows.Count; $i++) {
        if ($entraUserGrid.Rows[$i].Cells["Select"].Value -eq $true) {
            $upn = $entraUserGrid.Rows[$i].Cells["UserPrincipalName"].Value
            if (-not [string]::IsNullOrWhiteSpace($upn)) { $selectedUpns += $upn }
        }
    }
    if ($selectedUpns.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Select at least one user to unblock sign-in, or the operation will be performed on all loaded users.", "No Users Selected", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        # If no users selected, use all loaded users
        for ($i = 0; $i -lt $entraUserGrid.Rows.Count; $i++) {
            $upn = $entraUserGrid.Rows[$i].Cells["UserPrincipalName"].Value
            if (-not [string]::IsNullOrWhiteSpace($upn)) { $selectedUpns += $upn }
        }
        if ($selectedUpns.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("No users available to unblock.", "No Users Available", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
    }
    $confirm = [System.Windows.Forms.MessageBox]::Show("Unblock sign-in for the following user(s)?\n" + ($selectedUpns -join "\n"), "Confirm Unblock", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }
    try {
        Set-UserSignInBlockedState -UserPrincipalNames $selectedUpns -Blocked $false -StatusLabel $statusLabel -MainForm $mainForm
        [System.Windows.Forms.MessageBox]::Show("Unblocked sign-in for selected user(s).", "Unblock User", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to unblock user(s): $($_.Exception.Message)", "Unblock User Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

$entraRevokeSessionsButton.add_Click({
    $entraUserGrid.EndEdit()
    $selectedUpns = @()
    for ($i = 0; $i -lt $entraUserGrid.Rows.Count; $i++) {
        if ($entraUserGrid.Rows[$i].Cells["Select"].Value -eq $true) {
            $upn = $entraUserGrid.Rows[$i].Cells["UserPrincipalName"].Value
            if (-not [string]::IsNullOrWhiteSpace($upn)) { $selectedUpns += $upn }
        }
    }
    if ($selectedUpns.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Select at least one user to revoke sessions, or the operation will be performed on all loaded users.", "No Users Selected", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        # If no users selected, use all loaded users
        for ($i = 0; $i -lt $entraUserGrid.Rows.Count; $i++) {
            $upn = $entraUserGrid.Rows[$i].Cells["UserPrincipalName"].Value
            if (-not [string]::IsNullOrWhiteSpace($upn)) { $selectedUpns += $upn }
        }
        if ($selectedUpns.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("No users available to revoke sessions.", "No Users Available", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
    }
    Show-SessionRevocationTool -mainForm $mainForm -statusLabel $statusLabel -allLoadedMailboxUPNs $selectedUpns
})

$entraResetPasswordButton.add_Click({
    $entraUserGrid.EndEdit()
    $selectedUpns = @()
    for ($i = 0; $i -lt $entraUserGrid.Rows.Count; $i++) {
        if ($entraUserGrid.Rows[$i].Cells["Select"].Value -eq $true) {
            $upn = $entraUserGrid.Rows[$i].Cells["UserPrincipalName"].Value
            if (-not [string]::IsNullOrWhiteSpace($upn)) { $selectedUpns += $upn }
        }
    }
    if ($selectedUpns.Count -ne 1) {
        [System.Windows.Forms.MessageBox]::Show("Select exactly one user to reset password.", "Select One User", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }
    
    $userUpn = $selectedUpns[0]
    
    # Generate memorable password with validation
    try {
        $newPassword = New-XKCDPassword -WordCount 4 -IncludeSeparator
        
        # Validate password was generated
        if ([string]::IsNullOrWhiteSpace($newPassword)) {
            throw "Password generation failed - generated password is null or empty"
        }
        
        # Additional validation - ensure password meets minimum requirements
        if ($newPassword.Length -lt 8) {
            throw "Generated password is too short (length: $($newPassword.Length))"
        }
        
        Write-Host "Generated password length: $($newPassword.Length)" -ForegroundColor Green
        
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to generate password: $($_.Exception.Message)`n`nTrying fallback password generation...", "Password Generation Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        
        # Fallback password generation
        try {
            $newPassword = "TempPass" + (Get-Random -Minimum 1000 -Maximum 9999) + "!"
            Write-Host "Using fallback password: $newPassword" -ForegroundColor Yellow
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Failed to generate fallback password: $($_.Exception.Message)", "Password Generation Failed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            $statusLabel.Text = "Password generation failed"
            return
        }
    }
    
    try {
        # Reset user password via Microsoft Graph
        $statusLabel.Text = "Resetting password for $userUpn..."
        $mainForm.Refresh()
        
        # Check if connected to Microsoft Graph
        $context = Get-MgContext -ErrorAction Stop
        if (-not $context) {
            throw "Not connected to Microsoft Graph. Please connect first."
        }
        
        # Validate user exists before attempting password reset
        try {
            $user = Get-MgUser -UserId $userUpn -ErrorAction Stop
            if (-not $user) {
                throw "User not found: $userUpn"
            }
        } catch {
            throw "Failed to validate user $userUpn : $($_.Exception.Message)"
        }
        
        # Reset the password
        $passwordProfile = @{
            Password = $newPassword
            ForceChangePasswordNextSignIn = $true
        }
        
        Update-MgUser -UserId $userUpn -PasswordProfile $passwordProfile -ErrorAction Stop
        
        # Show success dialog with password
        $message = "Password reset successful for user: $userUpn`n`nNew Password: $newPassword`n`nThis password is memorable and secure. The user will be required to change it on next sign-in.`n`nCopy password to clipboard?"
        
        $result = [System.Windows.Forms.MessageBox]::Show(
            $message,
            "Password Reset Successful",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
        
        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            [System.Windows.Forms.Clipboard]::SetText($newPassword)
            [System.Windows.Forms.MessageBox]::Show("Password copied to clipboard!", "Copied", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
        
        $statusLabel.Text = "Password reset completed for $userUpn"
        
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to reset password for $userUpn : $($_.Exception.Message)", "Password Reset Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        $statusLabel.Text = "Password reset failed for $userUpn"
    }
})

$entraOpenDefenderRestrictedUsersButton.add_Click({
    Start-Process "https://security.microsoft.com/restrictedentities"
})

# Add click handlers for Select All/Deselect All buttons
$entraSelectAllButton.add_Click({
    for ($i = 0; $i -lt $entraUserGrid.Rows.Count; $i++) {
        $entraUserGrid.Rows[$i].Cells["Select"].Value = $true
    }
    UpdateEntraButtonStates
})

$entraDeselectAllButton.add_Click({
    for ($i = 0; $i -lt $entraUserGrid.Rows.Count; $i++) {
        $entraUserGrid.Rows[$i].Cells["Select"].Value = $false
    }
    UpdateEntraButtonStates
})

# Add click handler for refresh roles
$entraRefreshRolesButton.add_Click({
    $entraUserGrid.EndEdit()
    $selectedUpns = @()
    for ($i = 0; $i -lt $entraUserGrid.Rows.Count; $i++) {
        if ($entraUserGrid.Rows[$i].Cells["Select"].Value -eq $true) {
            $upn = $entraUserGrid.Rows[$i].Cells["UserPrincipalName"].Value
            if (-not [string]::IsNullOrWhiteSpace($upn)) { $selectedUpns += $upn }
        }
    }
    if ($selectedUpns.Count -eq 0) {
        $result = [System.Windows.Forms.MessageBox]::Show("No users selected. Would you like to refresh roles for ALL users?", "Refresh All Roles", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            # Get all users from the grid
            for ($i = 0; $i -lt $entraUserGrid.Rows.Count; $i++) {
                $upn = $entraUserGrid.Rows[$i].Cells["UserPrincipalName"].Value
                if (-not [string]::IsNullOrWhiteSpace($upn)) { $selectedUpns += $upn }
            }
        } else {
            return
        }
    }
    
    # Check if connected to Microsoft Graph
    $context = Get-MgContext -ErrorAction SilentlyContinue
    if (-not $context) {
        [System.Windows.Forms.MessageBox]::Show("Please connect to Microsoft Graph first.", "Not Connected", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        $statusLabel.Text = "Not connected to Microsoft Graph"
        return
    }
    
    $statusLabel.Text = "Fetching directory roles from server..."
    $mainForm.Refresh()
    
    # Get all directory roles once (server-side)
    $directoryRoles = Get-MgDirectoryRole -ErrorAction Stop
    
    $statusLabel.Text = "Refreshing roles for selected users..."
    $mainForm.Refresh()
    
    $processedCount = 0
    foreach ($userUpn in $selectedUpns) {
        $processedCount++
        $statusLabel.Text = "Refreshing roles for user $processedCount of $($selectedUpns.Count): $userUpn"
        $mainForm.Refresh()
        
        try {
            # Get the user's object ID (GUID) for matching
            $userObj = Get-MgUser -UserId $userUpn -Property Id -ErrorAction SilentlyContinue
            if (-not $userObj) {
                continue
            }
            $userId = $userObj.Id
            $userRoles = @()
            
            # Check each directory role for this user (server-side)
            foreach ($role in $directoryRoles) {
                try {
                    $roleMembers = Get-MgDirectoryRoleMember -DirectoryRoleId $role.Id -ErrorAction SilentlyContinue
                    if ($roleMembers) {
                        foreach ($member in $roleMembers) {
                            if ($member.Id -eq $userId) {
                                $userRoles += $role.DisplayName
                                break
                            }
                        }
                    }
                } catch {
                    # Silently continue if role member lookup fails
                }
            }
            $rolesText = if ($userRoles.Count -gt 0) { ($userRoles -join ", ") } else { "No Roles" }
            
            # Update the roles column for this user
            for ($i = 0; $i -lt $entraUserGrid.Rows.Count; $i++) {
                if ($entraUserGrid.Rows[$i].Cells["UserPrincipalName"].Value -eq $userUpn) {
                    $entraUserGrid.Rows[$i].Cells["Roles"].Value = $rolesText
                    
                    # Update highlighting
                    $highPrivilegeRoles = @("Global Administrator", "Company Administrator", "Exchange Administrator", "SharePoint Administrator", "Security Administrator", "Compliance Administrator", "User Administrator", "Billing Administrator", "Helpdesk Administrator", "Service Support Administrator", "Power Platform Administrator", "Teams Administrator", "Intune Administrator", "Application Administrator", "Cloud Application Administrator", "Privileged Role Administrator", "Privileged Authentication Administrator")
                    
                    $hasHighPrivilege = $false
                    foreach ($role in $userRoles) {
                        if ($highPrivilegeRoles -contains $role) {
                            $hasHighPrivilege = $true
                            break
                        }
                    }
                    
                    if ($hasHighPrivilege) {
                        $entraUserGrid.Rows[$i].DefaultCellStyle.BackColor = [System.Drawing.Color]::LightCoral
                        $entraUserGrid.Rows[$i].DefaultCellStyle.ForeColor = [System.Drawing.Color]::DarkRed
                    } else {
                        $entraUserGrid.Rows[$i].DefaultCellStyle.BackColor = [System.Drawing.Color]::White
                        $entraUserGrid.Rows[$i].DefaultCellStyle.ForeColor = [System.Drawing.Color]::Black
                    }
                    break
                }
            }
        } catch {
            # Silently continue if role refresh fails for this user
        }
    }
    
    $statusLabel.Text = "Roles refreshed for selected users"
    [System.Windows.Forms.MessageBox]::Show("Roles refreshed for selected users.", "Refresh Roles", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
})

# Add click handler for require password change
$entraRequirePwdChangeButton.add_Click({
    $entraUserGrid.EndEdit()
    $selectedUpns = @()
    for ($i = 0; $i -lt $entraUserGrid.Rows.Count; $i++) {
        if ($entraUserGrid.Rows[$i].Cells["Select"].Value -eq $true) {
            $upn = $entraUserGrid.Rows[$i].Cells["UserPrincipalName"].Value
            if (-not [string]::IsNullOrWhiteSpace($upn)) { $selectedUpns += $upn }
        }
    }
    if ($selectedUpns.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Select at least one user to require password change, or the operation will be performed on all loaded users.", "No Users Selected", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        # If no users selected, use all loaded users
        for ($i = 0; $i -lt $entraUserGrid.Rows.Count; $i++) {
            $upn = $entraUserGrid.Rows[$i].Cells["UserPrincipalName"].Value
            if (-not [string]::IsNullOrWhiteSpace($upn)) { $selectedUpns += $upn }
        }
        if ($selectedUpns.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("No users available to require password change.", "No Users Available", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
    }
    $confirm = [System.Windows.Forms.MessageBox]::Show("Require password change at next sign-in for the following user(s)?\n" + ($selectedUpns -join "\n"), "Confirm Require Password Change", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }
    try {
        foreach ($userUpn in $selectedUpns) {
            $statusLabel.Text = "Requiring password change for $userUpn..."
            $mainForm.Refresh()
            $context = Get-MgContext -ErrorAction Stop
            if (-not $context) { throw "Not connected to Microsoft Graph. Please connect first." }
            $passwordProfile = @{ ForceChangePasswordNextSignIn = $true }
            Update-MgUser -UserId $userUpn -PasswordProfile $passwordProfile -ErrorAction Stop
        }
        [System.Windows.Forms.MessageBox]::Show("Password change required at next sign-in for selected user(s).", "Require Password Change", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        $statusLabel.Text = "Password change required for selected user(s)"
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to require password change: $($_.Exception.Message)", "Require Password Change Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        $statusLabel.Text = "Require password change failed"
    }
})

# Add click handler for view admins
$entraViewAdminsButton.add_Click({
    try {
        $statusLabel.Text = "Querying server for admin users..."
        $mainForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        $mainForm.Refresh()
        
        # Check if connected to Microsoft Graph
        $context = Get-MgContext -ErrorAction SilentlyContinue
        if (-not $context) {
            [System.Windows.Forms.MessageBox]::Show("Please connect to Microsoft Graph first.", "Not Connected", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            $mainForm.Cursor = [System.Windows.Forms.Cursors]::Default
            $statusLabel.Text = "Not connected to Microsoft Graph"
            return
        }
        
        # Get all users with elevated roles using server-side filtering
        $adminUsers = @()
        $highPrivilegeRoles = @("Global Administrator", "Company Administrator", "Exchange Administrator", "SharePoint Administrator", "Security Administrator", "Compliance Administrator", "User Administrator", "Billing Administrator", "Helpdesk Administrator", "Service Support Administrator", "Power Platform Administrator", "Teams Administrator", "Intune Administrator", "Application Administrator", "Cloud Application Administrator", "Privileged Role Administrator", "Privileged Authentication Administrator")
        
        # Get all directory roles first
        $statusLabel.Text = "Fetching directory roles..."
        $mainForm.Refresh()
        $directoryRoles = Get-MgDirectoryRole -ErrorAction Stop
        
        # Get all users with their roles
        $statusLabel.Text = "Querying users with elevated roles..."
        $mainForm.Refresh()
        $allUsers = Get-MgUser -All -Property Id, UserPrincipalName, DisplayName, AssignedLicenses -ErrorAction Stop
        
        $processedCount = 0
        foreach ($user in $allUsers) {
            $processedCount++
            if ($processedCount % 50 -eq 0) {
                $statusLabel.Text = "Processing user $processedCount of $($allUsers.Count)..."
                $mainForm.Refresh()
            }
            
            $userRoles = @()
            $hasElevatedRole = $false
            $elevatedRoles = @()
            
            # Check each directory role for this user
            foreach ($role in $directoryRoles) {
                try {
                    $roleMembers = Get-MgDirectoryRoleMember -DirectoryRoleId $role.Id -ErrorAction SilentlyContinue
                    if ($roleMembers) {
                        foreach ($member in $roleMembers) {
                            if ($member.Id -eq $user.Id) {
                                $userRoles += $role.DisplayName
                                if ($highPrivilegeRoles -contains $role.DisplayName) {
                                    $hasElevatedRole = $true
                                    $elevatedRoles += $role.DisplayName
                                }
                                break
                            }
                        }
                    }
                } catch {
                    # Silently continue if role member lookup fails
                }
            }
            
            if ($hasElevatedRole) {
                $licensed = if ($user.AssignedLicenses -and $user.AssignedLicenses.Count -gt 0) { "Licensed" } else { "Unlicensed" }
                $adminUsers += [PSCustomObject]@{
                    UserPrincipalName = $user.UserPrincipalName
                    DisplayName = $user.DisplayName
                    Licensed = $licensed
                    ElevatedRoles = ($elevatedRoles -join ", ")
                    AllRoles = ($userRoles -join ", ")
                }
            }
        }
        
        if ($adminUsers.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("No users with elevated roles found. Make sure to refresh roles for users first.", "No Admins Found", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            $mainForm.Cursor = [System.Windows.Forms.Cursors]::Default
            $statusLabel.Text = "No users with elevated roles found"
            return
        }
        
        # Create admin report
        $report = @"
# Microsoft 365 Admin Users Report
**Generated:** $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
**Total Admin Users Found:** $($adminUsers.Count)

## Admin Users Summary

"@
        
        foreach ($admin in $adminUsers) {
            $report += @"

### User: $($admin.DisplayName)
- **UPN:** $($admin.UserPrincipalName)
- **Licensed:** $($admin.Licensed)
- **Elevated Roles:** $($admin.ElevatedRoles)
- **All Roles:** $($admin.AllRoles)

"@
        }
        
        $report += @"

## Security Recommendations
1. **Review Admin Access:** Verify all listed users should have elevated privileges
2. **Implement Just-In-Time Access:** Consider implementing privileged access management
3. **Enable MFA:** Ensure all admin accounts have Multi-Factor Authentication enabled
4. **Regular Audits:** Schedule regular reviews of admin access
5. **Monitor Sign-ins:** Enable sign-in monitoring for all admin accounts
"@
        
        # Create popup form to display the report
        $reportForm = New-Object System.Windows.Forms.Form
        $reportForm.Text = "Admin Users Report - $($adminUsers.Count) Users Found"
        $reportForm.Size = New-Object System.Drawing.Size(800, 600)
        $reportForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
        $reportForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
        $reportForm.MaximizeBox = $true
        
        # Create rich text box for the report
        $reportTextBox = New-Object System.Windows.Forms.RichTextBox
        $reportTextBox.Dock = 'Fill'
        $reportTextBox.Font = New-Object System.Drawing.Font('Consolas', 9)
        $reportTextBox.Text = $report
        $reportTextBox.ReadOnly = $true
        $reportForm.Controls.Add($reportTextBox)
        
        # Create button panel
        $buttonPanel = New-Object System.Windows.Forms.Panel
        $buttonPanel.Dock = 'Bottom'
        $buttonPanel.Height = 50
        $reportForm.Controls.Add($buttonPanel)
        
        # Add copy button
        $copyButton = New-Object System.Windows.Forms.Button
        $copyButton.Text = "Copy to Clipboard"
        $copyButton.Location = New-Object System.Drawing.Point(10, 10)
        $copyButton.Size = New-Object System.Drawing.Size(120, 30)
        $copyButton.add_Click({
            [System.Windows.Forms.Clipboard]::SetText($reportTextBox.Text)
            [System.Windows.Forms.MessageBox]::Show("Report copied to clipboard!", "Copied", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        })
        $buttonPanel.Controls.Add($copyButton)
        
        # Add export button
        $exportButton = New-Object System.Windows.Forms.Button
        $exportButton.Text = "Export to File"
        $exportButton.Location = New-Object System.Drawing.Point(140, 10)
        $exportButton.Size = New-Object System.Drawing.Size(120, 30)
        $exportButton.add_Click({
            $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
            $saveDialog.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
            $saveDialog.FileName = "AdminUsersReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
            if ($saveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                $reportTextBox.Text | Out-File -FilePath $saveDialog.FileName -Encoding UTF8
                [System.Windows.Forms.MessageBox]::Show("Report exported to: $($saveDialog.FileName)", "Exported", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            }
        })
        $buttonPanel.Controls.Add($exportButton)
        
        # Add close button
        $closeButton = New-Object System.Windows.Forms.Button
        $closeButton.Text = "Close"
        $closeButton.Location = New-Object System.Drawing.Point(270, 10)
        $closeButton.Size = New-Object System.Drawing.Size(80, 30)
        $closeButton.add_Click({ $reportForm.Close() })
        $buttonPanel.Controls.Add($closeButton)
        
        # Show the form
        $reportForm.ShowDialog()
        
        $mainForm.Cursor = [System.Windows.Forms.Cursors]::Default
        $statusLabel.Text = "Admin report generated successfully"
        
    } catch {
        $mainForm.Cursor = [System.Windows.Forms.Cursors]::Default
        $statusLabel.Text = "Error generating admin report: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Error generating admin report: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

# Update Reset Password button to always show confirmation before applying
$entraResetPasswordButton.add_Click({
    $entraUserGrid.EndEdit()
    $selectedUpns = @()
    for ($i = 0; $i -lt $entraUserGrid.Rows.Count; $i++) {
        if ($entraUserGrid.Rows[$i].Cells["Select"].Value -eq $true) {
            $upn = $entraUserGrid.Rows[$i].Cells["UserPrincipalName"].Value
            if (-not [string]::IsNullOrWhiteSpace($upn)) { $selectedUpns += $upn }
        }
    }
    if ($selectedUpns.Count -ne 1) {
        [System.Windows.Forms.MessageBox]::Show("Select exactly one user to reset password.", "Select One User", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }
    $userUpn = $selectedUpns[0]
    $confirm = [System.Windows.Forms.MessageBox]::Show("Reset password for user $userUpn? This will generate a new password and require the user to change it at next sign-in.", "Confirm Password Reset", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }
    
    # Generate memorable password with validation
    try {
        $newPassword = New-XKCDPassword -WordCount 4 -IncludeSeparator
        
        # Validate password was generated
        if ([string]::IsNullOrWhiteSpace($newPassword)) {
            throw "Password generation failed - generated password is null or empty"
        }
        
        # Additional validation - ensure password meets minimum requirements
        if ($newPassword.Length -lt 8) {
            throw "Generated password is too short (length: $($newPassword.Length))"
        }
        
        Write-Host "Generated password length: $($newPassword.Length)" -ForegroundColor Green
        
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to generate password: $($_.Exception.Message)`n`nTrying fallback password generation...", "Password Generation Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        
        # Fallback password generation
        try {
            $newPassword = "TempPass" + (Get-Random -Minimum 1000 -Maximum 9999) + "!"
            Write-Host "Using fallback password: $newPassword" -ForegroundColor Yellow
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Failed to generate fallback password: $($_.Exception.Message)", "Password Generation Failed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            $statusLabel.Text = "Password generation failed"
            return
        }
    }
    
    try {
        $statusLabel.Text = "Resetting password for $userUpn..."
        $mainForm.Refresh()
        $context = Get-MgContext -ErrorAction Stop
        if (-not $context) { throw "Not connected to Microsoft Graph. Please connect first." }
        
        # Validate user exists before attempting password reset
        try {
            $user = Get-MgUser -UserId $userUpn -ErrorAction Stop
            if (-not $user) {
                throw "User not found: $userUpn"
            }
        } catch {
            throw "Failed to validate user $userUpn : $($_.Exception.Message)"
        }
        
        $passwordProfile = @{ Password = $newPassword; ForceChangePasswordNextSignIn = $true }
        Update-MgUser -UserId $userUpn -PasswordProfile $passwordProfile -ErrorAction Stop
        
        $message = "Password reset successful for user: $userUpn`n`nNew Password: $newPassword`n`nThis password is memorable and secure. The user will be required to change it on next sign-in.`n`nCopy password to clipboard?"
        
        $result = [System.Windows.Forms.MessageBox]::Show($message, "Password Reset Successful", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Information)
        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            [System.Windows.Forms.Clipboard]::SetText($newPassword)
            [System.Windows.Forms.MessageBox]::Show("Password copied to clipboard!", "Copied", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
        $statusLabel.Text = "Password reset completed for $userUpn"
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to reset password for $userUpn : $($_.Exception.Message)", "Password Reset Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        $statusLabel.Text = "Password reset failed for $userUpn"
    }
})

# --- Keyboard Shortcuts ---
$mainForm.add_KeyDown({
    param($sender, $e)
    switch ($e.KeyCode) {
        "O" { if ($e.Control) { $connectButton.PerformClick() } }
        "D" { if ($e.Control) { $disconnectButton.PerformClick() } }
        "S" { if ($e.Control) { $getRulesButton.PerformClick() } }
        "F5" { 
            if ($tabControl.SelectedTab -eq $exchangeTab) {
                # Refresh Exchange data
                if ($connectButton.Enabled -eq $false) {
                    $connectButton.PerformClick()
                }
            } elseif ($tabControl.SelectedTab -eq $entraTab) {
                # Refresh Entra data
                if ($entraConnectGraphButton.Enabled -eq $false) {
                    $entraConnectGraphButton.PerformClick()
                }
            }
        }
        "A" { if ($e.Control) { 
            if ($tabControl.SelectedTab -eq $exchangeTab) {
                $selectAllButton.PerformClick()
            } elseif ($tabControl.SelectedTab -eq $entraTab) {
                for ($i = 0; $i -lt $entraUserGrid.Rows.Count; $i++) {
                    $entraUserGrid.Rows[$i].Cells["Select"].Value = $true
                }
            }
        }}
        "Escape" { $mainForm.Close() }
    }
})

# --- Show Form ---
# Remove all auto-connect logic from the form's Shown event
$mainForm.Add_Shown({ 
    $mainForm.Activate()
    
    # Force Entra ID grid headers to be visible
    $entraUserGrid.ColumnHeadersVisible = $true
    $entraUserGrid.EnableHeadersVisualStyles = $true
    $entraUserGrid.ColumnHeadersHeight = 30
    $entraUserGrid.PerformLayout()
    $entraUserGrid.Refresh()

    # Force the panel to refresh as well
    $entraGridPanel.PerformLayout()
    $entraGridPanel.Refresh()

    # Force grid headers to be properly set
    $entraUserGrid.ColumnHeadersHeight = 25
    $entraUserGrid.ColumnHeadersVisible = $true
    $entraUserGrid.EnableHeadersVisualStyles = $true
    $entraUserGrid.PerformLayout()
    $entraUserGrid.Refresh()

    


    
})
[void]$mainForm.ShowDialog()

# --- Script End ---
Write-Host "Script finished."
# No automatic disconnect on GUI close. User must use the "Disconnect" button.
# if ($script:currentExchangeConnection) { Write-Host "Disconnecting from Exchange Online..."; Disconnect-ExchangeOnline -Confirm:$false -EA SilentlyContinue }
# if ($script:graphConnection) { Write-Host "Disconnecting from Microsoft Graph..."; Disconnect-MgGraph -EA SilentlyContinue }

# --- Open Last Export button event handler ---
$entraOpenLastExportButton.add_Click({
    if ($script:lastExportedXlsxPath) {
        if (Test-Path $script:lastExportedXlsxPath) {
            try {
                $statusLabel.Text = "Opening: $script:lastExportedXlsxPath"
                Invoke-Item -Path $script:lastExportedXlsxPath -ErrorAction Stop
            } catch {
                $statusLabel.Text = "Failed to open: $script:lastExportedXlsxPath"
                [System.Windows.Forms.MessageBox]::Show("Could not open file: $($_.Exception.Message)\nPath: $script:lastExportedXlsxPath", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        } else {
            $statusLabel.Text = "File not found: $script:lastExportedXlsxPath"
            [System.Windows.Forms.MessageBox]::Show("No file exported or file not found.\nPath: $script:lastExportedXlsxPath", "No File", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
    } else {
        $statusLabel.Text = "No export path set."
        [System.Windows.Forms.MessageBox]::Show("No file exported yet.", "No File", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
})

# After every successful export, ensure the Open Last Export button is enabled
# (This is already handled by setting $script:lastExportedXlsxPath, but reinforce if needed)

# --- Disconnect Entra button event handler ---
$entraDisconnectGraphButton.add_Click({
    try {
        Disconnect-MgGraph -ErrorAction Stop
        $script:graphConnection = $null
        $entraUserGrid.Rows.Clear()
        $loadAllUsersButton.Enabled = $false
        $searchUsersButton.Enabled = $false
        $entraConnectGraphButton.Enabled = $true
        $entraDisconnectGraphButton.Enabled = $false
        $statusLabel.Text = "Disconnected from Microsoft Graph."
    } catch {
        $statusLabel.Text = "Error disconnecting from Microsoft Graph: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Error disconnecting from Microsoft Graph: $($_.Exception.Message)", "Disconnect Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

# Fix Module Conflicts button handler
$entraFixModulesButton.add_Click({
    $statusLabel.Text = "🔧 Fixing Microsoft Graph module conflicts..."
    $mainForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $entraFixModulesButton.Enabled = $false

    try {
        # Import the GraphOnline module
        Import-Module "$PSScriptRoot\Modules\GraphOnline.psm1" -Force -ErrorAction Stop

        # Run manual fix commands directly
        try {
            # Step 1: Disconnect and remove current modules
            $statusLabel.Text = "Step 1: Removing existing Microsoft Graph modules..."
            Disconnect-MgGraph -ErrorAction SilentlyContinue
            Get-Module -Name "Microsoft.Graph*" -All | Remove-Module -Force -ErrorAction SilentlyContinue
            Uninstall-Module -Name "Microsoft.Graph*" -AllVersions -Force -ErrorAction SilentlyContinue

            # Step 2: Clear module cache
            $statusLabel.Text = "Step 2: Clearing module cache..."
            Get-Module -Name "Microsoft.Graph*" -ListAvailable | ForEach-Object {
                try {
                    Remove-Item -Path $_.ModuleBase -Recurse -Force -ErrorAction SilentlyContinue
                } catch {
                    # Ignore errors removing module directories
                }
            }

            # Step 3: Reinstall required modules
            $statusLabel.Text = "Step 3: Reinstalling Microsoft Graph modules..."
            $modulesToInstall = @(
                "Microsoft.Graph.Authentication",
                "Microsoft.Graph.Users",
                "Microsoft.Graph.Users.Actions",
                "Microsoft.Graph.Identity.SignIns",
                "Microsoft.Graph.Reports"
            )

            foreach ($module in $modulesToInstall) {
                try {
                    Install-Module -Name $module -Scope CurrentUser -Repository PSGallery -Force -AllowClobber -ErrorAction Stop
                    Write-Host "✓ $module installed successfully" -ForegroundColor Green
                } catch {
                    Write-Host "✗ Failed to install $module`: $($_.Exception.Message)" -ForegroundColor Red
                    throw "Failed to install required modules"
                }
            }

            # Step 4: Clear authentication cache
            $statusLabel.Text = "Step 4: Clearing authentication cache..."
            try {
                $graphSession = [Microsoft.Graph.PowerShell.Authentication.GraphSession]::Instance
                if ($graphSession -and $graphSession.AuthContext) {
                    $graphSession.AuthContext.ClearTokenCache()
                }
            } catch {
                # Ignore errors clearing token cache
            }

            $statusLabel.Text = "✅ Microsoft Graph module conflicts fixed! Please restart PowerShell."
            [System.Windows.Forms.MessageBox]::Show(
                "Microsoft Graph module conflicts have been resolved!`n`n" +
                "✅ All conflicting modules uninstalled`n" +
                "✅ Required modules reinstalled with compatible versions`n" +
                "✅ Authentication cache cleared`n`n" +
                "Please restart PowerShell and try connecting again.",
                "Module Conflicts Fixed",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )

        } catch {
            $statusLabel.Text = "❌ Error fixing module conflicts: $($_.Exception.Message)"
            [System.Windows.Forms.MessageBox]::Show(
                "Error fixing Microsoft Graph module conflicts: $($_.Exception.Message)`n`n" +
                "Please manually run these commands:`n`n" +
                "1. Uninstall-Module Microsoft.Graph* -AllVersions -Force`n" +
                "2. Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force`n" +
                "3. Install-Module Microsoft.Graph.Users -Scope CurrentUser -Force`n" +
                "4. Install-Module Microsoft.Graph.Identity.SignIns -Scope CurrentUser -Force`n" +
                "5. Install-Module Microsoft.Graph.Reports -Scope CurrentUser -Force`n" +
                "6. Restart PowerShell",
                "Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
        }

    } catch {
        $statusLabel.Text = "❌ Error fixing module conflicts: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Error fixing module conflicts: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    } finally {
        $entraFixModulesButton.Enabled = $true
        $mainForm.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})

# Configure grids to auto-expand horizontally
$userMailboxGrid.AutoSizeColumnsMode = 'Fill'
$entraUserGrid.AutoSizeColumnsMode = 'Fill'

# Add a catch-all event to always enable the button after any grid change
$userMailboxGrid.add_SelectionChanged({ $manageRulesButton.Enabled = $true })
$userMailboxGrid.add_CellValueChanged({ $manageRulesButton.Enabled = $true })

# Add event handlers for Entra user grid to update button states
$entraUserGrid.add_CellContentClick({ UpdateEntraButtonStates })
$entraUserGrid.add_CellValueChanged({ UpdateEntraButtonStates })

# Enhanced error handling and resilience functions
function Invoke-WithRetry {
    param(
        [Parameter(Mandatory=$true)]
        [scriptblock]$ScriptBlock,
        [int]$MaxRetries = 3,
        [int]$DelaySeconds = 2,
        [string]$OperationName = "Operation",
        [System.Windows.Forms.ToolStripStatusLabel]$StatusLabel = $null
    )
    
    for ($attempt = 1; $attempt -le $MaxRetries; $attempt++) {
        try {
            if ($StatusLabel) { $StatusLabel.Text = "$OperationName (Attempt $attempt/$MaxRetries)..." }
            $result = & $ScriptBlock
            if ($StatusLabel) { $StatusLabel.Text = "$OperationName completed successfully." }
            return $result
        } catch {
            $errorMsg = $_.Exception.Message
            if ($attempt -lt $MaxRetries) {
                if ($StatusLabel) { $StatusLabel.Text = "$OperationName failed (Attempt $attempt/$MaxRetries). Retrying in $DelaySeconds seconds..." }
                Write-Warning "$OperationName failed (Attempt $attempt/$MaxRetries): $errorMsg. Retrying in $DelaySeconds seconds..."
                Start-Sleep -Seconds $DelaySeconds
            } else {
                if ($StatusLabel) { $StatusLabel.Text = "$OperationName failed after $MaxRetries attempts." }
                Write-Error "$OperationName failed after $MaxRetries attempts: $errorMsg"
                throw
            }
        }
    }
}

function Test-ConnectionHealth {
    param(
        [string]$ConnectionType = "Both"
    )
    
    $health = @{
        ExchangeOnline = $false
        MicrosoftGraph = $false
        LastCheck = Get-Date
    }
    
    if ($ConnectionType -in @("Exchange", "Both")) {
        try {
            $null = Get-PSSession | Where-Object { $_.ConfigurationName -eq "Microsoft.Exchange" }
            $health.ExchangeOnline = $true
        } catch {
            $health.ExchangeOnline = $false
        }
    }
    
    if ($ConnectionType -in @("Graph", "Both")) {
        try {
            $context = Get-MgContext -ErrorAction SilentlyContinue
            $health.MicrosoftGraph = $context -and $context.Account
        } catch {
            $health.MicrosoftGraph = $false
        }
    }
    
    return $health
}

# Performance optimization - Caching system
$script:dataCache = @{
    Mailboxes = $null
    Users = $null
    TransportRules = $null
    Connectors = $null
    LastRefresh = $null
    CacheExpiryMinutes = 5
}

function Get-CachedData {
    param(
        [string]$DataType,
        [scriptblock]$FetchScript,
        [int]$ExpiryMinutes = 5
    )
    
    $cacheKey = $DataType
    $now = Get-Date
    
    # Check if cache exists and is still valid
    if ($script:dataCache[$cacheKey] -and 
        $script:dataCache.LastRefresh -and 
        ($now - $script:dataCache.LastRefresh).TotalMinutes -lt $ExpiryMinutes) {
        return $script:dataCache[$cacheKey]
    }
    
    # Fetch fresh data
    try {
        $data = & $FetchScript
        $script:dataCache[$cacheKey] = $data
        $script:dataCache.LastRefresh = $now
        return $data
    } catch {
        Write-Warning "Failed to fetch $DataType data: $($_.Exception.Message)"
        return $script:dataCache[$cacheKey] # Return stale data if available
    }
}

function Clear-DataCache {
    param([string]$DataType = "All")
    
    if ($DataType -eq "All") {
        $script:dataCache = @{
            Mailboxes = $null
            Users = $null
            TransportRules = $null
            Connectors = $null
            LastRefresh = $null
            CacheExpiryMinutes = 5
        }
    } else {
        $script:dataCache[$DataType] = $null
    }
}

# Function to generate memorable but strong passwords
function New-MemorablePassword {
    [CmdletBinding()]
    param(
        [int]$WordCount = 4,
        [switch]$IncludeNumbers,
        [switch]$IncludeSymbols,
        [switch]$CapitalizeWords
    )
    
    # Common word list (you can expand this)
    $words = @(
        "apple", "banana", "cherry", "dragon", "eagle", "forest", "garden", "house", "island", "jungle",
        "knight", "lemon", "mountain", "ocean", "planet", "queen", "river", "sunset", "tiger", "umbrella",
        "village", "window", "yellow", "zebra", "anchor", "bridge", "castle", "diamond", "elephant", "firefly",
        "guitar", "hammer", "iceberg", "jacket", "kangaroo", "lighthouse", "moonlight", "notebook", "octopus", "penguin",
        "rainbow", "sailboat", "telescope", "umbrella", "volcano", "waterfall", "xylophone", "yacht", "zucchini"
    )
    
    # Generate random words
    $selectedWords = $words | Get-Random -Count $WordCount
    
    # Capitalize if requested
    if ($CapitalizeWords) {
        $selectedWords = $selectedWords | ForEach-Object { $_.Substring(0,1).ToUpper() + $_.Substring(1) }
    }
    
    # Join words
    $password = $selectedWords -join ""
    
    # Add numbers if requested
    if ($IncludeNumbers) {
        $numbers = 0..9 | Get-Random -Count 2
        $password += $numbers -join ""
    }
    
    # Add symbols if requested
    if ($IncludeSymbols) {
        $symbols = @("!", "@", "#", "$", "%", "^", "&", "*", "(", ")", "-", "_", "+", "=", "?", ".")
        $password += $symbols | Get-Random -Count 2
    }
    
    return $password
}

# Function to generate XKCD-style passphrase
function New-XKCDPassword {
    [CmdletBinding()]
    param(
        [int]$WordCount = 4,
        [switch]$IncludeSeparator
    )
    
    try {
        # Common words (expanded list)
        $words = @(
            "correct", "horse", "battery", "staple", "apple", "banana", "cherry", "dragon", "eagle", "forest",
            "garden", "house", "island", "jungle", "knight", "lemon", "mountain", "ocean", "planet", "queen",
            "river", "sunset", "tiger", "umbrella", "village", "window", "yellow", "zebra", "anchor", "bridge",
            "castle", "diamond", "elephant", "firefly", "guitar", "hammer", "iceberg", "jacket", "kangaroo",
            "lighthouse", "moonlight", "notebook", "octopus", "penguin", "rainbow", "sailboat", "telescope",
            "volcano", "waterfall", "xylophone", "yacht", "zucchini", "butterfly", "caterpillar", "dolphin",
            "flamingo", "giraffe", "hedgehog", "iguana", "jellyfish", "koala", "llama", "meerkat", "narwhal",
            "ostrich", "panda", "quokka", "raccoon", "sloth", "toucan", "unicorn", "vulture", "walrus", "xenops"
        )
        
        # Validate word count
        if ($WordCount -lt 1 -or $WordCount -gt 10) {
            throw "Word count must be between 1 and 10"
        }
        
        # Generate random words
        $selectedWords = $words | Get-Random -Count $WordCount
        
        # Validate selected words
        if (-not $selectedWords -or $selectedWords.Count -eq 0) {
            throw "Failed to generate random words"
        }
        
        # Join with separator if requested
        if ($IncludeSeparator) {
            $separators = @("-", "_", ".", "!")
            $password = ""
            for ($i = 0; $i -lt $selectedWords.Count; $i++) {
                $password += $selectedWords[$i]
                if ($i -lt $selectedWords.Count - 1) {
                    $password += $separators | Get-Random
                }
            }
        } else {
            $password = $selectedWords -join ""
        }
        
        # Final validation
        if ([string]::IsNullOrWhiteSpace($password)) {
            throw "Generated password is null or empty"
        }
        
        if ($password.Length -lt 8) {
            throw "Generated password is too short (length: $($password.Length))"
        }
        
        Write-Host "Generated XKCD password: $password (length: $($password.Length))" -ForegroundColor Green
        return $password
        
    } catch {
        Write-Host "Error in New-XKCDPassword: $($_.Exception.Message)" -ForegroundColor Red
        
        # Fallback to simple password generation
        $fallbackPassword = "Secure" + (Get-Random -Minimum 1000 -Maximum 9999) + "!"
        Write-Host "Using fallback password: $fallbackPassword" -ForegroundColor Yellow
        return $fallbackPassword
    }
}

# Function to generate pattern-based password
function New-PatternPassword {
    [CmdletBinding()]
    param(
        [string]$Pattern = "WWNSS"  # W=Word, N=Number, S=Symbol
    )
    
    $words = @("apple", "banana", "cherry", "dragon", "eagle", "forest", "garden", "house", "island", "jungle")
    $numbers = @("123", "456", "789", "2024", "2025", "99", "88", "77", "66", "55")
    $symbols = @("!", "@", "#", "$", "%", "^", "&", "*", "(", ")", "-", "_", "+", "=", "?", ".")
    
    $password = ""
    foreach ($char in $Pattern.ToCharArray()) {
        switch ($char) {
            "W" { $password += $words | Get-Random }
            "N" { $password += $numbers | Get-Random }
            "S" { $password += $symbols | Get-Random }
        }
    }
    
    return $password
}

