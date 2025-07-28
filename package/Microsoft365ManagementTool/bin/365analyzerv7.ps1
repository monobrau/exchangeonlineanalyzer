PS K:\exchangeonlineanalyzer\exchangeonlineanalyzer> # Build Script for Exchange Online Analyzer
PS K:\exchangeonlineanalyzer\exchangeonlineanalyzer> # This script converts the PowerShell application to an executable
PS K:\exchangeonlineanalyzer\exchangeonlineanalyzer> 
PS K:\exchangeonlineanalyzer\exchangeonlineanalyzer> param(
>>     [string]$OutputPath = ".\dist",
>>     [string]$AppName = "ExchangeOnlineAnalyzer"
>> )
PS K:\exchangeonlineanalyzer\exchangeonlineanalyzer> 
PS K:\exchangeonlineanalyzer\exchangeonlineanalyzer> Write-Host "Building Exchange Online Analyzer Executable..." -ForegroundColor Green
Building Exchange Online Analyzer Executable...
PS K:\exchangeonlineanalyzer\exchangeonlineanalyzer> 
PS K:\exchangeonlineanalyzer\exchangeonlineanalyzer> # Create output directory
PS K:\exchangeonlineanalyzer\exchangeonlineanalyzer> if (-not (Test-Path $OutputPath)) {
>>     New-Item -ItemType Directory -Path $OutputPath -Force
>> }

    Directory: K:\exchangeonlineanalyzer\exchangeonlineanalyzer

Mode                 LastWriteTime         Length Name
----                 -------------         ------ ----
d----           7/26/2025  2:25 PM                dist

PS K:\exchangeonlineanalyzer\exchangeonlineanalyzer> 
PS K:\exchangeonlineanalyzer\exchangeonlineanalyzer> # Check if PS2EXE is installed
PS K:\exchangeonlineanalyzer\exchangeonlineanalyzer> try {
>>     $ps2exeVersion = Get-Command ps2exe -ErrorAction Stop
>>     Write-Host "PS2EXE found: $($ps2exeVersion.Version)" -ForegroundColor Green
>> } catch {
>>     Write-Host "PS2EXE not found. Installing..." -ForegroundColor Yellow
>>     Install-Module -Name ps2exe -Force -Scope CurrentUser
>> }
PS2EXE not found. Installing...

NuGet provider is required to continue
This version of PowerShellGet requires minimum version '2.8.5.201' of NuGet provider to publish an item to NuGet-based repositories. The NuGet provider must be available in 'C:\Program Files\PackageManagement\ProviderAssemblies' or
'C:\Users\cknos\AppData\Local\PackageManagement\ProviderAssemblies'. You can also install the NuGet provider by running 'Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force'. Do you want PowerShellGet to install and import  
the NuGet provider now?
[Y] Yes  [N] No  [S] Suspend  [?] Help (default is "Y"): y
Install-PackageProvider: Unhandled Exception - Message:'The type initializer for 'Microsoft.PackageManagement.Internal.Utility.Extensions.FilesystemExtensions' threw an exception.' Name:'TypeInitializationException' Stack Trace:'   at
Microsoft.PackageManagement.Internal.Utility.Extensions.FilesystemExtensions.MakeSafeFileName(String input)    at Microsoft.PackageManagement.Internal.Utility.Plugin.DynamicType.DefineDynamicType(Type interfaceType)    at
Microsoft.PackageManagement.Internal.Utility.Plugin.DynamicType..ctor(Type interfaceType, OrderedDictionary`2 methods, List`2 delegates, List`1 stubs)    at
Microsoft.PackageManagement.Internal.Utility.Plugin.DynamicType.<>c__DisplayClass9_0.<Create>b__3()    at Microsoft.PackageManagement.Internal.Utility.Extensions.DictionaryExtensions.GetOrAdd[TKey,TValue](IDictionary`2 dictionary, TKey key, 
Func`1 valueFunction)    at Microsoft.PackageManagement.Internal.Utility.Plugin.DynamicType.Create(Type tInterface, OrderedDictionary`2 instanceMethods, List`2 delegateMethods, List`1 stubMethods, List`2 usedInstances)    at
Microsoft.PackageManagement.Internal.Utility.Plugin.DynamicInterface.CreateProxy(Type tInterface, Object[] instances)    at Microsoft.PackageManagement.Internal.Utility.Plugin.DynamicInterface.DynamicCast(Type tInterface, Object[]
instances)    at Microsoft.PackageManagement.Internal.Utility.Plugin.DynamicInterface.DynamicCast[TInterface](Object[] instances)    at Microsoft.PackageManagement.Internal.Utility.Plugin.DynamicInterfaceExtensions.As[TInterface](Object     
instance)    at Microsoft.PowerShell.PackageManagement.Cmdlets.CmdletBase.get_PackageManagementHost()    at Microsoft.PowerShell.PackageManagement.Cmdlets.CmdletBase.SelectProviders(String[] names)    at
Microsoft.PowerShell.PackageManagement.Cmdlets.CmdletWithProvider.get_SelectedProviders()    at Microsoft.PowerShell.PackageManagement.Cmdlets.InstallPackageProvider.get_SelectedProviders()    at
Microsoft.PowerShell.PackageManagement.Cmdlets.CmdletWithProvider.<get_CachedSelectedProviders>b__23_0()    at Microsoft.PackageManagement.Internal.Utility.Extensions.DictionaryExtensions.GetOrAdd[TKey,TValue](IDictionary`2 dictionary, TKey 
key, Func`1 valueFunction)    at Microsoft.PackageManagement.Internal.Utility.Extensions.Singleton`1.GetOrAdd(Func`1 newInstance, Object primaryKey, Object[] keys)    at
Microsoft.PackageManagement.Internal.Utility.Extensions.SingletonExtensions.GetOrAdd[TResult](Object primaryKey, Func`1 newInstance, Object[] keys)    at
Microsoft.PowerShell.PackageManagement.Cmdlets.CmdletWithProvider.get_CachedSelectedProviders()    at Microsoft.PowerShell.PackageManagement.Cmdlets.CmdletWithProvider.GenerateDynamicParameters()    at
Microsoft.PowerShell.PackageManagement.Cmdlets.AsyncCmdlet.<>c__DisplayClass85_0.<AsyncRun>b__0()'
Import-PackageProvider: No match was found for the specified search criteria and provider name 'NuGet'. Try 'Get-PackageProvider -ListAvailable' to see if the provider exists on the system.
Get-PackageProvider: Unhandled Exception - Message:'The type initializer for 'Microsoft.PackageManagement.Internal.Utility.Extensions.FilesystemExtensions' threw an exception.' Name:'TypeInitializationException' Stack Trace:'   at
Microsoft.PackageManagement.Internal.Utility.Extensions.FilesystemExtensions.MakeSafeFileName(String input)    at Microsoft.PackageManagement.Internal.Utility.Plugin.DynamicType.DefineDynamicType(Type interfaceType)    at
Microsoft.PackageManagement.Internal.Utility.Plugin.DynamicType..ctor(Type interfaceType, OrderedDictionary`2 methods, List`2 delegates, List`1 stubs)    at
Microsoft.PackageManagement.Internal.Utility.Plugin.DynamicType.<>c__DisplayClass9_0.<Create>b__3()    at Microsoft.PackageManagement.Internal.Utility.Extensions.DictionaryExtensions.GetOrAdd[TKey,TValue](IDictionary`2 dictionary, TKey key, 
Func`1 valueFunction)    at Microsoft.PackageManagement.Internal.Utility.Plugin.DynamicType.Create(Type tInterface, OrderedDictionary`2 instanceMethods, List`2 delegateMethods, List`1 stubMethods, List`2 usedInstances)    at
Microsoft.PackageManagement.Internal.Utility.Plugin.DynamicInterface.CreateProxy(Type tInterface, Object[] instances)    at Microsoft.PackageManagement.Internal.Utility.Plugin.DynamicInterface.DynamicCast(Type tInterface, Object[]
instances)    at Microsoft.PackageManagement.Internal.Utility.Plugin.DynamicInterface.DynamicCast[TInterface](Object[] instances)    at Microsoft.PackageManagement.Internal.Utility.Plugin.DynamicInterfaceExtensions.As[TInterface](Object     
instance)    at Microsoft.PowerShell.PackageManagement.Cmdlets.CmdletBase.get_PackageManagementHost()    at Microsoft.PowerShell.PackageManagement.Cmdlets.CmdletBase.SelectProviders(String name)    at
Microsoft.PowerShell.PackageManagement.Cmdlets.GetPackageProvider.ProcessProvidersFilteredByName()    at Microsoft.PowerShell.PackageManagement.Cmdlets.GetPackageProvider.ProcessRecordAsync()    at
Microsoft.PowerShell.PackageManagement.Cmdlets.AsyncCmdlet.<>c__DisplayClass85_0.<AsyncRun>b__0()'
Install-Module: 
Line |
   6 |      Install-Module -Name ps2exe -Force -Scope CurrentUser
     |      ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
     | NuGet provider is required to interact with NuGet-based repositories. Please ensure that '2.8.5.201' or newer version of NuGet provider is installed.
PS K:\exchangeonlineanalyzer\exchangeonlineanalyzer> 
PS K:\exchangeonlineanalyzer\exchangeonlineanalyzer> # Build the executable
PS K:\exchangeonlineanalyzer\exchangeonlineanalyzer> Write-Host "Converting PowerShell script to executable..." -ForegroundColor Yellow
Converting PowerShell script to executable...
PS K:\exchangeonlineanalyzer\exchangeonlineanalyzer> 
PS K:\exchangeonlineanalyzer\exchangeonlineanalyzer> $ps2exeArgs = @(
>>     "365analyzerv7.ps1",
>>     "$OutputPath\$AppName.exe",
>>     "-noConsole",
>>     "-noVisualStyles",
>>     "-noError",
>>     "-title", "Microsoft 365 Management Tool",
>>     "-version", "7.0",
>>     "-company", "Exchange Online Analyzer",
>>     "-product", "Microsoft 365 Management Tool",
>>     "-description", "Comprehensive Microsoft 365 security and management analysis tool"
>> )
PS K:\exchangeonlineanalyzer\exchangeonlineanalyzer> 
PS K:\exchangeonlineanalyzer\exchangeonlineanalyzer> & ps2exe @ps2exeArgs
&: The term 'ps2exe' is not recognized as a name of a cmdlet, function, script file, or executable program.
Check the spelling of the name, or if a path was included, verify that the path is correct and try again.
PS K:\exchangeonlineanalyzer\exchangeonlineanalyzer> 
PS K:\exchangeonlineanalyzer\exchangeonlineanalyzer> if (Test-Path "$OutputPath\$AppName.exe") {
>>     Write-Host "✅ Executable created successfully: $OutputPath\$AppName.exe" -ForegroundColor Green
>>
>>     # Create a shortcut for easy access
>>     $WshShell = New-Object -comObject WScript.Shell
>>     $Shortcut = $WshShell.CreateShortcut("$OutputPath\$AppName.lnk")
>>     $Shortcut.TargetPath = "$OutputPath\$AppName.exe"
>>     $Shortcut.WorkingDirectory = $OutputPath
>>     $Shortcut.Description = "Microsoft 365 Management Tool"
>>     $Shortcut.IconLocation = "$OutputPath\$AppName.exe,0"
>>     $Shortcut.Save()
>>
>>     Write-Host "✅ Shortcut created: $OutputPath\$AppName.lnk" -ForegroundColor Green
>>     Write-Host "`n📋 Instructions:" -ForegroundColor Cyan
>>     Write-Host "1. Copy the .exe file to your desired location" -ForegroundColor White
>>     Write-Host "2. Right-click the .exe and select 'Pin to taskbar' or 'Pin to Start'" -ForegroundColor White
>>     Write-Host "3. The application will run as a Windows GUI application" -ForegroundColor White
>> } else {
>>     Write-Host "❌ Failed to create executable" -ForegroundColor Red
>> }
❌ Failed to create executable
PS K:\exchangeonlineanalyzer\exchangeonlineanalyzer> <#
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

# Debug message box removed - script is confirmed working

# Import all modules with error handling
function Safe-ImportModule($modulePath) {
    try {
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
    $entraBlockUserButton.Enabled = $true
    $entraUnblockUserButton.Enabled = $true
    $entraRevokeSessionsButton.Enabled = $true
}

# Function to generate professional report
function Generate-ProfessionalReport {
    $report = @"
# Microsoft 365 Environment Analysis Report
**Generated:** $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
**Tool:** Microsoft 365 Management Tool v7.0

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
- **Tool Version:** 7.0
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
    $note += "Tool: Microsoft 365 Management Tool v7.0`n"
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
        $rowIdx = $unifiedAccountGrid.Rows.Add(
            $false, 
            $account.UPN, 
            $account.DisplayName, 
            $account.ExchangeStatus, 
            $account.EntraStatus
        )
        
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
    $report += "Tool: Microsoft 365 Management Tool v7.0`n`n"
    
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
                $report += "  - Licensed: $($account.Licensed)`n`n"
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
    $report += "- Tool Version: 7.0`n"
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
    $note += "- Tool Version: 7.0`n"
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
    @{Name="Microsoft.Graph.Identity.SignIns"; MinVersion="2.0"}
)
$script:graphScopes = @(
    "User.Read.All",
    "User.ReadWrite.All",
    "SecurityEvents.Read.All",
    "SecurityEvents.ReadWrite.All"
)

# --- GUI Setup ---
Add-Type -AssemblyName System.Windows.Forms; Add-Type -AssemblyName System.Drawing
$mainForm = New-Object System.Windows.Forms.Form; $mainForm.Text = "Microsoft 365 Management Tool"; $mainForm.Size = New-Object System.Drawing.Size(1100, 900); $mainForm.MinimumSize = New-Object System.Drawing.Size(900, 700); $mainForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen; $mainForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable; $mainForm.MaximizeBox = $true; $mainForm.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi
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

$keywordsLabel = New-Object System.Windows.Forms.Label
$keywordsLabel.Text = "Keywords:"

$keywordsTextBox = New-Object System.Windows.Forms.TextBox
$keywordsTextBox.Width = 200

$outputFolderLabel = New-Object System.Windows.Forms.Label
$outputFolderLabel.Text = "Output Folder:"

$outputFolderTextBox = New-Object System.Windows.Forms.TextBox
$outputFolderTextBox.Width = 200

$browseFolderButton = New-Object System.Windows.Forms.Button
$browseFolderButton.Text = "Browse..."
$browseFolderButton.Width = 100

$getRulesButton = New-Object System.Windows.Forms.Button
$getRulesButton.Text = "Export Rules"
$getRulesButton.Width = 120
$getRulesButtonTooltip = New-Object System.Windows.Forms.ToolTip
$getRulesButtonTooltip.SetToolTip($getRulesButton, "Export inbox rules for selected mailboxes (Ctrl+S)")

$manageRulesButton = New-Object System.Windows.Forms.Button
$manageRulesButton.Text = "Manage Rules"
$manageRulesButton.Width = 120
$manageRulesButton.Enabled = $true

$openFileButton = New-Object System.Windows.Forms.Button
$openFileButton.Text = "Open Last File"
$openFileButton.Width = 120

$blockUserButton = New-Object System.Windows.Forms.Button
$blockUserButton.Text = "Block User"
$blockUserButton.Width = 100
$blockUserButton.Enabled = $true

$unblockUserButton = New-Object System.Windows.Forms.Button
$unblockUserButton.Text = "Unblock User"
$unblockUserButton.Width = 100
$unblockUserButton.Enabled = $true

$revokeSessionsButton = New-Object System.Windows.Forms.Button
$revokeSessionsButton.Text = "Revoke Sessions"
$revokeSessionsButton.Width = 120

$manageRestrictedSendersButton = New-Object System.Windows.Forms.Button
$manageRestrictedSendersButton.Text = "Manage Restricted Senders"
$manageRestrictedSendersButton.Width = 180

$manageConnectorsButton = New-Object System.Windows.Forms.Button
$manageConnectorsButton.Text = "Manage Connectors"
$manageConnectorsButton.Width = 140

$manageTransportRulesButton = New-Object System.Windows.Forms.Button
$manageTransportRulesButton.Text = "Manage Transport Rules"
$manageTransportRulesButton.Width = 160

# Replace CheckedListBox with DataGridView for mailbox list
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

# Define columns
$colCheck = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
$colCheck.HeaderText = "Select"
$colCheck.Width = 40
$colCheck.Name = "Select"
$colCheck.ReadOnly = $false
$userMailboxGrid.Columns.Add($colCheck)

$colUPN = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colUPN.HeaderText = "UserPrincipalName"
$colUPN.DataPropertyName = "UserPrincipalName"
$colUPN.Width = 220
$colUPN.ReadOnly = $true
$userMailboxGrid.Columns.Add($colUPN)

$colDisplayName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colDisplayName.HeaderText = "DisplayName"
$colDisplayName.DataPropertyName = "DisplayName"
$colDisplayName.Width = 180
$colDisplayName.ReadOnly = $true
$userMailboxGrid.Columns.Add($colDisplayName)

$colBlocked = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colBlocked.HeaderText = "SignInBlocked"
$colBlocked.DataPropertyName = "SignInBlocked"
$colBlocked.Width = 100
$colBlocked.ReadOnly = $true
$userMailboxGrid.Columns.Add($colBlocked)

# Add columns for rule analysis
$colRulesCount = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colRulesCount.HeaderText = "RulesCount"
$colRulesCount.DataPropertyName = "RulesCount"
$colRulesCount.Width = 80
$colRulesCount.ReadOnly = $true
$userMailboxGrid.Columns.Add($colRulesCount)

$colSuspiciousRules = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colSuspiciousRules.HeaderText = "SuspiciousRules"
$colSuspiciousRules.DataPropertyName = "SuspiciousRules"
$colSuspiciousRules.Width = 100
$colSuspiciousRules.ReadOnly = $true
$userMailboxGrid.Columns.Add($colSuspiciousRules)

$colExternalForwarding = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colExternalForwarding.HeaderText = "ExternalForwarding"
$colExternalForwarding.DataPropertyName = "ExternalForwarding"
$colExternalForwarding.Width = 120
$colExternalForwarding.ReadOnly = $true
$userMailboxGrid.Columns.Add($colExternalForwarding)

$colDelegates = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colDelegates.HeaderText = "Delegates"
$colDelegates.DataPropertyName = "Delegates"
$colDelegates.Width = 80
$colDelegates.ReadOnly = $true
$userMailboxGrid.Columns.Add($colDelegates)

$colFullAccess = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colFullAccess.HeaderText = "FullAccess"
$colFullAccess.DataPropertyName = "FullAccess"
$colFullAccess.Width = 80
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
    foreach ($mbx in $script:allLoadedMailboxUPNs) {
        if ([string]::IsNullOrWhiteSpace($searchText) -or 
            $mbx.UserPrincipalName -like "*$searchText*" -or 
            $mbx.DisplayName -like "*$searchText*") {
            # Get rule analysis for this mailbox
            $rulesCount = "0"
            $suspiciousRules = "0"
            $externalForwarding = "Unknown"
            $delegates = "Unknown"
            $fullAccess = "Unknown"
            
            try {
                $rules = Get-InboxRule -Mailbox $mbx.UserPrincipalName -IncludeHidden -ErrorAction SilentlyContinue
                if ($rules) {
                    $rulesCount = $rules.Count.ToString()
                    $suspiciousCount = 0
                    $hasExternalForwarding = $false
                    
                    foreach ($rule in $rules) {
                        # Check for suspicious keywords
                        foreach ($kw in $BaseSuspiciousKeywords) {
                            if ($rule.Name -and $rule.Name -match [regex]::Escape($kw)) {
                                $suspiciousCount++
                                break
                            }
                        }
                        
                        # Check for symbols-only names
                        if ($rule.Name -and $rule.Name.Length -gt 0) {
                            $textCharacters = $rule.Name -replace '[^\p{L}\p{N}\s]', ''
                            if ([string]::IsNullOrWhiteSpace($textCharacters)) {
                                $suspiciousCount++
                            }
                        }
                        
                        # Check for hidden rules
                        if ($rule.IsHidden) {
                            $suspiciousCount++
                        }
                        
                        # Check for external forwarding
                        if ($rule.ForwardTo -and $rule.ForwardTo -match '@') {
                            $hasExternalForwarding = $true
                        }
                    }
                    
                    $suspiciousRules = $suspiciousCount.ToString()
                    $externalForwarding = if ($hasExternalForwarding) { "Yes" } else { "No" }
                }
            } catch {
                # Keep default values if analysis fails
            }
            
            $rowIdx = $userMailboxGrid.Rows.Add($false, $mbx.UserPrincipalName, $mbx.DisplayName, $mbx.SignInBlocked, $mbx.RecipientTypeDetails, $rulesCount, $suspiciousRules, $externalForwarding, $delegates, $fullAccess)
        }
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

$entraDisconnectGraphButton = New-Object System.Windows.Forms.Button
$entraDisconnectGraphButton.Text = "Disconnect Entra"
$entraDisconnectGraphButton.Width = 140

$entraOutputFolderLabel = New-Object System.Windows.Forms.Label
$entraOutputFolderLabel.Text = "Export Folder:"
$entraOutputFolderTextBox = New-Object System.Windows.Forms.TextBox
$entraOutputFolderTextBox.Width = 300
$entraBrowseFolderButton = New-Object System.Windows.Forms.Button
$entraBrowseFolderButton.Text = "Browse..."
$entraBrowseFolderButton.Width = 100

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

$entraSignInExportXlsxButton  = New-Object System.Windows.Forms.Button
$entraSignInExportXlsxButton.Text = "Export Sign-in XLSX"
$entraSignInExportXlsxButton.Width = 140
$entraSignInExportXlsxButton.Enabled = $false

$entraDetailsFetchButton      = New-Object System.Windows.Forms.Button
$entraDetailsFetchButton.Text = "User Details && Roles"
$entraDetailsFetchButton.Width = 140

$entraAuditFetchButton        = New-Object System.Windows.Forms.Button
$entraAuditFetchButton.Text   = "Fetch Audit Logs"
$entraAuditFetchButton.Width = 140

$entraAuditExportXlsxButton   = New-Object System.Windows.Forms.Button
$entraAuditExportXlsxButton.Text = "Export Audit XLSX"
$entraAuditExportXlsxButton.Width = 140
$entraAuditExportXlsxButton.Enabled = $false

$entraMfaFetchButton          = New-Object System.Windows.Forms.Button
$entraMfaFetchButton.Text     = "Analyze MFA"
$entraMfaFetchButton.Width = 120

# Add user management buttons for Entra ID tab
$entraBlockUserButton = New-Object System.Windows.Forms.Button
$entraBlockUserButton.Text = "Block User"
$entraBlockUserButton.Width = 100
$entraBlockUserButton.Enabled = $false

$entraUnblockUserButton = New-Object System.Windows.Forms.Button
$entraUnblockUserButton.Text = "Unblock User"
$entraUnblockUserButton.Width = 100
$entraUnblockUserButton.Enabled = $false

$entraRevokeSessionsButton = New-Object System.Windows.Forms.Button
$entraRevokeSessionsButton.Text = "Revoke Sessions"
$entraRevokeSessionsButton.Width = 120
$entraRevokeSessionsButton.Enabled = $false

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

# Top action panel for Connect/Disconnect/Select All/Deselect All/Block/Unblock User/Revoke Sessions/Manage Rules
$topActionPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$topActionPanel.Dock = 'Top'
$topActionPanel.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
$topActionPanel.WrapContents = $false
$topActionPanel.AutoSize = $true
$topActionPanel.Controls.AddRange(@($connectButton, $disconnectButton, $selectAllButton, $deselectAllButton, $manageRulesButton, $manageRestrictedSendersButton, $manageConnectorsButton, $manageTransportRulesButton))

# Add search to top action panel
$topActionPanel.Controls.Add($exchangeSearchLabel)
$topActionPanel.Controls.Add($exchangeSearchTextBox)

$exchangeTab.Controls.Add($topActionPanel)

# Panel for mailbox label and grid (fills remaining space)
$mailboxPanel = New-Object System.Windows.Forms.Panel
$mailboxPanel.Dock = 'Fill'
$mailboxPanel.Padding = New-Object System.Windows.Forms.Padding(5, 5, 5, 5)
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

# Test button removed - focusing on fixing the actual layout

# --- Entra ID Investigator Tab Layout (REBUILT FROM SCRATCH) ---
$entraTab = New-Object System.Windows.Forms.TabPage; $entraTab.Text = "Entra ID Investigator"

# Top action panel (EXACTLY like Exchange Online)
$entraTopPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$entraTopPanel.Dock = 'Top'
$entraTopPanel.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
$entraTopPanel.WrapContents = $false
$entraTopPanel.AutoSize = $true

$entraTopPanel.Controls.AddRange(@($entraConnectGraphButton, $entraDisconnectGraphButton, $entraViewSignInLogsButton, $entraViewAuditLogsButton, $entraDetailsFetchButton))

# Panel for user grid (EXACTLY like Exchange Online)
$entraGridPanel = New-Object System.Windows.Forms.Panel
$entraGridPanel.Dock = 'Fill'
$entraGridPanel.Padding = New-Object System.Windows.Forms.Padding(5, 30, 5, 15)
$entraGridPanel.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 0)

# User grid (simplified)
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

# Define columns
$colEntraCheck = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
$colEntraCheck.HeaderText = "Select"
$colEntraCheck.Name = "Select"
$colEntraCheck.ReadOnly = $false
$entraUserGrid.Columns.Add($colEntraCheck)

$colEntraUPN = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colEntraUPN.HeaderText = "UserPrincipalName"
$colEntraUPN.Name = "UserPrincipalName"
$colEntraUPN.ReadOnly = $true
$entraUserGrid.Columns.Add($colEntraUPN)

$colEntraDisplayName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colEntraDisplayName.HeaderText = "DisplayName"
$colEntraDisplayName.Name = "DisplayName"
$colEntraDisplayName.ReadOnly = $true
$entraUserGrid.Columns.Add($colEntraDisplayName)

$colEntraLicensed = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colEntraLicensed.HeaderText = "Licensed"
$colEntraLicensed.Name = "Licensed"
$colEntraLicensed.ReadOnly = $true
$entraUserGrid.Columns.Add($colEntraLicensed)

$entraGridPanel.Controls.Add($entraUserGrid)

# Simple bottom panel with buttons
$entraBottomPanel = New-Object System.Windows.Forms.Panel
$entraBottomPanel.Dock = 'Bottom'
$entraBottomPanel.Height = 70

# Add buttons to bottom panel with better spacing and larger size
$entraBrowseFolderButton.Location = New-Object System.Drawing.Point(10, 15)
$entraBrowseFolderButton.Size = New-Object System.Drawing.Size(120, 30)
$entraExportSignInLogsButton.Location = New-Object System.Drawing.Point(140, 15)
$entraExportSignInLogsButton.Size = New-Object System.Drawing.Size(140, 30)
$entraExportAuditLogsButton.Location = New-Object System.Drawing.Point(290, 15)
$entraExportAuditLogsButton.Size = New-Object System.Drawing.Size(140, 30)
$entraOpenLastExportButton.Location = New-Object System.Drawing.Point(440, 15)
$entraOpenLastExportButton.Size = New-Object System.Drawing.Size(120, 30)

# Add export path controls to the right of buttons
$entraOutputFolderLabel.Location = New-Object System.Drawing.Point(580, 18)
$entraOutputFolderTextBox.Location = New-Object System.Drawing.Point(680, 15)
$entraOutputFolderTextBox.Width = 200
$entraOutputFolderTextBox.Height = 25

$entraBottomPanel.Controls.AddRange(@($entraBrowseFolderButton, $entraExportSignInLogsButton, $entraExportAuditLogsButton, $entraOpenLastExportButton, $entraOutputFolderLabel, $entraOutputFolderTextBox))

# Add panels to tab in order
$entraTab.Controls.Add($entraTopPanel)
$entraTab.Controls.Add($entraGridPanel)
$entraTab.Controls.Add($entraBottomPanel)

Write-Host "=== REBUILT ENTRA ID LAYOUT ==="
Write-Host "Entra Tab Controls: $($entraTab.Controls.Count)"
Write-Host "Top Panel Controls: $($entraTopPanel.Controls.Count)"
Write-Host "Grid Panel Controls: $($entraGridPanel.Controls.Count)"
Write-Host "Bottom Panel Controls: $($entraBottomPanel.Controls.Count)"
Write-Host "Grid Header Height: $($entraUserGrid.ColumnHeadersHeight)"
Write-Host "Grid Size: $($entraUserGrid.Size)"
Write-Host "Grid Location: $($entraUserGrid.Location)"
Write-Host "Browse Button Size: $($entraBrowseFolderButton.Size)"
Write-Host "Export Sign-in Button Size: $($entraExportSignInLogsButton.Size)"

# Test removed - script execution confirmed

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

    try {
        if (Connect-EntraGraph) {
            $script:graphConnection = $true
            try {
                $users = Get-EntraUsers
                $entraUserGrid.Rows.Clear()
                foreach ($u in $users) {
                    try {
                        $userDetails = Get-MgUser -UserId $u.UserPrincipalName -Property AssignedLicenses
                        $isLicensed = $userDetails.AssignedLicenses.Count -gt 0
                    } catch {
                        $isLicensed = $false
                    }
                    $licensedText = if ($isLicensed) { "Licensed" } else { "Unlicensed" }
                    $entraUserGrid.Rows.Add($false, $u.UserPrincipalName, $u.DisplayName, $licensedText)
                    UpdateEntraButtonStates
                }
                $statusLabel.Text = "Connected to Microsoft Graph. Users loaded."
                $entraSignInExportButton.Enabled = $true; $entraDetailsFetchButton.Enabled = $true; $entraAuditFetchButton.Enabled = $true; $entraMfaFetchButton.Enabled = $true
                # User management buttons are always enabled when connected to Graph
                $entraBlockUserButton.Enabled = $true
                $entraUnblockUserButton.Enabled = $true
                $entraRevokeSessionsButton.Enabled = $true
                
                # Force headers to be visible after data is loaded
                $entraUserGrid.ColumnHeadersVisible = $true
                $entraUserGrid.EnableHeadersVisualStyles = $true
                $entraUserGrid.ColumnHeadersHeight = 25
                $entraUserGrid.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::DisableResizing
                $entraUserGrid.PerformLayout()
                $entraUserGrid.Refresh()
                
                # Force the panel to refresh as well
                $entraGridPanel.PerformLayout()
                $entraGridPanel.Refresh()
                
                UpdateEntraButtonStates
            } catch {
                $statusLabel.Text = "Failed to load users: $($_.Exception.Message)"
                [System.Windows.Forms.MessageBox]::Show("Failed to load users: $($_.Exception.Message)", "Graph Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
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
            $details = "User Principal Name: $($result.User.UserPrincipalName)`r`nDisplay Name: $($result.User.DisplayName)`r`nAccount Enabled: $($result.User.AccountEnabled)`r`nLast Password Change: $($result.User.LastPasswordChangeDateTime)`r`n" +
                "-----------------------------`r`nRoles:`r`n" +
                ($result.Roles.Count -gt 0 ? ($result.Roles -join "`r`n") : "None") +
                "`r`n-----------------------------`r`nGroups:`r`n" +
                ($result.Groups.Count -gt 0 ? ($result.Groups -join "`r`n") : "None")
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
    try {
        Connect-ExchangeOnline -ErrorAction Stop
        $script:currentExchangeConnection = $true
        Show-Progress -message "Connected. Loading mailboxes..." -progress 25
        
        $mailboxes = Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox,SharedMailbox,RoomMailbox,EquipmentMailbox | Select-Object UserPrincipalName, DisplayName, AccountDisabled, IsLicensed, RecipientTypeDetails | Sort-Object UserPrincipalName
        Show-Progress -message "Processing mailboxes..." -progress 50
        
        $userMailboxGrid.Rows.Clear()
        $script:allLoadedMailboxUPNs = @()  # Store for domain detection
        $totalMailboxes = $mailboxes.Count
        $processedCount = 0
        
        foreach ($mbx in $mailboxes) {
            $script:allLoadedMailboxUPNs += $mbx.UserPrincipalName
            try {
                $user = Get-User -Identity $mbx.UserPrincipalName -ErrorAction Stop
                if ($null -ne $user.AccountDisabled) {
                    $signInBlocked = if ($user.AccountDisabled) { "Blocked" } else { "Allowed" }
                } else {
                    $signInBlocked = "Unknown"
                }
            } catch {
                $signInBlocked = "Unknown"
            }
            # Get rule analysis for this mailbox
            $rulesCount = "0"
            $suspiciousRules = "0"
            $externalForwarding = "Unknown"
            $delegates = "Unknown"
            $fullAccess = "Unknown"
            
            try {
                $rules = Get-InboxRule -Mailbox $mbx.UserPrincipalName -IncludeHidden -ErrorAction SilentlyContinue
                if ($rules) {
                    $rulesCount = $rules.Count.ToString()
                    $suspiciousCount = 0
                    $hasExternalForwarding = $false
                    
                    foreach ($rule in $rules) {
                        # Check for suspicious keywords
                        foreach ($kw in $BaseSuspiciousKeywords) {
                            if ($rule.Name -and $rule.Name -match [regex]::Escape($kw)) {
                                $suspiciousCount++
                                break
                            }
                        }
                        
                        # Check for symbols-only names
                        if ($rule.Name -and $rule.Name.Length -gt 0) {
                            $textCharacters = $rule.Name -replace '[^\p{L}\p{N}\s]', ''
                            if ([string]::IsNullOrWhiteSpace($textCharacters)) {
                                $suspiciousCount++
                            }
                        }
                        
                        # Check for hidden rules
                        if ($rule.IsHidden) {
                            $suspiciousCount++
                        }
                        
                        # Check for external forwarding
                        if ($rule.ForwardTo -and $rule.ForwardTo -match '@') {
                            $hasExternalForwarding = $true
                        }
                    }
                    
                    $suspiciousRules = $suspiciousCount.ToString()
                    $externalForwarding = if ($hasExternalForwarding) { "Yes" } else { "No" }
                }
            } catch {
                # Keep default values if analysis fails
            }
            
            $rowIdx = $userMailboxGrid.Rows.Add($false, $mbx.UserPrincipalName, $mbx.DisplayName, $signInBlocked, $mbx.RecipientTypeDetails, $rulesCount, $suspiciousRules, $externalForwarding, $delegates, $fullAccess)
            $processedCount++
            if ($processedCount % 10 -eq 0) {
                Show-Progress -message "Processing mailboxes... ($processedCount/$totalMailboxes)" -progress (50 + ($processedCount / $totalMailboxes * 40))
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
        $selectAllButton.Enabled = $true; $deselectAllButton.Enabled = $true; $disconnectButton.Enabled = $true; $connectButton.Enabled = $false
        # Enable/disable action buttons based on selection
        $manageRulesButton.Enabled = $true
        $manageConnectorsButton.Enabled = $true
        $manageTransportRulesButton.Enabled = $true
        $blockUserButton.Enabled = $false
        $unblockUserButton.Enabled = $false
        # No automatic connect to Entra Graph here, user must click their button
        
        Show-Progress -message "Ready. Connected to Exchange Online." -progress -1
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
$disconnectButton.add_Click({
    $statusLabel.Text = "Disconnecting..."; $mainForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try { Disconnect-ExchangeOnline -Confirm:$false } catch {}
    $script:currentExchangeConnection = $null
    $userMailboxGrid.Rows.Clear(); $selectAllButton.Enabled = $false; $deselectAllButton.Enabled = $false; $disconnectButton.Enabled = $false; $connectButton.Enabled = $true
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
$openFileButton.add_Click({
    if ($openFileButton.Tag -and (Test-Path $openFileButton.Tag)) {
        try { Invoke-Item -Path $openFileButton.Tag -ErrorAction Stop } catch { [System.Windows.Forms.MessageBox]::Show("Could not open file: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) }
    } else {
        [System.Windows.Forms.MessageBox]::Show("No file exported or file not found.", "No File", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
})

$manageRestrictedSendersButton.add_Click({
    $userMailboxGrid.EndEdit()
    $checkedRows = @()
    for ($i = 0; $i -lt $userMailboxGrid.Rows.Count; $i++) {
        if ($userMailboxGrid.Rows[$i].Cells["Select"].Value -eq $true) {
            $checkedRows += $userMailboxGrid.Rows[$i]
        }
    }
    if ($checkedRows.Count -eq 1) {
        $row = $checkedRows[0]
        $upnCell = $row.Cells[1].Value  # Use index 1 for UPN
        $upn = if ($upnCell -ne $null) { $upnCell.ToString().Trim() } else { "" }
    } else {
        $upn = ""
    }
    if ($checkedRows.Count -ne 1 -or [string]::IsNullOrWhiteSpace($upn)) {
        [System.Windows.Forms.MessageBox]::Show("Select exactly one mailbox to manage restricted senders.", "Select One Mailbox", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information); return
    }
    [System.Windows.Forms.MessageBox]::Show("DEBUG: About to call Show-RestrictedSenderManagementDialog for UPN: $upn")
    Show-RestrictedSenderManagementDialog -UserPrincipalName $upn -ParentForm $mainForm -StatusLabelGlobal $statusLabel
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

# Note: Top panel is now populated during layout creation

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
$accountSelectorGroup.Location = New-Object System.Drawing.Point(10, 100)
$accountSelectorGroup.Size = New-Object System.Drawing.Size(800, 500)
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
$unifiedAccountGrid.Size = New-Object System.Drawing.Size(760, 400)
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
$refreshAccountsButton.Text = "Refresh Account List"
$refreshAccountsButton.Location = New-Object System.Drawing.Point(10, 160)
$refreshAccountsButton.Size = New-Object System.Drawing.Size(150, 30)
$refreshAccountsButton.Font = New-Object System.Drawing.Font('Segoe UI', 9)

$selectAllAccountsButton = New-Object System.Windows.Forms.Button
$selectAllAccountsButton.Text = "Select All"
$selectAllAccountsButton.Location = New-Object System.Drawing.Point(170, 160)
$selectAllAccountsButton.Size = New-Object System.Drawing.Size(100, 30)
$selectAllAccountsButton.Font = New-Object System.Drawing.Font('Segoe UI', 9)

$deselectAllAccountsButton = New-Object System.Windows.Forms.Button
$deselectAllAccountsButton.Text = "Deselect All"
$deselectAllAccountsButton.Location = New-Object System.Drawing.Point(280, 160)
$deselectAllAccountsButton.Size = New-Object System.Drawing.Size(100, 30)
$deselectAllAccountsButton.Font = New-Object System.Drawing.Font('Segoe UI', 9)

$accountSelectorGroup.Controls.AddRange(@($refreshAccountsButton, $selectAllAccountsButton, $deselectAllAccountsButton))

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

# Add account selector group to panel
$reportGeneratorPanel.Controls.Add($accountSelectorGroup)

# Add Report Generator tab to tab control
$tabControl.TabPages.Add($reportGeneratorTab)

# Initialize unified account grid when Report Generator tab is first shown
$reportGeneratorTab.add_Enter({
    if ($unifiedAccountGrid.Rows.Count -eq 0) {
        Update-UnifiedAccountGrid
    }
})

# Add panel to tab
$reportGeneratorTab.Controls.Add($reportGeneratorPanel)

# Account selector button event handlers
$refreshAccountsButton.add_Click({
    try {
        $statusLabel.Text = "Refreshing account list..."
        Update-UnifiedAccountGrid
        $statusLabel.Text = "Account list refreshed"
    } catch {
        $statusLabel.Text = "Error refreshing account list: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Error refreshing account list: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
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

    Write-Host "Test buttons added in Shown event"
    Write-Host "Entra Tab Controls: $($entraTab.Controls.Count)"
    Write-Host "Exchange Tab Controls: $($exchangeTab.Controls.Count)"

    # Debug message box removed - form creation is confirmed working
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
        $statusLabel.Text = "Disconnected from Microsoft Graph."
    } catch {
        $statusLabel.Text = "Error disconnecting from Microsoft Graph: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Error disconnecting from Microsoft Graph: $($_.Exception.Message)", "Disconnect Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
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

