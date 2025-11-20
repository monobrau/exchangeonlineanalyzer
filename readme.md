# Microsoft 365 Management Tool

A comprehensive PowerShell-based GUI tool for analyzing Exchange Online inbox rules, managing user accounts, monitoring security configurations, and investigating Entra ID (Azure AD) accounts. This tool provides administrators with powerful capabilities to detect suspicious inbox rules, manage user access, export detailed reports, and perform comprehensive Entra ID investigations.

## 🔧 **Version 8.1 - Now Available!**
**✅ Fixed MS Graph Integration** - User blocking and session revocation in Entra ID now working properly
**✅ Enhanced Module Management** - Automatic import and installation of Microsoft Graph PowerShell modules
**✅ Improved Error Handling** - Better diagnostics and error messages for Graph operations
**✅ Enhanced Security** - Comprehensive function validation and permission checks

![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue)
![Exchange Online](https://img.shields.io/badge/Exchange-Online-orange)
![Microsoft Graph](https://img.shields.io/badge/Microsoft-Graph-green)
![Entra ID](https://img.shields.io/badge/Entra-ID-purple)
![Version](https://img.shields.io/badge/Version-8.1-blue)

## 🚀 Entra Portal Shortcuts (v8.1)

Entra Portal Shortcuts with Firefox Multi‑Account Containers.

Requirements:
- Firefox (desktop)
- Extension: Open external links in a container (by Denys H)
  - Install: https://addons.mozilla.org/en-US/firefox/addon/open-url-in-container/
  - Enable handling of external `ext+container` links in the add‑on options

What it does:
- Opens Entra portal deep links directly in a chosen Firefox container
- Auto‑matches the best container by tenant name/domain (if Graph/EXO connected)

How to use (UI):
- Report Generator tab → Entra Portal Shortcuts (Preview)
- Choose Firefox profile and container (or accept auto‑match)
- Click a shortcut (Sign‑in Logs / Restricted Entities / Conditional Access)
- If the extension isn’t installed, links open as normal tabs

Quick test (CLI):
```powershell
Start-Process 'firefox.exe' -ArgumentList (
  "ext+container:name=MyContainer&url={0}" -f [uri]::EscapeUriString(
    'https://entra.microsoft.com/#view/Microsoft_AAD_UsersAndTenants/UserManagementMenuBlade/~/SignIns/menuId/'
  )
)
```

## 🚀 Features

### Core Functionality
- **📧 Inbox Rules Analysis**: Comprehensive analysis of Exchange Online inbox rules with suspicious activity detection
- **🔍 Auto-Domain Detection**: Automatically detects organization domains from loaded mailboxes
- **🔒 Security Investigation**: Comprehensive security analysis combining Exchange Online and Microsoft Graph data
- **📊 XLSX Export**: Formatted Excel reports with advanced formatting:
  - User-friendly columns (MailboxOwner, RuleName, IsEnabled, IsHidden, etc.)
  - All columns auto-fit for width
  - Description column auto-fit for row height and wrapped text
  - RuleID always exported as text (never scientific notation)
  - Conditional formatting: TRUE values highlighted light red, IsHidden rows highlighted bright yellow
  - SuspiciousKeywordsInName column: flags rules with suspicious keywords in the name
- **🎯 External Forwarding Detection**: Identifies rules forwarding emails to external domains
- **🔒 Hidden Rules Detection**: Discovers hidden or system-generated rules (improved logic)

### Security Management
- **👤 User Session Management**: Revoke active user sessions via Microsoft Graph
- **🚫 Sign-in Control**: Block/unblock user sign-ins through Entra ID (Azure AD)
- **📮 Sending Restrictions**: Manage user email sending restrictions (button now at top right)
- **🔐 Permissions Audit**: View mailbox delegates and full access permissions

### 🔍 Security Investigation Report
- **📊 Multi-Source Analysis**: Combines data from Exchange Online and Microsoft Graph
- **📧 Message Trace Collection**: Last 10 days of all email communications
- **📋 Inbox Rules Export**: Complete audit of all mailbox rules across the organization
- **🔐 Sign-in Log Analysis**: Authentication patterns and risk assessments
- **📋 Audit Log Collection**: Directory changes and administrative activities
- **🤖 AI Investigation Prompts**: Structured prompts for AI-powered security analysis
- **📝 Executive Summaries**: Non-technical reports for management and ticketing systems

### Advanced Features
- **🚦 Transport Rules Management**: Review, export, and delete Exchange Online transport rules
- **🔌 Connectors Management**: View, export, and delete inbound/outbound connectors
- **📝 Rule Management**: Interactive interface to view and delete specific inbox rules
- **📈 Progress Tracking**: Real-time progress indicators for long-running operations
- **⌨️ Keyboard Shortcuts**: Quick access to common functions (Ctrl+E for export, Ctrl+S for search, etc.)
- **🔍 Search & Filter**: Real-time search functionality for mailbox and user lists
- **❓ Help Tab**: In-app Help tab displays the full README for user reference
- **💡 Tooltips**: Contextual help for buttons and controls

### Entra ID Investigator
- **👥 User Management**: View, block, unblock, and revoke sessions for Entra ID users
- **📊 Sign-in Logs**: Export detailed sign-in activity logs to XLSX
- **🔍 Audit Logs**: Review administrative actions and security events
- **📈 User Analytics**: Comprehensive user account analysis and reporting
- **🎯 Bulk Operations**: Perform actions on multiple selected users or all loaded users

## 📋 Prerequisites

### Required Software
- **Windows Operating System** with PowerShell 5.1 or later
- **Microsoft Excel** (required for XLSX formatting and conversion)
- **Internet connectivity** for Exchange Online and Microsoft Graph access

### Required PowerShell Modules
The script will automatically prompt to install missing modules:

```powershell
# Exchange Online Management
Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force

# Microsoft Graph Modules
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force
Install-Module Microsoft.Graph.Users -Scope CurrentUser -Force
Install-Module Microsoft.Graph.Users.Actions -Scope CurrentUser -Force
Install-Module Microsoft.Graph.Identity.SignIns -Scope CurrentUser -Force
Install-Module Microsoft.Graph.Reports -Scope CurrentUser -Force
```

### Required Permissions

#### Exchange Online
- **Exchange Administrator** or **Global Administrator** role
- Permissions to read mailbox configurations and inbox rules
- Access to transport rules and connectors

#### Microsoft Graph (Optional but Recommended)
- `User.Read.All` - Read user profiles
- `User.ReadWrite.All` - Manage user accounts
- `SecurityEvents.Read.All` - Read security events
- `SecurityEvents.ReadWrite.All` - Manage security events

## 🔧 Installation

1. **Download the Script**
   ```bash
   git clone https://github.com/yourusername/exchange-online-analyzer.git
   cd exchange-online-analyzer
   ```

2. **Set Execution Policy** (if needed)
   ```powershell
   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
   ```

3. **Run the Script**
   ```powershell
   .\365analyzerv8.ps1
   ```

## 💻 Usage

### Getting Started

1. **Launch the Application**
   - Run the PowerShell script to open the GUI interface

2. **Connect to Exchange Online**
   - Click "Connect & Load Mailboxes" to authenticate and load mailbox data
   - The script will auto-detect organization domains from mailbox UPNs

3. **Configure Analysis Parameters**
   - Review and adjust auto-detected organization domains
   - Modify suspicious keywords if needed
   - Select output folder for reports

4. **Select Mailboxes**
   - Choose specific mailboxes or use "Select All"
   - Single mailbox selection enables rule management features

5. **Generate Reports**
   - Click "Export Rules" to analyze and export rules
   - Reports include suspicious rule detection, external forwarding analysis, and advanced formatting
   - **Security Investigation**: Click the "🔍 Security Investigation Report" button in the Report Generator tab for comprehensive security analysis

### Advanced Features

#### Microsoft Graph Integration
- Click "Connect to MS Graph" for enhanced user management features

#### Security Investigation Report
- Navigate to the "Report Generator" tab
- Click the blue "🔍 Security Investigation Report" button
- Configure investigation parameters:
  - Investigator Name (default: Security Administrator)
  - Company Name (default: Organization)
  - Days to Analyze (1, 3, 7, 10, or 30 days)
- Click "🚀 Generate Security Investigation" to start analysis
- Review results in three tabs:
  - **📋 Investigation Summary** - Technical analysis and findings
  - **🤖 AI Investigation Prompt** - Structured prompt for AI analysis
  - **📝 Non-Technical Summary** - Executive summary for management

#### Troubleshooting Microsoft Graph Connection Issues
If you encounter connection errors like "Method not found" or module version conflicts:

1. **Use the Fix Module Conflicts Button:**
   - In the Entra ID tab, click the yellow/orange "Fix Module Conflicts" button
   - This will automatically uninstall and reinstall all Microsoft Graph modules with compatible versions

2. **Manual Fix (if automatic fix fails):**
   ```powershell
   # Run these commands in PowerShell as Administrator
   Uninstall-Module Microsoft.Graph* -AllVersions -Force
   Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force
   Install-Module Microsoft.Graph.Users -Scope CurrentUser -Force
   Install-Module Microsoft.Graph.Identity.SignIns -Scope CurrentUser -Force
   Install-Module Microsoft.Graph.Reports -Scope CurrentUser -Force
   # Restart PowerShell and try connecting again
   ```

#### Transport Rules and Connectors
- Use dedicated buttons to view, export, and delete transport rules
- Review, export, and delete inbound/outbound connector configurations
- All manager windows are resizable for better usability

#### Rule Management
- Select a single mailbox and click "Manage Rules" to view/delete specific rules
- Interactive interface for rule administration

#### Entra ID Investigator
- Connect to Microsoft Graph to access Entra ID features
- View and manage user accounts, sign-in logs, and audit data
- Export comprehensive reports for security analysis
- Perform bulk operations on multiple users

#### Help Tab
- Access the full README and documentation from the in-app Help tab

#### Keyboard Shortcuts
- **Ctrl+E**: Export rules/data
- **Ctrl+S**: Search/filter lists
- **Ctrl+G**: Connect to Graph
- **Ctrl+X**: Connect to Exchange Online
- **F1**: Show help

## 📊 Report Output

### XLSX Report Features
- **User-Friendly Columns**: MailboxOwner, RuleName, IsEnabled, IsHidden, IsForwardingExternal, IsDeleting, IsMarkingAsRead, IsMovingToFolder, MoveToFolderName, SuspiciousKeywordsInName, Description, StopProcessingRules, Conditions, Actions, Exceptions, RuleID, etc.
- **Conditional Formatting**: TRUE values highlighted light red, IsHidden rows highlighted bright yellow
- **Suspicious Keyword Detection**: SuspiciousKeywordsInName column flags rules with keywords like invoice, payment, password, etc.
- **Comprehensive Data**: Includes rule details, mailbox forwarding, delegates, and permissions
- **Column and Row Auto-Fit**: All columns auto-fit for width, Description column auto-fit for row height and wrapped text
- **RuleID as Text**: RuleID is always exported and formatted as text (never scientific notation)
- **Color Coding**: 
  - 🟡 Yellow highlighting for hidden rules
  - 🔴 Red highlighting for TRUE boolean values

### Report Columns
- Mailbox owner and forwarding settings
- Rule name, priority, and status
- External forwarding detection
- Suspicious keyword matches
- Rule conditions, actions, and exceptions
- Delegate and full access permissions

## ⚙️ Configuration

### Suspicious Keywords
Default keywords include: `invoice`, `payment`, `password`, `confidential`, `urgent`, `bank`, `account`, `auto forward`, `external`, `hidden`

Customize keywords through the GUI interface before running analysis.

### Domain Detection
- **Automatic**: Script auto-detects domains from loaded mailbox UPNs
- **Manual Override**: Edit the organization domains field as needed
- **Priority**: Non-onmicrosoft.com domains are prioritized for detection

## 🔒 Security Considerations

### Data Handling
- All data processing occurs locally on the administrator's machine
- No data is transmitted to third-party services
- Temporary CSV files are automatically cleaned up after XLSX conversion

### Access Control
- Requires appropriate administrative permissions
- Graph features require explicit consent for requested scopes
- Session management affects user access - use with caution

### Best Practices
- Review auto-detected domains before analysis
- Test with small mailbox samples first
- Regularly update PowerShell modules
- Maintain current Exchange Online and Graph permissions

## 🐛 Troubleshooting

### Common Issues

#### Excel COM Errors
```
Solution: Ensure Microsoft Excel is installed and properly licensed
```

#### Module Installation Failures
```powershell
# Run PowerShell as Administrator and retry
Install-Module ExchangeOnlineManagement -Scope AllUsers -Force
```

#### Connection Timeouts
```
Solution: Check network connectivity and firewall settings
Verify Exchange Online and Graph service availability
```

#### Permission Errors
```
Solution: Verify role assignments in Exchange Admin Center
Ensure Graph permissions are properly consented
```

### Debug Mode
Enable verbose output by modifying the script's debug settings or checking console output for detailed error messages.

## 📝 Version History

### v8.1 (Current)
- ✅ Entra Portal Shortcuts with Firefox containers; tenant auto‑match and deep links
- ✅ AI Analysis tab: send report outputs to Gemini or Claude; Settings store API keys
- ✅ Tenant-scoped export folders: Documents\ExchangeOnlineAnalyzer\SecurityInvestigation\<Tenant>\timestamp
- ✅ Transport Rules and Connectors added to Security Investigation report exports
- ✅ MFA coverage and user security groups included in investigation outputs
- ✅ Robust domain/keyword detection on XOL tab for any subset of users
- ✅ Improved Graph connection handling; device code fallback on module conflicts
- ✅ Settings tab persists Investigator/Company and API keys

### v8.0
- ✅ **Fixed MS Graph Integration**: Resolved critical user blocking and session revocation issues in Entra ID
- ✅ **Enhanced Module Management**: Automatic import and installation of Microsoft Graph PowerShell modules
- ✅ **Improved Error Handling**: Better diagnostics and error messages for Graph operations
- ✅ **Enhanced Security**: Added comprehensive function validation and permission checks
- ✅ **Updated Dependencies**: Added Microsoft.Graph.Users.Actions and Microsoft.Graph.Reports module support
- ✅ **Robust API Integration**: Enhanced reliability with fallback mechanisms and detailed error reporting
- ✅ **Security Investigation Report**: Comprehensive security analysis tool with AI prompt generation
- ✅ **Multi-Source Data Collection**: Combines Exchange Online and Microsoft Graph data for complete analysis
- ✅ **Automated Audit Collection**: Pulls maximum audit logs, sign-in logs, message traces, and inbox rules
- ✅ **AI Investigation Prompts**: Generates structured prompts for AI-powered security analysis
- ✅ **Non-Technical Reporting**: Creates executive summaries for ticketing systems and management
- ✅ **Module Conflict Resolution**: Added automatic fix for Microsoft Graph version conflicts
- ✅ **Enhanced Troubleshooting**: Built-in diagnostics and repair tools for connection issues

### v7.0
- ✅ **New Application Name**: Changed from "Exchange Online Analyzer" to "Microsoft 365 Management Tool"
- ✅ **Enhanced GUI Layout**: Improved Entra ID tab layout with proper spacing, headers, and button positioning
- ✅ **Progress Indicators**: Real-time progress tracking for long-running operations
- ✅ **User-Friendly Error Handling**: Improved error messages and recovery options
- ✅ **Keyboard Shortcuts**: Quick access to common functions (Ctrl+E, Ctrl+S, Ctrl+G, Ctrl+X, F1)
- ✅ **Search & Filter**: Real-time search functionality for mailbox and user lists
- ✅ **Tooltips**: Contextual help for buttons and controls
- ✅ **Resizable Manager Windows**: Transport Rules and Connectors managers are now resizable
- ✅ **Delete Functionality**: Added delete capabilities for connectors and transport rules
- ✅ **Enhanced Entra ID Features**: Improved user management, session revocation, and bulk operations
- ✅ **Better Layout Management**: Fixed grid header visibility and control positioning issues
- ✅ **Export Path Controls**: Re-added export path controls to Entra ID tab with proper positioning

### v6.4 (Previous)
- ✅ Improved Excel export formatting: all columns auto-fit for width, Description column auto-fit for row height and wrapped text.
- ✅ RuleID always exported and formatted as text (never scientific notation).
- ✅ Conditional formatting: TRUE values highlighted light red, IsHidden rows highlighted bright yellow.
- ✅ SuspiciousKeywordsInName column: flags rules with suspicious keywords in the name.
- ✅ Help tab added to GUI for in-app documentation.
- ✅ Manage Restricted Senders button moved to top right of Exchange tab.
- ✅ Enhanced hidden inbox rule detection: expanded logic and debug output for more reliable detection and highlighting of hidden/system rules in exports.

### v6.3-FIXED-AUTODOMAINS-GRAPHCONTROL
- ✅ Added automatic domain detection from mailbox UPNs
- ✅ Enhanced Microsoft Graph integration with manual connect/disconnect
- ✅ Improved error handling for user management features
- ✅ Fixed domain prioritization logic
- ✅ Enhanced UI responsiveness and progress tracking

### Previous Versions
- v6.2: Added session revocation capabilities
- v6.1: Introduced transport rules and connectors viewers
- v6.0: Microsoft Graph integration
- v5.x: Core rule analysis and export functionality

## 🤝 Contributing

Contributions are welcome! Please follow these guidelines:

1. **Fork the Repository**
2. **Create a Feature Branch**
   ```bash
   git checkout -b feature/new-feature
   ```
3. **Commit Changes**
   ```bash
   git commit -m "Add new feature description"
   ```
4. **Submit Pull Request**

### Development Guidelines
- Follow PowerShell best practices
- Include error handling for new features
- Update documentation for new functionality
- Test with various Exchange Online configurations

## ⚠️ Disclaimer

This tool is provided as-is for educational and administrative purposes. Always test in non-production environments first. The authors are not responsible for any data loss or system issues resulting from the use of this tool.

## 📞 Support

For issues, questions, or feature requests:
- 🐛 [Open an Issue](../../issues)
- 💬 [Discussions](../../discussions)
- 📧 Contact: [maintainer@email.com]

## 🙏 Acknowledgments

- Microsoft Exchange Online and Graph API teams
- PowerShell community for GUI development patterns
- Contributors and testers who helped improve this tool

---

**Made with ❤️ for Microsoft 365 administrators**

## Entra ID Investigator
A comprehensive tabbed interface for investigating Entra ID (Azure AD) accounts with advanced user management capabilities.

### Features
- **User Management**: Block, unblock, and revoke sessions for Entra ID users
- **Sign-in Logs (Export)**: Export user sign-in logs to XLSX for forensic review
- **User Details & Roles**: View user account details, admin roles, and group memberships
- **User Audit Logs**: Review recent administrative actions for selected users
- **MFA Analysis**: Analyze and export user MFA status, including per-user, security defaults, and conditional access policies
- **Bulk Operations**: Perform actions on multiple selected users or all loaded users
- **Export Controls**: Comprehensive export functionality with customizable output paths

### Requirements
- Microsoft Graph PowerShell modules: Users, Reports, Identity.DirectoryManagement, Identity.SignIns
- Delegated permissions: User.Read.All, AuditLog.Read.All, Directory.Read.All, Policy.Read.All, UserAuthenticationMethod.Read.All
- Microsoft Excel (for XLSX export)