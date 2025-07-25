# Enhanced Exchange Online Analyzer GUI

A comprehensive PowerShell-based GUI tool for analyzing Exchange Online inbox rules, managing user accounts, and monitoring security configurations. This tool provides administrators with powerful capabilities to detect suspicious inbox rules, manage user access, and export detailed reports.

![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue)
![Exchange Online](https://img.shields.io/badge/Exchange-Online-orange)
![Microsoft Graph](https://img.shields.io/badge/Microsoft-Graph-green)

## 🚀 Features

### Core Functionality
- **📧 Inbox Rules Analysis**: Comprehensive analysis of Exchange Online inbox rules with suspicious activity detection
- **🔍 Auto-Domain Detection**: Automatically detects organization domains from loaded mailboxes
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
- **🚫 Sign-in Control**: Block/unblock user sign-ins through Azure AD
- **📮 Sending Restrictions**: Manage user email sending restrictions (button now at top right)
- **🔐 Permissions Audit**: View mailbox delegates and full access permissions

### Advanced Features
- **🚦 Transport Rules Viewer**: Review and export Exchange Online transport rules
- **🔌 Connectors Management**: View and manage inbound/outbound connectors
- **📝 Rule Management**: Interactive interface to view and delete specific inbox rules
- **📈 Progress Tracking**: Real-time progress indicators for long-running operations
- **❓ Help Tab**: In-app Help tab displays the full README for user reference

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
Install-Module Microsoft.Graph.Identity.SignIns -Scope CurrentUser -Force
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
   .\Enhanced_Exchange_Analyzer_GUI_v6_FIXED.ps1
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

### Advanced Features

#### Microsoft Graph Integration
- Click "Connect to MS Graph" for enhanced user management features
- Enable session revocation and sign-in blocking capabilities

#### Transport Rules and Connectors
- Use dedicated buttons to view and export transport rules
- Review inbound/outbound connector configurations

#### Rule Management
- Select a single mailbox and click "Manage Rules" to view/delete specific rules
- Interactive interface for rule administration

#### Help Tab
- Access the full README and documentation from the in-app Help tab

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

### v6.4 (Unreleased)
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

**Made with ❤️ for Exchange Online administrators**

## Entra ID Investigator (Preview)
A new tabbed interface for investigating Entra ID (Azure AD) accounts is now available. This feature is in preview and will be expanded in future updates.

### Features
- **Sign-in Logs (Export):** Export user sign-in logs to XLSX for forensic review.
- **User Details & Roles:** View user account details, admin roles, and group memberships.
- **User Audit Logs:** Review recent administrative actions for selected users.
- **MFA Analysis:** Analyze and export user MFA status, including per-user, security defaults, and conditional access policies.

### Requirements
- Microsoft Graph PowerShell modules: Users, Reports, Identity.DirectoryManagement, Identity.SignIns
- Delegated permissions: User.Read.All, AuditLog.Read.All, Directory.Read.All, Policy.Read.All, UserAuthenticationMethod.Read.All
- Microsoft Excel (for XLSX export)