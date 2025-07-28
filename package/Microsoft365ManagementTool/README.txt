# Microsoft 365 Management Tool Package

## Installation Options

### Option 1: Quick Install (Recommended)
1. Right-click on install.ps1
2. Select "Run as Administrator"
3. Follow the prompts

### Option 2: Manual Installation
1. Copy the entire package folder to your desired location
2. Run in\Launch.ps1 to start the application
3. Optionally create shortcuts to in\Launch.ps1

### Option 3: Create Executable
1. Install PS2EXE: Install-Module -Name ps2exe -Force
2. Run: ps2exe bin\launcher.ps1 bin\Microsoft365ManagementTool.exe
3. The .exe file can be pinned to taskbar

## Features
- Exchange Online management and analysis
- Entra ID (Azure AD) investigation tools
- Professional report generation
- Incident remediation checklists
- Security posture assessment

## Requirements
- Windows 10/11
- PowerShell 5.1 or later
- Microsoft 365 admin credentials
- Internet connection for Microsoft Graph API

## Support
For issues or questions, check the documentation in the docs folder.
