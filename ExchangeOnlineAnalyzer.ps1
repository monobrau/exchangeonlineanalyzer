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
Version: 8.3
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

#Requires -Version 5.1

if (-not $PSScriptRoot) { $PSScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path }

Add-Type -AssemblyName System.Windows.Forms

Import-Module (Join-Path $PSScriptRoot 'Modules\Logging.psm1') -Global -ErrorAction Stop
Import-Module (Join-Path $PSScriptRoot 'Modules\WinFormsHelpers.psm1') -Force -ErrorAction Stop
Import-Module (Join-Path $PSScriptRoot 'Modules\ExchangeOnlineAnalyzerApp.psm1') -Force -ErrorAction Stop

Start-ExchangeOnlineAnalyzer -AppRoot $PSScriptRoot

