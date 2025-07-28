# Microsoft 365 Management Tool Uninstaller
# Run as Administrator to uninstall

param(
    [string]$InstallPath = "$env:ProgramFiles\Microsoft365ManagementTool"
)

Write-Host "Uninstalling Microsoft 365 Management Tool..." -ForegroundColor Green

# Check if running as Administrator
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "This uninstaller requires Administrator privileges." -ForegroundColor Red
    exit 1
}

# Remove shortcuts
$startMenuPath = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Microsoft 365 Management Tool"
if (Test-Path "$startMenuPath\Microsoft 365 Management Tool.lnk") {
    Remove-Item "$startMenuPath\Microsoft 365 Management Tool.lnk" -Force
}

$desktopPath = [Environment]::GetFolderPath("Desktop")
if (Test-Path "$desktopPath\Microsoft 365 Management Tool.lnk") {
    Remove-Item "$desktopPath\Microsoft 365 Management Tool.lnk" -Force
}

# Remove installation directory
if (Test-Path $InstallPath) {
    Remove-Item $InstallPath -Recurse -Force
}

Write-Host "✅ Uninstallation completed successfully!" -ForegroundColor Green
