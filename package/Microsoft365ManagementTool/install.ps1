# Microsoft 365 Management Tool Installer
# Run as Administrator to install

param(
    [string]$InstallPath = "$env:ProgramFiles\Microsoft365ManagementTool"
)

Write-Host "Installing Microsoft 365 Management Tool..." -ForegroundColor Green

# Check if running as Administrator
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "This installer requires Administrator privileges." -ForegroundColor Red
    Write-Host "Please run PowerShell as Administrator and try again." -ForegroundColor Yellow
    exit 1
}

# Create installation directory
if (-not (Test-Path $InstallPath)) {
    New-Item -ItemType Directory -Path $InstallPath -Force
}

# Copy files
Write-Host "Copying application files..." -ForegroundColor Yellow
Copy-Item ".\*" -Destination $InstallPath -Recurse -Force

# Create Start Menu shortcut
$startMenuPath = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Microsoft 365 Management Tool"
if (-not (Test-Path $startMenuPath)) {
    New-Item -ItemType Directory -Path $startMenuPath -Force
}

$WshShell = New-Object -comObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut("$startMenuPath\Microsoft 365 Management Tool.lnk")
$Shortcut.TargetPath = "$InstallPath\bin\Launch.ps1"
$Shortcut.WorkingDirectory = "$InstallPath\bin"
$Shortcut.Description = "Microsoft 365 Management Tool"
$Shortcut.Save()

# Create Desktop shortcut
$desktopPath = [Environment]::GetFolderPath("Desktop")
$Shortcut = $WshShell.CreateShortcut("$desktopPath\Microsoft 365 Management Tool.lnk")
$Shortcut.TargetPath = "$InstallPath\bin\Launch.ps1"
$Shortcut.WorkingDirectory = "$InstallPath\bin"
$Shortcut.Description = "Microsoft 365 Management Tool"
$Shortcut.Save()

Write-Host "✅ Installation completed successfully!" -ForegroundColor Green
Write-Host "Shortcuts created in Start Menu and Desktop" -ForegroundColor Green
Write-Host "Installation path: $InstallPath" -ForegroundColor Cyan
