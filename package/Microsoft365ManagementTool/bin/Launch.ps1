# Microsoft 365 Management Tool Launcher
# Run this to start the application

param(
    [switch]$NoConsole
)

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $scriptDir

if ($NoConsole) {
    # Run without console window
    Start-Process powershell.exe -ArgumentList "-ExecutionPolicy Bypass -File "launcher.ps1"" -WindowStyle Hidden
} else {
    # Run with console window
    & ".\launcher.ps1"
}
