# Microsoft 365 Management Tool Launcher
# This script launches the main application with proper error handling

param(
    [string]$LogPath = "$env:TEMP\ExchangeOnlineAnalyzer.log"
)

# Set up logging
function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    Add-Content -Path $LogPath -Value $logMessage
}

# Create log file
New-Item -ItemType File -Path $LogPath -Force | Out-Null
Write-Log "Application started"

try {
    # Get the directory where this script is located
    $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
    
    # Set the working directory to the script directory
    Set-Location $scriptDir
    
    # Import required modules
    Write-Log "Loading modules..."
    Import-Module "$scriptDir\Modules\*.psm1" -Force -ErrorAction SilentlyContinue
    
    # Check if main script exists
    $mainScript = Join-Path $scriptDir "365analyzerv7.ps1"
    if (-not (Test-Path $mainScript)) {
        throw "Main script not found: $mainScript"
    }
    
    Write-Log "Launching main application..."
    
    # Launch the main application
    & $mainScript
    
    Write-Log "Application completed successfully"
    
} catch {
    Write-Log "Error: $($_.Exception.Message)" -Level "ERROR"
    
    # Show error message to user
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.MessageBox]::Show(
        "An error occurred while starting the application:`n`n$($_.Exception.Message)`n`nCheck the log file for details: $LogPath",
        "Microsoft 365 Management Tool - Error",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    )
} 