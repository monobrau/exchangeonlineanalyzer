<#
.SYNOPSIS
    ExchangeOnlineAnalyzer logging module - structured logging with levels, rotation, and multiple outputs.
.DESCRIPTION
    Provides configurable logging with:
    - Log levels: Verbose, Debug, Info, Warning, Error
    - File output with rotation (by size and retention)
    - Optional console output
    - Structured JSON format option
    - Session/component scoping for BulkTenantExporter
#>

$script:LogConfig = @{
    LogPath        = $null
    MinLevel       = 'Info'
    ConsoleOutput  = $true
    JsonFormat     = $false
    MaxFileSizeMB  = 10
    RetainDays     = 30
    SessionId      = $null
    CompanyName    = $null
    TicketNumbers  = $null  # array or single string, joined for display
    Component      = $null
    Initialized    = $false
    _fileLock      = $null
    _levelMap      = @{ Verbose = 0; Debug = 1; Info = 2; Warning = 3; Error = 4 }
}

function Get-LogLevelValue {
    param([string]$Level)
    $v = $script:LogConfig._levelMap[$Level]
    if ($null -eq $v) { return 2 }
    return $v
}

function Get-DefaultLogPath {
    $base = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'ExchangeOnlineAnalyzer'
    return Join-Path $base 'Logs'
}

function Initialize-Logger {
    <#
    .SYNOPSIS
        Initialize the logging module. Call once at script startup.
    .PARAMETER LogPath
        Directory for log files. Default: Documents\ExchangeOnlineAnalyzer\Logs
    .PARAMETER MinLevel
        Minimum level to log: Verbose, Debug, Info, Warning, Error
    .PARAMETER ConsoleOutput
        Whether to emit logs to console
    .PARAMETER JsonFormat
        Use JSON for log entries (structured)
    .PARAMETER MaxFileSizeMB
        Rotate log file when it exceeds this size (MB)
    .PARAMETER RetainDays
        Keep logs for this many days; older files are deleted
    .PARAMETER SessionId
        Optional session identifier (e.g. "Client1" for BulkTenantExporter)
    .PARAMETER CompanyName
        Optional company/client name for log line identification
    .PARAMETER TicketNumbers
        Optional ticket number(s) - array or string - for log line identification
    .PARAMETER Component
        Optional component name (e.g. "ExportUtils", "BulkTenantExporter")
    #>
    param(
        [string]$LogPath = (Get-DefaultLogPath),
        [ValidateSet('Verbose', 'Debug', 'Info', 'Warning', 'Error')]
        [string]$MinLevel = 'Info',
        [bool]$ConsoleOutput = $true,
        [bool]$JsonFormat = $false,
        [int]$MaxFileSizeMB = 10,
        [int]$RetainDays = 30,
        [string]$SessionId = $null,
        [string]$CompanyName = $null,
        [array]$TicketNumbers = @(),
        [string]$Component = $null
    )
    try {
        if (-not (Test-Path $LogPath)) {
            New-Item -ItemType Directory -Path $LogPath -Force | Out-Null
        }
        $script:LogConfig.LogPath = $LogPath
        $script:LogConfig.MinLevel = $MinLevel
        $script:LogConfig.ConsoleOutput = $ConsoleOutput
        $script:LogConfig.JsonFormat = $JsonFormat
        $script:LogConfig.MaxFileSizeMB = $MaxFileSizeMB
        $script:LogConfig.RetainDays = $RetainDays
        $script:LogConfig.SessionId = $SessionId
        $script:LogConfig.CompanyName = $CompanyName
        $script:LogConfig.TicketNumbers = if ($TicketNumbers) { @($TicketNumbers) } else { @() }
        $script:LogConfig.Component = $Component
        $script:LogConfig.Initialized = $true
        if ($null -eq $script:LogConfig._fileLock) {
            $script:LogConfig._fileLock = [System.Threading.Mutex]::new($false, "ExchangeOnlineAnalyzer_LogMutex")
        }
        # Remove old log files
        Remove-ExpiredLogFiles -LogPath $LogPath -RetainDays $RetainDays
        return $true
    } catch {
        Write-Warning "Logging initialization failed: $($_.Exception.Message)"
        $script:LogConfig.Initialized = $false
        return $false
    }
}

function Get-LogContextString {
    $parts = @()
    if ($script:LogConfig.CompanyName -and -not [string]::IsNullOrWhiteSpace($script:LogConfig.CompanyName)) {
        $parts += $script:LogConfig.CompanyName.Trim()
    }
    $tickets = $script:LogConfig.TicketNumbers
    if ($tickets -and $tickets.Count -gt 0) {
        $ticketStr = ($tickets | ForEach-Object { $_.ToString().Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join ', '
        if ($ticketStr) { $parts += $ticketStr }
    }
    if ($parts.Count -eq 0) { return $null }
    return $parts -join ' | '
}

function Remove-ExpiredLogFiles {
    param([string]$LogPath, [int]$RetainDays)
    try {
        $cutoff = (Get-Date).AddDays(-$RetainDays)
        Get-ChildItem -Path $LogPath -Filter "*.log" -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime -lt $cutoff } | Remove-Item -Force -ErrorAction SilentlyContinue
    } catch {
        # Ignore cleanup errors
    }
}

function Get-CurrentLogFilePath {
    $base = $script:LogConfig.LogPath
    if (-not $base) { $base = Get-DefaultLogPath }
    $date = Get-Date -Format 'yyyy-MM-dd'
    $suffix = if ($script:LogConfig.SessionId) { "_$($script:LogConfig.SessionId)" } else { '' }
    return Join-Path $base "ExchangeOnlineAnalyzer$suffix`_$date.log"
}

function Write-Log {
    <#
    .SYNOPSIS
        Write a log entry with automatic sanitization of sensitive data.
    .PARAMETER Message
        Log message (will be sanitized)
    .PARAMETER Level
        Verbose, Debug, Info, Warning, Error
    .PARAMETER Component
        Override component for this entry
    .PARAMETER Data
        Hashtable of additional structured data (will be sanitized)
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [ValidateSet('Verbose', 'Debug', 'Info', 'Warning', 'Error')]
        [string]$Level = 'Info',
        [string]$Component = $null,
        [hashtable]$Data = @{}
    )
    
    # SECURITY: Sanitize message and data before logging
    $sanitizedMessage = $Message
    $sanitizedData = @{}
    
    # Try to import SecurityHelpers for sanitization
    $securityHelpersPath = Join-Path $PSScriptRoot '..\Scripts\Common\SecurityHelpers.psm1'
    if (-not (Test-Path $securityHelpersPath)) {
        $securityHelpersPath = Join-Path (Split-Path $PSScriptRoot -Parent) 'Scripts\Common\SecurityHelpers.psm1'
    }
    
    if (Test-Path $securityHelpersPath) {
        try {
            Import-Module $securityHelpersPath -Force -ErrorAction SilentlyContinue
            if (Get-Command Remove-SensitiveDataFromText -ErrorAction SilentlyContinue) {
                $sanitizedMessage = Remove-SensitiveDataFromText -Text $Message
                foreach ($key in $Data.Keys) {
                    $value = if ($Data[$key] -is [string]) {
                        Remove-SensitiveDataFromText -Text $Data[$key]
                    } else {
                        $Data[$key]
                    }
                    $sanitizedData[$key] = $value
                }
            }
        } catch {
            # If sanitization fails, continue with original values (better than not logging)
        }
    } else {
        $sanitizedData = $Data
    }
    if (-not $script:LogConfig.Initialized) {
        $minVal = Get-LogLevelValue -Level $script:LogConfig.MinLevel
        if ($null -eq $minVal) { $minVal = 2 }
        $msgVal = Get-LogLevelValue -Level $Level
        if ($msgVal -lt $minVal) { return }
        if ($script:LogConfig.ConsoleOutput) {
            $color = switch ($Level) {
                'Error'   { 'Red' }
                'Warning' { 'Yellow' }
                'Info'    { 'Cyan' }
                default   { 'Gray' }
            }
            Write-Host "[$Level] $sanitizedMessage" -ForegroundColor $color
        }
        return
    }
    $minVal = Get-LogLevelValue -Level $script:LogConfig.MinLevel
    $msgVal = Get-LogLevelValue -Level $Level
    if ($msgVal -lt $minVal) { return }
    $comp = if ($Component) { $Component } elseif ($script:LogConfig.Component) { $script:LogConfig.Component } else { 'Main' }
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'
    $sessionId = $script:LogConfig.SessionId
    $contextStr = Get-LogContextString
    $logPath = Get-CurrentLogFilePath
    $acquired = $false
    try {
        if ($script:LogConfig._fileLock) {
            $acquired = $script:LogConfig._fileLock.WaitOne(2000)
        } else {
            $acquired = $true
        }
        if ($acquired) {
            if ($script:LogConfig.JsonFormat) {
                $entry = @{
                    Timestamp = $timestamp
                    Level     = $Level
                    Component = $comp
                    SessionId = $sessionId
                    Company   = $script:LogConfig.CompanyName
                    Ticket    = if ($script:LogConfig.TicketNumbers -and $script:LogConfig.TicketNumbers.Count -gt 0) { ($script:LogConfig.TicketNumbers -join ', ') } else { $null }
                    Message   = $sanitizedMessage
                } + $sanitizedData
                $line = $entry | ConvertTo-Json -Compress
            } else {
                $sessionPart = if ($sessionId) { "[$sessionId] " } else { '' }
                $contextPart = if ($contextStr) { "[$contextStr] " } else { '' }
                $line = "[$timestamp] [$Level] $sessionPart$contextPart[$comp] $Message"
            }
            Add-Content -Path $logPath -Value $line -Encoding UTF8 -ErrorAction SilentlyContinue
            # Check rotation
            $fi = Get-Item -Path $logPath -ErrorAction SilentlyContinue
            if ($fi -and ($fi.Length / 1MB) -gt $script:LogConfig.MaxFileSizeMB) {
                $rotated = $logPath -replace '\.log$', "_$(Get-Date -Format 'HHmmss').log"
                Rename-Item -Path $logPath -NewName (Split-Path $rotated -Leaf) -Force -ErrorAction SilentlyContinue
            }
        }
    } catch {
        # Fallback to console if file write fails
        Write-Host "[$Level] $sanitizedMessage" -ForegroundColor $(if ($Level -eq 'Error') { 'Red' } else { 'Gray' })
    } finally {
        if ($acquired -and $script:LogConfig._fileLock) {
            try { $script:LogConfig._fileLock.ReleaseMutex() } catch { }
        }
    }
    if ($script:LogConfig.ConsoleOutput) {
        $color = switch ($Level) {
            'Error'   { 'Red' }
            'Warning' { 'Yellow' }
            'Info'    { 'Cyan' }
            'Debug'   { 'DarkGray' }
            default   { 'Gray' }
        }
        Write-Host "[$Level] $sanitizedMessage" -ForegroundColor $color
    }
}

function Set-LogLevel {
    param([ValidateSet('Verbose', 'Debug', 'Info', 'Warning', 'Error')]
        [string]$Level)
    $script:LogConfig.MinLevel = $Level
}

function Set-LogComponent {
    param([string]$Component)
    $script:LogConfig.Component = $Component
}

function Set-LogSession {
    param([string]$SessionId)
    $script:LogConfig.SessionId = $SessionId
}

function Set-LogContext {
    <#
    .SYNOPSIS
        Set company and ticket identifiers for log line identification.
    .PARAMETER CompanyName
        Company or client name.
    .PARAMETER TicketNumbers
        Ticket number(s) - array or single string.
    #>
    param(
        [string]$CompanyName = $null,
        [array]$TicketNumbers = @()
    )
    if ($CompanyName -and -not [string]::IsNullOrWhiteSpace($CompanyName)) {
        $script:LogConfig.CompanyName = $CompanyName.Trim()
    }
    if ($TicketNumbers) {
        $script:LogConfig.TicketNumbers = @($TicketNumbers | ForEach-Object { $_.ToString().Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    }
}

function Close-Logger {
    try {
        if ($script:LogConfig._fileLock) {
            $script:LogConfig._fileLock.Dispose()
            $script:LogConfig._fileLock = $null
        }
        $script:LogConfig.Initialized = $false
    } catch { }
}

function Get-LogPath {
    return (Get-CurrentLogFilePath)
}

function Safe-ImportModule {
    <#
    .SYNOPSIS
        Safely imports a PowerShell module with error handling.
    
    .DESCRIPTION
        Imports a module, removing any existing version first to force reload.
        Handles errors gracefully with user-friendly messages.
    
    .PARAMETER ModulePath
        Full path to the module file (.psm1)
    
    .PARAMETER ShowSuccessMessage
        If true, displays a success message (default: false)
    
    .EXAMPLE
        Safe-ImportModule -ModulePath "$PSScriptRoot\Modules\Settings.psm1"
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$ModulePath,
        
        [Parameter(Mandatory=$false)]
        [switch]$ShowSuccessMessage
    )
    
    try {
        # Get the module name from the path
        $moduleName = [System.IO.Path]::GetFileNameWithoutExtension($ModulePath)
        
        # Remove the module if it's already loaded to force reload
        if (Get-Module -Name $moduleName -ErrorAction SilentlyContinue) {
            Remove-Module -Name $moduleName -Force -ErrorAction SilentlyContinue
        }
        
        Import-Module $ModulePath -Global -ErrorAction Stop
        
        if ($ShowSuccessMessage) {
            Write-Host "Successfully imported module: $moduleName" -ForegroundColor Green
        }
    } catch {
        $errorMsg = "Failed to import module: $ModulePath`nError: $($_.Exception.Message)"
        
        # Try to show MessageBox if Windows Forms is available, otherwise use Write-Error
        try {
            Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
            [System.Windows.Forms.MessageBox]::Show($errorMsg, "Module Import Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        } catch {
            Write-Error $errorMsg
        }
        
        exit 1
    }
}

Export-ModuleMember -Function Initialize-Logger, Write-Log, Set-LogLevel, Set-LogComponent, Set-LogSession, Set-LogContext, Close-Logger, Get-LogPath, Get-DefaultLogPath, Safe-ImportModule
