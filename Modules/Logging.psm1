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
        Write a log entry.
    .PARAMETER Message
        Log message
    .PARAMETER Level
        Verbose, Debug, Info, Warning, Error
    .PARAMETER Component
        Override component for this entry
    .PARAMETER Data
        Hashtable of additional structured data
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [ValidateSet('Verbose', 'Debug', 'Info', 'Warning', 'Error')]
        [string]$Level = 'Info',
        [string]$Component = $null,
        [hashtable]$Data = @{}
    )
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
            Write-Host "[$Level] $Message" -ForegroundColor $color
        }
        return
    }
    $minVal = Get-LogLevelValue -Level $script:LogConfig.MinLevel
    $msgVal = Get-LogLevelValue -Level $Level
    if ($msgVal -lt $minVal) { return }
    $comp = if ($Component) { $Component } elseif ($script:LogConfig.Component) { $script:LogConfig.Component } else { 'Main' }
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'
    $sessionId = $script:LogConfig.SessionId
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
                    Message   = $Message
                } + $Data
                $line = $entry | ConvertTo-Json -Compress
            } else {
                $sessionPart = if ($sessionId) { "[$sessionId] " } else { '' }
                $line = "[$timestamp] [$Level] $sessionPart[$comp] $Message"
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
        Write-Host "[$Level] $Message" -ForegroundColor $(if ($Level -eq 'Error') { 'Red' } else { 'Gray' })
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
        Write-Host "[$Level] $Message" -ForegroundColor $color
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

Export-ModuleMember -Function Initialize-Logger, Write-Log, Set-LogLevel, Set-LogComponent, Set-LogSession, Close-Logger, Get-LogPath, Get-DefaultLogPath
