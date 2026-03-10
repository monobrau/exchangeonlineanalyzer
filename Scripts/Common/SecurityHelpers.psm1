<#
.SYNOPSIS
    Security helper functions for input validation, sanitization, and safe error handling.
.DESCRIPTION
    Provides functions for:
    - Input validation and sanitization
    - Safe error message extraction
    - Sensitive data removal from text/logs
    - Path validation and traversal prevention
    - Argument escaping for PowerShell commands
#>

# Import ApiHelpers for Remove-SensitiveDataFromText (or move it here)
$script:ApiHelpersPath = Join-Path $PSScriptRoot 'ApiHelpers.psm1'
if (Test-Path $script:ApiHelpersPath) {
    Import-Module $script:ApiHelpersPath -Force -ErrorAction SilentlyContinue
}

function Remove-SensitiveDataFromText {
    <#
    .SYNOPSIS
        Removes sensitive data (passwords, API keys, tokens, emails, paths) from text.
    .PARAMETER Text
        Text to sanitize
    .PARAMETER AdditionalPatterns
        Additional regex patterns to match and redact
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$Text,
        
        [Parameter(Mandatory=$false)]
        [string[]]$AdditionalPatterns = @()
    )
    
    if ([string]::IsNullOrWhiteSpace($Text)) {
        return $Text
    }
    
    $sanitized = $Text
    
    # Redact API keys, tokens, passwords
    $sanitized = $sanitized -replace '(?i)(api[_-]?key|authorization|token|password|pwd|secret|credential|connectionstring)\s*[:=]\s*[\"'']?[^\"''\s]+[\"'']?', '$1: [REDACTED]'
    
    # Redact email addresses
    $sanitized = $sanitized -replace '[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}', '[EMAIL_REDACTED]'
    
    # Redact file paths with usernames
    $sanitized = $sanitized -replace 'C:\\Users\\[^\\]+', 'C:\Users\[REDACTED]'
    $sanitized = $sanitized -replace '/home/[^/]+', '/home/[REDACTED]'
    
    # Redact potential passwords in logs (common patterns)
    $sanitized = $sanitized -replace '(?i)(password|pwd|pass)\s*[:=]\s*[^\s]+', '$1: [REDACTED]'
    
    # Apply additional patterns
    foreach ($pattern in $AdditionalPatterns) {
        $sanitized = $sanitized -replace $pattern, '[REDACTED]'
    }
    
    return $sanitized
}

function Validate-SearchTerms {
    <#
    .SYNOPSIS
        Validates and sanitizes user search terms.
    .PARAMETER SearchTerms
        Array of search terms to validate
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string[]]$SearchTerms
    )
    
    $maxLength = 100
    $maxTerms = 50
    $validated = [System.Collections.ArrayList]::new()
    
    if ($null -eq $SearchTerms) { return @() }
    foreach ($term in $SearchTerms) {
        if ($null -eq $term -or [string]::IsNullOrWhiteSpace($term)) {
            continue
        }
        
        $trimmed = $term.Trim()
        
        # Validate length
        if ($trimmed.Length -gt $maxLength) {
            Write-Warning "Search term exceeds maximum length ($maxLength), truncating: $trimmed"
            $trimmed = $trimmed.Substring(0, $maxLength)
        }
        
        # Remove potentially dangerous characters (command injection prevention)
        $sanitized = $trimmed -replace '[<>"|&;`$\\]', ''
        
        # Validate it's not empty after sanitization
        if (-not [string]::IsNullOrWhiteSpace($sanitized)) {
            [void]$validated.Add($sanitized)
        }
    }
    
    if ($validated.Count -gt $maxTerms) {
        Write-Warning "Too many search terms ($($validated.Count)), limiting to $maxTerms"
        $validated = $validated[0..($maxTerms-1)]
    }
    
    return $validated.ToArray()
}

function Validate-FilePath {
    <#
    .SYNOPSIS
        Validates file paths and prevents path traversal attacks.
    .PARAMETER FilePath
        File path to validate
    .PARAMETER BaseDirectory
        Base directory that the file must be within
    .PARAMETER MustExist
        Whether the file must exist
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$FilePath,
        
        [Parameter(Mandatory=$false)]
        [string]$BaseDirectory = $null,
        
        [Parameter(Mandatory=$false)]
        [switch]$MustExist
    )
    
    if ([string]::IsNullOrWhiteSpace($FilePath)) {
        throw "File path cannot be empty"
    }
    
    # Resolve to absolute path
    try {
        $resolved = [System.IO.Path]::GetFullPath($FilePath)
    } catch {
        throw "Invalid file path: $FilePath"
    }
    
    # Validate against base directory if provided
    if ($BaseDirectory) {
        try {
            $baseResolved = [System.IO.Path]::GetFullPath($BaseDirectory)
            if (-not $resolved.StartsWith($baseResolved, [System.StringComparison]::OrdinalIgnoreCase)) {
                throw "File path outside allowed directory: $FilePath"
            }
        } catch {
            throw "Invalid base directory: $BaseDirectory"
        }
    }
    
    # Check existence if required
    if ($MustExist -and -not (Test-Path $resolved)) {
        throw "File does not exist: $FilePath"
    }
    
    return $resolved
}

function Get-SafeErrorMessage {
    <#
    .SYNOPSIS
        Extracts a safe error message from an exception object.
    .PARAMETER Error
        Error object (usually $_)
    .PARAMETER UserMessage
        Generic message to return if error is not safe to expose
    #>
    param(
        [Parameter(Mandatory=$true)]
        [object]$Error,
        
        [Parameter(Mandatory=$false)]
        [string]$UserMessage = "An error occurred"
    )
    
    if (-not $Error -or -not $Error.Exception) {
        return $UserMessage
    }
    
    $exceptionMessage = $Error.Exception.Message
    
    if ([string]::IsNullOrWhiteSpace($exceptionMessage)) {
        return $UserMessage
    }
    
    # Check if it's a known safe error type
    $safeErrorTypes = @(
        'System.IO.FileNotFoundException',
        'System.IO.DirectoryNotFoundException',
        'System.UnauthorizedAccessException',
        'System.ArgumentException'
    )
    
    $exceptionType = $Error.Exception.GetType().FullName
    
    if ($exceptionType -in $safeErrorTypes) {
        # These are safe to show to users (but still sanitize)
        return Remove-SensitiveDataFromText -Text $exceptionMessage
    }
    
    # For other errors, return generic message
    return $UserMessage
}

function Write-SafeError {
    <#
    .SYNOPSIS
        Writes a safe error message without exposing stack traces or sensitive data.
    .PARAMETER Error
        Error object
    .PARAMETER Context
        Context description for the error
    .PARAMETER LogToFile
        Whether to log full details to secure log file
    #>
    param(
        [Parameter(Mandatory=$true)]
        [object]$Error,
        
        [Parameter(Mandatory=$false)]
        [string]$Context = "Operation",
        
        [Parameter(Mandatory=$false)]
        [switch]$LogToFile
    )
    
    # Get safe error message
    $safeMessage = Get-SafeErrorMessage -Error $Error -UserMessage "$Context failed"
    
    # Write safe error to console
    Write-Error $safeMessage
    
    # Log full details to secure log file if requested
    if ($LogToFile -and (Get-Command Write-Log -ErrorAction SilentlyContinue)) {
        $fullMessage = if ($Error.Exception.Message) {
            Remove-SensitiveDataFromText -Text $Error.Exception.Message
        } else {
            "Unknown error"
        }
        
        Write-Log -Message "$Context failed: $fullMessage" -Level Error -Data @{
            ExceptionType = $Error.Exception.GetType().FullName
            # Don't log stack trace to user-visible logs
        }
    }
}

function Escape-PowerShellArgument {
    <#
    .SYNOPSIS
        Properly escapes PowerShell arguments to prevent command injection.
    .PARAMETER Argument
        Argument to escape
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$Argument
    )
    
    if ([string]::IsNullOrWhiteSpace($Argument)) {
        return "''"
    }
    
    # Escape single quotes and wrap in single quotes
    $escaped = $Argument.Replace("'", "''")
    return "'$escaped'"
}

function Validate-TicketContent {
    <#
    .SYNOPSIS
        Validates and sanitizes ticket content input.
    .PARAMETER TicketContent
        Ticket content to validate
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$TicketContent
    )
    
    $maxLength = 1000000  # 1MB limit
    
    if ([string]::IsNullOrWhiteSpace($TicketContent)) {
        return @{ IsValid = $true; Content = ''; WasTruncated = $false }
    }
    
    $wasTruncated = $false
    $sanitized = $TicketContent
    
    # Check length
    if ($sanitized.Length -gt $maxLength) {
        Write-Warning "Ticket content exceeds maximum length ($maxLength), truncating"
        $sanitized = $sanitized.Substring(0, $maxLength) + "`n...[TRUNCATED]"
        $wasTruncated = $true
    }
    
    # Basic sanitization - remove null bytes
    $sanitized = $sanitized -replace "`0", ''
    
    return @{ 
        IsValid = $true
        Content = $sanitized
        WasTruncated = $wasTruncated
    }
}

function Read-CommandFile {
    <#
    .SYNOPSIS
        Safely reads and validates command files.
    .PARAMETER CommandFilePath
        Path to command file
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$CommandFilePath
    )
    
    $maxCommandLength = 10000  # 10KB limit
    
    if (-not (Test-Path $CommandFilePath)) {
        return $null
    }
    
    # Validate file size
    try {
        $fileInfo = Get-Item $CommandFilePath -ErrorAction Stop
        if ($fileInfo.Length -gt $maxCommandLength) {
            Write-Warning "Command file exceeds maximum size ($maxCommandLength bytes), ignoring"
            return $null
        }
    } catch {
        Write-Warning "Failed to validate command file: $(Get-SafeErrorMessage -Error $_)"
        return $null
    }
    
    # Read and validate command
    try {
        $command = Get-Content $CommandFilePath -Raw -ErrorAction Stop
    } catch {
        Write-Warning "Failed to read command file: $(Get-SafeErrorMessage -Error $_)"
        return $null
    }
    
    if ([string]::IsNullOrWhiteSpace($command)) {
        return $null
    }
    
    $command = $command.Trim()
    
    # SECURITY: Use enhanced command whitelist validation
    if (Get-Command Validate-CommandWhitelist -ErrorAction SilentlyContinue) {
        $validation = Validate-CommandWhitelist -Command $command
        if (-not $validation.IsValid) {
            Write-Warning "Command validation failed: $($validation.Reason)"
            return $null
        }
    } else {
        # Fallback to basic validation
        $allowedCommands = @('GRAPH_AUTH', 'EXCHANGE_AUTH', 'VALIDATE_USERS', 'GENERATE_REPORTS', 'GENERATE_REPORTS_SEARCH', 'TICKET_DATA')
        $commandPrefix = $command -split '\|' | Select-Object -First 1
        
        if ($commandPrefix -notin $allowedCommands) {
            Write-Warning "Unknown command prefix: $commandPrefix, ignoring"
            return $null
        }
    }
    
    return $command
}

function New-SecurePassword {
    <#
    .SYNOPSIS
        Generates a secure random password.
    .PARAMETER Length
        Length of password to generate
    #>
    param(
        [Parameter(Mandatory=$false)]
        [int]$Length = 16
    )
    
    if ($Length -lt 8) {
        $Length = 8
    }
    if ($Length -gt 128) {
        $Length = 128
    }
    
    $chars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$%^&*"
    $password = ""
    
    for ($i = 0; $i -lt $Length; $i++) {
        $password += $chars[(Get-Random -Maximum $chars.Length)]
    }
    
    return $password
}

# Rate limiting storage (in-memory, per-process)
$script:RateLimitStore = @{}

function Test-RateLimit {
    <#
    .SYNOPSIS
        Tests if an operation should be rate-limited.
    .PARAMETER Key
        Unique identifier for the rate limit (e.g., "user-validation", "api-call")
    .PARAMETER MaxRequests
        Maximum number of requests allowed
    .PARAMETER WindowSeconds
        Time window in seconds
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$Key,
        
        [Parameter(Mandatory=$false)]
        [int]$MaxRequests = 10,
        
        [Parameter(Mandatory=$false)]
        [int]$WindowSeconds = 60
    )
    
    $now = Get-Date
    $windowStart = $now.AddSeconds(-$WindowSeconds)
    
    # Initialize or clean old entries
    if (-not $script:RateLimitStore.ContainsKey($Key)) {
        $script:RateLimitStore[$Key] = [System.Collections.ArrayList]::new()
    }
    
    $requests = $script:RateLimitStore[$Key]
    
    # Remove requests outside the time window
    $validRequests = @($requests | Where-Object { $_ -gt $windowStart })
    $script:RateLimitStore[$Key] = if ($validRequests.Count -eq 0) { [System.Collections.ArrayList]::new() } else { [System.Collections.ArrayList]::new($validRequests) }
    
    # Check if limit exceeded
    if ($script:RateLimitStore[$Key].Count -ge $MaxRequests) {
        $oldestRequest = ($script:RateLimitStore[$Key] | Sort-Object | Select-Object -First 1)
        $waitUntil = $oldestRequest.AddSeconds($WindowSeconds)
        $waitSeconds = [Math]::Ceiling(($waitUntil - $now).TotalSeconds)
        return @{
            Allowed = $false
            WaitSeconds = $waitSeconds
            Message = "Rate limit exceeded. Please wait $waitSeconds seconds before trying again."
        }
    }
    
    # Add current request
    [void]$script:RateLimitStore[$Key].Add($now)
    
    return @{
        Allowed = $true
        WaitSeconds = 0
        Message = $null
    }
}

function Clear-RateLimit {
    <#
    .SYNOPSIS
        Clears rate limit data for a specific key or all keys.
    .PARAMETER Key
        Specific key to clear, or omit to clear all
    #>
    param(
        [Parameter(Mandatory=$false)]
        [string]$Key = $null
    )
    
    if ($Key) {
        if ($script:RateLimitStore.ContainsKey($Key)) {
            $script:RateLimitStore.Remove($Key)
        }
    } else {
        $script:RateLimitStore.Clear()
    }
}

function Validate-CommandWhitelist {
    <#
    .SYNOPSIS
        Validates that a command is in the allowed whitelist.
    .PARAMETER Command
        Command string to validate
    .PARAMETER AllowedCommands
        Array of allowed command prefixes (defaults to standard set)
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$Command,
        
        [Parameter(Mandatory=$false)]
        [string[]]$AllowedCommands = @('GRAPH_AUTH', 'EXCHANGE_AUTH', 'VALIDATE_USERS', 'GENERATE_REPORTS', 'GENERATE_REPORTS_SEARCH', 'TICKET_DATA')
    )
    
    if ([string]::IsNullOrWhiteSpace($Command)) {
        return @{ IsValid = $false; Reason = "Command cannot be empty" }
    }
    
    # Extract command prefix (before first | or : or space)
    $commandPrefix = $Command -split '[|:\s]' | Select-Object -First 1
    
    if ([string]::IsNullOrWhiteSpace($commandPrefix)) {
        return @{ IsValid = $false; Reason = "Invalid command format" }
    }
    
    # Check against whitelist (case-insensitive)
    $commandPrefixUpper = $commandPrefix.ToUpper()
    $isAllowed = $false
    
    foreach ($allowed in $AllowedCommands) {
        if ($commandPrefixUpper -eq $allowed.ToUpper()) {
            $isAllowed = $true
            break
        }
    }
    
    if (-not $isAllowed) {
        return @{
            IsValid = $false
            Reason = "Command '$commandPrefix' is not in the allowed whitelist. Allowed commands: $($AllowedCommands -join ', ')"
        }
    }
    
    # Additional validation: check for suspicious patterns
    $suspiciousPatterns = @(
        '\.\.',           # Path traversal
        '[;&|`]',         # Command chaining
        '\$\(',           # Command substitution
        'Invoke-',         # PowerShell invocation
        'Start-Process',   # Process execution
        'Get-Content.*\.\.', # Path traversal in file operations
        'Remove-Item.*\.\.'  # Path traversal in deletion
    )
    
    foreach ($pattern in $suspiciousPatterns) {
        if ($Command -match $pattern) {
            return @{
                IsValid = $false
                Reason = "Command contains suspicious pattern: $pattern"
            }
        }
    }
    
    return @{ IsValid = $true; Reason = $null }
}

function Validate-AllFilePaths {
    <#
    .SYNOPSIS
        Validates multiple file paths at once, ensuring all are within allowed directories.
    .PARAMETER FilePaths
        Array of file paths to validate
    .PARAMETER BaseDirectory
        Base directory that all files must be within
    .PARAMETER MustExist
        Whether files must exist
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string[]]$FilePaths,
        
        [Parameter(Mandatory=$false)]
        [string]$BaseDirectory = $null,
        
        [Parameter(Mandatory=$false)]
        [switch]$MustExist
    )
    
    $results = @()
    
    foreach ($filePath in $FilePaths) {
        try {
            $validated = Validate-FilePath -FilePath $filePath -BaseDirectory $BaseDirectory -MustExist:$MustExist
            $results += @{
                FilePath = $filePath
                IsValid = $true
                ValidatedPath = $validated
                Error = $null
            }
        } catch {
            $results += @{
                FilePath = $filePath
                IsValid = $false
                ValidatedPath = $null
                Error = $_.Exception.Message
            }
        }
    }
    
    return $results
}

Export-ModuleMember -Function Remove-SensitiveDataFromText, Validate-SearchTerms, Validate-FilePath, Get-SafeErrorMessage, Write-SafeError, Escape-PowerShellArgument, Validate-TicketContent, Read-CommandFile, New-SecurePassword, Test-RateLimit, Clear-RateLimit, Validate-CommandWhitelist, Validate-AllFilePaths
