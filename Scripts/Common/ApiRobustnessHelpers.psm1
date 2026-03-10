<#
.SYNOPSIS
    API robustness helper functions for timeout handling, retry logic, response validation, and error tracking.
.DESCRIPTION
    Provides functions for:
    - Timeout wrappers for long-running operations
    - Retry logic with exponential backoff
    - Response structure validation
    - Failure tracking and reporting
    - Edge case handling
#>

# Import SecurityHelpers for safe error messages
$securityHelpersPath = Join-Path $PSScriptRoot 'SecurityHelpers.psm1'
if (Test-Path $securityHelpersPath) {
    Import-Module $securityHelpersPath -Force -ErrorAction SilentlyContinue
}

function Invoke-WithTimeout {
    <#
    .SYNOPSIS
        Executes a scriptblock with a timeout.
    .PARAMETER ScriptBlock
        Scriptblock to execute
    .PARAMETER TimeoutSeconds
        Maximum time to wait
    .PARAMETER OperationName
        Name of operation for logging
    #>
    param(
        [Parameter(Mandatory=$true)]
        [scriptblock]$ScriptBlock,
        
        [Parameter(Mandatory=$false)]
        [int]$TimeoutSeconds = 300,
        
        [Parameter(Mandatory=$false)]
        [string]$OperationName = "Operation"
    )
    
    $job = Start-Job -ScriptBlock $ScriptBlock
    
    try {
        if (Wait-Job $job -Timeout $TimeoutSeconds) {
            $result = Receive-Job $job
            Remove-Job $job -Force -ErrorAction SilentlyContinue
            return @{ Success = $true; Result = $result; TimedOut = $false }
        } else {
            Stop-Job $job -ErrorAction SilentlyContinue
            Remove-Job $job -Force -ErrorAction SilentlyContinue
            $errorMsg = "$OperationName timed out after $TimeoutSeconds seconds"
            Write-Warning $errorMsg
            if (Get-Command Write-Log -ErrorAction SilentlyContinue) {
                Write-Log -Message $errorMsg -Level Warning -Data @{ Operation = $OperationName; TimeoutSeconds = $TimeoutSeconds }
            }
            return @{ Success = $false; Result = $null; TimedOut = $true; Error = $errorMsg }
        }
    } catch {
        Stop-Job $job -ErrorAction SilentlyContinue
        Remove-Job $job -Force -ErrorAction SilentlyContinue
        $errorMsg = "$OperationName failed: $($_.Exception.Message)"
        Write-Warning $errorMsg
        if (Get-Command Write-Log -ErrorAction SilentlyContinue) {
            Write-Log -Message $errorMsg -Level Error -Data @{ Operation = $OperationName; Exception = $_.Exception.Message }
        }
        return @{ Success = $false; Result = $null; TimedOut = $false; Error = $errorMsg }
    }
}

function Invoke-ExchangeCmdletWithRetry {
    <#
    .SYNOPSIS
        Executes Exchange Online cmdlets with retry logic and timeout.
    .PARAMETER ScriptBlock
        Scriptblock containing Exchange cmdlet call
    .PARAMETER MaxRetries
        Maximum number of retry attempts
    .PARAMETER BaseDelaySeconds
        Base delay between retries (exponential backoff)
    .PARAMETER TimeoutSeconds
        Timeout per attempt
    .PARAMETER OperationName
        Name of operation for logging
    #>
    param(
        [Parameter(Mandatory=$true)]
        [scriptblock]$ScriptBlock,
        
        [Parameter(Mandatory=$false)]
        [int]$MaxRetries = 3,
        
        [Parameter(Mandatory=$false)]
        [int]$BaseDelaySeconds = 2,
        
        [Parameter(Mandatory=$false)]
        [int]$TimeoutSeconds = 300,
        
        [Parameter(Mandatory=$false)]
        [string]$OperationName = "Exchange operation"
    )
    
    $lastError = $null
    
    for ($attempt = 1; $attempt -le $MaxRetries; $attempt++) {
        # Use timeout wrapper
        $timeoutResult = Invoke-WithTimeout -ScriptBlock $ScriptBlock -TimeoutSeconds $TimeoutSeconds -OperationName "$OperationName (attempt $attempt)"
        
        if ($timeoutResult.Success) {
            return @{ Success = $true; Result = $timeoutResult.Result; Attempt = $attempt }
        }
        
        if ($timeoutResult.TimedOut) {
            $lastError = "Operation timed out"
            if ($attempt -lt $MaxRetries) {
                $delay = $BaseDelaySeconds * $attempt
                Write-Warning "$OperationName timed out (attempt $attempt/$MaxRetries). Retrying in $delay seconds..."
                Start-Sleep -Seconds $delay
                continue
            }
            break
        }
        
        # Check for retryable errors
        $errorMsg = if ($timeoutResult.Error) { $timeoutResult.Error } else { "Unknown error" }
        $isRetryable = $errorMsg -like "*429*" -or 
                      $errorMsg -like "*503*" -or 
                      $errorMsg -like "*502*" -or
                      $errorMsg -like "*throttle*" -or
                      $errorMsg -like "*timeout*" -or
                      $errorMsg -like "*connection*" -or
                      $errorMsg -like "*network*"
        
        if ($isRetryable -and $attempt -lt $MaxRetries) {
            $delay = $BaseDelaySeconds * $attempt
            Write-Warning "$OperationName failed (attempt $attempt/$MaxRetries): $errorMsg. Retrying in $delay seconds..."
            Start-Sleep -Seconds $delay
            continue
        }
        
        # Not retryable or max retries reached
        $lastError = $errorMsg
        break
    }
    
    $finalError = if ($lastError) { $lastError } else { "$OperationName failed after $MaxRetries attempts" }
    Write-Error $finalError
    if (Get-Command Write-Log -ErrorAction SilentlyContinue) {
        Write-Log -Message $finalError -Level Error -Data @{ Operation = $OperationName; Attempts = $MaxRetries }
    }
    
    return @{ Success = $false; Result = $null; Attempt = $MaxRetries; Error = $finalError }
}

function Invoke-GraphCmdletWithRetry {
    <#
    .SYNOPSIS
        Executes Microsoft Graph cmdlets with retry logic and timeout.
    .PARAMETER ScriptBlock
        Scriptblock containing Graph cmdlet call
    .PARAMETER MaxRetries
        Maximum number of retry attempts
    .PARAMETER BaseDelaySeconds
        Base delay between retries
    .PARAMETER TimeoutSeconds
        Timeout per attempt
    .PARAMETER OperationName
        Name of operation for logging
    #>
    param(
        [Parameter(Mandatory=$true)]
        [scriptblock]$ScriptBlock,
        
        [Parameter(Mandatory=$false)]
        [int]$MaxRetries = 3,
        
        [Parameter(Mandatory=$false)]
        [int]$BaseDelaySeconds = 2,
        
        [Parameter(Mandatory=$false)]
        [int]$TimeoutSeconds = 300,
        
        [Parameter(Mandatory=$false)]
        [string]$OperationName = "Graph operation"
    )
    
    # Graph cmdlets don't support timeout directly, so we use job wrapper
    return Invoke-ExchangeCmdletWithRetry -ScriptBlock $ScriptBlock -MaxRetries $MaxRetries -BaseDelaySeconds $BaseDelaySeconds -TimeoutSeconds $TimeoutSeconds -OperationName $OperationName
}

function Validate-ApiResponse {
    <#
    .SYNOPSIS
        Validates API response structure before parsing.
    .PARAMETER Response
        API response object to validate
    .PARAMETER RequiredProperties
        Array of required property paths (e.g., "candidates[0].content.parts")
    #>
    param(
        [Parameter(Mandatory=$true)]
        [object]$Response,
        
        [Parameter(Mandatory=$false)]
        [string[]]$RequiredProperties = @()
    )
    
    if ($null -eq $Response) {
        return @{ IsValid = $false; Reason = "Response is null"; MissingProperty = $null }
    }
    
    foreach ($propertyPath in $RequiredProperties) {
        $parts = $propertyPath -split '\.'
        $current = $Response
        $pathSoFar = ""
        
        foreach ($part in $parts) {
            if ([string]::IsNullOrWhiteSpace($pathSoFar)) {
                $pathSoFar = $part
            } else {
                $pathSoFar += ".$part"
            }
            
            # Handle array index
            if ($part -match '^(.+)\[(\d+)\]$') {
                $arrayName = $Matches[1]
                $index = [int]$Matches[2]
                
                if (-not $current.PSObject.Properties[$arrayName]) {
                    return @{ IsValid = $false; Reason = "Property '$pathSoFar' does not exist"; MissingProperty = $pathSoFar }
                }
                
                $array = $current.$arrayName
                if ($null -eq $array) {
                    return @{ IsValid = $false; Reason = "Property '$pathSoFar' is null"; MissingProperty = $pathSoFar }
                }
                
                if (-not ($array -is [System.Collections.IList])) {
                    return @{ IsValid = $false; Reason = "Property '$pathSoFar' is not an array"; MissingProperty = $pathSoFar }
                }
                
                if ($array.Count -le $index) {
                    return @{ IsValid = $false; Reason = "Array '$pathSoFar' has only $($array.Count) elements, index $index not available"; MissingProperty = $pathSoFar }
                }
                
                $current = $array[$index]
            } else {
                # Regular property
                if (-not $current.PSObject.Properties[$part]) {
                    return @{ IsValid = $false; Reason = "Property '$pathSoFar' does not exist"; MissingProperty = $pathSoFar }
                }
                
                $current = $current.$part
                if ($null -eq $current) {
                    return @{ IsValid = $false; Reason = "Property '$pathSoFar' is null"; MissingProperty = $pathSoFar }
                }
            }
        }
    }
    
    return @{ IsValid = $true; Reason = $null; MissingProperty = $null }
}

function Get-GeminiResponseText {
    <#
    .SYNOPSIS
        Safely extracts text from Gemini API response with full validation.
    .PARAMETER ApiResponse
        Gemini API response object
    #>
    param(
        [Parameter(Mandatory=$true)]
        [object]$ApiResponse
    )
    
    # Validate response structure
    $validation = Validate-ApiResponse -Response $ApiResponse -RequiredProperties @(
        "candidates",
        "candidates[0]",
        "candidates[0].content",
        "candidates[0].content.parts"
    )
    
    if (-not $validation.IsValid) {
        Write-Error "Invalid Gemini API response structure: $($validation.Reason)"
        if (Get-Command Write-Log -ErrorAction SilentlyContinue) {
            Write-Log -Message "Gemini response validation failed: $($validation.Reason)" -Level Error -Data @{ MissingProperty = $validation.MissingProperty }
        }
        return $null
    }
    
    if ($ApiResponse.candidates.Count -eq 0) {
        Write-Warning "Gemini API response has empty candidates array"
        return $null
    }
    
    $candidate = $ApiResponse.candidates[0]
    if ($candidate.content.parts.Count -eq 0) {
        Write-Warning "Gemini API response has empty parts array"
        return $null
    }
    
    $textParts = [System.Collections.ArrayList]::new()
    foreach ($part in $candidate.content.parts) {
        if ($part -and $part.PSObject.Properties['text'] -and $part.text) {
            [void]$textParts.Add($part.text)
        }
    }
    
    if ($textParts.Count -eq 0) {
        Write-Warning "No text content found in Gemini API response parts"
        return $null
    }
    
    return ($textParts -join "`n`n")
}

function Get-ClaudeResponseText {
    <#
    .SYNOPSIS
        Safely extracts text from Claude API response with full validation.
    .PARAMETER ApiResponse
        Claude API response object
    #>
    param(
        [Parameter(Mandatory=$true)]
        [object]$ApiResponse
    )
    
    # Validate response structure
    $validation = Validate-ApiResponse -Response $ApiResponse -RequiredProperties @("content")
    
    if (-not $validation.IsValid) {
        Write-Error "Invalid Claude API response structure: $($validation.Reason)"
        if (Get-Command Write-Log -ErrorAction SilentlyContinue) {
            Write-Log -Message "Claude response validation failed: $($validation.Reason)" -Level Error -Data @{ MissingProperty = $validation.MissingProperty }
        }
        return $null
    }
    
    if (-not $ApiResponse.content -or $ApiResponse.content.Count -eq 0) {
        Write-Warning "Claude API response has empty content array"
        return $null
    }
    
    $textParts = $ApiResponse.content | Where-Object { 
        $_.PSObject.Properties['type'] -and $_.type -eq 'text' -and 
        $_.PSObject.Properties['text'] -and $_.text 
    }
    
    if (-not $textParts -or $textParts.Count -eq 0) {
        Write-Warning "No text content found in Claude API response"
        return $null
    }
    
    return $textParts[0].text
}

function Track-ApiOperation {
    <#
    .SYNOPSIS
        Tracks API operation success/failure for reporting.
    .PARAMETER OperationName
        Name of the operation
    .PARAMETER Success
        Whether operation succeeded
    .PARAMETER ItemIdentifier
        Identifier for the item being processed (e.g., UPN)
    .PARAMETER Error
        Error message if failed
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$OperationName,
        
        [Parameter(Mandatory=$true)]
        [bool]$Success,
        
        [Parameter(Mandatory=$false)]
        [string]$ItemIdentifier = $null,
        
        [Parameter(Mandatory=$false)]
        [string]$Error = $null
    )
    
    # Store in script-level tracking hashtable
    if (-not $script:ApiOperationTracking) {
        $script:ApiOperationTracking = @{}
    }
    
    if (-not $script:ApiOperationTracking.ContainsKey($OperationName)) {
        $script:ApiOperationTracking[$OperationName] = @{
            Total = 0
            Successful = 0
            Failed = 0
            FailedItems = [System.Collections.ArrayList]::new()
        }
    }
    
    $tracking = $script:ApiOperationTracking[$OperationName]
    $tracking.Total++
    
    if ($Success) {
        $tracking.Successful++
    } else {
        $tracking.Failed++
        if ($ItemIdentifier) {
            [void]$tracking.FailedItems.Add(@{
                Item = $ItemIdentifier
                Error = $Error
            })
        }
    }
}

function Get-ApiOperationSummary {
    <#
    .SYNOPSIS
        Gets summary of tracked API operations.
    .PARAMETER OperationName
        Specific operation name, or omit for all
    #>
    param(
        [Parameter(Mandatory=$false)]
        [string]$OperationName = $null
    )
    
    if (-not $script:ApiOperationTracking) {
        return @()
    }
    
    $summaries = @()
    
    if ($OperationName) {
        if ($script:ApiOperationTracking.ContainsKey($OperationName)) {
            $tracking = $script:ApiOperationTracking[$OperationName]
            $summaries += [PSCustomObject]@{
                Operation = $OperationName
                Total = $tracking.Total
                Successful = $tracking.Successful
                Failed = $tracking.Failed
                SuccessRate = if ($tracking.Total -gt 0) { [Math]::Round(($tracking.Successful / $tracking.Total) * 100, 2) } else { 0 }
                FailedItems = $tracking.FailedItems.ToArray()
            }
        }
    } else {
        foreach ($opName in $script:ApiOperationTracking.Keys) {
            $tracking = $script:ApiOperationTracking[$opName]
            $summaries += [PSCustomObject]@{
                Operation = $opName
                Total = $tracking.Total
                Successful = $tracking.Successful
                Failed = $tracking.Failed
                SuccessRate = if ($tracking.Total -gt 0) { [Math]::Round(($tracking.Successful / $tracking.Total) * 100, 2) } else { 0 }
                FailedItems = $tracking.FailedItems.ToArray()
            }
        }
    }
    
    return $summaries
}

function Clear-ApiOperationTracking {
    <#
    .SYNOPSIS
        Clears API operation tracking data.
    .PARAMETER OperationName
        Specific operation to clear, or omit for all
    #>
    param(
        [Parameter(Mandatory=$false)]
        [string]$OperationName = $null
    )
    
    if (-not $script:ApiOperationTracking) {
        return
    }
    
    if ($OperationName) {
        if ($script:ApiOperationTracking.ContainsKey($OperationName)) {
            $script:ApiOperationTracking.Remove($OperationName)
        }
    } else {
        $script:ApiOperationTracking.Clear()
    }
}

function Invoke-ApiCallWithLogging {
    <#
    .SYNOPSIS
        Wrapper for API calls that logs failures and tracks operations.
    .PARAMETER ScriptBlock
        Scriptblock containing API call
    .PARAMETER OperationName
        Name of operation for tracking
    .PARAMETER ItemIdentifier
        Identifier for item being processed
    .PARAMETER ErrorAction
        Error action (Stop, Continue, SilentlyContinue)
    .PARAMETER LogErrors
        Whether to log errors (default: true)
    #>
    param(
        [Parameter(Mandatory=$true)]
        [scriptblock]$ScriptBlock,
        
        [Parameter(Mandatory=$true)]
        [string]$OperationName,
        
        [Parameter(Mandatory=$false)]
        [string]$ItemIdentifier = $null,
        
        [Parameter(Mandatory=$false)]
        [ValidateSet('Stop', 'Continue', 'SilentlyContinue')]
        [string]$ErrorAction = 'Continue',
        
        [Parameter(Mandatory=$false)]
        [bool]$LogErrors = $true
    )
    
    try {
        $result = & $ScriptBlock
        Track-ApiOperation -OperationName $OperationName -Success $true -ItemIdentifier $ItemIdentifier
        return $result
    } catch {
        $errorMsg = if (Get-Command Get-SafeErrorMessage -ErrorAction SilentlyContinue) {
            Get-SafeErrorMessage -Error $_ -UserMessage "API operation failed"
        } else {
            $_.Exception.Message
        }
        
        Track-ApiOperation -OperationName $OperationName -Success $false -ItemIdentifier $ItemIdentifier -Error $errorMsg
        
        if ($LogErrors) {
            $logMsg = "$OperationName failed"
            if ($ItemIdentifier) {
                $logMsg += " for $ItemIdentifier"
            }
            $logMsg += ": $errorMsg"
            
            Write-Warning $logMsg
            if (Get-Command Write-Log -ErrorAction SilentlyContinue) {
                Write-Log -Message $logMsg -Level Warning -Data @{
                    Operation = $OperationName
                    ItemIdentifier = $ItemIdentifier
                    ExceptionType = $_.Exception.GetType().FullName
                }
            }
        }
        
        if ($ErrorAction -eq 'Stop') {
            throw
        } elseif ($ErrorAction -eq 'Continue') {
            return $null
        } else {
            # SilentlyContinue - return null but don't log
            return $null
        }
    }
}

# Initialize tracking hashtable
if (-not $script:ApiOperationTracking) {
    $script:ApiOperationTracking = @{}
}

Export-ModuleMember -Function Invoke-WithTimeout, Invoke-ExchangeCmdletWithRetry, Invoke-GraphCmdletWithRetry, Validate-ApiResponse, Get-GeminiResponseText, Get-ClaudeResponseText, Track-ApiOperation, Get-ApiOperationSummary, Clear-ApiOperationTracking, Invoke-ApiCallWithLogging
