# API Helpers Module
# Shared functions for API requests, retries, and error handling

function Invoke-RestMethodWithRetry {
    <#
    .SYNOPSIS
        Invokes REST API with retry logic and timeout support.
    .PARAMETER TimeoutSeconds
        Timeout for the request (default: 300 seconds)
    #>
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$Headers,
        
        [Parameter(Mandatory=$true)]
        [string]$Uri,
        
        [Parameter(Mandatory=$true)]
        [string]$Body,
        
        [Parameter(Mandatory=$false)]
        [int]$MaxRetries = 3,
        
        [Parameter(Mandatory=$false)]
        [int]$BaseRetryDelaySeconds = 1,
        
        [Parameter(Mandatory=$false)]
        [int[]]$RetryableStatusCodes = @(429, 503, 502),
        
        [Parameter(Mandatory=$false)]
        [string]$Method = 'POST',
        
        [Parameter(Mandatory=$false)]
        [int]$TimeoutSeconds = 300,
        
        [Parameter(Mandatory=$false)]
        [switch]$VerboseOutput
    )
    
    $lastError = $null
    
    for ($attempt = 1; $attempt -le $MaxRetries; $attempt++) {
        try {
            if ($VerboseOutput) {
                Write-Verbose "API request attempt $attempt of $MaxRetries to: $Uri (timeout: $TimeoutSeconds seconds)"
            }
            
            # SECURITY/ROBUSTNESS: Use WebRequest with explicit timeout instead of Invoke-RestMethod
            $request = [System.Net.HttpWebRequest]::Create($Uri)
            $request.Method = $Method
            $request.Timeout = $TimeoutSeconds * 1000  # Convert to milliseconds
            $request.ReadWriteTimeout = $TimeoutSeconds * 1000
            
            # Add headers
            foreach ($key in $Headers.Keys) {
                if ($key -eq 'Content-Type') {
                    $request.ContentType = $Headers[$key]
                } elseif ($key -eq 'User-Agent') {
                    $request.UserAgent = $Headers[$key]
                } else {
                    try {
                        $request.Headers.Add($key, $Headers[$key])
                    } catch {
                        # Some headers can't be added this way, try alternative
                        if ($key -eq 'Authorization') {
                            $request.Headers['Authorization'] = $Headers[$key]
                        }
                    }
                }
            }
            
            # Add body for POST/PUT
            if ($Body -and ($Method -eq 'POST' -or $Method -eq 'PUT' -or $Method -eq 'PATCH')) {
                $bodyBytes = [System.Text.Encoding]::UTF8.GetBytes($Body)
                $request.ContentLength = $bodyBytes.Length
                $requestStream = $request.GetRequestStream()
                $requestStream.Write($bodyBytes, 0, $bodyBytes.Length)
                $requestStream.Close()
            }
            
            # Get response
            $response = $request.GetResponse()
            $reader = New-Object System.IO.StreamReader($response.GetResponseStream())
            $responseBody = $reader.ReadToEnd()
            $reader.Close()
            $response.Close()
            
            # Parse JSON if Content-Type indicates JSON
            $parsedResponse = $responseBody
            if ($response.ContentType -like "*json*" -or $responseBody.TrimStart().StartsWith('{') -or $responseBody.TrimStart().StartsWith('[')) {
                try {
                    $parsedResponse = $responseBody | ConvertFrom-Json
                } catch {
                    # If JSON parsing fails, return raw body
                    Write-Warning "Failed to parse JSON response, returning raw body"
                }
            }
            
            return @{ Success = $true; Response = $parsedResponse; Attempt = $attempt }
            
        } catch {
            $lastError = $_
            
            # Check for timeout
            if ($_.Exception -is [System.Net.WebException] -and 
                $_.Exception.Status -eq [System.Net.WebExceptionStatus]::Timeout) {
                Write-Warning "API request timed out after $TimeoutSeconds seconds (attempt $attempt/$MaxRetries)"
                if ($attempt -lt $MaxRetries) {
                    $delay = $BaseRetryDelaySeconds * $attempt
                    Start-Sleep -Seconds $delay
                    continue
                }
                break
            }
            
            if ($attempt -eq $MaxRetries) { break }
            
            $statusCode = $null
            try { 
                if ($_.Exception.Response) {
                    $statusCode = $_.Exception.Response.StatusCode.value__ 
                }
            } catch {}
            
            # Retry on specified status codes
            if ($statusCode -in $RetryableStatusCodes) {
                $delay = $BaseRetryDelaySeconds * $attempt
                Write-Warning "API request failed with status $statusCode, retrying in $delay seconds... (attempt $attempt of $MaxRetries)"
                Start-Sleep -Seconds $delay
                continue
            }
            
            # Don't retry on other errors
            break
        }
    }
    
    return @{ Success = $false; Error = $lastError; Attempt = $MaxRetries }
}

function Write-ApiErrorLog {
    param(
        [Parameter(Mandatory=$true)]
        [object]$Error,
        
        [Parameter(Mandatory=$true)]
        [string]$Endpoint,
        
        [Parameter(Mandatory=$true)]
        [string]$OutputFolder,
        
        [Parameter(Mandatory=$true)]
        [string]$ErrorFileName
    )
    
    $errorLogLines = [System.Collections.ArrayList]::new()
    [void]$errorLogLines.Add("Endpoint: $Endpoint")
    
    try { 
        $statusCode = $Error.Exception.Response.StatusCode
        $statusValue = $Error.Exception.Response.StatusCode.value__
        [void]$errorLogLines.Add("Status: $statusValue ($statusCode)")
    } catch {}
    
    if ($Error.Exception.Message) { 
        [void]$errorLogLines.Add("Exception: $($Error.Exception.Message)")
    }
    
    if ($Error.ErrorDetails -and $Error.ErrorDetails.Message) { 
        [void]$errorLogLines.Add("ErrorDetails: $($Error.ErrorDetails.Message)")
    }
    
    # Try to capture response body across frameworks
    try {
        $errorResponseBody = $null
        if ($Error.Exception.Response -is [System.Net.Http.HttpResponseMessage]) {
            $errorResponseBody = $Error.Exception.Response.Content.ReadAsStringAsync().Result
        } elseif ($Error.Exception.Response -and $Error.Exception.Response.GetResponseStream) {
            $reader = New-Object System.IO.StreamReader($Error.Exception.Response.GetResponseStream())
            $errorResponseBody = $reader.ReadToEnd()
        }
        
        if ($errorResponseBody) {
            $sanitizedBody = Remove-SensitiveDataFromText -Text $errorResponseBody
            [void]$errorLogLines.Add("Body:")
            [void]$errorLogLines.Add($sanitizedBody)
        }
    } catch {}
    
    $errorPath = Join-Path $OutputFolder $ErrorFileName
    ($errorLogLines -join "`r`n") | Out-File -FilePath $errorPath -Encoding utf8
}

function Remove-SensitiveDataFromText {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Text,
        
        [Parameter(Mandatory=$false)]
        [string[]]$AdditionalPatterns = @()
    )
    
    $patterns = @(
        '(?i)(api[_-]?key|authorization|token|password)\s*[:=]\s*["'']?[^"'']+["'']?'
    ) + $AdditionalPatterns
    
    $sanitized = $Text
    foreach ($pattern in $patterns) {
        $sanitized = $sanitized -replace $pattern, '$1: [REDACTED]'
    }
    
    return $sanitized
}

function Remove-SensitiveDataFromJson {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Json
    )
    
    $sanitized = $Json
    
    # Specific patterns for common API key fields
    $sanitized = $sanitized -replace '"x-api-key"\s*:\s*"[^"]*"', '"x-api-key": "[REDACTED]"'
    $sanitized = $sanitized -replace '"api[_-]?key"\s*:\s*"[^"]*"', '"api-key": "[REDACTED]"'
    
    # Generic pattern for any sensitive fields
    $sanitized = $sanitized -replace '(?i)(api[_-]?key|authorization|token|password)\s*[:=]\s*["'']?[^"'']+["'']?', '$1: "[REDACTED]"'
    
    return $sanitized
}

Export-ModuleMember -Function Invoke-RestMethodWithRetry, Write-ApiErrorLog, Remove-SensitiveDataFromText, Remove-SensitiveDataFromJson
