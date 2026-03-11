param(
    [Parameter(Mandatory=$false)][string]$ApiKey,
    [Parameter(Mandatory=$false)][string]$OutputFolder,
    [Parameter(Mandatory=$false)][string]$Model = 'models/gemini-2.5-pro',
    [Parameter(Mandatory=$false)][string[]]$ExtraFiles,
    [Parameter(Mandatory=$false)][string]$ResponseFile = 'Gemini_Response.md',
    [Parameter(Mandatory=$false)][switch]$DebugOutput
)

# Constants
$script:MaxFileSizeBytes = 20MB
$script:MaxSafeProcessingSizeBytes = 15MB
$script:MaxApiRetries = 3
$script:BaseRetryDelaySeconds = 1
$script:RetryableHttpStatusCodes = @(429, 503, 502)

# Import shared modules
$commonPath = Join-Path $PSScriptRoot 'Common'
Import-Module (Join-Path $commonPath 'InvestigationHelpers.psm1') -Force -ErrorAction Stop
Import-Module (Join-Path $commonPath 'FileCollection.psm1') -Force -ErrorAction Stop
Import-Module (Join-Path $commonPath 'SettingsHelpers.psm1') -Force -ErrorAction Stop
Import-Module (Join-Path $commonPath 'FileProcessing.psm1') -Force -ErrorAction Stop
Import-Module (Join-Path $commonPath 'ApiHelpers.psm1') -Force -ErrorAction Stop

# Find output folder if not provided
if (-not $OutputFolder -or -not (Test-Path $OutputFolder)) { 
    $OutputFolder = Get-LatestInvestigationFolder -VerboseOutput
}

# Load settings module and get API key
Import-SettingsModule -ScriptRoot $PSScriptRoot -VerboseOutput | Out-Null
if (-not $ApiKey) { 
    $ApiKey = Get-ApiKeyFromSettings -KeyName 'GeminiApiKey' -VerboseOutput
}
if (-not $ApiKey) { 
    Write-Error "Gemini API key is required. Provide -ApiKey or save it in Settings."
    exit 1
}

# Validate output folder
if (-not $OutputFolder -or -not (Test-Path $OutputFolder)) { 
    Write-Error "Could not find Security Investigation output folder. Run the report first or pass -OutputFolder."
    exit 1
}
Write-Verbose ("Using output folder: {0}" -f $OutputFolder)

# Prepare Gemini API endpoints
$modelName = $Model
if ($modelName -like 'models/*') { 
    $modelName = $modelName.Substring(7)
}

# SECURITY: Use Authorization header instead of URL query parameter
$baseEndpoints = @(
    "https://generativelanguage.googleapis.com/v1/models/{0}:generateContent" -f $modelName,
    "https://generativelanguage.googleapis.com/v1beta/models/{0}:generateContent" -f $modelName,
    "https://generativelanguage.googleapis.com/v1beta/{0}:generateContent" -f $Model
)
$headers = @{ 
    'Authorization' = "Bearer $ApiKey"
    'Content-Type' = 'application/json'
}
Write-Verbose ("Model endpoints (fallback order):`n - {0}`n - {1}`n - {2}" -f $baseEndpoints[0], $baseEndpoints[1], $baseEndpoints[2])

# Collect files to process
$validatedFiles = Get-InvestigationReportFiles -OutputFolder $OutputFolder -ExtraFiles $ExtraFiles
if ($validatedFiles.Count -eq 0) { 
    Write-Error "No known report files found in $OutputFolder."
    exit 1
}

Write-Verbose ("Attaching {0} file(s):" -f $validatedFiles.Count)
foreach ($filePath in $validatedFiles) { 
    Write-Verbose (" - {0} ({1} KB)" -f $filePath, [Math]::Round(((Get-Item $filePath).Length/1KB),0))
}

# Build request parts
$intro = @"
Please analyze the attached security investigation datasets and produce:
- Executive summary (non-technical)
- Timeline of events with timestamps and sources
- Evidence-backed findings
- Minimal immediate actions
Do not speculate beyond provided evidence. Use CSV content explicitly in references.
"@

$requestParts = @()
$requestParts += @{ text = $intro }

foreach ($filePath in $validatedFiles) {
    # Validate file can be processed
    $validation = Test-FileIsProcessable -FilePath $filePath -MaxSizeBytes $script:MaxFileSizeBytes -MaxSafeProcessingSizeBytes $script:MaxSafeProcessingSizeBytes
    if (-not $validation.IsValid) {
        Write-Warning $validation.Reason
        continue
    }
    
    # Read file and convert to base64 for Gemini API
    try {
        $fileBytes = [System.IO.File]::ReadAllBytes($filePath)
        $base64Data = [Convert]::ToBase64String($fileBytes)
        $mimeType = Get-FileMimeType -FilePath $filePath
        $requestParts += @{ inlineData = @{ mimeType = $mimeType; data = $base64Data } }
    } catch {
        Write-Warning "Failed to read file $filePath : $($_.Exception.Message)"
        # Continue with other files rather than failing completely
    }
}

# Build API request
$requestBody = @{ contents = @(@{ role = 'user'; parts = $requestParts }) }
$requestJson = $requestBody | ConvertTo-Json -Depth 6

# Save request JSON with redacted API key
$requestJsonPath = Join-Path $OutputFolder 'Gemini_Request.json'
$requestJsonSafe = Remove-SensitiveDataFromJson -Json $requestJson
$requestJsonSafe | Out-File -FilePath $requestJsonPath -Encoding utf8
Write-Verbose ("Saved request JSON (API key redacted): {0}" -f $requestJsonPath)

# Try each endpoint with retry logic
$apiResponse = $null
$lastError = $null
$usedEndpoint = $null

foreach ($endpoint in $baseEndpoints) {
    # ROBUSTNESS: Add timeout to API calls
    $apiResult = Invoke-RestMethodWithRetry `
        -Headers $headers `
        -Uri $endpoint `
        -Body $requestJson `
        -MaxRetries $script:MaxApiRetries `
        -BaseRetryDelaySeconds $script:BaseRetryDelaySeconds `
        -RetryableStatusCodes $script:RetryableHttpStatusCodes `
        -TimeoutSeconds 300 `
        -VerboseOutput
    
    if ($apiResult.Success) {
        $apiResponse = $apiResult.Response
        $usedEndpoint = $endpoint
        break
    }
    
    # Check if 404 - try next endpoint
    $statusCode = $null
    try { 
        $statusCode = $apiResult.Error.Exception.Response.StatusCode.value__
    } catch {}
    
    if ($statusCode -eq 404) {
        # Try next endpoint
        $lastError = $apiResult.Error
        continue
    }
    
    # For other errors, stop trying endpoints
    $lastError = $apiResult.Error
    break
}

if (-not $apiResponse) {
    Write-ApiErrorLog `
        -Error $lastError `
        -Endpoint $usedEndpoint `
        -OutputFolder $OutputFolder `
        -ErrorFileName 'Gemini_Error.txt'
    Write-Error ("Gemini request failed. See {0}" -f (Join-Path $OutputFolder 'Gemini_Error.txt'))
    exit 1
}

# ROBUSTNESS: Parse API response with proper validation
$responseText = $null

# Import robustness helpers if available
$robustnessHelpersPath = Join-Path $commonPath 'ApiRobustnessHelpers.psm1'
if (Test-Path $robustnessHelpersPath) {
    try {
        Import-Module $robustnessHelpersPath -Force -ErrorAction SilentlyContinue
        if (Get-Command Get-GeminiResponseText -ErrorAction SilentlyContinue) {
            $responseText = Get-GeminiResponseText -ApiResponse $apiResponse
        }
    } catch {
        Write-Warning "Failed to import ApiRobustnessHelpers, using fallback validation: $($_.Exception.Message)"
    }
}

# Fallback validation if helper not available
if (-not $responseText) {
    if (-not $apiResponse) {
        Write-Error "API response is null"
        exit 1
    }
    
    if (-not $apiResponse.PSObject.Properties['candidates']) {
        Write-Error "Invalid API response structure: 'candidates' property is missing"
        exit 1
    }
    
    if ($null -eq $apiResponse.candidates -or $apiResponse.candidates.Count -eq 0) {
        Write-Error "Invalid API response structure: 'candidates' array is empty or null"
        exit 1
    }
    
    try {
        $candidate = $apiResponse.candidates[0]
        if (-not $candidate.PSObject.Properties['content']) {
            Write-Error "Invalid API response structure: candidate missing 'content' property"
            exit 1
        }
        
        if (-not $candidate.content.PSObject.Properties['parts']) {
            Write-Error "Invalid API response structure: content missing 'parts' property"
            exit 1
        }
        
        if ($candidate.content.parts.Count -eq 0) {
            Write-Warning "API response has empty parts array"
            $responseText = $null
        } else {
            $textParts = @()
            foreach ($part in $candidate.content.parts) {
                if ($part -and $part.PSObject.Properties['text'] -and $part.text) {
                    $textParts += $part.text
                }
            }
            $responseText = if ($textParts.Count -gt 0) { ($textParts -join "`n`n") } else { $null }
        }
    } catch {
        Write-Error "Failed to parse API response: $($_.Exception.Message)"
        exit 1
    }
}

if ($DebugOutput) { 
    $rawResponsePath = Join-Path $OutputFolder 'Gemini_Response.raw.json'
    ($apiResponse | ConvertTo-Json -Depth 8) | Out-File -FilePath $rawResponsePath -Encoding utf8
    Write-Verbose ("Saved raw response JSON: {0}" -f $rawResponsePath)
}

if (-not $responseText) { 
    Write-Warning "No text content found in API response, saving raw JSON"
    $responseText = ($apiResponse | ConvertTo-Json -Depth 8)
    if ([string]::IsNullOrWhiteSpace($responseText)) {
        Write-Error "API response contains no usable content"
        exit 1
    }
}

# Save response
$responseFilePath = Join-Path $OutputFolder $ResponseFile
$responseText | Out-File -FilePath $responseFilePath -Encoding utf8
Write-Host ("Saved Gemini response to: {0}" -f $responseFilePath) -ForegroundColor Green
