param(
    [Parameter(Mandatory=$false)][string]$ApiKey,
    [Parameter(Mandatory=$false)][string]$OutputFolder,
    [Parameter(Mandatory=$false)][string]$Model = 'claude-3-5-sonnet-20241022',
    [Parameter(Mandatory=$false)][string[]]$ExtraFiles,
    [Parameter(Mandatory=$false)][string]$ResponseFile = 'Claude_Response.md',
    [Parameter(Mandatory=$false)][int]$MaxCsvRows = 0,
    [Parameter(Mandatory=$false)][switch]$VerboseOutput
)

# Constants
$script:MaxFileSizeBytes = 20MB
$script:MaxSafeProcessingSizeBytes = 15MB
$script:MaxContentCharacters = 300000
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
    $OutputFolder = Get-LatestInvestigationFolder -IncludeLegacy -VerboseOutput:$VerboseOutput
}

# Load settings module and get API key
Import-SettingsModule -ScriptRoot $PSScriptRoot -VerboseOutput:$VerboseOutput | Out-Null
if (-not $ApiKey) { 
    $ApiKey = Get-ApiKeyFromSettings -KeyName 'ClaudeApiKey' -VerboseOutput:$VerboseOutput
}
if (-not $ApiKey) { 
    Write-Error 'Claude API key is required. Save it in Settings or pass -ApiKey.'
    exit 1
}

# Validate output folder
if (-not $OutputFolder -or -not (Test-Path $OutputFolder)) { 
    Write-Error 'Could not find Security Investigation output folder. Pass -OutputFolder or generate a report.'
    exit 1
}
if ($VerboseOutput) { 
    Write-Host ("Using output folder: {0}" -f $OutputFolder) -ForegroundColor DarkGray
}

# Collect files to process
$validatedFiles = Get-InvestigationReportFiles -OutputFolder $OutputFolder -ExtraFiles $ExtraFiles
if ($validatedFiles.Count -eq 0) { 
    Write-Error 'No files to send.'
    exit 1
}

if ($VerboseOutput) {
    Write-Host ("Attaching {0} file(s):" -f $validatedFiles.Count) -ForegroundColor DarkGray
    foreach ($filePath in $validatedFiles) { 
        Write-Host (" - {0} ({1} KB)" -f $filePath, [Math]::Round(((Get-Item $filePath).Length/1KB),0)) -ForegroundColor DarkGray
    }
}

# Build Anthropic messages format
$contentParts = @()
$intro = @"
Please analyze the attached security investigation datasets and produce:
- Executive summary (non-technical)
- Timeline of events with timestamps and sources
- Evidence-backed findings
- Minimal immediate actions
Do not speculate beyond provided evidence. Use CSV content explicitly in references.
"@
$contentParts += @{ type = 'text'; text = $intro }

$tempFilesToCleanup = @()
foreach ($filePath in $validatedFiles) {
    # Validate file can be processed
    $validation = Test-FileIsProcessable -FilePath $filePath -MaxSizeBytes $script:MaxFileSizeBytes -MaxSafeProcessingSizeBytes $script:MaxSafeProcessingSizeBytes
    if (-not $validation.IsValid) {
        Write-Warning $validation.Reason
        continue
    }
    
    # Read and process file content
    $fileContent = Read-FileContent -FilePath $filePath -MaxCsvRows $MaxCsvRows -MaxChars $script:MaxContentCharacters
    
    # Track temp files for cleanup
    if ($fileContent.TempFilePath) {
        $tempFilesToCleanup += $fileContent.TempFilePath
    }
    
    # Format content for API
    $fileName = [System.IO.Path]::GetFileName($filePath)
    $contentParts += @{ type='text'; text=("=== {0} ===`n{1}" -f $fileName, $fileContent.Content) }
}

# Clean up temporary files
foreach ($tempFilePath in $tempFilesToCleanup) {
    if (Test-Path $tempFilePath) {
        try { Remove-Item $tempFilePath -Force -ErrorAction SilentlyContinue } catch {}
    }
}

# Build API request
$request = @{ 
    model = $Model
    max_tokens = 2048
    messages = @(@{ role = 'user'; content = $contentParts })
}
$requestJson = $request | ConvertTo-Json -Depth 8

# Save request JSON with redacted API key
$requestJsonPath = Join-Path $OutputFolder 'Claude_Request.json'
$requestJsonSafe = Remove-SensitiveDataFromJson -Json $requestJson
$requestJsonSafe | Out-File -FilePath $requestJsonPath -Encoding utf8
if ($VerboseOutput) { 
    Write-Host ("Saved request JSON (API key redacted): {0}" -f $requestJsonPath) -ForegroundColor DarkGray
}

# Prepare API call
$headers = @{ 
    'x-api-key' = $ApiKey
    'anthropic-version' = '2023-06-01'
    'Content-Type' = 'application/json'
}
$apiUrl = 'https://api.anthropic.com/v1/messages'

# ROBUSTNESS: Invoke API with retry logic and timeout
$apiResult = Invoke-RestMethodWithRetry `
    -Headers $headers `
    -Uri $apiUrl `
    -Body $requestJson `
    -MaxRetries $script:MaxApiRetries `
    -BaseRetryDelaySeconds $script:BaseRetryDelaySeconds `
    -RetryableStatusCodes $script:RetryableHttpStatusCodes `
    -TimeoutSeconds 300 `
    -VerboseOutput:$VerboseOutput

if (-not $apiResult.Success) {
    Write-ApiErrorLog `
        -Error $apiResult.Error `
        -Endpoint $apiUrl `
        -OutputFolder $OutputFolder `
        -ErrorFileName 'Claude_Error.txt'
    Write-Error ("Claude request failed. See {0}" -f (Join-Path $OutputFolder 'Claude_Error.txt'))
    exit 1
}

$apiResponse = $apiResult.Response

# ROBUSTNESS: Parse API response with proper validation
$responseText = $null

# Import robustness helpers if available
$robustnessHelpersPath = Join-Path $commonPath 'ApiRobustnessHelpers.psm1'
if (Test-Path $robustnessHelpersPath) {
    try {
        Import-Module $robustnessHelpersPath -Force -ErrorAction SilentlyContinue
        if (Get-Command Get-ClaudeResponseText -ErrorAction SilentlyContinue) {
            $responseText = Get-ClaudeResponseText -ApiResponse $apiResponse
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
    
    if (-not $apiResponse.PSObject.Properties['content']) {
        Write-Error "Invalid API response structure: 'content' property is missing"
        exit 1
    }
    
    if ($null -eq $apiResponse.content -or $apiResponse.content.Count -eq 0) {
        Write-Error "Invalid API response structure: 'content' array is empty or null"
        exit 1
    }
    
    try {
        $textParts = $apiResponse.content | Where-Object { 
            $_.PSObject.Properties['type'] -and $_.type -eq 'text' -and 
            $_.PSObject.Properties['text'] -and $_.text 
        }
        
        if (-not $textParts -or $textParts.Count -eq 0) {
            Write-Warning "No text content found in API response content array"
            $responseText = $null
        } else {
            $responseText = $textParts[0].text
        }
    } catch {
        Write-Error "Failed to parse API response: $($_.Exception.Message)"
        exit 1
    }
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
Write-Host ("Saved Claude response to: {0}" -f $responseFilePath) -ForegroundColor Green
