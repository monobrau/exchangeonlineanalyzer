# File Processing Module
# Shared functions for reading, validating, and processing files

function Get-FileMimeType {
    param(
        [Parameter(Mandatory=$true)]
        [string]$FilePath
    )
    
    switch -regex ($FilePath) {
        '\.csv$' { return 'text/csv' }
        '\.txt$' { return 'text/plain' }
        default { return 'application/octet-stream' }
    }
}

function Test-FileIsProcessable {
    param(
        [Parameter(Mandatory=$true)]
        [string]$FilePath,
        
        [Parameter(Mandatory=$true)]
        [long]$MaxSizeBytes,
        
        [Parameter(Mandatory=$false)]
        [long]$MaxSafeProcessingSizeBytes = 15MB
    )
    
    if (-not (Test-Path $FilePath)) {
        return @{ IsValid = $false; Reason = "File not found: $FilePath" }
    }
    
    $fileSize = (Get-Item $FilePath).Length
    
    if ($fileSize -gt $MaxSizeBytes) {
        return @{ IsValid = $false; Reason = "File exceeds $([Math]::Round($MaxSizeBytes/1MB, 2))MB limit: $FilePath" }
    }
    
    # Check safe processing size (accounts for base64 overhead)
    if ($fileSize -gt $MaxSafeProcessingSizeBytes) {
        return @{ IsValid = $false; Reason = "File too large for safe processing ($([Math]::Round($fileSize/1MB, 2))MB): $FilePath" }
    }
    
    # Validate file is text-based
    $mimeType = Get-FileMimeType -FilePath $FilePath
    if ($mimeType -eq 'application/octet-stream' -and $FilePath -notmatch '\.(csv|txt)$') {
        return @{ IsValid = $false; Reason = "Skipping non-text file: $FilePath" }
    }
    
    return @{ IsValid = $true; FileSize = $fileSize; MimeType = $mimeType }
}

function New-TemporaryFileFromCsv {
    param(
        [Parameter(Mandatory=$true)]
        [string]$FilePath,
        
        [Parameter(Mandatory=$true)]
        [int]$MaxRows
    )
    
    $tempFilePath = $null
    try {
        $linesNeeded = $MaxRows + 1 # header + rows
        $tempFilePath = [System.IO.Path]::GetTempFileName()
        Get-Content -Path $FilePath -TotalCount $linesNeeded | Set-Content -Path $tempFilePath -Encoding utf8
        return $tempFilePath
    } catch {
        # Clean up temp file on error
        if ($tempFilePath -and (Test-Path $tempFilePath)) {
            try { Remove-Item $tempFilePath -Force -ErrorAction SilentlyContinue } catch {}
        }
        return $FilePath
    }
}

function Read-FileContent {
    param(
        [Parameter(Mandatory=$true)]
        [string]$FilePath,
        
        [Parameter(Mandatory=$false)]
        [int]$MaxCsvRows = 0,
        
        [Parameter(Mandatory=$false)]
        [int]$MaxChars = 300000
    )
    
    $processedFilePath = $FilePath
    $isTempFile = $false
    
    # Create temp file if CSV truncation needed
    if ($MaxCsvRows -gt 0 -and $FilePath -match '\.csv$') {
        $processedFilePath = New-TemporaryFileFromCsv -FilePath $FilePath -MaxRows $MaxCsvRows
        $isTempFile = ($processedFilePath -ne $FilePath)
    }
    
    $fileContent = try { 
        Get-Content -Path $processedFilePath -Raw -Encoding UTF8 
    } catch { 
        Write-Warning "Failed to read file $FilePath : $($_.Exception.Message)"
        return @{ 
            Content = "(file read error: $($_.Exception.Message))"
            WasTruncated = $false
            OriginalSize = 0
            TempFilePath = $null
        }
    }
    
    if ([string]::IsNullOrWhiteSpace($fileContent)) { 
        Write-Warning "File $FilePath is empty or could not be read"
        $fileContent = "(empty file or read error)"
    }
    
    $originalSize = $fileContent.Length
    $wasTruncated = $false
    
    if ($fileContent.Length -gt $MaxChars) { 
        Write-Warning "File $(Split-Path $FilePath -Leaf) exceeds $MaxChars characters, truncating. Original size: $originalSize chars"
        $fileContent = $fileContent.Substring(0, $MaxChars) + "`n...[TRUNCATED - Original file was $originalSize characters]"
        $wasTruncated = $true
    }
    
    return @{
        Content = $fileContent
        WasTruncated = $wasTruncated
        OriginalSize = $originalSize
        TempFilePath = if ($isTempFile) { $processedFilePath } else { $null }
    }
}

Export-ModuleMember -Function Get-FileMimeType, Test-FileIsProcessable, New-TemporaryFileFromCsv, Read-FileContent
