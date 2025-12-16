# test_memberberry_execution.ps1
# Test script for memberberry automatic execution integration

param(
    [string]$MemberberryPath = "C:\git\memberberry",
    [string]$MemberberryExceptionsPath = "C:\git\memberberry\exceptions.json",
    [string]$CompanyName = "Test Company"
)

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $scriptRoot) { $scriptRoot = $PSScriptRoot }
if (-not $scriptRoot) { $scriptRoot = Get-Location }

Write-Host "`n=== Testing Memberberry Automatic Execution Integration ===" -ForegroundColor Cyan
Write-Host ""

# Import Settings module
Write-Host "Test 1: Import Settings Module" -ForegroundColor Yellow
try {
    Import-Module "$scriptRoot\Modules\Settings.psm1" -Force -ErrorAction Stop
    Write-Host "  [OK] Settings module imported successfully" -ForegroundColor Green
} catch {
    Write-Host "  [FAIL] Failed to import Settings module: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
Write-Host ""

# Test 2: Check if New-AIReadme function exists
Write-Host "Test 2: Check New-AIReadme Function" -ForegroundColor Yellow
if (Get-Command New-AIReadme -ErrorAction SilentlyContinue) {
    Write-Host "  [OK] New-AIReadme function found" -ForegroundColor Green
} else {
    Write-Host "  [FAIL] New-AIReadme function not found" -ForegroundColor Red
    exit 1
}
Write-Host ""

# Test 3: Verify memberberry path exists
Write-Host "Test 3: Verify Memberberry Path" -ForegroundColor Yellow
Write-Host "  Memberberry Path: $MemberberryPath" -ForegroundColor Gray
if (-not (Test-Path $MemberberryPath)) {
    Write-Host "  [WARN] Memberberry path does not exist: $MemberberryPath" -ForegroundColor Yellow
    Write-Host "  [INFO] Please update the path parameter if needed" -ForegroundColor Yellow
} else {
    Write-Host "  [OK] Memberberry path exists" -ForegroundColor Green
    
    # Check for memberberry script
    $scriptNames = @('memberberry.ps1', 'run.ps1', 'main.ps1', 'memberberry.py', 'run.py', 'main.py', 'memberberry.bat', 'run.bat')
    $foundScript = $null
    foreach ($scriptName in $scriptNames) {
        $scriptPath = Join-Path $MemberberryPath $scriptName
        if (Test-Path $scriptPath) {
            $foundScript = $scriptPath
            Write-Host "  [OK] Found memberberry script: $scriptName" -ForegroundColor Green
            break
        }
    }
    if (-not $foundScript) {
        Write-Host "  [WARN] No memberberry script found. Looking for: $($scriptNames -join ', ')" -ForegroundColor Yellow
    }
    
    # Check for output file
    $outputFile = Join-Path $MemberberryPath "output\memberberry.md"
    Write-Host "  Expected output file: $outputFile" -ForegroundColor Gray
    if (Test-Path $outputFile) {
        Write-Host "  [OK] Output file exists" -ForegroundColor Green
        $fileInfo = Get-Item $outputFile
        Write-Host "  File size: $($fileInfo.Length) bytes" -ForegroundColor Gray
        Write-Host "  Last modified: $($fileInfo.LastWriteTime)" -ForegroundColor Gray
    } else {
        Write-Host "  [WARN] Output file does not exist yet" -ForegroundColor Yellow
    }
}
Write-Host ""

# Test 4: Test New-AIReadme with memberberry enabled
Write-Host "Test 4: Test New-AIReadme with Memberberry Enabled" -ForegroundColor Yellow
Write-Host "  Company Name: $CompanyName" -ForegroundColor Gray
Write-Host "  Exceptions Path: $MemberberryExceptionsPath" -ForegroundColor Gray
Write-Host ""

# Create test settings object
$testSettings = @{
    MemberberryEnabled = $true
    MemberberryPath = $MemberberryPath
    MemberberryExceptionsPath = $MemberberryExceptionsPath
    CompanyName = $CompanyName
    InvestigatorName = "Test Investigator"
    InvestigatorTitle = "Security Analyst"
    TimeZone = "Eastern Standard Time"
}

try {
    Write-Host "  Running New-AIReadme (this will execute memberberry script if found)..." -ForegroundColor Cyan
    $startTime = Get-Date
    $readmeContent = New-AIReadme -Settings $testSettings
    $endTime = Get-Date
    $duration = ($endTime - $startTime).TotalSeconds
    
    if ($readmeContent -and $readmeContent.Length -gt 0) {
        Write-Host "  [OK] AI readme generated successfully" -ForegroundColor Green
        Write-Host "  Content length: $($readmeContent.Length) characters" -ForegroundColor Gray
        Write-Host "  Generation time: $([math]::Round($duration, 2)) seconds" -ForegroundColor Gray
        
        # Show preview
        Write-Host "`n  Content Preview (first 500 chars):" -ForegroundColor Cyan
        Write-Host "  $($readmeContent.Substring(0, [Math]::Min(500, $readmeContent.Length)))..." -ForegroundColor White
        
        # Check if memberberry content is present
        if ($readmeContent -match "ANALYSIS PRINCIPLES|THREAT CLASSIFICATION|REMEDIATION PROTOCOLS|PROCEDURES") {
            Write-Host "`n  [OK] Memberberry content detected in readme" -ForegroundColor Green
        } else {
            Write-Host "`n  [WARN] Memberberry content markers not found - may be using default instructions" -ForegroundColor Yellow
        }
        
        # Save test output
        $testOutputFile = Join-Path $scriptRoot "test_memberberry_execution_output.txt"
        $readmeContent | Out-File -FilePath $testOutputFile -Encoding utf8
        Write-Host "`n  [INFO] Test output saved to: $testOutputFile" -ForegroundColor Cyan
    } else {
        Write-Host "  [FAIL] AI readme generation returned empty content" -ForegroundColor Red
    }
} catch {
    Write-Host "  [FAIL] Exception generating AI readme: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "  Stack trace: $($_.ScriptStackTrace)" -ForegroundColor Gray
}
Write-Host ""

# Test 5: Test with ticket information
Write-Host "Test 5: Test New-AIReadme with Ticket Information" -ForegroundColor Yellow
$testTicketContent = @"
Ticket #1811523
Summary: Security Alert Investigation
Description: User reported suspicious email activity
Discussion: Investigating potential phishing attempt
"@

try {
    Write-Host "  Running New-AIReadme with ticket data..." -ForegroundColor Cyan
    $readmeWithTicket = New-AIReadme -Settings $testSettings -TicketNumbers @("1811523") -TicketContent $testTicketContent
    
    if ($readmeWithTicket -and $readmeWithTicket.Length -gt 0) {
        Write-Host "  [OK] AI readme with ticket generated successfully" -ForegroundColor Green
        Write-Host "  Content length: $($readmeWithTicket.Length) characters" -ForegroundColor Gray
        
        # Check if ticket information is present
        if ($readmeWithTicket -match "Ticket #1811523" -or $readmeWithTicket -match "ConnectWise Ticket Information") {
            Write-Host "  [OK] Ticket information detected in readme" -ForegroundColor Green
        } else {
            Write-Host "  [WARN] Ticket information not found in readme" -ForegroundColor Yellow
        }
        
        # Save test output
        $testOutputFileWithTicket = Join-Path $scriptRoot "test_memberberry_execution_output_with_ticket.txt"
        $readmeWithTicket | Out-File -FilePath $testOutputFileWithTicket -Encoding utf8
        Write-Host "  [INFO] Test output with ticket saved to: $testOutputFileWithTicket" -ForegroundColor Cyan
    } else {
        Write-Host "  [FAIL] AI readme with ticket generation returned empty content" -ForegroundColor Red
    }
} catch {
    Write-Host "  [FAIL] Exception generating AI readme with ticket: $($_.Exception.Message)" -ForegroundColor Red
}
Write-Host ""

# Test 6: Verify output file was updated
Write-Host "Test 6: Verify Output File Was Updated" -ForegroundColor Yellow
$outputFile = Join-Path $MemberberryPath "output\memberberry.md"
if (Test-Path $outputFile) {
    $fileInfoAfter = Get-Item $outputFile
    Write-Host "  Output file last modified: $($fileInfoAfter.LastWriteTime)" -ForegroundColor Gray
    
    # Compare with start time (allowing 5 second margin)
    $timeDiff = ($fileInfoAfter.LastWriteTime - $startTime).TotalSeconds
    if ($timeDiff -gt -5 -and $timeDiff -lt 60) {
        Write-Host "  [OK] Output file appears to have been updated recently" -ForegroundColor Green
    } else {
        Write-Host "  [INFO] Output file modification time: $($fileInfoAfter.LastWriteTime)" -ForegroundColor Gray
        Write-Host "  [INFO] Test start time: $startTime" -ForegroundColor Gray
    }
} else {
    Write-Host "  [WARN] Output file does not exist" -ForegroundColor Yellow
}
Write-Host ""

Write-Host "=== Test Complete ===" -ForegroundColor Cyan
Write-Host ""

