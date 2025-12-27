# verify_memberberry_integration.ps1
# Comprehensive verification script for memberberry integration
# This script verifies that the integration still works after memberberry changes

param(
    [string]$MemberberryPath = "C:\git\memberberry",
    [string]$MemberberryExceptionsPath = "C:\git\memberberry\exceptions.json",
    [string]$CompanyName = "Test Company"
)

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $scriptRoot) { $scriptRoot = $PSScriptRoot }
if (-not $scriptRoot) { $scriptRoot = Get-Location }

Write-Host "`n=== Memberberry Integration Verification ===" -ForegroundColor Cyan
Write-Host "This script verifies that memberberry integration works correctly" -ForegroundColor Gray
Write-Host ""

# Import Settings module
Write-Host "[1/8] Importing Settings Module..." -ForegroundColor Yellow
try {
    Import-Module "$scriptRoot\Modules\Settings.psm1" -Force -ErrorAction Stop
    Write-Host "  ✓ Settings module imported successfully" -ForegroundColor Green
} catch {
    Write-Host "  ✗ Failed to import Settings module: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# Verify function exists
Write-Host "[2/8] Verifying New-AIReadme Function..." -ForegroundColor Yellow
if (Get-Command New-AIReadme -ErrorAction SilentlyContinue) {
    Write-Host "  ✓ New-AIReadme function found" -ForegroundColor Green
} else {
    Write-Host "  ✗ New-AIReadme function not found" -ForegroundColor Red
    exit 1
}

# Check memberberry directory structure
Write-Host "[3/8] Verifying Memberberry Directory Structure..." -ForegroundColor Yellow
Write-Host "  Memberberry Path: $MemberberryPath" -ForegroundColor Gray

if (-not (Test-Path $MemberberryPath -PathType Container)) {
    Write-Host "  ✗ Memberberry directory does not exist: $MemberberryPath" -ForegroundColor Red
    Write-Host "  Please update the -MemberberryPath parameter" -ForegroundColor Yellow
    exit 1
}
Write-Host "  ✓ Memberberry directory exists" -ForegroundColor Green

# Check for expected script files (in order of preference)
Write-Host "  Checking for memberberry script files..." -ForegroundColor Gray
$scriptNames = @('compile.ps1', 'memberberry.ps1', 'run.ps1', 'main.ps1', 'memberberry.py', 'run.py', 'main.py', 'memberberry.bat', 'run.bat')
$foundScript = $null
foreach ($scriptName in $scriptNames) {
    $scriptPath = Join-Path $MemberberryPath $scriptName
    if (Test-Path $scriptPath) {
        $foundScript = $scriptPath
        Write-Host "  ✓ Found script: $scriptName" -ForegroundColor Green
        break
    }
}

if (-not $foundScript) {
    Write-Host "  ⚠ No memberberry script found. Looking for: $($scriptNames -join ', ')" -ForegroundColor Yellow
    Write-Host "  The integration will continue with existing output file if available." -ForegroundColor Yellow
} else {
    Write-Host "  Script location: $foundScript" -ForegroundColor Gray
}

# Check for output file
Write-Host "  Checking for output file..." -ForegroundColor Gray
$expectedOutputFile = Join-Path $MemberberryPath "output\memberberry.md"
Write-Host "  Expected output file: $expectedOutputFile" -ForegroundColor Gray

if (Test-Path $expectedOutputFile) {
    $fileInfo = Get-Item $expectedOutputFile
    Write-Host "  ✓ Output file exists" -ForegroundColor Green
    Write-Host "    Size: $($fileInfo.Length) bytes" -ForegroundColor Gray
    Write-Host "    Last modified: $($fileInfo.LastWriteTime)" -ForegroundColor Gray
} else {
    Write-Host "  ⚠ Output file not found: $expectedOutputFile" -ForegroundColor Yellow
    Write-Host "    Expected location: <MemberberryPath>\output\memberberry.md" -ForegroundColor Gray
    Write-Host "    If your memberberry script outputs to a different location, the integration may fail." -ForegroundColor Yellow
}

# Check exceptions file
Write-Host "[4/8] Verifying Exceptions File..." -ForegroundColor Yellow
Write-Host "  Exceptions Path: $MemberberryExceptionsPath" -ForegroundColor Gray

if (Test-Path $MemberberryExceptionsPath -PathType Leaf) {
    Write-Host "  ✓ Exceptions file exists" -ForegroundColor Green
    try {
        $exceptionsJson = Get-Content -Path $MemberberryExceptionsPath -Raw -ErrorAction Stop | ConvertFrom-Json
        Write-Host "  ✓ Exceptions JSON is valid" -ForegroundColor Green
        
        # Check for global exceptions
        if ($exceptionsJson._global) {
            Write-Host "  ✓ Global exceptions (_global) found" -ForegroundColor Green
        } else {
            Write-Host "  ⚠ Global exceptions (_global) not found" -ForegroundColor Yellow
        }
        
        # Check for client exceptions
        $clientKeys = $exceptionsJson.PSObject.Properties.Name | Where-Object { $_ -ne '_global' }
        if ($clientKeys.Count -gt 0) {
            Write-Host "  ✓ Found $($clientKeys.Count) client exception(s)" -ForegroundColor Green
            Write-Host "    Client keys: $($clientKeys -join ', ')" -ForegroundColor Gray
        } else {
            Write-Host "  ⚠ No client exceptions found" -ForegroundColor Yellow
        }
    } catch {
        Write-Host "  ✗ Exceptions JSON is invalid: $($_.Exception.Message)" -ForegroundColor Red
    }
} elseif (Test-Path $MemberberryExceptionsPath -PathType Container) {
    Write-Host "  ✗ Exceptions path points to a directory, not a file" -ForegroundColor Red
    Write-Host "    Expected: File path (e.g., C:\git\memberberry\exceptions.json)" -ForegroundColor Yellow
} else {
    Write-Host "  ⚠ Exceptions file not found: $MemberberryExceptionsPath" -ForegroundColor Yellow
    Write-Host "    Integration will work without exceptions, but client-specific rules won't be applied." -ForegroundColor Gray
}

# Test script execution (if script found)
Write-Host "[5/8] Testing Script Execution..." -ForegroundColor Yellow
if ($foundScript) {
    Write-Host "  Attempting to execute: $foundScript" -ForegroundColor Gray
    $scriptExtension = [System.IO.Path]::GetExtension($foundScript).ToLower()
    
    try {
        if ($scriptExtension -eq '.ps1') {
            Write-Host "  Executing PowerShell script..." -ForegroundColor Gray
            $startTime = Get-Date
            & $foundScript 2>&1 | Out-Null
            $endTime = Get-Date
            $duration = ($endTime - $startTime).TotalSeconds
            Write-Host "  ✓ Script executed successfully (took $([math]::Round($duration, 2)) seconds)" -ForegroundColor Green
            
            # Check if output file was updated
            if (Test-Path $expectedOutputFile) {
                $fileInfoAfter = Get-Item $expectedOutputFile
                $timeDiff = ($fileInfoAfter.LastWriteTime - $startTime).TotalSeconds
                if ($timeDiff -gt -5 -and $timeDiff -lt 60) {
                    Write-Host "  ✓ Output file was updated by script" -ForegroundColor Green
                } else {
                    Write-Host "  ⚠ Output file modification time suggests it wasn't updated" -ForegroundColor Yellow
                }
            }
        } elseif ($scriptExtension -eq '.py') {
            $pythonExe = Get-Command python -ErrorAction SilentlyContinue
            if (-not $pythonExe) {
                $pythonExe = Get-Command python3 -ErrorAction SilentlyContinue
            }
            if ($pythonExe) {
                Write-Host "  Executing Python script..." -ForegroundColor Gray
                $process = Start-Process -FilePath $pythonExe.Path -ArgumentList "`"$foundScript`"" -Wait -PassThru -NoNewWindow
                if ($process.ExitCode -eq 0) {
                    Write-Host "  ✓ Script executed successfully" -ForegroundColor Green
                } else {
                    Write-Host "  ⚠ Script exited with code $($process.ExitCode)" -ForegroundColor Yellow
                }
            } else {
                Write-Host "  ⚠ Python not found, skipping execution test" -ForegroundColor Yellow
            }
        } elseif ($scriptExtension -eq '.bat' -or $scriptExtension -eq '.cmd') {
            Write-Host "  Executing batch script..." -ForegroundColor Gray
            $process = Start-Process -FilePath "cmd.exe" -ArgumentList "/c", "`"$foundScript`"" -Wait -PassThru -NoNewWindow
            if ($process.ExitCode -eq 0) {
                Write-Host "  ✓ Script executed successfully" -ForegroundColor Green
            } else {
                Write-Host "  ⚠ Script exited with code $($process.ExitCode)" -ForegroundColor Yellow
            }
        }
    } catch {
        Write-Host "  ⚠ Script execution failed: $($_.Exception.Message)" -ForegroundColor Yellow
        Write-Host "    Integration will continue with existing output file if available." -ForegroundColor Gray
    }
} else {
    Write-Host "  ⚠ No script found, skipping execution test" -ForegroundColor Yellow
    Write-Host "    Integration will use existing output file if available." -ForegroundColor Gray
}

# Test New-AIReadme with memberberry enabled
Write-Host "[6/8] Testing New-AIReadme with Memberberry Enabled..." -ForegroundColor Yellow
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
    Write-Host "  Generating AI readme..." -ForegroundColor Gray
    $readmeContent = New-AIReadme -Settings $testSettings
    
    if ($readmeContent -and $readmeContent.Length -gt 0) {
        Write-Host "  ✓ AI readme generated successfully" -ForegroundColor Green
        Write-Host "    Content length: $($readmeContent.Length) characters" -ForegroundColor Gray
        
        # Check for memberberry content markers
        $hasMemberberryContent = $readmeContent -match "ANALYSIS PRINCIPLES|THREAT CLASSIFICATION|REMEDIATION PROTOCOLS|PROCEDURES"
        if ($hasMemberberryContent) {
            Write-Host "  ✓ Memberberry content detected in readme" -ForegroundColor Green
        } else {
            Write-Host "  ⚠ Memberberry content markers not found" -ForegroundColor Yellow
            Write-Host "    This may indicate the output file format has changed." -ForegroundColor Yellow
        }
        
        # Check for default template (should NOT be present if memberberry is working)
        $hasDefaultTemplate = $readmeContent -match "Master Prompt - Generic Template"
        if ($hasDefaultTemplate) {
            Write-Host "  ⚠ Default template still present (memberberry should replace it)" -ForegroundColor Yellow
        } else {
            Write-Host "  ✓ Default template replaced by memberberry content" -ForegroundColor Green
        }
        
        # Save test output
        $testOutputFile = Join-Path $scriptRoot "verify_memberberry_output.txt"
        $readmeContent | Out-File -FilePath $testOutputFile -Encoding utf8
        Write-Host "  Test output saved to: $testOutputFile" -ForegroundColor Cyan
    } else {
        Write-Host "  ✗ AI readme generation returned empty content" -ForegroundColor Red
    }
} catch {
    Write-Host "  ✗ Exception generating AI readme: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "    Stack trace: $($_.ScriptStackTrace)" -ForegroundColor Gray
}

# Test with ticket information
Write-Host "[7/8] Testing with Ticket Information..." -ForegroundColor Yellow
$testTicketContent = @"
Ticket #1811523
Summary: Security Alert Investigation
Description: User reported suspicious email activity
Discussion: Investigating potential phishing attempt
"@

try {
    $readmeWithTicket = New-AIReadme -Settings $testSettings -TicketNumbers @("1811523") -TicketContent $testTicketContent
    
    if ($readmeWithTicket -and $readmeWithTicket.Length -gt 0) {
        Write-Host "  ✓ AI readme with ticket generated successfully" -ForegroundColor Green
        
        # Check if ticket information is present
        if ($readmeWithTicket -match "Ticket.*1811523" -or $readmeWithTicket -match "ConnectWise Ticket Information") {
            Write-Host "  ✓ Ticket information detected in readme" -ForegroundColor Green
        } else {
            Write-Host "  ⚠ Ticket information not found in readme" -ForegroundColor Yellow
        }
        
        $testOutputFileWithTicket = Join-Path $scriptRoot "verify_memberberry_output_with_ticket.txt"
        $readmeWithTicket | Out-File -FilePath $testOutputFileWithTicket -Encoding utf8
        Write-Host "  Test output with ticket saved to: $testOutputFileWithTicket" -ForegroundColor Cyan
    }
} catch {
    Write-Host "  ⚠ Exception generating AI readme with ticket: $($_.Exception.Message)" -ForegroundColor Yellow
}

# Summary
Write-Host "[8/8] Integration Summary..." -ForegroundColor Yellow
Write-Host ""
Write-Host "=== Verification Complete ===" -ForegroundColor Cyan
Write-Host ""
Write-Host "Key Integration Points:" -ForegroundColor Yellow
Write-Host "  1. Script Location: $MemberberryPath" -ForegroundColor Gray
Write-Host "     Expected script names: $($scriptNames -join ', ')" -ForegroundColor Gray
Write-Host "  2. Output File: $expectedOutputFile" -ForegroundColor Gray
Write-Host "     The integration reads from: <MemberberryPath>\output\memberberry.md" -ForegroundColor Gray
Write-Host "  3. Exceptions File: $MemberberryExceptionsPath" -ForegroundColor Gray
Write-Host "     Format: JSON with '_global' key for global exceptions" -ForegroundColor Gray
Write-Host ""
Write-Host "If you've changed memberberry:" -ForegroundColor Yellow
Write-Host "  • Ensure your script outputs to: output\memberberry.md" -ForegroundColor Gray
Write-Host "  • Or update the code in Modules\Settings.psm1 line 482" -ForegroundColor Gray
Write-Host "  • Ensure your script name matches one of the expected names" -ForegroundColor Gray
Write-Host "  • Or update the scriptNames array in Modules\Settings.psm1 line 423" -ForegroundColor Gray
Write-Host ""
Write-Host "Test output files created in: $scriptRoot" -ForegroundColor Cyan
Get-ChildItem -Path $scriptRoot -Filter "verify_memberberry*.txt" -ErrorAction SilentlyContinue | 
    Select-Object -ExpandProperty Name | 
    ForEach-Object { Write-Host "  - $_" -ForegroundColor Green }
Write-Host ""

