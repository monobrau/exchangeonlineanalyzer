# test_memberberry_integration.ps1
# Test script for memberberry integration

param(
    [string]$MemberberryPath = "C:\git\memberberry\memberberry-complete-output.txt",
    [string]$MemberberryExceptionsPath = "C:\git\memberberry\exceptions.json",
    [string]$CompanyName = "Test Company"
)

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $scriptRoot) { $scriptRoot = $PSScriptRoot }
if (-not $scriptRoot) { $scriptRoot = Get-Location }

Write-Host "`n=== Testing Memberberry Integration ===" -ForegroundColor Cyan
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

# Test 2: Check if Get-MemberberryContent function exists
Write-Host "Test 2: Check Get-MemberberryContent Function" -ForegroundColor Yellow
if (Get-Command Get-MemberberryContent -ErrorAction SilentlyContinue) {
    Write-Host "  [OK] Get-MemberberryContent function found" -ForegroundColor Green
} else {
    Write-Host "  [FAIL] Get-MemberberryContent function not found" -ForegroundColor Red
    exit 1
}
Write-Host ""

# Test 3: Test Get-MemberberryContent with valid file
Write-Host "Test 3: Test Get-MemberberryContent Function" -ForegroundColor Yellow
Write-Host "  Memberberry Path: $MemberberryPath" -ForegroundColor Gray
Write-Host "  Exceptions Path: $MemberberryExceptionsPath" -ForegroundColor Gray
Write-Host "  Company Name: $CompanyName" -ForegroundColor Gray
Write-Host ""

if (-not (Test-Path $MemberberryPath)) {
    Write-Host "  [WARN] Memberberry file not found: $MemberberryPath" -ForegroundColor Yellow
    Write-Host "  Creating a test memberberry file..." -ForegroundColor Yellow
    
    # Create a minimal test memberberry file
    $testMemberberryContent = @"
# ANALYSIS PRINCIPLES

This is a test analysis principles section.

# THREAT CLASSIFICATION RULES

This is a test threat classification section.

# PROCEDURES

## Procedure 1: Test Procedure
This is a test procedure for memberberry integration testing.

## Procedure 2: Another Procedure
More test content here.

# CLIENT EXCEPTIONS

This section should be ignored in favor of JSON file.
"@
    
    $testMemberberryPath = Join-Path $scriptRoot "test_memberberry_output.txt"
    $testMemberberryContent | Out-File -FilePath $testMemberberryPath -Encoding utf8
    $MemberberryPath = $testMemberberryPath
    Write-Host "  Created test file: $testMemberberryPath" -ForegroundColor Green
}

try {
    $memberberryResult = Get-MemberberryContent -MemberberryPath $MemberberryPath -MemberberryExceptionsPath $MemberberryExceptionsPath -CompanyName $CompanyName
    
    if ($memberberryResult.Success) {
        Write-Host "  [OK] Memberberry content loaded successfully" -ForegroundColor Green
        Write-Host "  Global Instructions Length: $($memberberryResult.GlobalInstructions.Length) chars" -ForegroundColor Gray
        Write-Host "  Procedures Length: $($memberberryResult.Procedures.Length) chars" -ForegroundColor Gray
        Write-Host "  Client Exceptions Length: $($memberberryResult.ClientExceptions.Length) chars" -ForegroundColor Gray
        
        # Show previews
        if ($memberberryResult.GlobalInstructions.Length -gt 0) {
            Write-Host "`n  Global Instructions Preview (first 300 chars):" -ForegroundColor Cyan
            Write-Host "  $($memberberryResult.GlobalInstructions.Substring(0, [Math]::Min(300, $memberberryResult.GlobalInstructions.Length)))..." -ForegroundColor White
        }
        
        if ($memberberryResult.Procedures.Length -gt 0) {
            Write-Host "`n  Procedures Preview (first 200 chars):" -ForegroundColor Cyan
            Write-Host "  $($memberberryResult.Procedures.Substring(0, [Math]::Min(200, $memberberryResult.Procedures.Length)))..." -ForegroundColor White
        }
        
        if ($memberberryResult.ClientExceptions.Length -gt 0) {
            Write-Host "`n  Client Exceptions Preview:" -ForegroundColor Cyan
            Write-Host "  $($memberberryResult.ClientExceptions.Substring(0, [Math]::Min(200, $memberberryResult.ClientExceptions.Length)))..." -ForegroundColor White
        } else {
            Write-Host "`n  [WARN] No client exceptions found for '$CompanyName'" -ForegroundColor Yellow
        }
    } else {
        Write-Host "  [FAIL] Failed to load memberberry content: $($memberberryResult.ErrorMessage)" -ForegroundColor Red
    }
} catch {
    Write-Host "  [FAIL] Exception loading memberberry content: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "  Stack trace: $($_.ScriptStackTrace)" -ForegroundColor Red
}
Write-Host ""

# Test 4: Test New-AIReadme with memberberry enabled
Write-Host "Test 4: Test New-AIReadme with Memberberry Enabled" -ForegroundColor Yellow
try {
    # Create test settings with memberberry enabled
    $testSettings = [PSCustomObject]@{
        CompanyName = $CompanyName
        InvestigatorName = "Test Investigator"
        InvestigatorTitle = "Security Engineer"
        MemberberryEnabled = $true
        MemberberryPath = $MemberberryPath
        MemberberryExceptionsPath = $MemberberryExceptionsPath
        AdminUsernames = "admin, service_account"
        InternalTeamDisplayNames = "Managed Services"
        AuthorizedISPs = "Comcast, Verizon"
        TimeZone = "CST"
    }
    
    Write-Host "  Testing New-AIReadme with memberberry enabled..." -ForegroundColor Gray
    $aiReadme = New-AIReadme -Settings $testSettings
    
    if ($aiReadme) {
        Write-Host "  [OK] AI Readme generated successfully" -ForegroundColor Green
        Write-Host "  Length: $($aiReadme.Length) characters" -ForegroundColor Gray
        
        # Check if memberberry content is present
        $hasMemberberryContent = $aiReadme -match "ANALYSIS PRINCIPLES|THREAT CLASSIFICATION|PROCEDURES" -or 
                                 $aiReadme -match "Procedure 1|Procedure 2" -or
                                 ($aiReadme.Length -gt 1000)
        
        if ($hasMemberberryContent) {
            Write-Host "  [OK] Memberberry content appears to be included" -ForegroundColor Green
        } else {
            Write-Host "  [WARN] Memberberry content may not be included (check preview)" -ForegroundColor Yellow
        }
        
        # Check if default template is NOT present (memberberry should replace it)
        $hasDefaultTemplate = $aiReadme -match "Master Prompt - Generic Template" -and 
                             $aiReadme -match "II\. Classification Logic" -and
                             $aiReadme -match "A\. Authorized Activity"
        
        if ($hasDefaultTemplate) {
            Write-Host "  [WARN] Default template still present (memberberry should replace it)" -ForegroundColor Yellow
        } else {
            Write-Host "  [OK] Default template replaced by memberberry content" -ForegroundColor Green
        }
        
        Write-Host "`n  AI Readme Preview (first 500 chars):" -ForegroundColor Cyan
        Write-Host "  $($aiReadme.Substring(0, [Math]::Min(500, $aiReadme.Length)))..." -ForegroundColor White
        
        # Save to test file
        $outputFile = Join-Path $scriptRoot "test_memberberry_AI_Readme.txt"
        $aiReadme | Out-File -FilePath $outputFile -Encoding utf8
        Write-Host "`n  AI Readme saved to: $outputFile" -ForegroundColor Green
    } else {
        Write-Host "  [FAIL] Failed to generate AI Readme" -ForegroundColor Red
    }
} catch {
    Write-Host "  [FAIL] Exception generating AI Readme: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "  Stack trace: $($_.ScriptStackTrace)" -ForegroundColor Red
}
Write-Host ""

# Test 5: Test New-AIReadme with memberberry disabled (should use default)
Write-Host "Test 5: Test New-AIReadme with Memberberry Disabled" -ForegroundColor Yellow
try {
    $testSettingsDisabled = [PSCustomObject]@{
        CompanyName = $CompanyName
        InvestigatorName = "Test Investigator"
        InvestigatorTitle = "Security Engineer"
        MemberberryEnabled = $false
        AdminUsernames = "admin, service_account"
        InternalTeamDisplayNames = "Managed Services"
        AuthorizedISPs = "Comcast, Verizon"
        TimeZone = "CST"
    }
    
    Write-Host "  Testing New-AIReadme with memberberry disabled..." -ForegroundColor Gray
    $aiReadmeDefault = New-AIReadme -Settings $testSettingsDisabled
    
    if ($aiReadmeDefault) {
        Write-Host "  [OK] Default AI Readme generated successfully" -ForegroundColor Green
        Write-Host "  Length: $($aiReadmeDefault.Length) characters" -ForegroundColor Gray
        
        # Check if default template is present
        $hasDefaultTemplate = $aiReadmeDefault -match "Master Prompt - Generic Template" -and 
                             $aiReadmeDefault -match "II\. Classification Logic"
        
        if ($hasDefaultTemplate) {
            Write-Host "  [OK] Default template is present" -ForegroundColor Green
        } else {
            Write-Host "  [WARN] Default template may not be present" -ForegroundColor Yellow
        }
    }
} catch {
    Write-Host "  [FAIL] Exception generating default AI Readme: $($_.Exception.Message)" -ForegroundColor Red
}
Write-Host ""

# Test 6: Test with ticket data
Write-Host "Test 6: Test Memberberry Integration with Ticket Data" -ForegroundColor Yellow
try {
    $testTicketContent = "Service Ticket #12345 - Test alert for user@example.com"
    $testTicketNumbers = @("12345")
    
    Write-Host "  Testing with ticket data..." -ForegroundColor Gray
    $aiReadmeWithTicket = New-AIReadme -Settings $testSettings -TicketNumbers $testTicketNumbers -TicketContent $testTicketContent
    
    if ($aiReadmeWithTicket) {
        Write-Host "  [OK] AI Readme with ticket data generated successfully" -ForegroundColor Green
        
        # Check if ticket section is present
        $hasTicketSection = $aiReadmeWithTicket -match "## ConnectWise Ticket Information" -or
                           $aiReadmeWithTicket -match "Ticket.*12345"
        
        if ($hasTicketSection) {
            Write-Host "  [OK] Ticket information section found" -ForegroundColor Green
        } else {
            Write-Host "  [WARN] Ticket information section not found" -ForegroundColor Yellow
        }
        
        # Check if ticket number is in subject line
        $hasTicketInSubject = $aiReadmeWithTicket -match "Subject:.*Ticket.*12345"
        if ($hasTicketInSubject) {
            Write-Host "  [OK] Ticket number found in subject line" -ForegroundColor Green
        } else {
            Write-Host "  [WARN] Ticket number not found in subject line" -ForegroundColor Yellow
        }
        
        $outputFileTicket = Join-Path $scriptRoot "test_memberberry_AI_Readme_with_ticket.txt"
        $aiReadmeWithTicket | Out-File -FilePath $outputFileTicket -Encoding utf8
        Write-Host "  AI Readme with ticket saved to: $outputFileTicket" -ForegroundColor Green
    }
} catch {
    Write-Host "  [FAIL] Exception generating AI Readme with ticket: $($_.Exception.Message)" -ForegroundColor Red
}
Write-Host ""

# Test 7: Test with missing memberberry file (should fallback to default)
Write-Host "Test 7: Test Fallback to Default When Memberberry File Missing" -ForegroundColor Yellow
try {
    $testSettingsMissing = [PSCustomObject]@{
        CompanyName = $CompanyName
        InvestigatorName = "Test Investigator"
        MemberberryEnabled = $true
        MemberberryPath = "C:\nonexistent\file.txt"
        AdminUsernames = "admin"
    }
    
    Write-Host "  Testing with missing memberberry file..." -ForegroundColor Gray
    $aiReadmeFallback = New-AIReadme -Settings $testSettingsMissing
    
    if ($aiReadmeFallback) {
        Write-Host "  [OK] Fallback AI Readme generated" -ForegroundColor Green
        
        # Should use default template
        $hasDefaultTemplate = $aiReadmeFallback -match "Master Prompt - Generic Template"
        if ($hasDefaultTemplate) {
            Write-Host "  [OK] Default template used as fallback" -ForegroundColor Green
        } else {
            Write-Host "  [WARN] Default template not found in fallback" -ForegroundColor Yellow
        }
    }
} catch {
    Write-Host "  [FAIL] Exception testing fallback: $($_.Exception.Message)" -ForegroundColor Red
}
Write-Host ""

Write-Host "=== Test Complete ===" -ForegroundColor Cyan
Write-Host ""
Write-Host "Test output files created in: $scriptRoot" -ForegroundColor Gray
Get-ChildItem -Path $scriptRoot -Filter "test_memberberry*.txt" | Select-Object -ExpandProperty Name | ForEach-Object { Write-Host "  - $_" -ForegroundColor Green }
Write-Host ""

