# Test script for ticket parsing and AI readme generation
# This script tests the ticket integration functionality without running the full bulk exporter

param(
    [string]$TicketContent = @"
Calendar

Service Ticket

Service Ticket Search

Service Ticket
Service Ticket #1811523 - [126585259] Barracuda XDR - Microsoft Office 365 Anomalous Login RMM - Actionable Alert 
Special Configuration Note - Click to see more

Summary:
*
[126585259] Barracuda XDR - Microsoft Office 365 Anomalous Login RMM - Actionable Alert
Age: 2d 10h 28m
Company: Acieta
Company
:
*
Acieta
Contact
:
 
Steve Bucek
 
(712) 388-6410
Email
:
 
it@acieta.com
Site
:
 
Acieta - WI
Address 1
:
 
N25 W23790 Commerce Circle
Address 2
:
 
City
:
 
Waukesha
State
:
 
WI
Zip
:
 
53188
Country
:
 
United States
Ticket #1811523
Board
:
*
RMM - Actionable Alerts
Status
:
*
Dispatch
Type
:
 
SOC
Subtype
:
 
Item
:
 
Ticket Owner
:
 
(Unassigned)
 
IssueSKOUT CYBERSECURITY12/12/2025 10:07 AM
Description: Incident Name: Microsoft Office 365 Anomalous Login
Organization Name: Acieta
MITRE ATT&CK: Initial Access (TA0001), Valid Accounts (T1078), Cloud Accounts (T1078.004)
Risk: Medium
Barracuda XDR Risk Score: 0/1001
Ticket #: 126585259
Time the incident occurred: Friday December 12 2025, 04:04 PM UTC
 Data Source: Microsoft 365

What is the Threat:

Barracuda XDR has detected an anomalous login for the Office 365 user, "nvalentine@Acieta.com". Using machine learning to model 20+ login features to identify unusual user login activity, this detection flags events that are scored as 'highly anomalous' compared to the user's historical patterns and behaviors.

ANOMALY LOGIN EVENT
User nvalentine@Acieta.com
Login Time Friday December 12 2025, 03:40 PM UTC
Login IP 208.103.30.53 (Ligtel Communications, Inc.)
Login Origin Avilla, Indiana, United States
Login Device FA01536-PM18P
Application e87e3225-09ae-487a-83bd-abdccceb8fc5 (Unknown application)
"@
)

Write-Host ("=" * 80) -ForegroundColor Cyan
Write-Host "Ticket Integration Test Script" -ForegroundColor Cyan
Write-Host ("=" * 80) -ForegroundColor Cyan
Write-Host ""

# Get script root
$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $scriptRoot) {
    $scriptRoot = $PSScriptRoot
}
if (-not $scriptRoot) {
    $scriptRoot = Get-Location
}

Write-Host "Script root: $scriptRoot" -ForegroundColor Gray
Write-Host ""

# Import Settings module
Write-Host "Importing Settings module..." -ForegroundColor Yellow
try {
    Import-Module "$scriptRoot\Modules\Settings.psm1" -Force -ErrorAction Stop
    Write-Host "Settings module imported successfully" -ForegroundColor Green
} catch {
    Write-Error "Failed to import Settings module: $($_.Exception.Message)"
    exit 1
}

Write-Host ""

# Test 1: Extract Ticket Numbers
Write-Host ("=" * 80) -ForegroundColor Cyan
Write-Host "Test 1: Extract Ticket Numbers" -ForegroundColor Cyan
Write-Host ("=" * 80) -ForegroundColor Cyan
Write-Host ""

if (Get-Command Extract-TicketNumbers -ErrorAction SilentlyContinue) {
    Write-Host "Testing Extract-TicketNumbers function..." -ForegroundColor Yellow
    $ticketNumbers = Extract-TicketNumbers -TicketContent $TicketContent
    Write-Host "Extracted ticket numbers: $($ticketNumbers -join ', ')" -ForegroundColor Green
    Write-Host "Count: $($ticketNumbers.Count)" -ForegroundColor Green
    Write-Host ""
} else {
    Write-Error "Extract-TicketNumbers function not found!"
    exit 1
}

# Test 2: Filter Ticket Content
Write-Host ("=" * 80) -ForegroundColor Cyan
Write-Host "Test 2: Filter Ticket Content" -ForegroundColor Cyan
Write-Host ("=" * 80) -ForegroundColor Cyan
Write-Host ""

if (Get-Command Filter-TicketContent -ErrorAction SilentlyContinue) {
    Write-Host "Testing Filter-TicketContent function..." -ForegroundColor Yellow
    $filteredContent = Filter-TicketContent -TicketContent $TicketContent
    Write-Host "Original length: $($TicketContent.Length) characters" -ForegroundColor Gray
    Write-Host "Filtered length: $($filteredContent.Length) characters" -ForegroundColor Gray
    Write-Host ""
    Write-Host "First 500 characters of filtered content:" -ForegroundColor Yellow
    Write-Host $filteredContent.Substring(0, [Math]::Min(500, $filteredContent.Length)) -ForegroundColor White
    Write-Host ""
} else {
    Write-Error "Filter-TicketContent function not found!"
    exit 1
}

# Test 3: Load Settings
Write-Host ("=" * 80) -ForegroundColor Cyan
Write-Host "Test 3: Load Settings" -ForegroundColor Cyan
Write-Host ("=" * 80) -ForegroundColor Cyan
Write-Host ""

if (Get-Command Get-AppSettings -ErrorAction SilentlyContinue) {
    Write-Host "Loading app settings..." -ForegroundColor Yellow
    $settings = Get-AppSettings
    Write-Host "Settings loaded successfully" -ForegroundColor Green
    Write-Host "Investigator Name: $($settings.InvestigatorName)" -ForegroundColor Gray
    Write-Host "Company Name: $($settings.CompanyName)" -ForegroundColor Gray
    Write-Host ""
} else {
    Write-Error "Get-AppSettings function not found!"
    exit 1
}

# Test 4: Generate AI Readme with Ticket Data
Write-Host ("=" * 80) -ForegroundColor Cyan
Write-Host "Test 4: Generate AI Readme with Ticket Data" -ForegroundColor Cyan
Write-Host ("=" * 80) -ForegroundColor Cyan
Write-Host ""

if (Get-Command New-AIReadme -ErrorAction SilentlyContinue) {
    Write-Host "Generating AI readme with ticket data..." -ForegroundColor Yellow
    Write-Host "Ticket Numbers: $($ticketNumbers -join ', ')" -ForegroundColor Gray
    Write-Host "Ticket Content Length: $($filteredContent.Length) characters" -ForegroundColor Gray
    Write-Host ""
    
    $aiReadme = New-AIReadme -Settings $settings -TicketNumbers $ticketNumbers -TicketContent $filteredContent
    
    Write-Host "AI Readme generated successfully!" -ForegroundColor Green
    Write-Host "Length: $($aiReadme.Length) characters" -ForegroundColor Gray
    Write-Host ""
    
    # Check if ticket information is included
    if ($aiReadme -match "Ticket Information" -or $aiReadme -match "Ticket #") {
        Write-Host "✓ Ticket information section found in AI readme" -ForegroundColor Green
    } else {
        Write-Warning "✗ Ticket information section NOT found in AI readme"
    }
    
    # Check if ticket content is actually present (not just the header)
    $ticketContentStart = $aiReadme.IndexOf("**Ticket Content**:")
    if ($ticketContentStart -ge 0) {
        $ticketContentSection = $aiReadme.Substring($ticketContentStart, [Math]::Min(500, $aiReadme.Length - $ticketContentStart))
        if ($ticketContentSection -match "Calendar|Service Ticket|Barracuda|nvalentine" -or $ticketContentSection.Length -gt 50) {
            Write-Host "✓ Ticket content found in AI readme (length: $($ticketContentSection.Length) chars after header)" -ForegroundColor Green
            Write-Host "  Preview: $($ticketContentSection.Substring(0, [Math]::Min(200, $ticketContentSection.Length)))" -ForegroundColor Gray
        } else {
            Write-Warning "✗ Ticket content header found but content appears empty or too short"
            Write-Host "  Content after header: $ticketContentSection" -ForegroundColor Yellow
        }
    } else {
        Write-Warning "✗ Ticket content header NOT found in AI readme"
    }
    
    # Check if ticket number is in subject line
    if ($aiReadme -match "Ticket.*1811523" -or $aiReadme -match "#1811523") {
        Write-Host "✓ Ticket number found in AI readme" -ForegroundColor Green
    } else {
        Write-Warning "✗ Ticket number NOT found in AI readme"
    }
    
    Write-Host ""
    Write-Host "First 1000 characters of AI readme:" -ForegroundColor Yellow
    Write-Host $aiReadme.Substring(0, [Math]::Min(1000, $aiReadme.Length)) -ForegroundColor White
    Write-Host ""
    
    # Save to test file
    $testOutputPath = Join-Path $scriptRoot "test_output_AI_Readme.txt"
    $aiReadme | Out-File -FilePath $testOutputPath -Encoding utf8
    Write-Host "AI readme saved to: $testOutputPath" -ForegroundColor Green
    Write-Host ""
    
    # Test generating separate readme for each ticket
    if ($ticketNumbers.Count -gt 0) {
        Write-Host ("=" * 80) -ForegroundColor Cyan
        Write-Host "Test 5: Generate Separate AI Readme per Ticket" -ForegroundColor Cyan
        Write-Host ("=" * 80) -ForegroundColor Cyan
        Write-Host ""
        
        foreach ($ticketNum in $ticketNumbers) {
            Write-Host "Generating AI readme for ticket #$ticketNum..." -ForegroundColor Yellow
            $ticketReadme = New-AIReadme -Settings $settings -TicketNumbers @($ticketNum) -TicketContent $filteredContent
            
            $ticketOutputPath = Join-Path $scriptRoot "test_output_AI_Readme_Ticket_$ticketNum.txt"
            $ticketReadme | Out-File -FilePath $ticketOutputPath -Encoding utf8
            Write-Host "Ticket-specific readme saved to: $ticketOutputPath" -ForegroundColor Green
            
            # Verify ticket number is in filename and content
            if ($ticketReadme -match "Ticket.*$ticketNum" -or $ticketReadme -match "#$ticketNum") {
                Write-Host "✓ Ticket #$ticketNum found in readme content" -ForegroundColor Green
            } else {
                Write-Warning "✗ Ticket #$ticketNum NOT found in readme content"
            }
            Write-Host ""
        }
    }
    
} else {
    Write-Error "New-AIReadme function not found!"
    exit 1
}

Write-Host ("=" * 80) -ForegroundColor Cyan
Write-Host "All Tests Completed!" -ForegroundColor Green
Write-Host ("=" * 80) -ForegroundColor Cyan
Write-Host ""
Write-Host "Test output files created in: $scriptRoot" -ForegroundColor Gray
Write-Host "  - test_output_AI_Readme.txt" -ForegroundColor Gray
if ($ticketNumbers.Count -gt 0) {
    foreach ($ticketNum in $ticketNumbers) {
        Write-Host "  - test_output_AI_Readme_Ticket_$ticketNum.txt" -ForegroundColor Gray
    }
}
Write-Host ""

