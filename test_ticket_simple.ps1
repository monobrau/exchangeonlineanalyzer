# Simple test script to verify ticket parsing and AI readme generation
param(
    [string]$TicketContent = "Service Ticket #1811523 - Barracuda XDR Alert for nvalentine@Acieta.com"
)

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $scriptRoot) { $scriptRoot = $PSScriptRoot }
if (-not $scriptRoot) { $scriptRoot = Get-Location }

Write-Host "`n=== Testing Ticket Integration ===" -ForegroundColor Cyan
Write-Host ""

# Import Settings module
Import-Module "$scriptRoot\Modules\Settings.psm1" -Force -ErrorAction Stop

# Test 1: Extract ticket numbers
Write-Host "Test 1: Extract Ticket Numbers" -ForegroundColor Yellow
$ticketNumbers = Extract-TicketNumbers -TicketContent $TicketContent
Write-Host "  Extracted: $($ticketNumbers -join ', ')" -ForegroundColor $(if ($ticketNumbers.Count -gt 0) { 'Green' } else { 'Red' })
Write-Host ""

# Test 2: Filter ticket content
Write-Host "Test 2: Filter Ticket Content" -ForegroundColor Yellow
$filteredContent = Filter-TicketContent -TicketContent $TicketContent
Write-Host "  Original length: $($TicketContent.Length) chars" -ForegroundColor Gray
Write-Host "  Filtered length: $($filteredContent.Length) chars" -ForegroundColor Gray
Write-Host ""

# Test 3: Generate AI readme
Write-Host "Test 3: Generate AI Readme" -ForegroundColor Yellow
$settings = Get-AppSettings
$aiReadme = New-AIReadme -Settings $settings -TicketNumbers $ticketNumbers -TicketContent $filteredContent

# Check for ticket content
$hasTicketSection = $aiReadme -match "## ConnectWise Ticket Information"
$hasTicketContent = $aiReadme -match "Ticket Content"
$hasTicketNumber = $aiReadme -match "#1811523"
$hasTicketDetails = $aiReadme -match "nvalentine|Barracuda"

Write-Host "  Ticket section header: $(if ($hasTicketSection) { '✓ FOUND' } else { '✗ MISSING' })" -ForegroundColor $(if ($hasTicketSection) { 'Green' } else { 'Red' })
Write-Host "  Ticket content header: $(if ($hasTicketContent) { '✓ FOUND' } else { '✗ MISSING' })" -ForegroundColor $(if ($hasTicketContent) { 'Green' } else { 'Red' })
Write-Host "  Ticket number (#1811523): $(if ($hasTicketNumber) { '✓ FOUND' } else { '✗ MISSING' })" -ForegroundColor $(if ($hasTicketNumber) { 'Green' } else { 'Red' })
Write-Host "  Ticket details (nvalentine/Barracuda): $(if ($hasTicketDetails) { '✓ FOUND' } else { '✗ MISSING' })" -ForegroundColor $(if ($hasTicketDetails) { 'Green' } else { 'Red' })
Write-Host ""

# Show ticket content location
$ticketContentPos = $aiReadme.IndexOf("**Ticket Content**:")
if ($ticketContentPos -ge 0) {
    $afterHeader = $aiReadme.Substring($ticketContentPos + 20, [Math]::Min(200, $aiReadme.Length - $ticketContentPos - 20))
    Write-Host "  Content after 'Ticket Content:' header:" -ForegroundColor Yellow
    Write-Host "  $($afterHeader.Substring(0, [Math]::Min(150, $afterHeader.Length)))..." -ForegroundColor White
} else {
    Write-Host "  ✗ 'Ticket Content:' header not found in AI readme!" -ForegroundColor Red
}
Write-Host ""

# Save to file
$outputFile = Join-Path $scriptRoot "test_simple_output.txt"
$aiReadme | Out-File -FilePath $outputFile -Encoding utf8
Write-Host "AI readme saved to: $outputFile" -ForegroundColor Green
Write-Host "  File size: $((Get-Item $outputFile).Length) bytes" -ForegroundColor Gray
Write-Host ""

Write-Host "=== Test Complete ===" -ForegroundColor Cyan
Write-Host ""

