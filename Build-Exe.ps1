<#
.SYNOPSIS
    Builds ExchangeOnlineAnalyzer.exe and BulkTenantExporter.exe using ps2exe.
.DESCRIPTION
    Requires: Install-Module ps2exe -Scope CurrentUser
    Run: Install-Module ps2exe -Scope CurrentUser
    Output: Release\ folder with exes + Modules, Scripts (required at runtime)
#>
$ErrorActionPreference = "Stop"
$scriptDir = $PSScriptRoot
$releaseDir = Join-Path $scriptDir "Release"

if (-not (Get-Module ps2exe -ListAvailable)) {
    Write-Host "Installing ps2exe module..." -ForegroundColor Yellow
    try {
        Install-Module ps2exe -Scope CurrentUser -Force
    } catch {
        Write-Host "Failed to install ps2exe. Run: Install-Module ps2exe -Scope CurrentUser" -ForegroundColor Red
        exit 1
    }
}
Import-Module ps2exe -Force -ErrorAction Stop

if (-not (Test-Path $releaseDir)) {
    New-Item -ItemType Directory -Path $releaseDir -Force | Out-Null
}

# Build main application (GUI - no console)
Write-Host "Building ExchangeOnlineAnalyzer.exe..." -ForegroundColor Cyan
Invoke-PS2EXE -inputFile (Join-Path $scriptDir "ExchangeOnlineAnalyzer.ps1") `
    -outputFile (Join-Path $releaseDir "ExchangeOnlineAnalyzer.exe") `
    -noConsole `
    -title "Microsoft 365 Management Tool" `
    -description "Exchange Online and Entra ID analysis tool" `
    -version "8.4.0.0" `
    -company "ExchangeOnlineAnalyzer" `
    -product "Microsoft 365 Management Tool"

# Build bulk exporter (GUI - no console)
Write-Host "Building BulkTenantExporter.exe..." -ForegroundColor Cyan
Invoke-PS2EXE -inputFile (Join-Path $scriptDir "BulkTenantExporter.ps1") `
    -outputFile (Join-Path $releaseDir "BulkTenantExporter.exe") `
    -noConsole `
    -title "Bulk Tenant Report Exporter" `
    -description "Bulk security report export for multiple tenants" `
    -version "8.4.0.0" `
    -company "ExchangeOnlineAnalyzer" `
    -product "Bulk Tenant Report Exporter"

# Copy Modules and Scripts (required at runtime - exe uses $PSScriptRoot)
Write-Host "Copying Modules and Scripts..." -ForegroundColor Cyan
Copy-Item -Path (Join-Path $scriptDir "Modules") -Destination $releaseDir -Recurse -Force
Copy-Item -Path (Join-Path $scriptDir "Scripts") -Destination $releaseDir -Recurse -Force
Copy-Item -Path (Join-Path $scriptDir "readme.md") -Destination $releaseDir -Force -ErrorAction SilentlyContinue

Write-Host "Build complete. Output: $releaseDir" -ForegroundColor Green
Get-ChildItem $releaseDir -Recurse -File | Where-Object { $_.Extension -in '.exe','.psm1','.ps1','.md' } | Select-Object -First 20 | ForEach-Object { Write-Host "  $($_.FullName.Replace($releaseDir,''))" }
