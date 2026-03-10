<#
.SYNOPSIS
    Build exe, push changes, and create GitHub release v8.4.
.DESCRIPTION
    Requires: git, ps2exe (Install-Module ps2exe), gh CLI (optional)
#>
$ErrorActionPreference = "Stop"
$scriptDir = $PSScriptRoot
Set-Location $scriptDir

# 1. Build exe
Write-Host "Building exe..." -ForegroundColor Cyan
try {
    & (Join-Path $scriptDir "Build-Exe.ps1")
} catch {
    Write-Host "Build failed (ps2exe may not be installed). Continuing with push..." -ForegroundColor Yellow
}

# 2. Create release zip if Release exists
$releaseDir = Join-Path $scriptDir "Release"
$zipPath = Join-Path $scriptDir "ExchangeOnlineAnalyzer-v8.4.zip"
if (Test-Path $releaseDir) {
    if (Test-Path $zipPath) { Remove-Item $zipPath -Force }
    Compress-Archive -Path (Join-Path $releaseDir "*") -DestinationPath $zipPath -Force
    Write-Host "Created $zipPath" -ForegroundColor Green
}

# 3. Git add, commit, push
Write-Host "Git add/commit/push..." -ForegroundColor Cyan
git add -A
git status
git commit -m "v8.4 release - Validation fixes, Exchange auth speed, Security Incidents off by default" 2>$null
git push origin main

# 4. Tag and release
Write-Host "Creating tag v8.4..." -ForegroundColor Cyan
git tag -a v8.4 -m "v8.4 release" 2>$null
git push origin v8.4

# 5. GitHub release (if gh installed and zip exists)
if ((Get-Command gh -ErrorAction SilentlyContinue) -and (Test-Path $zipPath)) {
    gh release create v8.4 $zipPath --title "v8.4" --notes "Release v8.4 - Validation fixes, Exchange auth speed, Security Incidents off by default"
} elseif (Test-Path $zipPath) {
    Write-Host "Attach $zipPath manually to the v8.4 release on GitHub" -ForegroundColor Yellow
}
