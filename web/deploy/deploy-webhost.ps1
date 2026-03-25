#Requires -Version 5.1
<#
.SYNOPSIS
  Push current branch to origin, then SSH to the Linux webhost: git pull, pip install, restart eoa-api.

.EXAMPLE
  pwsh web/deploy/deploy-webhost.ps1
  pwsh web/deploy/deploy-webhost.ps1 -SshTarget 'cknospe@your-server.example.com' -Branch main
#>
param(
    [string]$Branch = "",
    [string]$SshTarget = "cknospe@eoa.knospe.org",
    [string]$RepoRootOnServer = "/home/cknospe/git/exchangeonlineanalyzer"
)

$ErrorActionPreference = "Stop"
$repo = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
Set-Location $repo

if (-not $Branch) {
    $Branch = (git rev-parse --abbrev-ref HEAD).Trim()
}

Write-Host "Pushing $Branch to origin..."
git push origin $Branch
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

$remote = "cd $RepoRootOnServer && git fetch origin && git checkout $Branch && git pull origin $Branch && cd web && .venv/bin/pip install -r requirements.txt && sudo systemctl restart eoa-api && sudo systemctl is-active eoa-api"
Write-Host "SSH $SshTarget ..."
ssh $SshTarget $remote
exit $LASTEXITCODE
