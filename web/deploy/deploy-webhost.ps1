#Requires -Version 5.1
<#
.SYNOPSIS
  Push current branch to origin, then SSH to the Linux webhost: git pull, pip install, restart eoa-api.

.EXAMPLE
  pwsh web/deploy/deploy-webhost.ps1
  pwsh web/deploy/deploy-webhost.ps1 -SshTarget 'cknospe@192.168.1.50' -Branch main

  # SSH must target the webhost LAN IP (or VPN/jump). The public site hostname (e.g. eoa.knospe.org) is HTTPS-only — not SSH.
#>
param(
    [string]$Branch = "",
    # Override with user@LAN_IP; public DNS hostname typically does not accept SSH.
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
