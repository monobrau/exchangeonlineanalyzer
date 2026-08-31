<#
.SYNOPSIS
    Runs New-GraphInboxRulesApp.ps1 in a fresh PowerShell process scope (child process = new default runspace, no host Graph state).
.DESCRIPTION
    The hosting app may set MSAL/Identity env vars or load Microsoft.Graph; this launcher clears those before delegating so the
    create-app script does not inherit another session's workspace. Invoked via Start-Process from the UI.
#>
#Requires -Version 5.1

param(
    [switch]$SaveToWCM = $false,
    [string]$TenantId = $null,
    [switch]$UseDeviceCode = $false,
    [switch]$UpdateExisting = $false
)

$ErrorActionPreference = 'Stop'

# Bypass WAM so the inner script gets browser-based auth with an account picker.
# Must be set before Microsoft.Graph.Authentication loads.
$env:MSAL_FORCE_WAM = '0'

foreach ($k in @(
        'MSAL_CACHE_DIR'
        'IDENTITY_SERVICE_CACHE_DIR'
        'AZURE_IDENTITY_DISABLE_BROKER'
        'MSAL_DISABLE_BROKER'
        'MSAL_EXPERIMENTAL_DISABLE_BROKER'
    )) {
    if (Test-Path "Env:\$k") { Remove-Item "Env:\$k" -ErrorAction SilentlyContinue }
}

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$inner = Join-Path $scriptRoot 'New-GraphInboxRulesApp.ps1'
if (-not (Test-Path $inner)) {
    throw "Script not found: $inner"
}

& $inner @PSBoundParameters
