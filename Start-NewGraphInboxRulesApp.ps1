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

# Browser sign-in with an account picker. Graph 2.37 keeps the Windows broker on unless these are set
# before Connect-MgGraph, and a broker failure closes this window with an empty error.
$env:MSAL_FORCE_WAM = '0'
$env:AZURE_IDENTITY_DISABLE_BROKER = 'true'
$env:MSAL_DISABLE_BROKER = '1'
$env:MSAL_EXPERIMENTAL_DISABLE_BROKER = '1'

foreach ($k in @(
        'MSAL_CACHE_DIR'
        'IDENTITY_SERVICE_CACHE_DIR'
    )) {
    if (Test-Path "Env:\$k") { Remove-Item "Env:\$k" -ErrorAction SilentlyContinue }
}

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$inner = Join-Path $scriptRoot 'New-GraphInboxRulesApp.ps1'
if (-not (Test-Path $inner)) {
    throw "Script not found: $inner"
}

& $inner @PSBoundParameters
