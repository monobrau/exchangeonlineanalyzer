<#
.SYNOPSIS
  Creates an Entra app registration (SPA) for EOA browser MSAL — User.Read, Organization.Read.All, Application.ReadWrite.All (delegated).

.DESCRIPTION
  Run once in a tenant where you can create app registrations. Copy the printed Application (client) ID into
  web/app/bundled_ms_graph.py as BUNDLED_MS_GRAPH_SPA_CLIENT_ID (or set EOA_MS_GRAPH_SPA_CLIENT_ID).

  By default the app is single-tenant (Accounts in this organizational directory only). Use -Multitenant when you
  publish one client ID for customers in many tenants (ISV / SaaS style).

  Requires: Microsoft.Graph.Authentication, Microsoft.Graph.Applications (Install-Module Microsoft.Graph)
  Connect with: Connect-MgGraph -Scopes Application.ReadWrite.All,DelegatedPermissionGrant.ReadWrite.All
#>
#Requires -Version 5.1
[CmdletBinding()]
param(
    [string]$DisplayName = "Exchange Online Analyzer - Web MSAL (SPA)",
    [switch]$Multitenant
)

$ErrorActionPreference = "Stop"

$extraModuleRoots = @(
    (Join-Path $env:USERPROFILE "Documents\PowerShell\Modules")
    (Join-Path $env:USERPROFILE "OneDrive\Documents\PowerShell\Modules")
)
foreach ($root in $extraModuleRoots) {
    if (Test-Path $root) {
        $env:PSModulePath = "$root;$env:PSModulePath"
    }
}

$authMod = Get-Module -ListAvailable Microsoft.Graph.Authentication | Sort-Object Version -Descending | Select-Object -First 1
$appMod = Get-Module -ListAvailable Microsoft.Graph.Applications | Sort-Object Version -Descending | Select-Object -First 1
if (-not $authMod -or -not $appMod) {
    Write-Error "Install Microsoft Graph PowerShell: Install-Module Microsoft.Graph -Scope CurrentUser"
}
Import-Module $authMod.Path -ErrorAction Stop
Import-Module $appMod.Path -ErrorAction Stop

$graphAppId = "00000003-0000-0000-c000-000000000000" # Microsoft Graph
$perms = @(
    @{ Id = "e1fe6dd8-ba31-4d61-89e7-88639da4683d"; Type = "Scope" }  # User.Read
    @{ Id = "498476ce-e0fe-48b0-b801-37ba7e2685c6"; Type = "Scope" }  # Organization.Read.All
    @{ Id = "1bfefb4e-e0b5-418b-a88f-73c46d2cc8e9"; Type = "Scope" }  # Application.ReadWrite.All
)

$redirectUris = @(
    "http://127.0.0.1:8080/"
    "http://127.0.0.1:8080/app"
    "http://localhost:8080/"
    "http://localhost:8080/app"
    "https://eoa.knospe.org/"
    "https://eoa.knospe.org/app"
)

$audience = if ($Multitenant) { "AzureADMultipleOrgs" } else { "AzureADMyOrg" }
Write-Host "signInAudience: $audience$(if ($Multitenant) { ' (any org may use this app registration)' } else { ' (this directory only)' })"

$body = @{
    displayName     = $DisplayName
    signInAudience  = $audience
    spa             = @{ redirectUris = $redirectUris }
    requiredResourceAccess = @(
        @{
            resourceAppId  = $graphAppId
            resourceAccess = $perms
        }
    )
}

Write-Host "Connecting to Microsoft Graph (sign in if prompted)..."
Connect-MgGraph -Scopes @(
    "Application.ReadWrite.All"
    "DelegatedPermissionGrant.ReadWrite.All"
) -NoWelcome

Write-Host "Creating application registration..."
$created = New-MgApplication -BodyParameter $body
$cid = $created.AppId
Write-Host ""
Write-Host "Application (client) ID: $cid"
Write-Host ""
Write-Host "Add to web/app/bundled_ms_graph.py:"
Write-Host "  BUNDLED_MS_GRAPH_SPA_CLIENT_ID = `"$cid`""
Write-Host ""
Write-Host "Then in Entra: API permissions, Grant admin consent (Application.ReadWrite.All needs admin consent per tenant)."
Disconnect-MgGraph | Out-Null
