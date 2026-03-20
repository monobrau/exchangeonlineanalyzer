<#
.SYNOPSIS
    Tests app-only Microsoft Graph login for EOA-GraphApp-* credentials in Windows Credential Manager.

.DESCRIPTION
    For each tenant ID, requests a client-credentials token (same as the bulk worker) and calls GET
    https://graph.microsoft.com/v1.0/organization to confirm the app can read directory data.

    Use this to verify which stored Graph apps work without signing in interactively.

.PARAMETER TenantId
    One or more directory (tenant) GUIDs. If omitted, every tenant discovered via cmdkey (EOA-GraphApp-*) is tested.

.PARAMETER PassThru
    Return result objects instead of only formatting to the host.

.EXAMPLE
    .\Test-GraphAppWcmLogin.ps1

.EXAMPLE
    .\Test-GraphAppWcmLogin.ps1 -TenantId 'c199e267-e7cc-4a40-9b71-45a404eaec0d'

.EXAMPLE
    .\Test-GraphAppWcmLogin.ps1 -Verbose
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string[]] $TenantId,

    [Parameter(Mandatory = $false)]
    [switch] $PassThru
)

$repoRoot = Split-Path -Parent $PSScriptRoot
$modPath = Join-Path $repoRoot 'Modules\GraphAppCredential.psm1'
if (-not (Test-Path -LiteralPath $modPath)) {
    throw "GraphAppCredential module not found: $modPath (run from repo Scripts folder)."
}
Import-Module $modPath -Force

if (-not $TenantId -or $TenantId.Count -eq 0) {
    $TenantId = @(Get-WCMTenantIds)
}

if ($TenantId.Count -eq 0) {
    Write-Host 'No EOA-GraphApp-* entries found (Windows Credential Manager).' -ForegroundColor Yellow
    Write-Host 'Use the analyzer Import App Creds / Create Graph App, or inspect: cmdkey /list' -ForegroundColor Gray
    if ($PassThru) { return @() }
    exit 1
}

$results = [System.Collections.Generic.List[object]]::new()
foreach ($tid in $TenantId) {
    $wcmErr = $null
    $gParams = @{ TenantId = $tid; FailureVariable = 'wcmErr' }
    if ($VerbosePreference -eq 'Continue') { $gParams['Verbose'] = $true }
    $tok = Get-GraphAppTokenFromWCM @gParams
    $orgName = $null
    $apiErr = $null
    if ($tok) {
        try {
            $headers = @{ Authorization = "Bearer $tok" }
            $orgResp = Invoke-RestMethod -Uri 'https://graph.microsoft.com/v1.0/organization' -Headers $headers -Method GET -ErrorAction Stop
            if ($orgResp.value -and $orgResp.value.Count -gt 0) {
                $orgName = $orgResp.value[0].displayName
            }
        }
        catch {
            $apiErr = $_.Exception.Message
        }
    }
    [void]$results.Add([pscustomobject]@{
            TenantId             = $tid
            TokenAcquired        = [bool]$tok
            WcmOrTokenError      = if ($tok) { $null } else { $wcmErr }
            OrganizationName     = $orgName
            OrganizationApiError = $apiErr
        })
}

if ($PassThru) {
    return $results
}

$results | Format-Table -AutoSize

$failed = @($results | Where-Object { -not $_.TokenAcquired -or $_.OrganizationApiError })
if ($failed.Count -eq 0) {
    Write-Host 'OK: token acquired and /organization succeeded for all tested tenant(s).' -ForegroundColor Green
    exit 0
}

Write-Host "$($failed.Count) tenant(s) failed token or Graph API check - see table above." -ForegroundColor Yellow
exit 2