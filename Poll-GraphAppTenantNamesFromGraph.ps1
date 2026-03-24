<#
.SYNOPSIS
    Lists every tenant that has EOA/ESR Graph app credentials in WCM and resolves the organization display name from Microsoft Graph.
.DESCRIPTION
    For each stored app (EOA-GraphApp-* / ESR-GraphApp-*), uses client_credentials from Windows Credential Manager and calls
    GET https://graph.microsoft.com/v1.0/organization to read displayName (same source as the in-app dropdown when Graph lookup runs).

    Does not list Entra "app registrations" globally — only tenants that have credentials saved locally in WCM.
.PARAMETER Prefix
    EOA, ESR, or Both (default).
.PARAMETER SaveToWcm
    After resolving each name from Graph, write EOA/ESR *-DisplayName entries via Save-GraphAppCredentialToWCM (same as Register-GraphAppTenantDisplayNamesInWCM).
.PARAMETER CsvPath
    Optional path to export results as CSV.
.PARAMETER PassThru
    Return objects to the pipeline instead of formatting a table (CSV still written if -CsvPath is set).
.EXAMPLE
    .\Poll-GraphAppTenantNamesFromGraph.ps1
.EXAMPLE
    .\Poll-GraphAppTenantNamesFromGraph.ps1 -Prefix EOA -SaveToWcm -CsvPath "$env:USERPROFILE\Desktop\graph-tenant-names.csv"
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [ValidateSet('EOA', 'ESR', 'Both')]
    [string]$Prefix = 'Both',

    [Parameter(Mandatory = $false)]
    [switch]$SaveToWcm,

    [Parameter(Mandatory = $false)]
    [string]$CsvPath,

    [Parameter(Mandatory = $false)]
    [switch]$PassThru
)

$ErrorActionPreference = 'Stop'
$graphMod = $null
foreach ($dir in @($PSScriptRoot, (Split-Path -Parent $PSScriptRoot))) {
    $candidate = Join-Path $dir 'Modules\GraphAppCredential.psm1'
    if (Test-Path -LiteralPath $candidate) {
        $graphMod = $candidate
        break
    }
}
if (-not $graphMod) {
    throw "GraphAppCredential.psm1 not found under repo root or parent of script folder (expected .\Modules\GraphAppCredential.psm1)."
}
Import-Module $graphMod -Force

$prefixes = if ($Prefix -eq 'Both') { @('EOA', 'ESR') } else { @($Prefix) }
$rows = [System.Collections.Generic.List[object]]::new()

foreach ($pfx in $prefixes) {
    foreach ($tid in @(Get-WCMTenantIds -Prefix $pfx)) {
        if ([string]::IsNullOrWhiteSpace($tid)) { continue }

        $tokenErr = $null
        $tok = Get-GraphAppTokenFromWCM -TenantId $tid -Prefix $pfx -FailureVariable tokenErr -ErrorAction SilentlyContinue
        if (-not $tok) {
            [void]$rows.Add([pscustomobject]@{
                    TenantId    = $tid
                    WcmPrefix   = $pfx
                    DisplayName = $null
                    Source      = 'Graph /organization'
                    Error       = $(if ($tokenErr) { [string]$tokenErr } else { 'Token request failed' })
                })
            continue
        }

        $name = Get-TenantDisplayNameFromWCM -TenantId $tid -Prefix $pfx -ForceRefresh
        if (-not $name) {
            $alt = if ($pfx -eq 'EOA') { 'ESR' } else { 'EOA' }
            $name = Get-TenantDisplayNameFromWCM -TenantId $tid -Prefix $alt -ForceRefresh
        }

        $err = $null
        if (-not $name) {
            $err = 'Graph /organization returned no displayName (check Organization.Read.All or Directory.Read.All for the app)'
        }

        [void]$rows.Add([pscustomobject]@{
                TenantId    = $tid
                WcmPrefix   = $pfx
                DisplayName = $name
                Source      = 'Graph /organization'
                Error       = $err
            })

        if ($SaveToWcm -and $name) {
            $c = Get-GraphAppCredentialFromWCM -TenantId $tid -Prefix $pfx
            if ($c) {
                try {
                    Save-GraphAppCredentialToWCM -TenantId $c.TenantId -ClientId $c.ClientId -ClientSecret $c.ClientSecret -TenantDisplayName $name -Prefix $pfx
                }
                catch {
                    Write-Warning "SaveToWcm failed for $pfx $tid : $($_.Exception.Message)"
                }
            }
        }
    }
}

$out = @($rows | Sort-Object WcmPrefix, DisplayName, TenantId)

if ($CsvPath) {
    $out | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8
    Write-Host "Wrote CSV: $CsvPath"
}

if ($SaveToWcm) {
    Write-Host "WCM *-DisplayName entries updated where Graph returned a name (per prefix)."
}

if ($PassThru) {
    return $out
}

$out | Format-Table -AutoSize
