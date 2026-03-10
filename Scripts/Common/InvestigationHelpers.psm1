# Investigation Helpers Module
# Shared functions for finding and working with investigation folders

function Get-LatestInvestigationFolder {
    param(
        [Parameter(Mandatory=$false)]
        [switch]$IncludeLegacy,
        
        [Parameter(Mandatory=$false)]
        [switch]$VerboseOutput
    )
    
    $base = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'ExchangeOnlineAnalyzer\SecurityInvestigation'
    if (-not (Test-Path $base)) { 
        if ($VerboseOutput) { Write-Verbose "Investigation base folder not found: $base" }
        return $null 
    }
    
    $investigationFolders = @()
    
    try {
        $tenants = Get-ChildItem -Path $base -Directory -ErrorAction Stop
        foreach ($tenant in $tenants) {
            try {
                $runs = Get-ChildItem -Path $tenant.FullName -Directory -ErrorAction Stop | Sort-Object LastWriteTime -Descending
                if ($runs -and $runs.Count -gt 0) { 
                    $investigationFolders += $runs 
                }
            } catch {
                if ($VerboseOutput) { Write-Verbose "Failed to enumerate runs in tenant folder $($tenant.FullName): $($_.Exception.Message)" }
            }
        }
        
        # Include legacy format folders if requested
        if ($IncludeLegacy) {
            $legacyFolders = Get-ChildItem -Path $base -Directory -ErrorAction Stop | Where-Object { $_.Name -match '^\d{8}_\d{6}$' }
            if ($legacyFolders) { 
                $investigationFolders += $legacyFolders 
            }
        }
    } catch {
        if ($VerboseOutput) { Write-Warning "Failed to enumerate investigation folders: $($_.Exception.Message)" }
        return $null
    }
    
    if (-not $investigationFolders -or $investigationFolders.Count -eq 0) { 
        return $null 
    }
    
    return ($investigationFolders | Sort-Object LastWriteTime -Descending | Select-Object -First 1).FullName
}

Export-ModuleMember -Function Get-LatestInvestigationFolder
