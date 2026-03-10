# File Collection Module
# Shared functions for collecting and validating investigation report files

# Import SecurityHelpers for enhanced validation if available
$securityHelpersPath = Join-Path (Split-Path $PSScriptRoot -Parent) 'Common\SecurityHelpers.psm1'
if (Test-Path $securityHelpersPath) {
    Import-Module $securityHelpersPath -Force -ErrorAction SilentlyContinue
}

function Get-InvestigationReportFiles {
    param(
        [Parameter(Mandatory=$true)]
        [string]$OutputFolder,
        
        [Parameter(Mandatory=$false)]
        [string[]]$ExtraFiles = @(),
        
        [Parameter(Mandatory=$false)]
        [string[]]$DefaultFiles = @(
            '_AI_Readme.txt',
            'MessageTrace.csv',
            'InboxRules.csv',
            'TransportRules.csv',
            'MailFlowConnectors.csv',
            'GraphAuditLogs.csv',
            'UserSecurityPosture.csv'
        )
    )
    
    $foundFiles = [System.Collections.ArrayList]::new()
    
    # Collect default files
    foreach ($fileName in $DefaultFiles) {
        $filePath = Join-Path $OutputFolder $fileName
        if (Test-Path $filePath) {
            [void]$foundFiles.Add($filePath)
        }
    }
    
    # Validate and add extra files
    if ($ExtraFiles.Count -gt 0) {
        # SECURITY: Use enhanced file path validation
        if (Get-Command Validate-AllFilePaths -ErrorAction SilentlyContinue) {
            $validationResults = Validate-AllFilePaths -FilePaths $ExtraFiles -BaseDirectory $OutputFolder -MustExist
            foreach ($result in $validationResults) {
                if ($result.IsValid) {
                    [void]$foundFiles.Add($result.ValidatedPath)
                } else {
                    Write-Warning "Extra file validation failed: $($result.FilePath) - $($result.Error)"
                }
            }
        } else {
            # Fallback to original validation
            $normalizedOutput = [System.IO.Path]::GetFullPath($OutputFolder)
            foreach ($extraFilePath in $ExtraFiles) {
                if ([string]::IsNullOrWhiteSpace($extraFilePath)) { continue }
                
                try {
                    $resolvedPath = (Resolve-Path $extraFilePath -ErrorAction Stop).Path
                    $normalizedFile = [System.IO.Path]::GetFullPath($resolvedPath)
                    
                    # Path traversal protection
                    if (-not $normalizedFile.StartsWith($normalizedOutput, [System.StringComparison]::OrdinalIgnoreCase)) {
                        Write-Warning "Extra file outside output folder, skipping (path traversal protection): $extraFilePath"
                        continue
                    }
                    
                    if (Test-Path $resolvedPath) {
                        [void]$foundFiles.Add($resolvedPath)
                    }
                } catch {
                    Write-Warning "Extra file not found or invalid: $extraFilePath - $($_.Exception.Message)"
                }
            }
        }
    }
    
    return ($foundFiles | Select-Object -Unique)
}

Export-ModuleMember -Function Get-InvestigationReportFiles
