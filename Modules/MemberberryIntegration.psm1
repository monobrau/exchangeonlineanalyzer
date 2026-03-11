# MemberberryIntegration.psm1
# Integration module for ExchangeOnlineAnalyzer to use memberberry's ticket processing scripts
# Enables reuse of memberberry's clean-ticket, extract-company, detect-alert-type, and compile scripts

# Module-level variables
$script:MemberberryPath = $null
$script:MemberberryPathCache = $null

<#
.SYNOPSIS
    Gets the memberberry installation path by checking common locations
.DESCRIPTION
    Searches for memberberry installation in common locations:
    - c:\git\memberberry
    - $env:USERPROFILE\Documents\memberberry
    - Current directory (if running from memberberry folder)
    - Checks config.json for custom path
#>
function Get-MemberberryScriptPath {
    # Return cached path if available
    if ($script:MemberberryPathCache) {
        return $script:MemberberryPathCache
    }
    
    $possiblePaths = @(
        "c:\git\memberberry",
        "$env:USERPROFILE\Documents\memberberry",
        $PSScriptRoot  # Current script directory
    )
    
    # Check if we're already in a memberberry directory
    foreach ($path in $possiblePaths) {
        if (Test-Path $path) {
            $cleanTicketPath = Join-Path $path "clean-ticket.ps1"
            if (Test-Path $cleanTicketPath) {
                $script:MemberberryPathCache = $path
                return $path
            }
        }
    }
    
    # Check config.json for custom path (if running from memberberry context)
    foreach ($path in $possiblePaths) {
        $configPath = Join-Path $path "config.json"
        if (Test-Path $configPath) {
            try {
                $config = Get-Content $configPath -Raw | ConvertFrom-Json
                if ($config.memberberry_path -and (Test-Path $config.memberberry_path)) {
                    $cleanTicketPath = Join-Path $config.memberberry_path "clean-ticket.ps1"
                    if (Test-Path $cleanTicketPath) {
                        $script:MemberberryPathCache = $config.memberberry_path
                        return $config.memberberry_path
                    }
                }
            } catch {
                # Ignore config errors
            }
        }
    }
    
    return $null
}

<#
.SYNOPSIS
    Cleans ticket content using memberberry's clean-ticket.ps1 script
.DESCRIPTION
    Removes configuration device list sections from ConnectWise tickets
    to reduce token usage and improve LLM analysis focus.
.PARAMETER TicketContent
    The raw ticket text to clean
.EXAMPLE
    $cleaned = Invoke-MemberberryCleanTicket -TicketContent $rawTicket
#>
function Invoke-MemberberryCleanTicket {
    param(
        [Parameter(Mandatory=$true)]
        [string]$TicketContent
    )
    
    $memberberryPath = Get-MemberberryScriptPath
    if (-not $memberberryPath) {
        Write-Warning "Memberberry not found. Cannot clean ticket content. Using original content."
        return $TicketContent
    }
    
    $cleanTicketScript = Join-Path $memberberryPath "clean-ticket.ps1"
    if (-not (Test-Path $cleanTicketScript)) {
        Write-Warning "clean-ticket.ps1 not found at $cleanTicketScript. Using original content."
        return $TicketContent
    }
    
    try {
        # Use PowerShell parameter passing to prevent injection
        # clean-ticket.ps1 accepts InputText parameter or pipeline input
        $cleanedContent = & $cleanTicketScript -InputText $TicketContent
        if ([string]::IsNullOrWhiteSpace($cleanedContent)) {
            # Fallback to pipeline if parameter doesn't work
            $cleanedContent = $TicketContent | & $cleanTicketScript
        }
        return $cleanedContent
    } catch {
        Write-Warning "Error cleaning ticket with memberberry: $($_.Exception.Message). Using original content."
        return $TicketContent
    }
}

<#
.SYNOPSIS
    Extracts company name from ticket content using memberberry's extract-company.ps1 script
.DESCRIPTION
    Searches ticket text for company name patterns and matches against known clients
    in memberberry's exceptions.json
.PARAMETER TicketContent
    The ticket text to search
.EXAMPLE
    $company = Invoke-MemberberryExtractCompany -TicketContent $ticketContent
#>
function Invoke-MemberberryExtractCompany {
    param(
        [Parameter(Mandatory=$true)]
        [string]$TicketContent
    )
    
    $memberberryPath = Get-MemberberryScriptPath
    if (-not $memberberryPath) {
        Write-Warning "Memberberry not found. Cannot extract company name."
        return ""
    }
    
    $extractCompanyScript = Join-Path $memberberryPath "extract-company.ps1"
    if (-not (Test-Path $extractCompanyScript)) {
        Write-Warning "extract-company.ps1 not found at $extractCompanyScript."
        return ""
    }
    
    try {
        # Use PowerShell parameter passing to prevent injection
        $companyName = & $extractCompanyScript -TicketText $TicketContent
        if ($companyName -and $companyName.Trim() -ne "") {
            return $companyName.Trim()
        }
        return ""
    } catch {
        Write-Warning "Error extracting company name with memberberry: $($_.Exception.Message)"
        return ""
    }
}

<#
.SYNOPSIS
    Detects alert types from ticket content using memberberry's detect-alert-type.ps1 script
.DESCRIPTION
    Analyzes ticket text to identify security alert types
.PARAMETER TicketContent
    The ticket content to analyze
.EXAMPLE
    $alertTypes = Invoke-MemberberryDetectAlertType -TicketContent $ticketContent
    # Returns: "suspicious_login,inbox_forwarding" or empty string
#>
function Invoke-MemberberryDetectAlertType {
    param(
        [Parameter(Mandatory=$true)]
        [string]$TicketContent
    )
    
    $memberberryPath = Get-MemberberryScriptPath
    if (-not $memberberryPath) {
        Write-Warning "Memberberry not found. Cannot detect alert types."
        return ""
    }
    
    $detectAlertTypeScript = Join-Path $memberberryPath "detect-alert-type.ps1"
    if (-not (Test-Path $detectAlertTypeScript)) {
        Write-Warning "detect-alert-type.ps1 not found at $detectAlertTypeScript."
        return ""
    }
    
    try {
        # Use PowerShell parameter passing to prevent injection
        $alertTypes = & $detectAlertTypeScript -TicketText $TicketContent
        if ($alertTypes -and $alertTypes.Trim() -ne "") {
            return $alertTypes.Trim()
        }
        return ""
    } catch {
        Write-Warning "Error detecting alert types with memberberry: $($_.Exception.Message)"
        return ""
    }
}

<#
.SYNOPSIS
    Compiles memberberry instructions using compile.ps1 script
.DESCRIPTION
    Merges general_rules.md, procedure files, and exceptions.json into a single
    output file for use with LLM security alert analysis
.PARAMETER Client
    Client name to apply specific exceptions (optional)
.PARAMETER AlertType
    Filter procedures by alert type (comma-separated, optional)
.PARAMETER Output
    Output file path (default: output/memberberry.md relative to memberberry path)
.EXAMPLE
    $outputPath = Invoke-MemberberryCompile -Client "Acme Corp" -AlertType "suspicious_login"
#>
function Invoke-MemberberryCompile {
    param(
        [Parameter(Mandatory=$false)]
        [string]$Client = "",
        
        [Parameter(Mandatory=$false)]
        [string]$AlertType = "",
        
        [Parameter(Mandatory=$false)]
        [string]$Output = ""
    )
    
    $memberberryPath = Get-MemberberryScriptPath
    if (-not $memberberryPath) {
        Write-Warning "Memberberry not found. Cannot compile instructions."
        return ""
    }
    
    $compileScript = Join-Path $memberberryPath "compile.ps1"
    if (-not (Test-Path $compileScript)) {
        Write-Warning "compile.ps1 not found at $compileScript."
        return ""
    }
    
    try {
        # Build parameter hashtable for splatting
        $params = @{}
        if ($Client -and $Client.Trim() -ne "") {
            $params['Client'] = $Client.Trim()
        }
        if ($AlertType -and $AlertType.Trim() -ne "") {
            $params['AlertType'] = $AlertType.Trim()
        }
        if ($Output -and $Output.Trim() -ne "") {
            $params['Output'] = $Output.Trim()
        }
        
        # Use PowerShell parameter passing to prevent injection
        $outputPath = & $compileScript @params
        
        # Return the output path (compile.ps1 returns the path it created)
        if ($outputPath -and (Test-Path $outputPath)) {
            return $outputPath
        }
        
        # Fallback: if no return value, check default output location
        if ([string]::IsNullOrWhiteSpace($Output)) {
            $defaultOutput = Join-Path $memberberryPath "output\memberberry.md"
            if (Test-Path $defaultOutput) {
                return $defaultOutput
            }
        } elseif (Test-Path $Output) {
            return $Output
        }
        
        return ""
    } catch {
        Write-Warning "Error compiling memberberry instructions: $($_.Exception.Message)"
        return ""
    }
}

<#
.SYNOPSIS
    Gets memberberry integration status information
.DESCRIPTION
    Returns information about whether memberberry is available and configured
.EXAMPLE
    $status = Get-MemberberryIntegrationStatus
    Write-Host "Memberberry Path: $($status.MemberberryPath)"
    Write-Host "Available: $($status.IsAvailable)"
#>
function Get-MemberberryIntegrationStatus {
    $memberberryPath = Get-MemberberryScriptPath
    $isAvailable = $false
    
    if ($memberberryPath) {
        $cleanTicketScript = Join-Path $memberberryPath "clean-ticket.ps1"
        $extractCompanyScript = Join-Path $memberberryPath "extract-company.ps1"
        $detectAlertTypeScript = Join-Path $memberberryPath "detect-alert-type.ps1"
        $compileScript = Join-Path $memberberryPath "compile.ps1"
        
        $isAvailable = (Test-Path $cleanTicketScript) -and
                      (Test-Path $extractCompanyScript) -and
                      (Test-Path $detectAlertTypeScript) -and
                      (Test-Path $compileScript)
    }
    
    return @{
        MemberberryPath = $memberberryPath
        IsAvailable = $isAvailable
        CleanTicketAvailable = if ($memberberryPath) { Test-Path (Join-Path $memberberryPath "clean-ticket.ps1") } else { $false }
        ExtractCompanyAvailable = if ($memberberryPath) { Test-Path (Join-Path $memberberryPath "extract-company.ps1") } else { $false }
        DetectAlertTypeAvailable = if ($memberberryPath) { Test-Path (Join-Path $memberberryPath "detect-alert-type.ps1") } else { $false }
        CompileAvailable = if ($memberberryPath) { Test-Path (Join-Path $memberberryPath "compile.ps1") } else { $false }
    }
}

# Export module members
Export-ModuleMember -Function @(
    'Get-MemberberryScriptPath',
    'Invoke-MemberberryCleanTicket',
    'Invoke-MemberberryExtractCompany',
    'Invoke-MemberberryDetectAlertType',
    'Invoke-MemberberryCompile',
    'Get-MemberberryIntegrationStatus'
)
