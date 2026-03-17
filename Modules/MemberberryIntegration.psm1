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
    Runs compile-slim.ps1 to produce GlobalExceptions.txt, ClientExceptions-*.txt, Settings.txt (no procedures).
.DESCRIPTION
    Same output as Memberberry-Slim.ahk: exceptions and settings only, for slim zip.
.PARAMETER Client
    Client name for client-specific exceptions (optional)
.PARAMETER OutputPath
    Directory where output files will be written (required)
.EXAMPLE
    Invoke-MemberberryCompileSlim -Client "Acme Corp" -OutputPath "C:\Reports\Tenant1"
#>
function Invoke-MemberberryCompileSlim {
    param(
        [Parameter(Mandatory=$false)]
        [string]$Client = "",
        [Parameter(Mandatory=$true)]
        [string]$OutputPath
    )
    $memberberryPath = Get-MemberberryScriptPath
    if (-not $memberberryPath) {
        Write-Warning "Memberberry not found. Cannot run compile-slim."
        return $null
    }
    $compileSlimScript = Join-Path $memberberryPath "compile-slim.ps1"
    if (-not (Test-Path $compileSlimScript)) {
        Write-Warning "compile-slim.ps1 not found at $compileSlimScript."
        return $null
    }
    try {
        $params = @{ OutputPath = $OutputPath }
        if ($Client -and $Client.Trim() -ne "") { $params['Client'] = $Client.Trim() }
        & $compileSlimScript @params | Out-Null
        $clientFilenamePath = Join-Path $OutputPath ".slim_client_exceptions_filename.txt"
        $clientFileName = "ClientExceptions-None.txt"
        if (Test-Path $clientFilenamePath) {
            $clientFileName = (Get-Content $clientFilenamePath -Raw -ErrorAction SilentlyContinue).Trim()
            Remove-Item $clientFilenamePath -Force -ErrorAction SilentlyContinue
        }
        return @{
            OutputPath = $OutputPath
            GlobalExceptionsPath = Join-Path $OutputPath "GlobalExceptions.txt"
            ClientExceptionsPath = Join-Path $OutputPath $clientFileName
            SettingsPath = Join-Path $OutputPath "Settings.txt"
            ClientExceptionsFileName = $clientFileName
        }
    } catch {
        Write-Warning "Invoke-MemberberryCompileSlim failed: $($_.Exception.Message)"
        return $null
    }
}

<#
.SYNOPSIS
    Produces Memberberry-Slim-style files by invoking memberberry's create-slim-package.ps1 directly.
.DESCRIPTION
    Calls memberberry's create-slim-package.ps1 (same flow as Memberberry-Slim.ahk).
    Use these files in the zip instead of _AI_Readme.txt.
.PARAMETER TicketContent
    Raw ticket text (ConnectWise clipboard format)
.PARAMETER TicketNumbers
    Array of ticket numbers (e.g. @('1838914'))
.PARAMETER CompanyName
    Optional. If not provided, extracted from ticket via extract-company.ps1
.PARAMETER OutputFolder
    Report output folder where files will be written
.PARAMETER MemberberryPath
    Optional. From Settings.MemberberryPath or Get-MemberberryScriptPath
.OUTPUTS
    Hashtable with paths: AlwaysIncludePath, TicketPath, ClientExceptionsPath, GlobalExceptionsPath, SettingsPath, or $null if memberberry unavailable
#>
function New-MemberberrySlimPackage {
    param(
        [Parameter(Mandatory=$false)]
        [string]$TicketContent = "",
        [Parameter(Mandatory=$false)]
        [string[]]$TicketNumbers = @(),
        [Parameter(Mandatory=$false)]
        [string]$CompanyName = "",
        [Parameter(Mandatory=$true)]
        [string]$OutputFolder,
        [Parameter(Mandatory=$false)]
        [string]$MemberberryPath = ""
    )
    $mbPath = if ($MemberberryPath -and (Test-Path $MemberberryPath)) { $MemberberryPath } else { Get-MemberberryScriptPath }
    if (-not $mbPath) {
        Write-Warning "Memberberry not found. Cannot create slim package."
        return $null
    }
    $createSlimScript = Join-Path $mbPath "create-slim-package.ps1"
    if (-not (Test-Path $createSlimScript)) {
        Write-Warning "create-slim-package.ps1 not found in memberberry. Falling back to local logic."
        return New-MemberberrySlimPackageLocal -TicketContent $TicketContent -TicketNumbers $TicketNumbers -CompanyName $CompanyName -OutputFolder $OutputFolder -MemberberryPath $mbPath
    }
    try {
        $params = @{ TicketContent = $TicketContent; OutputPath = $OutputFolder }
        if ($TicketNumbers -and $TicketNumbers.Count -gt 0) { $params['TicketNumbers'] = $TicketNumbers }
        if ($CompanyName) { $params['CompanyName'] = $CompanyName }
        $json = & $createSlimScript @params 2>$null
        if ([string]::IsNullOrWhiteSpace($json)) { return $null }
        $obj = $json | ConvertFrom-Json
        $result = @{
            AlwaysIncludePath = $obj.AlwaysIncludePath
            TicketPath = $obj.TicketPath
            ClientExceptionsPath = $obj.ClientExceptionsPath
            GlobalExceptionsPath = $obj.GlobalExceptionsPath
            SettingsPath = $obj.SettingsPath
        }
        if ((Test-Path $result.TicketPath)) { return $result }
    } catch {
        Write-Warning "create-slim-package.ps1 failed: $($_.Exception.Message). Falling back to local logic."
    }
    return New-MemberberrySlimPackageLocal -TicketContent $TicketContent -TicketNumbers $TicketNumbers -CompanyName $CompanyName -OutputFolder $OutputFolder -MemberberryPath $mbPath
}

function New-MemberberrySlimPackageLocal {
    param([string]$TicketContent, [string[]]$TicketNumbers, [string]$CompanyName, [string]$OutputFolder, [string]$MemberberryPath)
    $cleanedTicket = $TicketContent
    $detectedClient = $CompanyName
    if (-not [string]::IsNullOrWhiteSpace($TicketContent)) {
        $cleanedTicket = Invoke-MemberberryCleanTicket -TicketContent $TicketContent
        if ([string]::IsNullOrWhiteSpace($detectedClient)) {
            $detectedClient = Invoke-MemberberryExtractCompany -TicketContent $cleanedTicket
        }
    }
    $slimResult = Invoke-MemberberryCompileSlim -Client $detectedClient -OutputPath $OutputFolder
    if (-not $slimResult) { return $null }
    $ticketNumber = ""
    if ($TicketNumbers -and $TicketNumbers.Count -gt 0) { $ticketNumber = $TicketNumbers[0] }
    if ([string]::IsNullOrWhiteSpace($ticketNumber) -and -not [string]::IsNullOrWhiteSpace($cleanedTicket)) {
        if ($cleanedTicket -match "Service Ticket #(\d+)") { $ticketNumber = $Matches[1] }
        elseif ($cleanedTicket -match "Ticket #(\d+)") { $ticketNumber = $Matches[1] }
        elseif ($cleanedTicket -match "#(\d{6,})") { $ticketNumber = $Matches[1] }
    }
    $ticketFileName = if ($ticketNumber) { "Ticket-$ticketNumber.txt" } else { "Ticket-Information.txt" }
    $ticketPath = Join-Path $OutputFolder $ticketFileName
    $separator = "=================================================================================="
    $header = "TICKET INFORMATION"
    if ($ticketNumber) { $header = "$header - Ticket #$ticketNumber" }
    $ticketBody = if ([string]::IsNullOrWhiteSpace($cleanedTicket)) { $TicketContent } else { $cleanedTicket }
    "$header`n`n$separator`n`n$ticketBody" | Out-File -FilePath $ticketPath -Encoding UTF8 -NoNewline
    $alwaysIncludePath = Join-Path $MemberberryPath "always_include.md"
    if (-not (Test-Path $alwaysIncludePath)) { $alwaysIncludePath = Join-Path $OutputFolder "always_include.md" }
    return @{
        AlwaysIncludePath = $alwaysIncludePath
        TicketPath = $ticketPath
        ClientExceptionsPath = $slimResult.ClientExceptionsPath
        GlobalExceptionsPath = $slimResult.GlobalExceptionsPath
        SettingsPath = $slimResult.SettingsPath
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
        $compileSlimScript = Join-Path $memberberryPath "compile-slim.ps1"
        
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
        CompileSlimAvailable = if ($memberberryPath) { Test-Path (Join-Path $memberberryPath "compile-slim.ps1") } else { $false }
        CreateSlimPackageAvailable = if ($memberberryPath) { Test-Path (Join-Path $memberberryPath "create-slim-package.ps1") } else { $false }
    }
}

# Export module members
Export-ModuleMember -Function @(
    'Get-MemberberryScriptPath',
    'Invoke-MemberberryCleanTicket',
    'Invoke-MemberberryExtractCompany',
    'Invoke-MemberberryDetectAlertType',
    'Invoke-MemberberryCompile',
    'Invoke-MemberberryCompileSlim',
    'New-MemberberrySlimPackage',
    'Get-MemberberryIntegrationStatus'
)
