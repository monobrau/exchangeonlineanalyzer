function Get-BulkRunnerDefaultSettings {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProjectRoot
    )

    $investigator = 'Security Administrator'
    $company = 'Organization'
    $settingsPath = Join-Path $ProjectRoot 'settings.json'
    if (Test-Path -LiteralPath $settingsPath) {
        try {
            $settings = Get-Content -LiteralPath $settingsPath -Raw | ConvertFrom-Json
            if ($settings.InvestigatorName) { $investigator = [string]$settings.InvestigatorName }
            if ($settings.CompanyName) { $company = [string]$settings.CompanyName }
        } catch { }
    }
    return [pscustomobject]@{ InvestigatorName = $investigator; CompanyName = $company }
}

function New-BulkRunnerSession {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProjectRoot,

        [Parameter(Mandatory = $true)]
        [hashtable]$ReportSelections,

        [string]$InvestigatorName = '',
        [string]$CompanyName = '',
        [int]$DaysBack = 7
    )

    $workerTemplate = Join-Path $ProjectRoot 'Scripts\BulkExportWorker.ps1'
    if (-not (Test-Path -LiteralPath $workerTemplate)) {
        throw "Missing worker script: $workerTemplate"
    }

    $tempDir = Join-Path $env:TEMP "EOA_BulkWeb_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
    $null = New-Item -ItemType Directory -Path $tempDir -Force -ErrorAction Stop
    $commandDir = Join-Path $tempDir 'commands'
    $null = New-Item -ItemType Directory -Path $commandDir -Force -ErrorAction Stop

    $reportSelectionsFile = Join-Path $tempDir 'ReportSelections.json'
    $ReportSelections | ConvertTo-Json -Depth 6 | Set-Content -Path $reportSelectionsFile -Encoding UTF8

    $workerScriptFile = Join-Path $tempDir 'BulkTenantWorker.ps1'
    Copy-Item -LiteralPath $workerTemplate -Destination $workerScriptFile -Force

    return [pscustomobject]@{
        SessionId            = [System.IO.Path]::GetFileName($tempDir)
        ProjectRoot          = $ProjectRoot
        TempDir              = $tempDir
        CommandDir           = $commandDir
        ReportSelectionsFile = $reportSelectionsFile
        WorkerScriptFile     = $workerScriptFile
        InvestigatorName     = $InvestigatorName
        CompanyName          = $CompanyName
        DaysBack             = $DaysBack
        Tenants              = @{}
        NextClientNumber     = 1
        CreatedAt            = Get-Date
    }
}

function Add-BulkRunnerTenant {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Session
    )

    $clientNumber = $Session.NextClientNumber

    $statusFile = Join-Path $Session.TempDir "Client${clientNumber}_Status.txt"
    $resultFile = Join-Path $Session.TempDir "Client${clientNumber}_Result.txt"

    $investigatorName = if ([string]::IsNullOrWhiteSpace($Session.InvestigatorName)) { 'Security Administrator' } else { $Session.InvestigatorName.Trim() }
    $companyName = if ([string]::IsNullOrWhiteSpace($Session.CompanyName)) { 'Organization' } else { $Session.CompanyName.Trim() }

    $psExe = (Get-Process -Id $PID -ErrorAction Stop).Path
    $argList = @(
        '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $Session.WorkerScriptFile,
        '-ClientNumber', [string]$clientNumber,
        '-ScriptRoot', $Session.ProjectRoot,
        '-InvestigatorName', $investigatorName,
        '-CompanyName', $companyName,
        '-DaysBack', [string]$Session.DaysBack,
        '-ReportSelectionsFile', $Session.ReportSelectionsFile,
        '-StatusFile', $statusFile,
        '-ResultFile', $resultFile,
        '-CommandDir', $Session.CommandDir
    )

    $proc = Start-Process -FilePath $psExe -ArgumentList $argList -PassThru -WindowStyle Normal `
        -WorkingDirectory $Session.ProjectRoot -ErrorAction Stop

    $Session.NextClientNumber = $clientNumber + 1

    $tenant = [pscustomobject]@{
        ClientNumber          = $clientNumber
        ProcessId             = $proc.Id
        StatusFile            = $statusFile
        ResultFile            = $resultFile
        GraphAuthenticated    = $false
        ExchangeAuthenticated = $false
        LastResponse          = $null
    }
    $Session.Tenants[[string]$clientNumber] = $tenant
    return $tenant
}

function Send-BulkRunnerCommand {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Session,

        [Parameter(Mandatory = $true)]
        [int]$ClientNumber,

        [Parameter(Mandatory = $true)]
        [string]$Command,

        [int]$TimeoutSeconds = 60
    )

    if (-not $Session.Tenants.ContainsKey([string]$ClientNumber)) {
        throw "Unknown tenant client number: $ClientNumber"
    }

    $commandFile = Join-Path $Session.CommandDir "Client${ClientNumber}_Command.txt"
    $responseFile = Join-Path $Session.CommandDir "Client${ClientNumber}_Response.txt"

    if (Test-Path $responseFile) {
        Remove-Item $responseFile -Force -ErrorAction SilentlyContinue
        Start-Sleep -Milliseconds 100
    }

    $utf8NoBom = New-Object System.Text.UTF8Encoding $false
    [System.IO.File]::WriteAllText($commandFile, $Command, $utf8NoBom)

    $startTime = Get-Date
    while (((Get-Date) - $startTime).TotalSeconds -lt $TimeoutSeconds) {
        if (Test-Path $responseFile) {
            Start-Sleep -Milliseconds 200
            $response = (Get-Content $responseFile -Raw -ErrorAction SilentlyContinue)
            if ($response) {
                $response = $response.Trim()
                $Session.Tenants[[string]$ClientNumber].LastResponse = $response
                if ($response -like 'GRAPH_AUTH_SUCCESS*') {
                    $Session.Tenants[[string]$ClientNumber].GraphAuthenticated = $true
                }
                if ($response -eq 'EXCHANGE_AUTH_SUCCESS' -or $response -like 'EXCHANGE_AUTH_SUCCESS*') {
                    $Session.Tenants[[string]$ClientNumber].ExchangeAuthenticated = $true
                }
                return $response
            }
        }
        Start-Sleep -Milliseconds 200
    }

    return $null
}

function Get-BulkRunnerTenantStatus {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Session,

        [Parameter(Mandatory = $true)]
        [int]$ClientNumber,

        [int]$TailLines = 50
    )

    if (-not $Session.Tenants.ContainsKey([string]$ClientNumber)) {
        throw "Unknown tenant client number: $ClientNumber"
    }

    $statusFile = $Session.Tenants[[string]$ClientNumber].StatusFile
    if (-not (Test-Path $statusFile)) {
        return ''
    }

    $lines = Get-Content $statusFile -ErrorAction SilentlyContinue
    if (-not $lines) { return '' }
    if ($lines.Count -le $TailLines) {
        return ($lines -join "`n")
    }
    return ($lines[-$TailLines..-1] -join "`n")
}

function Get-BulkRunnerAppRegistrations {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProjectRoot,

        [switch]$ForceRefreshFromGraph
    )

    Import-Module (Join-Path $ProjectRoot 'Modules\GraphAppCredential.psm1') -Force -ErrorAction Stop
    $list = @()
    if (Get-Command Get-WCMTenantListWithNamesForAppRegCombo -ErrorAction SilentlyContinue) {
        if ($ForceRefreshFromGraph) {
            $list = Get-WCMTenantListWithNamesForAppRegCombo -ForceRefreshFromGraph
        } else {
            $list = Get-WCMTenantListWithNamesForAppRegCombo
        }
    } elseif (Get-Command Get-WCMTenantListWithNames -ErrorAction SilentlyContinue) {
        if ($ForceRefreshFromGraph) {
            $list = Get-WCMTenantListWithNames -ForceRefreshFromGraph
        } else {
            $list = Get-WCMTenantListWithNames
        }
    }

    return @($list | ForEach-Object {
        [pscustomobject]@{
            displayText = $_.DisplayText
            tenantId    = $_.TenantId
        }
    })
}

function Get-BulkRunnerSessionSummary {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Session
    )

    $tenants = @($Session.Tenants.Values | Sort-Object ClientNumber | ForEach-Object {
        [pscustomobject]@{
            clientNumber          = $_.ClientNumber
            processId             = $_.ProcessId
            graphAuthenticated    = $_.GraphAuthenticated
            exchangeAuthenticated = $_.ExchangeAuthenticated
            lastResponse          = $_.LastResponse
        }
    })

    return [pscustomobject]@{
        sessionId    = $Session.SessionId
        tempDir      = $Session.TempDir
        createdAt    = $Session.CreatedAt
        tenantCount  = $tenants.Count
        tenants      = $tenants
    }
}

Export-ModuleMember -Function New-BulkRunnerSession, Add-BulkRunnerTenant, Send-BulkRunnerCommand, Get-BulkRunnerTenantStatus, Get-BulkRunnerAppRegistrations, Get-BulkRunnerSessionSummary, Get-BulkRunnerDefaultSettings
