function Get-SecurityInvestigationRoot {
    Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'ExchangeOnlineAnalyzer\SecurityInvestigation'
}

function Get-SecurityIntegrationExportFolder {
    param(
        [Parameter(Mandatory = $true)]
        [string]$CompanyName,
        [string]$OutputFolder = '',
        [string]$TicketNumber = ''
    )

    if (-not [string]::IsNullOrWhiteSpace($OutputFolder) -and (Test-Path -LiteralPath $OutputFolder)) {
        return (Resolve-Path -LiteralPath $OutputFolder).Path
    }

    $safeCompany = ($CompanyName -replace '[\\/:*?"<>|]', '_').Trim()
    if ([string]::IsNullOrWhiteSpace($safeCompany)) { $safeCompany = 'UnknownClient' }

    $root = Get-SecurityInvestigationRoot
    $companyDir = Join-Path $root $safeCompany
    if (-not (Test-Path -LiteralPath $companyDir)) {
        New-Item -ItemType Directory -Path $companyDir -Force | Out-Null
    }

    $existing = Get-ChildItem -LiteralPath $companyDir -Directory -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1
    if ($existing -and ((Get-Date) - $existing.LastWriteTime).TotalHours -lt 12) {
        return $existing.FullName
    }

    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $folder = Join-Path $companyDir $timestamp
    New-Item -ItemType Directory -Path $folder -Force | Out-Null
    return $folder
}

function Export-SecurityIntegrationCsv {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Folder,
        [Parameter(Mandatory = $true)]
        [string]$BaseName,
        [Parameter(Mandatory = $true)]
        [object[]]$Rows,
        [string]$TicketNumber = ''
    )

    if (-not (Test-Path -LiteralPath $Folder)) {
        New-Item -ItemType Directory -Path $Folder -Force | Out-Null
    }

    $suffix = if ([string]::IsNullOrWhiteSpace($TicketNumber)) { '' } else { "_Ticket_$TicketNumber" }
    $path = Join-Path $Folder "$BaseName$suffix.csv"
    $Rows | Export-Csv -LiteralPath $path -NoTypeInformation -Encoding UTF8
    return $path
}

function Export-SecurityIntegrationJson {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Folder,
        [Parameter(Mandatory = $true)]
        [string]$BaseName,
        [Parameter(Mandatory = $true)]
        [object]$Object,
        [string]$TicketNumber = ''
    )

    if (-not (Test-Path -LiteralPath $Folder)) {
        New-Item -ItemType Directory -Path $Folder -Force | Out-Null
    }

    $suffix = if ([string]::IsNullOrWhiteSpace($TicketNumber)) { '' } else { "_Ticket_$TicketNumber" }
    $path = Join-Path $Folder "$BaseName$suffix.json"
    $Object | ConvertTo-Json -Depth 12 | Out-File -LiteralPath $path -Encoding utf8
    return $path
}

function Get-RecommendedSecurityPulls {
    param(
        [string]$AlertTypes = '',
        [hashtable]$HuntressCaps = $null,
        [hashtable]$SentinelOneCaps = $null,
        [string]$SocSource = 'unknown',
        [hashtable]$SecurityStack = $null
    )

    $types = @()
    if (-not [string]::IsNullOrWhiteSpace($AlertTypes)) {
        $types = @($AlertTypes -split '[,\s;]+' | ForEach-Object { $_.Trim().ToLowerInvariant() } | Where-Object { $_ })
    }

    $isEndpoint = $types | Where-Object { $_ -match 'malware|virus|endpoint|foothold|antivirus|process|sentinel|threat|edr' }
    $isIdentity = $types | Where-Object { $_ -match 'identity|m365|bec|inbox|forward|travel|sign.?in|login|itdr|email' }
    $isSiem = $types | Where-Object { $_ -match 'siem|lateral' }

    if ($SecurityStack) {
        if ($SecurityStack.isEndpointAlert) { $isEndpoint = @($isEndpoint) + @('endpoint') }
        if ($SecurityStack.isIdentityAlert) { $isIdentity = @($isIdentity) + @('identity') }
    }

    $wantHuntress = $true
    $wantS1 = $true
    if ($SecurityStack) {
        $wantHuntress = [bool]$SecurityStack.useHuntress
        $wantS1 = [bool]$SecurityStack.useS1ConnectWise -or [bool]$SecurityStack.useS1Barracuda
        if (-not $SecurityStack.useHuntress -and -not $SecurityStack.useS1ConnectWise -and -not $SecurityStack.useS1Barracuda) {
            if ($SecurityStack.isIdentityAlert) { $wantHuntress = [bool]$HuntressCaps.hasHuntress }
            if ($SecurityStack.isEndpointAlert) {
                $wantHuntress = [bool]$HuntressCaps.hasHuntress
                $wantS1 = [bool]$SentinelOneCaps.hasSentinelOne
            }
        }
    }

    $huntress = @{
        signalsFootholds      = $false
        signalsAntivirus      = $false
        signalsProcessInsights = $false
        signalsManagedItdr    = $false
        signalsSiem           = $false
        incidents             = $false
        agents                = $false
        identities            = $false
        escalations           = $false
    }

    $s1 = @{
        threats    = $false
        agents     = $false
        activities = $false
    }

    if ($wantHuntress -and $HuntressCaps -and $HuntressCaps.hasHuntress) {
        if ($isEndpoint -or (-not $types.Count -and -not $SecurityStack.isIdentityAlert)) {
            if ($HuntressCaps.edr -ne $false) {
                $huntress.signalsFootholds = $true
                $huntress.signalsAntivirus = $true
                $huntress.signalsProcessInsights = $true
                $huntress.incidents = $true
                $huntress.agents = $true
            }
        }
        if ($isIdentity) {
            if ($HuntressCaps.itdr -ne $false) {
                $huntress.signalsManagedItdr = $true
                $huntress.incidents = $true
                $huntress.identities = $true
            }
        }
        if ($isSiem -and $HuntressCaps.siem) {
            $huntress.signalsSiem = $true
            $huntress.incidents = $true
        }
        if (-not ($huntress.Values | Where-Object { $_ })) {
            $huntress.escalations = $true
            $huntress.agents = [bool]$HuntressCaps.edr
        }
    }

    if ($wantS1 -and $SentinelOneCaps -and $SentinelOneCaps.hasSentinelOne) {
        if ($isEndpoint -or (-not $types.Count -and -not $SecurityStack.isIdentityAlert)) {
            $s1.threats = $true
            $s1.agents = $true
            $s1.activities = $true
        }
    }

    return @{
        huntress      = $huntress
        sentinelOne   = $s1
        socSource     = $SocSource
        securityStack = $SecurityStack
    }
}

Export-ModuleMember -Function Get-SecurityInvestigationRoot,Get-SecurityIntegrationExportFolder,Export-SecurityIntegrationCsv,Export-SecurityIntegrationJson,Get-RecommendedSecurityPulls
