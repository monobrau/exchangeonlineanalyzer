function Get-VScanMagicConfigRoot {
    Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'VScanMagic'
}

function ConvertTo-VScanMagicSafeUserId {
    param([string]$UserId)

    $value = if ([string]::IsNullOrWhiteSpace($UserId)) { 'default' } else { $UserId.Trim() }
    foreach ($invalid in [IO.Path]::GetInvalidFileNameChars()) {
        $value = $value.Replace([string]$invalid, '_')
    }
    $value = $value.Trim().TrimEnd('.')
    if ([string]::IsNullOrWhiteSpace($value)) { return 'default' }
    return $value
}

function Get-VScanMagicManageCredentialCandidatePaths {
    $root = Get-VScanMagicConfigRoot
    $paths = [System.Collections.Generic.List[string]]::new()

    $userName = ConvertTo-VScanMagicSafeUserId -UserId $env:USERNAME
    $paths.Add((Join-Path $root "users\$userName\ConnectWise-Manage-Credentials.json"))
    $paths.Add((Join-Path $root 'users\default\ConnectWise-Manage-Credentials.json'))

    $usersDir = Join-Path $root 'users'
    if (Test-Path -LiteralPath $usersDir) {
        Get-ChildItem -LiteralPath $usersDir -Directory -ErrorAction SilentlyContinue | ForEach-Object {
            $path = Join-Path $_.FullName 'ConnectWise-Manage-Credentials.json'
            if (-not $paths.Contains($path)) { [void]$paths.Add($path) }
        }
    }

    $legacyPath = Join-Path $root 'ConnectWise-Manage-Credentials.json'
    if (-not $paths.Contains($legacyPath)) { [void]$paths.Add($legacyPath) }

    return $paths
}

function Read-ConnectWiseManageCredentialsFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $raw = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
    $creds = $raw | ConvertFrom-Json -ErrorAction Stop
    return [pscustomobject]@{
        ApiUrl     = [string]$creds.ApiUrl
        CompanyId  = [string]$creds.CompanyId
        PublicKey  = [string]$creds.PublicKey
        PrivateKey = [string]$creds.PrivateKey
        ClientId   = [string]$creds.ClientId
        Source     = 'VScanMagic'
        FilePath   = $Path
    }
}

function Get-VScanMagicManageCredentialsPath {
    foreach ($path in (Get-VScanMagicManageCredentialCandidatePaths)) {
        if (-not (Test-Path -LiteralPath $path)) { continue }
        try {
            $creds = Read-ConnectWiseManageCredentialsFile -Path $path
            if (Test-ConnectWiseManageCredentialsComplete -Credentials $creds) {
                return $path
            }
        } catch {}
    }

    foreach ($path in (Get-VScanMagicManageCredentialCandidatePaths)) {
        if (Test-Path -LiteralPath $path) { return $path }
    }
    return $null
}

function Import-ConnectWiseManageCredentialsFromVScan {
    [CmdletBinding()]
    param(
        [switch]$Quiet
    )

    foreach ($path in (Get-VScanMagicManageCredentialCandidatePaths)) {
        if (-not (Test-Path -LiteralPath $path)) { continue }
        try {
            $creds = Read-ConnectWiseManageCredentialsFile -Path $path
            if (Test-ConnectWiseManageCredentialsComplete -Credentials $creds) {
                return $creds
            }
        } catch {
            if (-not $Quiet) { Write-Warning "Could not read VScanMagic Manage credentials from ${path}: $($_.Exception.Message)" }
        }
    }

    if (-not $Quiet) { Write-Verbose 'VScanMagic Manage credentials file not found or incomplete.' }
    return $null
}

function Get-ConnectWiseManageCredentialsFromSettings {
    [CmdletBinding()]
    param(
        [object]$Settings = $null
    )

    if (-not $Settings) {
        if (Get-Command Get-AppSettings -ErrorAction SilentlyContinue) {
            $Settings = Get-AppSettings
        } else {
            return $null
        }
    }

    $fromEoa = [pscustomobject]@{
        ApiUrl     = [string]$Settings.ManageApiUrl
        CompanyId  = [string]$Settings.ManageCompanyId
        PublicKey  = [string]$Settings.ManagePublicKey
        PrivateKey = [string]$Settings.ManagePrivateKey
        ClientId   = [string]$Settings.ManageClientId
        Source     = 'EOA'
    }

    if (Test-ConnectWiseManageCredentialsComplete -Credentials $fromEoa) {
        return $fromEoa
    }

    if ($Settings.ManagePreferVScanCredentials -eq $true -or -not (Test-ConnectWiseManageCredentialsPartial -Credentials $fromEoa)) {
        $fromVScan = Import-ConnectWiseManageCredentialsFromVScan -Quiet
        if ($fromVScan -and (Test-ConnectWiseManageCredentialsComplete -Credentials $fromVScan)) {
            return $fromVScan
        }
    }

    if (Test-ConnectWiseManageCredentialsPartial -Credentials $fromEoa) {
        return $fromEoa
    }

    $fromVScan = Import-ConnectWiseManageCredentialsFromVScan -Quiet
    if ($fromVScan) { return $fromVScan }

    return $fromEoa
}

function Test-ConnectWiseManageCredentialsPartial {
    param([Parameter(Mandatory = $true)]$Credentials)
    foreach ($name in @('ApiUrl', 'CompanyId', 'PublicKey', 'PrivateKey', 'ClientId')) {
        if (-not [string]::IsNullOrWhiteSpace($Credentials.$name)) { return $true }
    }
    return $false
}

function Test-ConnectWiseManageCredentialsComplete {
    param([Parameter(Mandatory = $true)]$Credentials)
    foreach ($name in @('ApiUrl', 'CompanyId', 'PublicKey', 'PrivateKey', 'ClientId')) {
        if ([string]::IsNullOrWhiteSpace($Credentials.$name)) { return $false }
    }
    return $true
}

function Get-ConnectWiseManageApiBaseUri {
    param([Parameter(Mandatory = $true)][string]$ApiUrl)

    $baseUrl = $ApiUrl.Trim().TrimEnd('/')
    if ([string]::IsNullOrWhiteSpace($baseUrl)) {
        throw 'ConnectWise Manage API URL is required.'
    }
    if ($baseUrl -notmatch '/apis/3\.0$') {
        $baseUrl += '/apis/3.0'
    }
    return ($baseUrl + '/')
}

function Invoke-ConnectWiseManageApi {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Credentials,

        [Parameter(Mandatory = $true)]
        [ValidateSet('GET', 'POST', 'PATCH', 'PUT', 'DELETE')]
        [string]$Method,

        [Parameter(Mandatory = $true)]
        [string]$RelativePath,

        [object]$Body = $null
    )

    if (-not (Test-ConnectWiseManageCredentialsComplete -Credentials $Credentials)) {
        throw 'ConnectWise Manage is not fully configured. Set all fields on the Settings tab (or enable VScanMagic import).'
    }

    $baseUri = Get-ConnectWiseManageApiBaseUri -ApiUrl $Credentials.ApiUrl
    $relative = $RelativePath.TrimStart('/')
    $uri = [Uri]::new([Uri]$baseUri, $relative)

    $authBytes = [Text.Encoding]::UTF8.GetBytes("$($Credentials.CompanyId)+$($Credentials.PublicKey):$($Credentials.PrivateKey)")
    $authValue = [Convert]::ToBase64String($authBytes)

    $headers = @{
        Authorization = "Basic $authValue"
        clientId      = $Credentials.ClientId.Trim()
        Accept        = 'application/json'
    }

    $params = @{
        Method      = $Method
        Uri         = $uri.AbsoluteUri
        Headers     = $headers
        ErrorAction = 'Stop'
    }
    if ($null -ne $Body) {
        $params['Body'] = ($Body | ConvertTo-Json -Depth 8 -Compress)
        $params['ContentType'] = 'application/json'
    }

    try {
        return Invoke-RestMethod @params
    } catch {
        $msg = $_.Exception.Message
        if ($_.ErrorDetails -and $_.ErrorDetails.Message) {
            $detail = $_.ErrorDetails.Message
            if ($detail.Length -gt 500) { $detail = $detail.Substring(0, 500) + '...' }
            $msg = "$msg - $detail"
        }
        throw "ConnectWise Manage API request failed: $msg"
    }
}

function Test-ConnectWiseManageConnection {
    [CmdletBinding()]
    param(
        [object]$Settings = $null,
        [object]$Credentials = $null
    )

    if (-not $Credentials) {
        $Credentials = Get-ConnectWiseManageCredentialsFromSettings -Settings $Settings
    }
    if (-not (Test-ConnectWiseManageCredentialsComplete -Credentials $Credentials)) {
        return [pscustomobject]@{
            Success = $false
            Message = 'Missing one or more Manage fields: API URL, Company ID (auth), Public key, Private key, Client ID header.'
            Source  = $Credentials.Source
        }
    }

    try {
        $boards = Invoke-ConnectWiseManageApi -Credentials $Credentials -Method GET -RelativePath 'service/boards?pageSize=1'
        $count = @($boards).Count
        $sourceNote = if ($Credentials.Source) { " (credentials from $($Credentials.Source))" } else { '' }
        if ($count -eq 0) {
            return [pscustomobject]@{
                Success = $true
                Message = "Connected, but no service boards were returned.$sourceNote"
                Source  = $Credentials.Source
            }
        }
        return [pscustomobject]@{
            Success = $true
            Message = "Connected. $count service board(s) returned on test query.$sourceNote"
            Source  = $Credentials.Source
        }
    } catch {
        return [pscustomobject]@{
            Success = $false
            Message = $_.Exception.Message
            Source  = $Credentials.Source
        }
    }
}

function Get-ConnectWiseManageConfigurationStatus {
    [CmdletBinding()]
    param(
        [object]$Settings = $null
    )

    $vscanPath = Get-VScanMagicManageCredentialsPath
    $vscanExists = [bool]$vscanPath
    $vscanCreds = $null
    $vscanComplete = $false
    if ($vscanExists) {
        $vscanCreds = Import-ConnectWiseManageCredentialsFromVScan -Quiet
        $vscanComplete = $vscanCreds -and (Test-ConnectWiseManageCredentialsComplete -Credentials $vscanCreds)
    }

    $eoaCreds = Get-ConnectWiseManageCredentialsFromSettings -Settings $Settings
    $eoaComplete = $eoaCreds -and (Test-ConnectWiseManageCredentialsComplete -Credentials $eoaCreds)

    $settingsPath = $null
    if (Get-Command Get-SettingsPath -ErrorAction SilentlyContinue) {
        try { $settingsPath = Get-SettingsPath } catch {}
    }

    if ($eoaComplete) {
        return [pscustomobject]@{
            Configured    = $true
            Source        = $eoaCreds.Source
            Message       = "Manage credentials ready (source: $($eoaCreds.Source))."
            VScanFilePath = $vscanPath
            VScanComplete = $vscanComplete
            EoaComplete   = $true
            SettingsPath  = $settingsPath
        }
    }

    $parts = @(
        'ConnectWise Manage is not fully configured.',
        'In Exchange Online Analyzer Settings (ConnectWise Manage tab), set API URL, Company ID (auth), Public key, Private key, and Client ID header, then Save Settings.',
        'Or configure ConnectWise Manage in VScanMagic (stored per user under %LocalAppData%\VScanMagic\users\{username}\ConnectWise-Manage-Credentials.json) and enable Prefer VScanMagic in Settings.'
    )
    if ($vscanExists -and -not $vscanComplete) {
        $parts += "VScanMagic file found but fields are empty or incomplete: $vscanPath"
    }
    elseif (-not $vscanExists) {
        $userPath = Join-Path (Get-VScanMagicConfigRoot) "users\$(ConvertTo-VScanMagicSafeUserId -UserId $env:USERNAME)\ConnectWise-Manage-Credentials.json"
        $parts += "No complete VScanMagic Manage credentials found (checked per-user path: $userPath and legacy root file)."
    }
    if ($settingsPath) {
        $parts += "EOA settings: $settingsPath"
    }

    return [pscustomobject]@{
        Configured    = $false
        Source        = $null
        Message       = ($parts -join ' ')
        VScanFilePath = $vscanPath
        VScanComplete = $vscanComplete
        EoaComplete   = $false
        SettingsPath  = $settingsPath
    }
}

function Get-ConnectWiseManageServiceTicketText {
    <#
    .SYNOPSIS
        Fetches a ConnectWise Manage service ticket and formats it as text for EOA ticket parsing/export.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$TicketId,

        [object]$Settings = $null,
        [object]$Credentials = $null,
        [switch]$IncludeNotes = $true
    )

    $id = ($TicketId -replace '\D', '').Trim()
    if ([string]::IsNullOrWhiteSpace($id)) {
        throw 'Ticket ID must contain a numeric service ticket number.'
    }

    if (-not $Credentials) {
        $Credentials = Get-ConnectWiseManageCredentialsFromSettings -Settings $Settings
    }
    if (-not (Test-ConnectWiseManageCredentialsComplete -Credentials $Credentials)) {
        $status = Get-ConnectWiseManageConfigurationStatus -Settings $Settings
        throw $status.Message
    }

    $ticket = Invoke-ConnectWiseManageApi -Credentials $Credentials -Method GET -RelativePath "service/tickets/$id"
    if (-not $ticket) {
        throw "Ticket $id was not found in ConnectWise Manage."
    }

    $summary = [string]$ticket.summary
    $companyName = $null
    if ($ticket.company -and $ticket.company.name) { $companyName = [string]$ticket.company.name }
    $contactName = $null
    if ($ticket.contact -and $ticket.contact.name) { $contactName = [string]$ticket.contact.name }
    $statusName = $null
    if ($ticket.status -and $ticket.status.name) { $statusName = [string]$ticket.status.name }
    $boardName = $null
    if ($ticket.board -and $ticket.board.name) { $boardName = [string]$ticket.board.name }
    $priorityName = $null
    if ($ticket.priority -and $ticket.priority.name) { $priorityName = [string]$ticket.priority.name }

    $lines = [System.Collections.Generic.List[string]]::new()
    [void]$lines.Add("Service Ticket #$id")
    if ($summary) { [void]$lines.Add("Summary: $summary") }
    if ($companyName) { [void]$lines.Add("Company: $companyName") }
    if ($contactName) { [void]$lines.Add("Contact: $contactName") }
    if ($statusName) { [void]$lines.Add("Status: $statusName") }
    if ($boardName) { [void]$lines.Add("Board: $boardName") }
    if ($priorityName) { [void]$lines.Add("Priority: $priorityName") }
    [void]$lines.Add('')

    $initial = [string]$ticket.initialDescription
    if (-not [string]::IsNullOrWhiteSpace($initial)) {
        [void]$lines.Add('Initial Description:')
        [void]$lines.Add($initial.Trim())
        [void]$lines.Add('')
    }

    if ($IncludeNotes) {
        try {
            $notes = Invoke-ConnectWiseManageApi -Credentials $Credentials -Method GET -RelativePath "service/tickets/$id/notes?pageSize=250&orderBy=dateCreated asc"
            $noteList = @($notes)
            if ($noteList.Count -gt 0) {
                [void]$lines.Add('--- Discussion ---')
                foreach ($note in $noteList) {
                    $when = if ($note.dateCreated) { [string]$note.dateCreated } else { '' }
                    $who = if ($note.member -and $note.member.name) { [string]$note.member.name } elseif ($note.createdBy) { [string]$note.createdBy } else { 'Unknown' }
                    $text = if ($note.text) { [string]$note.text } elseif ($note.detailDescription) { [string]$note.detailDescription } else { '' }
                    if ([string]::IsNullOrWhiteSpace($text)) { continue }
                    [void]$lines.Add("")
                    if ($when) { [void]$lines.Add("[$when] $who") } else { [void]$lines.Add($who) }
                    [void]$lines.Add($text.Trim())
                }
            }
        } catch {
            Write-Warning "Could not load ticket notes for #$id : $($_.Exception.Message)"
        }
    }

    $rawContent = ($lines -join "`r`n").Trim()
    $filtered = $rawContent
    if (Get-Command Filter-TicketContent -ErrorAction SilentlyContinue) {
        try { $filtered = Filter-TicketContent -TicketContent $rawContent } catch { $filtered = $rawContent }
    }

    $ticketNumbers = @($id)
    if (Get-Command Extract-TicketNumbers -ErrorAction SilentlyContinue) {
        try {
            $extracted = @(Extract-TicketNumbers -TicketContent $filtered)
            if ($extracted.Count -gt 0) { $ticketNumbers = $extracted }
        } catch {}
    }

    return [pscustomobject]@{
        Success       = $true
        TicketId      = $id
        Summary       = $summary
        CompanyName   = $companyName
        TicketContent = $filtered
        TicketNumbers = $ticketNumbers
        RawLength     = $rawContent.Length
        FilteredLength = $filtered.Length
    }
}

Export-ModuleMember -Function `
    Get-VScanMagicManageCredentialsPath, `
    Import-ConnectWiseManageCredentialsFromVScan, `
    Get-ConnectWiseManageCredentialsFromSettings, `
    Test-ConnectWiseManageCredentialsComplete, `
    Test-ConnectWiseManageCredentialsPartial, `
    Get-ConnectWiseManageApiBaseUri, `
    Invoke-ConnectWiseManageApi, `
    Test-ConnectWiseManageConnection, `
    Get-ConnectWiseManageConfigurationStatus, `
    Get-ConnectWiseManageServiceTicketText
