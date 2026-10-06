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

function Resolve-ConnectWiseManageRedirectUri {
    param(
        [Parameter(Mandatory = $true)][Uri]$RequestUri,
        [Parameter(Mandatory = $true)][string]$Location
    )

    $locText = $Location.Trim()
    if (-not $locText) { throw 'ConnectWise Manage API request failed: redirect had no location.' }
    $locationUri = $null
    if (-not [Uri]::TryCreate($locText, [UriKind]::RelativeOrAbsolute, [ref]$locationUri)) {
        throw 'ConnectWise Manage API request failed: redirect location was not a URL.'
    }
    $next = if ($locationUri.IsAbsoluteUri) { $locationUri } else { [Uri]::new($RequestUri, $locationUri) }
    if ($RequestUri.Scheme -eq 'https' -and $next.Scheme -eq 'http') {
        $builder = [System.UriBuilder]::new($next)
        $builder.Scheme = 'https'
        if ($next.IsDefaultPort) { $builder.Port = -1 }
        $next = $builder.Uri
    }
    return $next
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
    $bodyText = $null
    if ($null -ne $Body) {
        if ($Body -is [string]) { $bodyText = $Body }
        else { $bodyText = ($Body | ConvertTo-Json -Depth 8 -Compress) }
    }

    # Manage often answers a successful POST with a redirect to an http URL.
    # PowerShell refuses that hop, so the note is saved and this client reports failure.
    # Follow the redirect ourselves and keep it on https.
    $handler = [System.Net.Http.HttpClientHandler]::new()
    $handler.AllowAutoRedirect = $false
    $client = [System.Net.Http.HttpClient]::new($handler)
    $client.Timeout = [TimeSpan]::FromSeconds(100)
    $current = $uri
    $verb = $Method
    $payload = $bodyText
    try {
        for ($hop = 0; $hop -lt 5; $hop++) {
            $request = [System.Net.Http.HttpRequestMessage]::new([System.Net.Http.HttpMethod]::new($verb), $current)
            try {
                [void]$request.Headers.TryAddWithoutValidation('Authorization', "Basic $authValue")
                [void]$request.Headers.TryAddWithoutValidation('clientId', $Credentials.ClientId.Trim())
                [void]$request.Headers.TryAddWithoutValidation('Accept', 'application/json')
                if ($null -ne $payload -and $verb -ne 'GET' -and $verb -ne 'DELETE') {
                    $request.Content = [System.Net.Http.StringContent]::new($payload, [Text.Encoding]::UTF8, 'application/json')
                }
                $pending = $client.SendAsync($request)
                if (-not $pending.Wait([TimeSpan]::FromSeconds(100))) {
                    throw 'ConnectWise Manage API request failed: the request timed out.'
                }
                if ($pending.IsFaulted) {
                    throw "ConnectWise Manage API request failed: $($pending.Exception.GetBaseException().Message)"
                }
                $response = $pending.Result
                try {
                    $read = $response.Content.ReadAsStringAsync()
                    [void]$read.Wait()
                    $text = [string]$read.Result
                    $code = [int]$response.StatusCode
                    if ($code -in 301, 302, 303, 307, 308) {
                        $location = ''
                        if ($response.Headers.Location) { $location = [string]$response.Headers.Location }
                        $current = Resolve-ConnectWiseManageRedirectUri -RequestUri $current -Location $location
                        if ($code -in 301, 302, 303) {
                            $verb = 'GET'
                            $payload = $null
                        }
                        continue
                    }
                    if ($code -ge 400) {
                        if ($text.Length -gt 500) { $text = $text.Substring(0, 500) + '...' }
                        throw "ConnectWise Manage API request failed: $code $text"
                    }
                    if ([string]::IsNullOrWhiteSpace($text)) { return $null }
                    $parsed = $text | ConvertFrom-Json
                    if ($parsed -is [System.Array]) { return ,$parsed }
                    return $parsed
                } finally {
                    $response.Dispose()
                }
            } finally {
                $request.Dispose()
            }
        }
        throw 'ConnectWise Manage API request failed: too many redirects.'
    } finally {
        $client.Dispose()
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

    $contactEmail = ''
    if ($ticket.contactEmailAddress) { $contactEmail = [string]$ticket.contactEmailAddress }
    $emailCc = ''
    if ($ticket.automaticEmailCc) { $emailCc = [string]$ticket.automaticEmailCc }

    return [pscustomobject]@{
        Success         = $true
        TicketId        = $id
        Summary         = $summary
        CompanyName     = $companyName
        ContactName     = $contactName
        StatusName      = $statusName
        ContactEmail    = $contactEmail.Trim()
        AutomaticEmailCc = $emailCc.Trim()
        EmailContact    = [bool]$ticket.automaticEmailContactFlag
        EmailCc         = [bool]$ticket.automaticEmailCcFlag
        TicketContent   = $filtered
        TicketNumbers   = $ticketNumbers
        RawLength       = $rawContent.Length
        FilteredLength  = $filtered.Length
    }
}

function Get-ConnectWiseManageHunterCatalog {
    <#
    River Run Manage, confirmed against the live API: Security board 97,
    New - Save 1st to send user Ticket Information 1831,
    Dispatch 1825, In Progress 1829, Review 1841, Waiting on Client 1854,
    Remote Consulting 3, Communication 82.
    Notes and time are entered as member cknospe. Status changes do not set closedFlag.
    #>
    [CmdletBinding()]
    param()
    return [pscustomobject]@{
        BoardId            = 97
        BoardName          = 'Security'
        MemberIdentifier   = 'cknospe'
        MemberId           = 301
        MemberName         = 'Chris Knospe'
        New                = [pscustomobject]@{ Key = 'new'; Id = 1831; Name = 'New - Save 1st to send user Ticket Information' }
        Dispatch           = [pscustomobject]@{ Key = 'dispatch'; Id = 1825; Name = 'Dispatch' }
        InProgress         = [pscustomobject]@{ Key = 'progress'; Id = 1829; Name = 'In Progress' }
        ResourceSchedule   = [pscustomobject]@{ TypeId = 4; StatusId = 2; SpanId = 1; WhereId = 4 }
        Statuses           = @(
            [pscustomobject]@{ Key = 'review'; Id = 1841; Name = 'Review' }
            [pscustomobject]@{ Key = 'waiting'; Id = 1854; Name = 'Waiting on Client' }
        )
        WorkTypes          = @(
            [pscustomobject]@{ Key = 'remote'; Id = 3; Name = 'Remote Consulting'; BillableOption = 'Billable' }
            [pscustomobject]@{ Key = 'communication'; Id = 82; Name = 'Communication'; BillableOption = 'DoNotBill' }
        )
    }
}

function Get-ConnectWiseManageHunterNoteText {
    param($Notes, [string]$Name)
    if (-not $Notes) { return '' }
    $value = $Notes.$Name
    if ($null -eq $value -and $Notes.PSObject) {
        $prop = $Notes.PSObject.Properties[$Name]
        if ($prop) { $value = $prop.Value }
    }
    if ($null -eq $value) { return '' }
    return ([string]$value).Trim()
}

function Resolve-ConnectWiseManageHunterPost {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Request
    )

    $catalog = Get-ConnectWiseManageHunterCatalog
    $ticketId = (([string]$Request.ticketId) -replace '\D', '').Trim()
    if ([string]::IsNullOrWhiteSpace($ticketId)) {
        throw 'Enter a ticket number.'
    }

    $statusKey = ([string]$Request.status).Trim().ToLowerInvariant()
    $keepStatus = $statusKey -eq 'keep'
    if ($keepStatus) {
        $status = [pscustomobject]@{ Id = 0; Name = '' }
    }
    else {
        $status = @($catalog.Statuses | Where-Object { $_.Key -eq $statusKey } | Select-Object -First 1)
        if (-not $status -or -not $status[0]) {
            throw 'Choose Review or Waiting on Client.'
        }
        $status = $status[0]
    }

    $client = Get-ConnectWiseManageHunterNoteText -Notes $Request.notes -Name 'client'
    $internal = Get-ConnectWiseManageHunterNoteText -Notes $Request.notes -Name 'internal'
    $incident = Get-ConnectWiseManageHunterNoteText -Notes $Request.notes -Name 'incident'
    if (-not $client -and -not $internal -and -not $incident) {
        throw 'Choose at least one note to post.'
    }

    $timeKey = ([string]$Request.time).Trim().ToLowerInvariant()
    if ([string]::IsNullOrWhiteSpace($timeKey)) { $timeKey = 'none' }
    $workTypeId = 0
    $billable = $null
    $hours = 0.0
    if ($timeKey -ne 'none') {
        if (-not $client) {
            throw 'Time is added only with the client email.'
        }
        $work = @($catalog.WorkTypes | Where-Object { $_.Key -eq $timeKey } | Select-Object -First 1)
        if (-not $work -or -not $work[0]) {
            throw 'Choose Remote Consulting, Communication, or no time.'
        }
        $parsedHours = 0.0
        $hoursOk = [double]::TryParse(
            [string]$Request.hours,
            [Globalization.NumberStyles]::Float,
            [Globalization.CultureInfo]::InvariantCulture,
            [ref]$parsedHours)
        if (-not $hoursOk -or $parsedHours -le 0 -or $parsedHours -gt 24) {
            throw 'Enter the hours for this time entry (up to 24).'
        }
        $hours = [math]::Round($parsedHours, 2)
        $workTypeId = [int]$work[0].Id
        $billable = [string]$work[0].BillableOption
    }

    $member = ([string]$Request.memberIdentifier).Trim()
    if ([string]::IsNullOrWhiteSpace($member)) { $member = [string]$catalog.MemberIdentifier }

    $notify = $false
    if ($Request.includeRecipients -eq $true -or "$($Request.includeRecipients)" -eq 'true') { $notify = $true }

    return [pscustomobject]@{
        TicketId           = $ticketId
        MemberIdentifier   = $member
        StatusId           = [int]$status.Id
        StatusName         = [string]$status.Name
        KeepStatus         = [bool]$keepStatus
        NotifyClient       = [bool]($notify -and $client)
        NotifyIncident     = [bool]($notify -and $incident)
        ClientText         = $client
        InternalText       = $internal
        IncidentText       = $incident
        Hours              = $hours
        WorkTypeId         = $workTypeId
        BillableOption     = $billable
        TimeMode           = $timeKey
    }
}

function ConvertTo-ConnectWiseManagePriorityColor {
    param([string]$Color, [string]$Name)
    switch (($Color).Trim().ToLowerInvariant()) {
        'black' { return '#111111' }
        'red' { return '#d0021b' }
        'yellow' { return '#f5c400' }
        'lime' { return '#32cd32' }
        'green' { return '#228b22' }
        'blue' { return '#1b6ec2' }
        'orange' { return '#f57500' }
    }
    $parts = @(($Name -split '\s+') | Where-Object { $_ })
    $last = if ($parts.Count) { [string]$parts[-1] } else { '' }
    switch ($last) {
        '1' { return '#111111' }
        '2' { return '#d0021b' }
        '3' { return '#f5c400' }
        '4' { return '#32cd32' }
        default { return '#6c757d' }
    }
}

function Get-ConnectWiseManageHunterStatusNames {
    param($Statuses)
    $pinned = @(
        'New - Save 1st to send user Ticket Information',
        'Dispatch',
        'In Progress',
        'Review',
        'Waiting on Client'
    )
    $source = New-Object System.Collections.Generic.List[string]
    foreach ($item in @($Statuses)) {
        if (-not $item) { continue }
        $name = [string]$item.Name
        if ([string]::IsNullOrWhiteSpace($name)) { $name = [string]$item.name }
        $name = $name.Trim()
        if ([string]::IsNullOrWhiteSpace($name)) { continue }
        if ($name.StartsWith('zClosed', [StringComparison]::OrdinalIgnoreCase)) { continue }
        $inactive = $false
        if ($item.PSObject.Properties['inactive'] -and $null -ne $item.inactive) { $inactive = [bool]$item.inactive }
        elseif ($item.PSObject.Properties['Inactive'] -and $null -ne $item.Inactive) { $inactive = [bool]$item.Inactive }
        if ($inactive) { continue }
        if ($source -notcontains $name) { $source.Add($name) }
    }
    $picked = New-Object System.Collections.Generic.List[string]
    foreach ($pin in $pinned) {
        $match = $source | Where-Object { $_ -eq $pin } | Select-Object -First 1
        if ($match) {
            $picked.Add($match)
            $source.Remove($match) | Out-Null
        }
    }
    foreach ($name in $source) { $picked.Add($name) }
    return @($picked.ToArray())
}

function Get-ConnectWiseManageHunterBoardStatusNames {
    param(
        [object]$Settings = $null,
        [object]$Credentials = $null
    )
    if (-not $Credentials) {
        $Credentials = Get-ConnectWiseManageCredentialsFromSettings -Settings $Settings
    }
    if (-not (Test-ConnectWiseManageCredentialsComplete -Credentials $Credentials)) {
        $status = Get-ConnectWiseManageConfigurationStatus -Settings $Settings
        throw $status.Message
    }
    $boardId = [int](Get-ConnectWiseManageHunterCatalog).BoardId
    $raw = Invoke-ConnectWiseManageApi -Credentials $Credentials -Method GET -RelativePath "service/boards/$boardId/statuses?pageSize=100&fields=id,name,inactive"
    $rows = @()
    if ($raw -is [System.Array]) { $rows = $raw }
    elseif ($null -ne $raw) { $rows = @($raw) }
    return @(Get-ConnectWiseManageHunterStatusNames -Statuses $rows)
}

function Get-ConnectWiseManageHunterBoardStatuses {
    param(
        [object]$Settings = $null,
        [object]$Credentials = $null
    )
    if (-not $Credentials) {
        $Credentials = Get-ConnectWiseManageCredentialsFromSettings -Settings $Settings
    }
    if (-not (Test-ConnectWiseManageCredentialsComplete -Credentials $Credentials)) {
        $status = Get-ConnectWiseManageConfigurationStatus -Settings $Settings
        throw $status.Message
    }
    $boardId = [int](Get-ConnectWiseManageHunterCatalog).BoardId
    $raw = Invoke-ConnectWiseManageApi -Credentials $Credentials -Method GET -RelativePath "service/boards/$boardId/statuses?pageSize=100&fields=id,name,inactive,closedStatus"
    $rows = @()
    if ($raw -is [System.Array]) { $rows = $raw }
    elseif ($null -ne $raw) { $rows = @($raw) }
    $names = @(Get-ConnectWiseManageHunterStatusNames -Statuses $rows)
    $byName = @{}
    foreach ($row in $rows) {
        $name = ([string]$row.name).Trim()
        if ($name -and -not $byName.ContainsKey($name)) { $byName[$name] = $row }
    }
    $ordered = New-Object System.Collections.Generic.List[object]
    foreach ($name in $names) {
        if (-not $byName.ContainsKey($name)) { continue }
        $row = $byName[$name]
        $ordered.Add([pscustomobject]@{
            id     = [int]$row.id
            name   = $name
            closed = [bool]$row.closedStatus
        })
    }
    return @($ordered.ToArray())
}

function Resolve-ConnectWiseManageHunterStatusChange {
    param(
        [Parameter(Mandatory = $true)]
        $Request,
        [Parameter(Mandatory = $true)]
        $Choices
    )
    $catalog = Get-ConnectWiseManageHunterCatalog
    $ticketId = (([string]$Request.ticketId) -replace '\D', '').Trim()
    if ([string]::IsNullOrWhiteSpace($ticketId)) { throw 'Enter a ticket number.' }
    $statusId = 0
    [void][int]::TryParse(([string]$Request.statusId).Trim(), [ref]$statusId)
    $match = @($Choices | Where-Object { $_ -and [int]$_.id -eq $statusId } | Select-Object -First 1)
    if ($statusId -le 0 -or -not $match -or -not $match[0]) {
        throw 'Choose a status from the Security board.'
    }
    $choice = $match[0]
    $closed = $false
    if ($choice.PSObject.Properties['closed'] -and $choice.closed) { $closed = $true }
    $timeKey = 'none'
    $workTypeId = 0
    $billable = $null
    $hours = 0.0
    $start = $null
    $end = $null
    if ($closed) {
        $timeKey = ([string]$Request.time).Trim().ToLowerInvariant()
        if ([string]::IsNullOrWhiteSpace($timeKey)) { $timeKey = 'none' }
    }
    if ($timeKey -ne 'none') {
        $work = @($catalog.WorkTypes | Where-Object { $_.Key -eq $timeKey } | Select-Object -First 1)
        if (-not $work -or -not $work[0]) {
            throw 'Choose Remote Consulting, Communication, or no time.'
        }
        $parsedHours = 0.0
        $hoursOk = [double]::TryParse(
            [string]$Request.hours,
            [Globalization.NumberStyles]::Float,
            [Globalization.CultureInfo]::InvariantCulture,
            [ref]$parsedHours)
        if (-not $hoursOk -or $parsedHours -le 0 -or $parsedHours -gt 24) {
            throw 'Enter the hours for this time entry (up to 24).'
        }
        $hours = [math]::Round($parsedHours, 2)
        $workTypeId = [int]$work[0].Id
        $billable = [string]$work[0].BillableOption
        $parsedStart = [datetimeoffset]::MinValue
        $startText = ([string]$Request.timeStart).Trim()
        if ([string]::IsNullOrWhiteSpace($startText) -or -not [datetimeoffset]::TryParse($startText, [ref]$parsedStart)) {
            throw 'Enter the start time for this time entry.'
        }
        $start = $parsedStart
        $end = $start.AddMinutes([math]::Round($hours * 60))
        if ($end -le $start) { $end = $start.AddMinutes(15) }
    }
    return [pscustomobject]@{
        TicketId         = $ticketId
        MemberIdentifier = [string]$catalog.MemberIdentifier
        StatusId         = $statusId
        StatusName       = [string]$choice.name
        Closed           = $closed
        Hours            = $hours
        WorkTypeId       = $workTypeId
        BillableOption   = $billable
        TimeMode         = $timeKey
        TimeStart        = $start
        TimeEnd          = $end
    }
}

function ConvertTo-ConnectWiseManageApiTime {
    param([Parameter(Mandatory = $true)][datetimeoffset]$Value)
    # Manage rejects offsets such as -05:00 on timeStart and timeEnd. It accepts UTC with a Z.
    return $Value.UtcDateTime.ToString('yyyy-MM-ddTHH:mm:ss', [Globalization.CultureInfo]::InvariantCulture) + 'Z'
}

function Resolve-ConnectWiseManageHunterNoteTime {
    param(
        [Parameter(Mandatory = $true)]
        $Request,
        [Parameter(Mandatory = $true)]
        $Note
    )
    $catalog = Get-ConnectWiseManageHunterCatalog
    $ticketId = (([string]$Request.ticketId) -replace '\D', '').Trim()
    if ([string]::IsNullOrWhiteSpace($ticketId)) { throw 'Enter a ticket number.' }
    $noteId = (([string]$Request.noteId) -replace '\D', '').Trim()
    if ([string]::IsNullOrWhiteSpace($noteId)) { throw 'Choose a discussion note.' }
    if ((ConvertTo-ConnectWiseManageHunterNoteKind $Note) -ne 'discussion') {
        throw 'Time can be added only to a discussion note.'
    }
    $text = ''
    if ($Note.text) { $text = [string]$Note.text }
    elseif ($Note.detailDescription) { $text = [string]$Note.detailDescription }
    $text = $text.Trim()
    if ([string]::IsNullOrWhiteSpace($text)) { throw 'That discussion note has no text.' }
    $timeKey = ([string]$Request.time).Trim().ToLowerInvariant()
    $work = @($catalog.WorkTypes | Where-Object { $_.Key -eq $timeKey } | Select-Object -First 1)
    if (-not $work -or -not $work[0]) {
        throw 'Choose Remote Consulting or Communication.'
    }
    $parsedHours = 0.0
    $hoursOk = [double]::TryParse(
        [string]$Request.hours,
        [Globalization.NumberStyles]::Float,
        [Globalization.CultureInfo]::InvariantCulture,
        [ref]$parsedHours)
    if (-not $hoursOk -or $parsedHours -le 0 -or $parsedHours -gt 24) {
        throw 'Enter the hours for this time entry (up to 24).'
    }
    $hours = [math]::Round($parsedHours, 2)
    $parsedStart = [datetimeoffset]::MinValue
    $startText = ([string]$Request.timeStart).Trim()
    if ([string]::IsNullOrWhiteSpace($startText) -or -not [datetimeoffset]::TryParse($startText, [ref]$parsedStart)) {
        throw 'Enter the start time for this time entry.'
    }
    $end = $parsedStart.AddMinutes([math]::Round($hours * 60))
    if ($end -le $parsedStart) { $end = $parsedStart.AddMinutes(15) }
    return [pscustomobject]@{
        TicketId         = $ticketId
        NoteId           = $noteId
        MemberIdentifier = [string]$catalog.MemberIdentifier
        NoteText         = $text
        Hours            = $hours
        WorkTypeId       = [int]$work[0].Id
        BillableOption   = [string]$work[0].BillableOption
        TimeMode         = $timeKey
        TimeStart        = $parsedStart
        TimeEnd          = $end
    }
}

function Get-ConnectWiseManageHunterTriagedIds {
    param(
        $Credentials,
        [string[]]$TicketIds
    )
    $ids = @($TicketIds | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
    $found = New-Object 'System.Collections.Generic.HashSet[string]'
    if (-not $ids.Count) { return $found }
    $baseUri = Get-ConnectWiseManageApiBaseUri -ApiUrl $Credentials.ApiUrl
    $authBytes = [Text.Encoding]::UTF8.GetBytes("$($Credentials.CompanyId)+$($Credentials.PublicKey):$($Credentials.PrivateKey)")
    $authValue = [Convert]::ToBase64String($authBytes)
    $clientId = $Credentials.ClientId.Trim()
    $cond = [Uri]::EscapeDataString('text contains "HATZ-AI POWERED REVIEW"')
    $handler = [System.Net.Http.HttpClientHandler]::new()
    $handler.AllowAutoRedirect = $false
    $client = [System.Net.Http.HttpClient]::new($handler)
    $client.Timeout = [TimeSpan]::FromSeconds(30)
    try {
        $batchSize = 16
        for ($offset = 0; $offset -lt $ids.Count; $offset += $batchSize) {
            $end = [Math]::Min($offset + $batchSize - 1, $ids.Count - 1)
            $slice = @($ids[$offset..$end])
            $pending = New-Object System.Collections.Generic.List[object]
            foreach ($id in $slice) {
                $uri = [Uri]::new([Uri]$baseUri, "service/tickets/$id/notes/count?conditions=$cond")
                $request = [System.Net.Http.HttpRequestMessage]::new([System.Net.Http.HttpMethod]::Get, $uri)
                [void]$request.Headers.TryAddWithoutValidation('Authorization', "Basic $authValue")
                [void]$request.Headers.TryAddWithoutValidation('clientId', $clientId)
                [void]$request.Headers.TryAddWithoutValidation('Accept', 'application/json')
                $pending.Add([pscustomobject]@{
                    Id      = [string]$id
                    Request = $request
                    Task    = $client.SendAsync($request)
                })
            }
            try { [void][System.Threading.Tasks.Task]::WaitAll(@($pending | ForEach-Object { $_.Task })) } catch {}
            foreach ($item in $pending) {
                try {
                    if (-not $item.Task.IsCompleted -or $item.Task.IsFaulted -or $item.Task.IsCanceled) { continue }
                    $response = $item.Task.Result
                    try {
                        if ([int]$response.StatusCode -ge 400) { continue }
                        $read = $response.Content.ReadAsStringAsync()
                        [void]$read.Wait()
                        if ([string]$read.Result -match '"count"\s*:\s*([1-9]\d*)') { [void]$found.Add($item.Id) }
                    } finally {
                        if ($response) { $response.Dispose() }
                    }
                } catch {}
                finally { $item.Request.Dispose() }
            }
        }
    } finally {
        $client.Dispose()
    }
    return $found
}

function Get-ConnectWiseManageSecurityQueue {
    [CmdletBinding()]
    param(
        [object]$Settings = $null,
        [object]$Credentials = $null
    )

    if (-not $Credentials) {
        $Credentials = Get-ConnectWiseManageCredentialsFromSettings -Settings $Settings
    }
    if (-not (Test-ConnectWiseManageCredentialsComplete -Credentials $Credentials)) {
        $status = Get-ConnectWiseManageConfigurationStatus -Settings $Settings
        throw $status.Message
    }

    $catalog = Get-ConnectWiseManageHunterCatalog
    $boardId = [int]$catalog.BoardId
    $priorityColors = @{}
    $prioritySorts = @{}
    try {
        $rawPri = Invoke-ConnectWiseManageApi -Credentials $Credentials -Method GET -RelativePath 'service/priorities?pageSize=50&fields=id,name,color,sortOrder'
        $priRows = @()
        if ($rawPri -is [System.Array]) { $priRows = $rawPri }
        elseif ($null -ne $rawPri) { $priRows = @($rawPri) }
        foreach ($pri in $priRows) {
            if (-not $pri.id) { continue }
            $priorityColors[[string]$pri.id] = ConvertTo-ConnectWiseManagePriorityColor -Color ([string]$pri.color) -Name ([string]$pri.name)
            $sortValue = 0
            if ($null -ne $pri.sortOrder -and [string]$pri.sortOrder -ne '') { $sortValue = [int]$pri.sortOrder }
            elseif ($null -ne $pri.sort -and [string]$pri.sort -ne '') { $sortValue = [int]$pri.sort }
            $prioritySorts[[string]$pri.id] = $sortValue
        }
    } catch {
        $priorityColors = @{}
    }
    $rows = New-Object System.Collections.Generic.List[object]
    for ($page = 1; $page -le 3; $page++) {
        $cond = [Uri]::EscapeDataString("board/id=$boardId AND closedFlag=false")
        $relative = "service/tickets?conditions=$cond&fields=id,summary,status,company,contact,owner,priority,type,parentTicketId,_info&pageSize=100&page=$page&orderBy=id%20desc"
        $raw = Invoke-ConnectWiseManageApi -Credentials $Credentials -Method GET -RelativePath $relative
        if ($null -eq $raw) {
            $batch = @()
        } elseif ($raw -is [System.Array]) {
            $batch = $raw
        } else {
            $batch = @($raw)
        }
        $batch = @($batch | Where-Object { $_ -and $null -ne $_.id -and -not ($_.id -is [System.Array]) })
        if ($batch.Count -eq 0) { break }
        foreach ($row in $batch) {
            $entered = ''
            if ($row._info -and $row._info.dateEntered) { $entered = [string]$row._info.dateEntered }
            elseif ($row.dateEntered) { $entered = [string]$row.dateEntered }
            $ageDays = $null
            if ($entered) {
                $enteredAt = [datetimeoffset]::MinValue
                if ([datetimeoffset]::TryParse($entered, [ref]$enteredAt)) {
                    $ageDays = [math]::Floor(((Get-Date).ToUniversalTime() - $enteredAt.UtcDateTime).TotalDays)
                    if ($ageDays -lt 0) { $ageDays = 0 }
                }
            }
            $priorityId = if ($row.priority) { [string]$row.priority.id } else { '' }
            $priorityName = if ($row.priority) { [string]$row.priority.name } else { '' }
            $prioritySort = 0
            if ($row.priority -and $null -ne $row.priority.sort -and [string]$row.priority.sort -ne '') {
                $prioritySort = [int]$row.priority.sort
            } elseif ($prioritySorts.ContainsKey($priorityId)) {
                $prioritySort = [int]$prioritySorts[$priorityId]
            }
            $rows.Add([pscustomobject]@{
                id            = [string]$row.id
                summary       = [string]$row.summary
                status        = if ($row.status) { [string]$row.status.name } else { '' }
                company       = if ($row.company) { [string]$row.company.name } else { '' }
                contact       = if ($row.contact) { [string]$row.contact.name } else { '' }
                owner         = if ($row.owner) { [string]$row.owner.identifier } else { '' }
                type          = if ($row.type) { [string]$row.type.name } else { '' }
                parentId      = if ($row.parentTicketId) { [string]$row.parentTicketId } else { '' }
                priority      = $priorityName.Trim()
                priorityColor = if ($priorityColors.ContainsKey($priorityId)) { [string]$priorityColors[$priorityId] } else { ConvertTo-ConnectWiseManagePriorityColor -Color '' -Name $priorityName }
                prioritySort  = $prioritySort
                entered       = $entered
                ageDays       = $ageDays
            })
        }
        if ($batch.Count -lt 100) { break }
    }
    $triagedIds = $null
    try {
        $triagedIds = Get-ConnectWiseManageHunterTriagedIds -Credentials $Credentials -TicketIds @($rows | ForEach-Object { [string]$_.id })
    } catch {
        $triagedIds = $null
    }
    foreach ($row in $rows) {
        $marked = $false
        if ($triagedIds) { $marked = $triagedIds.Contains([string]$row.id) }
        $row | Add-Member -NotePropertyName triaged -NotePropertyValue $marked -Force
    }
    return $rows.ToArray()
}

function ConvertTo-ConnectWiseManageHunterNoteKind {
    param($Note)
    if ($Note.resolutionFlag -eq $true) { return 'resolution' }
    if ($Note.internalAnalysisFlag -eq $true) { return 'internal' }
    if ($Note.detailDescriptionFlag -eq $true) { return 'discussion' }
    return 'other'
}

function Get-ConnectWiseManageHunterAuditPath {
    param([string]$Path)
    if (-not [string]::IsNullOrWhiteSpace($Path)) { return $Path }
    return Join-Path $env:APPDATA 'ExchangeOnlineAnalyzer\hunter-manage-log.jsonl'
}

function Get-ConnectWiseManageHunterAuditAction {
    param([string]$Summary)
    if ($Summary -match '^(Status|Owner|Resource|Internal note|Client email|Time entry|Incident report)\b') {
        return $Matches[1]
    }
    return 'write'
}

function Add-ConnectWiseManageHunterAudit {
    param(
        [string]$TicketId,
        [string]$Action,
        [string]$Outcome,
        [string]$Summary,
        [string]$Member = 'cknospe',
        [string]$Path
    )
    try {
        $file = Get-ConnectWiseManageHunterAuditPath -Path $Path
        $dir = Split-Path -Parent $file
        if ($dir -and -not (Test-Path -LiteralPath $dir)) {
            New-Item -ItemType Directory -Path $dir -Force | Out-Null
        }
        $entry = [ordered]@{
            at       = [datetimeoffset]::UtcNow.ToString('o')
            ticketId = [string]$TicketId
            member   = [string]$Member
            action   = [string]$Action
            outcome  = [string]$Outcome
            summary  = [string]$Summary
        }
        Add-Content -LiteralPath $file -Value ($entry | ConvertTo-Json -Compress) -Encoding utf8
    } catch {
        Write-Warning "Could not write the hunter Manage log: $($_.Exception.Message)"
    }
}

function Write-ConnectWiseManageHunterAuditResult {
    param(
        [string]$TicketId,
        [string]$Member,
        $Completed,
        $Errors
    )
    foreach ($line in @($Completed)) {
        $text = [string]$line
        if ([string]::IsNullOrWhiteSpace($text)) { continue }
        Add-ConnectWiseManageHunterAudit -TicketId $TicketId -Member $Member -Outcome 'ok' -Summary $text -Action (Get-ConnectWiseManageHunterAuditAction $text)
    }
    foreach ($line in @($Errors)) {
        $text = [string]$line
        if ([string]::IsNullOrWhiteSpace($text)) { continue }
        Add-ConnectWiseManageHunterAudit -TicketId $TicketId -Member $Member -Outcome 'failed' -Summary $text -Action (Get-ConnectWiseManageHunterAuditAction $text)
    }
}

function Get-ConnectWiseManageHunterAudit {
    param(
        [string]$TicketId,
        [int]$Limit = 40,
        [string]$Path
    )
    $file = Get-ConnectWiseManageHunterAuditPath -Path $Path
    if (-not (Test-Path -LiteralPath $file)) { return @() }
    $lines = @(Get-Content -LiteralPath $file -Tail 400 -ErrorAction SilentlyContinue)
    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($line in $lines) {
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        try { $row = $line | ConvertFrom-Json } catch { continue }
        if ($TicketId -and [string]$row.ticketId -ne [string]$TicketId) { continue }
        $rows.Add($row)
    }
    $ordered = @($rows | Sort-Object { [string]$_.at } -Descending)
    if ($Limit -gt 0 -and $ordered.Count -gt $Limit) {
        $ordered = @($ordered | Select-Object -First $Limit)
    }
    return @($ordered)
}

function Get-ConnectWiseManageHunterTicketThread {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TicketId,
        $Credentials
    )
    $id = ($TicketId -replace '\D', '').Trim()
    if ([string]::IsNullOrWhiteSpace($id)) { throw 'Enter a ticket number.' }
    $ticket = Invoke-ConnectWiseManageApi -Credentials $Credentials -Method GET -RelativePath "service/tickets/${id}?fields=id,initialDescription,contact"
    if ($ticket -is [System.Array]) { $ticket = $ticket | Select-Object -First 1 }
    $items = New-Object System.Collections.Generic.List[object]
    $initial = ''
    if ($ticket -and $ticket.initialDescription) { $initial = ([string]$ticket.initialDescription).Trim() }
    if ($initial) {
        $who = ''
        if ($ticket.contact -and $ticket.contact.name) { $who = [string]$ticket.contact.name }
        $items.Add([pscustomobject]@{
            kind = 'description'
            when = ''
            who  = $who
            text = $initial
        })
    }
    $raw = Invoke-ConnectWiseManageApi -Credentials $Credentials -Method GET -RelativePath "service/tickets/$id/notes?pageSize=250&orderBy=dateCreated%20asc"
    $notes = @()
    if ($raw -is [System.Array]) { $notes = $raw }
    elseif ($null -ne $raw) { $notes = @($raw) }
    foreach ($note in $notes) {
        $text = ''
        if ($note.text) { $text = [string]$note.text }
        elseif ($note.detailDescription) { $text = [string]$note.detailDescription }
        $text = $text.Trim()
        if (-not $text) { continue }
        $who = 'Unknown'
        if ($note.member -and $note.member.name) { $who = [string]$note.member.name }
        elseif ($note.createdBy) { $who = [string]$note.createdBy }
        elseif ($note.contact -and $note.contact.name) { $who = [string]$note.contact.name }
        $items.Add([pscustomobject]@{
            id   = if ($note.id) { [string]$note.id } else { '' }
            kind = ConvertTo-ConnectWiseManageHunterNoteKind $note
            when = if ($note.dateCreated) { [string]$note.dateCreated } else { '' }
            who  = $who
            text = $text
        })
    }
    return @($items.ToArray())
}

function Get-ConnectWiseManageReturnedId {
    param($Value)
    if ($null -eq $Value) { return $null }
    if ($Value -is [System.Array]) { $Value = $Value | Select-Object -First 1 }
    if ($null -eq $Value -or $null -eq $Value.id) { return $null }
    return $Value.id
}

function Test-ConnectWiseManageRedirectFailure {
    param([string]$Message)
    return $Message -match 'insecure redirection|AllowInsecureRedirect'
}

function Find-ConnectWiseManageHunterNoteJustPosted {
    param(
        $Credentials,
        [string]$TicketId,
        [string]$Text,
        [bool]$Discussion
    )
    $wanted = $Text.Trim()
    if (-not $wanted) { return $null }
    $raw = Invoke-ConnectWiseManageApi -Credentials $Credentials -Method GET -RelativePath "service/tickets/$TicketId/notes?pageSize=15&orderBy=id%20desc"
    $notes = @()
    if ($raw -is [System.Array]) { $notes = $raw }
    elseif ($null -ne $raw) { $notes = @($raw) }
    foreach ($note in $notes) {
        $noteText = ''
        if ($note.text) { $noteText = [string]$note.text }
        elseif ($note.detailDescription) { $noteText = [string]$note.detailDescription }
        if ($noteText.Trim() -ne $wanted) { continue }
        $isDiscussion = $false
        if ($null -ne $note.detailDescriptionFlag) { $isDiscussion = [bool]$note.detailDescriptionFlag }
        if ($isDiscussion -ne $Discussion) { continue }
        return $note
    }
    return $null
}

function Send-ConnectWiseManageHunterNote {
    param(
        $Credentials,
        [string]$TicketId,
        [string]$Text,
        [bool]$Discussion,
        [bool]$Notify,
        [string]$Member
    )
    $body = @{
        text                    = $Text
        detailDescriptionFlag   = [bool]$Discussion
        internalAnalysisFlag    = (-not $Discussion)
        internalFlag            = $false
        resolutionFlag          = $false
        processNotifications    = [bool]$Notify
        member                  = @{ identifier = $Member }
    }
    try {
        return Invoke-ConnectWiseManageApi -Credentials $Credentials -Method POST -RelativePath "service/tickets/$TicketId/notes" -Body $body
    } catch {
        $msg = $_.Exception.Message
        if (-not (Test-ConnectWiseManageRedirectFailure $msg)) { throw }
        # Manage saves the note, then redirects http. The client used to report failure
        # and skip time and status even though the note was already on the ticket.
        $found = Find-ConnectWiseManageHunterNoteJustPosted -Credentials $Credentials -TicketId $TicketId -Text $Text -Discussion:$Discussion
        if ($found) { return $found }
        throw
    }
}

function Set-ConnectWiseManageTicketStatus {
    param(
        $Credentials,
        [string]$TicketId,
        [int]$StatusId,
        [string]$StatusName
    )

    $patch = '[{"op":"replace","path":"/status/id","value":' + $StatusId + '}]'
    $lastError = $null
    foreach ($attempt in 1, 2) {
        try {
            Invoke-ConnectWiseManageApi -Credentials $Credentials -Method PATCH -RelativePath "service/tickets/$TicketId" -Body $patch | Out-Null
            $lastError = $null
        } catch {
            $lastError = $_.Exception.Message
        }
        $check = Invoke-ConnectWiseManageApi -Credentials $Credentials -Method GET -RelativePath "service/tickets/${TicketId}?fields=id,status"
        if ($check -is [System.Array]) { $check = $check | Select-Object -First 1 }
        $got = 0
        $gotName = ''
        if ($check -and $check.status) {
            $got = [int]$check.status.id
            $gotName = [string]$check.status.name
        }
        if ($got -eq $StatusId) { return }
        if (-not $lastError) {
            $lastError = "Manage left the status at $(if ($gotName) { $gotName } else { 'unknown' })."
        }
        if ($attempt -eq 1) { Start-Sleep -Seconds 1 }
    }
    throw "Could not set status to ${StatusName}. $lastError"
}

function Set-ConnectWiseManageHunterProgress {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$TicketId,
        [object]$Settings = $null,
        [object]$Credentials = $null
    )

    $id = ($TicketId -replace '\D', '').Trim()
    if ([string]::IsNullOrWhiteSpace($id)) { throw 'Enter a ticket number.' }
    $catalog = Get-ConnectWiseManageHunterCatalog
    $progress = $catalog.InProgress
    $member = [string]$catalog.MemberIdentifier
    if (-not $Credentials) {
        $Credentials = Get-ConnectWiseManageCredentialsFromSettings -Settings $Settings
    }
    if (-not (Test-ConnectWiseManageCredentialsComplete -Credentials $Credentials)) {
        $status = Get-ConnectWiseManageConfigurationStatus -Settings $Settings
        throw $status.Message
    }

    $completed = New-Object System.Collections.Generic.List[string]
    $errors = New-Object System.Collections.Generic.List[string]
    $statusUpdated = $false

    try {
        Set-ConnectWiseManageTicketStatus -Credentials $Credentials -TicketId $id -StatusId ([int]$progress.Id) -StatusName ([string]$progress.Name)
        $statusUpdated = $true
        $completed.Add("Status set to $($progress.Name).")
    } catch {
        $errors.Add("Status: $($_.Exception.Message)")
    }

    try {
        $ownerPatch = '[{"op":"replace","path":"/owner/id","value":' + [int]$catalog.MemberId + '}]'
        Invoke-ConnectWiseManageApi -Credentials $Credentials -Method PATCH -RelativePath "service/tickets/$id" -Body $ownerPatch | Out-Null
        $completed.Add("Owner set to $member.")
    } catch {
        $errors.Add("Owner: $($_.Exception.Message)")
    }

    try {
        $schedule = $catalog.ResourceSchedule
        $conditions = "objectId=$id and member/identifier=`"$member`" and doneFlag=false and type/id=$($schedule.TypeId)"
        $existingPath = 'schedule/entries?conditions=' + [uri]::EscapeDataString($conditions) + '&fields=id,ownerFlag&pageSize=10'
        $raw = Invoke-ConnectWiseManageApi -Credentials $Credentials -Method GET -RelativePath $existingPath
        $rows = @()
        if ($raw -is [System.Array]) { $rows = $raw }
        elseif ($null -ne $raw) { $rows = @($raw) }
        $owned = @($rows | Where-Object { $_.ownerFlag -eq $true })
        if ($owned.Count -gt 0) {
            $completed.Add('Already assigned as a resource.')
        } elseif ($rows.Count -gt 0) {
            $entryPatch = '[{"op":"replace","path":"/ownerFlag","value":true}]'
            Invoke-ConnectWiseManageApi -Credentials $Credentials -Method PATCH -RelativePath "schedule/entries/$($rows[0].id)" -Body $entryPatch | Out-Null
            $completed.Add('Marked the existing schedule entry as the owner resource.')
        } else {
            $body = @{
                objectId                  = [int]$id
                member                    = @{ id = [int]$catalog.MemberId }
                where                     = @{ id = [int]$schedule.WhereId }
                status                    = @{ id = [int]$schedule.StatusId }
                type                      = @{ id = [int]$schedule.TypeId }
                span                      = @{ id = [int]$schedule.SpanId }
                ownerFlag                 = $true
                doneFlag                  = $false
                hours                     = 0
                allowScheduleConflictsFlag = $true
            }
            Invoke-ConnectWiseManageApi -Credentials $Credentials -Method POST -RelativePath 'schedule/entries' -Body $body | Out-Null
            $completed.Add('Added as a resource.')
        }
    } catch {
        $errors.Add("Resource: $($_.Exception.Message)")
    }

    Write-ConnectWiseManageHunterAuditResult -TicketId $id -Member $member -Completed $completed -Errors $errors
    return [pscustomobject]@{
        ok               = ($errors.Count -eq 0)
        ticketId         = $id
        memberIdentifier = $member
        statusName       = [string]$progress.Name
        statusUpdated    = $statusUpdated
        notes            = @{ client = $null; internal = $null; incident = $null }
        timeEntryId      = $null
        completed        = @($completed.ToArray())
        errors           = @($errors.ToArray())
    }
}

function Set-ConnectWiseManageHunterDispatch {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$TicketId,
        [object]$Settings = $null,
        [object]$Credentials = $null
    )

    $id = ($TicketId -replace '\D', '').Trim()
    if ([string]::IsNullOrWhiteSpace($id)) { throw 'Enter a ticket number.' }
    $catalog = Get-ConnectWiseManageHunterCatalog
    $dispatch = $catalog.Dispatch
    $member = [string]$catalog.MemberIdentifier
    if (-not $Credentials) {
        $Credentials = Get-ConnectWiseManageCredentialsFromSettings -Settings $Settings
    }
    if (-not (Test-ConnectWiseManageCredentialsComplete -Credentials $Credentials)) {
        $status = Get-ConnectWiseManageConfigurationStatus -Settings $Settings
        throw $status.Message
    }

    $completed = New-Object System.Collections.Generic.List[string]
    $errors = New-Object System.Collections.Generic.List[string]
    $statusUpdated = $false

    try {
        Set-ConnectWiseManageTicketStatus -Credentials $Credentials -TicketId $id -StatusId ([int]$dispatch.Id) -StatusName ([string]$dispatch.Name)
        $statusUpdated = $true
        $completed.Add("Status set to $($dispatch.Name).")
    } catch {
        $errors.Add("Status: $($_.Exception.Message)")
    }

    $ownerError = $null
    foreach ($body in @(
            '[{"op":"replace","path":"/owner","value":null}]',
            '[{"op":"remove","path":"/owner","value":null}]'
        )) {
        try {
            Invoke-ConnectWiseManageApi -Credentials $Credentials -Method PATCH -RelativePath "service/tickets/$id" -Body $body | Out-Null
        } catch {
            $ownerError = $_.Exception.Message
        }
        $check = Invoke-ConnectWiseManageApi -Credentials $Credentials -Method GET -RelativePath "service/tickets/${id}?fields=id,owner"
        if (-not $check.owner.id) { break }
    }
    $check = Invoke-ConnectWiseManageApi -Credentials $Credentials -Method GET -RelativePath "service/tickets/${id}?fields=id,owner"
    if ($check.owner.id) {
        $detail = if ($ownerError) { $ownerError } else { 'Manage left the owner in place.' }
        $errors.Add("Owner: $detail")
    } else {
        $completed.Add('Owner removed.')
    }

    try {
        $schedule = $catalog.ResourceSchedule
        $conditions = "objectId=$id and member/identifier=`"$member`" and doneFlag=false and type/id=$($schedule.TypeId)"
        $existingPath = 'schedule/entries?conditions=' + [uri]::EscapeDataString($conditions) + '&fields=id&pageSize=25'
        $raw = Invoke-ConnectWiseManageApi -Credentials $Credentials -Method GET -RelativePath $existingPath
        $rows = @()
        if ($raw -is [System.Array]) { $rows = $raw }
        elseif ($null -ne $raw) { $rows = @($raw) }
        $removed = 0
        $resourceErrors = 0
        foreach ($row in $rows) {
            if (-not $row.id) { continue }
            try {
                Invoke-ConnectWiseManageApi -Credentials $Credentials -Method DELETE -RelativePath "schedule/entries/$($row.id)" | Out-Null
                $removed++
            } catch {
                $resourceErrors++
                $errors.Add("Resource $($row.id): $($_.Exception.Message)")
            }
        }
        if ($removed -gt 0) {
            $completed.Add("Removed $removed resource assignment(s).")
        } elseif ($resourceErrors -eq 0) {
            $completed.Add('No open resource assignment to remove.')
        }
    } catch {
        $errors.Add("Resource: $($_.Exception.Message)")
    }

    Write-ConnectWiseManageHunterAuditResult -TicketId $id -Member $member -Completed $completed -Errors $errors
    return [pscustomobject]@{
        ok               = ($errors.Count -eq 0)
        ticketId         = $id
        memberIdentifier = $member
        statusName       = [string]$dispatch.Name
        statusUpdated    = $statusUpdated
        notes            = @{ client = $null; internal = $null; incident = $null }
        timeEntryId      = $null
        completed        = @($completed.ToArray())
        errors           = @($errors.ToArray())
    }
}

function Send-ConnectWiseManageHunterPost {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Request,
        [object]$Settings = $null,
        [object]$Credentials = $null
    )

    $statusKey = ([string]$Request.status).Trim().ToLowerInvariant()
    if ($statusKey -eq 'board-statuses') {
        $choices = @(Get-ConnectWiseManageHunterBoardStatuses -Settings $Settings -Credentials $Credentials)
        if ($choices.Count -eq 1 -and $choices[0] -is [System.Array]) { $choices = @($choices[0]) }
        $names = @($choices | ForEach-Object { [string]$_.name })
        $choiceRows = @($choices | ForEach-Object { @{ id = [int]$_.id; name = [string]$_.name; closed = [bool]$_.closed } })
        return [pscustomobject]@{
            ok               = $true
            ticketId         = $null
            memberIdentifier = (Get-ConnectWiseManageHunterCatalog).MemberIdentifier
            statusName       = $null
            statusUpdated    = $false
            notes            = @{ statuses = $names; choices = $choiceRows }
            timeEntryId      = $null
            completed        = @()
            errors           = @()
        }
    }
    if ($statusKey -eq 'ticket-thread' -or $statusKey -eq 'manage-log') {
        $id = ([string]$Request.ticketId -replace '\D', '').Trim()
        if ([string]::IsNullOrWhiteSpace($id)) { throw 'Enter a ticket number.' }
        $thread = @()
        $threadErrors = New-Object System.Collections.Generic.List[string]
        if ($statusKey -eq 'ticket-thread') {
            if (-not $Credentials) {
                $Credentials = Get-ConnectWiseManageCredentialsFromSettings -Settings $Settings
            }
            if (-not (Test-ConnectWiseManageCredentialsComplete -Credentials $Credentials)) {
                $status = Get-ConnectWiseManageConfigurationStatus -Settings $Settings
                throw $status.Message
            }
            try {
                $thread = @(Get-ConnectWiseManageHunterTicketThread -TicketId $id -Credentials $Credentials)
                if ($thread.Count -eq 1 -and $thread[0] -is [System.Array]) { $thread = @($thread[0]) }
            } catch {
                $threadErrors.Add($_.Exception.Message)
            }
        }
        $audit = @(Get-ConnectWiseManageHunterAudit -TicketId $id)
        if ($audit.Count -eq 1 -and $audit[0] -is [System.Array]) { $audit = @($audit[0]) }
        return [pscustomobject]@{
            ok               = ($threadErrors.Count -eq 0)
            ticketId         = $id
            memberIdentifier = (Get-ConnectWiseManageHunterCatalog).MemberIdentifier
            statusName       = $null
            statusUpdated    = $false
            notes            = @{ thread = @($thread); audit = @($audit) }
            timeEntryId      = $null
            completed        = @()
            errors           = @($threadErrors.ToArray())
        }
    }
    if ($statusKey -eq 'delete-note') {
        if (-not $Credentials) {
            $Credentials = Get-ConnectWiseManageCredentialsFromSettings -Settings $Settings
        }
        if (-not (Test-ConnectWiseManageCredentialsComplete -Credentials $Credentials)) {
            $status = Get-ConnectWiseManageConfigurationStatus -Settings $Settings
            throw $status.Message
        }
        $ticketId = (([string]$Request.ticketId) -replace '\D', '').Trim()
        $noteId = (([string]$Request.noteId) -replace '\D', '').Trim()
        if ([string]::IsNullOrWhiteSpace($ticketId)) { throw 'Enter a ticket number.' }
        if ([string]::IsNullOrWhiteSpace($noteId)) { throw 'Choose a note to delete.' }
        $member = [string](Get-ConnectWiseManageHunterCatalog).MemberIdentifier
        $completed = New-Object System.Collections.Generic.List[string]
        $errors = New-Object System.Collections.Generic.List[string]
        try {
            $existing = Invoke-ConnectWiseManageApi -Credentials $Credentials -Method GET -RelativePath "service/tickets/$ticketId/notes/$noteId"
            if ($existing -is [System.Array]) { $existing = $existing | Select-Object -First 1 }
            if (-not $existing) { throw 'That note was not found.' }
            $kind = ConvertTo-ConnectWiseManageHunterNoteKind $existing
            Invoke-ConnectWiseManageApi -Credentials $Credentials -Method DELETE -RelativePath "service/tickets/$ticketId/notes/$noteId" | Out-Null
            $completed.Add("Deleted note ${noteId} ($kind).")
        } catch {
            $errors.Add("Delete note: $($_.Exception.Message)")
        }
        Write-ConnectWiseManageHunterAuditResult -TicketId $ticketId -Member $member -Completed $completed -Errors $errors
        return [pscustomobject]@{
            ok               = ($errors.Count -eq 0)
            ticketId         = $ticketId
            memberIdentifier = $member
            statusName       = $null
            statusUpdated    = $false
            notes            = @{ client = $null; internal = $null; incident = $null }
            timeEntryId      = $null
            completed        = @($completed.ToArray())
            errors           = @($errors.ToArray())
        }
    }
    if ($statusKey -eq 'add-time') {
        if (-not $Credentials) {
            $Credentials = Get-ConnectWiseManageCredentialsFromSettings -Settings $Settings
        }
        if (-not (Test-ConnectWiseManageCredentialsComplete -Credentials $Credentials)) {
            $status = Get-ConnectWiseManageConfigurationStatus -Settings $Settings
            throw $status.Message
        }
        $ticketId = (([string]$Request.ticketId) -replace '\D', '').Trim()
        $noteId = (([string]$Request.noteId) -replace '\D', '').Trim()
        $rawNote = Invoke-ConnectWiseManageApi -Credentials $Credentials -Method GET -RelativePath "service/tickets/$ticketId/notes/$noteId"
        if ($rawNote -is [System.Array]) { $rawNote = $rawNote | Select-Object -First 1 }
        if (-not $rawNote) { throw 'That discussion note was not found.' }
        $plan = Resolve-ConnectWiseManageHunterNoteTime -Request $Request -Note $rawNote
        $completed = New-Object System.Collections.Generic.List[string]
        $errors = New-Object System.Collections.Generic.List[string]
        $timeEntryId = $null
        try {
            $timeBody = @{
                chargeToId                 = [int]$plan.TicketId
                chargeToType               = 'ServiceTicket'
                member                     = @{ identifier = $plan.MemberIdentifier }
                timeStart                  = ConvertTo-ConnectWiseManageApiTime $plan.TimeStart
                timeEnd                    = ConvertTo-ConnectWiseManageApiTime $plan.TimeEnd
                workType                   = @{ id = [int]$plan.WorkTypeId }
                billableOption             = [string]$plan.BillableOption
                notes                      = $plan.NoteText
                addToDetailDescriptionFlag = $false
            }
            $entry = Invoke-ConnectWiseManageApi -Credentials $Credentials -Method POST -RelativePath 'time/entries' -Body $timeBody
            $timeEntryId = $entry.id
            $workName = if ($plan.TimeMode -eq 'remote') { 'Remote Consulting' } else { 'Communication' }
            $completed.Add("Time entry: $workName, $($plan.Hours) h, on discussion note $($plan.NoteId).")
        } catch {
            $errors.Add("Time entry: $($_.Exception.Message)")
        }
        Write-ConnectWiseManageHunterAuditResult -TicketId $plan.TicketId -Member $plan.MemberIdentifier -Completed $completed -Errors $errors
        return [pscustomobject]@{
            ok               = ($errors.Count -eq 0)
            ticketId         = $plan.TicketId
            memberIdentifier = $plan.MemberIdentifier
            statusName       = $null
            statusUpdated    = $false
            notes            = @{ client = $null; internal = $null; incident = $null }
            timeEntryId      = $timeEntryId
            completed        = @($completed.ToArray())
            errors           = @($errors.ToArray())
        }
    }
    if ($statusKey -eq 'set-status') {
        if (-not $Credentials) {
            $Credentials = Get-ConnectWiseManageCredentialsFromSettings -Settings $Settings
        }
        if (-not (Test-ConnectWiseManageCredentialsComplete -Credentials $Credentials)) {
            $status = Get-ConnectWiseManageConfigurationStatus -Settings $Settings
            throw $status.Message
        }
        $choices = @(Get-ConnectWiseManageHunterBoardStatuses -Settings $Settings -Credentials $Credentials)
        if ($choices.Count -eq 1 -and $choices[0] -is [System.Array]) { $choices = @($choices[0]) }
        $plan = Resolve-ConnectWiseManageHunterStatusChange -Request $Request -Choices $choices
        $completed = New-Object System.Collections.Generic.List[string]
        $errors = New-Object System.Collections.Generic.List[string]
        $timeEntryId = $null
        $statusUpdated = $false
        if ($plan.TimeMode -ne 'none') {
            try {
                $timeBody = @{
                    chargeToId                 = [int]$plan.TicketId
                    chargeToType               = 'ServiceTicket'
                    member                     = @{ identifier = $plan.MemberIdentifier }
                    timeStart                  = ConvertTo-ConnectWiseManageApiTime $plan.TimeStart
                    timeEnd                    = ConvertTo-ConnectWiseManageApiTime $plan.TimeEnd
                    workType                   = @{ id = [int]$plan.WorkTypeId }
                    billableOption             = [string]$plan.BillableOption
                    notes                      = 'Time entered while closing the ticket.'
                    addToDetailDescriptionFlag = $false
                }
                $entry = Invoke-ConnectWiseManageApi -Credentials $Credentials -Method POST -RelativePath 'time/entries' -Body $timeBody
                $timeEntryId = $entry.id
                $workName = if ($plan.TimeMode -eq 'remote') { 'Remote Consulting' } else { 'Communication' }
                $completed.Add("Added $workName time ($($plan.Hours) h) starting $($plan.TimeStart.ToString('g')).")
            } catch {
                $errors.Add("Time entry: $($_.Exception.Message)")
            }
        }
        if ($errors.Count -eq 0) {
            try {
                Set-ConnectWiseManageTicketStatus -Credentials $Credentials -TicketId $plan.TicketId -StatusId ([int]$plan.StatusId) -StatusName ([string]$plan.StatusName)
                $statusUpdated = $true
                if ($plan.Closed) { $completed.Add("Closed the ticket as $($plan.StatusName).") }
                else { $completed.Add("Status set to $($plan.StatusName).") }
            } catch {
                $errors.Add("Status: $($_.Exception.Message)")
            }
        }
        Write-ConnectWiseManageHunterAuditResult -TicketId $plan.TicketId -Member $plan.MemberIdentifier -Completed $completed -Errors $errors
        return [pscustomobject]@{
            ok               = ($errors.Count -eq 0)
            ticketId         = $plan.TicketId
            memberIdentifier = $plan.MemberIdentifier
            statusName       = [string]$plan.StatusName
            statusUpdated    = $statusUpdated
            notes            = @{ client = $null; internal = $null; incident = $null }
            timeEntryId      = $timeEntryId
            completed        = @($completed.ToArray())
            errors           = @($errors.ToArray())
        }
    }
    if ($statusKey -eq 'progress') {
        return Set-ConnectWiseManageHunterProgress -TicketId ([string]$Request.ticketId) -Settings $Settings -Credentials $Credentials
    }
    if ($statusKey -eq 'dispatch') {
        return Set-ConnectWiseManageHunterDispatch -TicketId ([string]$Request.ticketId) -Settings $Settings -Credentials $Credentials
    }

    $plan = Resolve-ConnectWiseManageHunterPost -Request $Request
    if (-not $Credentials) {
        $Credentials = Get-ConnectWiseManageCredentialsFromSettings -Settings $Settings
    }
    if (-not (Test-ConnectWiseManageCredentialsComplete -Credentials $Credentials)) {
        $status = Get-ConnectWiseManageConfigurationStatus -Settings $Settings
        throw $status.Message
    }

    $completed = New-Object System.Collections.Generic.List[string]
    $errors = New-Object System.Collections.Generic.List[string]
    $noteIds = @{ client = $null; internal = $null; incident = $null }
    $timeEntryId = $null
    $statusUpdated = $false

    if ($plan.InternalText) {
        try {
            $note = Send-ConnectWiseManageHunterNote -Credentials $Credentials -TicketId $plan.TicketId -Text $plan.InternalText -Discussion:$false -Notify:$false -Member $plan.MemberIdentifier
            $completed.Add('Posted the internal note.')
            $noteIds.internal = Get-ConnectWiseManageReturnedId $note
        } catch {
            $errors.Add("Internal note: $($_.Exception.Message)")
        }
    }

    $clientPosted = $false
    if ($plan.ClientText) {
        try {
            $note = Send-ConnectWiseManageHunterNote -Credentials $Credentials -TicketId $plan.TicketId -Text $plan.ClientText -Discussion:$true -Notify:$plan.NotifyClient -Member $plan.MemberIdentifier
            $clientPosted = $true
            if ($plan.NotifyClient) { $completed.Add('Posted the client email and sent it to the ticket recipients.') }
            else { $completed.Add('Posted the client email without emailing recipients.') }
            $noteIds.client = Get-ConnectWiseManageReturnedId $note
        } catch {
            $errors.Add("Client email: $($_.Exception.Message)")
        }
    }

    if ($plan.TimeMode -ne 'none') {
        if (-not $clientPosted) {
            $errors.Add('Time entry: skipped because the client email was not posted.')
        } else {
            try {
                $end = [datetimeoffset]::Now
                $start = $end.AddMinutes(-1 * [math]::Round($plan.Hours * 60))
                if ($start -ge $end) { $start = $end.AddMinutes(-15) }
                $timeBody = @{
                    chargeToId                 = [int]$plan.TicketId
                    chargeToType               = 'ServiceTicket'
                    member                     = @{ identifier = $plan.MemberIdentifier }
                    timeStart                  = ConvertTo-ConnectWiseManageApiTime $start
                    timeEnd                    = ConvertTo-ConnectWiseManageApiTime $end
                    workType                   = @{ id = [int]$plan.WorkTypeId }
                    billableOption             = [string]$plan.BillableOption
                    notes                      = $plan.ClientText
                    addToDetailDescriptionFlag = $false
                }
                $entry = Invoke-ConnectWiseManageApi -Credentials $Credentials -Method POST -RelativePath 'time/entries' -Body $timeBody
                $timeEntryId = $entry.id
                $workName = if ($plan.TimeMode -eq 'remote') { 'Remote Consulting' } else { 'Communication' }
                $completed.Add("Added $workName time ($($plan.Hours) h) on the client email.")
            } catch {
                $errors.Add("Time entry: $($_.Exception.Message)")
            }
        }
    }

    if ($plan.IncidentText) {
        try {
            $note = Send-ConnectWiseManageHunterNote -Credentials $Credentials -TicketId $plan.TicketId -Text $plan.IncidentText -Discussion:$true -Notify:$plan.NotifyIncident -Member $plan.MemberIdentifier
            if ($plan.NotifyIncident) { $completed.Add('Posted the incident report and sent it to the ticket recipients.') }
            else { $completed.Add('Posted the incident report without emailing recipients.') }
            $noteIds.incident = Get-ConnectWiseManageReturnedId $note
        } catch {
            $errors.Add("Incident report: $($_.Exception.Message)")
        }
    }

    $anyNote = $completed.Count -gt 0
    if ($plan.KeepStatus) {
        if ($anyNote) { $completed.Add('Status was left unchanged.') }
    }
    elseif ($anyNote) {
        try {
            Set-ConnectWiseManageTicketStatus -Credentials $Credentials -TicketId $plan.TicketId -StatusId ([int]$plan.StatusId) -StatusName ([string]$plan.StatusName)
            $statusUpdated = $true
            $completed.Add("Status set to $($plan.StatusName).")
        } catch {
            $errors.Add("Status: $($_.Exception.Message)")
        }
    } else {
        $errors.Add('Status was not changed because no note was posted.')
    }

    Write-ConnectWiseManageHunterAuditResult -TicketId $plan.TicketId -Member $plan.MemberIdentifier -Completed $completed -Errors $errors
    return [pscustomobject]@{
        ok               = ($errors.Count -eq 0)
        ticketId         = $plan.TicketId
        memberIdentifier = $plan.MemberIdentifier
        statusName       = $plan.StatusName
        statusUpdated    = $statusUpdated
        notes            = $noteIds
        timeEntryId      = $timeEntryId
        completed        = @($completed.ToArray())
        errors           = @($errors.ToArray())
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
    Get-ConnectWiseManageServiceTicketText, `
    Get-ConnectWiseManageHunterCatalog, `
    Get-ConnectWiseManageHunterStatusNames, `
    ConvertTo-ConnectWiseManageHunterNoteKind, `
    ConvertTo-ConnectWiseManageApiTime, `
    Add-ConnectWiseManageHunterAudit, `
    Get-ConnectWiseManageHunterAudit, `
    Resolve-ConnectWiseManageHunterPost, `
    Resolve-ConnectWiseManageHunterStatusChange, `
    Resolve-ConnectWiseManageHunterNoteTime, `
    Get-ConnectWiseManageSecurityQueue, `
    Send-ConnectWiseManageHunterPost
