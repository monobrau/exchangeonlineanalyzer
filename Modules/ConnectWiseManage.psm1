function Get-VScanMagicManageCredentialsPath {
    $path = Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'VScanMagic\ConnectWise-Manage-Credentials.json'
    if (Test-Path -LiteralPath $path) { return $path }
    return $null
}

function Import-ConnectWiseManageCredentialsFromVScan {
    [CmdletBinding()]
    param(
        [switch]$Quiet
    )

    $path = Get-VScanMagicManageCredentialsPath
    if (-not $path) {
        if (-not $Quiet) { Write-Verbose 'VScanMagic Manage credentials file not found.' }
        return $null
    }

    try {
        $raw = Get-Content -LiteralPath $path -Raw -ErrorAction Stop
        $creds = $raw | ConvertFrom-Json -ErrorAction Stop
        return [pscustomobject]@{
            ApiUrl     = [string]$creds.ApiUrl
            CompanyId  = [string]$creds.CompanyId
            PublicKey  = [string]$creds.PublicKey
            PrivateKey = [string]$creds.PrivateKey
            ClientId   = [string]$creds.ClientId
            Source     = 'VScanMagic'
        }
    } catch {
        if (-not $Quiet) { Write-Warning "Could not read VScanMagic Manage credentials: $($_.Exception.Message)" }
        return $null
    }
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

Export-ModuleMember -Function `
    Get-VScanMagicManageCredentialsPath, `
    Import-ConnectWiseManageCredentialsFromVScan, `
    Get-ConnectWiseManageCredentialsFromSettings, `
    Test-ConnectWiseManageCredentialsComplete, `
    Test-ConnectWiseManageCredentialsPartial, `
    Get-ConnectWiseManageApiBaseUri, `
    Invoke-ConnectWiseManageApi, `
    Test-ConnectWiseManageConnection
