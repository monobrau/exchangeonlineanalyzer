param()

function Get-TestConfig {
    [CmdletBinding()]
    param([string]$Path = (Join-Path (Split-Path -Parent $PSScriptRoot) 'test.config.json'))
    if (-not (Test-Path $Path)) { return @{} }
    try { return (Get-Content -Raw -Path $Path | ConvertFrom-Json) } catch { return @{} }
}

function Get-GraphAppToken {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][string]$TenantId,
        [Parameter(Mandatory=$true)][string]$ClientId,
        [Parameter(Mandatory=$true)][string]$ClientSecret,
        [string]$Scope = 'https://graph.microsoft.com/.default'
    )
    $body = [ordered]@{
        client_id     = $ClientId
        client_secret = $ClientSecret
        scope         = $Scope
        grant_type    = 'client_credentials'
    }
    $uri = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
    $resp = Invoke-RestMethod -Method Post -Uri $uri -Body $body -ContentType 'application/x-www-form-urlencoded' -ErrorAction Stop
    return $resp.access_token
}

function Get-BackoffDelayMs {
    param([int]$Attempt)
    $base = 500
    [int][Math]::Min(30000, $base * [Math]::Pow(2, [Math]::Max(0, $Attempt-1)))
}

function Invoke-GraphRequest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][string]$Method,
        [Parameter(Mandatory=$true)][string]$Uri,
        [hashtable]$Headers,
        $Body,
        [int]$MaxRetries = 5
    )
    $attempt = 0
    do {
        try {
            return Invoke-RestMethod -Method $Method -Uri $Uri -Headers $Headers -Body $Body -ErrorAction Stop
        } catch {
            $status = $_.Exception.Response.StatusCode.Value__
            if ($status -eq 429 -or ($status -ge 500 -and $status -lt 600)) {
                $retryAfter = $null
                try { $retryAfter = $_.Exception.Response.Headers['Retry-After'] } catch {}
                $delayMs = if ($retryAfter) { [int]$retryAfter * 1000 } else { Get-BackoffDelayMs -Attempt ($attempt+1) }
                Start-Sleep -Milliseconds $delayMs
                $attempt++
            } else { throw }
        }
    } while ($attempt -lt $MaxRetries)
    throw "Graph request failed after $MaxRetries retries: $Uri"
}

function Invoke-GraphBatch {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][string]$AccessToken,
        [Parameter(Mandatory=$true)][array]$Requests, # each: @{ id=string; method='GET'; url='...' }
        [int]$ChunkSize = 20
    )
    $headers = @{ Authorization = "Bearer $AccessToken"; 'Content-Type' = 'application/json' }
    $endpoint = 'https://graph.microsoft.com/v1.0/$batch'
    $allResponses = @()
    for ($i=0; $i -lt $Requests.Count; $i += $ChunkSize) {
        $chunk = $Requests[$i..([Math]::Min($i+$ChunkSize-1, $Requests.Count-1))]
        $payload = @{ requests = $chunk } | ConvertTo-Json -Depth 8
        $resp = Invoke-GraphRequest -Method POST -Uri $endpoint -Headers $headers -Body $payload
        if ($resp.responses) { $allResponses += $resp.responses }
    }
    return $allResponses
}

Export-ModuleMember -Function Get-TestConfig,Get-GraphAppToken,Invoke-GraphRequest,Invoke-GraphBatch


