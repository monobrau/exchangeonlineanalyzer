function Get-SettingsPath {
    $dir = Join-Path ([Environment]::GetFolderPath('ApplicationData')) 'ExchangeOnlineAnalyzer'
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    return (Join-Path $dir 'settings.json')
}

function Get-AppSettings {
    try {
        $path = Get-SettingsPath
        if (Test-Path $path) {
            $raw = Get-Content -Path $path -Raw -ErrorAction Stop
            if ($raw.Trim().Length -gt 0) { return ($raw | ConvertFrom-Json) }
        }
    } catch {}
    return [pscustomobject]@{
        InvestigatorName = 'Security Administrator'
        CompanyName = 'Organization'
        GeminiApiKey = ''
        ClaudeApiKey = ''
    }
}

function Save-AppSettings {
    param([Parameter(Mandatory=$true)][object]$Settings)
    try {
        $json = $Settings | ConvertTo-Json -Depth 4
        $path = Get-SettingsPath
        $json | Out-File -FilePath $path -Encoding utf8
        return $true
    } catch {
        Write-Error "Failed to save settings: $($_.Exception.Message)"; return $false
    }
}

Export-ModuleMember -Function Get-AppSettings,Save-AppSettings,Get-SettingsPath


