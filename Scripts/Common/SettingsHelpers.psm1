# Settings Helpers Module
# Shared functions for loading settings and API keys

function Import-SettingsModule {
    param(
        [Parameter(Mandatory=$false)]
        [string]$ScriptRoot = $PSScriptRoot,
        
        [Parameter(Mandatory=$false)]
        [switch]$VerboseOutput
    )
    
    $settingsPath = Join-Path $ScriptRoot '..\Modules\Settings.psm1'
    
    try {
        if (Test-Path $settingsPath) {
            Import-Module $settingsPath -Force -ErrorAction Stop
            if ($VerboseOutput) { Write-Verbose "Settings module loaded successfully" }
            return $true
        }
    } catch {
        if ($VerboseOutput) { Write-Warning "Failed to load settings module: $($_.Exception.Message)" }
    }
    
    return $false
}

function Get-ApiKeyFromSettings {
    param(
        [Parameter(Mandatory=$true)]
        [string]$KeyName,  # 'ClaudeApiKey' or 'GeminiApiKey'
        
        [Parameter(Mandatory=$false)]
        [switch]$VerboseOutput
    )
    
    try {
        $settings = Get-AppSettings
        if ($settings -and $settings.$KeyName) {
            return $settings.$KeyName
        }
    } catch {
        if ($VerboseOutput) { Write-Verbose "Failed to get API key from settings: $($_.Exception.Message)" }
    }
    
    return $null
}

Export-ModuleMember -Function Import-SettingsModule, Get-ApiKeyFromSettings
