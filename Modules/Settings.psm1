function Get-SettingsPath {
    $dir = Join-Path ([Environment]::GetFolderPath('ApplicationData')) 'ExchangeOnlineAnalyzer'
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    return (Join-Path $dir 'settings.json')
}

# Encrypt a string using Windows DPAPI (Data Protection API)
function Protect-ApiKey {
    param([string]$PlainText)

    if ([string]::IsNullOrWhiteSpace($PlainText)) {
        return ''
    }

    try {
        $secureString = ConvertTo-SecureString $PlainText -AsPlainText -Force
        $encrypted = ConvertFrom-SecureString $secureString
        return $encrypted
    } catch {
        Write-Verbose "Failed to encrypt API key: $($_.Exception.Message)"
        return $PlainText  # Fallback to plain text if encryption fails
    }
}

# Decrypt a string using Windows DPAPI
function Unprotect-ApiKey {
    param([string]$EncryptedText)

    if ([string]::IsNullOrWhiteSpace($EncryptedText)) {
        return ''
    }

    try {
        # Check if it's already encrypted (DPAPI format starts with certain patterns)
        if ($EncryptedText.Length -gt 50 -and $EncryptedText -match '^[0-9a-fA-F]+\|') {
            $secureString = ConvertTo-SecureString $EncryptedText
            $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureString)
            $decrypted = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
            [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
            return $decrypted
        } else {
            # Not encrypted, return as-is
            return $EncryptedText
        }
    } catch {
        Write-Verbose "Failed to decrypt API key, treating as plain text: $($_.Exception.Message)"
        return $EncryptedText  # Fallback to returning as-is if decryption fails
    }
}

function Get-AppSettings {
    try {
        $path = Get-SettingsPath
        if (Test-Path $path) {
            $raw = Get-Content -Path $path -Raw -ErrorAction Stop
            if ($raw.Trim().Length -gt 0) {
                $settings = $raw | ConvertFrom-Json

                # Decrypt API keys if they exist
                if ($settings.GeminiApiKey) {
                    $settings.GeminiApiKey = Unprotect-ApiKey -EncryptedText $settings.GeminiApiKey
                }
                if ($settings.ClaudeApiKey) {
                    $settings.ClaudeApiKey = Unprotect-ApiKey -EncryptedText $settings.ClaudeApiKey
                }

                return $settings
            }
        }
    } catch {
        Write-Verbose "Failed to load settings: $($_.Exception.Message)"
    }
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
        # Create a copy of settings for encryption
        $settingsToSave = [pscustomobject]@{
            InvestigatorName = $Settings.InvestigatorName
            CompanyName = $Settings.CompanyName
            GeminiApiKey = ''
            ClaudeApiKey = ''
        }

        # Encrypt API keys before saving
        if ($Settings.GeminiApiKey) {
            $settingsToSave.GeminiApiKey = Protect-ApiKey -PlainText $Settings.GeminiApiKey
        }
        if ($Settings.ClaudeApiKey) {
            $settingsToSave.ClaudeApiKey = Protect-ApiKey -PlainText $Settings.ClaudeApiKey
        }

        $json = $settingsToSave | ConvertTo-Json -Depth 4
        $path = Get-SettingsPath
        $json | Out-File -FilePath $path -Encoding utf8

        # Set file permissions to restrict access to current user only (on Windows)
        if ($PSVersionTable.PSVersion.Major -ge 6 -or $PSVersionTable.Platform -eq 'Win32NT' -or [Environment]::OSVersion.Platform -eq 'Win32NT') {
            try {
                $acl = Get-Acl $path
                # Remove inheritance
                $acl.SetAccessRuleProtection($true, $false)
                # Remove all existing rules
                $acl.Access | ForEach-Object { $acl.RemoveAccessRule($_) | Out-Null }
                # Add rule for current user only
                $currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
                $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule($currentUser, 'FullControl', 'Allow')
                $acl.SetAccessRule($accessRule)
                Set-Acl -Path $path -AclObject $acl
            } catch {
                Write-Verbose "Failed to set file permissions: $($_.Exception.Message)"
            }
        }

        return $true
    } catch {
        Write-Error "Failed to save settings: $($_.Exception.Message)"
        return $false
    }
}

Export-ModuleMember -Function Get-AppSettings,Save-AppSettings,Get-SettingsPath


