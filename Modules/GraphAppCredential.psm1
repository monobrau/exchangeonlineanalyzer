<#
.SYNOPSIS
    Store and retrieve Graph app credentials (app-only) in Windows Credential Manager.
.DESCRIPTION
    Uses CredentialManager module. Target format: EOA-GraphApp-{tenantId}
    UserName stores "TenantId|ClientId", Password stores ClientSecret.
.NOTES
    Requires: Install-Module CredentialManager
#>

$script:credTargetPrefix = 'EOA-GraphApp-'

function Get-GraphAppCredentialFromWCM {
    <#
    .SYNOPSIS
        Retrieves Graph app credentials from Windows Credential Manager for a tenant.
    .OUTPUTS
        @{ TenantId; ClientId; ClientSecret } or $null if not found
    .NOTES
        Tries CredentialManager first, falls back to CredRead P/Invoke (for pwsh compatibility).
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$TenantId
    )
    $target = "$script:credTargetPrefix$TenantId"

    # Try CredentialManager first (works in Windows PowerShell 5.1)
    if (Get-Module -ListAvailable -Name CredentialManager) {
        try {
            Import-Module CredentialManager -ErrorAction Stop
            $cred = Get-StoredCredential -Target $target -ErrorAction SilentlyContinue
            if ($cred) {
                $parts = $cred.UserName -split '\|', 2
                if ($parts.Count -ge 2) {
                    return [pscustomobject]@{
                        TenantId     = $parts[0]
                        ClientId     = $parts[1]
                        ClientSecret = $cred.GetNetworkCredential().Password
                    }
                }
            }
        } catch {
            # CredentialManager may fail in pwsh
        }
    }

    # Fallback: CredRead P/Invoke (works in pwsh)
    try {
        $credObj = _ReadCredentialViaCredRead -Target $target
        if (-not $credObj) { return $null }
        $parts = $credObj.UserName -split '\|', 2
        if ($parts.Count -lt 2) { return $null }
        return [pscustomobject]@{
            TenantId     = $parts[0]
            ClientId     = $parts[1]
            ClientSecret = $credObj.CredentialBlob
        }
    } catch {
        return $null
    }
}

function Save-GraphAppCredentialToWCM {
    <#
    .SYNOPSIS
        Saves Graph app credentials to Windows Credential Manager.
    .PARAMETER TenantDisplayName
        Optional. Tenant display name to store for dropdown display (avoids Graph API lookup later).
    .NOTES
        Uses CredentialManager module when available. Falls back to cmdkey on pwsh (CredentialManager
        fails in pwsh due to System.Web dependency).
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$TenantId,
        [Parameter(Mandatory = $true)]
        [string]$ClientId,
        [Parameter(Mandatory = $true)]
        [string]$ClientSecret,
        [Parameter(Mandatory = $false)]
        [string]$TenantDisplayName
    )
    $target = "$script:credTargetPrefix$TenantId"
    $userName = "${TenantId}|${ClientId}"

    # Try CredentialManager first (works in Windows PowerShell 5.1)
    # Use CurrentUser persistence to avoid UAC prompt (LocalMachine can freeze waiting for hidden elevation dialog)
    $usedCredMgr = $false
    if (Get-Module -ListAvailable -Name CredentialManager) {
        try {
            Import-Module CredentialManager -ErrorAction Stop
            $cred = New-Object PSCredential $userName, (ConvertTo-SecureString $ClientSecret -AsPlainText -Force)
            New-StoredCredential -Target $target -Credentials $cred -ErrorAction Stop | Out-Null
            $usedCredMgr = $true
        } catch {
            # CredentialManager fails in pwsh (System.Web.Membership not in .NET Core)
        }
    }

    if (-not $usedCredMgr) {
        # Fallback: cmdkey (built-in, works in pwsh)
        try {
            $targetArg = "/generic:$target"
            $userArg = "/user:$userName"
            $passArg = "/pass:$ClientSecret"
            $proc = Start-Process -FilePath "cmdkey.exe" -ArgumentList $targetArg, $userArg, $passArg -Wait -PassThru -WindowStyle Hidden
            if ($proc.ExitCode -ne 0) {
                throw "cmdkey exited with code $($proc.ExitCode)"
            }
        } catch {
            throw "Could not save to WCM: $($_.Exception.Message). Ensure CredentialManager is installed (Install-Module CredentialManager -Scope CurrentUser) or run from Windows PowerShell 5.1."
        }
    }

    # Store tenant display name for dropdown (avoids Graph API lookup later)
    if ($TenantDisplayName -and -not [string]::IsNullOrWhiteSpace($TenantDisplayName)) {
        $nameTarget = "${script:credTargetPrefix}${TenantId}-DisplayName"
        try {
            if (Get-Module -ListAvailable -Name CredentialManager) {
                Import-Module CredentialManager -ErrorAction Stop
                $nameCred = New-Object PSCredential 'DisplayName', (ConvertTo-SecureString $TenantDisplayName -AsPlainText -Force)
                New-StoredCredential -Target $nameTarget -Credentials $nameCred -ErrorAction Stop | Out-Null
            } else {
                Start-Process -FilePath "cmdkey.exe" -ArgumentList "/generic:$nameTarget", "/user:DisplayName", "/pass:$TenantDisplayName" -Wait -PassThru -WindowStyle Hidden | Out-Null
            }
        } catch { /* non-fatal */ }
    }
}

function _Get-StoredDisplayName {
    param([string]$TenantId)
    $target = "${script:credTargetPrefix}${TenantId}-DisplayName"
    try {
        if (Get-Module -ListAvailable -Name CredentialManager) {
            Import-Module CredentialManager -ErrorAction Stop
            $c = Get-StoredCredential -Target $target -ErrorAction SilentlyContinue
            if ($c) { return $c.GetNetworkCredential().Password }
        }
        $obj = _ReadCredentialViaCredRead -Target $target
        if ($obj -and $obj.CredentialBlob) { return $obj.CredentialBlob }
    } catch {}
    return $null
}

function Remove-GraphAppCredentialFromWCM {
    param([Parameter(Mandatory = $true)][string]$TenantId)
    $target = "$script:credTargetPrefix$TenantId"
    $nameTarget = "${script:credTargetPrefix}${TenantId}-DisplayName"
    if (Get-Module -ListAvailable -Name CredentialManager) {
        try {
            Import-Module CredentialManager -ErrorAction Stop
            Remove-StoredCredential -Target $target -ErrorAction SilentlyContinue
            Remove-StoredCredential -Target $nameTarget -ErrorAction SilentlyContinue
            return
        } catch {}
    }
    try {
        Start-Process -FilePath "cmdkey.exe" -ArgumentList "/delete:$target" -Wait -WindowStyle Hidden -ErrorAction SilentlyContinue
        Start-Process -FilePath "cmdkey.exe" -ArgumentList "/delete:$nameTarget" -Wait -WindowStyle Hidden -ErrorAction SilentlyContinue
    } catch {}
}

function Get-WCMTenantIds {
    <#
    .SYNOPSIS
        Returns tenant IDs that have Graph app credentials stored in Windows Credential Manager.
    .OUTPUTS
        [string[]] Tenant IDs, or @() if none found
    #>
    $tenantIds = @()
    try {
        $output = cmdkey /list 2>$null
        if ($output) {
            $text = $output | Out-String
            $prefix = $script:credTargetPrefix
            $pattern = [regex]::Escape($prefix) + '([a-fA-F0-9\-]{36})'
            $m = [regex]::Matches($text, $pattern)
            foreach ($match in $m) {
                if ($match.Success -and $match.Groups[1].Value) {
                    $tid = $match.Groups[1].Value
                    if ($tid -notin $tenantIds) { $tenantIds += $tid }
                }
            }
        }
    } catch {}
    return $tenantIds
}

function Get-TenantDisplayNameFromWCM {
    <#
    .SYNOPSIS
        Resolves tenant ID to display name using Graph API (requires WCM credentials).
    .OUTPUTS
        Display name string, or $null if resolution fails
    #>
    param([Parameter(Mandatory = $true)][string]$TenantId)
    $token = Get-GraphAppTokenFromWCM -TenantId $TenantId
    if (-not $token) { return $null }
    try {
        $headers = @{ Authorization = "Bearer $token" }
        $resp = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/organization" -Headers $headers -Method Get -ErrorAction Stop
        if ($resp.value -and $resp.value.Count -gt 0 -and $resp.value[0].displayName) {
            return $resp.value[0].displayName
        }
    } catch {}
    return $null
}

function Get-WCMTenantListWithNames {
    <#
    .SYNOPSIS
        Returns WCM tenants with display names for dropdown display, sorted alphabetically by DisplayText.
    .OUTPUTS
        @(@{ TenantId; DisplayName; DisplayText }, ...)
    #>
    $result = @()
    $ids = Get-WCMTenantIds
    foreach ($tid in $ids) {
        $name = _Get-StoredDisplayName -TenantId $tid
        if (-not $name) { $name = Get-TenantDisplayNameFromWCM -TenantId $tid }
        $displayText = if ($name) { "$name ($tid)" } else { $tid }
        $result += [pscustomobject]@{ TenantId = $tid; DisplayName = $name; DisplayText = $displayText }
    }
    return $result | Sort-Object -Property DisplayText
}

function Get-GraphAppTokenFromWCM {
    <#
    .SYNOPSIS
        Gets an app-only access token using credentials from WCM. Returns $null if not found
    #>
    param([Parameter(Mandatory = $true)][string]$TenantId)
    $cred = Get-GraphAppCredentialFromWCM -TenantId $TenantId
    if (-not $cred) { return $null }
    $tokenUrl = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
    $body = @{
        client_id     = $cred.ClientId
        client_secret = $cred.ClientSecret
        scope         = 'https://graph.microsoft.com/.default'
        grant_type    = 'client_credentials'
    }
    try {
        $resp = Invoke-RestMethod -Uri $tokenUrl -Method POST -Body $body -ContentType 'application/x-www-form-urlencoded' -ErrorAction Stop
        return $resp.access_token
    } catch {
        return $null
    }
}

function _ReadCredentialViaCredRead {
    param([string]$Target)
    if (-not $Target) { return $null }
    $sig = @'
[DllImport("Advapi32.dll", EntryPoint = "CredReadW", CharSet = CharSet.Unicode, SetLastError = true)]
public static extern bool CredRead(string target, uint type, int reservedFlag, out IntPtr credentialPtr);

[DllImport("Advapi32.dll", EntryPoint = "CredFree", SetLastError = true)]
public static extern bool CredFree(IntPtr cred);

[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
public struct NativeCredential {
    public uint Flags;
    public uint Type;
    public IntPtr TargetName;
    public IntPtr Comment;
    public long LastWritten;
    public uint CredentialBlobSize;
    public IntPtr CredentialBlob;
    public uint Persist;
    public uint AttributeCount;
    public IntPtr Attributes;
    public IntPtr TargetAlias;
    public IntPtr UserName;
}
'@
    try {
        Add-Type -MemberDefinition $sig -Namespace 'EOACredRead' -Name 'Util' -ErrorAction Stop
    } catch {
        if ($_.Exception.Message -notmatch 'already exists') { return $null }
    }
    $ptr = [IntPtr]::Zero
    $ok = [EOACredRead.Util]::CredRead($Target, 1, 0, [ref]$ptr)
    if (-not $ok -or $ptr -eq [IntPtr]::Zero) { return $null }
    try {
        $ncred = [System.Runtime.InteropServices.Marshal]::PtrToStructure($ptr, [EOACredRead.Util+NativeCredential])
        $userName = if ($ncred.UserName -ne [IntPtr]::Zero) { [System.Runtime.InteropServices.Marshal]::PtrToStringUni($ncred.UserName) } else { $null }
        $blob = $null
        if ($ncred.CredentialBlob -ne [IntPtr]::Zero -and $ncred.CredentialBlobSize -gt 0) {
            $blob = [System.Runtime.InteropServices.Marshal]::PtrToStringUni($ncred.CredentialBlob, [int]$ncred.CredentialBlobSize / 2)
        }
        [EOACredRead.Util]::CredFree($ptr) | Out-Null
        return [pscustomobject]@{ UserName = $userName; CredentialBlob = $blob }
    } catch {
        try { [EOACredRead.Util]::CredFree($ptr) | Out-Null } catch {}
        return $null
    }
}

function Export-GraphAppCredentialsToFile {
    <#
    .SYNOPSIS
        Exports all Graph app credentials from WCM to an encrypted file.
    .PARAMETER Path
        Output file path (e.g. .eoa-creds). Will be overwritten.
    .PARAMETER Password
        SecureString password for encryption. Required for security.
    #>
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [Parameter(Mandatory=$true)][SecureString]$Password
    )
    $ids = Get-WCMTenantIds
    if ($ids.Count -eq 0) {
        throw "No app credentials found in Windows Credential Manager."
    }
    $creds = @()
    foreach ($tid in $ids) {
        $c = Get-GraphAppCredentialFromWCM -TenantId $tid
        if ($c) {
            $dn = _Get-StoredDisplayName -TenantId $tid
            if ($dn) { $c | Add-Member -NotePropertyName 'TenantDisplayName' -NotePropertyValue $dn -Force }
            $creds += $c
        }
    }
    if ($creds.Count -eq 0) { throw "Could not read any credentials." }
    $json = $creds | ConvertTo-Json -Compress
    $key = [System.Security.Cryptography.SHA256]::Create().ComputeHash([System.Text.Encoding]::UTF8.GetBytes([System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Password))))[0..31]
    $secure = ConvertTo-SecureString $json -AsPlainText -Force
    $encrypted = $secure | ConvertFrom-SecureString -Key $key
    $header = "EOA-CREDS-1`n"
    [System.IO.File]::WriteAllText($Path, $header + $encrypted, [System.Text.Encoding]::UTF8)
}

function Import-GraphAppCredentialsFromFile {
    <#
    .SYNOPSIS
        Imports Graph app credentials from an encrypted file into WCM.
    .PARAMETER Path
        Input file path (e.g. .eoa-creds).
    .PARAMETER Password
        SecureString password used when the file was exported.
    .OUTPUTS
        Number of credentials imported.
    #>
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [Parameter(Mandatory=$true)][SecureString]$Password
    )
    if (-not (Test-Path $Path)) { throw "File not found: $Path" }
    $content = [System.IO.File]::ReadAllText($Path, [System.Text.Encoding]::UTF8)
    if ($content -notmatch '^EOA-CREDS-1\r?\n(.+)$') { throw "Invalid file format. File must be exported by this tool." }
    $encrypted = $Matches[1]
    $key = [System.Security.Cryptography.SHA256]::Create().ComputeHash([System.Text.Encoding]::UTF8.GetBytes([System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Password))))[0..31]
    try {
        $secure = $encrypted | ConvertTo-SecureString -Key $key
        $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
        $json = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
        [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
    } catch {
        throw "Decryption failed. Wrong password?"
    }
    $creds = $json | ConvertFrom-Json
    if (-not $creds) { return 0 }
    if ($creds -isnot [Array]) { $creds = @($creds) }
    $count = 0
    foreach ($c in $creds) {
        $tid = if ($c.TenantId) { $c.TenantId } elseif ($c.PSObject.Properties['TenantId']) { $c.TenantId } else { continue }
        $cid = if ($c.ClientId) { $c.ClientId } elseif ($c.PSObject.Properties['ClientId']) { $c.ClientId } else { continue }
        $secret = if ($c.ClientSecret) { $c.ClientSecret } elseif ($c.PSObject.Properties['ClientSecret']) { $c.ClientSecret } else { continue }
        $displayName = if ($c.TenantDisplayName) { $c.TenantDisplayName } elseif ($c.PSObject.Properties['TenantDisplayName']) { $c.TenantDisplayName } else { $null }
        try {
            Save-GraphAppCredentialToWCM -TenantId $tid -ClientId $cid -ClientSecret $secret -TenantDisplayName $displayName
            $count++
        } catch { Write-Warning "Failed to import $tid : $($_.Exception.Message)" }
    }
    return $count
}

Export-ModuleMember -Function Get-GraphAppCredentialFromWCM, Save-GraphAppCredentialToWCM, Remove-GraphAppCredentialFromWCM, Get-GraphAppTokenFromWCM, Get-WCMTenantIds, Get-TenantDisplayNameFromWCM, Get-WCMTenantListWithNames, Export-GraphAppCredentialsToFile, Import-GraphAppCredentialsFromFile
