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
# CredRead P/Invoke: load native helper type once per session (see _Ensure-CredReadNativeType)
$script:credReadNativeTypeLoaded = $false

# CredentialManager: avoid repeated Get-Module -ListAvailable / Import-Module on every WCM call
$script:credMgrListAvailable = $null   # $null = not yet checked
$script:credMgrImported = $false
$script:credMgrImportFailed = $false

# Compiled regex for Get-WCMTenantIds (prefix is fixed at module load)
$script:wcmTenantIdRegex = [regex]::new(
    [regex]::Escape($script:credTargetPrefix) + '([a-fA-F0-9\-]{36})',
    [System.Text.RegularExpressions.RegexOptions]::Compiled)

# Cache Graph /organization displayName per tenant per session (avoids duplicate token + HTTP)
$script:tenantOrgDisplayNameCache = [System.Collections.Generic.Dictionary[string, object]]::new([StringComparer]::OrdinalIgnoreCase)

function _EnsureCredentialManagerImported {
    <#
    .SYNOPSIS
        Returns $true if CredentialManager module is loaded; caches list/import state for the session.
    #>
    if ($script:credMgrImported) { return $true }
    if ($script:credMgrImportFailed) { return $false }
    if ($null -eq $script:credMgrListAvailable) {
        $script:credMgrListAvailable = [bool](Get-Module -ListAvailable -Name CredentialManager)
        if (-not $script:credMgrListAvailable) { return $false }
    }
    elseif (-not $script:credMgrListAvailable) { return $false }
    try {
        Import-Module CredentialManager -ErrorAction Stop
        $script:credMgrImported = $true
        return $true
    }
    catch {
        $script:credMgrImportFailed = $true
        return $false
    }
}

function _Get-SecureStringAsPlainForKey {
    <#
    .SYNOPSIS
        Derives UTF-8 bytes from a SecureString for key material, then zeroes the intermediate BSTR.
    #>
    param([Parameter(Mandatory = $true)][SecureString]$SecureString)
    $bstr = [IntPtr]::Zero
    try {
        $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureString)
        $plain = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
        return [System.Text.Encoding]::UTF8.GetBytes($plain)
    }
    finally {
        if ($bstr -ne [IntPtr]::Zero) {
            [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) | Out-Null
        }
    }
}

function _Ensure-CredReadNativeType {
    <#
    .SYNOPSIS
        Ensures EOACredRead.Util (CredRead/CredFree P/Invoke) is loaded exactly once.
    #>
    if ($script:credReadNativeTypeLoaded) { return $true }
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
        $script:credReadNativeTypeLoaded = $true
        return $true
    }
    catch {
        $msg = $_.Exception.Message
        if ($msg -match 'already exists|duplicate type name|already been added|Cannot add type') {
            $script:credReadNativeTypeLoaded = $true
            return $true
        }
        Write-Warning "GraphAppCredential: Could not load CredRead native type (P/Invoke). WCM read fallback may fail: $msg"
        return $false
    }
}

function _Get-ImportedCredProperty {
    param(
        [Parameter(Mandatory = $true)]$Object,
        [Parameter(Mandatory = $true)][string]$Name
    )
    if ($null -eq $Object) { return $null }
    if ($Object -is [hashtable]) {
        if ($Object.ContainsKey($Name)) { return $Object[$Name] }
        return $null
    }
    $p = $Object.PSObject.Properties[$Name]
    if ($p) { return $p.Value }
    return $null
}

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
    if (_EnsureCredentialManagerImported) {
        try {
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
    if (_EnsureCredentialManagerImported) {
        try {
            $cred = New-Object PSCredential $userName, (ConvertTo-SecureString $ClientSecret -AsPlainText -Force)
            New-StoredCredential -Target $target -Credentials $cred -ErrorAction Stop | Out-Null
            $usedCredMgr = $true
        } catch {
            Write-Warning "CredentialManager failed; falling back to cmdkey. Install CredentialManager for secure storage: Install-Module CredentialManager -Scope CurrentUser"
        }
    }

    if (-not $usedCredMgr) {
        # Fallback: cmdkey (built-in, works in pwsh). SECURITY: /pass: exposes secret in process argv.
        Write-Warning "CredentialManager not installed. Client secret may be visible in process argv. Install for secure storage: Install-Module CredentialManager -Scope CurrentUser"
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
            if (_EnsureCredentialManagerImported) {
                $nameCred = New-Object PSCredential 'DisplayName', (ConvertTo-SecureString $TenantDisplayName -AsPlainText -Force)
                New-StoredCredential -Target $nameTarget -Credentials $nameCred -ErrorAction Stop | Out-Null
            } else {
                Start-Process -FilePath "cmdkey.exe" -ArgumentList "/generic:$nameTarget", "/user:DisplayName", "/pass:$TenantDisplayName" -Wait -PassThru -WindowStyle Hidden | Out-Null
            }
        } catch { /* non-fatal */ }
    }
    [void]$script:tenantOrgDisplayNameCache.Remove($TenantId)
}

function _Get-StoredDisplayName {
    param([string]$TenantId)
    $target = "${script:credTargetPrefix}${TenantId}-DisplayName"
    try {
        if (_EnsureCredentialManagerImported) {
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
    if (_EnsureCredentialManagerImported) {
        try {
            Remove-StoredCredential -Target $target -ErrorAction SilentlyContinue
            Remove-StoredCredential -Target $nameTarget -ErrorAction SilentlyContinue
            return
        } catch {}
    }
    try {
        Start-Process -FilePath "cmdkey.exe" -ArgumentList "/delete:$target" -Wait -WindowStyle Hidden -ErrorAction SilentlyContinue
        Start-Process -FilePath "cmdkey.exe" -ArgumentList "/delete:$nameTarget" -Wait -WindowStyle Hidden -ErrorAction SilentlyContinue
    } catch {}
    [void]$script:tenantOrgDisplayNameCache.Remove($TenantId)
}

function Get-WCMTenantIds {
    <#
    .SYNOPSIS
        Returns tenant IDs that have Graph app credentials stored in Windows Credential Manager.
    .OUTPUTS
        [string[]] Tenant IDs, or @() if none found
    .NOTES
        Implementation runs "cmdkey /list" and parses stdout with a regex (EOA-GraphApp-{GUID}).
        Output format can vary by Windows locale or cmdkey version; if discovery fails unexpectedly,
        verify credentials exist in Credential Manager and that targets still use prefix EOA-GraphApp-.
    #>
    $tenantIds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    try {
        $output = cmdkey /list 2>$null
        if ($output) {
            $text = if ($output -is [string]) { $output } else { [string]::Join([Environment]::NewLine, @($output)) }
            foreach ($match in $script:wcmTenantIdRegex.Matches($text)) {
                if ($match.Success -and $match.Groups[1].Value) {
                    [void]$tenantIds.Add($match.Groups[1].Value)
                }
            }
        }
    } catch {}
    # HashSet iteration order is undefined; sort for stable callers (e.g. export/tests)
    return @($tenantIds | Sort-Object)
}

function Get-TenantDisplayNameFromWCM {
    <#
    .SYNOPSIS
        Resolves tenant ID to display name using Graph API (requires WCM credentials).
    .OUTPUTS
        Display name string, or $null if resolution fails
    #>
    param([Parameter(Mandatory = $true)][string]$TenantId)
    if ($script:tenantOrgDisplayNameCache.ContainsKey($TenantId)) {
        $cached = $script:tenantOrgDisplayNameCache[$TenantId]
        return $cached
    }
    $token = Get-GraphAppTokenFromWCM -TenantId $TenantId
    if (-not $token) {
        $script:tenantOrgDisplayNameCache[$TenantId] = $null
        return $null
    }
    try {
        $headers = @{ Authorization = "Bearer $token" }
        $resp = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/organization" -Headers $headers -Method Get -ErrorAction Stop
        if ($resp.value -and $resp.value.Count -gt 0 -and $resp.value[0].displayName) {
            $name = $resp.value[0].displayName
            $script:tenantOrgDisplayNameCache[$TenantId] = $name
            return $name
        }
    } catch {}
    $script:tenantOrgDisplayNameCache[$TenantId] = $null
    return $null
}

function Get-WCMTenantListWithNames {
    <#
    .SYNOPSIS
        Returns WCM tenants with display names for dropdown display, sorted alphabetically by DisplayText.
    .OUTPUTS
        @(@{ TenantId; DisplayName; DisplayText }, ...)
    #>
    $result = [System.Collections.ArrayList]::new()
    $ids = Get-WCMTenantIds
    foreach ($tid in $ids) {
        $name = _Get-StoredDisplayName -TenantId $tid
        if (-not $name) { $name = Get-TenantDisplayNameFromWCM -TenantId $tid }
        $displayText = if ($name) { "$name ($tid)" } else { $tid }
        [void]$result.Add([pscustomobject]@{ TenantId = $tid; DisplayName = $name; DisplayText = $displayText })
    }
    return @($result | Sort-Object -Property DisplayText)
}

function _Set-GraphAppFailureInCallerScope {
    <#
    .SYNOPSIS
        Sets a variable in the caller's scope (not the module scope). Exported module functions use Scope 1 = module;
        Set-Variable -Scope 1 from Get-GraphAppTokenFromWCM did not update the worker/GUI script's $wcmErr.
    #>
    param([string]$Name, [string]$Message)
    if (-not $Name) { return }
    foreach ($s in 2..25) {
        try {
            Set-Variable -Name $Name -Value $Message -Scope $s -ErrorAction Stop
            return
        } catch { }
    }
    try { Set-Variable -Name $Name -Value $Message -Scope Global -ErrorAction SilentlyContinue } catch { }
}

function _Report-GraphAppTokenFailure {
    <#
    .SYNOPSIS
        Sets FailureVariable in caller scope when possible, and always emits WARNING so bulk worker consoles show the reason
        even when nested scopes block Set-Variable (e.g. invoked scriptblocks).
    #>
    param([string]$FailureVariable, [string]$TenantId, [string]$Message)
    if (-not $FailureVariable) { return }
    _Set-GraphAppFailureInCallerScope -Name $FailureVariable -Message $Message
    Write-Warning "Get-GraphAppTokenFromWCM [$TenantId]: $Message"
}

function Get-GraphAppTokenFromWCM {
    <#
    .SYNOPSIS
        Gets an app-only access token using credentials from WCM. Returns $null if not found or token request fails.
    .PARAMETER FailureVariable
        Optional. Name of a variable in the caller's scope to set with a short failure reason (for diagnostics).
    .NOTES
        Use -Verbose for additional detail. Use -FailureVariable err to capture why $null was returned.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$TenantId,
        [Parameter(Mandatory = $false)][string]$FailureVariable
    )
    $cred = Get-GraphAppCredentialFromWCM -TenantId $TenantId
    if (-not $cred) {
        $msg = "No app credentials found in WCM for tenant $TenantId."
        Write-Verbose "Get-GraphAppTokenFromWCM: $msg"
        _Report-GraphAppTokenFailure -FailureVariable $FailureVariable -TenantId $TenantId -Message $msg
        return $null
    }
    $tokenUrl = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
    $body = @{
        client_id     = $cred.ClientId
        client_secret = $cred.ClientSecret
        scope         = 'https://graph.microsoft.com/.default'
        grant_type    = 'client_credentials'
    }
    try {
        $resp = Invoke-RestMethod -Uri $tokenUrl -Method POST -Body $body -ContentType 'application/x-www-form-urlencoded' -ErrorAction Stop
        if (-not $resp.access_token) {
            $msg = 'Token endpoint returned no access_token (check app registration and tenant).'
            Write-Verbose "Get-GraphAppTokenFromWCM: $msg"
            _Report-GraphAppTokenFailure -FailureVariable $FailureVariable -TenantId $TenantId -Message $msg
            return $null
        }
        return $resp.access_token
    }
    catch {
        $detail = $_.Exception.Message
        if ($_.ErrorDetails.Message) {
            try {
                $j = $_.ErrorDetails.Message | ConvertFrom-Json -ErrorAction SilentlyContinue
                if ($j.error_description) { $detail = $j.error_description }
                elseif ($j.error) { $detail = $j.error }
            }
            catch { /* keep Exception.Message */ }
        }
        $msg = "Token request failed: $detail"
        Write-Verbose "Get-GraphAppTokenFromWCM: $msg"
        _Report-GraphAppTokenFailure -FailureVariable $FailureVariable -TenantId $TenantId -Message $msg
        return $null
    }
}

function _ReadCredentialViaCredRead {
    param([string]$Target)
    if (-not $Target) { return $null }
    if (-not (_Ensure-CredReadNativeType)) { return $null }
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
    $creds = [System.Collections.ArrayList]::new()
    foreach ($tid in $ids) {
        $c = Get-GraphAppCredentialFromWCM -TenantId $tid
        if ($c) {
            $dn = _Get-StoredDisplayName -TenantId $tid
            if ($dn) { $c | Add-Member -NotePropertyName 'TenantDisplayName' -NotePropertyValue $dn -Force }
            [void]$creds.Add($c)
        }
    }
    if ($creds.Count -eq 0) { throw "Could not read any credentials." }
    $json = @($creds) | ConvertTo-Json -Compress
    $pwdBytes = _Get-SecureStringAsPlainForKey -SecureString $Password
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $key = $sha.ComputeHash($pwdBytes)[0..31]
    }
    finally {
        $sha.Dispose()
    }
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
    $pwdBytes = _Get-SecureStringAsPlainForKey -SecureString $Password
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $key = $sha.ComputeHash($pwdBytes)[0..31]
    }
    finally {
        $sha.Dispose()
    }
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
        $tid = [string](_Get-ImportedCredProperty -Object $c -Name 'TenantId')
        $cid = [string](_Get-ImportedCredProperty -Object $c -Name 'ClientId')
        $secret = [string](_Get-ImportedCredProperty -Object $c -Name 'ClientSecret')
        if ([string]::IsNullOrWhiteSpace($tid) -or [string]::IsNullOrWhiteSpace($cid) -or [string]::IsNullOrWhiteSpace($secret)) { continue }
        $displayName = _Get-ImportedCredProperty -Object $c -Name 'TenantDisplayName'
        if ([string]::IsNullOrWhiteSpace([string]$displayName)) { $displayName = $null }
        try {
            Save-GraphAppCredentialToWCM -TenantId $tid -ClientId $cid -ClientSecret $secret -TenantDisplayName $displayName
            $count++
        } catch { Write-Warning "Failed to import $tid : $($_.Exception.Message)" }
    }
    return $count
}

Export-ModuleMember -Function Get-GraphAppCredentialFromWCM, Save-GraphAppCredentialToWCM, Remove-GraphAppCredentialFromWCM, Get-GraphAppTokenFromWCM, Get-WCMTenantIds, Get-TenantDisplayNameFromWCM, Get-WCMTenantListWithNames, Export-GraphAppCredentialsToFile, Import-GraphAppCredentialsFromFile
