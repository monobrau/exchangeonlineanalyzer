<#
.SYNOPSIS
    Store and retrieve Graph app credentials (app-only) in Windows Credential Manager.
.DESCRIPTION
    Uses CredentialManager module. Target formats: EOA-GraphApp-{tenantId} (Exchange Online Analyzer),
    ESR-GraphApp-{tenantId} (Entra Secret Rotate). UserName stores "TenantId|ClientId", Password stores ClientSecret.
.NOTES
    Requires: Install-Module CredentialManager
#>

$script:credTargetPrefixEOA = 'EOA-GraphApp-'
$script:credTargetPrefixESR = 'ESR-GraphApp-'
# Legacy alias (EOA); used where a single default prefix string is needed
$script:credTargetPrefix = $script:credTargetPrefixEOA
# CredRead P/Invoke: load native helper type once per session (see _Ensure-CredReadNativeType)
$script:credReadNativeTypeLoaded = $false

# CredentialManager: avoid repeated Get-Module -ListAvailable / Import-Module on every WCM call
$script:credMgrListAvailable = $null   # $null = not yet checked
$script:credMgrImported = $false
$script:credMgrImportFailed = $false

# Cache Graph /organization displayName per tenant+prefix per session (avoids duplicate token + HTTP)
$script:tenantOrgDisplayNameCache = [System.Collections.Generic.Dictionary[string, object]]::new([StringComparer]::OrdinalIgnoreCase)

function _Get-CredPrefixString {
    param([Parameter(Mandatory = $true)][ValidateSet('EOA', 'ESR')][string]$Prefix)
    if ($Prefix -eq 'ESR') { return $script:credTargetPrefixESR }
    return $script:credTargetPrefixEOA
}

function _Get-WcmGraphAppTenantIdSuffixVariants {
    <#
    .SYNOPSIS
        Suffix variants after EOA-GraphApp- / ESR-GraphApp- (cmdkey targets differ by braces or GUID casing).
    #>
    param([string]$TenantId)
    $cand = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $t = $TenantId.Trim()
    if ($t) { [void]$cand.Add($t) }
    $noBrace = $t -replace '[\{\}]', ''
    if ($noBrace) {
        [void]$cand.Add($noBrace)
        [void]$cand.Add($noBrace.ToLowerInvariant())
        [void]$cand.Add($noBrace.ToUpperInvariant())
        [void]$cand.Add('{' + $noBrace + '}')
        [void]$cand.Add('{' + $noBrace.ToUpperInvariant() + '}')
    }
    return @($cand)
}

function _Get-WcmCredReadTargetVariants {
    <#
    .SYNOPSIS
        Windows stores generic credentials as LegacyGeneric:target=<name> in many cases; CredRead may need either form.
    #>
    param([string]$BaseTarget)
    $set = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    if ([string]::IsNullOrWhiteSpace($BaseTarget)) { return @() }
    $b = $BaseTarget.Trim()
    [void]$set.Add($b)
    if ($b -notlike 'LegacyGeneric:target=*') {
        [void]$set.Add('LegacyGeneric:target=' + $b)
    }
    return @($set)
}

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
        Use -Prefix ESR for Entra Secret Rotate targets (ESR-GraphApp-...).
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$TenantId,
        [Parameter(Mandatory = $false)]
        [ValidateSet('EOA', 'ESR')]
        [string]$Prefix = 'EOA'
    )
    $credPrefix = _Get-CredPrefixString -Prefix $Prefix
    foreach ($tidCand in (_Get-WcmGraphAppTenantIdSuffixVariants -TenantId $TenantId)) {
        $baseTarget = "$credPrefix$tidCand"
        foreach ($tn in (_Get-WcmCredReadTargetVariants -BaseTarget $baseTarget)) {
            if (_EnsureCredentialManagerImported) {
                try {
                    $cred = Get-StoredCredential -Target $tn -ErrorAction SilentlyContinue
                    if ($cred) {
                        $parts = $cred.UserName -split '\|', 2
                        if ($parts.Count -ge 2) {
                            $pw = $cred.GetNetworkCredential().Password
                            if (-not [string]::IsNullOrWhiteSpace($pw)) {
                                return [pscustomobject]@{
                                    TenantId       = $parts[0]
                                    ClientId       = $parts[1]
                                    ClientSecret   = $pw
                                }
                            }
                        }
                    }
                } catch {
                    # CredentialManager may fail in pwsh
                }
            }
            try {
                $credObj = _ReadCredentialViaCredRead -Target $tn
                if (-not $credObj) { continue }
                $parts = $credObj.UserName -split '\|', 2
                if ($parts.Count -lt 2) { continue }
                if ([string]::IsNullOrWhiteSpace($credObj.CredentialBlob)) { continue }
                return [pscustomobject]@{
                    TenantId     = $parts[0]
                    ClientId     = $parts[1]
                    ClientSecret = $credObj.CredentialBlob
                }
            } catch { }
        }
    }
    return $null
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
        [string]$TenantDisplayName,
        [Parameter(Mandatory = $false)]
        [ValidateSet('EOA', 'ESR')]
        [string]$Prefix = 'EOA'
    )
    $credPrefix = _Get-CredPrefixString -Prefix $Prefix
    $target = "$credPrefix$TenantId"
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
        $nameTarget = "${credPrefix}${TenantId}-DisplayName"
        try {
            if (_EnsureCredentialManagerImported) {
                $nameCred = New-Object PSCredential 'DisplayName', (ConvertTo-SecureString $TenantDisplayName -AsPlainText -Force)
                New-StoredCredential -Target $nameTarget -Credentials $nameCred -ErrorAction Stop | Out-Null
            } else {
                Start-Process -FilePath "cmdkey.exe" -ArgumentList "/generic:$nameTarget", "/user:DisplayName", "/pass:$TenantDisplayName" -Wait -PassThru -WindowStyle Hidden | Out-Null
            }
        } catch { /* non-fatal */ }
    }
    [void]$script:tenantOrgDisplayNameCache.Remove("${Prefix}|$TenantId")
    [void]$script:tenantOrgDisplayNameCache.Remove($TenantId)
}

function _Get-StoredDisplayName {
    param(
        [string]$TenantId,
        [ValidateSet('EOA', 'ESR')]
        [string]$Prefix = 'EOA'
    )
    $credPrefix = _Get-CredPrefixString -Prefix $Prefix
    foreach ($tidCand in (_Get-WcmGraphAppTenantIdSuffixVariants -TenantId $TenantId)) {
        $base = "${credPrefix}${tidCand}-DisplayName"
        foreach ($tn in (_Get-WcmCredReadTargetVariants -BaseTarget $base)) {
            try {
                if (_EnsureCredentialManagerImported) {
                    $c = Get-StoredCredential -Target $tn -ErrorAction SilentlyContinue
                    if ($c) {
                        $pwd = $c.GetNetworkCredential().Password
                        if (-not [string]::IsNullOrWhiteSpace($pwd)) { return $pwd }
                    }
                }
                $obj = _ReadCredentialViaCredRead -Target $tn
                if ($obj -and $obj.CredentialBlob) { return $obj.CredentialBlob }
            } catch { }
        }
    }
    return $null
}

function Remove-GraphAppCredentialFromWCM {
    param(
        [Parameter(Mandatory = $true)][string]$TenantId,
        [Parameter(Mandatory = $false)][ValidateSet('EOA', 'ESR')][string]$Prefix = 'EOA'
    )
    $credPrefix = _Get-CredPrefixString -Prefix $Prefix
    foreach ($tidCand in (_Get-WcmGraphAppTenantIdSuffixVariants -TenantId $TenantId)) {
        $targets = @("$credPrefix$tidCand", "${credPrefix}${tidCand}-DisplayName")
        foreach ($base in $targets) {
            foreach ($tn in (_Get-WcmCredReadTargetVariants -BaseTarget $base)) {
                if (_EnsureCredentialManagerImported) {
                    try { Remove-StoredCredential -Target $tn -ErrorAction SilentlyContinue } catch { }
                }
                try {
                    Start-Process -FilePath "cmdkey.exe" -ArgumentList "/delete:$tn" -Wait -WindowStyle Hidden -ErrorAction SilentlyContinue
                } catch { }
            }
        }
        $ck = "${Prefix}|$($tidCand.Trim())"
        [void]$script:tenantOrgDisplayNameCache.Remove($ck)
    }
    [void]$script:tenantOrgDisplayNameCache.Remove($TenantId)
    [void]$script:tenantOrgDisplayNameCache.Remove("${Prefix}|$TenantId")
}

function _Get-GraphAppShortTargetsFromCmdKeyList {
    <#
    .SYNOPSIS
        Parses "cmdkey /list" output for stored target names like EOA-GraphApp-{guid} (short form used by WCM APIs).
    #>
    param([Parameter(Mandatory = $true)][string]$NamePrefix)
    $set = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    try {
        $output = cmdkey /list 2>$null
        $text = if ($output -is [string]) { $output } else { [string]::Join([Environment]::NewLine, @($output)) }
        foreach ($line in $text -split '\r?\n') {
            if ($line -notmatch 'Target:\s*(.+)$') { continue }
            $rest = $Matches[1].Trim()
            $short = $null
            if ($rest -match 'target=(.+)$') { $short = $Matches[1].Trim() }
            elseif ($rest.StartsWith($NamePrefix, [StringComparison]::OrdinalIgnoreCase)) { $short = $rest }
            if ($short -and $short.StartsWith($NamePrefix, [StringComparison]::OrdinalIgnoreCase)) {
                [void]$set.Add($short)
            }
        }
    } catch {}
    return @($set)
}

function Get-WCMUnrecognizedGraphAppTargets {
    <#
    .SYNOPSIS
        Lists GraphApp-* credential targets for the given prefix family that do not match the expected tenant GUID (or GUID-DisplayName) pattern.
        Use Remove-WCMGraphCredentialTarget to delete these individually.
    #>
    param(
        [Parameter(Mandatory = $false)]
        [ValidateSet('EOA', 'ESR')]
        [string]$Prefix = 'EOA'
    )
    $prefix = _Get-CredPrefixString -Prefix $Prefix
    $esc = [regex]::Escape($prefix)
    $validMain = "^$esc[a-fA-F0-9]{8}-[a-fA-F0-9]{4}-[a-fA-F0-9]{4}-[a-fA-F0-9]{4}-[a-fA-F0-9]{12}$"
    $validDisp = "^$esc[a-fA-F0-9]{8}-[a-fA-F0-9]{4}-[a-fA-F0-9]{4}-[a-fA-F0-9]{4}-[a-fA-F0-9]{12}-DisplayName$"
    $all = _Get-GraphAppShortTargetsFromCmdKeyList -NamePrefix $prefix
    $orphans = [System.Collections.Generic.List[string]]::new()
    foreach ($t in $all) {
        if ($t -notmatch $validMain -and $t -notmatch $validDisp) {
            $orphans.Add($t)
        }
    }
    return @($orphans | Sort-Object)
}

function Remove-WCMGraphCredentialTarget {
    <#
    .SYNOPSIS
        Removes a single Windows Credential Manager entry by its short target name (e.g. EOA-GraphApp-...).
        Does not call Microsoft Graph or delete Entra app registrations.
    #>
    param([Parameter(Mandatory = $true)][string]$TargetName)
    if (_EnsureCredentialManagerImported) {
        try { Remove-StoredCredential -Target $TargetName -ErrorAction SilentlyContinue } catch { }
    }
    try {
        Start-Process -FilePath "cmdkey.exe" -ArgumentList "/delete:$TargetName" -Wait -WindowStyle Hidden -ErrorAction SilentlyContinue
    } catch { }
}

function Remove-GraphAppCredentialsLocalOnly {
    <#
    .SYNOPSIS
        Removes stored Graph app credentials for the given tenant(s) from Windows Credential Manager only.
        Does not delete app registrations in Entra ID.
    #>
    param(
        [Parameter(Mandatory = $true)][string[]]$TenantId,
        [Parameter(Mandatory = $false)][ValidateSet('EOA', 'ESR')][string]$Prefix = 'EOA'
    )
    foreach ($tid in $TenantId) {
        if ([string]::IsNullOrWhiteSpace($tid)) { continue }
        Remove-GraphAppCredentialFromWCM -TenantId $tid.Trim() -Prefix $Prefix
    }
}

function Get-WCMTenantIds {
    <#
    .SYNOPSIS
        Returns tenant IDs that have Graph app credentials stored in Windows Credential Manager.
    .OUTPUTS
        [string[]] Tenant IDs, or @() if none found
    .NOTES
        Parses cmdkey /list for EOA-GraphApp-{GUID} or ESR-GraphApp-{GUID} (excludes *-DisplayName rows).
    #>
    param(
        [Parameter(Mandatory = $false)]
        [ValidateSet('EOA', 'ESR')]
        [string]$Prefix = 'EOA'
    )
    $credPrefix = _Get-CredPrefixString -Prefix $Prefix
    $tenantIds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    try {
        $output = cmdkey /list 2>$null
        if ($output) {
            $text = if ($output -is [string]) { $output } else { [string]::Join([Environment]::NewLine, @($output)) }
            $pattern = [regex]::Escape($credPrefix) + '(?:\{)?([a-fA-F0-9]{8}-[a-fA-F0-9]{4}-[a-fA-F0-9]{4}-[a-fA-F0-9]{4}-[a-fA-F0-9]{12})(?:\})?(?!-DisplayName)'
            $m = [regex]::Matches($text, $pattern)
            foreach ($match in $m) {
                if ($match.Success -and $match.Groups[1].Value) {
                    [void]$tenantIds.Add($match.Groups[1].Value)
                }
            }
        }
    } catch {}
    return @($tenantIds | Sort-Object)
}

function Get-TenantDisplayNameFromWCM {
    <#
    .SYNOPSIS
        Resolves tenant ID to display name using Graph API (requires WCM credentials).
    .OUTPUTS
        Display name string, or $null if resolution fails
    #>
    param(
        [Parameter(Mandatory = $true)][string]$TenantId,
        [Parameter(Mandatory = $false)][ValidateSet('EOA', 'ESR')][string]$Prefix = 'EOA'
    )
    $cacheKey = "${Prefix}|$TenantId"
    if ($script:tenantOrgDisplayNameCache.ContainsKey($cacheKey)) {
        $cached = $script:tenantOrgDisplayNameCache[$cacheKey]
        return $cached
    }
    $token = Get-GraphAppTokenFromWCM -TenantId $TenantId -Prefix $Prefix
    if (-not $token) {
        $script:tenantOrgDisplayNameCache[$cacheKey] = $null
        return $null
    }
    try {
        $headers = @{ Authorization = "Bearer $token" }
        $resp = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/organization" -Headers $headers -Method Get -ErrorAction Stop
        if ($resp.value -and $resp.value.Count -gt 0 -and $resp.value[0].displayName) {
            $name = $resp.value[0].displayName
            $script:tenantOrgDisplayNameCache[$cacheKey] = $name
            return $name
        }
    } catch {}
    $script:tenantOrgDisplayNameCache[$cacheKey] = $null
    return $null
}

function Get-WCMTenantListWithNames {
    <#
    .SYNOPSIS
        Returns WCM tenants with display names for dropdown display, sorted alphabetically by DisplayText.
    .OUTPUTS
        @(@{ TenantId; DisplayName; DisplayText; Source }, ...)
    #>
    param(
        [Parameter(Mandatory = $false)][ValidateSet('EOA', 'ESR')][string]$Prefix = 'EOA',
        [Parameter(Mandatory = $false)][switch]$SkipGraphLookup
    )
    $result = [System.Collections.ArrayList]::new()
    $ids = Get-WCMTenantIds -Prefix $Prefix
    foreach ($tid in $ids) {
        $name = _Get-StoredDisplayName -TenantId $tid -Prefix $Prefix
        if (-not $name -and -not $SkipGraphLookup) {
            $name = Get-TenantDisplayNameFromWCM -TenantId $tid -Prefix $Prefix
        }
        if (-not $name -and -not $SkipGraphLookup) {
            $alt = if ($Prefix -eq 'EOA') { 'ESR' } else { 'EOA' }
            $name = Get-TenantDisplayNameFromWCM -TenantId $tid -Prefix $alt
        }
        $displayText = if ($name) {
            if ($Prefix -eq 'ESR') { "$name ($tid) (ESR)" } else { "$name ($tid)" }
        } else {
            if ($Prefix -eq 'ESR') { "$tid (ESR)" } else { $tid }
        }
        [void]$result.Add([pscustomobject]@{
                TenantId    = $tid
                DisplayName = $name
                DisplayText = $displayText
                Source      = $Prefix
            })
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
        [Parameter(Mandatory = $false)][ValidateSet('EOA', 'ESR')][string]$Prefix = 'EOA',
        [Parameter(Mandatory = $false)][string]$FailureVariable
    )
    $cred = Get-GraphAppCredentialFromWCM -TenantId $TenantId -Prefix $Prefix
    if (-not $cred) {
        $msg = "No app credentials found in WCM for tenant $TenantId (prefix $Prefix)."
        Write-Verbose "Get-GraphAppTokenFromWCM: $msg"
        _Report-GraphAppTokenFailure -FailureVariable $FailureVariable -TenantId $TenantId -Message $msg
        return $null
    }
    $tenantForUrl = ($cred.TenantId -replace '[\{\}]', '').Trim()
    if (-not $tenantForUrl) { $tenantForUrl = ($TenantId -replace '[\{\}]', '').Trim() }
    $tokenUrl = "https://login.microsoftonline.com/$tenantForUrl/oauth2/v2.0/token"
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
        [Parameter(Mandatory=$true)][SecureString]$Password,
        [Parameter(Mandatory=$false)][ValidateSet('EOA', 'ESR')][string]$Prefix = 'EOA'
    )
    $ids = Get-WCMTenantIds -Prefix $Prefix
    if ($ids.Count -eq 0) {
        throw "No app credentials found in Windows Credential Manager for prefix $Prefix."
    }
    $creds = [System.Collections.ArrayList]::new()
    foreach ($tid in $ids) {
        $c = Get-GraphAppCredentialFromWCM -TenantId $tid -Prefix $Prefix
        if ($c) {
            $dn = _Get-StoredDisplayName -TenantId $tid -Prefix $Prefix
            if ($dn) { $c | Add-Member -NotePropertyName 'TenantDisplayName' -NotePropertyValue $dn -Force }
            $c | Add-Member -NotePropertyName 'WcmPrefix' -NotePropertyValue $Prefix -Force
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
        $wcmPfx = _Get-ImportedCredProperty -Object $c -Name 'WcmPrefix'
        $savePrefix = if ([string]$wcmPfx -eq 'ESR') { 'ESR' } else { 'EOA' }
        try {
            Save-GraphAppCredentialToWCM -TenantId $tid -ClientId $cid -ClientSecret $secret -TenantDisplayName $displayName -Prefix $savePrefix
            $count++
        } catch { Write-Warning "Failed to import $tid : $($_.Exception.Message)" }
    }
    return $count
}

function Show-ClearLocalGraphWcmPicker {
    <#
    .SYNOPSIS
        UI: pick stored EOA Graph app credential(s) to remove from Windows Credential Manager only (Entra unchanged).
    .OUTPUTS
        Number of tenant/orphan entries cleared.
    #>
    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
    } catch {
        Write-Warning "Show-ClearLocalGraphWcmPicker: System.Windows.Forms not available: $($_.Exception.Message)"
        return 0
    }
    $rowList = [System.Collections.ArrayList]::new()
    foreach ($pfx in @('EOA', 'ESR')) {
        foreach ($t in @(Get-WCMTenantListWithNames -Prefix $pfx -SkipGraphLookup)) {
            $label = if ($t.DisplayName) { $t.DisplayText } else { "$($t.TenantId)  (tenant ID - display name unknown)" }
            [void]$rowList.Add([pscustomobject]@{ DisplayText = $label; Kind = 'Tenant'; TenantId = $t.TenantId; WcmPrefix = $pfx; OrphanTarget = [string]$null })
        }
    }
    foreach ($pfx in @('EOA', 'ESR')) {
        foreach ($o in @(Get-WCMUnrecognizedGraphAppTargets -Prefix $pfx)) {
            [void]$rowList.Add([pscustomobject]@{ DisplayText = "Unrecognized WCM target ($pfx): $o"; Kind = 'Orphan'; TenantId = [string]$null; WcmPrefix = $pfx; OrphanTarget = $o })
        }
    }
    if ($rowList.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "No Graph app credentials (EOA/ESR) found in Windows Credential Manager.",
            "Clear local credentials",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
        return 0
    }
    $sorted = @($rowList | Sort-Object -Property DisplayText)
    $selForm = New-Object System.Windows.Forms.Form
    $selForm.Text = "Clear local credentials (this PC only)"
    $selForm.Size = New-Object System.Drawing.Size(520, 400)
    $selForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
    $selForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = "Removes entries from Windows Credential Manager only.`r`nDoes NOT delete app registrations in Entra ID.`r`n`r`nSelect one or more rows (tenant ID shown even when the name is unknown):"
    $lbl.Location = New-Object System.Drawing.Point(10, 10)
    $lbl.Size = New-Object System.Drawing.Size(490, 70)
    $clb = New-Object System.Windows.Forms.CheckedListBox
    $clb.Location = New-Object System.Drawing.Point(10, 85)
    $clb.Size = New-Object System.Drawing.Size(490, 220)
    $clb.CheckOnClick = $true
    foreach ($r in $sorted) { [void]$clb.Items.Add($r.DisplayText, $false) }
    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = "Remove selected"
    $btnOk.Location = New-Object System.Drawing.Point(200, 315)
    $btnOk.Size = New-Object System.Drawing.Size(140, 28)
    $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel"
    $btnCancel.Location = New-Object System.Drawing.Point(350, 315)
    $btnCancel.Size = New-Object System.Drawing.Size(100, 28)
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $selForm.AcceptButton = $btnOk
    $selForm.CancelButton = $btnCancel
    $selForm.Controls.AddRange(@($lbl, $clb, $btnOk, $btnCancel))
    if ($selForm.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return 0 }
    $picked = @()
    for ($i = 0; $i -lt $clb.Items.Count; $i++) {
        if ($clb.GetItemChecked($i)) { $picked += $sorted[$i] }
    }
    if ($picked.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No rows selected.", "Clear local credentials", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return 0
    }
    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "Remove $($picked.Count) stored credential entry/entries from this PC only?`n`nEntra app registrations will NOT be changed.",
        "Confirm",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return 0 }
    $removed = 0
    foreach ($p in $picked) {
        if ($p.Kind -eq 'Tenant' -and $p.TenantId) {
            $pfx = if ($p.WcmPrefix -eq 'ESR') { 'ESR' } else { 'EOA' }
            Remove-GraphAppCredentialsLocalOnly -TenantId @($p.TenantId) -Prefix $pfx
            $removed++
        }
        elseif ($p.Kind -eq 'Orphan' -and $p.OrphanTarget) {
            Remove-WCMGraphCredentialTarget -TargetName $p.OrphanTarget
            $removed++
        }
    }
    [System.Windows.Forms.MessageBox]::Show(
        "Removed $removed local credential entry/entries from Windows Credential Manager.",
        "Clear local credentials",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    )
    return $removed
}

Export-ModuleMember -Function Get-GraphAppCredentialFromWCM, Save-GraphAppCredentialToWCM, Remove-GraphAppCredentialFromWCM, Get-GraphAppTokenFromWCM, Get-WCMTenantIds, Get-TenantDisplayNameFromWCM, Get-WCMTenantListWithNames, Export-GraphAppCredentialsToFile, Import-GraphAppCredentialsFromFile, Get-WCMUnrecognizedGraphAppTargets, Remove-WCMGraphCredentialTarget, Remove-GraphAppCredentialsLocalOnly, Show-ClearLocalGraphWcmPicker
