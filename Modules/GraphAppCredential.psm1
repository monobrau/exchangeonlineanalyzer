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
# CredWrite P/Invoke for *-DisplayName (avoids cmdkey argv issues with Unicode / special chars in PS Core)
$script:credWriteNativeTypeLoaded = $false

# CredentialManager: avoid repeated Get-Module -ListAvailable / Import-Module on every WCM call
$script:credMgrListAvailable = $null   # $null = not yet checked
$script:credMgrImported = $false
$script:credMgrImportFailed = $false
$script:credMgrCmdkeyFallbackWarned = $false

# Cache Graph /organization displayName per tenant+prefix per session (avoids duplicate token + HTTP)
$script:tenantOrgDisplayNameCache = [System.Collections.Generic.Dictionary[string, object]]::new([StringComparer]::OrdinalIgnoreCase)
# Last token failure detail for Get-TenantDisplayNameFromWCM / UI (Get-GraphAppTokenFromWCM clears on entry; sets before each return $null)
$script:_GraphAppTokenLastFailureMessage = $null

function _Get-CredPrefixString {
    param([Parameter(Mandatory = $true)][ValidateSet('EOA', 'ESR')][string]$Prefix)
    if ($Prefix -eq 'ESR') { return $script:credTargetPrefixESR }
    return $script:credTargetPrefixEOA
}

function _Normalize-GraphAppTenantIdForWcm {
    <#
    .SYNOPSIS
        Canonical tenant id string (no braces, hyphenated GUID) so WCM targets match Get-WCMTenantIds after import/export.
    #>
    param([string]$TenantId)
    if ([string]::IsNullOrWhiteSpace($TenantId)) { return $TenantId }
    $t = ($TenantId.Trim() -replace '[\{\}]', '')
    if ([string]::IsNullOrWhiteSpace($t)) { return $TenantId.Trim() }
    try {
        return [Guid]::Parse($t).ToString('d')
    }
    catch {
        return $t
    }
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

function Reset-GraphAppCredentialManagerImportCache {
    <#
    .SYNOPSIS
        Clears cached CredentialManager import state so the next WCM call retries Import-Module (e.g. after Install-Module).
    #>
    $script:credMgrListAvailable = $null
    $script:credMgrImported = $false
    $script:credMgrImportFailed = $false
}

function _Warn-OnceCmdkeyCredentialManager {
    param([string]$Reason)
    if ($script:credMgrCmdkeyFallbackWarned) { return }
    $script:credMgrCmdkeyFallbackWarned = $true
    $r = if ($Reason) { " ($Reason)" } else { '' }
    Write-Warning @"
GraphAppCredential: Using built-in cmdkey for WCM$r  - client secrets can appear in process arguments.
Install: Install-Module CredentialManager -Scope CurrentUser -Force
Then restart this PowerShell session (or run Reset-GraphAppCredentialManagerImportCache and re-import this module).
If Import-Module CredentialManager still fails in PowerShell 7, use Windows PowerShell 5.1 (powershell.exe) for bulk registration.
"@
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
        Ensures EOACredRead.CredReadHelper (CredRead/CredFree in C#) is loaded exactly once.
    #>
    if ($script:credReadNativeTypeLoaded) { return $true }
    $csharp = @'
using System;
using System.Runtime.InteropServices;

namespace EOACredRead {
    public static class CredReadHelper {
        [DllImport("Advapi32.dll", EntryPoint = "CredReadW", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern bool CredRead(string target, int type, int reservedFlag, out IntPtr credentialPtr);

        [DllImport("Advapi32.dll", EntryPoint = "CredFree", SetLastError = true)]
        private static extern bool CredFree(IntPtr cred);

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct NativeCredential {
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

        public static bool TryRead(string target, out string userName, out string secret) {
            userName = null;
            secret = null;
            IntPtr ptr;
            if (!CredRead(target, 1, 0, out ptr) || ptr == IntPtr.Zero) {
                return false;
            }
            try {
                NativeCredential n = (NativeCredential)Marshal.PtrToStructure(ptr, typeof(NativeCredential));
                if (n.UserName != IntPtr.Zero) {
                    userName = Marshal.PtrToStringUni(n.UserName);
                }
                if (n.CredentialBlob != IntPtr.Zero && n.CredentialBlobSize > 0) {
                    secret = Marshal.PtrToStringUni(n.CredentialBlob, (int)(n.CredentialBlobSize / 2));
                    if (secret != null) {
                        secret = secret.TrimEnd('\0');
                    }
                }
                return !string.IsNullOrEmpty(secret);
            }
            finally {
                CredFree(ptr);
            }
        }
    }
}
'@
    try {
        Add-Type -TypeDefinition $csharp -Language CSharp -ErrorAction Stop
        $script:credReadNativeTypeLoaded = $true
        return $true
    }
    catch {
        $msg = $_.Exception.Message
        if ($msg -match 'already exists|duplicate type name|already been added|Cannot add type') {
            $script:credReadNativeTypeLoaded = $true
            return $true
        }
        Write-Warning "GraphAppCredential: Could not load CredRead helper (P/Invoke). WCM read fallback may fail: $msg"
        return $false
    }
}

function _Ensure-CredWriteNativeType {
    if ($script:credWriteNativeTypeLoaded) { return $true }
    $sig = @'
public class CredWriteInterop {
    [System.Runtime.InteropServices.DllImport("Advapi32.dll", EntryPoint = "CredWriteW", CharSet = System.Runtime.InteropServices.CharSet.Unicode, SetLastError = true)]
    public static extern bool CredWrite(ref CREDENTIAL cred, uint Flags);

    [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential, CharSet = System.Runtime.InteropServices.CharSet.Unicode)]
    public struct CREDENTIAL {
        public uint Flags;
        public uint Type;
        public System.IntPtr TargetName;
        public System.IntPtr Comment;
        public long LastWritten;
        public uint CredentialBlobSize;
        public System.IntPtr CredentialBlob;
        public uint Persist;
        public uint AttributeCount;
        public System.IntPtr Attributes;
        public System.IntPtr TargetAlias;
        public System.IntPtr UserName;
    }
}
'@
    try {
        Add-Type -TypeDefinition $sig -Language CSharp -ErrorAction Stop
        $script:credWriteNativeTypeLoaded = $true
        return $true
    }
    catch {
        $msg = $_.Exception.Message
        if ($msg -match 'already exists|duplicate type name|already been added|Cannot add type') {
            $script:credWriteNativeTypeLoaded = $true
            return $true
        }
        Write-Warning "GraphAppCredential: CredWrite native type not loaded: $msg"
        return $false
    }
}

function _Normalize-DisplayNameBlob {
    param([AllowNull()][string]$Text)
    if ($null -eq $Text) { return $null }
    return $Text.TrimEnd([char]0).Trim()
}

function _Write-GenericWcmSecretNative {
    <#
    .SYNOPSIS
        Writes a generic credential via CredWriteW (UTF-16 secret). Avoids cmdkey argv issues (&, quotes, Unicode in secrets).
        Tries CRED_PERSIST_LOCAL_MACHINE (2) then CRED_PERSIST_ENTERPRISE (3); some vaults reject one or the other.
    #>
    param(
        [Parameter(Mandatory = $true)][string]$TargetName,
        [Parameter(Mandatory = $true)][string]$UserName,
        [Parameter(Mandatory = $true)][string]$SecretText
    )
    if (-not (_Ensure-CredWriteNativeType)) { return $false }
    try {
        $interopType = [CredWriteInterop]
        $credType = $interopType.GetNestedType('CREDENTIAL', [System.Reflection.BindingFlags]::Public)
    }
    catch {
        return $false
    }
    if (-not $credType) { return $false }
    $tgtPtr = [IntPtr]::Zero
    $usrPtr = [IntPtr]::Zero
    $blobPtr = [IntPtr]::Zero
    try {
        $tgtPtr = [System.Runtime.InteropServices.Marshal]::StringToCoTaskMemUni($TargetName)
        $usrPtr = [System.Runtime.InteropServices.Marshal]::StringToCoTaskMemUni($UserName)
        $secretWithNul = $SecretText + [char]0
        $bytes = [System.Text.Encoding]::Unicode.GetBytes($secretWithNul)
        $blobPtr = [System.Runtime.InteropServices.Marshal]::AllocCoTaskMem($bytes.Length)
        [System.Runtime.InteropServices.Marshal]::Copy($bytes, 0, $blobPtr, $bytes.Length)
        $lastErr = 0
        foreach ($persist in @(2, 3)) {
            $cred = [System.Activator]::CreateInstance($credType)
            $cred.Flags = 0
            $cred.Type = 1
            $cred.TargetName = $tgtPtr
            $cred.Comment = [IntPtr]::Zero
            $cred.LastWritten = 0
            $cred.CredentialBlobSize = [uint32]$bytes.Length
            $cred.CredentialBlob = $blobPtr
            $cred.Persist = [uint32]$persist
            $cred.AttributeCount = 0
            $cred.Attributes = [IntPtr]::Zero
            $cred.TargetAlias = [IntPtr]::Zero
            $cred.UserName = $usrPtr
            $ok = [CredWriteInterop]::CredWrite([ref]$cred, 0)
            if ($ok) { return $true }
            $lastErr = [System.Runtime.InteropServices.Marshal]::GetLastWin32Error()
            Write-Verbose "_Write-GenericWcmSecretNative: CredWrite Persist=$persist failed LastError=$lastErr Target=$TargetName"
        }
        Write-Warning "GraphAppCredential: CredWrite could not save WCM target '$TargetName' (tried persist 2 and 3). Last Win32 error: $lastErr"
        return $false
    }
    finally {
        if ($blobPtr -ne [IntPtr]::Zero) { [System.Runtime.InteropServices.Marshal]::FreeCoTaskMem($blobPtr) }
        if ($usrPtr -ne [IntPtr]::Zero) { [System.Runtime.InteropServices.Marshal]::FreeCoTaskMem($usrPtr) }
        if ($tgtPtr -ne [IntPtr]::Zero) { [System.Runtime.InteropServices.Marshal]::FreeCoTaskMem($tgtPtr) }
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
        foreach ($k in @($Object.Keys)) {
            if ([string]::Equals([string]$k, $Name, [StringComparison]::OrdinalIgnoreCase)) { return $Object[$k] }
        }
        return $null
    }
    $p = $Object.PSObject.Properties[$Name]
    if ($p) { return $p.Value }
    foreach ($prop in $Object.PSObject.Properties) {
        if ([string]::Equals($prop.Name, $Name, [StringComparison]::OrdinalIgnoreCase)) { return $prop.Value }
    }
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
    $TenantId = _Normalize-GraphAppTenantIdForWcm -TenantId $TenantId
    $credPrefix = _Get-CredPrefixString -Prefix $Prefix
    $target = "$credPrefix$TenantId"
    $userName = "${TenantId}|${ClientId}"

    # CredentialManager's New-StoredCredential uses System.Web (.NET Framework) - not available in PowerShell 7+ (Core).
    $psCore = ($PSVersionTable.PSEdition -eq 'Core')
    $mainCredSaved = $false
    if (-not $psCore -and (_EnsureCredentialManagerImported)) {
        try {
            $cred = New-Object PSCredential $userName, (ConvertTo-SecureString $ClientSecret -AsPlainText -Force)
            New-StoredCredential -Target $target -Credentials $cred -ErrorAction Stop | Out-Null
            $mainCredSaved = $true
        } catch {
            _Warn-OnceCmdkeyCredentialManager -Reason "New-StoredCredential failed: $($_.Exception.Message)"
        }
    }
    elseif (-not $mainCredSaved -and -not $script:credMgrCmdkeyFallbackWarned) {
        if ($psCore) {
            _Warn-OnceCmdkeyCredentialManager -Reason 'PowerShell 7+ cannot use CredentialManager for WCM writes (System.Web). Using CredWrite/cmdkey fallback.'
        }
        elseif (-not (_EnsureCredentialManagerImported)) {
            _Warn-OnceCmdkeyCredentialManager -Reason 'CredentialManager module not available or Import-Module failed'
        }
    }

    if (-not $mainCredSaved) {
        if (_Write-GenericWcmSecretNative -TargetName $target -UserName $userName -SecretText $ClientSecret) {
            $mainCredSaved = $true
        }
        else {
            # Last resort: cmdkey (breaks when secret or user contains &, ", etc.)
            try {
                $targetArg = "/generic:$target"
                $userArg = "/user:$userName"
                $passArg = "/pass:$ClientSecret"
                $proc = Start-Process -FilePath "cmdkey.exe" -ArgumentList $targetArg, $userArg, $passArg -Wait -PassThru -WindowStyle Hidden
                if ($proc.ExitCode -ne 0) {
                    throw "cmdkey exited with code $($proc.ExitCode)"
                }
                $mainCredSaved = $true
            } catch {
                throw "Could not save to WCM: $($_.Exception.Message). Install-Module CredentialManager -Scope CurrentUser, or run Create Graph App from Windows PowerShell 5.1."
            }
        }
    }

    $readBack = Get-GraphAppCredentialFromWCM -TenantId $TenantId -Prefix $Prefix
    if (-not $readBack -or [string]::IsNullOrWhiteSpace($readBack.ClientSecret)) {
        throw "Graph app credential could not be read back from Windows Credential Manager for tenant $TenantId."
    }

    # Store tenant display name for dropdown (avoids Graph API lookup later)
    if ($TenantDisplayName -and -not [string]::IsNullOrWhiteSpace($TenantDisplayName)) {
        $nameTarget = "${credPrefix}${TenantId}-DisplayName"
        $dnOk = $false
        try {
            if (-not $psCore -and (_EnsureCredentialManagerImported)) {
                try {
                    $nameCred = New-Object PSCredential 'DisplayName', (ConvertTo-SecureString $TenantDisplayName -AsPlainText -Force)
                    New-StoredCredential -Target $nameTarget -Credentials $nameCred -ErrorAction Stop | Out-Null
                    $dnOk = $true
                }
                catch { }
            }
            if (-not $dnOk -and (_Write-GenericWcmSecretNative -TargetName $nameTarget -UserName 'DisplayName' -SecretText $TenantDisplayName)) {
                $dnOk = $true
            }
            if (-not $dnOk) {
                Write-Warning "GraphAppCredential: Could not save *-DisplayName for $TenantId (name will still resolve via Graph when online)."
            }
        } catch {
            Write-Warning "GraphAppCredential: Could not save *-DisplayName for $TenantId : $($_.Exception.Message)"
        }
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
            if (_EnsureCredentialManagerImported) {
                try {
                    $c = Get-StoredCredential -Target $tn -ErrorAction SilentlyContinue
                    if ($c) {
                        $pwd = _Normalize-DisplayNameBlob -Text ($c.GetNetworkCredential().Password)
                        if (-not [string]::IsNullOrWhiteSpace($pwd)) { return $pwd }
                    }
                }
                catch { }
            }
            try {
                $obj = _ReadCredentialViaCredRead -Target $tn
                if ($obj -and $obj.CredentialBlob) { return (_Normalize-DisplayNameBlob -Text $obj.CredentialBlob) }
            }
            catch { }
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

function _Format-GraphRestExceptionDetail {
    param([System.Management.Automation.ErrorRecord]$ErrorRecord)
    if (-not $ErrorRecord) { return 'Unknown error' }
    $msg = $ErrorRecord.Exception.Message
    if ($ErrorRecord.ErrorDetails -and $ErrorRecord.ErrorDetails.Message) {
        $raw = [string]$ErrorRecord.ErrorDetails.Message
        try {
            $j = $raw | ConvertFrom-Json -ErrorAction Stop
            if ($j.error.message) { return ($msg + '  - ' + [string]$j.error.message) }
            if ($j.error.code) { return ($msg + '  - ' + [string]$j.error.code) }
        } catch {
            if ($raw.Length -lt 500) { return ($msg + '  - ' + $raw) }
        }
    }
    return $msg
}

function _Set-DisplayNameLookupErrorRef {
    param($Ref, $Message)
    try {
        if ($null -ne $Ref -and $Ref -is [System.Management.Automation.PSReference]) {
            $Ref.Value = $Message
        }
    } catch { }
}

function Get-TenantDisplayNameFromWCM {
    <#
    .SYNOPSIS
        Resolves tenant ID to display name using Graph API (requires WCM credentials).
    .PARAMETER ForceRefresh
        Ignore session cache for this tenant/prefix (retry after failures or before re-register).
    .PARAMETER LastError
        Optional [ref] string assigned when this function returns $null (token or Graph failure reason).
    .OUTPUTS
        Display name string, or $null if resolution fails
    #>
    param(
        [Parameter(Mandatory = $true)][string]$TenantId,
        [Parameter(Mandatory = $false)][ValidateSet('EOA', 'ESR')][string]$Prefix = 'EOA',
        [Parameter(Mandatory = $false)][switch]$ForceRefresh,
        $LastError
    )
    _Set-DisplayNameLookupErrorRef -Ref $LastError -Message $null
    $cacheKey = "${Prefix}|$TenantId"
    if ($ForceRefresh) {
        [void]$script:tenantOrgDisplayNameCache.Remove($cacheKey)
    }
    elseif ($script:tenantOrgDisplayNameCache.ContainsKey($cacheKey)) {
        $cached = $script:tenantOrgDisplayNameCache[$cacheKey]
        # Do not treat cached failure ($null) as final  - allow retry on next call (same session).
        if ($null -ne $cached -and -not [string]::IsNullOrWhiteSpace([string]$cached)) {
            return [string]$cached
        }
    }
    $token = Get-GraphAppTokenFromWCM -TenantId $TenantId -Prefix $Prefix
    if (-not $token) {
        $why = $script:_GraphAppTokenLastFailureMessage
        if (-not $why) { $why = 'App-only token request failed (check WCM client id/secret and app registration).' }
        _Set-DisplayNameLookupErrorRef -Ref $LastError -Message $why
        return $null
    }
    try {
        $headers = @{ Authorization = "Bearer $token" }
        $resp = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/organization" -Headers $headers -Method Get -ErrorAction Stop
        $vals = @($resp.value)
        if ($vals.Count -gt 0) {
            $dn = $vals[0].displayName
            if (-not [string]::IsNullOrWhiteSpace([string]$dn)) {
                $name = [string]$dn.Trim()
                $script:tenantOrgDisplayNameCache[$cacheKey] = $name
                return $name
            }
        }
        _Set-DisplayNameLookupErrorRef -Ref $LastError -Message 'Microsoft Graph returned no organization displayName (empty or missing field). Grant Organization.Read.All or Directory.Read.All + admin consent for the app.'
    } catch {
        $detail = _Format-GraphRestExceptionDetail -ErrorRecord $_
        Write-Warning "Get-TenantDisplayNameFromWCM: Graph GET /organization failed ($Prefix, tenant $TenantId): $detail"
        _Set-DisplayNameLookupErrorRef -Ref $LastError -Message $detail
    }
    return $null
}

function Get-WCMTenantListWithNames {
    <#
    .SYNOPSIS
        Returns WCM tenants with display names for dropdown display, sorted alphabetically by DisplayText.
    .PARAMETER SkipGraphLookup
        When set, skips per-tenant client_credentials token + Graph /organization calls. Use for responsive UI
        (same pattern as CA Manager / XOA); labels use stored WCM *-DisplayName entries or raw tenant GUIDs.
    .PARAMETER ForceRefreshFromGraph
        When set (and SkipGraphLookup is not set), calls Get-TenantDisplayNameFromWCM -ForceRefresh so dropdowns
        pick up current /organization names even if a previous lookup cached a failure in-session.
    .OUTPUTS
        @(@{ TenantId; DisplayName; DisplayText; Source }, ...)
    #>
    param(
        [Parameter(Mandatory = $false)][ValidateSet('EOA', 'ESR')][string]$Prefix = 'EOA',
        [Parameter(Mandatory = $false)][switch]$SkipGraphLookup,
        [Parameter(Mandatory = $false)][switch]$ForceRefreshFromGraph
    )
    $result = [System.Collections.ArrayList]::new()
    $ids = Get-WCMTenantIds -Prefix $Prefix
    foreach ($tid in $ids) {
        $name = _Get-StoredDisplayName -TenantId $tid -Prefix $Prefix
        if (-not $SkipGraphLookup) {
            if ($ForceRefreshFromGraph) {
                $g = Get-TenantDisplayNameFromWCM -TenantId $tid -Prefix $Prefix -ForceRefresh
                if ($g) { $name = $g }
                elseif (-not $name) {
                    $alt = if ($Prefix -eq 'EOA') { 'ESR' } else { 'EOA' }
                    $g2 = Get-TenantDisplayNameFromWCM -TenantId $tid -Prefix $alt -ForceRefresh
                    if ($g2) { $name = $g2 }
                }
            }
            else {
                if (-not $name) {
                    $name = Get-TenantDisplayNameFromWCM -TenantId $tid -Prefix $Prefix
                }
                if (-not $name) {
                    $alt = if ($Prefix -eq 'EOA') { 'ESR' } else { 'EOA' }
                    $name = Get-TenantDisplayNameFromWCM -TenantId $tid -Prefix $alt
                }
            }
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

function Get-WCMTenantListWithNamesForAppRegCombo {
    <#
    .SYNOPSIS
        Merged EOA + ESR WCM tenants for client-auth "App reg tenant" dropdowns (Exchange Online Analyzer / Bulk Tenant Exporter).
    .DESCRIPTION
        Calls Get-WCMTenantListWithNames **without** -SkipGraphLookup so missing *-DisplayName WCM entries still get a friendly
        name from Graph. Merges duplicate tenant IDs (prefers EOA row unless EOA has no DisplayName and ESR does).
    .PARAMETER ForceRefreshFromGraph
        Passed through to Get-WCMTenantListWithNames (e.g. after Refresh tenant names in the auth console).
    .PARAMETER SkipGraphLookup
        When set, uses WCM *-DisplayName entries only (fast; no per-tenant Graph token calls).
    #>
    param(
        [Parameter(Mandatory = $false)]
        [switch]$ForceRefreshFromGraph,
        [Parameter(Mandatory = $false)]
        [switch]$SkipGraphLookup
    )
    $merged = @{}
    foreach ($pfx in @('EOA', 'ESR')) {
        foreach ($row in @(Get-WCMTenantListWithNames -Prefix $pfx -ForceRefreshFromGraph:$ForceRefreshFromGraph -SkipGraphLookup:$SkipGraphLookup -ErrorAction SilentlyContinue)) {
            $tid = [string]$row.TenantId
            if (-not $merged.ContainsKey($tid)) {
                $merged[$tid] = $row
                continue
            }
            $cur = $merged[$tid]
            $curWeak = [string]::IsNullOrWhiteSpace($cur.DisplayName)
            $newStrong = -not [string]::IsNullOrWhiteSpace($row.DisplayName)
            if ($curWeak -and $newStrong) {
                $merged[$tid] = $row
            }
        }
    }
    return @($merged.Values | Sort-Object DisplayText)
}

function Register-GraphAppTenantDisplayNamesInWCM {
    <#
    .SYNOPSIS
        Creates WCM *-DisplayName entries for each stored Graph app tenant by calling Graph /organization.
    .DESCRIPTION
        Export-GraphAppCredentialsToFile only embeds TenantDisplayName when these entries exist (or when
        -ResolveMissingDisplayNamesFromGraph is used). After registering on a PC where Graph works, export/import
        carries friendly names to machines where Graph lookup may fail.
    .PARAMETER Prefix
        EOA, ESR, or Both (default).
    .PARAMETER ForceRefresh
        When set, re-queries Microsoft Graph /organization for every stored tenant and rewrites WCM *-DisplayName even if one already exists.
    .OUTPUTS
        Count of tenants where *-DisplayName was verified in WCM after save (not merely Save invoked).
    .PARAMETER DiagnosticMessages
        Optional [ref] to an object; after the run, .Value is a string[] of per-tenant failure lines (for UI).
    #>
    param(
        [Parameter(Mandatory = $false)]
        [ValidateSet('EOA', 'ESR', 'Both')]
        [string]$Prefix = 'Both',

        [Parameter(Mandatory = $false)]
        [switch]$ForceRefresh,

        $DiagnosticMessages
    )
    $prefixes = if ($Prefix -eq 'Both') { @('EOA', 'ESR') } else { @($Prefix) }
    $registered = 0
    $diag = [System.Collections.Generic.List[string]]::new()
    foreach ($pfx in $prefixes) {
        foreach ($tidRaw in @(Get-WCMTenantIds -Prefix $pfx)) {
            if ([string]::IsNullOrWhiteSpace($tidRaw)) { continue }
            $tid = _Normalize-GraphAppTenantIdForWcm -TenantId $tidRaw
            if (-not $ForceRefresh -and (_Get-StoredDisplayName -TenantId $tid -Prefix $pfx)) { continue }
            $c = Get-GraphAppCredentialFromWCM -TenantId $tid -Prefix $pfx
            if (-not $c) {
                [void]$diag.Add("[$pfx $tid] No app credentials found in WCM for this prefix.")
                continue
            }
            $le = $null
            $name = Get-TenantDisplayNameFromWCM -TenantId $tid -Prefix $pfx -ForceRefresh -LastError ([ref]$le)
            if (-not $name) {
                $alt = if ($pfx -eq 'EOA') { 'ESR' } else { 'EOA' }
                $le2 = $null
                $name = Get-TenantDisplayNameFromWCM -TenantId $tid -Prefix $alt -ForceRefresh -LastError ([ref]$le2)
            }
            if (-not $name) {
                $parts = [System.Collections.Generic.List[string]]::new()
                if ($le) { [void]$parts.Add("$pfx`: $le") }
                if ($le2) { [void]$parts.Add("$alt`: $le2") }
                $line = if ($parts.Count -gt 0) { ($parts -join ' | ') } else { 'Could not resolve display name from Microsoft Graph.' }
                [void]$diag.Add("[$pfx $tid] $line")
                Write-Warning "Register-GraphAppTenantDisplayNamesInWCM [$pfx $tid]: $line"
                continue
            }
            try {
                Save-GraphAppCredentialToWCM -TenantId $c.TenantId -ClientId $c.ClientId -ClientSecret $c.ClientSecret -TenantDisplayName $name -Prefix $pfx
            }
            catch {
                [void]$diag.Add("[$pfx $tid] Save to WCM failed: $($_.Exception.Message)")
                Write-Warning "Register-GraphAppTenantDisplayNamesInWCM [$pfx $tid]: $($_.Exception.Message)"
                continue
            }
            $verify = _Get-StoredDisplayName -TenantId $tid -Prefix $pfx
            if ([string]::IsNullOrWhiteSpace($verify)) {
                [void]$diag.Add("[$pfx $tid] Graph returned a display name but *-DisplayName is missing in WCM afterward (CredentialManager/cmdkey write failed; try Windows PowerShell 5.1).")
                continue
            }
            $registered++
        }
    }
    try {
        if ($null -ne $DiagnosticMessages -and $DiagnosticMessages -is [System.Management.Automation.PSReference]) {
            $DiagnosticMessages.Value = @($diag)
        }
    } catch { }
    return $registered
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
        Optional. Name (string) of a variable in a parent scope to set with a short failure reason, e.g. -FailureVariable 'wcmErr' (quotes required  - not $wcmErr).
    .NOTES
        Use -Verbose for additional detail. Also sets script-level detail for Get-TenantDisplayNameFromWCM when token acquisition fails.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$TenantId,
        [Parameter(Mandatory = $false)][ValidateSet('EOA', 'ESR')][string]$Prefix = 'EOA',
        [Parameter(Mandatory = $false)][string]$FailureVariable
    )
    $script:_GraphAppTokenLastFailureMessage = $null
    $cred = Get-GraphAppCredentialFromWCM -TenantId $TenantId -Prefix $Prefix
    if (-not $cred) {
        $msg = "No app credentials found in WCM for tenant $TenantId (prefix $Prefix)."
        Write-Verbose "Get-GraphAppTokenFromWCM: $msg"
        $script:_GraphAppTokenLastFailureMessage = $msg
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
            $script:_GraphAppTokenLastFailureMessage = $msg
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
            catch {
                # JSON parse failed; keep $detail from Exception.Message
            }
        }
        $msg = "Token request failed: $detail"
        Write-Verbose "Get-GraphAppTokenFromWCM: $msg"
        $script:_GraphAppTokenLastFailureMessage = $msg
        _Report-GraphAppTokenFailure -FailureVariable $FailureVariable -TenantId $TenantId -Message $msg
        return $null
    }
}

function _ReadCredentialViaCredRead {
    param([string]$Target)
    if (-not $Target) { return $null }
    if (-not (_Ensure-CredReadNativeType)) { return $null }
    try {
        $userName = [string]::Empty
        $secret = [string]::Empty
        $ok = [EOACredRead.CredReadHelper]::TryRead($Target, [ref]$userName, [ref]$secret)
        if (-not $ok -or [string]::IsNullOrWhiteSpace($secret)) { return $null }
        return [pscustomobject]@{ UserName = $userName; CredentialBlob = $secret }
    } catch {
        Write-Warning "GraphAppCredential: CredRead failed for target '$Target': $($_.Exception.Message)"
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
    .PARAMETER ResolveMissingDisplayNamesFromGraph
        When WCM has no *-DisplayName entry for a tenant, resolve display name via Graph and embed TenantDisplayName
        in the file so Import recreates friendly labels on PCs where Graph lookup fails.
    #>
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [Parameter(Mandatory=$true)][SecureString]$Password,
        [Parameter(Mandatory=$false)][ValidateSet('EOA', 'ESR')][string]$Prefix = 'EOA',
        [Parameter(Mandatory=$false)]
        [switch]$ResolveMissingDisplayNamesFromGraph
    )
    $ids = Get-WCMTenantIds -Prefix $Prefix
    if ($ids.Count -eq 0) {
        throw "No app credentials found in Windows Credential Manager for prefix $Prefix."
    }
    $creds = [System.Collections.ArrayList]::new()
    foreach ($tid in $ids) {
        $c = Get-GraphAppCredentialFromWCM -TenantId $tid -Prefix $Prefix
        if ($c) {
            $normTid = _Normalize-GraphAppTenantIdForWcm -TenantId ([string]$c.TenantId)
            if ($normTid -and $normTid -ne [string]$c.TenantId) {
                $c | Add-Member -NotePropertyName TenantId -NotePropertyValue $normTid -Force
            }
            $dn = _Get-StoredDisplayName -TenantId $tid -Prefix $Prefix
            # Always fill a missing label from Graph so Import on another PC gets names, not raw GUIDs.
            if (-not $dn) {
                $dn = Get-TenantDisplayNameFromWCM -TenantId $tid -Prefix $Prefix -ForceRefresh
                if (-not $dn) {
                    $alt = if ($Prefix -eq 'EOA') { 'ESR' } else { 'EOA' }
                    $dn = Get-TenantDisplayNameFromWCM -TenantId $tid -Prefix $alt -ForceRefresh
                }
            }
            if ($dn) { $c | Add-Member -NotePropertyName 'TenantDisplayName' -NotePropertyValue $dn -Force }
            $c | Add-Member -NotePropertyName 'WcmPrefix' -NotePropertyValue $Prefix -Force
            [void]$creds.Add($c)
        }
    }
    if ($creds.Count -eq 0) { throw "Could not read any credentials." }
    $namesInExport = 0
    foreach ($row in $creds) {
        $rdn = _Get-ImportedCredProperty -Object $row -Name 'TenantDisplayName'
        if (-not [string]::IsNullOrWhiteSpace([string]$rdn)) { $namesInExport++ }
    }
    if ($namesInExport -eq 0) {
        Write-Warning "Export: no tenant display names will be in this file. Import on another PC will show GUIDs until names exist in WCM there. Fix: check 'Embed tenant display names' in the app, or run Register-GraphAppTenantDisplayNamesInWCM on this PC, then export again."
    }
    # -Depth ensures TenantDisplayName and long secrets serialize reliably; -EnumsAsStrings not needed
    $json = @($creds) | ConvertTo-Json -Compress -Depth 8
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
    return $creds.Count
}

function Get-GraphAppCredentialEncryptedFileSummary {
    <#
    .SYNOPSIS
        Decrypts an .eoa-creds export and reports how many rows include TenantDisplayName (does not print secrets).
    .DESCRIPTION
        Run on the export PC before copying the file, or on the import PC to verify the file before Import-GraphAppCredentialsFromFile.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][SecureString]$Password
    )
    if (-not (Test-Path $Path)) { throw "File not found: $Path" }
    $content = [System.IO.File]::ReadAllText($Path, [System.Text.Encoding]::UTF8)
    if ($content -notmatch '^EOA-CREDS-1\r?\n(.+)$') { throw "Invalid file format (expected EOA-CREDS-1 header)." }
    $encrypted = $Matches[1]
    $pwdBytes = _Get-SecureStringAsPlainForKey -SecureString $Password
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $key = $sha.ComputeHash($pwdBytes)[0..31]
    }
    finally {
        $sha.Dispose()
    }
    $secure = $encrypted | ConvertTo-SecureString -Key $key
    $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
    try {
        $json = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
    }
    finally {
        [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
    }
    $creds = $json | ConvertFrom-Json
    if (-not $creds) {
        return [pscustomobject]@{
            TotalCredentials          = 0
            WithTenantDisplayName     = 0
            WithoutTenantDisplayName  = 0
            SampleTenantIds           = @()
        }
    }
    if ($creds -isnot [Array]) { $creds = @($creds) }
    $withDn = 0
    $sampleTids = [System.Collections.Generic.List[string]]::new()
    foreach ($row in $creds) {
        $tid = [string](_Get-ImportedCredProperty -Object $row -Name 'TenantId')
        $dn = _Get-ImportedCredProperty -Object $row -Name 'TenantDisplayName'
        if ($sampleTids.Count -lt 8 -and -not [string]::IsNullOrWhiteSpace($tid)) { [void]$sampleTids.Add($tid) }
        if (-not [string]::IsNullOrWhiteSpace([string]$dn)) { $withDn++ }
    }
    return [pscustomobject]@{
        TotalCredentials          = $creds.Count
        WithTenantDisplayName     = $withDn
        WithoutTenantDisplayName  = ($creds.Count - $withDn)
        SampleTenantIds           = @($sampleTids)
    }
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
    $withDisplayName = 0
    foreach ($ic in $creds) {
        $dnProbe = _Get-ImportedCredProperty -Object $ic -Name 'TenantDisplayName'
        if (-not [string]::IsNullOrWhiteSpace([string]$dnProbe)) { $withDisplayName++ }
    }
    if ($withDisplayName -eq 0 -and $creds.Count -gt 0) {
        Write-Warning "Import: no TenantDisplayName fields in this file  - export on a PC with 'Embed tenant display names' checked (or run Register-GraphAppTenantDisplayNamesInWCM before export). Dropdowns may show GUIDs until names are stored."
    }
    $count = 0
    foreach ($c in $creds) {
        $tid = [string](_Get-ImportedCredProperty -Object $c -Name 'TenantId')
        $cid = [string](_Get-ImportedCredProperty -Object $c -Name 'ClientId')
        $secret = [string](_Get-ImportedCredProperty -Object $c -Name 'ClientSecret')
        if ([string]::IsNullOrWhiteSpace($tid) -or [string]::IsNullOrWhiteSpace($cid) -or [string]::IsNullOrWhiteSpace($secret)) { continue }
        $tid = _Normalize-GraphAppTenantIdForWcm -TenantId $tid
        $displayName = _Get-ImportedCredProperty -Object $c -Name 'TenantDisplayName'
        if ($null -ne $displayName -and $displayName -isnot [string]) { $displayName = "$displayName" }
        if ([string]::IsNullOrWhiteSpace([string]$displayName)) { $displayName = $null }
        else { $displayName = $displayName.Trim() }
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

function Invoke-GraphAppCreateWithWcmSave {
    <#
    .SYNOPSIS
        Runs Start-NewGraphInboxRulesApp.ps1 -SaveToWCM in the same PowerShell host family as the caller (pwsh stays pwsh).
    .OUTPUTS
        PSCustomObject with ExitCode, Result, LogPath
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProjectRoot,

        [Parameter(Mandatory = $false)]
        [string]$TenantId
    )
    $launcherPath = Join-Path $ProjectRoot 'Start-NewGraphInboxRulesApp.ps1'
    if (-not (Test-Path -LiteralPath $launcherPath)) {
        throw "Script not found: $launcherPath"
    }
    # Same executable as the GUI (pwsh when you launch from pwsh). Do not switch to Windows PowerShell 5.1:
    # -NoProfile there hides Microsoft.Graph modules installed for pwsh.
    $psExe = (Get-Process -Id $PID -ErrorAction Stop).Path
    $resultPath = Join-Path $env:TEMP 'EOA-GraphAppCreate-result.json'
    $logPath = Join-Path $env:TEMP 'EOA-GraphAppCreate-last.log'
    foreach ($p in @($resultPath, $logPath, "${logPath}.err")) {
        if ($p -and (Test-Path -LiteralPath $p)) {
            Remove-Item -LiteralPath $p -Force -ErrorAction SilentlyContinue
        }
    }
    $argList = [System.Collections.ArrayList]@(
        '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $launcherPath, '-SaveToWCM'
    )
    if (-not [string]::IsNullOrWhiteSpace($TenantId)) {
        $tid = $TenantId.Trim()
        [void]$argList.Add('-TenantId')
        [void]$argList.Add($tid)
    }
    # Interactive console required for browser sign-in and Read-Host (Y/n). Redirecting stdout breaks both.
    $proc = Start-Process -FilePath $psExe -ArgumentList $argList -Wait -PassThru `
        -WorkingDirectory $ProjectRoot
    $result = $null
    if (Test-Path -LiteralPath $resultPath) {
        try {
            $result = Get-Content -LiteralPath $resultPath -Raw -ErrorAction Stop | ConvertFrom-Json
        } catch { }
    }
    return [pscustomobject]@{
        ExitCode = $proc.ExitCode
        Result   = $result
        LogPath  = $logPath
    }
}

function Show-GraphAppCreateResultMessage {
    <#
    .SYNOPSIS
        MessageBox text for Invoke-GraphAppCreateWithWcmSave outcome (Information vs Warning).
    #>
    param(
        [Parameter(Mandatory = $true)]
        $CreateOutcome
    )
    try { Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue } catch { }
    $r = $CreateOutcome.Result
    if ($CreateOutcome.ExitCode -eq 0 -and $r -and $r.WcmSaved) {
        $dn = if ($r.TenantDisplayName) { $r.TenantDisplayName } else { $r.TenantId }
        return @{
            Text = "App created and saved to Windows Credential Manager for:`n$dn`n($($r.TenantId))`n`nSelect this tenant in App reg tenant, then run Graph Auth."
            Title = 'Create Graph App'
            Icon  = [System.Windows.Forms.MessageBoxIcon]::Information
        }
    }
    $detailParts = [System.Collections.Generic.List[string]]::new()
    if ($r -and $r.ScriptError) { [void]$detailParts.Add([string]$r.ScriptError) }
    elseif ($r -and $r.WcmError) { [void]$detailParts.Add([string]$r.WcmError) }
    elseif ($CreateOutcome.ExitCode -eq 2) { [void]$detailParts.Add('Credential Manager save failed or could not be verified.') }
    elseif ($CreateOutcome.ExitCode -eq -1073741510) {
        [void]$detailParts.Add('The Create Graph App console was closed or cancelled before the script finished.')
        [void]$detailParts.Add('Sign in to the correct tenant in the browser, then type Y at "Is this the correct tenant?" and complete any replace (y/n) prompts.')
    }
    else { [void]$detailParts.Add("Script exit code $($CreateOutcome.ExitCode).") }
    $logPath = $CreateOutcome.LogPath
    if ($logPath -and (Test-Path -LiteralPath $logPath)) {
        try {
            $tail = @(Get-Content -LiteralPath $logPath -Tail 6 -ErrorAction SilentlyContinue)
            if ($tail.Count -gt 0) {
                [void]$detailParts.Add('')
                [void]$detailParts.Add('Last output:')
                [void]$detailParts.Add(($tail -join "`n"))
            }
        } catch { }
    }
    $detail = $detailParts -join "`n"
    $tid = if ($r -and $r.TenantId) { "`nTenant: $($r.TenantId)" } else { '' }
    $title = if ($CreateOutcome.ExitCode -eq 2) { 'Create Graph App - WCM save failed' } else { 'Create Graph App failed' }
    $intro = if ($CreateOutcome.ExitCode -eq 2) {
        "The Entra app may exist, but credentials were NOT stored in Windows Credential Manager on this PC.$tid"
    } else {
        "Create Graph App did not finish successfully.$tid"
    }
    return @{
        Text  = "$intro`n`n$detail`n`nIf Graph module errors persist, in pwsh run:`n  Update-Module Microsoft.Graph* -Scope CurrentUser -Force`n  (or: Install-Module Microsoft.Graph -Scope CurrentUser -Force)"
        Title = $title
        Icon  = [System.Windows.Forms.MessageBoxIcon]::Warning
    }
}

function Connect-MgGraphWithWcmApp {
    <#
    .SYNOPSIS
        Connects to Microsoft Graph using app-only client secret credentials from WCM.
    .NOTES
        Prefer this over Connect-MgGraph -AccessToken; UserProvidedTokenCredential is broken in some Graph module versions.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$TenantId,
        [Parameter(Mandatory = $false)]
        [ValidateSet('EOA', 'ESR')]
        [string]$Prefix = 'EOA'
    )
    if (-not (Get-Command Connect-MgGraph -ErrorAction SilentlyContinue)) {
        Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    }
    $cred = Get-GraphAppCredentialFromWCM -TenantId $TenantId -Prefix $Prefix
    if (-not $cred) {
        throw "No Graph app credentials in Windows Credential Manager for tenant $TenantId (prefix $Prefix)."
    }
    $tenantGuid = ($cred.TenantId -replace '[\{\}]', '').Trim()
    if (-not $tenantGuid) { $tenantGuid = ($TenantId -replace '[\{\}]', '').Trim() }
    $sec = ConvertTo-SecureString $cred.ClientSecret -AsPlainText -Force
    $clientCred = New-Object System.Management.Automation.PSCredential($cred.ClientId, $sec)
    $connectParams = @{
        TenantId               = $tenantGuid
        ClientSecretCredential = $clientCred
        NoWelcome              = $true
        ErrorAction            = 'Stop'
    }
    Connect-MgGraph @connectParams | Out-Null
}

function Search-GraphUsersWithWcm {
    <#
    .SYNOPSIS
        Searches Graph users using app-only WCM token and REST (no Connect-MgGraph / Get-MgUser).
    .OUTPUTS
        @(@{ UserPrincipalName; DisplayName }, ...)
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$TenantId,
        [Parameter(Mandatory = $true)]
        [string[]]$SearchTerms,
        [Parameter(Mandatory = $false)]
        [ValidateSet('EOA', 'ESR')]
        [string]$Prefix = 'EOA',
        [Parameter(Mandatory = $false)]
        [string]$AccessToken
    )

    $token = $AccessToken
    if ([string]::IsNullOrWhiteSpace($token)) {
        $token = Get-GraphAppTokenFromWCM -TenantId $TenantId -Prefix $Prefix
    }
    if ([string]::IsNullOrWhiteSpace($token)) {
        throw "Could not obtain Graph app-only token for tenant $TenantId."
    }

    $baseHeaders = @{
        Authorization = "Bearer $token"
        Accept        = 'application/json'
    }
    $allUsers = [System.Collections.ArrayList]::new()
    $seen = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)

    foreach ($searchTerm in $SearchTerms) {
        $term = $searchTerm.Trim()
        if ([string]::IsNullOrWhiteSpace($term)) { continue }

        Write-Host "  Searching for users matching: '$term'" -ForegroundColor Gray
        $batch = @()

        if ($term -match '@') {
            try {
                $enc = [Uri]::EscapeDataString($term)
                $u = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users/$enc`?$select=userPrincipalName,displayName,id" -Headers $baseHeaders -Method GET -ErrorAction Stop
                if ($u -and $u.userPrincipalName) { $batch = @($u); Write-Host '    Found via direct UPN lookup' -ForegroundColor Gray }
            } catch {
                Write-Host "    Direct UPN lookup failed: $($_.Exception.Message)" -ForegroundColor DarkGray
            }
        }

        if ($batch.Count -eq 0) {
            $escaped = $term.Replace("'", "''")
            $filterAttempts = @(
                "startswith(displayName,'$escaped') or startswith(userPrincipalName,'$escaped')",
                "contains(displayName,'$escaped') or contains(userPrincipalName,'$escaped')"
            )
            foreach ($filter in $filterAttempts) {
                if ($batch.Count -gt 0) { break }
                try {
                    $uri = 'https://graph.microsoft.com/v1.0/users?$filter=' + [Uri]::EscapeDataString($filter) + '&$select=userPrincipalName,displayName,id&$top=50'
                    $headers = @{
                        Authorization    = $baseHeaders.Authorization
                        Accept           = $baseHeaders.Accept
                        ConsistencyLevel = 'eventual'
                    }
                    $resp = Invoke-RestMethod -Uri $uri -Headers $headers -Method GET -ErrorAction Stop
                    if ($resp.value -and $resp.value.Count -gt 0) {
                        $batch = @($resp.value)
                        Write-Host "    Found $($batch.Count) user(s) via Graph REST filter" -ForegroundColor Gray
                    }
                } catch {
                    Write-Host "    Graph REST filter failed: $($_.Exception.Message)" -ForegroundColor DarkGray
                }
            }
        }

        foreach ($u in $batch) {
            $upn = [string]$u.userPrincipalName
            if (-not [string]::IsNullOrWhiteSpace($upn) -and $seen.Add($upn)) {
                [void]$allUsers.Add([pscustomobject]@{
                        UserPrincipalName = $upn
                        DisplayName       = [string]$u.displayName
                    })
            }
        }
    }

    return @($allUsers)
}

Export-ModuleMember -Function Get-GraphAppCredentialFromWCM, Save-GraphAppCredentialToWCM, Remove-GraphAppCredentialFromWCM, Get-GraphAppTokenFromWCM, Connect-MgGraphWithWcmApp, Search-GraphUsersWithWcm, Get-WCMTenantIds, Get-TenantDisplayNameFromWCM, Get-WCMTenantListWithNames, Get-WCMTenantListWithNamesForAppRegCombo, Register-GraphAppTenantDisplayNamesInWCM, Export-GraphAppCredentialsToFile, Import-GraphAppCredentialsFromFile, Get-GraphAppCredentialEncryptedFileSummary, Get-WCMUnrecognizedGraphAppTargets, Remove-WCMGraphCredentialTarget, Remove-GraphAppCredentialsLocalOnly, Show-ClearLocalGraphWcmPicker, Reset-GraphAppCredentialManagerImportCache, Invoke-GraphAppCreateWithWcmSave, Show-GraphAppCreateResultMessage
