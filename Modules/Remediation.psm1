# Web-runner containment: Graph REST (revoke / block / unblock) + EXO Restricted Users + inbox rules.
# Graph writes use ExportUtils Set-GraphRestBearerToken / Invoke-GraphRestRequest — no Connect-MgGraph required.

function Get-RemediationJwtPayload {
    param([string]$Token)
    if ([string]::IsNullOrWhiteSpace($Token)) { return $null }
    $parts = $Token.Split('.')
    if ($parts.Count -lt 2) { return $null }
    try {
        $p = $parts[1].Replace('-', '+').Replace('_', '/')
        switch ($p.Length % 4) {
            2 { $p += '==' }
            3 { $p += '=' }
        }
        $json = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($p))
        return $json | ConvertFrom-Json
    } catch {
        return $null
    }
}

function Get-GraphContainmentCapabilities {
    <#
    .SYNOPSIS
        Inspect the current Graph REST bearer token for revoke / enable-disable scopes (roles or scp).
    #>
    $result = [ordered]@{
        canRevoke        = $null
        canBlock         = $null
        canAuthWrite     = $null
        canDeviceDelete  = $null
        canAppWrite      = $null
        canPasswordReset = $null
        canOauthWrite    = $null
        canIntune        = $null
        canIntuneWipe    = $null
        canRoles         = $null
        canGroupWrite    = $null
        reason           = ''
        scopes           = @()
    }
    $token = $null
    if (Get-Command Test-GraphRestBearerToken -ErrorAction SilentlyContinue) {
        if (-not (Test-GraphRestBearerToken)) {
            $result.canRevoke = $false
            $result.canBlock = $false
            $result.canAuthWrite = $false
            $result.canDeviceDelete = $false
            $result.canAppWrite = $false
            $result.canPasswordReset = $false
            $result.canOauthWrite = $false
            $result.canIntune = $false
            $result.canIntuneWipe = $false
            $result.canRoles = $false
            $result.canGroupWrite = $false
            $result.reason = 'Graph REST bearer token is not set. Complete Graph Auth, then retry.'
            return [pscustomobject]$result
        }
    }
    try {
        if (Get-Command Get-GraphRestBearerToken -ErrorAction SilentlyContinue) {
            $token = Get-GraphRestBearerToken
        }
    } catch {}
    if ([string]::IsNullOrWhiteSpace($token)) {
        $result.reason = 'Token present but claims could not be read. Graph buttons stay enabled until the API returns an error.'
        return [pscustomobject]$result
    }
    $payload = Get-RemediationJwtPayload -Token $token
    if (-not $payload) {
        $result.reason = 'Could not decode Graph token claims. Graph buttons stay enabled until the API returns an error.'
        return [pscustomobject]$result
    }
    $claimed = [System.Collections.Generic.List[string]]::new()
    if ($payload.roles) {
        foreach ($r in @($payload.roles)) { if ($r) { [void]$claimed.Add([string]$r) } }
    }
    if ($payload.scp) {
        foreach ($s in ([string]$payload.scp -split '\s+')) { if ($s) { [void]$claimed.Add($s) } }
    }
    $result.scopes = @($claimed)
    $hasRevoke = $claimed -contains 'User.RevokeSessions.All' -or $claimed -contains 'User.ReadWrite.All'
    $hasBlock = $claimed -contains 'User.EnableDisableAccount.All' -or $claimed -contains 'User.ReadWrite.All'
    $hasAuthWrite = $claimed -contains 'UserAuthenticationMethod.ReadWrite.All'
    $hasDeviceDelete = $claimed -contains 'Device.ReadWrite.All' -or $claimed -contains 'Directory.AccessAsUser.All'
    $hasAppWrite = $claimed -contains 'Application.ReadWrite.All'
    $hasPasswordReset = $claimed -contains 'User-PasswordProfile.ReadWrite.All' -or $claimed -contains 'User.ReadWrite.All' -or $claimed -contains 'Directory.AccessAsUser.All'
    $hasOauthWrite = $claimed -contains 'DelegatedPermissionGrant.ReadWrite.All' -or $claimed -contains 'Directory.ReadWrite.All' -or $claimed -contains 'Directory.AccessAsUser.All'
    $hasIntune = $claimed -contains 'DeviceManagementManagedDevices.ReadWrite.All' -or $claimed -contains 'DeviceManagementManagedDevices.Read.All' -or $claimed -contains 'DeviceManagementManagedDevices.PrivilegedOperations.All'
    $hasIntuneWipe = $claimed -contains 'DeviceManagementManagedDevices.PrivilegedOperations.All'
    $hasRoles = $claimed -contains 'RoleManagement.ReadWrite.Directory' -or $claimed -contains 'Directory.AccessAsUser.All'
    $hasGroupWrite = $claimed -contains 'GroupMember.ReadWrite.All' -or $claimed -contains 'Directory.AccessAsUser.All' -or $claimed -contains 'Directory.ReadWrite.All'
    $result.canRevoke = [bool]$hasRevoke
    $result.canBlock = [bool]$hasBlock
    $result.canAuthWrite = [bool]$hasAuthWrite
    $result.canDeviceDelete = [bool]$hasDeviceDelete
    $result.canAppWrite = [bool]$hasAppWrite
    $result.canPasswordReset = [bool]$hasPasswordReset
    $result.canOauthWrite = [bool]$hasOauthWrite
    $result.canIntune = [bool]$hasIntune
    $result.canIntuneWipe = [bool]$hasIntuneWipe
    $result.canRoles = [bool]$hasRoles
    $result.canGroupWrite = [bool]$hasGroupWrite
    $missing = @()
    if (-not $hasRevoke) { $missing += 'User.RevokeSessions.All' }
    if (-not $hasBlock) { $missing += 'User.EnableDisableAccount.All' }
    if (-not $hasAuthWrite) { $missing += 'UserAuthenticationMethod.ReadWrite.All' }
    if (-not $hasDeviceDelete) { $missing += 'Device.ReadWrite.All (WCM) or Directory.AccessAsUser.All (interactive)' }
    if (-not $hasAppWrite) { $missing += 'Application.ReadWrite.All' }
    if (-not $hasPasswordReset) { $missing += 'User-PasswordProfile.ReadWrite.All' }
    if (-not $hasOauthWrite) { $missing += 'DelegatedPermissionGrant.ReadWrite.All' }
    if (-not $hasIntune) { $missing += 'DeviceManagementManagedDevices.Read.All' }
    if (-not $hasIntuneWipe) { $missing += 'DeviceManagementManagedDevices.PrivilegedOperations.All (Intune wipe/retire)' }
    if (-not $hasRoles) { $missing += 'RoleManagement.ReadWrite.Directory' }
    if (-not $hasGroupWrite) { $missing += 'GroupMember.ReadWrite.All' }
    if ($missing.Count -gt 0) {
        $result.reason = "Graph token lacks $($missing -join '; '). For interactive Graph, sign in again so the new scopes can be consented. For WCM app-only, use Update Graph App scopes (App registrations), finish admin consent in the console, then Graph Auth again. Recreate the app only if it is missing."
    }
    return [pscustomobject]$result
}

function Convert-RemediationExoAddressList {
    param($Value)
    if ($null -eq $Value) { return '' }
    $items = @($Value)
    $parts = foreach ($item in $items) {
        if ($null -eq $item) { continue }
        if ($item.Address) { [string]$item.Address }
        elseif ($item.PrimarySmtpAddress) { [string]$item.PrimarySmtpAddress }
        else { [string]$item }
    }
    return (($parts | Where-Object { $_ }) -join '; ')
}

function Write-RemediationAuditCsv {
    param(
        [string]$Folder,
        [string]$Tenant,
        [int]$ClientNumber,
        [string]$Upn,
        [string]$Action,
        [string]$Result,
        [string]$Detail
    )
    if ([string]::IsNullOrWhiteSpace($Folder)) { return $null }
    try {
        if (-not (Test-Path -LiteralPath $Folder)) {
            New-Item -ItemType Directory -Path $Folder -Force | Out-Null
        }
        $path = Join-Path $Folder 'Remediation.csv'
        $row = [pscustomobject]@{
            Timestamp    = (Get-Date).ToString('o')
            Tenant       = $Tenant
            ClientNumber = $ClientNumber
            UPN          = $Upn
            Action       = $Action
            Result       = $Result
            Detail       = $Detail
        }
        $row | Export-Csv -LiteralPath $path -NoTypeInformation -Append -Encoding UTF8
    } catch {
        Write-Warning "Remediation audit CSV failed: $($_.Exception.Message)"
    }
}

function Get-RemediationGraphUser {
    param([Parameter(Mandatory = $true)][string]$UserPrincipalName)
    if (-not (Get-Command Invoke-GraphRestRequest -ErrorAction SilentlyContinue)) {
        throw 'Invoke-GraphRestRequest is not available. Import ExportUtils in the worker.'
    }
    $enc = [Uri]::EscapeDataString($UserPrincipalName.Trim())
    return Invoke-GraphRestRequest -Uri "https://graph.microsoft.com/v1.0/users/$enc`?`$select=id,userPrincipalName,displayName,accountEnabled"
}

function Invoke-RemediationRevokeSessions {
    param([Parameter(Mandatory = $true)][string]$UserPrincipalName)
    $user = Get-RemediationGraphUser -UserPrincipalName $UserPrincipalName
    $id = if ($user.id) { [string]$user.id } else { $UserPrincipalName.Trim() }
    $enc = [Uri]::EscapeDataString($id)
    $resp = Invoke-GraphRestRequest -Uri "https://graph.microsoft.com/v1.0/users/$enc/revokeSignInSessions" -Method POST -Body '{}'
    return [pscustomobject]@{
        UserPrincipalName = $user.userPrincipalName
        Id                = $id
        AccountEnabled    = $user.accountEnabled
        Value             = $resp.value
    }
}

function Invoke-RemediationSetAccountEnabled {
    param(
        [Parameter(Mandatory = $true)][string]$UserPrincipalName,
        [Parameter(Mandatory = $true)][bool]$Enabled
    )
    $user = Get-RemediationGraphUser -UserPrincipalName $UserPrincipalName
    $id = if ($user.id) { [string]$user.id } else { $UserPrincipalName.Trim() }
    $enc = [Uri]::EscapeDataString($id)
    $body = @{ accountEnabled = $Enabled } | ConvertTo-Json -Compress
    Invoke-GraphRestRequest -Uri "https://graph.microsoft.com/v1.0/users/$enc" -Method PATCH -Body $body | Out-Null
    $after = $null
    try { $after = Get-RemediationGraphUser -UserPrincipalName $id } catch {}
    return [pscustomobject]@{
        UserPrincipalName = $user.userPrincipalName
        Id                = $id
        AccountEnabled    = if ($after) { $after.accountEnabled } else { $Enabled }
        PreviousEnabled   = $user.accountEnabled
    }
}

function Get-RemediationRandomChar {
    param([string]$Chars, [System.Security.Cryptography.RandomNumberGenerator]$Rng)
    $b = [byte[]]::new(4)
    $Rng.GetBytes($b)
    $n = [BitConverter]::ToUInt32($b, 0)
    return $Chars[$n % $Chars.Length]
}

function New-RemediationRandomPassword {
    $lower = 'abcdefghijkmnopqrstuvwxyz'
    $upper = 'ABCDEFGHJKLMNPQRSTUVWXYZ'
    $digits = '23456789'
    $symbols = '!@#$%^&*-_+'
    $all = $lower + $upper + $digits + $symbols
    $rng = [System.Security.Cryptography.RandomNumberGenerator]::Create()
    $chars = [System.Collections.Generic.List[char]]::new()
    [void]$chars.Add((Get-RemediationRandomChar -Chars $lower -Rng $rng))
    [void]$chars.Add((Get-RemediationRandomChar -Chars $upper -Rng $rng))
    [void]$chars.Add((Get-RemediationRandomChar -Chars $digits -Rng $rng))
    [void]$chars.Add((Get-RemediationRandomChar -Chars $symbols -Rng $rng))
    for ($i = 0; $i -lt 16; $i++) {
        [void]$chars.Add((Get-RemediationRandomChar -Chars $all -Rng $rng))
    }
    $bytes = [byte[]]::new($chars.Count)
    $rng.GetBytes($bytes)
    for ($i = $chars.Count - 1; $i -gt 0; $i--) {
        $j = $bytes[$i] % ($i + 1)
        $tmp = $chars[$i]
        $chars[$i] = $chars[$j]
        $chars[$j] = $tmp
    }
    $rng.Dispose()
    return (-join $chars)
}

function Get-RemediationSsprUrl {
    param([string]$UserPrincipalName)
    $domain = ''
    if ($UserPrincipalName -match '@(.+)$') { $domain = $Matches[1].Trim() }
    if ($domain) {
        return "https://passwordreset.microsoftonline.com/?whr=$([Uri]::EscapeDataString($domain))"
    }
    return 'https://passwordreset.microsoftonline.com/'
}

function Invoke-RemediationResetPassword {
    param(
        [Parameter(Mandatory = $true)][string]$UserPrincipalName,
        [string]$Password
    )
    $user = Get-RemediationGraphUser -UserPrincipalName $UserPrincipalName
    $id = if ($user.id) { [string]$user.id } else { $UserPrincipalName.Trim() }
    $enc = [Uri]::EscapeDataString($id)
    if ([string]::IsNullOrWhiteSpace($Password)) {
        $Password = New-RemediationRandomPassword
    }
    $body = @{
        passwordProfile = @{
            password = $Password
            forceChangePasswordNextSignIn = $false
        }
    } | ConvertTo-Json -Compress -Depth 4
    Invoke-GraphRestRequest -Uri "https://graph.microsoft.com/v1.0/users/$enc" -Method PATCH -Body $body | Out-Null
    return [pscustomobject]@{
        UserPrincipalName = [string]$user.userPrincipalName
        SsprUrl           = Get-RemediationSsprUrl -UserPrincipalName $user.userPrincipalName
    }
}

function Invoke-RemediationGraphDelete {
    param([Parameter(Mandatory = $true)][string]$Uri)
    if (-not (Get-Command Get-GraphRestBearerToken -ErrorAction SilentlyContinue)) {
        throw 'Graph REST bearer token is not set.'
    }
    $token = Get-GraphRestBearerToken
    if ([string]::IsNullOrWhiteSpace($token)) { throw 'Graph REST bearer token is not set.' }
    $resp = Invoke-WebRequest -Uri $Uri -Method DELETE -Headers @{
        Authorization = "Bearer $token"
        Accept        = 'application/json'
    } -UseBasicParsing
    if ($resp.StatusCode -lt 200 -or $resp.StatusCode -ge 300) {
        throw "Graph DELETE failed ($($resp.StatusCode)): $($resp.Content)"
    }
}

function Get-RemediationAuthMethodMeta {
    param([string]$ODataType)
    $t = [string]$ODataType
    switch ($t) {
        '#microsoft.graph.microsoftAuthenticatorAuthenticationMethod' { return [pscustomobject]@{ Segment = 'microsoftAuthenticatorMethods'; Label = 'Authenticator' } }
        '#microsoft.graph.phoneAuthenticationMethod' { return [pscustomobject]@{ Segment = 'phoneMethods'; Label = 'Phone' } }
        '#microsoft.graph.emailAuthenticationMethod' { return [pscustomobject]@{ Segment = 'emailMethods'; Label = 'Email' } }
        '#microsoft.graph.fido2AuthenticationMethod' { return [pscustomobject]@{ Segment = 'fido2Methods'; Label = 'FIDO2' } }
        '#microsoft.graph.softwareOathAuthenticationMethod' { return [pscustomobject]@{ Segment = 'softwareOathMethods'; Label = 'Software OATH' } }
        '#microsoft.graph.windowsHelloForBusinessAuthenticationMethod' { return [pscustomobject]@{ Segment = 'windowsHelloForBusinessMethods'; Label = 'Windows Hello' } }
        '#microsoft.graph.temporaryAccessPassAuthenticationMethod' { return [pscustomobject]@{ Segment = 'temporaryAccessPassMethods'; Label = 'Temporary Access Pass' } }
        '#microsoft.graph.platformCredentialAuthenticationMethod' { return [pscustomobject]@{ Segment = 'platformCredentialMethods'; Label = 'Platform credential' } }
        '#microsoft.graph.hardwareOathAuthenticationMethod' { return [pscustomobject]@{ Segment = 'hardwareOathMethods'; Label = 'Hardware OATH' } }
        default { return [pscustomobject]@{ Segment = $null; Label = if ($t -match 'password') { 'Password' } else { ($t -replace '#microsoft\.graph\.', '' -replace 'AuthenticationMethod$', '') } } }
    }
}

function Get-RemediationAuthMethodDetails {
    param($Method)
    $parts = @(
        $Method.displayName
        $Method.phoneNumber
        $Method.phoneType
        $Method.emailAddress
        $Method.deviceTag
        $Method.model
        if ($Method.createdDateTime) { "created $($Method.createdDateTime)" }
    ) | Where-Object { $_ }
    return ($parts -join ' · ')
}

function Get-RemediationAuthMethods {
    param([Parameter(Mandatory = $true)][string[]]$UserPrincipalNames)
    foreach ($upn in $UserPrincipalNames) {
        $user = $null
        try {
            $user = Get-RemediationGraphUser -UserPrincipalName $upn
        } catch {
            [pscustomobject]@{
                UserPrincipalName = $upn
                Id                = ''
                ODataType         = ''
                Type              = 'Error'
                Details           = $_.Exception.Message
                CanDelete         = $false
                Error             = (Get-RemediationGraphErrorText $_)
            }
            continue
        }
        $id = if ($user.id) { [string]$user.id } else { $upn.Trim() }
        $enc = [Uri]::EscapeDataString($id)
        $methods = @()
        try {
            $methods = @(Invoke-GraphRestPaged -Uri "https://graph.microsoft.com/v1.0/users/$enc/authentication/methods")
        } catch {
            [pscustomobject]@{
                UserPrincipalName = [string]$user.userPrincipalName
                Id                = ''
                ODataType         = ''
                Type              = 'Error'
                Details           = $_.Exception.Message
                CanDelete         = $false
                Error             = (Get-RemediationGraphErrorText $_)
            }
            continue
        }
        foreach ($m in $methods) {
            $meta = Get-RemediationAuthMethodMeta -ODataType $m.'@odata.type'
            [pscustomobject]@{
                UserPrincipalName = [string]$user.userPrincipalName
                Id                = [string]$m.id
                ODataType         = [string]$m.'@odata.type'
                Type              = $meta.Label
                Details           = Get-RemediationAuthMethodDetails -Method $m
                CanDelete         = [bool]$meta.Segment
                Error             = ''
            }
        }
    }
}

function Remove-RemediationAuthMethods {
    param([Parameter(Mandatory = $true)][object[]]$Items)
    foreach ($item in $Items) {
        $upn = if ($item.UserPrincipalName) { [string]$item.UserPrincipalName } elseif ($item.User) { [string]$item.User } else { '' }
        $methodId = if ($item.Id) { [string]$item.Id } else { '' }
        $odata = if ($item.ODataType) { [string]$item.ODataType } else { '' }
        if ([string]::IsNullOrWhiteSpace($upn) -or [string]::IsNullOrWhiteSpace($methodId)) { continue }
        $meta = Get-RemediationAuthMethodMeta -ODataType $odata
        if (-not $meta.Segment) {
            [pscustomobject]@{ UserPrincipalName = $upn; Id = $methodId; Success = $false; Error = "Cannot delete $($meta.Label) methods via Graph" }
            continue
        }
        try {
            $user = Get-RemediationGraphUser -UserPrincipalName $upn
            $uid = if ($user.id) { [string]$user.id } else { $upn.Trim() }
            $encUser = [Uri]::EscapeDataString($uid)
            $encMethod = [Uri]::EscapeDataString($methodId)
            Invoke-RemediationGraphDelete -Uri "https://graph.microsoft.com/v1.0/users/$encUser/authentication/$($meta.Segment)/$encMethod"
            [pscustomobject]@{ UserPrincipalName = $upn; Id = $methodId; Success = $true; Error = '' }
        } catch {
            [pscustomobject]@{ UserPrincipalName = $upn; Id = $methodId; Success = $false; Error = (Get-RemediationGraphErrorText $_) }
        }
    }
}

function Get-RemediationUserDevices {
    param([Parameter(Mandatory = $true)][string[]]$UserPrincipalNames)
    foreach ($upn in $UserPrincipalNames) {
        $user = $null
        try {
            $user = Get-RemediationGraphUser -UserPrincipalName $upn
        } catch {
            [pscustomobject]@{
                UserPrincipalName = $upn
                Id                = ''
                DisplayName       = ''
                OperatingSystem   = ''
                TrustType         = ''
                Relation          = 'Error'
                LastSignIn        = ''
                Enabled           = $null
                Error             = (Get-RemediationGraphErrorText $_)
            }
            continue
        }
        $uid = if ($user.id) { [string]$user.id } else { $upn.Trim() }
        $enc = [Uri]::EscapeDataString($uid)
        $select = '$select=id,displayName,operatingSystem,operatingSystemVersion,trustType,approximateLastSignInDateTime,deviceId,accountEnabled'
        $byId = @{}
        foreach ($rel in @('registeredDevices', 'ownedDevices')) {
            $rows = @()
            try {
                $rows = @(Invoke-GraphRestPaged -Uri "https://graph.microsoft.com/v1.0/users/$enc/$rel`?$select")
            } catch {
                continue
            }
            $label = if ($rel -eq 'registeredDevices') { 'Registered' } else { 'Owned' }
            foreach ($d in $rows) {
                $did = [string]$d.id
                if (-not $did) { continue }
                if ($byId.ContainsKey($did)) {
                    if ($byId[$did].Relation -notmatch [regex]::Escape($label)) {
                        $byId[$did].Relation = "$($byId[$did].Relation)+$label"
                    }
                    continue
                }
                $os = [string]$d.operatingSystem
                if ($d.operatingSystemVersion) { $os = "$os $($d.operatingSystemVersion)".Trim() }
                $byId[$did] = [pscustomobject]@{
                    UserPrincipalName = [string]$user.userPrincipalName
                    Id                = $did
                    DisplayName       = [string]$d.displayName
                    OperatingSystem   = $os
                    TrustType         = [string]$d.trustType
                    Relation          = $label
                    LastSignIn        = if ($d.approximateLastSignInDateTime) { [string]$d.approximateLastSignInDateTime } else { '' }
                    Enabled           = [bool]$d.accountEnabled
                    Error             = ''
                }
            }
        }
        if ($byId.Count -eq 0) {
            [pscustomobject]@{
                UserPrincipalName = [string]$user.userPrincipalName
                Id                = ''
                DisplayName       = '(none)'
                OperatingSystem   = ''
                TrustType         = ''
                Relation          = ''
                LastSignIn        = ''
                Enabled           = $null
                Error             = ''
            }
        } else {
            foreach ($row in $byId.Values) { $row }
        }
    }
}

function Remove-RemediationDevices {
    param([Parameter(Mandatory = $true)][object[]]$Items)
    foreach ($item in $Items) {
        $id = if ($item.Id) { [string]$item.Id } elseif ($item.DeviceId) { [string]$item.DeviceId } else { [string]$item }
        if ([string]::IsNullOrWhiteSpace($id)) { continue }
        try {
            $enc = [Uri]::EscapeDataString($id)
            Invoke-RemediationGraphDelete -Uri "https://graph.microsoft.com/v1.0/devices/$enc"
            [pscustomobject]@{ Id = $id; Success = $true; Error = '' }
        } catch {
            [pscustomobject]@{ Id = $id; Success = $false; Error = (Get-RemediationGraphErrorText $_) }
        }
    }
}

function Get-RemediationDirectoryApps {
    $msft = 'f8cdef31-a31e-4b4a-93e4-5f571e91255a'
    $tenantId = ''
    try {
        $org = Invoke-GraphRestRequest -Uri 'https://graph.microsoft.com/v1.0/organization?$select=id'
        $tenantId = [string]@($org.value)[0].id
    } catch {}
    $rows = [System.Collections.ArrayList]::new()
    try {
        $apps = @(Invoke-GraphRestPaged -Uri 'https://graph.microsoft.com/v1.0/applications?$select=id,appId,displayName,createdDateTime,signInAudience,publisherDomain')
        foreach ($app in $apps) {
            [void]$rows.Add([pscustomobject]@{
                Kind        = 'AppRegistration'
                Id          = [string]$app.id
                AppId       = [string]$app.appId
                DisplayName = [string]$app.displayName
                Created     = if ($app.createdDateTime) { [string]$app.createdDateTime } else { '' }
                Publisher   = [string]$app.publisherDomain
                Audience    = [string]$app.signInAudience
                OwnerTenant = $tenantId
            })
        }
    } catch {
        throw
    }
    try {
        $sps = @(Invoke-GraphRestPaged -Uri 'https://graph.microsoft.com/v1.0/servicePrincipals?$select=id,appId,displayName,createdDateTime,appOwnerOrganizationId,servicePrincipalType,accountEnabled,publisherName&$filter=servicePrincipalType eq ''Application''')
        foreach ($sp in $sps) {
            $owner = [string]$sp.appOwnerOrganizationId
            if ($owner -eq $msft) { continue }
            if ($tenantId -and $owner -eq $tenantId) { continue }
            [void]$rows.Add([pscustomobject]@{
                Kind        = 'EnterpriseApp'
                Id          = [string]$sp.id
                AppId       = [string]$sp.appId
                DisplayName = [string]$sp.displayName
                Created     = if ($sp.createdDateTime) { [string]$sp.createdDateTime } else { '' }
                Publisher   = if ($sp.publisherName) { [string]$sp.publisherName } else { $owner }
                Audience    = 'external-tenant'
                OwnerTenant = $owner
            })
        }
    } catch {}
    return @($rows)
}

function Remove-RemediationDirectoryApps {
    param([Parameter(Mandatory = $true)][object[]]$Items)
    foreach ($item in $Items) {
        $id = if ($item.Id) { [string]$item.Id } else { '' }
        $kind = if ($item.Kind) { [string]$item.Kind } else { 'AppRegistration' }
        if ([string]::IsNullOrWhiteSpace($id)) { continue }
        $path = if ($kind -ieq 'EnterpriseApp' -or $kind -ieq 'ServicePrincipal') { 'servicePrincipals' } else { 'applications' }
        try {
            $enc = [Uri]::EscapeDataString($id)
            Invoke-RemediationGraphDelete -Uri "https://graph.microsoft.com/v1.0/$path/$enc"
            [pscustomobject]@{ Id = $id; Kind = $kind; Success = $true; Error = '' }
        } catch {
            [pscustomobject]@{ Id = $id; Kind = $kind; Success = $false; Error = (Get-RemediationGraphErrorText $_) }
        }
    }
}

function Get-RemediationRestrictedEmailStatus {
    param([Parameter(Mandatory = $true)][string[]]$UserPrincipalNames)
    if (-not (Get-Command Get-BlockedSenderAddress -ErrorAction SilentlyContinue)) {
        throw 'Get-BlockedSenderAddress is not available. Complete Exchange Auth (Security/Exchange admin).'
    }
    $results = foreach ($upn in $UserPrincipalNames) {
        $needle = $upn.Trim()
        $hit = $null
        try {
            $direct = @(Get-BlockedSenderAddress -SenderAddress $needle -ErrorAction Stop)
            if ($direct.Count -gt 0) { $hit = $direct[0] }
        } catch {
            try {
                $direct = @(Get-BlockedSenderAddress -Identity $needle -ErrorAction Stop)
                if ($direct.Count -gt 0) { $hit = $direct[0] }
            } catch {}
        }
        [pscustomobject]@{
            UserPrincipalName = $needle
            Restricted        = [bool]$hit
            SenderAddress     = if ($hit) { [string]$hit.SenderAddress } else { '' }
            Reason            = if ($hit) { [string]$hit.Reason } else { '' }
            CreatedDateTime   = if ($hit -and $hit.CreatedDateTime) { [string]$hit.CreatedDateTime } elseif ($hit -and $hit.CreatedDatetime) { [string]$hit.CreatedDatetime } else { '' }
        }
    }
    return @($results)
}

function Invoke-RemediationUnrestrictEmail {
    param([Parameter(Mandatory = $true)][string]$SenderAddress)
    if (-not (Get-Command Remove-BlockedSenderAddress -ErrorAction SilentlyContinue)) {
        throw 'Remove-BlockedSenderAddress is not available. Complete Exchange Auth (Security/Exchange admin).'
    }
    $addr = $SenderAddress.Trim()
    Remove-BlockedSenderAddress -SenderAddress $addr -Confirm:$false -ErrorAction Stop
    return [pscustomobject]@{
        SenderAddress = $addr
        Removed       = $true
    }
}

function Get-RemediationInboxRules {
    param([Parameter(Mandatory = $true)][string]$Mailbox)
    if (-not (Get-Command Get-InboxRule -ErrorAction SilentlyContinue)) {
        throw 'Get-InboxRule is not available. Complete Exchange Auth.'
    }
    $rules = @()
    try {
        $rules = @(Get-InboxRule -Mailbox $Mailbox -IncludeHidden -ErrorAction Stop)
    } catch {
        $rules = @(Get-InboxRule -Mailbox $Mailbox -ErrorAction Stop)
    }
    foreach ($rule in $rules) {
        $identity = if ($rule.Name) { [string]$rule.Name } else { [string]$rule.Identity }
        [pscustomobject]@{
            Identity      = $identity
            Name          = [string]$rule.Name
            Enabled       = [bool]$rule.Enabled
            Priority      = $rule.Priority
            Hidden        = [bool]$rule.Hidden
            Description   = [string]$rule.Description
            RedirectTo    = Convert-RemediationExoAddressList $rule.RedirectTo
            ForwardTo     = Convert-RemediationExoAddressList $rule.ForwardTo
            ForwardAsAttachmentTo = Convert-RemediationExoAddressList $rule.ForwardAsAttachmentTo
            DeleteMessage = [bool]$rule.DeleteMessage
            From          = Convert-RemediationExoAddressList $rule.From
            SentTo        = Convert-RemediationExoAddressList $rule.SentTo
            RuleIdentity  = if ($rule.RuleIdentity) { [string]$rule.RuleIdentity } else { '' }
        }
    }
}

function Remove-RemediationInboxRules {
    param(
        [Parameter(Mandatory = $true)][string]$Mailbox,
        [Parameter(Mandatory = $true)][string[]]$Identities
    )
    if (-not (Get-Command Remove-InboxRule -ErrorAction SilentlyContinue)) {
        throw 'Remove-InboxRule is not available. Complete Exchange Auth.'
    }
    $results = foreach ($id in $Identities) {
        $name = [string]$id
        if ([string]::IsNullOrWhiteSpace($name)) { continue }
        try {
            Remove-InboxRule -Mailbox $Mailbox -Identity $name -Confirm:$false -ErrorAction Stop
            [pscustomobject]@{ Identity = $name; Success = $true; Error = '' }
        } catch {
            [pscustomobject]@{ Identity = $name; Success = $false; Error = $_.Exception.Message }
        }
    }
    return @($results)
}

function Convert-RemediationUsersFromCommand {
    param([string]$Command)
    if ($Command -match '\|USERS:(\[.*?\])(?:\||$)') {
        $raw = $Matches[1]
    } elseif ($Command -match '\|USERS:(.+)$') {
        $raw = $Matches[1]
        if ($raw -match '^(\[.*\])') { $raw = $Matches[1] }
    } else {
        return @()
    }
    try {
        $parsed = $raw | ConvertFrom-Json -ErrorAction Stop
        if ($parsed -is [array]) {
            return @($parsed | ForEach-Object { [string]$_ } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        }
        if ($parsed -is [string] -and -not [string]::IsNullOrWhiteSpace($parsed)) {
            return @([string]$parsed)
        }
        return @($parsed | ForEach-Object { [string]$_ } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    } catch {
        return @($raw -split ',' | ForEach-Object { $_.Trim().Trim('"') } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    }
}

function Convert-RemediationRuleIdentitiesFromCommand {
    param([string]$Command)
    if ($Command -notmatch '\|RULES:(.+)$') { return @() }
    $raw = $Matches[1]
    try {
        $parsed = $raw | ConvertFrom-Json -ErrorAction Stop
        $list = @($parsed)
        $ids = foreach ($item in $list) {
            if ($null -eq $item) { continue }
            if ($item -is [string]) { $item }
            elseif ($item.Identity) { [string]$item.Identity }
            elseif ($item.Name) { [string]$item.Name }
            else { [string]$item }
        }
        return @($ids | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    } catch {
        return @($raw -split ',' | ForEach-Object { $_.Trim().Trim('"') } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    }
}

function Get-RemediationCommandToken {
    param([string]$Command, [string]$Name)
    if ($Command -match "\|${Name}:([^|]+)") { return $Matches[1].Trim() }
    return $null
}

function Convert-RemediationRecipientName {
    param($Value)
    if ($null -eq $Value) { return '' }
    if ($Value.PrimarySmtpAddress) { return [string]$Value.PrimarySmtpAddress }
    if ($Value.UserPrincipalName) { return [string]$Value.UserPrincipalName }
    return [string]$Value
}

function Test-RemediationAccessRight {
    param($AccessRights, [string]$Name)
    foreach ($r in @($AccessRights)) {
        if ([string]$r -eq $Name -or [string]$r -match [regex]::Escape($Name)) { return $true }
    }
    return $false
}

function Convert-RemediationSmtpAddress {
    param($Value)
    $s = [string]$Value
    if ([string]::IsNullOrWhiteSpace($s)) { return '' }
    if ($s -match '^(?i)smtp:(.+)$') { return $Matches[1].Trim() }
    return $s.Trim()
}

function Get-RemediationMailboxAccess {
    param([Parameter(Mandatory = $true)][string[]]$UserPrincipalNames)
    foreach ($upn in $UserPrincipalNames) {
        $mbx = Get-Mailbox -Identity $upn -ErrorAction Stop
        $full = @()
        try {
            $full = @(Get-MailboxPermission -Identity $upn -ErrorAction SilentlyContinue |
                Where-Object { -not $_.IsInherited -and (Test-RemediationAccessRight $_.AccessRights 'FullAccess') -and [string]$_.User -notmatch 'NT AUTHORITY|S-1-|SELF' } |
                ForEach-Object { [string]$_.User })
        } catch {}
        $sendAs = @()
        try {
            $sendAs = @(Get-RecipientPermission -Identity $upn -ErrorAction SilentlyContinue |
                Where-Object { (Test-RemediationAccessRight $_.AccessRights 'SendAs') -and [string]$_.Trustee -notmatch 'NT AUTHORITY|S-1-|SELF' } |
                ForEach-Object { [string]$_.Trustee })
        } catch {}
        $sob = @()
        if ($mbx.GrantSendOnBehalfTo) {
            $sob = @($mbx.GrantSendOnBehalfTo | ForEach-Object { Convert-RemediationRecipientName $_ } | Where-Object { $_ })
        }
        $delegates = [System.Collections.ArrayList]::new()
        foreach ($u in @($full | Select-Object -Unique)) { [void]$delegates.Add([pscustomobject]@{ User = $u; Right = 'FullAccess' }) }
        foreach ($u in @($sendAs | Select-Object -Unique)) { [void]$delegates.Add([pscustomobject]@{ User = $u; Right = 'SendAs' }) }
        foreach ($u in @($sob | Select-Object -Unique)) { [void]$delegates.Add([pscustomobject]@{ User = $u; Right = 'SendOnBehalf' }) }
        $forwards = [System.Collections.ArrayList]::new()
        $smtp = Convert-RemediationSmtpAddress $mbx.ForwardingSmtpAddress
        $recip = Convert-RemediationRecipientName $mbx.ForwardingAddress
        if ($smtp) {
            [void]$forwards.Add([pscustomobject]@{ Field = 'Smtp'; Address = $smtp })
        }
        if ($recip) {
            [void]$forwards.Add([pscustomobject]@{ Field = 'Recipient'; Address = $recip })
        }
        [pscustomobject]@{
            UserPrincipalName          = [string]$mbx.UserPrincipalName
            ForwardingAddress          = $recip
            ForwardingSmtpAddress      = $smtp
            DeliverToMailboxAndForward = [bool]$mbx.DeliverToMailboxAndForward
            Forwards                   = @($forwards)
            Delegates                  = @($delegates)
            Error                      = ''
        }
    }
}

function Set-RemediationMailboxForwarding {
    param(
        [Parameter(Mandatory = $true)][string]$Mailbox,
        [string]$SmtpAddress,
        [bool]$DeliverToMailboxAndForward = $true
    )
    if ([string]::IsNullOrWhiteSpace($SmtpAddress)) {
        Set-Mailbox -Identity $Mailbox -ForwardingAddress $null -ForwardingSmtpAddress $null -DeliverToMailboxAndForward $false -ErrorAction Stop
    } else {
        Set-Mailbox -Identity $Mailbox -ForwardingSmtpAddress $SmtpAddress.Trim() -DeliverToMailboxAndForward $DeliverToMailboxAndForward -ErrorAction Stop
    }
}

function Clear-RemediationMailboxForwardingField {
    param(
        [Parameter(Mandatory = $true)][string]$Mailbox,
        [Parameter(Mandatory = $true)][ValidateSet('Smtp', 'Recipient')][string]$Field
    )
    if ($Field -eq 'Smtp') {
        Set-Mailbox -Identity $Mailbox -ForwardingSmtpAddress $null -ErrorAction Stop
    } else {
        Set-Mailbox -Identity $Mailbox -ForwardingAddress $null -ErrorAction Stop
    }
    $after = Get-Mailbox -Identity $Mailbox -ErrorAction SilentlyContinue
    if ($after -and -not $after.ForwardingAddress -and -not $after.ForwardingSmtpAddress) {
        Set-Mailbox -Identity $Mailbox -DeliverToMailboxAndForward $false -ErrorAction SilentlyContinue
    }
}

function Add-RemediationMailboxDelegation {
    param(
        [Parameter(Mandatory = $true)][string]$Mailbox,
        [Parameter(Mandatory = $true)][string]$Delegate,
        [Parameter(Mandatory = $true)][ValidateSet('FullAccess', 'SendAs', 'SendOnBehalf')][string]$Right
    )
    $del = $Delegate.Trim()
    switch ($Right) {
        'FullAccess' {
            Add-MailboxPermission -Identity $Mailbox -User $del -AccessRights FullAccess -InheritanceType All -AutoMapping:$false -Confirm:$false -ErrorAction Stop | Out-Null
        }
        'SendAs' {
            Add-RecipientPermission -Identity $Mailbox -Trustee $del -AccessRights SendAs -Confirm:$false -ErrorAction Stop | Out-Null
        }
        'SendOnBehalf' {
            Set-Mailbox -Identity $Mailbox -GrantSendOnBehalfTo @{ Add = $del } -ErrorAction Stop
        }
    }
}

function Remove-RemediationMailboxDelegation {
    param(
        [Parameter(Mandatory = $true)][string]$Mailbox,
        [Parameter(Mandatory = $true)][string]$Delegate,
        [Parameter(Mandatory = $true)][ValidateSet('FullAccess', 'SendAs', 'SendOnBehalf')][string]$Right
    )
    $del = $Delegate.Trim()
    switch ($Right) {
        'FullAccess' {
            Remove-MailboxPermission -Identity $Mailbox -User $del -AccessRights FullAccess -Confirm:$false -ErrorAction Stop
        }
        'SendAs' {
            Remove-RecipientPermission -Identity $Mailbox -Trustee $del -AccessRights SendAs -Confirm:$false -ErrorAction Stop | Out-Null
        }
        'SendOnBehalf' {
            Set-Mailbox -Identity $Mailbox -GrantSendOnBehalfTo @{ Remove = $del } -ErrorAction Stop
        }
    }
}

function Get-RemediationTransportRules {
    if (-not (Get-Command Get-TransportRule -ErrorAction SilentlyContinue)) {
        throw 'Get-TransportRule is not available. Complete Exchange Auth.'
    }
    foreach ($rule in @(Get-TransportRule -ErrorAction Stop)) {
        $desc = [string]$rule.Description
        if ($desc.Length -gt 400) { $desc = $desc.Substring(0, 400) + '…' }
        [pscustomobject]@{
            Identity    = [string]$rule.Identity
            Name        = [string]$rule.Name
            Priority    = $rule.Priority
            State       = [string]$rule.State
            Mode        = [string]$rule.Mode
            Description = $desc
        }
    }
}

function Remove-RemediationTransportRules {
    param([Parameter(Mandatory = $true)][string[]]$Identities)
    foreach ($id in $Identities) {
        if ([string]::IsNullOrWhiteSpace($id)) { continue }
        try {
            Remove-TransportRule -Identity $id -Confirm:$false -ErrorAction Stop
            [pscustomobject]@{ Identity = $id; Success = $true; Error = '' }
        } catch {
            [pscustomobject]@{ Identity = $id; Success = $false; Error = $_.Exception.Message }
        }
    }
}

function Convert-RemediationConnectorList {
    param($Connectors, [string]$Direction)
    foreach ($c in @($Connectors)) {
        $smart = if ($c.SmartHosts) { @($c.SmartHosts) -join '; ' } else { '' }
        $ips = if ($c.SenderIPAddresses) { @($c.SenderIPAddresses) -join '; ' } else { '' }
        $domains = if ($c.SenderDomains) { @($c.SenderDomains) -join '; ' } elseif ($c.RecipientDomains) { @($c.RecipientDomains) -join '; ' } else { '' }
        [pscustomobject]@{
            Direction = $Direction
            Name      = [string]$c.Name
            Identity  = if ($c.Identity) { [string]$c.Identity } else { [string]$c.Name }
            Enabled   = [bool]$c.Enabled
            ConnectorType = [string]$c.ConnectorType
            SmartHosts = $smart
            SenderIPAddresses = $ips
            Domains   = $domains
        }
    }
}

function Get-RemediationConnectors {
    $rows = [System.Collections.ArrayList]::new()
    try {
        foreach ($c in @(Convert-RemediationConnectorList -Connectors (Get-InboundConnector -ErrorAction Stop) -Direction 'Inbound')) {
            [void]$rows.Add($c)
        }
    } catch {}
    try {
        foreach ($c in @(Convert-RemediationConnectorList -Connectors (Get-OutboundConnector -ErrorAction Stop) -Direction 'Outbound')) {
            [void]$rows.Add($c)
        }
    } catch {}
    return @($rows)
}

function Remove-RemediationConnectors {
    param([Parameter(Mandatory = $true)][object[]]$Items)
    foreach ($item in $Items) {
        $name = if ($item.Name) { [string]$item.Name } elseif ($item.Identity) { [string]$item.Identity } else { [string]$item }
        $dir = if ($item.Direction) { [string]$item.Direction } elseif ($item.Type) { [string]$item.Type } else { 'Inbound' }
        if ([string]::IsNullOrWhiteSpace($name)) { continue }
        try {
            if ($dir -ieq 'Outbound') {
                Remove-OutboundConnector -Identity $name -Confirm:$false -ErrorAction Stop
            } else {
                Remove-InboundConnector -Identity $name -Confirm:$false -ErrorAction Stop
            }
            [pscustomobject]@{ Name = $name; Direction = $dir; Success = $true; Error = '' }
        } catch {
            [pscustomobject]@{ Name = $name; Direction = $dir; Success = $false; Error = $_.Exception.Message }
        }
    }
}

function Convert-RemediationConnectorsFromCommand {
    param([string]$Command)
    return @(Convert-RemediationJsonTailFromCommand -Command $Command -Name 'CONNECTORS')
}

function Convert-RemediationJsonTailFromCommand {
    param([string]$Command, [string]$Name)
    if ($Command -notmatch "\|${Name}:(.+)$") { return @() }
    $raw = $Matches[1]
    try {
        return @($raw | ConvertFrom-Json -ErrorAction Stop)
    } catch {
        return @()
    }
}

function Get-RemediationGraphErrorText {
    param($ErrorRecord)
    $msg = if ($ErrorRecord.Exception.Message) { $ErrorRecord.Exception.Message } else { [string]$ErrorRecord }
    $resp = $null
    try {
        if ($ErrorRecord.ErrorDetails.Message) { $resp = $ErrorRecord.ErrorDetails.Message }
        elseif ($ErrorRecord.Exception.Response) {
            $stream = $ErrorRecord.Exception.Response.GetResponseStream()
            if ($stream) {
                $reader = New-Object System.IO.StreamReader($stream)
                $resp = $reader.ReadToEnd()
            }
        }
    } catch {}
    if ($resp) { $msg = "$msg $resp" }
    if ($msg -match 'federated') {
        return "Password reset is not available for federated users via Graph. Use the IdP or SSPR if enabled. Detail: $msg"
    }
    if ($msg -match 'Authorization_RequestDenied|Insufficient privileges|Access is denied|403') {
        return "Graph refused the write (missing User.RevokeSessions.All, User.EnableDisableAccount.All, UserAuthenticationMethod.ReadWrite.All, Device.ReadWrite.All / Directory.AccessAsUser.All, Application.ReadWrite.All, or User-PasswordProfile.ReadWrite.All, or admin consent). For interactive Graph, sign in again. For WCM, use Update Graph App scopes, re-consent, then Graph Auth again. App-only password reset also needs the User Administrator directory role on the app. Detail: $msg"
    }
    return $msg
}

function Invoke-RemediationWorkerCommand {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Command,
        [int]$ClientNumber,
        [string]$CompanyName,
        [string]$TenantName,
        [string]$AuditFolder,
        [bool]$GraphAuthenticated,
        [bool]$ExchangeAuthenticated
    )
    $tenantLabel = if ($TenantName) { $TenantName } elseif ($CompanyName) { $CompanyName } else { "Client$ClientNumber" }
    $caps = $null
    $needsGraph = $Command -match '^REMEDIATE_(REVOKE_SESSIONS|BLOCK|UNBLOCK|SIGNIN_STATUS|RESET_PASSWORD|LIST_AUTH_METHODS|DELETE_AUTH_METHODS|LIST_DEVICES|DELETE_DEVICES|LIST_APPS|DELETE_APPS|LIST_OAUTH_GRANTS|DELETE_OAUTH_GRANTS|REREGISTER_MFA|LIST_INTUNE|WIPE_INTUNE|RETIRE_INTUNE|LIST_ROLES|DELETE_ROLES|LIST_APP_CREDS|DELETE_APP_CREDS)'
    $needsExo = $Command -match '^REMEDIATE_(RESTRICTED_EMAIL_STATUS|UNRESTRICT_EMAIL|LIST_INBOX_RULES|DELETE_INBOX_RULES|MAILBOX_STATUS|SET_FORWARDING|CLEAR_FORWARDING|REMOVE_FORWARDING|ADD_DELEGATION|REMOVE_DELEGATION|LIST_TRANSPORT_RULES|DELETE_TRANSPORT_RULES|LIST_CONNECTORS|DELETE_CONNECTORS|LIST_MOBILE_DEVICES|DELETE_MOBILE_DEVICES|LIST_FOLDER_PERMS|DELETE_FOLDER_PERMS|GET_AUTOREPLY|DISABLE_AUTOREPLY|LIST_ORG_FORWARD|SET_ORG_FORWARD|LIST_JUNK|DELETE_JUNK|LIST_JOURNAL|DELETE_JOURNAL|GET_MAILBOX_HOLD|SET_MAILBOX_HOLD|LIST_ELSEWHERE|DELETE_ELSEWHERE)'
    if ($needsGraph -and -not $GraphAuthenticated) {
        return "REMEDIATE_FAILED:Graph authentication not completed"
    }
    if ($needsExo -and -not $ExchangeAuthenticated) {
        return "REMEDIATE_FAILED:Exchange authentication not completed"
    }
    if ($needsGraph) {
        if (-not (Get-Command Test-GraphRestBearerToken -ErrorAction SilentlyContinue) -or -not (Test-GraphRestBearerToken)) {
            return "REMEDIATE_FAILED:Graph REST bearer token is not set. Complete Graph Auth (worker stores the token; Connect-MgGraph is not required)."
        }
        $caps = Get-GraphContainmentCapabilities
    }

    try {
        if ($Command -match '^REMEDIATE_SIGNIN_STATUS') {
            $users = Convert-RemediationUsersFromCommand -Command $Command
            $rows = foreach ($upn in $users) {
                try {
                    $u = Get-RemediationGraphUser -UserPrincipalName $upn
                    [pscustomobject]@{
                        UserPrincipalName = [string]$u.userPrincipalName
                        AccountEnabled    = [bool]$u.accountEnabled
                        DisplayName       = [string]$u.displayName
                        Error             = ''
                    }
                } catch {
                    [pscustomobject]@{
                        UserPrincipalName = $upn
                        AccountEnabled    = $null
                        DisplayName       = ''
                        Error             = (Get-RemediationGraphErrorText $_)
                    }
                }
            }
            $payload = @{
                Capabilities = $caps
                Users        = @($rows)
            } | ConvertTo-Json -Compress -Depth 6
            return "REMEDIATE_SUCCESS:$payload"
        }

        if ($Command -match '^REMEDIATE_REVOKE_SESSIONS') {
            if ($caps -and $caps.canRevoke -eq $false) {
                return "REMEDIATE_FAILED:$($caps.reason)"
            }
            $users = Convert-RemediationUsersFromCommand -Command $Command
            if ($users.Count -eq 0) { return 'REMEDIATE_FAILED:No users provided' }
            $ok = 0; $fail = 0; $details = [System.Collections.ArrayList]::new()
            foreach ($upn in $users) {
                try {
                    $r = Invoke-RemediationRevokeSessions -UserPrincipalName $upn
                    $ok++
                    [void]$details.Add("$upn : revoked")
                    Write-RemediationAuditCsv -Folder $AuditFolder -Tenant $tenantLabel -ClientNumber $ClientNumber -Upn $upn -Action 'revoke' -Result 'success' -Detail 'revokeSignInSessions'
                } catch {
                    $fail++
                    $err = Get-RemediationGraphErrorText $_
                    [void]$details.Add("$upn : $err")
                    Write-RemediationAuditCsv -Folder $AuditFolder -Tenant $tenantLabel -ClientNumber $ClientNumber -Upn $upn -Action 'revoke' -Result 'failed' -Detail $err
                }
            }
            $payload = @{ SuccessCount = $ok; FailCount = $fail; Details = @($details); Capabilities = $caps } | ConvertTo-Json -Compress -Depth 6
            if ($fail -gt 0 -and $ok -eq 0) { return "REMEDIATE_FAILED:$payload" }
            return "REMEDIATE_SUCCESS:$payload"
        }

        if ($Command -match '^REMEDIATE_BLOCK' -or $Command -match '^REMEDIATE_UNBLOCK') {
            $enable = $Command -match '^REMEDIATE_UNBLOCK'
            $action = if ($enable) { 'unblock' } else { 'block' }
            if ($caps -and $caps.canBlock -eq $false) {
                return "REMEDIATE_FAILED:$($caps.reason)"
            }
            $users = Convert-RemediationUsersFromCommand -Command $Command
            if ($users.Count -eq 0) { return 'REMEDIATE_FAILED:No users provided' }
            $ok = 0; $fail = 0; $details = [System.Collections.ArrayList]::new()
            foreach ($upn in $users) {
                try {
                    $r = Invoke-RemediationSetAccountEnabled -UserPrincipalName $upn -Enabled $enable
                    $ok++
                    [void]$details.Add("$($r.UserPrincipalName) : accountEnabled=$($r.AccountEnabled)")
                    Write-RemediationAuditCsv -Folder $AuditFolder -Tenant $tenantLabel -ClientNumber $ClientNumber -Upn $upn -Action $action -Result 'success' -Detail "accountEnabled=$($r.AccountEnabled)"
                } catch {
                    $fail++
                    $err = Get-RemediationGraphErrorText $_
                    [void]$details.Add("$upn : $err")
                    Write-RemediationAuditCsv -Folder $AuditFolder -Tenant $tenantLabel -ClientNumber $ClientNumber -Upn $upn -Action $action -Result 'failed' -Detail $err
                }
            }
            $payload = @{ SuccessCount = $ok; FailCount = $fail; Details = @($details); Capabilities = $caps } | ConvertTo-Json -Compress -Depth 6
            if ($fail -gt 0 -and $ok -eq 0) { return "REMEDIATE_FAILED:$payload" }
            return "REMEDIATE_SUCCESS:$payload"
        }

        if ($Command -match '^REMEDIATE_RESET_PASSWORD') {
            if ($caps -and $caps.canPasswordReset -eq $false) {
                return "REMEDIATE_FAILED:$($caps.reason)"
            }
            $users = Convert-RemediationUsersFromCommand -Command $Command
            if ($users.Count -eq 0) { return 'REMEDIATE_FAILED:No users provided' }
            $opts = @(Convert-RemediationJsonTailFromCommand -Command $Command -Name 'OPTIONS')
            $opt = if ($opts.Count -gt 0) { $opts[0] } else { $null }
            $mode = if ($opt -and $opt.Mode) { [string]$opt.Mode } else { 'random' }
            $assigned = if ($opt -and $opt.Password) { [string]$opt.Password } else { '' }
            if ($mode -ieq 'assign' -and [string]::IsNullOrWhiteSpace($assigned)) {
                return 'REMEDIATE_FAILED:Assigned password is empty'
            }
            $ok = 0; $fail = 0; $details = [System.Collections.ArrayList]::new(); $sspr = ''
            foreach ($upn in $users) {
                try {
                    $pw = if ($mode -ieq 'assign') { $assigned } else { $null }
                    $r = Invoke-RemediationResetPassword -UserPrincipalName $upn -Password $pw
                    $ok++
                    if (-not $sspr) { $sspr = $r.SsprUrl }
                    [void]$details.Add("$($r.UserPrincipalName) : password reset ($mode). Do not send a password - send SSPR: $($r.SsprUrl)")
                    Write-RemediationAuditCsv -Folder $AuditFolder -Tenant $tenantLabel -ClientNumber $ClientNumber -Upn $upn -Action 'reset-password' -Result 'success' -Detail $mode
                } catch {
                    $fail++
                    $err = Get-RemediationGraphErrorText $_
                    [void]$details.Add("$upn : $err")
                    Write-RemediationAuditCsv -Folder $AuditFolder -Tenant $tenantLabel -ClientNumber $ClientNumber -Upn $upn -Action 'reset-password' -Result 'failed' -Detail $err
                }
            }
            if (-not $sspr -and $users.Count -gt 0) { $sspr = Get-RemediationSsprUrl -UserPrincipalName $users[0] }
            $payload = @{
                SuccessCount = $ok
                FailCount    = $fail
                Details      = @($details)
                Mode         = $mode
                SsprUrl      = $sspr
                SsprHint     = 'https://aka.ms/sspr'
                Capabilities = $caps
            } | ConvertTo-Json -Compress -Depth 6
            if ($fail -gt 0 -and $ok -eq 0) { return "REMEDIATE_FAILED:$payload" }
            return "REMEDIATE_SUCCESS:$payload"
        }

        if ($Command -match '^REMEDIATE_LIST_AUTH_METHODS') {
            $users = Convert-RemediationUsersFromCommand -Command $Command
            if ($users.Count -eq 0) { return 'REMEDIATE_FAILED:No users provided' }
            $rows = @(Get-RemediationAuthMethods -UserPrincipalNames $users)
            Write-RemediationAuditCsv -Folder $AuditFolder -Tenant $tenantLabel -ClientNumber $ClientNumber -Upn ($users -join ';') -Action 'list-auth-methods' -Result 'success' -Detail "$($rows.Count) method(s)"
            $payload = @{ Methods = @($rows); Capabilities = $caps } | ConvertTo-Json -Compress -Depth 6
            return "REMEDIATE_AUTHMETHODS:$payload"
        }

        if ($Command -match '^REMEDIATE_DELETE_AUTH_METHODS') {
            $items = @(Convert-RemediationJsonTailFromCommand -Command $Command -Name 'ITEMS')
            if ($items.Count -eq 0) { return 'REMEDIATE_FAILED:No authentication methods provided' }
            $deleted = @(Remove-RemediationAuthMethods -Items $items)
            foreach ($row in $deleted) {
                $res = if ($row.Success) { 'success' } else { 'failed' }
                Write-RemediationAuditCsv -Folder $AuditFolder -Tenant $tenantLabel -ClientNumber $ClientNumber -Upn $row.UserPrincipalName -Action 'delete-auth-method' -Result $res -Detail ("$($row.Id) $($row.Error)")
            }
            $upns = @($items | ForEach-Object { if ($_.UserPrincipalName) { $_.UserPrincipalName } elseif ($_.User) { $_.User } } | Select-Object -Unique)
            $remaining = @()
            if ($upns.Count -gt 0) { $remaining = @(Get-RemediationAuthMethods -UserPrincipalNames $upns) }
            $ok = @($deleted | Where-Object { $_.Success }).Count
            $fail = @($deleted | Where-Object { -not $_.Success }).Count
            $payload = @{ Deleted = @($deleted); Methods = @($remaining); SuccessCount = $ok; FailCount = $fail; Capabilities = $caps } | ConvertTo-Json -Compress -Depth 6
            if ($fail -gt 0 -and $ok -eq 0) { return "REMEDIATE_FAILED:$payload" }
            return "REMEDIATE_AUTHMETHODS:$payload"
        }

        if ($Command -match '^REMEDIATE_LIST_DEVICES') {
            $users = Convert-RemediationUsersFromCommand -Command $Command
            if ($users.Count -eq 0) { return 'REMEDIATE_FAILED:No users provided' }
            $rows = @(Get-RemediationUserDevices -UserPrincipalNames $users)
            Write-RemediationAuditCsv -Folder $AuditFolder -Tenant $tenantLabel -ClientNumber $ClientNumber -Upn ($users -join ';') -Action 'list-devices' -Result 'success' -Detail "$($rows.Count) device(s)"
            $payload = @{ Devices = @($rows); Capabilities = $caps } | ConvertTo-Json -Compress -Depth 6
            return "REMEDIATE_DEVICES:$payload"
        }

        if ($Command -match '^REMEDIATE_DELETE_DEVICES') {
            $items = @(Convert-RemediationJsonTailFromCommand -Command $Command -Name 'DEVICES')
            if ($items.Count -eq 0) { return 'REMEDIATE_FAILED:No devices provided' }
            $deleted = @(Remove-RemediationDevices -Items $items)
            foreach ($row in $deleted) {
                $res = if ($row.Success) { 'success' } else { 'failed' }
                Write-RemediationAuditCsv -Folder $AuditFolder -Tenant $tenantLabel -ClientNumber $ClientNumber -Upn '' -Action 'delete-device' -Result $res -Detail ("$($row.Id) $($row.Error)")
            }
            $upns = @($items | ForEach-Object { if ($_.UserPrincipalName) { $_.UserPrincipalName } elseif ($_.User) { $_.User } } | Where-Object { $_ } | Select-Object -Unique)
            $remaining = @()
            if ($upns.Count -gt 0) { $remaining = @(Get-RemediationUserDevices -UserPrincipalNames $upns) }
            $ok = @($deleted | Where-Object { $_.Success }).Count
            $fail = @($deleted | Where-Object { -not $_.Success }).Count
            $payload = @{ Deleted = @($deleted); Devices = @($remaining); SuccessCount = $ok; FailCount = $fail; Capabilities = $caps } | ConvertTo-Json -Compress -Depth 6
            if ($fail -gt 0 -and $ok -eq 0) { return "REMEDIATE_FAILED:$payload" }
            return "REMEDIATE_DEVICES:$payload"
        }

        if ($Command -match '^REMEDIATE_LIST_APPS') {
            $rows = @(Get-RemediationDirectoryApps)
            Write-RemediationAuditCsv -Folder $AuditFolder -Tenant $tenantLabel -ClientNumber $ClientNumber -Upn '' -Action 'list-apps' -Result 'success' -Detail "$($rows.Count) app(s)"
            $payload = @{ Apps = @($rows); Capabilities = $caps } | ConvertTo-Json -Compress -Depth 6
            return "REMEDIATE_APPS:$payload"
        }

        if ($Command -match '^REMEDIATE_DELETE_APPS') {
            $items = @(Convert-RemediationJsonTailFromCommand -Command $Command -Name 'APPS')
            if ($items.Count -eq 0) { return 'REMEDIATE_FAILED:No apps provided' }
            $deleted = @(Remove-RemediationDirectoryApps -Items $items)
            foreach ($row in $deleted) {
                $res = if ($row.Success) { 'success' } else { 'failed' }
                Write-RemediationAuditCsv -Folder $AuditFolder -Tenant $tenantLabel -ClientNumber $ClientNumber -Upn '' -Action 'delete-app' -Result $res -Detail ("$($row.Kind) $($row.Id) $($row.Error)")
            }
            $remaining = @(Get-RemediationDirectoryApps)
            $ok = @($deleted | Where-Object { $_.Success }).Count
            $fail = @($deleted | Where-Object { -not $_.Success }).Count
            $payload = @{ Deleted = @($deleted); Apps = @($remaining); SuccessCount = $ok; FailCount = $fail; Capabilities = $caps } | ConvertTo-Json -Compress -Depth 6
            if ($fail -gt 0 -and $ok -eq 0) { return "REMEDIATE_FAILED:$payload" }
            return "REMEDIATE_APPS:$payload"
        }

        if ($Command -match '^REMEDIATE_RESTRICTED_EMAIL_STATUS') {
            $users = Convert-RemediationUsersFromCommand -Command $Command
            if ($users.Count -eq 0) { return 'REMEDIATE_FAILED:No users provided' }
            $rows = Get-RemediationRestrictedEmailStatus -UserPrincipalNames $users
            foreach ($row in $rows) {
                Write-RemediationAuditCsv -Folder $AuditFolder -Tenant $tenantLabel -ClientNumber $ClientNumber -Upn $row.UserPrincipalName -Action 'restricted-status' -Result 'success' -Detail ("restricted=$($row.Restricted); reason=$($row.Reason); date=$($row.CreatedDateTime)")
            }
            $payload = @{ Users = @($rows) } | ConvertTo-Json -Compress -Depth 6
            return "REMEDIATE_RESTRICTED:$payload"
        }

        if ($Command -match '^REMEDIATE_UNRESTRICT_EMAIL') {
            $users = Convert-RemediationUsersFromCommand -Command $Command
            if ($users.Count -eq 0) { return 'REMEDIATE_FAILED:No users provided' }
            $status = Get-RemediationRestrictedEmailStatus -UserPrincipalNames $users
            $ok = 0; $fail = 0; $details = [System.Collections.ArrayList]::new()
            foreach ($row in $status) {
                if (-not $row.Restricted) {
                    [void]$details.Add("$($row.UserPrincipalName) : not on Restricted Users list")
                    Write-RemediationAuditCsv -Folder $AuditFolder -Tenant $tenantLabel -ClientNumber $ClientNumber -Upn $row.UserPrincipalName -Action 'unrestrict' -Result 'skipped' -Detail 'not restricted'
                    continue
                }
                $addr = if ($row.SenderAddress) { $row.SenderAddress } else { $row.UserPrincipalName }
                try {
                    Invoke-RemediationUnrestrictEmail -SenderAddress $addr | Out-Null
                    $ok++
                    [void]$details.Add("$addr : unrestricted")
                    Write-RemediationAuditCsv -Folder $AuditFolder -Tenant $tenantLabel -ClientNumber $ClientNumber -Upn $row.UserPrincipalName -Action 'unrestrict' -Result 'success' -Detail $addr
                } catch {
                    $fail++
                    $err = $_.Exception.Message
                    [void]$details.Add("$addr : $err")
                    Write-RemediationAuditCsv -Folder $AuditFolder -Tenant $tenantLabel -ClientNumber $ClientNumber -Upn $row.UserPrincipalName -Action 'unrestrict' -Result 'failed' -Detail $err
                }
            }
            $payload = @{ SuccessCount = $ok; FailCount = $fail; Details = @($details); Status = @($status) } | ConvertTo-Json -Compress -Depth 6
            if ($fail -gt 0 -and $ok -eq 0) { return "REMEDIATE_FAILED:$payload" }
            return "REMEDIATE_SUCCESS:$payload"
        }

        if ($Command -match '^REMEDIATE_LIST_INBOX_RULES') {
            $mailbox = $null
            if ($Command -match '\|USER:([^|]+)') { $mailbox = $Matches[1].Trim() }
            if ([string]::IsNullOrWhiteSpace($mailbox)) { return 'REMEDIATE_FAILED:USER (mailbox UPN) is required' }
            $rules = @(Get-RemediationInboxRules -Mailbox $mailbox)
            Write-RemediationAuditCsv -Folder $AuditFolder -Tenant $tenantLabel -ClientNumber $ClientNumber -Upn $mailbox -Action 'list-rules' -Result 'success' -Detail "$($rules.Count) rule(s)"
            $payload = @{ User = $mailbox; Rules = @($rules) } | ConvertTo-Json -Compress -Depth 6
            return "REMEDIATE_RULES:$payload"
        }

        if ($Command -match '^REMEDIATE_DELETE_INBOX_RULES') {
            $mailbox = $null
            if ($Command -match '\|USER:([^|]+)') { $mailbox = $Matches[1].Trim() }
            if ([string]::IsNullOrWhiteSpace($mailbox)) { return 'REMEDIATE_FAILED:USER (mailbox UPN) is required' }
            $ids = Convert-RemediationRuleIdentitiesFromCommand -Command $Command
            if ($ids.Count -eq 0) { return 'REMEDIATE_FAILED:No rule identities provided' }
            $rows = Remove-RemediationInboxRules -Mailbox $mailbox -Identities $ids
            foreach ($row in $rows) {
                $res = if ($row.Success) { 'success' } else { 'failed' }
                Write-RemediationAuditCsv -Folder $AuditFolder -Tenant $tenantLabel -ClientNumber $ClientNumber -Upn $mailbox -Action 'delete-rule' -Result $res -Detail ("$($row.Identity) $($row.Error)")
            }
            $remaining = @()
            try { $remaining = @(Get-RemediationInboxRules -Mailbox $mailbox) } catch {}
            $ok = @($rows | Where-Object { $_.Success }).Count
            $fail = @($rows | Where-Object { -not $_.Success }).Count
            $payload = @{ User = $mailbox; Deleted = @($rows); Rules = @($remaining); SuccessCount = $ok; FailCount = $fail } | ConvertTo-Json -Compress -Depth 6
            if ($fail -gt 0 -and $ok -eq 0) { return "REMEDIATE_FAILED:$payload" }
            return "REMEDIATE_SUCCESS:$payload"
        }

        if ($Command -match '^REMEDIATE_MAILBOX_STATUS') {
            $users = Convert-RemediationUsersFromCommand -Command $Command
            if ($users.Count -eq 0) { return 'REMEDIATE_FAILED:No users provided' }
            $rows = foreach ($upn in $users) {
                try {
                    Get-RemediationMailboxAccess -UserPrincipalNames @($upn)
                } catch {
                    [pscustomobject]@{
                        UserPrincipalName = $upn
                        ForwardingAddress = ''
                        ForwardingSmtpAddress = ''
                        DeliverToMailboxAndForward = $false
                        Delegates = @()
                        Error = $_.Exception.Message
                    }
                }
            }
            Write-RemediationAuditCsv -Folder $AuditFolder -Tenant $tenantLabel -ClientNumber $ClientNumber -Upn ($users -join ';') -Action 'mailbox-status' -Result 'success' -Detail "$($rows.Count) mailbox(es)"
            $payload = @{ Users = @($rows) } | ConvertTo-Json -Compress -Depth 8
            return "REMEDIATE_MAILBOX:$payload"
        }

        if ($Command -match '^REMEDIATE_SET_FORWARDING') {
            $mailbox = Get-RemediationCommandToken -Command $Command -Name 'USER'
            $smtp = Get-RemediationCommandToken -Command $Command -Name 'SMTP'
            $deliverTok = Get-RemediationCommandToken -Command $Command -Name 'DELIVER'
            $deliver = $deliverTok -ne '0'
            if ([string]::IsNullOrWhiteSpace($mailbox) -or [string]::IsNullOrWhiteSpace($smtp)) {
                return 'REMEDIATE_FAILED:USER and SMTP are required'
            }
            Set-RemediationMailboxForwarding -Mailbox $mailbox -SmtpAddress $smtp -DeliverToMailboxAndForward $deliver
            Write-RemediationAuditCsv -Folder $AuditFolder -Tenant $tenantLabel -ClientNumber $ClientNumber -Upn $mailbox -Action 'set-forwarding' -Result 'success' -Detail "$smtp deliver=$deliver"
            $rows = @(Get-RemediationMailboxAccess -UserPrincipalNames @($mailbox))
            $payload = @{ Users = @($rows) } | ConvertTo-Json -Compress -Depth 8
            return "REMEDIATE_MAILBOX:$payload"
        }

        if ($Command -match '^REMEDIATE_CLEAR_FORWARDING') {
            $users = Convert-RemediationUsersFromCommand -Command $Command
            if ($users.Count -eq 0) { return 'REMEDIATE_FAILED:No users provided' }
            $ok = 0; $fail = 0; $details = [System.Collections.ArrayList]::new()
            foreach ($upn in $users) {
                try {
                    Set-RemediationMailboxForwarding -Mailbox $upn -SmtpAddress ''
                    $ok++
                    [void]$details.Add("$upn : forwarding cleared")
                    Write-RemediationAuditCsv -Folder $AuditFolder -Tenant $tenantLabel -ClientNumber $ClientNumber -Upn $upn -Action 'clear-forwarding' -Result 'success' -Detail ''
                } catch {
                    $fail++
                    [void]$details.Add("$upn : $($_.Exception.Message)")
                    Write-RemediationAuditCsv -Folder $AuditFolder -Tenant $tenantLabel -ClientNumber $ClientNumber -Upn $upn -Action 'clear-forwarding' -Result 'failed' -Detail $_.Exception.Message
                }
            }
            $rows = foreach ($upn in $users) {
                try { Get-RemediationMailboxAccess -UserPrincipalNames @($upn) } catch { $null }
            }
            $payload = @{ SuccessCount = $ok; FailCount = $fail; Details = @($details); Users = @($rows | Where-Object { $_ }) } | ConvertTo-Json -Compress -Depth 8
            if ($fail -gt 0 -and $ok -eq 0) { return "REMEDIATE_FAILED:$payload" }
            return "REMEDIATE_MAILBOX:$payload"
        }

        if ($Command -match '^REMEDIATE_REMOVE_FORWARDING') {
            $items = @(Convert-RemediationJsonTailFromCommand -Command $Command -Name 'ITEMS')
            if ($items.Count -eq 0) { return 'REMEDIATE_FAILED:No forwarding entries provided' }
            $ok = 0; $fail = 0; $details = [System.Collections.ArrayList]::new(); $mailboxes = [System.Collections.Generic.List[string]]::new()
            foreach ($item in $items) {
                $mailbox = if ($item.UserPrincipalName) { [string]$item.UserPrincipalName } elseif ($item.User) { [string]$item.User } elseif ($item.Mailbox) { [string]$item.Mailbox } else { '' }
                $field = if ($item.Field) { [string]$item.Field } else { '' }
                if ([string]::IsNullOrWhiteSpace($mailbox) -or [string]::IsNullOrWhiteSpace($field)) { continue }
                if ($mailboxes -notcontains $mailbox) { [void]$mailboxes.Add($mailbox) }
                try {
                    Clear-RemediationMailboxForwardingField -Mailbox $mailbox -Field $field
                    $ok++
                    [void]$details.Add("$mailbox : cleared $field forwarding")
                    Write-RemediationAuditCsv -Folder $AuditFolder -Tenant $tenantLabel -ClientNumber $ClientNumber -Upn $mailbox -Action 'remove-forwarding' -Result 'success' -Detail $field
                } catch {
                    $fail++
                    [void]$details.Add("$mailbox $field : $($_.Exception.Message)")
                    Write-RemediationAuditCsv -Folder $AuditFolder -Tenant $tenantLabel -ClientNumber $ClientNumber -Upn $mailbox -Action 'remove-forwarding' -Result 'failed' -Detail $_.Exception.Message
                }
            }
            $rows = foreach ($upn in $mailboxes) {
                try { Get-RemediationMailboxAccess -UserPrincipalNames @($upn) } catch { $null }
            }
            $payload = @{ SuccessCount = $ok; FailCount = $fail; Details = @($details); Users = @($rows | Where-Object { $_ }) } | ConvertTo-Json -Compress -Depth 8
            if ($fail -gt 0 -and $ok -eq 0) { return "REMEDIATE_FAILED:$payload" }
            return "REMEDIATE_MAILBOX:$payload"
        }

        if ($Command -match '^REMEDIATE_ADD_DELEGATION') {
            $mailbox = Get-RemediationCommandToken -Command $Command -Name 'USER'
            $delegate = Get-RemediationCommandToken -Command $Command -Name 'DELEGATE'
            $right = Get-RemediationCommandToken -Command $Command -Name 'RIGHT'
            if ([string]::IsNullOrWhiteSpace($mailbox) -or [string]::IsNullOrWhiteSpace($delegate) -or [string]::IsNullOrWhiteSpace($right)) {
                return 'REMEDIATE_FAILED:USER, DELEGATE, and RIGHT are required'
            }
            Add-RemediationMailboxDelegation -Mailbox $mailbox -Delegate $delegate -Right $right
            Write-RemediationAuditCsv -Folder $AuditFolder -Tenant $tenantLabel -ClientNumber $ClientNumber -Upn $mailbox -Action 'add-delegation' -Result 'success' -Detail "$right $delegate"
            $rows = @(Get-RemediationMailboxAccess -UserPrincipalNames @($mailbox))
            $payload = @{ Users = @($rows) } | ConvertTo-Json -Compress -Depth 8
            return "REMEDIATE_MAILBOX:$payload"
        }

        if ($Command -match '^REMEDIATE_REMOVE_DELEGATION') {
            $mailbox = Get-RemediationCommandToken -Command $Command -Name 'USER'
            $delegate = Get-RemediationCommandToken -Command $Command -Name 'DELEGATE'
            $right = Get-RemediationCommandToken -Command $Command -Name 'RIGHT'
            if ([string]::IsNullOrWhiteSpace($mailbox) -or [string]::IsNullOrWhiteSpace($delegate) -or [string]::IsNullOrWhiteSpace($right)) {
                return 'REMEDIATE_FAILED:USER, DELEGATE, and RIGHT are required'
            }
            Remove-RemediationMailboxDelegation -Mailbox $mailbox -Delegate $delegate -Right $right
            Write-RemediationAuditCsv -Folder $AuditFolder -Tenant $tenantLabel -ClientNumber $ClientNumber -Upn $mailbox -Action 'remove-delegation' -Result 'success' -Detail "$right $delegate"
            $rows = @(Get-RemediationMailboxAccess -UserPrincipalNames @($mailbox))
            $payload = @{ Users = @($rows) } | ConvertTo-Json -Compress -Depth 8
            return "REMEDIATE_MAILBOX:$payload"
        }

        if ($Command -match '^REMEDIATE_LIST_TRANSPORT_RULES') {
            $rules = @(Get-RemediationTransportRules)
            Write-RemediationAuditCsv -Folder $AuditFolder -Tenant $tenantLabel -ClientNumber $ClientNumber -Upn '' -Action 'list-transport-rules' -Result 'success' -Detail "$($rules.Count) rule(s)"
            $payload = @{ Rules = @($rules) } | ConvertTo-Json -Compress -Depth 6
            return "REMEDIATE_TRANSPORT:$payload"
        }

        if ($Command -match '^REMEDIATE_DELETE_TRANSPORT_RULES') {
            $ids = Convert-RemediationRuleIdentitiesFromCommand -Command $Command
            if ($ids.Count -eq 0) { return 'REMEDIATE_FAILED:No transport rule identities provided' }
            $rows = @(Remove-RemediationTransportRules -Identities $ids)
            foreach ($row in $rows) {
                $res = if ($row.Success) { 'success' } else { 'failed' }
                Write-RemediationAuditCsv -Folder $AuditFolder -Tenant $tenantLabel -ClientNumber $ClientNumber -Upn '' -Action 'delete-transport-rule' -Result $res -Detail ("$($row.Identity) $($row.Error)")
            }
            $remaining = @(Get-RemediationTransportRules)
            $ok = @($rows | Where-Object { $_.Success }).Count
            $fail = @($rows | Where-Object { -not $_.Success }).Count
            $payload = @{ Deleted = @($rows); Rules = @($remaining); SuccessCount = $ok; FailCount = $fail } | ConvertTo-Json -Compress -Depth 6
            if ($fail -gt 0 -and $ok -eq 0) { return "REMEDIATE_FAILED:$payload" }
            return "REMEDIATE_TRANSPORT:$payload"
        }

        if ($Command -match '^REMEDIATE_LIST_CONNECTORS') {
            $rows = @(Get-RemediationConnectors)
            Write-RemediationAuditCsv -Folder $AuditFolder -Tenant $tenantLabel -ClientNumber $ClientNumber -Upn '' -Action 'list-connectors' -Result 'success' -Detail "$($rows.Count) connector(s)"
            $payload = @{ Connectors = @($rows) } | ConvertTo-Json -Compress -Depth 6
            return "REMEDIATE_CONNECTORS:$payload"
        }

        if ($Command -match '^REMEDIATE_DELETE_CONNECTORS') {
            $items = @(Convert-RemediationConnectorsFromCommand -Command $Command)
            if ($items.Count -eq 0) { return 'REMEDIATE_FAILED:No connectors provided' }
            $rows = @(Remove-RemediationConnectors -Items $items)
            foreach ($row in $rows) {
                $res = if ($row.Success) { 'success' } else { 'failed' }
                Write-RemediationAuditCsv -Folder $AuditFolder -Tenant $tenantLabel -ClientNumber $ClientNumber -Upn '' -Action 'delete-connector' -Result $res -Detail ("$($row.Direction) $($row.Name) $($row.Error)")
            }
            $remaining = @(Get-RemediationConnectors)
            $ok = @($rows | Where-Object { $_.Success }).Count
            $fail = @($rows | Where-Object { -not $_.Success }).Count
            $payload = @{ Deleted = @($rows); Connectors = @($remaining); SuccessCount = $ok; FailCount = $fail } | ConvertTo-Json -Compress -Depth 6
            if ($fail -gt 0 -and $ok -eq 0) { return "REMEDIATE_FAILED:$payload" }
            return "REMEDIATE_CONNECTORS:$payload"
        }

        $extended = Invoke-RemediationExtendedCommand -Command $Command -Capabilities $caps -AuditFolder $AuditFolder -TenantLabel $tenantLabel -ClientNumber $ClientNumber
        if ($null -ne $extended) { return $extended }

        return "REMEDIATE_FAILED:Unknown remediation command"
    } catch {
        $err = if ($needsGraph) { Get-RemediationGraphErrorText $_ } else { $_.Exception.Message }
        return "REMEDIATE_FAILED:$err"
    }
}

. (Join-Path $PSScriptRoot 'RemediationExtended.ps1')

Export-ModuleMember -Function Get-GraphContainmentCapabilities, Get-RemediationGraphUser
Export-ModuleMember -Function Invoke-RemediationRevokeSessions, Invoke-RemediationSetAccountEnabled, Invoke-RemediationResetPassword
Export-ModuleMember -Function Get-RemediationAuthMethods, Remove-RemediationAuthMethods
Export-ModuleMember -Function Get-RemediationUserDevices, Remove-RemediationDevices
Export-ModuleMember -Function Get-RemediationDirectoryApps, Remove-RemediationDirectoryApps
Export-ModuleMember -Function Get-RemediationRestrictedEmailStatus, Invoke-RemediationUnrestrictEmail
Export-ModuleMember -Function Get-RemediationInboxRules, Remove-RemediationInboxRules
Export-ModuleMember -Function Get-RemediationMailboxAccess, Set-RemediationMailboxForwarding
Export-ModuleMember -Function Add-RemediationMailboxDelegation, Remove-RemediationMailboxDelegation
Export-ModuleMember -Function Get-RemediationTransportRules, Remove-RemediationTransportRules
Export-ModuleMember -Function Get-RemediationConnectors, Remove-RemediationConnectors
Export-ModuleMember -Function Write-RemediationAuditCsv, Invoke-RemediationWorkerCommand
