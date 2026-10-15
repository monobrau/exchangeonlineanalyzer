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
    $key = $UserPrincipalName.Trim().ToLowerInvariant()
    if (-not $script:RemediationUserCache) { $script:RemediationUserCache = @{} }
    $hit = $script:RemediationUserCache[$key]
    if ($hit -and $hit.At -gt (Get-Date).AddMinutes(-2)) { return $hit.User }
    $enc = [Uri]::EscapeDataString($UserPrincipalName.Trim())
    $user = Invoke-GraphRestRequest -Uri "https://graph.microsoft.com/v1.0/users/$enc`?`$select=id,userPrincipalName,displayName,accountEnabled"
    $script:RemediationUserCache[$key] = @{ At = Get-Date; User = $user }
    return $user
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
    # Canonical commercial SSPR page (aka.ms/sspr only redirects here). ?whr= lands the user in the right tenant.
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
                Created           = ''
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
                Created           = ''
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
                Created           = (Get-RemediationObjectDateText -Object $m -PropertyNames @('createdDateTime', 'CreatedDateTime', 'Created'))
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
            if (Test-RemediationGraphAlreadyGone $_) {
                [pscustomobject]@{ UserPrincipalName = $upn; Id = $methodId; Success = $true; Error = '' }
            } else {
                [pscustomobject]@{ UserPrincipalName = $upn; Id = $methodId; Success = $false; Error = (Get-RemediationGraphErrorText $_) }
            }
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
                    LastSignIn        = (Get-RemediationObjectDateText -Object $d -PropertyNames @('approximateLastSignInDateTime', 'ApproximateLastSignInDateTime', 'lastSignInDateTime'))
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
            if (Test-RemediationGraphAlreadyGone $_) {
                [pscustomobject]@{ Id = $id; Success = $true; Error = '' }
            } else {
                [pscustomobject]@{ Id = $id; Success = $false; Error = (Get-RemediationGraphErrorText $_) }
            }
        }
    }
}

function Get-RemediationDirectoryApps {
    param([int]$Top = 0)
    $msft = 'f8cdef31-a31e-4b4a-93e4-5f571e91255a'
    $tenantId = ''
    try {
        $org = Invoke-GraphRestRequest -Uri 'https://graph.microsoft.com/v1.0/organization?$select=id'
        $tenantId = [string]@($org.value)[0].id
    } catch {}
    $rows = [System.Collections.ArrayList]::new()
    try {
        $appSelect = 'id,appId,displayName,createdDateTime,signInAudience,publisherDomain'
        if ($Top -gt 0) {
            $ordered = "https://graph.microsoft.com/v1.0/applications?`$select=$appSelect&`$top=$Top&`$orderby=createdDateTime desc"
            try {
                $resp = Invoke-GraphRestRequest -Uri $ordered
                $apps = @($resp.value | Where-Object { $_ })
            } catch {
                $apps = @(Invoke-GraphRestPaged -Uri "https://graph.microsoft.com/v1.0/applications?`$select=$appSelect" -Maximum $Top)
            }
        } else {
            $apps = Invoke-GraphRestPaged -Uri "https://graph.microsoft.com/v1.0/applications?`$select=$appSelect"
        }
        foreach ($app in $apps) {
            [void]$rows.Add([pscustomobject]@{
                Kind        = 'AppRegistration'
                Id          = [string]$app.id
                AppId       = [string]$app.appId
                DisplayName = [string]$app.displayName
                Created     = (Get-RemediationObjectDateText -Object $app -PropertyNames @('createdDateTime', 'CreatedDateTime'))
                Publisher   = [string]$app.publisherDomain
                Audience    = [string]$app.signInAudience
                OwnerTenant = $tenantId
            })
        }
    } catch {
        throw
    }
    try {
        $spSelect = 'id,appId,displayName,createdDateTime,appOwnerOrganizationId,servicePrincipalType,accountEnabled,publisherName'
        $spFilter = "servicePrincipalType eq 'Application'"
        if ($Top -gt 0) {
            $spFilter = "servicePrincipalType eq 'Application' and appOwnerOrganizationId ne '$msft'"
            if ($tenantId) { $spFilter = "$spFilter and appOwnerOrganizationId ne '$tenantId'" }
            $ordered = "https://graph.microsoft.com/v1.0/servicePrincipals?`$select=$spSelect&`$filter=$spFilter&`$top=$Top&`$orderby=createdDateTime desc"
            try {
                $resp = Invoke-GraphRestRequest -Uri $ordered
                $sps = @($resp.value | Where-Object { $_ })
            } catch {
                try {
                    $resp = Invoke-GraphRestRequest -Uri "https://graph.microsoft.com/v1.0/servicePrincipals?`$select=$spSelect&`$filter=$spFilter&`$top=$Top"
                    $sps = @($resp.value | Where-Object { $_ })
                } catch {
                    $broad = "servicePrincipalType eq 'Application'"
                    $sps = @(Invoke-GraphRestPaged -Uri "https://graph.microsoft.com/v1.0/servicePrincipals?`$select=$spSelect&`$filter=$broad" -Maximum 400)
                }
            }
        } else {
            $sps = Invoke-GraphRestPaged -Uri "https://graph.microsoft.com/v1.0/servicePrincipals?`$select=$spSelect&`$filter=$spFilter"
        }
        foreach ($sp in $sps) {
            $owner = [string]$sp.appOwnerOrganizationId
            if ($owner -eq $msft) { continue }
            if ($tenantId -and $owner -eq $tenantId) { continue }
            [void]$rows.Add([pscustomobject]@{
                Kind        = 'EnterpriseApp'
                Id          = [string]$sp.id
                AppId       = [string]$sp.appId
                DisplayName = [string]$sp.displayName
                Created     = (Get-RemediationObjectDateText -Object $sp -PropertyNames @('createdDateTime', 'CreatedDateTime'))
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
            if (Test-RemediationGraphAlreadyGone $_) {
                [pscustomobject]@{ Id = $id; Kind = $kind; Success = $true; Error = '' }
            } else {
                [pscustomobject]@{ Id = $id; Kind = $kind; Success = $false; Error = (Get-RemediationGraphErrorText $_) }
            }
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

function Convert-RemediationDateToText {
    param($Value, [int]$Depth = 0)
    if ($null -eq $Value -or $Depth -gt 4) { return '' }
    if ($Value -is [datetime]) {
        if ($Value.Year -le 1601) { return '' }
        try { return $Value.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ') } catch { return $Value.ToString('o') }
    }
    if ($Value -is [datetimeoffset]) {
        if ($Value.Year -le 1601) { return '' }
        return $Value.UtcDateTime.ToString('yyyy-MM-ddTHH:mm:ssZ')
    }
    if ($Value -is [string]) {
        $s = $Value.Trim()
        if ([string]::IsNullOrWhiteSpace($s) -or $s -eq '[object Object]') { return '' }
        return $s
    }
    if ($Value -is [valuetype]) { return '' }
    $props = @($Value.PSObject.Properties)
    foreach ($name in @('DateTime', 'dateTime', 'UtcDateTime', 'utcDateTime', 'Value', 'value', 'DisplayDateTime')) {
        $p = $props | Where-Object { $_.Name -eq $name } | Select-Object -First 1
        if ($p -and $null -ne $p.Value) {
            $inner = Convert-RemediationDateToText -Value $p.Value -Depth ($Depth + 1)
            if ($inner) { return $inner }
        }
    }
    return ''
}

function Get-RemediationObjectDateText {
    param($Object, [string[]]$PropertyNames)
    if ($null -eq $Object) { return '' }
    $wanted = @($PropertyNames)
    foreach ($name in $wanted) {
        $prop = $Object.PSObject.Properties | Where-Object { $_.Name -ieq $name } | Select-Object -First 1
        if (-not $prop) { continue }
        $text = Convert-RemediationDateToText $prop.Value
        if ($text) { return $text }
    }
    foreach ($prop in $Object.PSObject.Properties) {
        if ($prop.Name -notmatch 'WhenChanged|WhenCreated|LastModified|CreatedTime|CreationTime|ModifiedTime|startDateTime|createdDateTime|LitigationHoldDate|StartTime|EndTime|approximateLastSignIn') {
            continue
        }
        $text = Convert-RemediationDateToText $prop.Value
        if ($text) { return $text }
    }
    return ''
}

function Get-RemediationHydratedDateText {
    param(
        $Listed,
        [string]$Identity,
        [scriptblock]$Loader,
        [string[]]$PropertyNames,
        $State
    )
    $text = Get-RemediationObjectDateText -Object $Listed -PropertyNames $PropertyNames
    if ($text) { return $text }
    if ($State.Skip -or [string]::IsNullOrWhiteSpace($Identity) -or -not $Loader) { return '' }
    try {
        $full = & $Loader $Identity
        $text = Get-RemediationObjectDateText -Object $full -PropertyNames $PropertyNames
        if ($text) {
            $State.Hits++
            return $text
        }
        $State.Misses++
    } catch {
        $State.Misses++
    }
    if ($State.Hits -eq 0 -and $State.Misses -ge 3) { $State.Skip = $true }
    return ''
}

function Get-RemediationTransportRules {
    if (-not (Get-Command Get-TransportRule -ErrorAction SilentlyContinue)) {
        throw 'Get-TransportRule is not available. Complete Exchange Auth.'
    }
    $dateNames = @('WhenChanged', 'WhenChangedUTC', 'LastModifiedTime', 'LastModifiedDateTime', 'ModifiedTime')
    $hydrate = @{ Hits = 0; Misses = 0; Skip = $false }
    foreach ($rule in @(Get-TransportRule -ErrorAction Stop)) {
        $desc = [string]$rule.Description
        if ($desc.Length -gt 400) { $desc = $desc.Substring(0, 400) + '…' }
        $identity = [string]$rule.Identity
        $changed = Get-RemediationHydratedDateText -Listed $rule -Identity $identity -PropertyNames $dateNames -State $hydrate -Loader {
            param($id) Get-TransportRule -Identity $id -ErrorAction Stop
        }
        [pscustomobject]@{
            Identity    = $identity
            Name        = [string]$rule.Name
            Priority    = $rule.Priority
            State       = [string]$rule.State
            Mode        = [string]$rule.Mode
            Description = $desc
            WhenChanged = $changed
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
            WhenCreated = (Get-RemediationObjectDateText -Object $c -PropertyNames @('WhenCreated', 'WhenCreatedUTC', 'CreationTime', 'CreatedTime', 'WhenChanged', 'WhenChangedUTC'))
        }
    }
}

function Get-RemediationConnectors {
    $rows = [System.Collections.ArrayList]::new()
    $dateNames = @('WhenCreated', 'WhenCreatedUTC', 'CreationTime', 'CreatedTime', 'WhenChanged', 'WhenChangedUTC')
    $hydrate = @{ Hits = 0; Misses = 0; Skip = $false }
    try {
        foreach ($c in @(Convert-RemediationConnectorList -Connectors (Get-InboundConnector -ErrorAction Stop) -Direction 'Inbound')) {
            if (-not $c.WhenCreated) {
                $c.WhenCreated = Get-RemediationHydratedDateText -Listed $c -Identity $c.Identity -PropertyNames $dateNames -State $hydrate -Loader {
                    param($id) Get-InboundConnector -Identity $id -ErrorAction Stop
                }
            }
            [void]$rows.Add($c)
        }
    } catch {}
    try {
        foreach ($c in @(Convert-RemediationConnectorList -Connectors (Get-OutboundConnector -ErrorAction Stop) -Direction 'Outbound')) {
            if (-not $c.WhenCreated) {
                $c.WhenCreated = Get-RemediationHydratedDateText -Listed $c -Identity $c.Identity -PropertyNames $dateNames -State $hydrate -Loader {
                    param($id) Get-OutboundConnector -Identity $id -ErrorAction Stop
                }
            }
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

function Get-RemediationGraphErrorTextFromAny {
    param($ErrorRecordOrText)
    if ($ErrorRecordOrText -is [System.Management.Automation.ErrorRecord]) {
        return Get-RemediationGraphErrorText $ErrorRecordOrText
    }
    return [string]$ErrorRecordOrText
}

function Test-RemediationGraphAlreadyGone {
    param($ErrorRecordOrText)
    $text = Get-RemediationGraphErrorTextFromAny $ErrorRecordOrText
    if ([string]::IsNullOrWhiteSpace($text)) { return $false }
    return [bool]($text -match 'Request_ResourceNotFound|does not exist or one of its queried reference-property objects are not present')
}

function Test-RemediationIntuneUnavailable {
    param($ErrorRecordOrText)
    $text = Get-RemediationGraphErrorTextFromAny $ErrorRecordOrText
    if ([string]::IsNullOrWhiteSpace($text)) { return $false }
    return [bool]($text -match 'not applicable to target tenant')
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

function Test-RemediationGraphTokenExpiring {
    param([string]$Token, [int]$MarginMinutes = 5)
    $parts = $Token.Split('.')
    if ($parts.Count -lt 2) { return $false }
    try {
        $payload = $parts[1].Replace('-', '+').Replace('_', '/')
        switch ($payload.Length % 4) { 2 { $payload += '==' } 3 { $payload += '=' } }
        $claims = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($payload)) | ConvertFrom-Json
        if (-not $claims.exp) { return $false }
        $expires = [DateTimeOffset]::FromUnixTimeSeconds([int64]$claims.exp).UtcDateTime
        return ($expires -le [DateTime]::UtcNow.AddMinutes($MarginMinutes))
    } catch {
        return $false
    }
}

function Restore-RemediationGraphToken {
    # The worker restores the token only when Set-GraphRestBearerToken exists in its session.
    # If ExportUtils was unloaded, load it globally again and fetch a fresh app-only token.
    if (-not (Get-Command Set-GraphRestBearerToken -ErrorAction SilentlyContinue)) {
        $exportUtils = Join-Path $PSScriptRoot 'ExportUtils.psm1'
        if (Test-Path -LiteralPath $exportUtils) {
            Import-Module $exportUtils -Global -DisableNameChecking -ErrorAction SilentlyContinue
        }
    }
    if (-not (Get-Command Test-GraphRestBearerToken -ErrorAction SilentlyContinue)) { return }
    if ((Test-GraphRestBearerToken) -and -not (Test-RemediationGraphTokenExpiring ([string](Get-GraphRestBearerToken)))) { return }
    if (-not (Get-Command Get-GraphAppTokenFromWCM -ErrorAction SilentlyContinue)) {
        $appMod = Join-Path $PSScriptRoot 'GraphAppCredential.psm1'
        if (Test-Path -LiteralPath $appMod) {
            Import-Module $appMod -Global -DisableNameChecking -ErrorAction SilentlyContinue
        }
    }
    if (-not (Get-Command Get-GraphAppTokenFromWCM -ErrorAction SilentlyContinue)) { return }
    if (-not (Get-Command Get-ConnectionInformation -ErrorAction SilentlyContinue)) { return }
    $tenantId = ''
    try {
        $conn = @(Get-ConnectionInformation -ErrorAction Stop | Where-Object { $_.State -eq 'Connected' -and $_.TenantID }) | Select-Object -First 1
        if ($conn) { $tenantId = [string]$conn.TenantID }
    } catch {}
    if (-not $tenantId) { return }
    try {
        $fresh = Get-GraphAppTokenFromWCM -TenantId $tenantId
        if ($fresh) { Set-GraphRestBearerToken -Token ([string]$fresh) }
    } catch {}
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
    if ($GraphAuthenticated) { Restore-RemediationGraphToken }
    $caps = $null
    $needsGraph = $Command -match '^REMEDIATE_(REVOKE_SESSIONS|BLOCK|UNBLOCK|SIGNIN_STATUS|RESET_PASSWORD|LIST_AUTH_METHODS|DELETE_AUTH_METHODS|LIST_DEVICES|DELETE_DEVICES|LIST_APPS|DELETE_APPS|LIST_OAUTH_GRANTS|DELETE_OAUTH_GRANTS|REREGISTER_MFA|LIST_INTUNE|WIPE_INTUNE|RETIRE_INTUNE|LIST_ROLES|DELETE_ROLES|LIST_APP_CREDS|DELETE_APP_CREDS|HUNTER_GRAPH)'
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
                    [void]$details.Add("$($r.UserPrincipalName) : password reset ($mode). Do not send a password - send $($r.SsprUrl)")
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
            $goneIds = @($deleted | Where-Object { $_.Success -and $_.Id } | ForEach-Object { [string]$_.Id })
            if ($goneIds.Count -gt 0) {
                $remaining = @($remaining | Where-Object { -not $_.Id -or ($goneIds -notcontains [string]$_.Id) })
            }
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
            $goneIds = @($deleted | Where-Object { $_.Success -and $_.Id } | ForEach-Object { [string]$_.Id })
            if ($goneIds.Count -gt 0) {
                $remaining = @($remaining | Where-Object { -not $_.Id -or ($goneIds -notcontains [string]$_.Id) })
            }
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

        if ($Command -match '^REMEDIATE_HUNTER_LISTS') {
            $hunterUser = ''
            if ($Command -match '\|USER:([^|]+)') { $hunterUser = $Matches[1].Trim() }
            if (-not $hunterUser) { return 'REMEDIATE_FAILED:Missing mailbox' }
            $listErrors = @{}
            $rules = @()
            $users = @()
            $mobile = @()
            $transport = @()
            $connectors = @()
            try { $rules = @(Get-RemediationInboxRules -Mailbox $hunterUser) } catch { $listErrors.rules = $_.Exception.Message }
            try { $users = @(Get-RemediationMailboxAccess -UserPrincipalNames @($hunterUser)) } catch { $listErrors.mailbox = $_.Exception.Message }
            try { $mobile = @(Get-RemediationMobileDevices -UserPrincipalNames @($hunterUser)) } catch { $listErrors.mobile = $_.Exception.Message }
            try { $transport = @(Get-RemediationTransportRules) } catch { $listErrors.transport = $_.Exception.Message }
            try { $connectors = @(Get-RemediationConnectors) } catch { $listErrors.connectors = $_.Exception.Message }
            $listPayload = @{
                Rules      = @($rules)
                Users      = @($users)
                Devices    = @($mobile)
                Transport  = @($transport)
                Connectors = @($connectors)
                Errors     = $listErrors
            }
            if ($Command -match '\|WITHGRAPH\b') {
                $packed = @(Invoke-RemediationHunterGraphPack -User $hunterUser)
                $byKind = @{}
                foreach ($item in $packed) { $byKind[[string]$item.Kind] = $item }
                $graphPayload = @{
                    Methods = @(if ($byKind.auth) { $byKind.auth.Rows })
                    Devices = @(if ($byKind.devices) { $byKind.devices.Rows })
                    Grants  = @(if ($byKind.oauth) { $byKind.oauth.Rows })
                    Intune  = @(if ($byKind.intune) { $byKind.intune.Rows })
                    Roles   = @(if ($byKind.roles) { $byKind.roles.Rows })
                    Apps    = @(if ($byKind.apps) { $byKind.apps.Rows })
                    Errors  = @{
                        auth    = $(if ($byKind.auth) { [string]$byKind.auth.Error })
                        devices = $(if ($byKind.devices) { [string]$byKind.devices.Error })
                        oauth   = $(if ($byKind.oauth) { [string]$byKind.oauth.Error })
                        intune  = $(if ($byKind.intune) { [string]$byKind.intune.Error })
                        roles   = $(if ($byKind.roles) { [string]$byKind.roles.Error })
                        apps    = $(if ($byKind.apps) { [string]$byKind.apps.Error })
                    }
                }
                $payload = @{ Lists = $listPayload; Graph = $graphPayload } | ConvertTo-Json -Compress -Depth 8
                return "REMEDIATE_HUNTER_PACK:$payload"
            }
            $payload = $listPayload | ConvertTo-Json -Compress -Depth 6
            return "REMEDIATE_HUNTER_LISTS:$payload"
        }

        if ($Command -match '^REMEDIATE_HUNTER_GRAPH') {
            $hunterUser = ''
            if ($Command -match '\|USER:([^|]+)') { $hunterUser = $Matches[1].Trim() }
            if (-not $hunterUser) { return 'REMEDIATE_FAILED:Missing mailbox' }
            $packed = @(Invoke-RemediationHunterGraphPack -User $hunterUser)
            $byKind = @{}
            foreach ($item in $packed) { $byKind[[string]$item.Kind] = $item }
            foreach ($pair in @(
                @{ Kind = 'auth'; Action = 'list-auth-methods'; Noun = 'method' },
                @{ Kind = 'devices'; Action = 'list-devices'; Noun = 'device' },
                @{ Kind = 'apps'; Action = 'list-apps'; Noun = 'app' }
            )) {
                $item = $byKind[$pair.Kind]
                $err = if ($item) { [string]$item.Error } else { 'missing' }
                $count = if ($item) { @($item.Rows).Count } else { 0 }
                $result = if ($err) { 'failed' } else { 'success' }
                $detail = if ($err) { $err } else { "$count $($pair.Noun)(s)" }
                Write-RemediationAuditCsv -Folder $AuditFolder -Tenant $tenantLabel -ClientNumber $ClientNumber -Upn $(if ($pair.Kind -eq 'apps') { '' } else { $hunterUser }) -Action $pair.Action -Result $result -Detail $detail
            }
            $payload = @{
                Methods = @(if ($byKind.auth) { $byKind.auth.Rows })
                Devices = @(if ($byKind.devices) { $byKind.devices.Rows })
                Grants  = @(if ($byKind.oauth) { $byKind.oauth.Rows })
                Intune  = @(if ($byKind.intune) { $byKind.intune.Rows })
                Roles   = @(if ($byKind.roles) { $byKind.roles.Rows })
                Apps    = @(if ($byKind.apps) { $byKind.apps.Rows })
                Errors  = @{
                    auth    = $(if ($byKind.auth) { [string]$byKind.auth.Error })
                    devices = $(if ($byKind.devices) { [string]$byKind.devices.Error })
                    oauth   = $(if ($byKind.oauth) { [string]$byKind.oauth.Error })
                    intune  = $(if ($byKind.intune) { [string]$byKind.intune.Error })
                    roles   = $(if ($byKind.roles) { [string]$byKind.roles.Error })
                    apps    = $(if ($byKind.apps) { [string]$byKind.apps.Error })
                }
            } | ConvertTo-Json -Compress -Depth 6
            return "REMEDIATE_HUNTER_GRAPH:$payload"
        }

        if ($Command -match '^REMEDIATE_BEC_AUDIT') {
            return (Get-RemediationBecAudit -Command $Command -GraphAuthenticated $GraphAuthenticated -ExchangeAuthenticated $ExchangeAuthenticated)
        }

        $extended = Invoke-RemediationExtendedCommand -Command $Command -Capabilities $caps -AuditFolder $AuditFolder -TenantLabel $tenantLabel -ClientNumber $ClientNumber
        if ($null -ne $extended) { return $extended }

        return "REMEDIATE_FAILED:Unknown remediation command"
    } catch {
        $err = if ($needsGraph) { Get-RemediationGraphErrorText $_ } else { $_.Exception.Message }
        return "REMEDIATE_FAILED:$err"
    }
}

function Test-RemediationLogLicenseLimit {
    param([string]$Message)
    return [string]$Message -match '(?i)licen[cs]|premium|retention|more than \d+ days|within the last \d+|older than \d+|only available for the last|maximum retention|start date .{0,80}(past|ago|retention|days)|start time .{0,80}(past|ago|within|days)'
}

function Get-RemediationShorterLogDays {
    param([int]$Days)
    foreach ($step in @(60, 30, 21, 14, 7)) {
        if ($step -lt $Days) { return $step }
    }
    return 0
}

function Get-RemediationBecAuditClip {
    param([string]$Text, [int]$Max = 120)
    $text = [string]$Text
    if ($text.Length -gt $Max) { return $text.Substring(0, $Max) }
    return $text
}

function Test-RemediationBecAuditPriority {
    param($Row)
    $op = [string]$Row.Operations
    if (-not $op) { $op = [string]$Row.Operation }
    $rt = [string]$Row.RecordType
    if ($rt -match 'ComplianceDLP|DlpEndpoint|ThreatIntelligence|SecurityCompliance') { return $true }
    if ($rt -match 'SharePointSharing|SharePointFile|OneDrive') { return $true }
    if ($op -match 'InboxRule|Forwarding|Set-Mailbox|MailboxPermission|SendAs|SendOnBehalf|MailItemsAccessed|FileDownloaded|FileSyncDownloaded|AnonymousLink|SharingSet|SecureLink|CompanyLink|SoftDelete|HardDelete|MoveToDeletedItems|FileCopied|FileUploaded|Consent to application|Add service principal|Add member to role|Add app role|Reset user password|Disable Strong Authentication') { return $true }
    return $false
}

function ConvertTo-RemediationBecAuditRow {
    param($Row)
    $op = [string]$Row.Operations
    if (-not $op) { $op = [string]$Row.Operation }
    $target = [string]$Row.ObjectId
    if (-not $target -and $Row.AuditData) {
        try {
            $event = if ($Row.AuditData -is [string]) { $Row.AuditData | ConvertFrom-Json -ErrorAction Stop } else { $Row.AuditData }
            if ($event.ObjectId) { $target = [string]$event.ObjectId }
            elseif ($event.SourceFileName) { $target = [string]$event.SourceFileName }
            elseif ($event.FileName) { $target = [string]$event.FileName }
        } catch { }
    }
    return [pscustomobject]@{
        Created    = [string]$Row.CreationDate
        RecordType = [string]$Row.RecordType
        Operation  = $op
        User       = [string]$Row.UserIds
        Target     = (Get-RemediationBecAuditClip $target)
    }
}

function Get-RemediationBecAuditLogSet {
    param([string]$Command)
    if ($Command -notmatch '\|LOGS:([^|]+)') { return $null }
    $set = @{}
    foreach ($part in ($Matches[1] -split ',')) {
        $key = $part.Trim().ToLowerInvariant()
        if ($key -in @('trace', 'ual', 'sharing', 'files', 'dlp', 'alerts', 'signins', 'apps')) {
            $set[$key] = $true
        }
    }
    if ($set.Count -eq 0) { return $null }
    return $set
}

function Start-RemediationGraphGets {
    param([Parameter(Mandatory = $true)][hashtable]$Requests)
    $token = Get-GraphRestBearerToken
    if ([string]::IsNullOrWhiteSpace($token)) { throw 'Graph REST bearer token is not set.' }
    $client = [System.Net.Http.HttpClient]::new()
    $client.Timeout = [TimeSpan]::FromSeconds(90)
    $client.DefaultRequestHeaders.Authorization = [System.Net.Http.Headers.AuthenticationHeaderValue]::new('Bearer', $token)
    $client.DefaultRequestHeaders.Accept.ParseAdd('application/json')
    [void]$client.DefaultRequestHeaders.TryAddWithoutValidation('ConsistencyLevel', 'eventual')
    $tasks = @{}
    foreach ($name in @($Requests.Keys)) {
        $tasks[$name] = $client.GetAsync([string]$Requests[$name])
    }
    return ,@{ Client = $client; Tasks = $tasks }
}

function Complete-RemediationGraphGets {
    param($Batch)
    $results = @{}
    try {
        $pending = New-Object System.Collections.Generic.List[System.Threading.Tasks.Task]
        foreach ($name in @($Batch.Tasks.Keys)) { [void]$pending.Add($Batch.Tasks[$name]) }
        if ($pending.Count -gt 0) {
            try {
                [void][System.Threading.Tasks.Task]::WaitAll($pending.ToArray(), [TimeSpan]::FromSeconds(95))
            } catch {
                # A failed request is recorded on that task below.
            }
        }
        foreach ($name in @($Batch.Tasks.Keys)) {
            $task = $Batch.Tasks[$name]
            if (-not $task.IsCompleted) {
                $results[$name] = @{ Ok = $false; Error = 'Graph request timed out.'; Json = $null }
                continue
            }
            if ($task.IsFaulted) {
                $results[$name] = @{ Ok = $false; Error = $task.Exception.GetBaseException().Message; Json = $null }
                continue
            }
            $response = $task.Result
            $read = $response.Content.ReadAsStringAsync()
            [void]$read.Wait()
            $body = [string]$read.Result
            $code = [int]$response.StatusCode
            if ($code -ge 400) {
                if ($body.Length -gt 300) { $body = $body.Substring(0, 300) }
                $results[$name] = @{ Ok = $false; Error = "$code $body"; Json = $null }
                continue
            }
            $json = $null
            if ($body) { $json = $body | ConvertFrom-Json }
            $results[$name] = @{ Ok = $true; Error = ''; Json = $json }
        }
    } finally {
        if ($Batch.Client) { $Batch.Client.Dispose() }
    }
    return ,$results
}

function Invoke-RemediationHunterGraphPack {
    param([Parameter(Mandatory = $true)][string]$User)
    if ([string]::IsNullOrWhiteSpace((Get-GraphRestBearerToken))) { throw 'Graph REST bearer token is not set.' }
    $calls = [ordered]@{
        auth    = { @(Get-RemediationAuthMethods -UserPrincipalNames @($User)) }
        devices = { @(Get-RemediationUserDevices -UserPrincipalNames @($User)) }
        oauth   = { @(Get-RemediationOAuthGrants -UserPrincipalNames @($User)) }
        intune  = { @(Get-RemediationIntuneDevices -UserPrincipalNames @($User)) }
        roles   = { @(Get-RemediationDirectoryRoles -UserPrincipalNames @($User)) }
        apps    = { @(Get-RemediationDirectoryApps -Top 200) }
    }
    $packed = foreach ($kind in @($calls.Keys)) {
        $rows = @()
        $err = ''
        try { $rows = & $calls[$kind] }
        catch {
            $err = $_.Exception.Message
            if ($_.Exception.InnerException) { $err = $_.Exception.InnerException.Message }
        }
        [pscustomobject]@{ Kind = $kind; Error = $err; Rows = @($rows) }
    }
    return @($packed)
}

function ConvertTo-RemediationReportedPlainText {
    param([string]$Html)
    $text = [string]$Html
    if ($text.Length -gt 20000) { $text = $text.Substring(0, 20000) }
    $text = [regex]::Replace($text, '(?is)<script\b.*?</script>', ' ')
    $text = [regex]::Replace($text, '(?is)<style\b.*?</style>', ' ')
    $text = [regex]::Replace($text, '(?i)<br\s*/?>', "`n")
    $text = [regex]::Replace($text, '(?i)</p>', "`n")
    $text = [regex]::Replace($text, '(?s)<[^>]+>', ' ')
    $text = [System.Net.WebUtility]::HtmlDecode($text)
    $text = [regex]::Replace($text, '[ \t]+', ' ')
    $text = [regex]::Replace($text, '(\r?\n[ \t]*){2,}', "`n")
    return $text.Trim()
}

function Get-RemediationReportedLinks {
    param([string]$Html, [string]$Plain)
    $links = New-Object System.Collections.Generic.List[string]
    $seen = @{}
    $add = {
        param([string]$Raw)
        $link = [System.Net.WebUtility]::HtmlDecode([string]$Raw).Trim().TrimEnd('.,);]')
        if ($link -notmatch '^(?i)https?://') { return }
        if ($link.Length -gt 140) { $link = $link.Substring(0, 140) }
        $key = $link.ToLowerInvariant()
        if ($seen.ContainsKey($key)) { return }
        $seen[$key] = $true
        [void]$links.Add($link)
    }
    foreach ($hit in [regex]::Matches([string]$Html, '(?i)(?:href|src)\s*=\s*"([^"]+)"')) { if ($links.Count -ge 8) { break }; & $add $hit.Groups[1].Value }
    if ($links.Count -lt 8) {
        foreach ($hit in [regex]::Matches([string]$Html, "(?i)(?:href|src)\s*=\s*'([^']+)'")) { if ($links.Count -ge 8) { break }; & $add $hit.Groups[1].Value }
    }
    if ($links.Count -lt 8) {
        foreach ($hit in [regex]::Matches([string]$Plain, 'https?://[^\s<>"'']+')) { if ($links.Count -ge 8) { break }; & $add $hit.Value }
    }
    return @($links.ToArray())
}

function Get-RemediationReportedMessages {
    param(
        [Parameter(Mandatory = $true)][string]$User,
        [Parameter(Mandatory = $true)][string[]]$Senders,
        [datetime]$Start,
        [datetime]$End,
        [string]$SubjectHint = ''
    )
    $encUser = [Uri]::EscapeDataString($User.Trim())
    $select = 'id,subject,from,replyTo,receivedDateTime,internetMessageHeaders,bodyPreview,hasAttachments,body'
    $found = New-Object System.Collections.Generic.List[object]
    $hint = [string]$SubjectHint
    foreach ($sender in @($Senders | Select-Object -Unique -First 3)) {
        $addr = [string]$sender
        if ($addr -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') { continue }
        $search = [Uri]::EscapeDataString(('"{0}"' -f ('from:' + $addr)))
        $uri = "https://graph.microsoft.com/v1.0/users/$encUser/messages?`$search=$search&`$top=6&`$select=$select"
        try {
            $page = Invoke-GraphRestRequest -Uri $uri -Headers @{ ConsistencyLevel = 'eventual' }
        } catch {
            $detail = $_.Exception.Message
            if ($_.ErrorDetails -and $_.ErrorDetails.Message) { $detail = [string]$_.ErrorDetails.Message }
            if ($detail -match '403|ErrorAccessDenied|Authorization_RequestDenied|Access is denied') {
                throw 'Graph refused to read the mailbox. The signed-in app needs Mail.Read.'
            }
            throw $detail
        }
        foreach ($msg in @($page.value)) {
            if (-not $msg) { continue }
            $received = [datetimeoffset]::MinValue
            $receivedText = [string]$msg.receivedDateTime
            if ($receivedText -and -not [datetimeoffset]::TryParse($receivedText, [ref]$received)) { continue }
            if ($received -ne [datetimeoffset]::MinValue) {
                $utc = $received.UtcDateTime
                if ($utc -lt $Start.AddHours(-2) -or $utc -gt $End.AddHours(2)) { continue }
            }
            $fromAddr = ''
            $fromName = ''
            if ($msg.from -and $msg.from.emailAddress) {
                $fromAddr = [string]$msg.from.emailAddress.address
                $fromName = [string]$msg.from.emailAddress.name
            }
            $replyTo = @()
            foreach ($who in @($msg.replyTo)) {
                if ($who.emailAddress -and $who.emailAddress.address) { $replyTo += [string]$who.emailAddress.address }
            }
            $headerMap = @{}
            foreach ($header in @($msg.internetMessageHeaders)) {
                $name = [string]$header.name
                if ($name -notmatch '^(?i)(from|reply-to|return-path|authentication-results|x-forefront-antispam-report)$') { continue }
                $key = $name.ToLowerInvariant()
                if (-not $headerMap.ContainsKey($key)) { $headerMap[$key] = [string]$header.value }
            }
            $auth = [string]$headerMap['authentication-results']
            $front = [string]$headerMap['x-forefront-antispam-report']
            $authBits = New-Object System.Collections.Generic.List[string]
            foreach ($name in @('spf', 'dkim', 'dmarc', 'compauth')) {
                $hit = [regex]::Match($auth, ('(?i)\b{0}=([a-z]+)' -f $name))
                if ($hit.Success) { [void]$authBits.Add(('{0}={1}' -f $name, $hit.Groups[1].Value.ToLowerInvariant())) }
            }
            foreach ($name in @('SCL', 'CAT', 'CTRY', 'CIP')) {
                $hit = [regex]::Match($front, ('(?i)\b{0}:([^;]+)' -f $name))
                if ($hit.Success) { [void]$authBits.Add(('{0}:{1}' -f $name, $hit.Groups[1].Value.Trim())) }
            }
            $rawBody = ''
            if ($msg.body -and $msg.body.content) { $rawBody = [string]$msg.body.content }
            $plain = ''
            if ($rawBody) { $plain = ConvertTo-RemediationReportedPlainText $rawBody }
            if (-not $plain) { $plain = [string]$msg.bodyPreview }
            $visible = [regex]::Replace($plain, "[\u200B-\u200D\uFEFF\u2060\u180E]", '').Trim()
            $links = @(Get-RemediationReportedLinks -Html $rawBody -Plain $plain)
            $files = New-Object System.Collections.Generic.List[string]
            if ($msg.hasAttachments -and $msg.id) {
                $encId = [Uri]::EscapeDataString([string]$msg.id)
                try {
                    $attached = Invoke-GraphRestRequest -Uri "https://graph.microsoft.com/v1.0/users/$encUser/messages/$encId/attachments?`$select=name,contentType,size"
                    foreach ($file in @($attached.value)) {
                        if (-not $file -or -not $file.name) { continue }
                        $size = ''
                        if ($null -ne $file.size) { $size = [string]$file.size }
                        [void]$files.Add(('{0} {1}' -f [string]$file.name, $size).Trim())
                        if ($files.Count -ge 8) { break }
                    }
                } catch {
                    [void]$files.Add('attachment list failed')
                }
            }
            $preview = $visible
            if (-not $preview -and $plain) { $preview = 'No visible text. The body is zero-width characters.' }
            elseif ($preview.Length -gt 400) { $preview = $preview.Substring(0, 400) }
            $subject = [string]$msg.subject
            $boost = 0
            if ($hint -and $subject -match [regex]::Escape($hint)) { $boost = 1 }
            [void]$found.Add([pscustomobject]@{
                Boost    = $boost
                Received = $receivedText
                From     = $fromAddr
                FromName = (Get-RemediationBecAuditClip $fromName 80)
                ReplyTo  = (($replyTo | Select-Object -Unique) -join ', ')
                Subject  = (Get-RemediationBecAuditClip $subject 120)
                Auth     = (Get-RemediationBecAuditClip (($authBits -join ' ')) 240)
                Links    = @($links.ToArray())
                Files    = @($files.ToArray())
                Preview  = $preview
            })
        }
    }
    $picked = New-Object System.Collections.Generic.List[object]
    foreach ($row in @($found | Sort-Object Boost -Descending)) {
        if ($picked.Count -ge 3) { break }
        [void]$picked.Add([pscustomobject]@{
            Received = $row.Received
            From     = $row.From
            FromName = $row.FromName
            ReplyTo  = $row.ReplyTo
            Subject  = $row.Subject
            Auth     = $row.Auth
            Links    = $row.Links
            Files    = $row.Files
            Preview  = $row.Preview
        })
    }
    return @($picked.ToArray())
}

function Get-RemediationBecAudit {
    param(
        [Parameter(Mandatory = $true)][string]$Command,
        [bool]$GraphAuthenticated,
        [bool]$ExchangeAuthenticated
    )
    $notes = New-Object System.Collections.ArrayList
    $user = ''
    if ($Command -match '\|USER:([^|]+)') { $user = $Matches[1].Trim() }
    if (-not $user) { return 'REMEDIATE_FAILED:Missing mailbox' }
    $leadSenders = @()
    if ($Command -match '\|FROM:([^|]+)') {
        $leadSenders = @($Matches[1].Split(',') | ForEach-Object { $_.Trim().ToLowerInvariant() } | Where-Object { $_ -match '^[^@\s]+@[^@\s]+\.[^@\s]+$' } | Select-Object -Unique -First 3)
    }
    $subjectHint = ''
    if ($Command -match '\|SUBJECT:([^|]+)') { $subjectHint = $Matches[1].Trim() }
    $start = (Get-Date).ToUniversalTime().AddDays(-2)
    $end = (Get-Date).ToUniversalTime()
    if ($Command -match '\|START:([^|]+)') {
        $parsed = [datetime]::MinValue
        if ([datetime]::TryParse($Matches[1].Trim(), [ref]$parsed)) { $start = $parsed.ToUniversalTime() }
    }
    if ($Command -match '\|END:([^|]+)') {
        $parsedEnd = [datetime]::MinValue
        if ([datetime]::TryParse($Matches[1].Trim(), [ref]$parsedEnd)) { $end = $parsedEnd.ToUniversalTime() }
    }
    if ($end -lt $start) { $end = (Get-Date).ToUniversalTime() }
    $days = [int][Math]::Ceiling(($end - $start).TotalDays)
    if ($days -gt 90) {
        $start = $end.AddDays(-90)
        $days = 90
        [void]$notes.Add('Lookback was capped at 90 days. Unified audit retention is often 90 days, and longer only when the tenant has extended audit.')
    }
    $only = Get-RemediationBecAuditLogSet -Command $Command
    $wantLog = {
        param([string]$Name)
        if ($null -eq $only) { return $true }
        return [bool]$only.ContainsKey($Name)
    }
    $pulledStart = $start.ToLocalTime().ToString('yyyy-MM-dd h:mm tt')
    $pulledEnd = $end.ToLocalTime().ToString('yyyy-MM-dd h:mm tt')
    $pulledZone = 'local time'
    try {
        $tzInfo = [TimeZoneInfo]::Local
        $zoneName = if ($tzInfo.IsDaylightSavingTime($end.ToLocalTime())) { $tzInfo.DaylightName } else { $tzInfo.StandardName }
        if ($zoneName) { $pulledZone = $zoneName }
    } catch {}
    if ($null -eq $only) {
        [void]$notes.Add(('Pulled {0} through {1} ({2}, {3} days). Message trace uses the tenant API limit (legacy about 10 days, V2 about 90). Unified audit stops where the license retention stops.' -f $pulledStart, $pulledEnd, $pulledZone, $days))
    } else {
        $onlyNames = @($only.Keys | Sort-Object) -join ', '
        [void]$notes.Add(('Pulled {0} through {1} ({2}, {3} days) for {4} only. Other logs stay on the earlier window.' -f $pulledStart, $pulledEnd, $pulledZone, $days, $onlyNames))
    }

    $trace = New-Object System.Collections.ArrayList
    $ual = New-Object System.Collections.ArrayList
    $dlp = New-Object System.Collections.ArrayList
    $files = New-Object System.Collections.ArrayList
    $sharing = New-Object System.Collections.ArrayList
    $alerts = New-Object System.Collections.ArrayList
    $signIns = New-Object System.Collections.ArrayList
    $apps = New-Object System.Collections.ArrayList
    $blasts = New-Object System.Collections.ArrayList
    $reported = New-Object System.Collections.ArrayList
    $posture = $null
    $recordTypes = @(
        'ExchangeAdmin', 'ExchangeItem', 'ExchangeItemGroup', 'AzureActiveDirectory',
        'SharePointFileOperation', 'SharePointSharingOperation', 'OneDrive',
        'ComplianceDLPExchange', 'ComplianceDLPSharePoint', 'ComplianceDLPEndpoint',
        'ThreatIntelligence', 'SecurityComplianceAlerts'
    )
    if ($null -ne $only) {
        $pickedTypes = New-Object System.Collections.Generic.List[string]
        if ($only.ContainsKey('ual')) {
            foreach ($t in @('ExchangeAdmin', 'ExchangeItem', 'ExchangeItemGroup', 'AzureActiveDirectory', 'ThreatIntelligence', 'SecurityComplianceAlerts')) {
                [void]$pickedTypes.Add($t)
            }
        }
        if ($only.ContainsKey('sharing') -or $only.ContainsKey('files')) {
            foreach ($t in @('SharePointFileOperation', 'SharePointSharingOperation', 'OneDrive')) {
                [void]$pickedTypes.Add($t)
            }
        }
        if ($only.ContainsKey('dlp')) {
            foreach ($t in @('ComplianceDLPExchange', 'ComplianceDLPSharePoint', 'ComplianceDLPEndpoint')) {
                [void]$pickedTypes.Add($t)
            }
        }
        $recordTypes = $pickedTypes.ToArray()
    }
    $sharingCap = 25
    if ($null -ne $only -and $only.ContainsKey('sharing')) { $sharingCap = 60 }

    $graphBatch = $null
    $mfaJob = $null
    $graphTokenMissing = $false
    $needGraph = ($null -eq $only) -or (& $wantLog 'signins') -or (& $wantLog 'alerts') -or (& $wantLog 'apps')
    if ($GraphAuthenticated -and $needGraph) {
        $graphToken = ''
        if (Get-Command Get-GraphRestBearerToken -ErrorAction SilentlyContinue) { $graphToken = [string](Get-GraphRestBearerToken) }
        if ([string]::IsNullOrWhiteSpace($graphToken)) {
            $graphTokenMissing = $true
            [void]$notes.Add('Graph auth was present, but the bearer token was not set, so MFA posture, sign-ins, security alerts, and new app registrations were not pulled.')
        } else {
            $graphUris = @{}
            $startZ = $start.ToString('yyyy-MM-ddTHH:mm:ssZ')
            $endZ = $end.ToString('yyyy-MM-ddTHH:mm:ssZ')
            $safeUser = $user.Replace("'", "''")
            $signUserId = ''
            $signByUserId = $false
            if (& $wantLog 'signins') {
                try {
                    $graphUser = Get-MgUserCompat -UserId $user -Property id
                    if ($graphUser.id) { $signUserId = [string]$graphUser.id }
                    elseif ($graphUser.Id) { $signUserId = [string]$graphUser.Id }
                } catch {
                    [void]$notes.Add(('Sign-in user id lookup failed: {0}' -f $_.Exception.Message))
                }
                if ($signUserId) {
                    $signByUserId = $true
                    $filter = "createdDateTime ge $startZ and createdDateTime le $endZ and userId eq '$signUserId'"
                } else {
                    $filter = "userPrincipalName eq '$safeUser' and createdDateTime ge $startZ and createdDateTime le $endZ"
                }
                # Without an explicit event type Graph returns interactive sign-ins only; token refreshes from mail clients are non-interactive.
                $filter = "$filter and signInEventTypes/any(t: t eq 'interactiveUser' or t eq 'nonInteractiveUser')"
                $graphUris.signins = 'https://graph.microsoft.com/beta/auditLogs/signIns?$top=500&$filter=' + [Uri]::EscapeDataString($filter)
            }
            if (& $wantLog 'apps') {
                $filter = "createdDateTime ge $startZ"
                $graphUris.apps = 'https://graph.microsoft.com/v1.0/applications?$top=20&$select=displayName,appId,createdDateTime&$filter=' + [Uri]::EscapeDataString($filter)
            }
            if (& $wantLog 'alerts') {
                $filter = "createdDateTime ge $startZ"
                $graphUris.alerts = 'https://graph.microsoft.com/v1.0/security/alerts?$top=50&$filter=' + [Uri]::EscapeDataString($filter)
            }
            if ($graphUris.Count -gt 0) {
                try { $graphBatch = Start-RemediationGraphGets -Requests $graphUris }
                catch { [void]$notes.Add(('Graph log pull failed to start: {0}' -f $_.Exception.Message)) }
            }
            if ($null -eq $only -and (Get-Command Get-MfaCoverageReport -ErrorAction SilentlyContinue)) {
                Import-Module ThreadJob -ErrorAction SilentlyContinue
                if (Get-Command Start-ThreadJob -ErrorAction SilentlyContinue) {
                    $exportUtils = Join-Path $PSScriptRoot 'ExportUtils.psm1'
                    try {
                        $mfaJob = Start-ThreadJob -ScriptBlock {
                            param($ExportUtils, $Token, $Mailbox)
                            Import-Module $ExportUtils -DisableNameChecking -ErrorAction Stop
                            Set-GraphRestBearerToken -Token $Token
                            Get-MfaCoverageReport -SelectedUsers @($Mailbox)
                        } -ArgumentList $exportUtils, $graphToken, $user
                    } catch {
                        $mfaJob = $null
                    }
                }
            }
        }
    }

    if ($null -eq $only -and $leadSenders.Count -gt 0) {
        if (-not $GraphAuthenticated -or $graphTokenMissing) {
            [void]$notes.Add(('Reported message from {0} was not opened. Graph auth was not available.' -f ($leadSenders -join ', ')))
        } else {
            try {
                $opened = @(Get-RemediationReportedMessages -User $user -Senders $leadSenders -Start $start -End $end -SubjectHint $subjectHint)
                foreach ($row in $opened) { if ($row) { [void]$reported.Add($row) } }
                if ($reported.Count -eq 0) {
                    [void]$notes.Add(('No message from {0} was found in this mailbox for this window.' -f ($leadSenders -join ', ')))
                } else {
                    [void]$notes.Add(('Opened {0} reported message(s) from the mailbox.' -f $reported.Count))
                }
            } catch {
                [void]$notes.Add(('Reported message was not opened: {0}' -f $_.Exception.Message))
            }
        }
    }

    if (-not $ExchangeAuthenticated) {
        [void]$notes.Add('Exchange auth was not available, so inbox state, message trace, and unified audit were not pulled.')
    } else {
        if (-not (Get-Command Get-ExchangeMessageTrace -ErrorAction SilentlyContinue) -and (Test-Path -LiteralPath (Join-Path $PSScriptRoot 'ExportUtils.psm1'))) {
            Import-Module (Join-Path $PSScriptRoot 'ExportUtils.psm1') -Global -DisableNameChecking -ErrorAction SilentlyContinue
        }
        $traceReady = (Get-Command Get-MessageTraceV2 -ErrorAction SilentlyContinue) -or (Get-Command Get-MessageTrace -ErrorAction SilentlyContinue)
        if ((& $wantLog 'trace') -and (Get-Command Get-ExchangeMessageTrace -ErrorAction SilentlyContinue)) {
            try {
                $traceArgs = @{
                    StartDate = $start
                    EndDate = $end
                    SelectedUsers = @($user)
                }
                $traceCmd = Get-Command Get-ExchangeMessageTrace -ErrorAction SilentlyContinue
                if ($traceCmd -and $traceCmd.Parameters.ContainsKey('NoGraph')) { $traceArgs.NoGraph = $true }
                $rows = Get-ExchangeMessageTrace @traceArgs
                if ($null -eq $rows) { $rows = @() }
                $large = New-Object System.Collections.Generic.List[object]
                $blastGroups = @{}
                foreach ($row in $rows) {
                    if (-not $row) { continue }
                    $subject = Get-RemediationBecAuditClip ([string]$row.Subject) 80
                    $sizeText = ''
                    $sizeBytes = 0L
                    $rawSize = $row.Size
                    if ($null -ne $rawSize -and [string]$rawSize -match '^\d+$') {
                        $sizeBytes = [int64]$rawSize
                        $sizeText = [string]$sizeBytes
                    }
                    $item = [pscustomobject]@{
                        Received  = [string]$row.Received
                        Sender    = [string]$row.SenderAddress
                        Recipient = [string]$row.RecipientAddress
                        Status    = [string]$row.Status
                        Subject   = $subject
                        Size      = $sizeText
                    }
                    if ($trace.Count -lt 40) {
                        [void]$trace.Add($item)
                    } elseif ($sizeBytes -ge 5242880 -and $large.Count -lt 8) {
                        [void]$large.Add($item)
                    }
                    $senderAddr = [string]$row.SenderAddress
                    $recipAddr = [string]$row.RecipientAddress
                    $statusText = [string]$row.Status
                    $mailboxKey = $user.Trim().ToLowerInvariant()
                    if ($senderAddr.Trim().ToLowerInvariant() -eq $mailboxKey -and $statusText -match '^(?i)delivered\b' -and $recipAddr.Trim() -and $recipAddr.Trim().ToLowerInvariant() -ne $mailboxKey) {
                        $groupKey = ([string]$row.Subject).Trim().ToLowerInvariant()
                        if (-not $blastGroups.ContainsKey($groupKey)) {
                            $blastGroups[$groupKey] = [pscustomobject]@{
                                Subject    = $(if ($subject.Trim()) { $subject.Trim() } else { '(no subject)' })
                                Received   = [string]$row.Received
                                Recipients = New-Object System.Collections.ArrayList
                                Seen       = @{}
                            }
                        }
                        $bucket = $blastGroups[$groupKey]
                        $addrKey = $recipAddr.Trim().ToLowerInvariant()
                        if (-not $bucket.Seen.ContainsKey($addrKey) -and $bucket.Recipients.Count -lt 2000) {
                            $bucket.Seen[$addrKey] = $true
                            [void]$bucket.Recipients.Add($recipAddr.Trim())
                        }
                    }
                }
                foreach ($item in $large) { [void]$trace.Add($item) }
                foreach ($bucket in $blastGroups.Values) {
                    if ($bucket.Recipients.Count -ge 5) {
                        [void]$blasts.Add([pscustomobject]@{
                            Subject    = [string]$bucket.Subject
                            Received   = [string]$bucket.Received
                            Count      = [int]$bucket.Recipients.Count
                            Recipients = $bucket.Recipients
                        })
                    }
                }
                if ($rows.Count -gt 40) {
                    $extra = if ($large.Count) { " $([int]$large.Count) larger message(s) of at least 5 MB were kept past that cap." } else { '' }
                    [void]$notes.Add(('Message trace was capped at 40 rows for the review.{0}' -f $extra))
                }
            } catch {
                [void]$notes.Add(('Message trace failed: {0}' -f $_.Exception.Message))
            }
        } elseif (& $wantLog 'trace') {
            $traceNote = if ($traceReady) {
                'Message trace was not pulled. The Exchange session is signed in, but the trace helper did not load. Run the threat check again.'
            } else {
                'Message trace was not pulled. Get-MessageTraceV2 is not in this Exchange session. Re-run Exchange Auth.'
            }
            [void]$notes.Add($traceNote)
        }
        $auditReady = Get-Command Search-UnifiedAuditLog -ErrorAction SilentlyContinue
        $needAudit = (& $wantLog 'ual') -or (& $wantLog 'sharing') -or (& $wantLog 'files') -or (& $wantLog 'dlp')
        if ($needAudit -and $auditReady -and (Get-Command Get-UnifiedAuditLogs -ErrorAction SilentlyContinue)) {
            $ualDays = $days
            $ualStart = $start
            $ualTries = 0
            $ualShortened = $false
            $rows = @()
            $licenseHit = @()
            while ($ualTries -lt 6) {
                $ualTries++
                $searchErrors = New-Object System.Collections.Generic.List[string]
                try {
                    $ualArgs = @{
                        StartDate = $ualStart
                        EndDate = $end
                        SelectedUsers = @($user)
                        RecordTypes = $recordTypes
                        IncludeExchangeItemAggregated = $false
                    }
                    $ualCmd = Get-Command Get-UnifiedAuditLogs -ErrorAction SilentlyContinue
                    if ($ualCmd -and $ualCmd.Parameters.ContainsKey('SearchErrors')) {
                        $ualArgs.SearchErrors = $searchErrors
                    }
                    # Get-UnifiedAuditLogs returns ,ArrayList; @() around the call would keep it as one row.
                    $ualRaw = Get-UnifiedAuditLogs @ualArgs
                    $rows = @($ualRaw)
                } catch {
                    $searchErrors.Add([string]$_.Exception.Message)
                    $rows = @()
                }
                $licenseHit = @($searchErrors | Where-Object { Test-RemediationLogLicenseLimit $_ } | Select-Object -First 1)
                if ($licenseHit.Count -gt 0 -and @($rows).Count -eq 0) {
                    $shorter = Get-RemediationShorterLogDays -Days $ualDays
                    if ($shorter -gt 0) {
                        [void]$notes.Add(('{0}-day unified audit was refused by the license. Trying {1} days.' -f $ualDays, $shorter))
                        $ualDays = $shorter
                        $ualStart = $end.AddDays(-$shorter)
                        $ualShortened = $true
                        continue
                    }
                    [void]$notes.Add(('Unified audit failed: {0}' -f $licenseHit[0]))
                    $rows = @()
                } elseif (@($rows).Count -eq 0 -and $searchErrors.Count -gt 0) {
                    [void]$notes.Add(('Unified audit failed: {0}' -f $searchErrors[0]))
                }
                break
            }
            if ($ualShortened -and $licenseHit.Count -eq 0) {
                [void]$notes.Add(('Unified audit used {0} days because a longer window was refused by the license. Do not ask for that longer unified audit window again.' -f $ualDays))
            }
            try {
                if ($null -eq $rows) { $rows = @() }
                $priority = New-Object System.Collections.Generic.List[object]
                $other = New-Object System.Collections.Generic.List[object]
                foreach ($row in $rows) {
                    if (-not $row) { continue }
                    $item = ConvertTo-RemediationBecAuditRow $row
                    if (-not $item.Operation) { continue }
                    $rt = [string]$item.RecordType
                    $takeSharing = (& $wantLog 'sharing')
                    $takeFiles = (& $wantLog 'files')
                    $takeDlp = (& $wantLog 'dlp')
                    $takeUal = (& $wantLog 'ual')
                    if ($takeSharing -and $rt -match 'SharePointSharing|OneDrive' -and $sharing.Count -lt $sharingCap) {
                        [void]$sharing.Add($item)
                    } elseif ($takeFiles -and $rt -match 'SharePointFile' -and $files.Count -lt 15) {
                        [void]$files.Add($item)
                    }
                    if ($takeDlp -and $rt -match 'ComplianceDLP|DlpEndpoint' -and $dlp.Count -lt 15) {
                        $policy = ''
                        if ($row.AuditData) {
                            try {
                                $event = if ($row.AuditData -is [string]) { $row.AuditData | ConvertFrom-Json -ErrorAction Stop } else { $row.AuditData }
                                if ($event.PolicyDetails) {
                                    $pd = @($event.PolicyDetails)[0]
                                    if ($pd.PolicyName) { $policy = [string]$pd.PolicyName }
                                }
                            } catch { }
                        }
                        [void]$dlp.Add([pscustomobject]@{
                            Created   = $item.Created
                            Operation = $item.Operation
                            Policy    = (Get-RemediationBecAuditClip $policy 80)
                            Target    = $item.Target
                            User      = $item.User
                        })
                    }
                    if ($takeUal -and (Test-RemediationBecAuditPriority $row) -and $priority.Count -lt 25) {
                        [void]$priority.Add($item)
                    } elseif ($takeUal -and $other.Count -lt 20) {
                        [void]$other.Add($item)
                    }
                }
                foreach ($item in $priority) { if ($ual.Count -lt 40) { [void]$ual.Add($item) } }
                foreach ($item in $other) { if ($ual.Count -lt 40) { [void]$ual.Add($item) } }
                if ($rows.Count -gt 40) { [void]$notes.Add('Unified audit was capped at 40 rows, keeping DLP, sharing, download, and mailbox-change events ahead of routine mail activity. Exchange item aggregates were not included.') }
                if ($null -eq $only) {
                    [void]$notes.Add('Also pulled for this mailbox: DLP audit (Exchange, SharePoint, endpoint), SharePoint and OneDrive file and sharing operations, and threat or compliance alerts in unified audit.')
                }
            } catch {
                [void]$notes.Add(('Unified audit failed: {0}' -f $_.Exception.Message))
            }
        } elseif ($needAudit) {
            $auditNote = if ($auditReady) {
                'Unified audit was not pulled. The Exchange session is signed in, but the audit helper did not load. Run the threat check again.'
            } else {
                'Unified audit was not pulled. Search-UnifiedAuditLog is not in this Exchange session. Re-run Exchange Auth.'
            }
            [void]$notes.Add($auditNote)
        }
    }

    if (-not $GraphAuthenticated) {
        if ($needGraph) {
            [void]$notes.Add('Graph auth was not available, so MFA posture, sign-ins, security alerts, and new app registrations were not pulled.')
        }
    } else {
        $got = @{}
        if ($graphBatch) {
            try { $got = Complete-RemediationGraphGets -Batch $graphBatch }
            catch { [void]$notes.Add(('Graph log pull failed: {0}' -f $_.Exception.Message)) }
            $graphBatch = $null
        }
        if ($mfaJob) {
            try {
                $finished = Wait-Job -Job $mfaJob -Timeout 120
                if (-not $finished) {
                    Stop-Job -Job $mfaJob -ErrorAction SilentlyContinue
                    throw 'MFA posture timed out.'
                }
                $report = Receive-Job -Job $mfaJob -ErrorAction Stop
                $one = @($report.Users) | Select-Object -First 1
                $posture = [pscustomobject]@{
                    perUser          = if ($one) { [string]$one.PerUserMfaEnabled } else { '' }
                    securityDefaults = [string]$report.SecurityDefaultsEnabled
                    caRequiresMfa    = if ($one) { [string]$one.CARequiresMfa } else { [string]$report.CAPoliciesRequireMfa }
                    covered          = if ($one) { [string]$one.MfaCovered } else { '' }
                }
            } catch {
                [void]$notes.Add(('MFA posture failed: {0}' -f $_.Exception.Message))
            } finally {
                Remove-Job -Job $mfaJob -Force -ErrorAction SilentlyContinue
                $mfaJob = $null
            }
        } elseif (-not $graphTokenMissing -and $null -eq $only -and (Get-Command Get-MfaCoverageReport -ErrorAction SilentlyContinue)) {
            try {
                $report = Get-MfaCoverageReport -SelectedUsers @($user)
                $one = @($report.Users) | Select-Object -First 1
                $posture = [pscustomobject]@{
                    perUser          = if ($one) { [string]$one.PerUserMfaEnabled } else { '' }
                    securityDefaults = [string]$report.SecurityDefaultsEnabled
                    caRequiresMfa    = if ($one) { [string]$one.CARequiresMfa } else { [string]$report.CAPoliciesRequireMfa }
                    covered          = if ($one) { [string]$one.MfaCovered } else { '' }
                }
            } catch {
                [void]$notes.Add(('MFA posture failed: {0}' -f $_.Exception.Message))
            }
        }
        if ($got.ContainsKey('signins') -and -not $got.signins.Ok -and (Test-RemediationLogLicenseLimit ([string]$got.signins.Error))) {
            $signDays = $days
            $signTries = 0
            while ($signTries -lt 5) {
                $signTries++
                $shorter = Get-RemediationShorterLogDays -Days $signDays
                if ($shorter -lt 1) {
                    if ($signByUserId) {
                        $got.signins.Error = 'Sign-in logs for this window were refused by the tenant license. A shorter window will not change that. Do not ask for sign-ins again.'
                    }
                    break
                }
                [void]$notes.Add(('{0}-day sign-in log was refused by the license. Trying {1} days.' -f $signDays, $shorter))
                $signStart = $end.AddDays(-$shorter).ToString('yyyy-MM-ddTHH:mm:ssZ')
                $signEnd = $end.ToString('yyyy-MM-ddTHH:mm:ssZ')
                if ($signByUserId -and $signUserId) {
                    $signFilter = "createdDateTime ge $signStart and createdDateTime le $signEnd and userId eq '$signUserId'"
                } else {
                    $safeSignUser = $user.Replace("'", "''")
                    $signFilter = "userPrincipalName eq '$safeSignUser' and createdDateTime ge $signStart and createdDateTime le $signEnd"
                }
                $signFilter = "$signFilter and signInEventTypes/any(t: t eq 'interactiveUser' or t eq 'nonInteractiveUser')"
                $signUri = 'https://graph.microsoft.com/beta/auditLogs/signIns?$top=500&$filter=' + [Uri]::EscapeDataString($signFilter)
                try {
                    $signPageRetry = Invoke-GraphRestRequest -Uri $signUri -Headers @{ ConsistencyLevel = 'eventual' }
                    $got.signins = @{ Ok = $true; Error = ''; Json = $signPageRetry }
                    [void]$notes.Add(('Sign-in log used {0} days because a longer window was refused by the license. Do not ask for that longer sign-in window again.' -f $shorter))
                    break
                } catch {
                    $detail = $_.Exception.Message
                    if ($_.ErrorDetails -and $_.ErrorDetails.Message) { $detail = [string]$_.ErrorDetails.Message }
                    $got.signins = @{ Ok = $false; Error = $detail; Json = $null }
                    if (-not (Test-RemediationLogLicenseLimit $detail)) { break }
                    $signDays = $shorter
                }
            }
        }
        foreach ($pair in @(
            @{ Key = 'signins'; Label = 'Sign-in log' },
            @{ Key = 'apps'; Label = 'App registration search' },
            @{ Key = 'alerts'; Label = 'Security alerts' }
        )) {
            if (-not $got.ContainsKey($pair.Key)) { continue }
            $page = $got[$pair.Key]
            if (-not $page.Ok) {
                [void]$notes.Add(('{0} failed: {1}' -f $pair.Label, $page.Error))
            }
        }
        $signPage = $null
        if ($got.ContainsKey('signins') -and $got.signins.Ok) { $signPage = $got.signins.Json }
        $signRows = New-Object System.Collections.Generic.List[object]
        if ($signPage) {
            foreach ($row in @($signPage.value)) { if ($row) { [void]$signRows.Add($row) } }
            # Filtered sign-in queries often return empty pages that still carry a nextLink.
            $signNext = [string]$signPage.'@odata.nextLink'
            $signPages = 1
            while ($signNext -and $signRows.Count -lt 2000 -and $signPages -lt 20) {
                $signPages++
                try {
                    $more = Invoke-GraphRestRequest -Uri $signNext -Headers @{ ConsistencyLevel = 'eventual' }
                } catch {
                    [void]$notes.Add(('Sign-in log paging stopped: {0}' -f $_.Exception.Message))
                    break
                }
                foreach ($row in @($more.value)) { if ($row) { [void]$signRows.Add($row) } }
                $signNext = [string]$more.'@odata.nextLink'
            }
            if ($signRows.Count -eq 0) {
                $signHow = if ($signByUserId) { 'user id' } else { 'user principal name' }
                [void]$notes.Add(('Graph returned no sign-ins for this mailbox (by {0}) in this window after {1} page(s).' -f $signHow, $signPages))
            }
        }
        $interactive = New-Object System.Collections.Generic.List[object]
        $interactiveTotal = 0
        $quietGroups = [ordered]@{}
        foreach ($row in $signRows) {
            if (-not $row) { continue }
            $status = ''
            if ($row.status -and $row.status.errorCode -ne $null) { $status = [string]$row.status.errorCode }
            $where = ''
            if ($row.location) {
                $where = (@($row.location.city, $row.location.state, $row.location.countryOrRegion) | Where-Object { $_ }) -join ', '
            }
            $isQuiet = (@($row.signInEventTypes) -contains 'nonInteractiveUser') -or ($row.isInteractive -eq $false)
            if (-not $isQuiet) {
                $interactiveTotal++
                if ($interactive.Count -lt 25) {
                    [void]$interactive.Add([pscustomobject]@{
                        Created  = [string]$row.createdDateTime
                        Ip       = [string]$row.ipAddress
                        Location = [string]$where
                        App      = [string]$row.appDisplayName
                        Result   = $status
                        Kind     = 'interactive'
                    })
                }
                continue
            }
            # Rows arrive newest first, so the first row of a group is its last-seen time.
            $key = '{0}|{1}|{2}' -f $row.ipAddress, $row.appDisplayName, $status
            if ($quietGroups.Contains($key)) {
                $group = $quietGroups[$key]
                $group.Count++
                $group.FirstSeen = [string]$row.createdDateTime
            } else {
                $quietGroups[$key] = [pscustomobject]@{
                    Created   = [string]$row.createdDateTime
                    Ip        = [string]$row.ipAddress
                    Location  = [string]$where
                    App       = [string]$row.appDisplayName
                    Result    = $status
                    Kind      = 'non-interactive'
                    Count     = 1
                    FirstSeen = [string]$row.createdDateTime
                }
            }
        }
        $quietShown = @($quietGroups.Values | Select-Object -First 12)
        foreach ($group in $quietShown) { [void]$signIns.Add($group) }
        foreach ($row in $interactive) {
            if ($signIns.Count -ge 25) { break }
            [void]$signIns.Add($row)
        }
        if ($signRows.Count -gt 0 -and -not $signNext) {
            $oldest = $null
            foreach ($row in $signRows) {
                $stamp = [datetime]::MinValue
                if ($row.createdDateTime -is [datetime]) {
                    $stamp = $row.createdDateTime.ToUniversalTime()
                    if ($null -eq $oldest -or $stamp -lt $oldest) { $oldest = $stamp }
                } elseif ([datetime]::TryParse([string]$row.createdDateTime, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::AdjustToUniversal, [ref]$stamp)) {
                    if ($null -eq $oldest -or $stamp -lt $oldest) { $oldest = $stamp }
                }
            }
            if ($oldest -and ($oldest - $start).TotalDays -gt 2) {
                [void]$notes.Add(('Sign-ins only reach back to {0:yyyy-MM-dd} UTC although the window started {1:yyyy-MM-dd}. That is all the sign-in history the tenant keeps (about 7 days without a premium license). A longer window will not return more sign-ins. Do not ask for more sign-in history.' -f $oldest, $start))
            }
        }
        if ($signRows.Count -gt 0) {
            [void]$notes.Add(('Sign-ins pulled: {0} interactive, {1} non-interactive in {2} IP/app group(s). Non-interactive sign-ins are token refreshes from apps already signed in.' -f $interactiveTotal, ($signRows.Count - $interactiveTotal), $quietGroups.Count))
        }
        $appPage = $null
        if ($got.ContainsKey('apps') -and $got.apps.Ok) { $appPage = $got.apps.Json }
        foreach ($row in @($appPage.value)) {
            if ($apps.Count -ge 20) { break }
            if (-not $row) { continue }
            [void]$apps.Add([pscustomobject]@{
                DisplayName = [string]$row.displayName
                Created     = [string]$row.createdDateTime
            })
        }
        if ($got.ContainsKey('alerts') -and $got.alerts.Ok) {
            $want = $user.ToLowerInvariant()
            $alertRows = @()
            if ($got.alerts.Json -and $got.alerts.Json.value) { $alertRows = @($got.alerts.Json.value) }
            foreach ($row in $alertRows) {
                if ($alerts.Count -ge 15) { break }
                if (-not $row) { continue }
                $names = New-Object System.Collections.Generic.List[string]
                foreach ($state in @($row.userStates)) {
                    $upn = [string]$state.userPrincipalName
                    if ($upn) { [void]$names.Add($upn.ToLowerInvariant()) }
                }
                $title = [string]$row.title
                $matched = $names -contains $want -or ($title -and $title.ToLowerInvariant().Contains($want))
                if (-not $matched) { continue }
                [void]$alerts.Add([pscustomobject]@{
                    Created  = [string]$row.createdDateTime
                    Title    = (Get-RemediationBecAuditClip $title 120)
                    Severity = [string]$row.severity
                    Category = [string]$row.category
                    Status   = [string]$row.status
                })
            }
            [void]$notes.Add('Security alerts are the newest 50 in the window, then kept only when they name this mailbox.')
        }
    }
    if ($null -eq $only) {
        [void]$notes.Add('Not pulled unless a later pass needs them: folder permissions, junk trusted senders, transport rules, org auto-forward, rights on other mailboxes, Teams activity, and Exchange item aggregates.')
    }
    $payload = @{
        Days         = $days
        StartUtc     = $start.ToString('yyyy-MM-ddTHH:mm:ssZ')
        EndUtc       = $end.ToString('yyyy-MM-ddTHH:mm:ssZ')
        MfaPosture   = $posture
        MessageTrace = $trace
        ReportedMessages = @($reported.ToArray())
        Ual          = $ual
        Dlp          = $dlp
        Files        = $files
        Sharing      = $sharing
        Alerts       = $alerts
        SignIns      = $signIns
        Apps         = $apps
        Blasts       = $blasts
        PullNotes    = $notes
    } | ConvertTo-Json -Compress -Depth 6
    return "REMEDIATE_BEC:$payload"
}

. (Join-Path $PSScriptRoot 'RemediationExtended.ps1')

Export-ModuleMember -Function Get-GraphContainmentCapabilities, Get-RemediationGraphUser
Export-ModuleMember -Function Test-RemediationGraphAlreadyGone, Test-RemediationIntuneUnavailable
Export-ModuleMember -Function Test-RemediationLogLicenseLimit, Get-RemediationShorterLogDays
Export-ModuleMember -Function Convert-RemediationDateToText, Get-RemediationObjectDateText
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
