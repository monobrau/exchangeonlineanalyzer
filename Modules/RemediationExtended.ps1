# Dot-sourced by Remediation.psm1. Extra BEC containment actions (OAuth, Intune, roles, EXO extras).

function Get-RemediationSpDisplayName {
    param([string]$Id, [hashtable]$Cache)
    if ([string]::IsNullOrWhiteSpace($Id)) { return '' }
    if ($Cache.ContainsKey($Id)) { return [string]$Cache[$Id] }
    try {
        $enc = [Uri]::EscapeDataString($Id)
        $sp = Invoke-GraphRestRequest -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$enc`?`$select=displayName,appId"
        $name = if ($sp.displayName) { [string]$sp.displayName } else { $Id }
        $Cache[$Id] = $name
        return $name
    } catch {
        $Cache[$Id] = $Id
        return $Id
    }
}

function Get-RemediationOAuthGrants {
    param([Parameter(Mandatory = $true)][string[]]$UserPrincipalNames)
    $cache = @{}
    foreach ($upn in $UserPrincipalNames) {
        try {
            $user = Get-RemediationGraphUser -UserPrincipalName $upn
            $id = [string]$user.id
            $enc = [Uri]::EscapeDataString($id)
            $grants = @()
            try {
                $grants = @(Invoke-GraphRestPaged -Uri "https://graph.microsoft.com/v1.0/users/$enc/oauth2PermissionGrants")
            } catch {
                $filter = [Uri]::EscapeDataString("principalId eq '$id'")
                $grants = @(Invoke-GraphRestPaged -Uri "https://graph.microsoft.com/v1.0/oauth2PermissionGrants?`$filter=$filter")
            }
            if ($grants.Count -eq 0) {
                [pscustomobject]@{
                    UserPrincipalName = [string]$user.userPrincipalName
                    Id                = ''
                    App               = '(none)'
                    Scope             = ''
                    ConsentType       = ''
                    Error             = ''
                }
                continue
            }
            foreach ($g in $grants) {
                $clientId = [string]$g.clientId
                [pscustomobject]@{
                    UserPrincipalName = [string]$user.userPrincipalName
                    Id                = [string]$g.id
                    App               = (Get-RemediationSpDisplayName -Id $clientId -Cache $cache)
                    Scope             = [string]$g.scope
                    ConsentType       = [string]$g.consentType
                    Error             = ''
                }
            }
        } catch {
            [pscustomobject]@{
                UserPrincipalName = $upn
                Id                = ''
                App               = ''
                Scope             = ''
                ConsentType       = ''
                Error             = (Get-RemediationGraphErrorText $_)
            }
        }
    }
}

function Remove-RemediationOAuthGrants {
    param([Parameter(Mandatory = $true)][object[]]$Items)
    foreach ($item in $Items) {
        $id = if ($item.Id) { [string]$item.Id } else { [string]$item }
        if ([string]::IsNullOrWhiteSpace($id)) { continue }
        try {
            $enc = [Uri]::EscapeDataString($id)
            Invoke-RemediationGraphDelete -Uri "https://graph.microsoft.com/v1.0/oauth2PermissionGrants/$enc"
            [pscustomobject]@{ Id = $id; Success = $true; Error = '' }
        } catch {
            [pscustomobject]@{ Id = $id; Success = $false; Error = (Get-RemediationGraphErrorText $_) }
        }
    }
}

function Invoke-RemediationReregisterMfa {
    param([Parameter(Mandatory = $true)][string[]]$UserPrincipalNames)
    foreach ($upn in $UserPrincipalNames) {
        try {
            $methods = @(Get-RemediationAuthMethods -UserPrincipalNames @($upn) | Where-Object { $_.CanDelete -and $_.Id })
            $deleted = 0
            $failed = 0
            $errors = [System.Collections.Generic.List[string]]::new()
            if ($methods.Count -gt 0) {
                $rows = @(Remove-RemediationAuthMethods -Items $methods)
                foreach ($row in $rows) {
                    if ($row.Success) { $deleted++ } else {
                        $failed++
                        if ($row.Error) { [void]$errors.Add([string]$row.Error) }
                    }
                }
            }
            $revoked = $false
            try {
                Invoke-RemediationRevokeSessions -UserPrincipalName $upn | Out-Null
                $revoked = $true
            } catch {
                [void]$errors.Add((Get-RemediationGraphErrorText $_))
            }
            [pscustomobject]@{
                UserPrincipalName = $upn
                DeletedCount      = $deleted
                FailCount         = $failed
                SessionsRevoked   = $revoked
                Success           = ($failed -eq 0)
                Error             = ($errors -join '; ')
            }
        } catch {
            [pscustomobject]@{
                UserPrincipalName = $upn
                DeletedCount      = 0
                FailCount         = 1
                SessionsRevoked   = $false
                Success           = $false
                Error             = (Get-RemediationGraphErrorText $_)
            }
        }
    }
}

function Get-RemediationIntuneDevices {
    param([Parameter(Mandatory = $true)][string[]]$UserPrincipalNames)
    foreach ($upn in $UserPrincipalNames) {
        try {
            $user = Get-RemediationGraphUser -UserPrincipalName $upn
            $id = [string]$user.id
            $enc = [Uri]::EscapeDataString($id)
            $rows = @()
            try {
                $rows = @(Invoke-GraphRestPaged -Uri "https://graph.microsoft.com/v1.0/users/$enc/managedDevices")
            } catch {
                $filter = [Uri]::EscapeDataString("userPrincipalName eq '$($user.userPrincipalName)'")
                $rows = @(Invoke-GraphRestPaged -Uri "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$filter=$filter")
            }
            if ($rows.Count -eq 0) {
                [pscustomobject]@{
                    UserPrincipalName = [string]$user.userPrincipalName
                    Id                = ''
                    DeviceName        = '(none)'
                    Os                = ''
                    Management        = ''
                    Compliance        = ''
                    LastSync          = ''
                    Error             = ''
                }
                continue
            }
            foreach ($d in $rows) {
                [pscustomobject]@{
                    UserPrincipalName = [string]$user.userPrincipalName
                    Id                = [string]$d.id
                    DeviceName        = [string]$d.deviceName
                    Os                = ("$($d.operatingSystem) $($d.osVersion)").Trim()
                    Management        = [string]$d.managementAgent
                    Compliance        = [string]$d.complianceState
                    LastSync          = if ($d.lastSyncDateTime) { [string]$d.lastSyncDateTime } else { '' }
                    Error             = ''
                }
            }
        } catch {
            [pscustomobject]@{
                UserPrincipalName = $upn
                Id                = ''
                DeviceName        = ''
                Os                = ''
                Management        = ''
                Compliance        = ''
                LastSync          = ''
                Error             = (Get-RemediationGraphErrorText $_)
            }
        }
    }
}

function Invoke-RemediationIntuneAction {
    param(
        [Parameter(Mandatory = $true)][object[]]$Items,
        [ValidateSet('wipe', 'retire')][string]$Action
    )
    foreach ($item in $Items) {
        $id = if ($item.Id) { [string]$item.Id } else { [string]$item }
        if ([string]::IsNullOrWhiteSpace($id)) { continue }
        try {
            $enc = [Uri]::EscapeDataString($id)
            $body = if ($Action -eq 'wipe') { '{"keepEnrollmentData":false,"keepUserData":false}' } else { '{}' }
            Invoke-GraphRestRequest -Uri "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices/$enc/$Action" -Method POST -Body $body | Out-Null
            [pscustomobject]@{ Id = $id; Action = $Action; Success = $true; Error = '' }
        } catch {
            [pscustomobject]@{ Id = $id; Action = $Action; Success = $false; Error = (Get-RemediationGraphErrorText $_) }
        }
    }
}

function Get-RemediationDirectoryRoles {
    param([Parameter(Mandatory = $true)][string[]]$UserPrincipalNames)
    foreach ($upn in $UserPrincipalNames) {
        try {
            $user = Get-RemediationGraphUser -UserPrincipalName $upn
            $id = [string]$user.id
            $enc = [Uri]::EscapeDataString($id)
            $any = $false
            $memberOf = @()
            try {
                $memberOf = @(Invoke-GraphRestPaged -Uri "https://graph.microsoft.com/v1.0/users/$enc/memberOf?`$select=id,displayName,mail,@odata.type,securityEnabled,mailEnabled,isAssignableToRole")
            } catch {}
            foreach ($obj in $memberOf) {
                $odata = [string]$obj.'@odata.type'
                if ($odata -match 'directoryRole') {
                    $any = $true
                    [pscustomobject]@{
                        UserPrincipalName = [string]$user.userPrincipalName
                        Kind              = 'DirectoryRole'
                        Id                = [string]$obj.id
                        Name              = [string]$obj.displayName
                        Details           = 'directoryRole member'
                        CanRemove         = $true
                        Error             = ''
                    }
                } elseif ($odata -match 'group') {
                    $any = $true
                    $flags = @()
                    if ($obj.isAssignableToRole) { $flags += 'role-assignable' }
                    if ($obj.securityEnabled) { $flags += 'security' }
                    if ($obj.mailEnabled) { $flags += 'mail' }
                    [pscustomobject]@{
                        UserPrincipalName = [string]$user.userPrincipalName
                        Kind              = 'Group'
                        Id                = [string]$obj.id
                        Name              = [string]$obj.displayName
                        Details           = ($flags -join ', ')
                        CanRemove         = $true
                        Error             = ''
                    }
                }
            }
            try {
                $filter = [Uri]::EscapeDataString("principalId eq '$id'")
                $assigns = @(Invoke-GraphRestPaged -Uri "https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignments?`$filter=$filter")
                foreach ($a in $assigns) {
                    $roleName = [string]$a.roleDefinitionId
                    try {
                        $rdEnc = [Uri]::EscapeDataString([string]$a.roleDefinitionId)
                        $rd = Invoke-GraphRestRequest -Uri "https://graph.microsoft.com/v1.0/roleManagement/directory/roleDefinitions/$rdEnc`?`$select=displayName"
                        if ($rd.displayName) { $roleName = [string]$rd.displayName }
                    } catch {}
                    $any = $true
                    [pscustomobject]@{
                        UserPrincipalName = [string]$user.userPrincipalName
                        Kind              = 'RoleAssignment'
                        Id                = [string]$a.id
                        Name              = $roleName
                        Details           = "roleDefinitionId=$($a.roleDefinitionId)"
                        CanRemove         = $true
                        Error             = ''
                    }
                }
            } catch {}
            if (Get-Command Get-ManagementRoleAssignment -ErrorAction SilentlyContinue) {
                try {
                    $rbac = @(Get-ManagementRoleAssignment -RoleAssignee $upn -ErrorAction SilentlyContinue)
                    foreach ($r in $rbac) {
                        $any = $true
                        [pscustomobject]@{
                            UserPrincipalName = [string]$user.userPrincipalName
                            Kind              = 'ExchangeRbac'
                            Id                = [string]$r.Identity
                            Name              = [string]$r.Role
                            Details           = "assignee=$($r.RoleAssigneeName); method=$($r.AssignmentMethod)"
                            CanRemove         = $true
                            Error             = ''
                        }
                    }
                } catch {}
            }
            if (-not $any) {
                [pscustomobject]@{
                    UserPrincipalName = [string]$user.userPrincipalName
                    Kind              = ''
                    Id                = ''
                    Name              = '(none)'
                    Details           = ''
                    CanRemove         = $false
                    Error             = ''
                }
            }
        } catch {
            [pscustomobject]@{
                UserPrincipalName = $upn
                Kind              = ''
                Id                = ''
                Name              = ''
                Details           = ''
                CanRemove         = $false
                Error             = (Get-RemediationGraphErrorText $_)
            }
        }
    }
}

function Remove-RemediationDirectoryRoles {
    param([Parameter(Mandatory = $true)][object[]]$Items)
    foreach ($item in $Items) {
        $kind = [string]$item.Kind
        $id = [string]$item.Id
        $upn = [string]$item.UserPrincipalName
        if ([string]::IsNullOrWhiteSpace($id)) { continue }
        try {
            if ($kind -eq 'RoleAssignment') {
                $enc = [Uri]::EscapeDataString($id)
                Invoke-RemediationGraphDelete -Uri "https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignments/$enc"
            } elseif ($kind -eq 'DirectoryRole') {
                $user = Get-RemediationGraphUser -UserPrincipalName $upn
                $roleEnc = [Uri]::EscapeDataString($id)
                $userEnc = [Uri]::EscapeDataString([string]$user.id)
                Invoke-RemediationGraphDelete -Uri "https://graph.microsoft.com/v1.0/directoryRoles/$roleEnc/members/$userEnc/`$ref"
            } elseif ($kind -eq 'Group') {
                $user = Get-RemediationGraphUser -UserPrincipalName $upn
                $gEnc = [Uri]::EscapeDataString($id)
                $userEnc = [Uri]::EscapeDataString([string]$user.id)
                Invoke-RemediationGraphDelete -Uri "https://graph.microsoft.com/v1.0/groups/$gEnc/members/$userEnc/`$ref"
            } elseif ($kind -eq 'ExchangeRbac') {
                Remove-ManagementRoleAssignment -Identity $id -Confirm:$false -ErrorAction Stop
            } else {
                throw "Unknown role kind '$kind'"
            }
            [pscustomobject]@{ Id = $id; Kind = $kind; Success = $true; Error = '' }
        } catch {
            [pscustomobject]@{ Id = $id; Kind = $kind; Success = $false; Error = (Get-RemediationGraphErrorText $_) }
        }
    }
}

function Get-RemediationAppCredentials {
    $rows = [System.Collections.Generic.List[object]]::new()
    $apps = @(Invoke-GraphRestPaged -Uri 'https://graph.microsoft.com/v1.0/applications?$select=id,appId,displayName,passwordCredentials,keyCredentials')
    foreach ($app in $apps) {
        $appId = [string]$app.id
        $name = [string]$app.displayName
        foreach ($p in @($app.passwordCredentials)) {
            [void]$rows.Add([pscustomobject]@{
                Kind        = 'Secret'
                AppId       = $appId
                AppName     = $name
                KeyId       = [string]$p.keyId
                DisplayName = [string]$p.displayName
                End         = if ($p.endDateTime) { [string]$p.endDateTime } else { '' }
                OwnerId     = ''
                Error       = ''
            })
        }
        foreach ($k in @($app.keyCredentials)) {
            [void]$rows.Add([pscustomobject]@{
                Kind        = 'Certificate'
                AppId       = $appId
                AppName     = $name
                KeyId       = [string]$k.keyId
                DisplayName = [string]$k.displayName
                End         = if ($k.endDateTime) { [string]$k.endDateTime } else { '' }
                OwnerId     = ''
                Error       = ''
            })
        }
        try {
            $enc = [Uri]::EscapeDataString($appId)
            $owners = @(Invoke-GraphRestPaged -Uri "https://graph.microsoft.com/v1.0/applications/$enc/owners?`$select=id,displayName,userPrincipalName")
            foreach ($o in $owners) {
                $ownerLabel = if ($o.userPrincipalName) { [string]$o.userPrincipalName } elseif ($o.displayName) { [string]$o.displayName } else { [string]$o.id }
                [void]$rows.Add([pscustomobject]@{
                    Kind        = 'Owner'
                    AppId       = $appId
                    AppName     = $name
                    KeyId       = ''
                    DisplayName = $ownerLabel
                    End         = ''
                    OwnerId     = [string]$o.id
                    Error       = ''
                })
            }
        } catch {}
    }
    if ($rows.Count -eq 0) {
        [void]$rows.Add([pscustomobject]@{
            Kind = ''; AppId = ''; AppName = '(none)'; KeyId = ''; DisplayName = ''; End = ''; OwnerId = ''; Error = ''
        })
    }
    return @($rows)
}

function Remove-RemediationAppCredentials {
    param([Parameter(Mandatory = $true)][object[]]$Items)
    foreach ($item in $Items) {
        $kind = [string]$item.Kind
        $appId = [string]$item.AppId
        try {
            if ([string]::IsNullOrWhiteSpace($appId)) { throw 'Missing AppId' }
            $enc = [Uri]::EscapeDataString($appId)
            if ($kind -eq 'Secret') {
                $kid = [string]$item.KeyId
                $body = @{ keyId = $kid } | ConvertTo-Json -Compress
                Invoke-GraphRestRequest -Uri "https://graph.microsoft.com/v1.0/applications/$enc/removePassword" -Method POST -Body $body | Out-Null
            } elseif ($kind -eq 'Certificate') {
                throw 'Certificate removal needs proof-of-possession. Remove the cert in Entra or replace the app.'
            } elseif ($kind -eq 'Owner') {
                $oid = [Uri]::EscapeDataString([string]$item.OwnerId)
                Invoke-RemediationGraphDelete -Uri "https://graph.microsoft.com/v1.0/applications/$enc/owners/$oid/`$ref"
            } else {
                throw "Unknown credential kind '$kind'"
            }
            [pscustomobject]@{ Kind = $kind; AppId = $appId; KeyId = [string]$item.KeyId; Success = $true; Error = '' }
        } catch {
            [pscustomobject]@{ Kind = $kind; AppId = $appId; KeyId = [string]$item.KeyId; Success = $false; Error = (Get-RemediationGraphErrorText $_) }
        }
    }
}

function Get-RemediationMobileDevices {
    param([Parameter(Mandatory = $true)][string[]]$UserPrincipalNames)
    foreach ($upn in $UserPrincipalNames) {
        try {
            $devs = @(Get-MobileDevice -Mailbox $upn -ErrorAction Stop)
            if ($devs.Count -eq 0) {
                [pscustomobject]@{
                    UserPrincipalName = $upn
                    Identity          = ''
                    DeviceId          = ''
                    DeviceType        = ''
                    FriendlyName      = '(none)'
                    FirstSync         = ''
                    Error             = ''
                }
                continue
            }
            foreach ($d in $devs) {
                [pscustomobject]@{
                    UserPrincipalName = $upn
                    Identity          = [string]$d.Identity
                    DeviceId          = [string]$d.DeviceId
                    DeviceType        = [string]$d.DeviceType
                    FriendlyName      = [string]$d.FriendlyName
                    FirstSync         = if ($d.FirstSyncTime) { [string]$d.FirstSyncTime } else { '' }
                    Error             = ''
                }
            }
        } catch {
            [pscustomobject]@{
                UserPrincipalName = $upn
                Identity          = ''
                DeviceId          = ''
                DeviceType        = ''
                FriendlyName      = ''
                FirstSync         = ''
                Error             = $_.Exception.Message
            }
        }
    }
}

function Remove-RemediationMobileDevices {
    param([Parameter(Mandatory = $true)][object[]]$Items)
    foreach ($item in $Items) {
        $identity = if ($item.Identity) { [string]$item.Identity } else { [string]$item }
        if ([string]::IsNullOrWhiteSpace($identity)) { continue }
        try {
            Remove-MobileDevice -Identity $identity -Confirm:$false -ErrorAction Stop
            [pscustomobject]@{ Identity = $identity; Success = $true; Error = '' }
        } catch {
            [pscustomobject]@{ Identity = $identity; Success = $false; Error = $_.Exception.Message }
        }
    }
}

function Get-RemediationFolderPermissions {
    param([Parameter(Mandatory = $true)][string[]]$UserPrincipalNames)
    $folders = @('Inbox', 'Calendar', 'SentItems', 'Contacts', 'Drafts')
    foreach ($upn in $UserPrincipalNames) {
        $any = $false
        foreach ($folder in $folders) {
            $identity = "${upn}:\$folder"
            try {
                $perms = @(Get-MailboxFolderPermission -Identity $identity -ErrorAction Stop)
                foreach ($p in $perms) {
                    $user = [string]$p.User
                    if ([string]::IsNullOrWhiteSpace($user)) { continue }
                    $rights = if ($p.AccessRights) { ($p.AccessRights | ForEach-Object { [string]$_ }) -join ', ' } else { '' }
                    if ($user -match '^(Default|Anonymous)$' -and $rights -match '^(None)?$') { continue }
                    $any = $true
                    [pscustomobject]@{
                        UserPrincipalName = $upn
                        Folder            = $folder
                        User              = $user
                        AccessRights      = $rights
                        Error             = ''
                    }
                }
            } catch {
                $any = $true
                [pscustomobject]@{
                    UserPrincipalName = $upn
                    Folder            = $folder
                    User              = ''
                    AccessRights      = ''
                    Error             = $_.Exception.Message
                }
            }
        }
        if (-not $any) {
            [pscustomobject]@{
                UserPrincipalName = $upn
                Folder            = ''
                User              = '(none)'
                AccessRights      = ''
                Error             = ''
            }
        }
    }
}

function Remove-RemediationFolderPermissions {
    param([Parameter(Mandatory = $true)][object[]]$Items)
    foreach ($item in $Items) {
        $upn = [string]$item.UserPrincipalName
        $folder = [string]$item.Folder
        $user = [string]$item.User
        $identity = "${upn}:\$folder"
        try {
            if ($user -match '^(Default|Anonymous)$') {
                Set-MailboxFolderPermission -Identity $identity -User $user -AccessRights None -ErrorAction Stop
            } else {
                Remove-MailboxFolderPermission -Identity $identity -User $user -Confirm:$false -ErrorAction Stop
            }
            [pscustomobject]@{ UserPrincipalName = $upn; Folder = $folder; User = $user; Success = $true; Error = '' }
        } catch {
            [pscustomobject]@{ UserPrincipalName = $upn; Folder = $folder; User = $user; Success = $false; Error = $_.Exception.Message }
        }
    }
}

function Get-RemediationAutoReply {
    param([Parameter(Mandatory = $true)][string[]]$UserPrincipalNames)
    foreach ($upn in $UserPrincipalNames) {
        try {
            $cfg = Get-MailboxAutoReplyConfiguration -Identity $upn -ErrorAction Stop
            [pscustomobject]@{
                UserPrincipalName = $upn
                AutoReplyState    = [string]$cfg.AutoReplyState
                ExternalAudience  = [string]$cfg.ExternalAudience
                StartTime         = if ($cfg.StartTime) { [string]$cfg.StartTime } else { '' }
                EndTime           = if ($cfg.EndTime) { [string]$cfg.EndTime } else { '' }
                InternalMessage   = [string]$cfg.InternalMessage
                ExternalMessage   = [string]$cfg.ExternalMessage
                Error             = ''
            }
        } catch {
            [pscustomobject]@{
                UserPrincipalName = $upn
                AutoReplyState    = ''
                ExternalAudience  = ''
                StartTime         = ''
                EndTime           = ''
                InternalMessage   = ''
                ExternalMessage   = ''
                Error             = $_.Exception.Message
            }
        }
    }
}

function Disable-RemediationAutoReply {
    param([Parameter(Mandatory = $true)][string[]]$UserPrincipalNames)
    foreach ($upn in $UserPrincipalNames) {
        try {
            Set-MailboxAutoReplyConfiguration -Identity $upn -AutoReplyState Disabled -ErrorAction Stop
            [pscustomobject]@{ UserPrincipalName = $upn; Success = $true; Error = '' }
        } catch {
            [pscustomobject]@{ UserPrincipalName = $upn; Success = $false; Error = $_.Exception.Message }
        }
    }
}

function Get-RemediationOrgForward {
    $rows = [System.Collections.Generic.List[object]]::new()
    try {
        foreach ($d in @(Get-RemoteDomain -ErrorAction Stop)) {
            [void]$rows.Add([pscustomobject]@{
                Kind              = 'RemoteDomain'
                Identity          = [string]$d.Identity
                Name              = [string]$d.DomainName
                AutoForward       = [string]$d.AutoForwardEnabled
                AutoForwardingMode = ''
                Error             = ''
            })
        }
    } catch {
        [void]$rows.Add([pscustomobject]@{
            Kind = 'RemoteDomain'; Identity = ''; Name = ''; AutoForward = ''; AutoForwardingMode = ''; Error = $_.Exception.Message
        })
    }
    try {
        foreach ($p in @(Get-HostedOutboundSpamFilterPolicy -ErrorAction Stop)) {
            [void]$rows.Add([pscustomobject]@{
                Kind              = 'OutboundSpam'
                Identity          = [string]$p.Identity
                Name              = [string]$p.Name
                AutoForward       = ''
                AutoForwardingMode = [string]$p.AutoForwardingMode
                Error             = ''
            })
        }
    } catch {
        [void]$rows.Add([pscustomobject]@{
            Kind = 'OutboundSpam'; Identity = ''; Name = ''; AutoForward = ''; AutoForwardingMode = ''; Error = $_.Exception.Message
        })
    }
    return @($rows)
}

function Set-RemediationOrgForward {
    param([Parameter(Mandatory = $true)][object[]]$Items)
    foreach ($item in $Items) {
        $kind = [string]$item.Kind
        $identity = [string]$item.Identity
        try {
            if ($kind -eq 'RemoteDomain') {
                Set-RemoteDomain -Identity $identity -AutoForwardEnabled $false -ErrorAction Stop
            } elseif ($kind -eq 'OutboundSpam') {
                Set-HostedOutboundSpamFilterPolicy -Identity $identity -AutoForwardingMode Off -ErrorAction Stop
            } else {
                throw "Unknown org-forward kind '$kind'"
            }
            [pscustomobject]@{ Kind = $kind; Identity = $identity; Success = $true; Error = '' }
        } catch {
            [pscustomobject]@{ Kind = $kind; Identity = $identity; Success = $false; Error = $_.Exception.Message }
        }
    }
}

function Get-RemediationJunk {
    param([Parameter(Mandatory = $true)][string[]]$UserPrincipalNames)
    foreach ($upn in $UserPrincipalNames) {
        try {
            $cfg = Get-MailboxJunkEmailConfiguration -Identity $upn -ErrorAction Stop
            $any = $false
            foreach ($addr in @($cfg.TrustedSendersAndDomains)) {
                if (-not $addr) { continue }
                $any = $true
                [pscustomobject]@{
                    UserPrincipalName = $upn
                    List              = 'TrustedSenders'
                    Address           = [string]$addr
                    ContactsTrusted   = [bool]$cfg.ContactsTrusted
                    Error             = ''
                }
            }
            foreach ($addr in @($cfg.TrustedRecipientsAndDomains)) {
                if (-not $addr) { continue }
                $any = $true
                [pscustomobject]@{
                    UserPrincipalName = $upn
                    List              = 'TrustedRecipients'
                    Address           = [string]$addr
                    ContactsTrusted   = [bool]$cfg.ContactsTrusted
                    Error             = ''
                }
            }
            if (-not $any) {
                [pscustomobject]@{
                    UserPrincipalName = $upn
                    List              = ''
                    Address           = '(none)'
                    ContactsTrusted   = [bool]$cfg.ContactsTrusted
                    Error             = ''
                }
            }
        } catch {
            [pscustomobject]@{
                UserPrincipalName = $upn
                List              = ''
                Address           = ''
                ContactsTrusted   = $null
                Error             = $_.Exception.Message
            }
        }
    }
}

function Remove-RemediationJunk {
    param([Parameter(Mandatory = $true)][object[]]$Items)
    $byUser = @{}
    foreach ($item in $Items) {
        $upn = [string]$item.UserPrincipalName
        if (-not $byUser.ContainsKey($upn)) { $byUser[$upn] = [System.Collections.Generic.List[object]]::new() }
        [void]$byUser[$upn].Add($item)
    }
    foreach ($upn in $byUser.Keys) {
        foreach ($item in $byUser[$upn]) {
            $list = [string]$item.List
            $addr = [string]$item.Address
            try {
                if ($list -eq 'TrustedSenders') {
                    Set-MailboxJunkEmailConfiguration -Identity $upn -TrustedSendersAndDomains @{ Remove = $addr } -ErrorAction Stop
                } elseif ($list -eq 'TrustedRecipients') {
                    Set-MailboxJunkEmailConfiguration -Identity $upn -TrustedRecipientsAndDomains @{ Remove = $addr } -ErrorAction Stop
                } else {
                    throw "Unknown junk list '$list'"
                }
                [pscustomobject]@{ UserPrincipalName = $upn; List = $list; Address = $addr; Success = $true; Error = '' }
            } catch {
                [pscustomobject]@{ UserPrincipalName = $upn; List = $list; Address = $addr; Success = $false; Error = $_.Exception.Message }
            }
        }
    }
}

function Get-RemediationJournalRules {
    try {
        $rules = @(Get-JournalRule -ErrorAction Stop)
        if ($rules.Count -eq 0) {
            return @([pscustomobject]@{ Identity = ''; Name = '(none)'; Recipient = ''; JournalEmailAddress = ''; Enabled = $null; Error = '' })
        }
        foreach ($r in $rules) {
            [pscustomobject]@{
                Identity            = [string]$r.Identity
                Name                = [string]$r.Name
                Recipient           = [string]$r.Recipient
                JournalEmailAddress = [string]$r.JournalEmailAddress
                Enabled             = [bool]$r.Enabled
                Error               = ''
            }
        }
    } catch {
        [pscustomobject]@{ Identity = ''; Name = ''; Recipient = ''; JournalEmailAddress = ''; Enabled = $null; Error = $_.Exception.Message }
    }
}

function Remove-RemediationJournalRules {
    param([Parameter(Mandatory = $true)][object[]]$Items)
    foreach ($item in $Items) {
        $identity = if ($item.Identity) { [string]$item.Identity } elseif ($item.Name) { [string]$item.Name } else { [string]$item }
        if ([string]::IsNullOrWhiteSpace($identity)) { continue }
        try {
            Remove-JournalRule -Identity $identity -Confirm:$false -ErrorAction Stop
            [pscustomobject]@{ Identity = $identity; Success = $true; Error = '' }
        } catch {
            [pscustomobject]@{ Identity = $identity; Success = $false; Error = $_.Exception.Message }
        }
    }
}

function Get-RemediationMailboxHold {
    param([Parameter(Mandatory = $true)][string[]]$UserPrincipalNames)
    foreach ($upn in $UserPrincipalNames) {
        try {
            $mbx = Get-Mailbox -Identity $upn -ErrorAction Stop
            [pscustomobject]@{
                UserPrincipalName      = [string]$mbx.UserPrincipalName
                LitigationHoldEnabled  = [bool]$mbx.LitigationHoldEnabled
                RetainDeletedItemsFor  = [string]$mbx.RetainDeletedItemsFor
                AuditEnabled           = [bool]$mbx.AuditEnabled
                Error                  = ''
            }
        } catch {
            [pscustomobject]@{
                UserPrincipalName     = $upn
                LitigationHoldEnabled = $null
                RetainDeletedItemsFor = ''
                AuditEnabled          = $null
                Error                 = $_.Exception.Message
            }
        }
    }
}

function Set-RemediationMailboxHold {
    param([Parameter(Mandatory = $true)][string[]]$UserPrincipalNames)
    foreach ($upn in $UserPrincipalNames) {
        try {
            Set-Mailbox -Identity $upn -LitigationHoldEnabled $true -RetainDeletedItemsFor 30 -AuditEnabled $true -ErrorAction Stop
            [pscustomobject]@{ UserPrincipalName = $upn; Success = $true; Error = '' }
        } catch {
            [pscustomobject]@{ UserPrincipalName = $upn; Success = $false; Error = $_.Exception.Message }
        }
    }
}

function Get-RemediationMailboxElsewhere {
    param([Parameter(Mandatory = $true)][string[]]$UserPrincipalNames)
    foreach ($upn in $UserPrincipalNames) {
        $any = $false
        try {
            foreach ($p in @(Get-RecipientPermission -Trustee $upn -ErrorAction SilentlyContinue)) {
                $any = $true
                [pscustomobject]@{
                    UserPrincipalName = $upn
                    Mailbox           = [string]$p.Identity
                    Right             = 'SendAs'
                    Trustee           = [string]$p.Trustee
                    Error             = ''
                }
            }
        } catch {}
        try {
            $self = Get-Mailbox -Identity $upn -ErrorAction Stop
            $dn = [string]$self.DistinguishedName
            if ($dn) {
                $escaped = $dn.Replace("'", "''")
                foreach ($mbx in @(Get-Mailbox -Filter "GrantSendOnBehalfTo -eq '$escaped'" -ResultSize Unlimited -ErrorAction SilentlyContinue)) {
                    if ([string]$mbx.UserPrincipalName -eq $upn) { continue }
                    $any = $true
                    [pscustomobject]@{
                        UserPrincipalName = $upn
                        Mailbox           = [string]$mbx.UserPrincipalName
                        Right             = 'SendOnBehalf'
                        Trustee           = $upn
                        Error             = ''
                    }
                }
            }
        } catch {}
        try {
            $mailboxes = @(Get-Mailbox -RecipientTypeDetails UserMailbox, SharedMailbox, RoomMailbox, EquipmentMailbox -ResultSize Unlimited -ErrorAction Stop)
            foreach ($mbx in $mailboxes) {
                if ([string]$mbx.UserPrincipalName -eq $upn) { continue }
                $perms = @(Get-MailboxPermission -Identity $mbx.Identity -User $upn -ErrorAction SilentlyContinue |
                    Where-Object { -not $_.IsInherited -and $_.AccessRights -and ($_.AccessRights -contains 'FullAccess') })
                foreach ($p in $perms) {
                    $any = $true
                    [pscustomobject]@{
                        UserPrincipalName = $upn
                        Mailbox           = [string]$mbx.UserPrincipalName
                        Right             = 'FullAccess'
                        Trustee           = $upn
                        Error             = ''
                    }
                }
            }
        } catch {
            $any = $true
            [pscustomobject]@{
                UserPrincipalName = $upn
                Mailbox           = ''
                Right             = 'FullAccess'
                Trustee           = $upn
                Error             = "Full Access scan: $($_.Exception.Message)"
            }
        }
        if (-not $any) {
            [pscustomobject]@{
                UserPrincipalName = $upn
                Mailbox           = '(none)'
                Right             = ''
                Trustee           = ''
                Error             = ''
            }
        }
    }
}

function Remove-RemediationMailboxElsewhere {
    param([Parameter(Mandatory = $true)][object[]]$Items)
    foreach ($item in $Items) {
        $mailbox = [string]$item.Mailbox
        $right = [string]$item.Right
        $trustee = if ($item.Trustee) { [string]$item.Trustee } else { [string]$item.UserPrincipalName }
        try {
            if ($right -eq 'SendAs') {
                Remove-RecipientPermission -Identity $mailbox -Trustee $trustee -AccessRights SendAs -Confirm:$false -ErrorAction Stop
            } elseif ($right -eq 'FullAccess') {
                Remove-MailboxPermission -Identity $mailbox -User $trustee -AccessRights FullAccess -Confirm:$false -ErrorAction Stop
            } elseif ($right -eq 'SendOnBehalf') {
                Set-Mailbox -Identity $mailbox -GrantSendOnBehalfTo @{ Remove = $trustee } -ErrorAction Stop
            } else {
                throw "Unknown right '$right'"
            }
            [pscustomobject]@{ Mailbox = $mailbox; Right = $right; Trustee = $trustee; Success = $true; Error = '' }
        } catch {
            [pscustomobject]@{ Mailbox = $mailbox; Right = $right; Trustee = $trustee; Success = $false; Error = $_.Exception.Message }
        }
    }
}

function Get-RemediationFlows {
    if (-not (Get-Command Get-AdminFlow -ErrorAction SilentlyContinue)) {
        return [pscustomobject]@{
            Available = $false
            Message   = 'Power Automate listing is not available in this session (Get-AdminFlow / Power Platform admin module is not loaded). Review flows at https://make.powerautomate.com for the compromised user.'
            Flows     = @()
        }
    }
    $rows = [System.Collections.Generic.List[object]]::new()
    try {
        foreach ($flow in @(Get-AdminFlow -ErrorAction Stop)) {
            [void]$rows.Add([pscustomobject]@{
                Id          = [string]$flow.FlowName
                Name        = [string]$flow.DisplayName
                Environment = [string]$flow.EnvironmentName
                Enabled     = if ($null -ne $flow.Enabled) { [bool]$flow.Enabled } else { $null }
                CreatedBy   = [string]$flow.CreatedBy.userId
                Error       = ''
            })
        }
    } catch {
        return [pscustomobject]@{
            Available = $false
            Message   = $_.Exception.Message
            Flows     = @()
        }
    }
    return [pscustomobject]@{
        Available = $true
        Message   = ''
        Flows     = @($rows)
    }
}

function Remove-RemediationFlows {
    param([Parameter(Mandatory = $true)][object[]]$Items)
    if (-not (Get-Command Remove-AdminFlow -ErrorAction SilentlyContinue)) {
        foreach ($item in $Items) {
            [pscustomobject]@{ Id = [string]$item.Id; Success = $false; Error = 'Remove-AdminFlow is not available in this session.' }
        }
        return
    }
    foreach ($item in $Items) {
        $id = [string]$item.Id
        $env = [string]$item.Environment
        try {
            Remove-AdminFlow -FlowName $id -EnvironmentName $env -ErrorAction Stop
            [pscustomobject]@{ Id = $id; Success = $true; Error = '' }
        } catch {
            [pscustomobject]@{ Id = $id; Success = $false; Error = $_.Exception.Message }
        }
    }
}

function Invoke-RemediationExtendedCommand {
    param(
        [Parameter(Mandatory = $true)][string]$Command,
        $Capabilities,
        [string]$AuditFolder,
        [string]$TenantLabel,
        [int]$ClientNumber
    )
    $caps = $Capabilities

    if ($Command -match '^REMEDIATE_LIST_OAUTH_GRANTS') {
        $users = Convert-RemediationUsersFromCommand -Command $Command
        $rows = @(Get-RemediationOAuthGrants -UserPrincipalNames $users)
        $payload = @{ Grants = $rows; Capabilities = $caps } | ConvertTo-Json -Compress -Depth 6
        return "REMEDIATE_OAUTH:$payload"
    }
    if ($Command -match '^REMEDIATE_DELETE_OAUTH_GRANTS') {
        $items = @(Convert-RemediationJsonTailFromCommand -Command $Command -Name 'ITEMS')
        $rows = @(Remove-RemediationOAuthGrants -Items $items)
        $remaining = @()
        $users = @($items | ForEach-Object { $_.UserPrincipalName } | Select-Object -Unique)
        if ($users.Count -gt 0) { $remaining = @(Get-RemediationOAuthGrants -UserPrincipalNames $users) }
        $ok = @($rows | Where-Object { $_.Success }).Count
        $fail = @($rows | Where-Object { -not $_.Success }).Count
        $payload = @{ Deleted = $rows; Grants = $remaining; SuccessCount = $ok; FailCount = $fail; Capabilities = $caps } | ConvertTo-Json -Compress -Depth 6
        if ($fail -gt 0 -and $ok -eq 0) { return "REMEDIATE_FAILED:$payload" }
        return "REMEDIATE_OAUTH:$payload"
    }

    if ($Command -match '^REMEDIATE_REREGISTER_MFA') {
        $users = Convert-RemediationUsersFromCommand -Command $Command
        $rows = @(Invoke-RemediationReregisterMfa -UserPrincipalNames $users)
        foreach ($r in $rows) {
            $result = if ($r.Success) { 'success' } else { 'failed' }
            Write-RemediationAuditCsv -Folder $AuditFolder -Tenant $TenantLabel -ClientNumber $ClientNumber -Upn $r.UserPrincipalName -Action 'reregister-mfa' -Result $result -Detail "deleted=$($r.DeletedCount); revoked=$($r.SessionsRevoked); $($r.Error)"
        }
        $ok = @($rows | Where-Object { $_.Success }).Count
        $fail = @($rows | Where-Object { -not $_.Success }).Count
        $payload = @{ Users = $rows; SuccessCount = $ok; FailCount = $fail; Capabilities = $caps } | ConvertTo-Json -Compress -Depth 6
        if ($fail -gt 0 -and $ok -eq 0) { return "REMEDIATE_FAILED:$payload" }
        return "REMEDIATE_SUCCESS:$payload"
    }

    if ($Command -match '^REMEDIATE_LIST_INTUNE') {
        $users = Convert-RemediationUsersFromCommand -Command $Command
        $rows = @(Get-RemediationIntuneDevices -UserPrincipalNames $users)
        $payload = @{ Devices = $rows; Capabilities = $caps } | ConvertTo-Json -Compress -Depth 6
        return "REMEDIATE_INTUNE:$payload"
    }
    if ($Command -match '^REMEDIATE_WIPE_INTUNE' -or $Command -match '^REMEDIATE_RETIRE_INTUNE') {
        $action = if ($Command -match '^REMEDIATE_WIPE_INTUNE') { 'wipe' } else { 'retire' }
        $items = @(Convert-RemediationJsonTailFromCommand -Command $Command -Name 'DEVICES')
        $rows = @(Invoke-RemediationIntuneAction -Items $items -Action $action)
        $ok = @($rows | Where-Object { $_.Success }).Count
        $fail = @($rows | Where-Object { -not $_.Success }).Count
        $payload = @{ Results = $rows; SuccessCount = $ok; FailCount = $fail; Capabilities = $caps } | ConvertTo-Json -Compress -Depth 6
        if ($fail -gt 0 -and $ok -eq 0) { return "REMEDIATE_FAILED:$payload" }
        return "REMEDIATE_INTUNE:$payload"
    }

    if ($Command -match '^REMEDIATE_LIST_ROLES') {
        $users = Convert-RemediationUsersFromCommand -Command $Command
        $rows = @(Get-RemediationDirectoryRoles -UserPrincipalNames $users)
        $payload = @{ Roles = $rows; Capabilities = $caps } | ConvertTo-Json -Compress -Depth 6
        return "REMEDIATE_ROLES:$payload"
    }
    if ($Command -match '^REMEDIATE_DELETE_ROLES') {
        $items = @(Convert-RemediationJsonTailFromCommand -Command $Command -Name 'ITEMS')
        $rows = @(Remove-RemediationDirectoryRoles -Items $items)
        $users = @($items | ForEach-Object { $_.UserPrincipalName } | Select-Object -Unique)
        $remaining = @()
        if ($users.Count -gt 0) { $remaining = @(Get-RemediationDirectoryRoles -UserPrincipalNames $users) }
        $ok = @($rows | Where-Object { $_.Success }).Count
        $fail = @($rows | Where-Object { -not $_.Success }).Count
        $payload = @{ Deleted = $rows; Roles = $remaining; SuccessCount = $ok; FailCount = $fail; Capabilities = $caps } | ConvertTo-Json -Compress -Depth 6
        if ($fail -gt 0 -and $ok -eq 0) { return "REMEDIATE_FAILED:$payload" }
        return "REMEDIATE_ROLES:$payload"
    }

    if ($Command -match '^REMEDIATE_LIST_APP_CREDS') {
        $rows = @(Get-RemediationAppCredentials)
        $payload = @{ Credentials = $rows; Capabilities = $caps } | ConvertTo-Json -Compress -Depth 8
        return "REMEDIATE_APPCREDS:$payload"
    }
    if ($Command -match '^REMEDIATE_DELETE_APP_CREDS') {
        $items = @(Convert-RemediationJsonTailFromCommand -Command $Command -Name 'ITEMS')
        $rows = @(Remove-RemediationAppCredentials -Items $items)
        $remaining = @(Get-RemediationAppCredentials)
        $ok = @($rows | Where-Object { $_.Success }).Count
        $fail = @($rows | Where-Object { -not $_.Success }).Count
        $payload = @{ Deleted = $rows; Credentials = $remaining; SuccessCount = $ok; FailCount = $fail; Capabilities = $caps } | ConvertTo-Json -Compress -Depth 8
        if ($fail -gt 0 -and $ok -eq 0) { return "REMEDIATE_FAILED:$payload" }
        return "REMEDIATE_APPCREDS:$payload"
    }

    if ($Command -match '^REMEDIATE_LIST_MOBILE_DEVICES') {
        $users = Convert-RemediationUsersFromCommand -Command $Command
        $rows = @(Get-RemediationMobileDevices -UserPrincipalNames $users)
        $payload = @{ Devices = $rows } | ConvertTo-Json -Compress -Depth 6
        return "REMEDIATE_MOBILE:$payload"
    }
    if ($Command -match '^REMEDIATE_DELETE_MOBILE_DEVICES') {
        $items = @(Convert-RemediationJsonTailFromCommand -Command $Command -Name 'ITEMS')
        $rows = @(Remove-RemediationMobileDevices -Items $items)
        $users = @($items | ForEach-Object { $_.UserPrincipalName } | Select-Object -Unique)
        $remaining = @()
        if ($users.Count -gt 0) { $remaining = @(Get-RemediationMobileDevices -UserPrincipalNames $users) }
        $ok = @($rows | Where-Object { $_.Success }).Count
        $fail = @($rows | Where-Object { -not $_.Success }).Count
        $payload = @{ Deleted = $rows; Devices = $remaining; SuccessCount = $ok; FailCount = $fail } | ConvertTo-Json -Compress -Depth 6
        if ($fail -gt 0 -and $ok -eq 0) { return "REMEDIATE_FAILED:$payload" }
        return "REMEDIATE_MOBILE:$payload"
    }

    if ($Command -match '^REMEDIATE_LIST_FOLDER_PERMS') {
        $users = Convert-RemediationUsersFromCommand -Command $Command
        $rows = @(Get-RemediationFolderPermissions -UserPrincipalNames $users)
        $payload = @{ Permissions = $rows } | ConvertTo-Json -Compress -Depth 6
        return "REMEDIATE_FOLDERS:$payload"
    }
    if ($Command -match '^REMEDIATE_DELETE_FOLDER_PERMS') {
        $items = @(Convert-RemediationJsonTailFromCommand -Command $Command -Name 'ITEMS')
        $rows = @(Remove-RemediationFolderPermissions -Items $items)
        $users = @($items | ForEach-Object { $_.UserPrincipalName } | Select-Object -Unique)
        $remaining = @()
        if ($users.Count -gt 0) { $remaining = @(Get-RemediationFolderPermissions -UserPrincipalNames $users) }
        $ok = @($rows | Where-Object { $_.Success }).Count
        $fail = @($rows | Where-Object { -not $_.Success }).Count
        $payload = @{ Deleted = $rows; Permissions = $remaining; SuccessCount = $ok; FailCount = $fail } | ConvertTo-Json -Compress -Depth 6
        if ($fail -gt 0 -and $ok -eq 0) { return "REMEDIATE_FAILED:$payload" }
        return "REMEDIATE_FOLDERS:$payload"
    }

    if ($Command -match '^REMEDIATE_GET_AUTOREPLY') {
        $users = Convert-RemediationUsersFromCommand -Command $Command
        $rows = @(Get-RemediationAutoReply -UserPrincipalNames $users)
        $payload = @{ Users = $rows } | ConvertTo-Json -Compress -Depth 6
        return "REMEDIATE_AUTOREPLY:$payload"
    }
    if ($Command -match '^REMEDIATE_DISABLE_AUTOREPLY') {
        $users = Convert-RemediationUsersFromCommand -Command $Command
        $rows = @(Disable-RemediationAutoReply -UserPrincipalNames $users)
        $remaining = @(Get-RemediationAutoReply -UserPrincipalNames $users)
        $ok = @($rows | Where-Object { $_.Success }).Count
        $fail = @($rows | Where-Object { -not $_.Success }).Count
        $payload = @{ Users = $remaining; SuccessCount = $ok; FailCount = $fail } | ConvertTo-Json -Compress -Depth 6
        if ($fail -gt 0 -and $ok -eq 0) { return "REMEDIATE_FAILED:$payload" }
        return "REMEDIATE_AUTOREPLY:$payload"
    }

    if ($Command -match '^REMEDIATE_LIST_ORG_FORWARD') {
        $rows = @(Get-RemediationOrgForward)
        $payload = @{ Policies = $rows } | ConvertTo-Json -Compress -Depth 6
        return "REMEDIATE_ORGFWD:$payload"
    }
    if ($Command -match '^REMEDIATE_SET_ORG_FORWARD') {
        $items = @(Convert-RemediationJsonTailFromCommand -Command $Command -Name 'ITEMS')
        $rows = @(Set-RemediationOrgForward -Items $items)
        $remaining = @(Get-RemediationOrgForward)
        $ok = @($rows | Where-Object { $_.Success }).Count
        $fail = @($rows | Where-Object { -not $_.Success }).Count
        $payload = @{ Updated = $rows; Policies = $remaining; SuccessCount = $ok; FailCount = $fail } | ConvertTo-Json -Compress -Depth 6
        if ($fail -gt 0 -and $ok -eq 0) { return "REMEDIATE_FAILED:$payload" }
        return "REMEDIATE_ORGFWD:$payload"
    }

    if ($Command -match '^REMEDIATE_LIST_JUNK') {
        $users = Convert-RemediationUsersFromCommand -Command $Command
        $rows = @(Get-RemediationJunk -UserPrincipalNames $users)
        $payload = @{ Entries = $rows } | ConvertTo-Json -Compress -Depth 6
        return "REMEDIATE_JUNK:$payload"
    }
    if ($Command -match '^REMEDIATE_DELETE_JUNK') {
        $items = @(Convert-RemediationJsonTailFromCommand -Command $Command -Name 'ITEMS')
        $rows = @(Remove-RemediationJunk -Items $items)
        $users = @($items | ForEach-Object { $_.UserPrincipalName } | Select-Object -Unique)
        $remaining = @()
        if ($users.Count -gt 0) { $remaining = @(Get-RemediationJunk -UserPrincipalNames $users) }
        $ok = @($rows | Where-Object { $_.Success }).Count
        $fail = @($rows | Where-Object { -not $_.Success }).Count
        $payload = @{ Deleted = $rows; Entries = $remaining; SuccessCount = $ok; FailCount = $fail } | ConvertTo-Json -Compress -Depth 6
        if ($fail -gt 0 -and $ok -eq 0) { return "REMEDIATE_FAILED:$payload" }
        return "REMEDIATE_JUNK:$payload"
    }

    if ($Command -match '^REMEDIATE_LIST_JOURNAL') {
        $rows = @(Get-RemediationJournalRules)
        $payload = @{ Rules = $rows } | ConvertTo-Json -Compress -Depth 6
        return "REMEDIATE_JOURNAL:$payload"
    }
    if ($Command -match '^REMEDIATE_DELETE_JOURNAL') {
        $items = @(Convert-RemediationJsonTailFromCommand -Command $Command -Name 'ITEMS')
        $rows = @(Remove-RemediationJournalRules -Items $items)
        $remaining = @(Get-RemediationJournalRules)
        $ok = @($rows | Where-Object { $_.Success }).Count
        $fail = @($rows | Where-Object { -not $_.Success }).Count
        $payload = @{ Deleted = $rows; Rules = $remaining; SuccessCount = $ok; FailCount = $fail } | ConvertTo-Json -Compress -Depth 6
        if ($fail -gt 0 -and $ok -eq 0) { return "REMEDIATE_FAILED:$payload" }
        return "REMEDIATE_JOURNAL:$payload"
    }

    if ($Command -match '^REMEDIATE_GET_MAILBOX_HOLD') {
        $users = Convert-RemediationUsersFromCommand -Command $Command
        $rows = @(Get-RemediationMailboxHold -UserPrincipalNames $users)
        $payload = @{ Users = $rows } | ConvertTo-Json -Compress -Depth 6
        return "REMEDIATE_HOLD:$payload"
    }
    if ($Command -match '^REMEDIATE_SET_MAILBOX_HOLD') {
        $users = Convert-RemediationUsersFromCommand -Command $Command
        $rows = @(Set-RemediationMailboxHold -UserPrincipalNames $users)
        $remaining = @(Get-RemediationMailboxHold -UserPrincipalNames $users)
        $ok = @($rows | Where-Object { $_.Success }).Count
        $fail = @($rows | Where-Object { -not $_.Success }).Count
        $payload = @{ Users = $remaining; SuccessCount = $ok; FailCount = $fail } | ConvertTo-Json -Compress -Depth 6
        if ($fail -gt 0 -and $ok -eq 0) { return "REMEDIATE_FAILED:$payload" }
        return "REMEDIATE_HOLD:$payload"
    }

    if ($Command -match '^REMEDIATE_LIST_ELSEWHERE') {
        $users = Convert-RemediationUsersFromCommand -Command $Command
        $rows = @(Get-RemediationMailboxElsewhere -UserPrincipalNames $users)
        $payload = @{ Grants = $rows } | ConvertTo-Json -Compress -Depth 6
        return "REMEDIATE_ELSEWHERE:$payload"
    }
    if ($Command -match '^REMEDIATE_DELETE_ELSEWHERE') {
        $items = @(Convert-RemediationJsonTailFromCommand -Command $Command -Name 'ITEMS')
        $rows = @(Remove-RemediationMailboxElsewhere -Items $items)
        $users = @($items | ForEach-Object { $_.UserPrincipalName } | Select-Object -Unique)
        $remaining = @()
        if ($users.Count -gt 0) { $remaining = @(Get-RemediationMailboxElsewhere -UserPrincipalNames $users) }
        $ok = @($rows | Where-Object { $_.Success }).Count
        $fail = @($rows | Where-Object { -not $_.Success }).Count
        $payload = @{ Deleted = $rows; Grants = $remaining; SuccessCount = $ok; FailCount = $fail } | ConvertTo-Json -Compress -Depth 6
        if ($fail -gt 0 -and $ok -eq 0) { return "REMEDIATE_FAILED:$payload" }
        return "REMEDIATE_ELSEWHERE:$payload"
    }

    if ($Command -match '^REMEDIATE_LIST_FLOWS') {
        $data = Get-RemediationFlows
        $payload = $data | ConvertTo-Json -Compress -Depth 6
        return "REMEDIATE_FLOWS:$payload"
    }
    if ($Command -match '^REMEDIATE_DELETE_FLOWS') {
        $items = @(Convert-RemediationJsonTailFromCommand -Command $Command -Name 'ITEMS')
        $rows = @(Remove-RemediationFlows -Items $items)
        $data = Get-RemediationFlows
        $ok = @($rows | Where-Object { $_.Success }).Count
        $fail = @($rows | Where-Object { -not $_.Success }).Count
        $payload = @{ Deleted = $rows; Flows = $data.Flows; Available = $data.Available; Message = $data.Message; SuccessCount = $ok; FailCount = $fail } | ConvertTo-Json -Compress -Depth 6
        if ($fail -gt 0 -and $ok -eq 0) { return "REMEDIATE_FAILED:$payload" }
        return "REMEDIATE_FLOWS:$payload"
    }

    return $null
}
