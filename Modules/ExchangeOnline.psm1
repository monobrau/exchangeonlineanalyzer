function Test-ExchangeModule {
    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        return $false
    }
    return $true
}

function Install-ExchangeModule {
    Write-Host "Attempting to install ExchangeOnlineManagement module..." -ForegroundColor Yellow
    try {
        Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Repository PSGallery -Force -AllowClobber -ErrorAction Stop
        Write-Host "ExchangeOnlineManagement module installed successfully. Please restart the script." -ForegroundColor Green
        return $true
    } catch {
        $ex = $_.Exception 
        Write-Error ("Failed to install ExchangeOnlineManagement module. Please install it manually: Install-Module ExchangeOnlineManagement -Scope CurrentUser. Error: {0}" -f $ex.Message)
        return $false
    }
}

function Get-ConnectExchangeOnlineParams {
    [CmdletBinding()]
    param(
        [hashtable]$AdditionalParams = @{}
    )

    $params = @{
        ErrorAction = 'Stop'
    }
    foreach ($key in $AdditionalParams.Keys) {
        $params[$key] = $AdditionalParams[$key]
    }

    $exoConnect = Get-Command Connect-ExchangeOnline -ErrorAction SilentlyContinue
    if ($exoConnect) {
        $supportedParams = $exoConnect.Parameters.Keys
        if ($supportedParams -contains 'ShowBanner') {
            $params['ShowBanner'] = $false
        }
        if ($supportedParams -contains 'DisableWAM') {
            $params['DisableWAM'] = $true
        }
        if ($supportedParams -contains 'SkipLoadingCmdletHelp') {
            $params['SkipLoadingCmdletHelp'] = $true
        }
    }

    return $params
}

function Clear-ExchangeOnlineConnectionState {
    [CmdletBinding()]
    param(
        [int]$SettleMilliseconds = 750
    )

    $global:EOA_MessageTraceV2Command = $null
    $global:EOA_ExchangeHasUserMailboxes = $null
    try {
        if (Get-Command Disconnect-ExchangeOnline -ErrorAction SilentlyContinue) {
            Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
        }
    } catch { }

    try {
        Get-PSSession -ErrorAction SilentlyContinue |
            Where-Object {
                $_.ConfigurationName -eq 'Microsoft.Exchange' -or
                $_.ComputerName -like '*.outlook.office365.com' -or
                $_.ComputerName -like '*.outlook.com'
            } |
            ForEach-Object { Remove-PSSession -Session $_ -ErrorAction SilentlyContinue }
    } catch { }

    if ($SettleMilliseconds -gt 0) {
        Start-Sleep -Milliseconds $SettleMilliseconds
    }
}

function Import-ExchangeOnlineManagementFirst {
    [CmdletBinding()]
    param()

    if (Get-Module ExchangeOnlineManagement -ErrorAction SilentlyContinue) {
        return
    }

    foreach ($name in @('AZURE_IDENTITY_DISABLE_BROKER', 'MSAL_DISABLE_BROKER', 'MSAL_EXPERIMENTAL_DISABLE_BROKER')) {
        Set-Item -Path "Env:$name" -Value '1' -ErrorAction SilentlyContinue
    }
    $env:MSAL_FORCE_WAM = '0'

    $latest = Get-Module -ListAvailable -Name ExchangeOnlineManagement |
        Sort-Object Version -Descending |
        Select-Object -First 1
    if ($latest -and $latest.Version -ge [version]'3.7.0') {
        Import-Module ExchangeOnlineManagement -RequiredVersion $latest.Version -Force -ErrorAction Stop
    } else {
        Import-Module ExchangeOnlineManagement -Force -ErrorAction Stop
    }
}

function Connect-ExchangeOnlineWithDefaults {
    [CmdletBinding()]
    param(
        [hashtable]$AdditionalParams = @{}
    )

    Import-ExchangeOnlineManagementFirst

    foreach ($name in @('AZURE_IDENTITY_DISABLE_BROKER', 'MSAL_DISABLE_BROKER', 'MSAL_EXPERIMENTAL_DISABLE_BROKER')) {
        Set-Item -Path "Env:$name" -Value '1' -ErrorAction SilentlyContinue
    }
    $env:MSAL_FORCE_WAM = '0'

    $connectParams = Get-ConnectExchangeOnlineParams -AdditionalParams $AdditionalParams
    $exoConnect = Get-Command Connect-ExchangeOnline -ErrorAction SilentlyContinue

    try {
        Connect-ExchangeOnline @connectParams
        Write-ExchangeMessageTraceV2Availability
        return
    } catch {
        $msg = $_.Exception.Message
        if ($msg -match 'Method not found' -and $msg -match 'WithBroker|BrokerExtension|BrokerOptions') {
            throw (
                'Exchange Online authentication failed due to an MSAL assembly conflict in this PowerShell process. ' +
                'Restart the tenant worker, run Exchange Auth before Graph Auth, and prefer PowerShell 7 when available. ' +
                "Detail: $msg"
            )
        }
        if ($msg -notmatch 'WithBroker|BrokerOptions|BrokerExtension|Broker') {
            throw
        }
    }

    Write-Warning 'Exchange Online broker/WAM connect failed; retrying with -UseRPSSession (legacy auth path).'
    $retryParams = Get-ConnectExchangeOnlineParams -AdditionalParams $AdditionalParams
    if ($exoConnect -and ($exoConnect.Parameters.Keys -contains 'UseRPSSession')) {
        $retryParams['UseRPSSession'] = $true
    }
    try {
        Connect-ExchangeOnline @retryParams
        Write-ExchangeMessageTraceV2Availability
    } catch {
        $msg = $_.Exception.Message
        if ($msg -match 'Method not found' -and $msg -match 'WithBroker|BrokerExtension|BrokerOptions') {
            throw (
                'Exchange Online authentication failed due to an MSAL assembly conflict in this PowerShell process. ' +
                'Restart the tenant worker, run Exchange Auth before Graph Auth, and prefer PowerShell 7 when available. ' +
                "Detail: $msg"
            )
        }
        throw
    }
}

function Get-TmpExoModules {
    return @(Get-Module -All -ErrorAction SilentlyContinue | Where-Object { $_.Name -like 'tmpEXO_*' })
}

function Get-ExchangeMessageTraceV2Command {
    if ($global:EOA_MessageTraceV2Command) { return $global:EOA_MessageTraceV2Command }
    foreach ($mod in (Get-TmpExoModules)) {
        if ($mod.ExportedCommands -and $mod.ExportedCommands.ContainsKey('Get-MessageTraceV2') -and $mod.Name -like 'tmpEXO_*') {
            return $mod.ExportedCommands['Get-MessageTraceV2']
        }
    }
    $cmd = Microsoft.PowerShell.Core\Get-Command -Name Get-MessageTraceV2 -All -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($cmd -and $cmd.Source -like 'tmpEXO_*') { return $cmd }
    return $null
}

function Get-ExchangeMessageTraceV2RoleEntryHint {
    # A tenant/admin can be missing the new Message Trace cmdlets from RBAC role entries,
    # which is why the REST session exports only the legacy MessageTrace cmdlets.
    if (-not (Get-Command Get-ManagementRoleEntry -ErrorAction SilentlyContinue)) { return '' }
    try {
        $entries = @(Get-ManagementRoleEntry '*\Get-MessageTraceV2' -ErrorAction Stop)
        if ($entries.Count -gt 0) {
            return (" RBAC: Get-MessageTraceV2 exists in {0} role(s) (e.g. {1}) - the cmdlet is assignable but was not exported to this session." -f $entries.Count, (@($entries | Select-Object -First 3 | ForEach-Object { $_.Role }) -join ', '))
        }
        return ' RBAC: no management role contains Get-MessageTraceV2 in this tenant, so the REST session cannot export it.'
    } catch {
        $msg = [string]$_.Exception.Message
        if ($msg -match "couldn't be found|not found|does not exist") {
            return ' RBAC: no management role contains Get-MessageTraceV2 in this tenant, so the REST session cannot export it.'
        }
        return " RBAC check failed: $msg"
    }
}

function Write-ExchangeMessageTraceV2Availability {
    $cmd = Get-ExchangeMessageTraceV2Command
    if ($cmd) {
        $global:EOA_MessageTraceV2Command = $cmd
        $msg = 'Message trace: Get-MessageTraceV2 is available in this Exchange session.'
        Write-Host $msg -ForegroundColor Green
        if (Get-Command Write-Status -ErrorAction SilentlyContinue) { Write-Status $msg }
        return
    }
    $names = [System.Collections.Generic.List[string]]::new()
    $modNames = [System.Collections.Generic.List[string]]::new()
    foreach ($mod in (Get-TmpExoModules)) {
        [void]$modNames.Add($mod.Name)
        foreach ($key in @($mod.ExportedCommands.Keys)) {
            if ($key -like '*MessageTrace*' -or $key -like '*Trace*') { [void]$names.Add($key) }
        }
    }
    $connHint = ''
    try {
        $conn = Get-ConnectionInformation -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($conn) {
            $method = $conn.ConnectionMethod
            if (-not $method) { $method = $conn.TokenStatus }
            $connHint = " connection=$method"
        }
    } catch { }
    $listed = if ($names.Count) { ($names | Select-Object -Unique) -join ', ' } else { '(none)' }
    $mods = if ($modNames.Count) { ($modNames | Select-Object -Unique) -join ', ' } else { '(no tmpEXO module)' }
    $roleHint = Get-ExchangeMessageTraceV2RoleEntryHint
    $warn = "Get-MessageTraceV2 not loaded. tmpEXO=$mods MessageTrace cmdlets: $listed.$connHint$roleHint Generate will try Graph messageTraces, then Start-HistoricalSearch."
    Write-Warning $warn
    if (Get-Command Write-Status -ErrorAction SilentlyContinue) { Write-Status $warn }
}

function Get-ExchangeOnlineSendingRestrictions {
    param(
        [Parameter(Mandatory=$true)]
        [string]$UserPrincipalName
    )
    try {
        $mailbox = Get-Mailbox -Identity $UserPrincipalName -ErrorAction Stop
        $restrictions = @{
            RequireSenderAuthenticationEnabled = $mailbox.RequireSenderAuthenticationEnabled
            AcceptMessagesOnlyFrom = $mailbox.AcceptMessagesOnlyFrom
            AcceptMessagesOnlyFromDLMembers = $mailbox.AcceptMessagesOnlyFromDLMembers
            RejectMessagesFrom = $mailbox.RejectMessagesFrom
            RejectMessagesFromDLMembers = $mailbox.RejectMessagesFromDLMembers
        }
        try {
            $orgConfig = Get-OrganizationConfig
            $restrictions.OutboundSpamFilteringEnabled = $orgConfig.OutboundSpamFilteringEnabled
        } catch {}
        return $restrictions
    } catch {
        Write-Error "Could not retrieve sending restrictions for $UserPrincipalName : $($_.Exception.Message)"
        return $null
    }
}

Export-ModuleMember -Function Test-ExchangeModule,Install-ExchangeModule,Get-ConnectExchangeOnlineParams,Clear-ExchangeOnlineConnectionState,Import-ExchangeOnlineManagementFirst,Connect-ExchangeOnlineWithDefaults,Get-TmpExoModules,Get-ExchangeMessageTraceV2Command,Get-ExchangeMessageTraceV2RoleEntryHint,Write-ExchangeMessageTraceV2Availability,Get-ExchangeOnlineSendingRestrictions 