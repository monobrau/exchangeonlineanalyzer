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

    Import-Module ExchangeOnlineManagement -Force -ErrorAction Stop
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

Export-ModuleMember -Function Test-ExchangeModule,Install-ExchangeModule,Get-ConnectExchangeOnlineParams,Clear-ExchangeOnlineConnectionState,Import-ExchangeOnlineManagementFirst,Connect-ExchangeOnlineWithDefaults,Get-ExchangeOnlineSendingRestrictions 