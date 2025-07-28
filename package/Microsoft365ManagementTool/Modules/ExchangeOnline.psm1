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

Export-ModuleMember -Function Test-ExchangeModule,Install-ExchangeModule,Connect-ExchangeOnlineAnalyzer,Disconnect-ExchangeOnlineAnalyzer 