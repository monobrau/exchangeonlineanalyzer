#Requires -Version 5.1
#Requires -Modules Pester

$modulePath = Join-Path (Split-Path $PSScriptRoot -Parent) 'Modules\ExportUtils.psm1'
Import-Module $modulePath -Force

Describe 'Sign-in log source' {
    It 'treats the non-premium Graph error as a license block' {
        InModuleScope ExportUtils {
            $body = '{"error":{"code":"Authentication_RequestFromNonPremiumTenantOrB2CTenant","message":"Tenant is not a B2C tenant and doesn''t have premium license"}}'
            Test-SignInLogLicenseBlock $body | Should -Be $true
            Test-SignInLogLicenseBlock 'Response status code does not indicate success: 403 (Forbidden).' | Should -Be $false
            Test-SignInLogLicenseBlock 'Sign-in logs are not included in this tenant license.' | Should -Be $true
        }
    }

    It 'only maps UserLoggedIn / UserLoginFailed unified audit operations' {
        InModuleScope ExportUtils {
            Test-UnifiedAuditSignInOperation 'UserLoggedIn' | Should -Be $true
            Test-UnifiedAuditSignInOperation 'UserLoginFailed' | Should -Be $true
            Test-UnifiedAuditSignInOperation 'Add user.' | Should -Be $false
        }
    }

    It 'maps a unified audit STS logon into a sign-in row' {
        InModuleScope ExportUtils {
            $audit = @{
                CreationTime = '2026-10-20T15:04:00Z'
                UserId = 'pat@contoso.com'
                UserKey = 'user-key'
                ClientIP = '203.0.113.10'
                ResultStatus = 'Success'
                ApplicationId = '00000002-0000-0ff1-ce00-000000000000'
                ExtendedProperties = @(
                    @{ Name = 'RequestType'; Value = 'OAuth2:Authorize' }
                    @{ Name = 'UserAuthenticationMethod'; Value = 'Password' }
                )
            } | ConvertTo-Json -Depth 5
            $row = [pscustomobject]@{
                CreationDate = '2026-10-20T15:04:00Z'
                Operations = 'UserLoggedIn'
                UserIds = 'pat@contoso.com'
                AuditData = $audit
            }
            $signIn = ConvertTo-SignInLogFromUnifiedAudit $row
            $signIn.UserPrincipalName | Should -Be 'pat@contoso.com'
            $signIn.IPAddress | Should -Be '203.0.113.10'
            $signIn.Status | Should -Be 'Success'
            $signIn.ClientAppUsed | Should -Be 'OAuth2:Authorize'
            $signIn.AuthMethods | Should -Be 'Password'
            $signIn.AppDisplayName | Should -Be '00000002-0000-0ff1-ce00-000000000000'
            $signIn.SignInSource | Should -Be 'UnifiedAudit'
        }
    }

    It 'skips non-sign-in unified audit operations' {
        InModuleScope ExportUtils {
            $row = [pscustomobject]@{
                CreationDate = '2026-10-20T15:04:00Z'
                Operations = 'Add user.'
                UserIds = 'pat@contoso.com'
                AuditData = '{}'
            }
            ConvertTo-SignInLogFromUnifiedAudit $row | Should -Be $null
        }
    }

    It 'uses Exchange first and skips Graph when UAL returns rows' {
        InModuleScope ExportUtils {
            $script:graphCalls = 0
            $ualRow = [pscustomobject]@{
                CreatedDateTime = '2026-10-20T15:04:00Z'
                UserPrincipalName = 'pat@contoso.com'
                Status = 'Success'
                SignInSource = 'UnifiedAudit'
            }
            Mock Get-Command {
                param($Name)
                if ($Name -eq 'Search-UnifiedAuditLog') { return @{ Name = $Name } }
                return $null
            }
            Mock Add-UnifiedAuditSignInLogsToReport {
                param($Report)
                $Report.SignInLogs = @($ualRow)
                $Report.SignInLogsCoverageNote = 'from-ual'
                return $true
            }
            Mock Add-GraphSignInLogsToReport {
                $script:graphCalls++
                throw 'Graph should not be called when UAL succeeds'
            }

            $report = @{ SignInLogs = @(); SignInLogsError = $null; SignInLogsCoverageNote = $null }
            Add-SignInLogsToReport -Report $report -DaysBack 7 -ExchangeAvailable:$true -GraphAvailable:$true
            $report.SignInLogs.Count | Should -Be 1
            $report.SignInLogs[0].UserPrincipalName | Should -Be 'pat@contoso.com'
            $report.SignInLogsCoverageNote | Should -Be 'from-ual'
            $script:graphCalls | Should -Be 0
        }
    }

    It 'falls back to Graph when UAL returns no rows' {
        InModuleScope ExportUtils {
            $script:graphCalls = 0
            Mock Get-Command {
                param($Name)
                if ($Name -eq 'Search-UnifiedAuditLog') { return @{ Name = $Name } }
                return $null
            }
            Mock Add-UnifiedAuditSignInLogsToReport {
                param($Report)
                $Report.SignInLogs = @()
                return $false
            }
            Mock Add-GraphSignInLogsToReport {
                param($Report)
                $script:graphCalls++
                $Report.SignInLogs = @([pscustomobject]@{ UserPrincipalName = 'graph@contoso.com'; Status = 'Success' })
                $Report.SignInLogsError = $null
                $Report.SignInLogsCoverageNote = $null
                return $true
            }

            $report = @{ SignInLogs = @(); SignInLogsError = $null; SignInLogsCoverageNote = $null }
            Add-SignInLogsToReport -Report $report -DaysBack 7 -ExchangeAvailable:$true -GraphAvailable:$true
            $report.SignInLogs.Count | Should -Be 1
            $report.SignInLogs[0].UserPrincipalName | Should -Be 'graph@contoso.com'
            $script:graphCalls | Should -Be 1
        }
    }
}

$settingsPath = Join-Path (Split-Path $PSScriptRoot -Parent) 'Modules\Settings.psm1'
Import-Module $settingsPath -Force

Describe 'Sign-in auth requirements' {
    It 'requires Exchange for IncludeSignInLogs and does not force Graph' {
        $auth = Get-RequiredAuthFromReportSelections -ReportSelections @{ IncludeSignInLogs = $true }
        $auth.NeedsExchange | Should -Be $true
        $auth.NeedsGraph | Should -Be $false
    }
}
