# Returns:
#   @{ SecurityDefaultsEnabled = <bool>; CAPoliciesRequireMfa = <bool>; Users = <list of user objects> }
function Get-MfaCoverageReport {
    param(
        [Parameter(Mandatory=$false)]
        [array]$SelectedUsers = @()
    )
    
    try {
        # 1) Security Defaults status (authoritative)
        $secDefaultsEnabled = $false
        try {
            $secDefaults = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/policies/identitySecurityDefaultsEnforcementPolicy' -ErrorAction Stop
            if ($secDefaults -and $secDefaults.isEnabled -ne $null) { $secDefaultsEnabled = [bool]$secDefaults.isEnabled }
        } catch {}

        # 2) Conditional Access policies requiring MFA (tenant-wide set)
        $caPolicies = @()
        try {
            $resp = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies?$top=999' -ErrorAction SilentlyContinue
            if ($resp.value) { $caPolicies = $resp.value }
        } catch {}

        # Filter enabled policies that require MFA
        $mfaPolicies = @()
        foreach ($p in $caPolicies) {
            $enabled = ($p.state -eq 'enabled')
            if (-not $enabled) { continue }
            $grant = $p.grantControls
            $requiresMfa = $false
            if ($grant) {
                if ($grant.builtInControls -contains 'mfa') { $requiresMfa = $true }
                # authenticationStrength also implies MFA, but skip for simplicity if missing
            }
            if ($requiresMfa) { $mfaPolicies += $p }
        }

        # 3) Users and per-user evaluation - filter server-side if SelectedUsers provided
        $users = @()
        try {
            # Directory roles map (for policy role assignment evaluation) - needed for CA policy evaluation
            $roles = @(); $roleIdToName = @{}
            try { $roles = Get-MgDirectoryRole -All -ErrorAction SilentlyContinue } catch {}
            foreach ($r in $roles) { $roleIdToName[$r.Id] = $r.DisplayName }

            # If SelectedUsers provided, only query those users (server-side filtering)
            if ($SelectedUsers -and $SelectedUsers.Count -gt 0) {
                $userPage = @()
                foreach ($user in $SelectedUsers) {
                    $upn = if ($user -is [string]) { $user } elseif ($user.UserPrincipalName) { $user.UserPrincipalName } else { continue }
                    try {
                        $u = Get-MgUser -UserId $upn -Property 'id,displayName,userPrincipalName' -ErrorAction Stop
                        if ($u) {
                            $userPage += $u
                        }
                    } catch {
                        Write-Warning "User not found for MFA coverage: ${upn}: $($_.Exception.Message)"
                    }
                }
            } else {
                # No selection - get all users
                $userPage = Get-MgUser -All -Property 'id,displayName,userPrincipalName' -ErrorAction Stop
            }

            foreach ($u in $userPage) {
                $directMfa = $false
                # Attempt to detect registered methods (requires UserAuthenticationMethod.Read.All). Best-effort.
                try {
                    $methods = Invoke-MgGraphRequest -Method GET -Uri ("https://graph.microsoft.com/v1.0/users/{0}/authentication/methods" -f $u.Id) -ErrorAction SilentlyContinue
                    if ($methods.value) {
                        foreach ($m in $methods.value) {
                            $otype = $m.'@odata.type'
                            if ($otype -eq '#microsoft.graph.microsoftAuthenticatorAuthenticationMethod' -or
                                $otype -eq '#microsoft.graph.phoneAuthenticationMethod' -or
                                $otype -eq '#microsoft.graph.softwareOathAuthenticationMethod' -or
                                $otype -eq '#microsoft.graph.fido2AuthenticationMethod' -or
                                $otype -eq '#microsoft.graph.temporaryAccessPassAuthenticationMethod') {
                                $directMfa = $true; break
                            }
                        }
                    }
                } catch {}

                # Determine if any MFA CA policy applies to this user
                $userGroups = @()
                $userRoles = @()
                try {
                    $mem = Get-MgUserMemberOf -UserId $u.Id -All -ErrorAction SilentlyContinue
                    foreach ($m in $mem) {
                        if ($m.'@odata.type' -eq '#microsoft.graph.group') { $userGroups += $m.Id }
                        elseif ($m.'@odata.type' -eq '#microsoft.graph.directoryRole') { $userRoles += $m.Id }
                    }
                } catch {}

                $userCaRequiresMfa = $false
                foreach ($p in $mfaPolicies) {
                    $conds = $p.conditions
                    if (-not $conds) { continue }
                    $usersCond = $conds.users
                    $incAll = $false
                    $incUser = $false
                    $excluded = $false

                    if ($usersCond) {
                        # Include
                        if ($usersCond.includeUsers -and ($usersCond.includeUsers -contains 'All' -or $usersCond.includeUsers -contains $u.Id)) { $incAll = $usersCond.includeUsers -contains 'All'; if (-not $incAll) { $incUser = $true } }
                        if (-not $incUser -and $usersCond.includeGroups) { if (@($usersCond.includeGroups) -ne $null) { if ($usersCond.includeGroups | Where-Object { $userGroups -contains $_ }) { $incUser = $true } } }
                        if (-not $incUser -and $usersCond.includeRoles) { if (@($usersCond.includeRoles) -ne $null) { if ($usersCond.includeRoles | Where-Object { $userRoles -contains $_ }) { $incUser = $true } } }

                        # Exclude
                        if ($usersCond.excludeUsers -and ($usersCond.excludeUsers -contains $u.Id)) { $excluded = $true }
                        if ($usersCond.excludeGroups) { if (@($usersCond.excludeGroups) -ne $null) { if ($usersCond.excludeGroups | Where-Object { $userGroups -contains $_ }) { $excluded = $true } } }
                        if ($usersCond.excludeRoles) { if (@($usersCond.excludeRoles) -ne $null) { if ($usersCond.excludeRoles | Where-Object { $userRoles -contains $_ }) { $excluded = $true } } }
                    }

                    $applies = ($incAll -or $incUser)
                    if ($applies -and -not $excluded) { $userCaRequiresMfa = $true; break }
                }

                $covered = ($directMfa -or $secDefaultsEnabled -or $userCaRequiresMfa)
                $users += [pscustomobject]@{
                    DisplayName        = $u.displayName
                    UserPrincipalName  = $u.userPrincipalName
                    PerUserMfaEnabled  = $directMfa
                    SecurityDefaults   = $secDefaultsEnabled
                    CARequiresMfa      = $userCaRequiresMfa
                    MfaCovered         = $covered
                }
            }
        } catch {}

        $tenantLevelCaMfa = ($mfaPolicies.Count -gt 0)
        return @{ SecurityDefaultsEnabled = $secDefaultsEnabled; CAPoliciesRequireMfa = $tenantLevelCaMfa; Users = $users }
    } catch {
        Write-Error "Get-MfaCoverageReport failed: $($_.Exception.Message)"; return @{ SecurityDefaultsEnabled=$false; CAPoliciesRequireMfa=$false; Users=@() }
    }
}

# Flattens user membership in directory roles and security groups
function Get-UserSecurityGroupsReport {
    param(
        [Parameter(Mandatory=$false)]
        [array]$SelectedUsers = @()
    )
    
    try {
        $results = New-Object System.Collections.Generic.List[object]

        # Directory roles (e.g., Global Administrator)
        $roles = @()
        try { $roles = Get-MgDirectoryRole -All -ErrorAction SilentlyContinue } catch {}
        $roleIdToName = @{}
        foreach ($r in $roles) { $roleIdToName[$r.Id] = $r.DisplayName }

        # Users - filter server-side if SelectedUsers provided
        $users = @()
        if ($SelectedUsers -and $SelectedUsers.Count -gt 0) {
            foreach ($user in $SelectedUsers) {
                $upn = if ($user -is [string]) { $user } elseif ($user.UserPrincipalName) { $user.UserPrincipalName } else { continue }
                try {
                    $u = Get-MgUser -UserId $upn -Property 'id,displayName,userPrincipalName' -ErrorAction Stop
                    if ($u) { $users += $u }
                } catch {
                    Write-Warning "User not found: ${upn}: $($_.Exception.Message)"
                }
            }
        } else {
            # No selection - get all users
            try { $users = Get-MgUser -All -Property 'id,displayName,userPrincipalName' -ErrorAction Stop } catch {}
        }

        foreach ($u in $users) {
            $groups = @()
            try {
                $mem = Get-MgUserMemberOf -UserId $u.Id -All -ErrorAction SilentlyContinue
                foreach ($m in $mem) {
                    $name = $null
                    if ($m.'@odata.type' -eq '#microsoft.graph.group') { $name = $m.DisplayName }
                    elseif ($m.'@odata.type' -eq '#microsoft.graph.directoryRole') { $name = if ($roleIdToName.ContainsKey($m.Id)) { $roleIdToName[$m.Id] } else { 'Directory Role' } }
                    if ($name) { $groups += $name }
                }
            } catch {}

            $results.Add([pscustomobject]@{
                DisplayName       = $u.DisplayName
                UserPrincipalName = $u.UserPrincipalName
                GroupsAndRoles    = ($groups | Sort-Object -Unique) -join '; '
            }) | Out-Null
        }

        return [System.Collections.ArrayList]$results
    } catch { Write-Error "Get-UserSecurityGroupsReport failed: $($_.Exception.Message)"; return @() }
}
function Format-InboxRuleXlsx {
    param(
        [Parameter(Mandatory=$true)]
        [string]$CsvPath,
        [Parameter(Mandatory=$true)]
        [string]$XlsxPath
    )

    $excel = $null; $workbook = $null; $worksheet = $null; $usedRange = $null; $columns = $null; $rows = $null; $headerRange = $null
    $xlOpenXMLWorkbook = 51
    $missing = [System.Reflection.Missing]::Value

    try { $excel = New-Object -ComObject Excel.Application -ErrorAction Stop } 
    catch { 
        $ex = $_.Exception
        Write-Error ("Excel COM object creation failed: {0}" -f $ex.Message)
        return $false 
    }

    try {
        $excel.Visible = $false; $excel.DisplayAlerts = $false    
        $workbook = $excel.Workbooks.Open($CsvPath); $workbook.SaveAs($XlsxPath, $xlOpenXMLWorkbook); $workbook.Close($false) 
        $workbook = $excel.Workbooks.Open($XlsxPath); $worksheet = $workbook.Worksheets.Item(1); $usedRange = $worksheet.UsedRange; $columns = $usedRange.Columns; $rows = $usedRange.Rows

        if ($usedRange.Rows.Count -gt 0) {
            $columns.AutoFit() | Out-Null
            $rows.AutoFit() | Out-Null
            $headerRange = $worksheet.Rows.Item(1)
            $headerRange.Font.Bold = $true
            $headerRange.Interior.Color = 15773696 # Blue header (RGB: 224, 235, 255)
            $headerRange.Font.Color = 1 # Black text
            $headerRange.Borders.LineStyle = 1
            # Find Description column
            $descCol = 0
            $isHiddenCol = 0
            $isCols = @{}
            for ($i = 1; $i -le $usedRange.Columns.Count; $i++) {
                $header = $worksheet.Cells.Item(1, $i).Text
                if ($header -eq 'Description') { $descCol = $i }
                if ($header -eq 'IsHidden') { $isHiddenCol = $i }
                if ($header -like 'Is*') { $isCols[$i] = $header }
            }
            # Wrap and autofit Description column
            if ($descCol -gt 0) {
                $descRange = $worksheet.Columns.Item($descCol)
                $descRange.WrapText = $true
                $descRange.EntireColumn.AutoFit() | Out-Null
            }
            # Apply alternating white/grey background to all data rows
            if ($usedRange.Rows.Count -gt 1) {
                $dataRange = $usedRange.Offset(1,0).Resize($usedRange.Rows.Count -1)
                for ($i = 1; $i -le $dataRange.Rows.Count; $i++) {
                    $rowRange = $dataRange.Rows.Item($i)
                    $rowNum = $i + 1
                    $isHidden = $isHiddenCol -gt 0 -and $worksheet.Cells.Item($rowNum, $isHiddenCol).Value2 -eq $true
                    if ($isHidden) {
                        $rowRange.Interior.Color = 65535 # Bright yellow
                    } elseif ($i % 2 -eq 1) {
                        $rowRange.Interior.Color = 16777215 # White
                    } else {
                        $rowRange.Interior.Color = 15132390 # Light grey (RGB: 230, 230, 230)
                    }
                    $rowRange.Borders.LineStyle = 1
                    # Highlight Is<item> columns that are TRUE
                    for ($colIdx = 1; $colIdx -le $usedRange.Columns.Count; $colIdx++) {
                        $cell = $worksheet.Cells.Item($rowNum, $colIdx)
                        if ($cell.Value2 -eq $true -or ($cell.Value2 -is [string] -and $cell.Value2.ToLower() -eq 'true')) {
                            $cell.Interior.Color = 13421823 # Light red
                        }
                    }
                    # Wrap and autofit Description cell height
                    if ($descCol -gt 0) {
                        $descCell = $worksheet.Cells.Item($rowNum, $descCol)
                        $descCell.WrapText = $true
                        $descCell.EntireRow.AutoFit() | Out-Null
                    }
                }
            }
            # Set RuleID column to text format
            $ruleIdCol = 0
            for ($i = 1; $i -le $usedRange.Columns.Count; $i++) {
                if ($worksheet.Cells.Item(1, $i).Text -eq 'RuleID') { $ruleIdCol = $i; break }
            }
            if ($ruleIdCol -gt 0) {
                $worksheet.Columns.Item($ruleIdCol).NumberFormat = "@"
            }
        }
        $workbook.Save(); $workbook.Close()
    } catch {
        $ex = $_.Exception
        Write-Error ("Excel formatting/conversion error: {0}`n{1}" -f $ex.Message, $ex.ScriptStackTrace)
        try { if ($workbook -ne $null) { $workbook.Close($false) } } catch {}
        return $false 
    } finally {
        if ($columns) {[System.Runtime.InteropServices.Marshal]::ReleaseComObject($columns) | Out-Null}
        if ($rows) {[System.Runtime.InteropServices.Marshal]::ReleaseComObject($rows) | Out-Null}
        if ($usedRange) {[System.Runtime.InteropServices.Marshal]::ReleaseComObject($usedRange) | Out-Null}
        if ($worksheet) {[System.Runtime.InteropServices.Marshal]::ReleaseComObject($worksheet) | Out-Null}
        if ($workbook) {[System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null}
        if ($excel) {$excel.Quit();[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null}
        [gc]::Collect(); [gc]::WaitForPendingFinalizers();
    }
    return $true
}

function New-SecurityInvestigationReport {
    param(
        [Parameter(Mandatory=$false)]
        [string]$InvestigatorName = "Security Administrator",
        [Parameter(Mandatory=$false)]
        [string]$CompanyName = "Organization",
        [Parameter(Mandatory=$false)]
        [int]$DaysBack = 10,
        [Parameter(Mandatory=$false)]
        [object]$StatusLabel,
        [Parameter(Mandatory=$false)]
        [object]$MainForm,
        [Parameter(Mandatory=$false)]
        [string]$OutputFolder,
        [Parameter(Mandatory=$false)]
        [bool]$IncludeMessageTrace = $true,
        [Parameter(Mandatory=$false)]
        [bool]$IncludeInboxRules = $true,
        [Parameter(Mandatory=$false)]
        [bool]$IncludeTransportRules = $true,
        [Parameter(Mandatory=$false)]
        [bool]$IncludeMailFlowConnectors = $true,
        [Parameter(Mandatory=$false)]
        [bool]$IncludeMailboxForwarding = $true,
        [Parameter(Mandatory=$false)]
        [bool]$IncludeAuditLogs = $true,
        [Parameter(Mandatory=$false)]
        [bool]$IncludeConditionalAccessPolicies = $true,
        [Parameter(Mandatory=$false)]
        [bool]$IncludeAppRegistrations = $true,
        [Parameter(Mandatory=$false)]
        [bool]$IncludeSignInLogs = $false,
        [Parameter(Mandatory=$false)]
        [int]$SignInLogsDaysBack = 7,
        [Parameter(Mandatory=$false)]
        [array]$SelectedUsers = @()
    )

    try {
        if ($StatusLabel -and $StatusLabel.GetType().Name -eq "Label") {
            $StatusLabel.Text = "Starting comprehensive security investigation..."
        }
        if ($MainForm -and $MainForm.GetType().Name -eq "Form") {
            $MainForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        }
    } catch {
        # Ignore Windows Forms errors when running outside GUI context
    }

    $report = @{}
    $report.Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $report.Investigator = $InvestigatorName
    $report.Company = $CompanyName
    # Display intent: 10 days for message trace; sign-ins use max available. Keep DaysAnalyzed consistent with 10 unless explicitly provided.
    if (-not $PSBoundParameters.ContainsKey('DaysBack')) { $DaysBack = 10 }
    $report.DaysAnalyzed = $DaysBack
    $report.DataSources = @("Exchange Online", "Microsoft Graph", "Entra ID")
    $report.FilePaths = @{}

    # Resolve output folder (tenant-scoped/timestamped)
    try {
        if ([string]::IsNullOrWhiteSpace($OutputFolder)) {
            $root = Join-Path -Path ([Environment]::GetFolderPath('MyDocuments')) -ChildPath "ExchangeOnlineAnalyzer\SecurityInvestigation"

            # Try to get tenant display name for folder scoping
            $tenantName = $null
            try {
                # Prefer BrowserIntegration helper for a unified identity fetch
                $bi = Join-Path $PSScriptRoot 'BrowserIntegration.psm1'
                if (Test-Path $bi) { Import-Module $bi -Force -ErrorAction SilentlyContinue }
                $ti = $null; try { $ti = Get-TenantIdentity } catch {}
                if ($ti) { if ($ti.TenantDisplayName) { $tenantName = $ti.TenantDisplayName } elseif ($ti.PrimaryDomain) { $tenantName = $ti.PrimaryDomain } }
                if (-not $tenantName) {
                    # Fallback to EXO org display name if available
                    try { $org = Get-OrganizationConfig -ErrorAction Stop; if ($org.DisplayName) { $tenantName = $org.DisplayName } elseif ($org.Name) { $tenantName = $org.Name } } catch {}
                }
            } catch {}

            if (-not $tenantName -or [string]::IsNullOrWhiteSpace($tenantName)) { $tenantName = 'Tenant' }

            # Sanitize folder name
            $invalid = [System.IO.Path]::GetInvalidFileNameChars()
            $safeName = ($tenantName.ToCharArray() | ForEach-Object { if ($invalid -contains $_) { '-' } else { $_ } }) -join ''
            $safeName = ($safeName -replace '\s+', ' ').Trim()
            if ($safeName.Length -gt 80) { $safeName = $safeName.Substring(0,80) }

            $tenantRoot = Join-Path $root $safeName
            $ts   = Get-Date -Format "yyyyMMdd_HHmmss"
            $OutputFolder = Join-Path $tenantRoot $ts
        }
        if (-not (Test-Path $OutputFolder)) { New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null }
        $report.OutputFolder = $OutputFolder
    } catch {}

    # Check connections (robust detection outside UI context)
    $exchangeConnected = $false
    try {
        # Lightweight call; succeeds only when connected to EXO
        Get-OrganizationConfig -ErrorAction Stop | Out-Null
        $exchangeConnected = $true
    } catch {
        # Fallback to UI flag if present
        if (Get-Variable -Name currentExchangeConnection -Scope Script -ErrorAction SilentlyContinue) {
            $exchangeConnected = ($script:currentExchangeConnection -eq $true)
        }
    }

    $graphConnected = $false
    try {
        $mgCtx = Get-MgContext -ErrorAction Stop
        if ($mgCtx -and $mgCtx.Account) { $graphConnected = $true }
    } catch {
        # Fallback to legacy/global flag if present
        if (Get-Variable -Name graphConnection -Scope Global -ErrorAction SilentlyContinue) {
            $graphConnected = ($global:graphConnection -ne $null)
        } elseif (Get-Variable -Name graphConnection -Scope Script -ErrorAction SilentlyContinue) {
            $graphConnected = ($script:graphConnection -ne $null)
        }
    }

    if (-not $exchangeConnected) {
        Write-Warning "Exchange Online connection required for complete analysis"
        $report.ExchangeConnection = "Not Connected"
    } else {
        $report.ExchangeConnection = "Connected"
    }

    if (-not $graphConnected) {
        Write-Warning "Microsoft Graph connection required for complete analysis"
        $report.GraphConnection = "Not Connected"
    } else {
        $report.GraphConnection = "Connected"
    }

    # Collect data from Exchange Online
    if ($exchangeConnected) {
        try {
            if ($IncludeMessageTrace) {
                if ($StatusLabel -and $StatusLabel.GetType().Name -eq "Label") { $StatusLabel.Text = "Collecting message trace data (last $DaysBack days)..." }
                $report.MessageTrace = Get-ExchangeMessageTrace -DaysBack 10 -SelectedUsers $SelectedUsers # always 10 days per requirement
            }

            if ($IncludeInboxRules) {
                if ($StatusLabel -and $StatusLabel.GetType().Name -eq "Label") { $StatusLabel.Text = "Exporting inbox rules..." }
                $report.InboxRules = Get-ExchangeInboxRules -SelectedUsers $SelectedUsers
            }

            if ($IncludeTransportRules) {
                if ($StatusLabel -and $StatusLabel.GetType().Name -eq "Label") { $StatusLabel.Text = "Collecting transport rules..." }
                $report.TransportRules = Get-ExchangeTransportRules
            }

            if ($IncludeMailFlowConnectors) {
                if ($StatusLabel -and $StatusLabel.GetType().Name -eq "Label") { $StatusLabel.Text = "Collecting mail flow connectors..." }
                $report.MailFlowConnectors = Get-MailFlowConnectors
            }

            if ($IncludeMailboxForwarding) {
                if ($StatusLabel -and $StatusLabel.GetType().Name -eq "Label") { $StatusLabel.Text = "Collecting mailbox forwarding and delegation..." }
                $report.MailboxForwarding = Get-MailboxForwardingAndDelegation -SelectedUsers $SelectedUsers
            }
        } catch {
            Write-Warning "Failed to collect Exchange Online data: $($_.Exception.Message)"
            $report.ExchangeDataError = $_.Exception.Message
        }
    }

    # Collect data from Microsoft Graph (audit logs and sign-in logs)
    if ($graphConnected) {
        try {
            if ($IncludeAuditLogs) {
                if ($StatusLabel -and $StatusLabel.GetType().Name -eq "Label") { $StatusLabel.Text = "Collecting audit logs from Microsoft Graph..." }
                $report.AuditLogs = Get-GraphAuditLogs -DaysBack $DaysBack -SelectedUsers $SelectedUsers
            }

            if ($IncludeSignInLogs) {
                try {
                    if ($StatusLabel -and $StatusLabel.GetType().Name -eq "Label") { $StatusLabel.Text = "Collecting sign-in logs (last $SignInLogsDaysBack days)... This requires Azure AD Premium license." }
                    $report.SignInLogs = Get-GraphSignInLogs -DaysBack $SignInLogsDaysBack -SelectedUsers $SelectedUsers
                    if ($report.SignInLogs -and $report.SignInLogs.Count -gt 0) {
                        Write-Host "Collected $($report.SignInLogs.Count) sign-in log entries" -ForegroundColor Green
                    }
                } catch {
                    if ($_.Exception.Message -like "*insufficient privileges*" -or $_.Exception.Message -like "*permission*" -or $_.Exception.Message -like "*access denied*") {
                        Write-Warning "Insufficient permissions to read sign-in logs. Requires 'AuditLog.Read.All' permission."
                        $report.SignInLogs = @()
                        $report.SignInLogsError = "Permission denied - requires AuditLog.Read.All"
                    } elseif ($_.Exception.Message -like "*license*" -or $_.Exception.Message -like "*subscription*" -or $_.Exception.Message -like "*premium*") {
                        Write-Warning "Sign-in logs require Azure AD Premium P1 or P2 license. Free tenants can only access last 7 days."
                        $report.SignInLogs = @()
                        $report.SignInLogsError = "License required - Azure AD Premium P1/P2 (free tenants limited to 7 days)"
                    } else {
                        Write-Warning "Failed to collect sign-in logs: $($_.Exception.Message)"
                        $report.SignInLogs = @()
                        $report.SignInLogsError = $_.Exception.Message
                    }
                }
            } else {
                $report.SignInLogs = @()
            }
        } catch {
            Write-Warning "Failed to collect Microsoft Graph data: $($_.Exception.Message)"
            $report.GraphDataError = $_.Exception.Message
        }
    } else {
        $report.SignInLogs = @()
    }

    # MFA Coverage and User Security Groups
    if ($graphConnected) {
        try {
            if ($StatusLabel -and $StatusLabel.GetType().Name -eq "Label") { $StatusLabel.Text = "Evaluating MFA coverage (Security Defaults / CA / Per-user)..." }
            $report.MfaCoverage = Get-MfaCoverageReport -SelectedUsers $SelectedUsers

            if ($StatusLabel -and $StatusLabel.GetType().Name -eq "Label") { $StatusLabel.Text = "Collecting user security groups and roles..." }
            $report.UserSecurityGroups = Get-UserSecurityGroupsReport -SelectedUsers $SelectedUsers
        } catch {
            Write-Warning "Failed to build MFA/Groups reports: $($_.Exception.Message)"
        }
    }

    # Conditional Access Policies and App Registrations
    if ($graphConnected -and ($IncludeConditionalAccessPolicies -or $IncludeAppRegistrations)) {
        try {
            # Import SecurityAnalysis module if available
            $securityAnalysisModule = Join-Path $PSScriptRoot "SecurityAnalysis.psm1"
            if (Test-Path $securityAnalysisModule) {
                Import-Module $securityAnalysisModule -Force -ErrorAction SilentlyContinue
                
                # Collect Conditional Access Policies
                if ($IncludeConditionalAccessPolicies) {
                    try {
                        if ($StatusLabel -and $StatusLabel.GetType().Name -eq "Label") { $StatusLabel.Text = "Collecting Conditional Access policies..." }
                        $report.ConditionalAccessPolicies = Get-ConditionalAccessPolicies -ErrorAction Stop
                        Write-Host "Collected $($report.ConditionalAccessPolicies.Count) Conditional Access policies" -ForegroundColor Green
                    } catch {
                        if ($_.Exception.Message -like "*insufficient privileges*" -or $_.Exception.Message -like "*permission*" -or $_.Exception.Message -like "*access denied*") {
                            Write-Warning "Insufficient permissions to read Conditional Access policies. Requires 'Policy.Read.All' permission."
                            $report.ConditionalAccessPolicies = @()
                            $report.CAPoliciesError = "Permission denied - requires Policy.Read.All"
                        } elseif ($_.Exception.Message -like "*license*" -or $_.Exception.Message -like "*subscription*") {
                            Write-Warning "Conditional Access requires Azure AD Premium P1 license."
                            $report.ConditionalAccessPolicies = @()
                            $report.CAPoliciesError = "License required - Azure AD Premium P1"
                        } else {
                            Write-Warning "Failed to collect Conditional Access policies: $($_.Exception.Message)"
                            $report.ConditionalAccessPolicies = @()
                            $report.CAPoliciesError = $_.Exception.Message
                        }
                    }
                } else {
                    $report.ConditionalAccessPolicies = @()
                }
                
                # Collect App Registrations
                if ($IncludeAppRegistrations) {
                    try {
                        if ($StatusLabel -and $StatusLabel.GetType().Name -eq "Label") { $StatusLabel.Text = "Collecting app registrations..." }
                        $report.AppRegistrations = Get-AppRegistrations -ErrorAction Stop
                        Write-Host "Collected $($report.AppRegistrations.Count) app registrations" -ForegroundColor Green
                    } catch {
                        if ($_.Exception.Message -like "*insufficient privileges*" -or $_.Exception.Message -like "*permission*" -or $_.Exception.Message -like "*access denied*") {
                            Write-Warning "Insufficient permissions to read app registrations. Requires 'Application.Read.All' permission."
                            $report.AppRegistrations = @()
                            $report.AppRegistrationsError = "Permission denied - requires Application.Read.All"
                        } else {
                            Write-Warning "Failed to collect app registrations: $($_.Exception.Message)"
                            $report.AppRegistrations = @()
                            $report.AppRegistrationsError = $_.Exception.Message
                        }
                    }
                } else {
                    $report.AppRegistrations = @()
                }
            } else {
                Write-Warning "SecurityAnalysis module not found. CA Policies and App Registrations will not be collected."
                if ($IncludeConditionalAccessPolicies) { $report.ConditionalAccessPolicies = @() }
                if ($IncludeAppRegistrations) { $report.AppRegistrations = @() }
            }
        } catch {
            Write-Warning "Failed to collect CA Policies or App Registrations: $($_.Exception.Message)"
            if ($IncludeConditionalAccessPolicies -and -not $report.ConditionalAccessPolicies) { $report.ConditionalAccessPolicies = @() }
            if ($IncludeAppRegistrations -and -not $report.AppRegistrations) { $report.AppRegistrations = @() }
        }
    } else {
        # Both are disabled or not connected, set empty arrays
        $report.ConditionalAccessPolicies = @()
        $report.AppRegistrations = @()
    }

    # Generate AI Investigation Prompt
    if ($StatusLabel -and $StatusLabel.GetType().Name -eq "Label") { $StatusLabel.Text = "Generating AI investigation prompts..." }
    $report.AIPrompt = New-AISecurityInvestigationPrompt -Report $report

    # Generate Ticketing System Message
    if ($StatusLabel -and $StatusLabel.GetType().Name -eq "Label") { $StatusLabel.Text = "Generating non-technical incident summary..." }
    $report.TicketMessage = New-TicketSecuritySummary -Report $report

    # Generate comprehensive report
    $report.Summary = New-SecurityInvestigationSummary -Report $report

    if ($StatusLabel -and $StatusLabel.GetType().Name -eq "Label") { $StatusLabel.Text = "Security investigation report completed" }
    if ($MainForm -and $MainForm.GetType().Name -eq "Form") { $MainForm.Cursor = [System.Windows.Forms.Cursors]::Default }

    # Export datasets to CSV (and JSON fallback) if we have an output folder
    # Note: Data is already filtered server-side by the collection functions
    if ($report.OutputFolder) {
        $exportError = $null
        try {
            $csv = Join-Path $report.OutputFolder "MessageTrace.csv"
            $json = Join-Path $report.OutputFolder "MessageTrace.json"
            if ($report.MessageTrace -and $report.MessageTrace.Count -gt 0) {
                try { $report.MessageTrace | Export-Csv -Path $csv -NoTypeInformation -Encoding UTF8; $report.FilePaths.MessageTraceCsv = $csv }
                catch { $report.MessageTrace | ConvertTo-Json -Depth 8 | Out-File -FilePath $json -Encoding utf8; $report.FilePaths.MessageTraceJson = $json }
            }

            $csv = Join-Path $report.OutputFolder "InboxRules.csv"
            $json = Join-Path $report.OutputFolder "InboxRules.json"
            if ($report.InboxRules -and $report.InboxRules.Count -gt 0) {
                try { $report.InboxRules | Export-Csv -Path $csv -NoTypeInformation -Encoding UTF8; $report.FilePaths.InboxRulesCsv = $csv }
                catch { $report.InboxRules | ConvertTo-Json -Depth 6 | Out-File -FilePath $json -Encoding utf8; $report.FilePaths.InboxRulesJson = $json }
            }

            # Transport Rules export
            $csv = Join-Path $report.OutputFolder "TransportRules.csv"
            $json = Join-Path $report.OutputFolder "TransportRules.json"
            if ($report.TransportRules -and $report.TransportRules.Count -gt 0) {
                try { $report.TransportRules | Export-Csv -Path $csv -NoTypeInformation -Encoding UTF8; $report.FilePaths.TransportRulesCsv = $csv }
                catch { $report.TransportRules | ConvertTo-Json -Depth 8 | Out-File -FilePath $json -Encoding utf8; $report.FilePaths.TransportRulesJson = $json }
            }

            # Mail Flow Connectors export (combined Inbound + Outbound)
            $csv = Join-Path $report.OutputFolder "MailFlowConnectors.csv"
            $json = Join-Path $report.OutputFolder "MailFlowConnectors.json"
            if ($report.MailFlowConnectors -and $report.MailFlowConnectors.Count -gt 0) {
                try { $report.MailFlowConnectors | Export-Csv -Path $csv -NoTypeInformation -Encoding UTF8; $report.FilePaths.MailFlowConnectorsCsv = $csv }
                catch { $report.MailFlowConnectors | ConvertTo-Json -Depth 8 | Out-File -FilePath $json -Encoding utf8; $report.FilePaths.MailFlowConnectorsJson = $json }
            }

            $csv = Join-Path $report.OutputFolder "GraphAuditLogs.csv"
            $json = Join-Path $report.OutputFolder "GraphAuditLogs.json"
            if ($report.AuditLogs -and $report.AuditLogs.Count -gt 0) {
                try { $report.AuditLogs | Export-Csv -Path $csv -NoTypeInformation -Encoding UTF8; $report.FilePaths.AuditLogsCsv = $csv }
                catch { $report.AuditLogs | ConvertTo-Json -Depth 8 | Out-File -FilePath $json -Encoding utf8; $report.FilePaths.AuditLogsJson = $json }
            }

            # Sign-in Logs export
            $csv = Join-Path $report.OutputFolder "SignInLogs.csv"
            $json = Join-Path $report.OutputFolder "SignInLogs.json"
            if ($report.SignInLogs -and $report.SignInLogs.Count -gt 0) {
                try { 
                    $report.SignInLogs | Export-Csv -Path $csv -NoTypeInformation -Encoding UTF8
                    $report.FilePaths.SignInLogsCsv = $csv
                    Write-Host "Exported $($report.SignInLogs.Count) sign-in log entries to SignInLogs.csv" -ForegroundColor Green
                }
                catch { 
                    $report.SignInLogs | ConvertTo-Json -Depth 8 | Out-File -FilePath $json -Encoding utf8
                    $report.FilePaths.SignInLogsJson = $json
                    Write-Warning "Failed to export sign-in logs to CSV, exported to JSON instead"
                }
            } elseif ($report.SignInLogsError) {
                # Write error to a text file
                $errorFile = Join-Path $report.OutputFolder "SignInLogs_Error.txt"
                "Error collecting Sign-in Logs:`n$($report.SignInLogsError)`n`nNote: Sign-in logs require Azure AD Premium P1 or P2 license. Free tenants are limited to 7 days of data." | Out-File -FilePath $errorFile -Encoding utf8
                $report.FilePaths.SignInLogsError = $errorFile
            }

            # Conditional Access Policies export
            $csv = Join-Path $report.OutputFolder "ConditionalAccessPolicies.csv"
            $json = Join-Path $report.OutputFolder "ConditionalAccessPolicies.json"
            if ($report.ConditionalAccessPolicies -and $report.ConditionalAccessPolicies.Count -gt 0) {
                try {
                    # Flatten CA policies for CSV export
                    $caPoliciesFlat = $report.ConditionalAccessPolicies | ForEach-Object {
                        [PSCustomObject]@{
                            PolicyId = $_.Id
                            DisplayName = $_.DisplayName
                            State = $_.State
                            CreatedDateTime = $_.CreatedDateTime
                            ModifiedDateTime = $_.ModifiedDateTime
                            RiskLevel = $_.RiskLevel
                            RiskScore = $_.Analysis.RiskScore
                            IsEnabled = $_.Analysis.IsEnabled
                            HasSuspiciousConditions = $_.Analysis.HasSuspiciousConditions
                            HasSuspiciousControls = $_.Analysis.HasSuspiciousControls
                            SuspiciousIndicators = ($_.Analysis.SuspiciousIndicators -join "; ")
                            UserIncludeAll = if ($_.Conditions.Users.IncludeUsers) { $_.Conditions.Users.IncludeUsers -contains "All" } else { $false }
                            UserExcludeCount = if ($_.Conditions.Users.ExcludeUsers) { $_.Conditions.Users.ExcludeUsers.Count } else { 0 }
                            LocationIncludeAll = if ($_.Conditions.Locations.IncludeLocations) { $_.Conditions.Locations.IncludeLocations -contains "All" } else { $false }
                            RequiresMfa = if ($_.GrantControls.BuiltInControls) { $_.GrantControls.BuiltInControls -contains "mfa" } else { $false }
                            RequiresCompliantDevice = if ($_.GrantControls.BuiltInControls) { $_.GrantControls.BuiltInControls -contains "compliantDevice" } else { $false }
                            RequiresHybridDevice = if ($_.GrantControls.BuiltInControls) { $_.GrantControls.BuiltInControls -contains "domainJoinedDevice" } else { $false }
                            SignInFrequencyHours = if ($_.SessionControls.SignInFrequency) { $_.SessionControls.SignInFrequency.Value } else { $null }
                        }
                    }
                    $caPoliciesFlat | Export-Csv -Path $csv -NoTypeInformation -Encoding UTF8
                    $report.FilePaths.ConditionalAccessPoliciesCsv = $csv
                } catch {
                    Write-Warning "Failed to export CA Policies to CSV: $($_.Exception.Message)"
                    try {
                        $report.ConditionalAccessPolicies | ConvertTo-Json -Depth 10 | Out-File -FilePath $json -Encoding utf8
                        $report.FilePaths.ConditionalAccessPoliciesJson = $json
                    } catch {
                        Write-Warning "Failed to export CA Policies to JSON: $($_.Exception.Message)"
                    }
                }
            } elseif ($report.CAPoliciesError) {
                # Write error to a text file
                $errorFile = Join-Path $report.OutputFolder "ConditionalAccessPolicies_Error.txt"
                "Error collecting Conditional Access Policies:`n$($report.CAPoliciesError)" | Out-File -FilePath $errorFile -Encoding utf8
                $report.FilePaths.ConditionalAccessPoliciesError = $errorFile
            }

            # App Registrations export
            $csv = Join-Path $report.OutputFolder "AppRegistrations.csv"
            $json = Join-Path $report.OutputFolder "AppRegistrations.json"
            if ($report.AppRegistrations -and $report.AppRegistrations.Count -gt 0) {
                try {
                    # Flatten app registrations for CSV export
                    $appRegsFlat = $report.AppRegistrations | ForEach-Object {
                        [PSCustomObject]@{
                            AppId = $_.AppId
                            DisplayName = $_.DisplayName
                            PublisherDomain = $_.PublisherDomain
                            CreatedDateTime = $_.CreatedDateTime
                            RiskLevel = $_.RiskLevel
                            RiskScore = $_.Analysis.RiskScore
                            HasHighPrivilegePermissions = $_.Analysis.HasHighPrivilegePermissions
                            HasSuspiciousPermissions = $_.Analysis.HasSuspiciousPermissions
                            HasUserConsent = $_.Analysis.HasUserConsent
                            SuspiciousIndicators = ($_.Analysis.SuspiciousIndicators -join "; ")
                            RequiredPermissions = ($_.RequiredPermissions -join "; ")
                            HasCertificates = $_.HasCertificates
                            HasPasswordCredentials = $_.HasPasswordCredentials
                            WebRedirectUris = ($_.WebRedirectUris -join "; ")
                            ServicePrincipalId = if ($_.ServicePrincipal) { $_.ServicePrincipal.Id } else { $null }
                            ServicePrincipalType = if ($_.ServicePrincipal) { $_.ServicePrincipal.ServicePrincipalType } else { $null }
                        }
                    }
                    $appRegsFlat | Export-Csv -Path $csv -NoTypeInformation -Encoding UTF8
                    $report.FilePaths.AppRegistrationsCsv = $csv
                } catch {
                    Write-Warning "Failed to export App Registrations to CSV: $($_.Exception.Message)"
                    try {
                        $report.AppRegistrations | ConvertTo-Json -Depth 10 | Out-File -FilePath $json -Encoding utf8
                        $report.FilePaths.AppRegistrationsJson = $json
                    } catch {
                        Write-Warning "Failed to export App Registrations to JSON: $($_.Exception.Message)"
                    }
                }
            } elseif ($report.AppRegistrationsError) {
                # Write error to a text file
                $errorFile = Join-Path $report.OutputFolder "AppRegistrations_Error.txt"
                "Error collecting App Registrations:`n$($report.AppRegistrationsError)" | Out-File -FilePath $errorFile -Encoding utf8
                $report.FilePaths.AppRegistrationsError = $errorFile
            }

            # User Security Posture export (combined MFA + Groups + Mailbox Forwarding/Delegation)
            try {
                $userPosture = New-Object System.Collections.Generic.List[object]

                # Create lookup dictionary for mailbox forwarding/delegation by UPN
                $mbxLookup = @{}
                if ($report.MailboxForwarding) {
                    foreach ($mbx in $report.MailboxForwarding) {
                        $mbxLookup[$mbx.UserPrincipalName] = $mbx
                    }
                }

                # Create lookup dictionary for user groups by UPN
                $groupsLookup = @{}
                if ($report.UserSecurityGroups) {
                    foreach ($userGroup in $report.UserSecurityGroups) {
                        $groupsLookup[$userGroup.UserPrincipalName] = $userGroup.GroupsAndRoles
                    }
                }

                # Determine which users to include in the export
                $usersToExport = @()
                if ($SelectedUsers -and $SelectedUsers.Count -gt 0) {
                    # Filter to selected users only
                    $selectedUserSet = @{}
                    foreach ($user in $SelectedUsers) {
                        if ($user -is [string]) {
                            $selectedUserSet[$user.ToLower()] = $user
                        } elseif ($user.UserPrincipalName) {
                            $selectedUserSet[$user.UserPrincipalName.ToLower()] = $user.UserPrincipalName
                        }
                    }
                    
                    # Build list from selected users
                    foreach ($userKey in $selectedUserSet.Keys) {
                        $upn = $selectedUserSet[$userKey]
                        # Try to find user in MFA coverage, groups, or mailbox forwarding
                        $found = $false
                        if ($report.MfaCoverage -and $report.MfaCoverage.Users) {
                            $mfaUser = $report.MfaCoverage.Users | Where-Object { $_.UserPrincipalName -eq $upn -or $_.UserPrincipalName.ToLower() -eq $upn.ToLower() } | Select-Object -First 1
                            if ($mfaUser) {
                                $usersToExport += $mfaUser
                                $found = $true
                            }
                        }
                        if (-not $found) {
                            # Create a basic user object if not found in MFA coverage
                            $usersToExport += [pscustomobject]@{
                                UserPrincipalName = $upn
                                DisplayName = $upn
                                PerUserMfaEnabled = $false
                                SecurityDefaults = $false
                                CARequiresMfa = $false
                                MfaCovered = $false
                            }
                        }
                    }
                } else {
                    # No selection - use all MFA coverage users
                    if ($report.MfaCoverage -and $report.MfaCoverage.Users) {
                        $usersToExport = $report.MfaCoverage.Users
                    }
                }

                # Import EntraInvestigator module for per-user MFA status
                $entraModuleLoaded = $false
                try {
                    $entraModulePath = Join-Path $PSScriptRoot 'EntraInvestigator.psm1'
                    if (Test-Path $entraModulePath) {
                        Import-Module $entraModulePath -Force -ErrorAction SilentlyContinue
                        $entraModuleLoaded = $true
                    }
                } catch {}

                # Build user posture for each user
                foreach ($mfaUser in $usersToExport) {
                    $upn = $mfaUser.UserPrincipalName
                    $mbxData = $mbxLookup[$upn]
                    $groupsData = $groupsLookup[$upn]

                    # Get detailed per-user MFA status if module is available
                    $perUserMfaStatus = $null
                    $perUserMfaDetails = $null
                    $perUserMfaOverallStatus = $null
                    $perUserMfaSummary = $null
                    if ($entraModuleLoaded -and $graphConnected) {
                        try {
                            $mfaStatus = Get-EntraUserMfaStatus -UserPrincipalName $upn -ErrorAction SilentlyContinue
                            if ($mfaStatus) {
                                $perUserMfaStatus = $mfaStatus.PerUserMfa.Enabled
                                $perUserMfaDetails = $mfaStatus.PerUserMfa.Details
                                $perUserMfaOverallStatus = $mfaStatus.OverallStatus
                                $perUserMfaSummary = $mfaStatus.Summary
                            }
                        } catch {
                            # Ignore errors getting per-user MFA status
                        }
                    }

                    $userPosture.Add([pscustomobject]@{
                        UserPrincipalName           = $upn
                        DisplayName                 = $mfaUser.DisplayName
                        RecipientType               = if ($mbxData) { $mbxData.RecipientType } else { $null }
                        # MFA columns (from organization-wide MFA coverage)
                        PerUserMfaEnabled           = $mfaUser.PerUserMfaEnabled
                        SecurityDefaults            = $mfaUser.SecurityDefaults
                        CARequiresMfa               = $mfaUser.CARequiresMfa
                        MfaCovered                  = $mfaUser.MfaCovered
                        # Per-user detailed MFA status
                        PerUserMfaStatus            = $perUserMfaStatus
                        PerUserMfaDetails           = $perUserMfaDetails
                        PerUserMfaOverallStatus     = $perUserMfaOverallStatus
                        PerUserMfaSummary           = $perUserMfaSummary
                        # Groups/Roles
                        GroupsAndRoles              = $groupsData
                        # Mailbox Forwarding
                        ForwardingAddress           = if ($mbxData) { $mbxData.ForwardingAddress } else { $null }
                        ForwardingSmtpAddress       = if ($mbxData) { $mbxData.ForwardingSmtpAddress } else { $null }
                        DeliverToMailboxAndForward  = if ($mbxData) { $mbxData.DeliverToMailboxAndForward } else { $null }
                        # Delegation
                        FullAccessUsers             = if ($mbxData) { $mbxData.FullAccessUsers } else { $null }
                        SendAsUsers                 = if ($mbxData) { $mbxData.SendAsUsers } else { $null }
                        SendOnBehalfUsers           = if ($mbxData) { $mbxData.SendOnBehalfUsers } else { $null }
                    }) | Out-Null
                }

                # Export combined user security posture
                if ($userPosture.Count -gt 0) {
                    $csv = Join-Path $report.OutputFolder "UserSecurityPosture.csv"
                    try { $userPosture | Export-Csv -Path $csv -NoTypeInformation -Encoding UTF8; $report.FilePaths.UserSecurityPostureCsv = $csv } catch {}
                }
            } catch {
                Write-Warning "Failed to create UserSecurityPosture export: $($_.Exception.Message)"
            }
        } catch { $exportError = $_ }

        # Save only LLM instructions as TXT (no other text files on disk)
        try {
            $report.LLMInstructions = New-LLMInvestigationInstructions -Report $report
            $llmPath = Join-Path $report.OutputFolder "_AI_Readme.txt"
            if ($report.LLMInstructions) { $report.LLMInstructions | Out-File -FilePath $llmPath -Encoding utf8 }
            $report.FilePaths.LLMInstructionsTxt = $llmPath
        } catch {}

        # Automatically create zip file of all reports (excluding _AI_Readme.txt)
        try {
            if ($StatusLabel -and $StatusLabel.GetType().Name -eq "Label") {
                $StatusLabel.Text = "Creating zip archive of security reports..."
            }
            $zipPath = New-SecurityInvestigationZip -OutputFolder $report.OutputFolder
            if ($zipPath) {
                $report.FilePaths.ZipFile = $zipPath
                Write-Host "Zip archive created: $zipPath" -ForegroundColor Green
            }
        } catch {
            Write-Warning "Failed to create zip archive: $($_.Exception.Message)"
        }
    }

    return $report
}

function Get-ExchangeMessageTrace {
    param(
        [int]$DaysBack = 10,
        [Parameter(Mandatory=$false)]
        [array]$SelectedUsers = @()
    )

    try {
        Write-Host "Collecting message trace data..." -ForegroundColor Yellow
        $end = (Get-Date).ToUniversalTime()
        $start = $end.AddDays(-10).Date # always 10 full days; start at 00:00Z

        $results = New-Object System.Collections.Generic.List[object]

        $hasV2 = $null -ne (Get-Command Get-MessageTraceV2 -ErrorAction SilentlyContinue)

        # If SelectedUsers provided, filter server-side by querying per user
        if ($SelectedUsers -and $SelectedUsers.Count -gt 0) {
            $selectedUserList = @()
            foreach ($user in $SelectedUsers) {
                $upn = if ($user -is [string]) { $user } elseif ($user.UserPrincipalName) { $user.UserPrincipalName } else { continue }
                $selectedUserList += $upn
            }
            
            # Query message trace for each selected user (server-side filtering)
            foreach ($upn in $selectedUserList) {
                # Chunk by day to avoid server-side caps
                for ($d = 0; $d -lt 10; $d++) {
                    $winStart = $start.AddDays($d)
                    $winEnd   = $winStart.AddDays(1)

                    try {
                        if ($hasV2) {
                            # Query by sender - handle pagination
                            try {
                                $params = @{ StartDate = $winStart; EndDate = $winEnd; SenderAddress = $upn; ErrorAction = 'Stop' }
                                $params.ResultSize = 5000  # Use ResultSize for pagination
                                $chunk = Get-MessageTraceV2 @params
                                if ($chunk) {
                                    # Handle both single objects and collections
                                    $chunkArray = @($chunk)
                                    foreach ($item in $chunkArray) {
                                        if ($item) { [void]$results.Add($item) }
                                    }
                                }
                            } catch {
                                Write-Warning "Failed to get message trace by sender for ${upn}: $($_.Exception.Message)"
                            }
                            
                            # Query by recipient - handle pagination
                            try {
                                $params = @{ StartDate = $winStart; EndDate = $winEnd; RecipientAddress = $upn; ErrorAction = 'Stop' }
                                $params.ResultSize = 5000  # Use ResultSize for pagination
                                $chunk = Get-MessageTraceV2 @params
                                if ($chunk) {
                                    # Handle both single objects and collections
                                    $chunkArray = @($chunk)
                                    foreach ($item in $chunkArray) {
                                        if ($item) { [void]$results.Add($item) }
                                    }
                                }
                            } catch {
                                Write-Warning "Failed to get message trace by recipient for ${upn}: $($_.Exception.Message)"
                            }
                        } else {
                            # Legacy Get-MessageTrace - filter by sender
                            try {
                                $batch = Get-MessageTrace -StartDate $winStart -EndDate $winEnd -SenderAddress $upn -ErrorAction Stop
                                if ($batch) {
                                    $batchArray = @($batch)
                                    foreach ($item in $batchArray) {
                                        if ($item) { [void]$results.Add($item) }
                                    }
                                }
                            } catch {
                                Write-Warning "Failed to get message trace by sender for ${upn}: $($_.Exception.Message)"
                            }
                            
                            # Filter by recipient
                            try {
                                $batch = Get-MessageTrace -StartDate $winStart -EndDate $winEnd -RecipientAddress $upn -ErrorAction Stop
                                if ($batch) {
                                    $batchArray = @($batch)
                                    foreach ($item in $batchArray) {
                                        if ($item) { [void]$results.Add($item) }
                                    }
                                }
                            } catch {
                                Write-Warning "Failed to get message trace by recipient for ${upn}: $($_.Exception.Message)"
                            }
                        }
                    } catch {
                        Write-Warning "Failed to get message trace for ${upn}: $($_.Exception.Message)"
                    }
                }
            }
            
            # Remove duplicates (same message might appear as both sender and recipient)
            # Use MessageId if available, otherwise use a combination of properties
            $uniqueResults = @()
            $seenIds = @{}
            foreach ($item in $results) {
                $uniqueKey = $null
                if ($item.MessageId) {
                    $uniqueKey = $item.MessageId
                } elseif ($item.MessageID) {
                    $uniqueKey = $item.MessageID
                } else {
                    # Fallback: use combination of properties
                    $uniqueKey = "$($item.SenderAddress)_$($item.RecipientAddress)_$($item.Subject)_$($item.MessageTraceId)"
                }
                
                if ($uniqueKey -and -not $seenIds.ContainsKey($uniqueKey)) {
                    $seenIds[$uniqueKey] = $true
                    $uniqueResults += $item
                }
            }
            
            return [System.Collections.ArrayList]$uniqueResults
        } else {
            # No selection - get all message traces
            # Chunk by day to avoid server-side caps; try paged in each window
            for ($d = 0; $d -lt 10; $d++) {
                $winStart = $start.AddDays($d)
                $winEnd   = $winStart.AddDays(1)

                try {
                    if ($hasV2) {
                        # Seek-based pagination using StartingRecipientAddress and ResultSize
                        $startRecipient = $null
                        $iterations = 0
                        do {
                            $params = @{ StartDate = $winStart; EndDate = $winEnd; ErrorAction = 'Stop' }
                            $params.ResultSize = 1000
                            if ($startRecipient) { $params.StartingRecipientAddress = $startRecipient }
                            $chunk = Get-MessageTraceV2 @params
                            if ($chunk) {
                                # Avoid duplicate loops when StartingRecipientAddress is inclusive
                                if ($startRecipient) {
                                    $filtered = $chunk | Where-Object { $_.RecipientAddress -gt $startRecipient }
                                } else {
                                    $filtered = $chunk
                                }
                                if ($filtered) { [void]$results.AddRange($filtered) }

                                $prev = $startRecipient
                                $last = $chunk[-1]
                                $startRecipient = $last.RecipientAddress
                                if (-not $startRecipient -or ($prev -and $startRecipient -le $prev)) { break }
                            } else {
                                $startRecipient = $null
                            }
                            $iterations++
                        } while ($chunk -and $chunk.Count -eq 1000 -and $startRecipient -and $iterations -lt 500)
                    } else {
                        $batch = Get-MessageTrace -StartDate $winStart -EndDate $winEnd -ErrorAction Stop
                        if ($batch) { [void]$results.AddRange($batch) }
                    }
                } catch {}
            }
        }

        return [System.Collections.ArrayList]$results
    } catch {
        Write-Error "Failed to collect message trace: $($_.Exception.Message)"
        return @()
    }
}

function Get-ExchangeInboxRules {
    param(
        [Parameter(Mandatory=$false)]
        [array]$SelectedUsers = @()
    )
    
    try {
        Write-Host "Exporting inbox rules..." -ForegroundColor Yellow

        $mailboxes = @()
        
        # If SelectedUsers provided, only query those mailboxes (server-side filtering)
        if ($SelectedUsers -and $SelectedUsers.Count -gt 0) {
            foreach ($user in $SelectedUsers) {
                $upn = if ($user -is [string]) { $user } elseif ($user.UserPrincipalName) { $user.UserPrincipalName } else { continue }
                try {
                    $mbx = Get-Mailbox -Identity $upn -ErrorAction Stop
                    if ($mbx) { $mailboxes += $mbx }
                } catch {
                    Write-Warning "Mailbox not found for ${upn}: $($_.Exception.Message)"
                }
            }
        } else {
            # No selection - get all mailboxes
            try {
                $mailboxes = Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox,SharedMailbox -ErrorAction Stop
            } catch {
                # Fallback narrower call if needed
                $mailboxes = Get-Mailbox -ResultSize 2000 -ErrorAction Stop
            }
        }

        $allRules = New-Object System.Collections.Generic.List[object]
        foreach ($mbx in $mailboxes) {
            $upn = if ($mbx.UserPrincipalName) { $mbx.UserPrincipalName } else { $mbx.PrimarySmtpAddress }
            try {
                $rules = Get-InboxRule -Mailbox $upn -ErrorAction Stop
                foreach ($r in $rules) {
                    $obj = [pscustomobject]@{
                        MailboxOwner        = $upn
                        Name                = $r.Name
                        Enabled             = $r.Enabled
                        Priority            = $r.Priority
                        FromAddressContains = ($r.FromAddressContainsWords -join ';')
                        SubjectContains     = ($r.SubjectContainsWords -join ';')
                        SentTo              = ($r.SentTo -join ';')
                        RedirectTo          = ($r.RedirectTo -join ';')
                        ForwardTo           = ($r.ForwardTo -join ';')
                        ForwardAsAttachment = ($r.ForwardAsAttachmentTo -join ';')
                        DeleteMessage       = $r.DeleteMessage
                        StopProcessing      = $r.StopProcessingRules
                        IsHidden            = $false
                        Description         = ($r.Description -join ' ')
                    }
                    [void]$allRules.Add($obj)
                }
            } catch {
                Write-Warning "Get-InboxRule failed for ${upn}: $($_.Exception.Message)"
            }
        }

        return [System.Collections.ArrayList]$allRules
    } catch {
        Write-Error "Failed to export inbox rules: $($_.Exception.Message)"
        return @()
    }
}

function Get-ExchangeTransportRules {
    try {
        Write-Host "Exporting transport (mail flow) rules..." -ForegroundColor Yellow
        $rules = @()
        try { $rules = Get-TransportRule -ResultSize Unlimited -ErrorAction Stop } catch { $rules = Get-TransportRule -ErrorAction Stop }

        function Convert-ShortJson($obj) { try { return ($obj | ConvertTo-Json -Depth 12 -Compress) } catch { return "" } }

        $results = New-Object System.Collections.Generic.List[object]
        foreach ($r in $rules) {
            $results.Add([pscustomobject]@{
                Name               = $r.Name
                Priority           = $r.Priority
                State              = $r.State
                Mode               = $r.Mode
                Comments           = $r.Comments
                RuleVersion        = $r.RuleVersion
                ActivationDate     = $r.ActivationDate
                ExpiryDate         = $r.ExpiryDate
                ConditionsSummary  = (Convert-ShortJson $r.Conditions)
                ExceptionsSummary  = (Convert-ShortJson $r.Exceptions)
                ActionsSummary     = (Convert-ShortJson $r.Actions)
                ImmutableId        = $r.ImmutableId
                Guid               = $r.Guid
                DlpPolicy          = $r.DlpPolicy
            }) | Out-Null
        }
        return [System.Collections.ArrayList]$results
    } catch {
        Write-Error "Failed to export transport rules: $($_.Exception.Message)"; return @()
    }
}

function Get-ExchangeInboundConnectors {
    try {
        Write-Host "Exporting inbound connectors..." -ForegroundColor Yellow
        $conns = @()
        try {
            $params = @{ ErrorAction = 'Stop'; WarningAction = 'SilentlyContinue' }
            $gc = Get-Command Get-InboundConnector -ErrorAction SilentlyContinue
            if ($gc -and $gc.Parameters.ContainsKey('IncludeTestModeConnectors')) { $params.IncludeTestModeConnectors = $true }
            $conns = Get-InboundConnector @params
        } catch { $conns = @() }
        $results = New-Object System.Collections.Generic.List[object]
        foreach ($c in $conns) {
            $results.Add([pscustomobject]@{
                Name                          = $c.Name
                ConnectorType                 = $c.ConnectorType
                Enabled                       = $c.Enabled
                SenderDomains                 = ($c.SenderDomains -join ';')
                SenderIPAddresses             = ($c.SenderIPAddresses -join ';')
                RestrictDomainsToCertificate  = $c.RestrictDomainsToCertificate
                RestrictDomainsToIPAddresses  = $c.RestrictDomainsToIPAddresses
                TlsSenderCertificateName      = $c.TlsSenderCertificateName
                RequireTls                    = $c.RequireTls
                CloudServicesMailEnabled      = $c.CloudServicesMailEnabled
                Comment                       = $c.Comment
                Identity                      = $c.Identity
                Guid                           = $c.Guid
                TestMode                      = $(if ($c.PSObject.Properties['TestMode']) { $c.TestMode } elseif ($c.PSObject.Properties['IsTestMode']) { $c.IsTestMode } else { $null })
            }) | Out-Null
        }
        return [System.Collections.ArrayList]$results
    } catch {
        Write-Error "Failed to export inbound connectors: $($_.Exception.Message)"; return @()
    }
}

function Get-ExchangeOutboundConnectors {
    try {
        Write-Host "Exporting outbound connectors..." -ForegroundColor Yellow
        $conns = @()
        try {
            $params = @{ ErrorAction = 'Stop'; WarningAction = 'SilentlyContinue' }
            $gc = Get-Command Get-OutboundConnector -ErrorAction SilentlyContinue
            if ($gc -and $gc.Parameters.ContainsKey('IncludeTestModeConnectors')) { $params.IncludeTestModeConnectors = $true }
            $conns = Get-OutboundConnector @params
        } catch { $conns = @() }
        $results = New-Object System.Collections.Generic.List[object]
        foreach ($c in $conns) {
            $results.Add([pscustomobject]@{
                Name                     = $c.Name
                ConnectorType            = $c.ConnectorType
                Enabled                  = $c.Enabled
                SmartHosts               = ($c.SmartHosts -join ';')
                RecipientDomains         = ($c.RecipientDomains -join ';')
                UseMXRecord              = $c.UseMXRecord
                TlsSettings              = $c.TlsSettings
                TlsDomain                = $c.TlsDomain
                CloudServicesMailEnabled = $c.CloudServicesMailEnabled
                Comment                  = $c.Comment
                Identity                 = $c.Identity
                Guid                      = $c.Guid
                TestMode                 = $(if ($c.PSObject.Properties['TestMode']) { $c.TestMode } elseif ($c.PSObject.Properties['IsTestMode']) { $c.IsTestMode } else { $null })
            }) | Out-Null
        }
        return [System.Collections.ArrayList]$results
    } catch {
        Write-Error "Failed to export outbound connectors: $($_.Exception.Message)"; return @()
    }
}

function Get-GraphAuditLogs {
    param(
        [int]$DaysBack = 10,
        [Parameter(Mandatory=$false)]
        [array]$SelectedUsers = @()
    )

    try {
        Write-Host "Collecting audit logs..." -ForegroundColor Yellow
        # Ensure identity modules are available
        if (-not (Get-Command Get-MgAuditLogDirectoryAudit -ErrorAction SilentlyContinue)) {
            Import-Module Microsoft.Graph.Reports -ErrorAction SilentlyContinue | Out-Null
            Import-Module Microsoft.Graph.Authentication -ErrorAction SilentlyContinue | Out-Null
        }

        $startUtc = (Get-Date).ToUniversalTime().AddDays(-[Math]::Max(1,$DaysBack))
        $startIso = $startUtc.ToString("s") + "Z"

        $raw = New-Object System.Collections.Generic.List[object]
        
        # If SelectedUsers provided, filter server-side by target user IDs
        if ($SelectedUsers -and $SelectedUsers.Count -gt 0) {
            # Get user IDs for selected users (server-side filtering)
            $userIds = @()
            $userIdToUpn = @{}  # Map user IDs back to UPNs for better error messages
            foreach ($user in $SelectedUsers) {
                $upn = if ($user -is [string]) { $user } elseif ($user.UserPrincipalName) { $user.UserPrincipalName } else { continue }
                try {
                    Write-Host "  Looking up user ID for: $upn" -ForegroundColor Gray
                    $mgUser = Get-MgUser -UserId $upn -Property Id -ErrorAction Stop
                    if ($mgUser -and $mgUser.Id) {
                        $userIds += $mgUser.Id
                        $userIdToUpn[$mgUser.Id] = $upn
                        Write-Host "  ✓ Found user ID $($mgUser.Id) for $upn" -ForegroundColor Gray
                    } else {
                        Write-Warning "User ID lookup returned null for ${upn}"
                    }
                } catch {
                    Write-Warning "Failed to get user ID for ${upn}: $($_.Exception.Message)"
                }
            }
            
            Write-Host "  Successfully resolved $($userIds.Count) of $($SelectedUsers.Count) user IDs" -ForegroundColor Gray
            
            # Query audit logs filtered by target user IDs (server-side filtering)
            if ($userIds.Count -gt 0) {
                foreach ($userId in $userIds) {
                    $upn = $userIdToUpn[$userId]
                    try {
                        Write-Host "  Querying audit logs for: $upn (ID: $userId)" -ForegroundColor Gray
                        # Filter by target resource ID (server-side)
                        $filter = "activityDateTime ge $startIso and targetResources/any(t:t/id eq '$userId')"
                        $page = Get-MgAuditLogDirectoryAudit -All -Filter $filter -ErrorAction Stop
                        if ($page) {
                            # Handle both single objects and collections
                            $pageArray = @($page)
                            $count = 0
                            foreach ($item in $pageArray) {
                                if ($item) {
                                    [void]$raw.Add($item)
                                    $count++
                                }
                            }
                            Write-Host "  ✓ Found $count audit log entries for $upn" -ForegroundColor Gray
                        } else {
                            Write-Host "  ℹ No audit log entries found for $upn (this is normal if no activity)" -ForegroundColor Gray
                        }
                    } catch {
                        Write-Warning "Failed to get audit logs for $upn (ID: ${userId}): $($_.Exception.Message)"
                        # Don't silently fail - log the error details
                        Write-Host "  Error details: $($_.Exception.GetType().FullName)" -ForegroundColor Yellow
                        if ($_.Exception.InnerException) {
                            Write-Host "  Inner exception: $($_.Exception.InnerException.Message)" -ForegroundColor Yellow
                        }
                    }
                }
            } else {
                Write-Warning "No user IDs were successfully resolved. Cannot query audit logs."
            }
        } else {
            # No selection - get all audit logs
            $page = Get-MgAuditLogDirectoryAudit -All -Filter "activityDateTime ge $startIso" -ErrorAction Stop
            if ($page) {
                # Handle both single objects and collections
                $pageArray = @($page)
                foreach ($item in $pageArray) {
                    if ($item) {
                        [void]$raw.Add($item)
                    }
                }
            }
        }

        # Flatten for CSV detail richness
        $flattened = New-Object System.Collections.Generic.List[object]

        foreach ($r in $raw) {
            try {
                $userObj  = $r.InitiatedBy.User
                $appObj   = $r.InitiatedBy.App
                $ipAddr   = $null
                if ($userObj -and $userObj.IpAddress) { $ipAddr = $userObj.IpAddress }

                $targets = @()
                if ($r.TargetResources) {
                    foreach ($t in $r.TargetResources) {
                        $tname = $t.DisplayName
                        $tid   = $t.Id
                        $ttype = $t.Type
                        $targets += ("{0} ({1}, {2})" -f $tname,$tid,$ttype)
                    }
                }

                $modProps = @()
                if ($r.TargetResources -and $r.TargetResources[0] -and $r.TargetResources[0].ModifiedProperties) {
                    foreach ($p in $r.TargetResources[0].ModifiedProperties) {
                        $pname = $p.DisplayName
                        $oldV  = $p.OldValue
                        $newV  = $p.NewValue
                        $modProps += ("{0}: '{1}' → '{2}'" -f $pname,$oldV,$newV)
                    }
                }

                $details = @()
                if ($r.AdditionalDetails) {
                    foreach ($d in $r.AdditionalDetails) {
                        $details += ("{0}={1}" -f $d.Key, $d.Value)
                    }
                }

                $flattened.Add([pscustomobject]@{
                    ActivityDateTime         = $r.ActivityDateTime
                    ActivityDisplayName      = $r.ActivityDisplayName
                    Category                 = $r.Category
                    CorrelationId            = $r.CorrelationId
                    Result                   = $r.Result
                    ResultReason             = $r.ResultReason
                    LoggedByService          = $r.LoggedByService
                    IPAddress                = $ipAddr
                    InitiatedByUserId        = if ($userObj) { $userObj.Id } else { $null }
                    InitiatedByUPN           = if ($userObj) { $userObj.UserPrincipalName } else { $null }
                    InitiatedByUserDisplay   = if ($userObj) { $userObj.DisplayName } else { $null }
                    InitiatedByAppId         = if ($appObj) { $appObj.ServicePrincipalId } else { $null }
                    InitiatedByAppDisplay    = if ($appObj) { $appObj.DisplayName } else { $null }
                    TargetResources          = ($targets -join '; ')
                    ModifiedProperties       = ($modProps -join '; ')
                    AdditionalDetails        = ($details -join '; ')
                    RawId                    = $r.Id
                }) | Out-Null
            } catch {
                # If flattening fails for a record, fall back to a minimal projection
                $flattened.Add([pscustomobject]@{
                    ActivityDateTime    = $r.ActivityDateTime
                    ActivityDisplayName = $r.ActivityDisplayName
                    Category            = $r.Category
                    Result              = $r.Result
                    RawId               = $r.Id
                }) | Out-Null
            }
        }

        Write-Host "  Total audit log entries collected: $($flattened.Count)" -ForegroundColor Gray
        
        return [System.Collections.ArrayList]$flattened
    } catch {
        Write-Error "Failed to collect audit logs: $($_.Exception.Message)"
        return @()
    }
}

function Get-GraphSignInLogs {
    param(
        [Parameter(Mandatory=$false)]
        [int]$DaysBack = 7,
        [Parameter(Mandatory=$false)]
        [array]$SelectedUsers = @()
    )
    
    try {
        Write-Host "Collecting sign-in logs (last $DaysBack days)..." -ForegroundColor Yellow
        
        # Check if Microsoft Graph is connected
        $context = Get-MgContext -ErrorAction SilentlyContinue
        if (-not $context) {
            Write-Warning "Microsoft Graph not connected. Cannot collect sign-in logs."
            return @()
        }
        
        # Free tenants are limited to 7 days, Premium tenants can go up to 30 days
        if ($DaysBack -gt 7) {
            Write-Host "  Note: Retrieving more than 7 days requires Azure AD Premium license." -ForegroundColor Cyan
        }
        
        $allLogs = New-Object System.Collections.ArrayList
        $startDate = (Get-Date).AddDays(-$DaysBack).ToUniversalTime()
        $startIso = $startDate.ToString("yyyy-MM-ddTHH:mm:ssZ")
        
        # Build filter for date range
        $filter = "createdDateTime ge $startIso"
        
        # If specific users are selected, filter by user IDs (per-user mode)
        # If no users selected, collect all sign-in logs (all-users mode)
        if ($SelectedUsers -and $SelectedUsers.Count -gt 0) {
            Write-Host "  Per-user mode: Filtering sign-in logs for $($SelectedUsers.Count) selected user(s)..." -ForegroundColor Cyan
            
            $userIds = @()
            foreach ($upn in $SelectedUsers) {
                try {
                    $user = Get-MgUser -UserId $upn -Property Id -ErrorAction Stop
                    if ($user -and $user.Id) {
                        $userIds += $user.Id
                    }
                } catch {
                    Write-Warning "  Could not resolve user ID for $upn : $($_.Exception.Message)"
                }
            }
            
            if ($userIds.Count -gt 0) {
                # Build filter with user IDs (OR condition) for server-side filtering
                $userIdFilters = $userIds | ForEach-Object { "userId eq '$_'" }
                $userFilter = "(" + ($userIdFilters -join " or ") + ")"
                $filter = "$filter and $userFilter"
                Write-Host "  Server-side filtering: Querying sign-in logs for $($userIds.Count) user ID(s)..." -ForegroundColor Cyan
            } else {
                Write-Warning "  No valid user IDs found for selected users. Falling back to all-users mode."
            }
        } else {
            Write-Host "  All-users mode: Collecting sign-in logs for ALL users in the organization..." -ForegroundColor Cyan
        }
        
        Write-Host "  Querying Microsoft Graph API for sign-in logs..." -ForegroundColor Cyan
        
        try {
            # Attempt to retrieve sign-in logs
            $signIns = Get-MgAuditLogSignIn -Filter $filter -All -ErrorAction Stop
            
            if ($signIns) {
                Write-Host "  Retrieved $($signIns.Count) sign-in log entries" -ForegroundColor Green
                
                # Flatten sign-in logs for easier export
                foreach ($signIn in $signIns) {
                    try {
                        $logEntry = [PSCustomObject]@{
                            CreatedDateTime = $signIn.CreatedDateTime
                            UserPrincipalName = if ($signIn.UserId) { 
                                try {
                                    $user = Get-MgUser -UserId $signIn.UserId -Property UserPrincipalName -ErrorAction SilentlyContinue
                                    if ($user) { $user.UserPrincipalName } else { $signIn.UserId }
                                } catch {
                                    $signIn.UserId
                                }
                            } else { "Unknown" }
                            UserId = $signIn.UserId
                            AppDisplayName = $signIn.AppDisplayName
                            ClientAppUsed = $signIn.ClientAppUsed
                            IPAddress = $signIn.IpAddress
                            Location = if ($signIn.Location) {
                                $locParts = @()
                                if ($signIn.Location.City) { $locParts += $signIn.Location.City }
                                if ($signIn.Location.State) { $locParts += $signIn.Location.State }
                                if ($signIn.Location.CountryOrRegion) { $locParts += $signIn.Location.CountryOrRegion }
                                if ($locParts.Count -gt 0) { $locParts -join ", " } else { "Unknown" }
                            } else { "Unknown" }
                            Status = if ($signIn.Status) {
                                if ($signIn.Status.AdditionalDetails) { $signIn.Status.AdditionalDetails } else { $signIn.Status.ErrorCode }
                            } else { "Unknown" }
                            RiskLevelAggregated = $signIn.RiskLevelAggregated
                            RiskLevelDuringSignIn = $signIn.RiskLevelDuringSignIn
                            RiskState = $signIn.RiskState
                            ConditionalAccessStatus = if ($signIn.ConditionalAccessStatus) { $signIn.ConditionalAccessStatus } else { "Not Applied" }
                            DeviceDetail = if ($signIn.DeviceDetail) {
                                $deviceParts = @()
                                if ($signIn.DeviceDetail.Browser) { $deviceParts += $signIn.DeviceDetail.Browser }
                                if ($signIn.DeviceDetail.OperatingSystem) { $deviceParts += $signIn.DeviceDetail.OperatingSystem }
                                if ($deviceParts.Count -gt 0) { $deviceParts -join " / " } else { "Unknown" }
                            } else { "Unknown" }
                            ResourceDisplayName = $signIn.ResourceDisplayName
                            ResourceId = $signIn.ResourceId
                        }
                        [void]$allLogs.Add($logEntry)
                    } catch {
                        Write-Warning "  Error processing sign-in log entry: $($_.Exception.Message)"
                    }
                }
            } else {
                Write-Host "  No sign-in logs found for the specified time range" -ForegroundColor Yellow
            }
        } catch {
            # Check for licensing/permission errors
            $errorMsg = $_.Exception.Message
            if ($errorMsg -like "*insufficient privileges*" -or $errorMsg -like "*permission*" -or $errorMsg -like "*access denied*" -or $errorMsg -like "*Forbidden*") {
                Write-Warning "Permission denied: Sign-in logs require 'AuditLog.Read.All' permission."
                throw "Permission denied - requires AuditLog.Read.All permission"
            } elseif ($errorMsg -like "*license*" -or $errorMsg -like "*subscription*" -or $errorMsg -like "*premium*" -or $errorMsg -like "*not available*") {
                Write-Warning "License required: Sign-in logs require Azure AD Premium P1 or P2 license. Free tenants are limited to 7 days."
                throw "License required - Azure AD Premium P1/P2 (free tenants limited to 7 days)"
            } else {
                Write-Error "Failed to retrieve sign-in logs: $errorMsg"
                throw
            }
        }
        
        Write-Host "  Total sign-in log entries collected: $($allLogs.Count)" -ForegroundColor Gray
        
        return [System.Collections.ArrayList]$allLogs
    } catch {
        Write-Error "Failed to collect sign-in logs: $($_.Exception.Message)"
        return @()
    }
}

function Get-MailboxForwardingAndDelegation {
    param(
        [Parameter(Mandatory=$false)]
        [array]$SelectedUsers = @()
    )
    
    try {
        Write-Host "Collecting mailbox forwarding and delegation settings..." -ForegroundColor Yellow

        $mailboxes = @()
        
        # If SelectedUsers provided, only query those mailboxes (server-side filtering)
        if ($SelectedUsers -and $SelectedUsers.Count -gt 0) {
            foreach ($user in $SelectedUsers) {
                $upn = if ($user -is [string]) { $user } elseif ($user.UserPrincipalName) { $user.UserPrincipalName } else { continue }
                try {
                    $mbx = Get-Mailbox -Identity $upn -ErrorAction Stop
                    if ($mbx) { $mailboxes += $mbx }
                } catch {
                    Write-Warning "Mailbox not found for ${upn}: $($_.Exception.Message)"
                }
            }
        } else {
            # No selection - get all mailboxes
            try {
                $mailboxes = Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox,SharedMailbox -ErrorAction Stop
            } catch {
                # Fallback narrower call if needed
                $mailboxes = Get-Mailbox -ResultSize 2000 -ErrorAction Stop
            }
        }

        $results = New-Object System.Collections.Generic.List[object]

        foreach ($mbx in $mailboxes) {
            $upn = if ($mbx.UserPrincipalName) { $mbx.UserPrincipalName } else { $mbx.PrimarySmtpAddress }

            try {
                # Get mailbox-level forwarding settings
                $forwardingAddress = $mbx.ForwardingAddress
                $forwardingSmtpAddress = $mbx.ForwardingSmtpAddress
                $deliverToMailboxAndForward = $mbx.DeliverToMailboxAndForward

                # Get delegation/permissions
                $fullAccessUsers = @()
                $sendAsUsers = @()
                $sendOnBehalfUsers = @()

                try {
                    $permissions = Get-MailboxPermission -Identity $upn -ErrorAction SilentlyContinue |
                                   Where-Object { $_.User -notlike "*NT AUTHORITY*" -and $_.User -notlike "*S-1-*" -and $_.IsInherited -eq $false }

                    $fullAccessUsers = $permissions | Where-Object { $_.AccessRights -contains "FullAccess" } |
                                      Select-Object -ExpandProperty User | ForEach-Object { $_.ToString() }

                    $sendAsUsers = $permissions | Where-Object { $_.AccessRights -contains "SendAs" } |
                                  Select-Object -ExpandProperty User | ForEach-Object { $_.ToString() }
                } catch {}

                # Get Send-On-Behalf
                try {
                    if ($mbx.GrantSendOnBehalfTo -and $mbx.GrantSendOnBehalfTo.Count -gt 0) {
                        $sendOnBehalfUsers = $mbx.GrantSendOnBehalfTo | ForEach-Object { $_.ToString() }
                    }
                } catch {}

                $obj = [pscustomobject]@{
                    UserPrincipalName           = $upn
                    DisplayName                 = $mbx.DisplayName
                    RecipientType               = $mbx.RecipientTypeDetails
                    ForwardingAddress           = if ($forwardingAddress) { $forwardingAddress.ToString() } else { $null }
                    ForwardingSmtpAddress       = $forwardingSmtpAddress
                    DeliverToMailboxAndForward  = $deliverToMailboxAndForward
                    FullAccessUsers             = ($fullAccessUsers -join '; ')
                    SendAsUsers                 = ($sendAsUsers -join '; ')
                    SendOnBehalfUsers           = ($sendOnBehalfUsers -join '; ')
                }
                [void]$results.Add($obj)

            } catch {
                Write-Warning "Failed to process mailbox ${upn}: $($_.Exception.Message)"
            }
        }

        return [System.Collections.ArrayList]$results
    } catch {
        Write-Error "Failed to collect mailbox forwarding and delegation: $($_.Exception.Message)"
        return @()
    }
}

function Get-MailFlowConnectors {
    try {
        Write-Host "Collecting mail flow connectors..." -ForegroundColor Yellow

        $results = New-Object System.Collections.Generic.List[object]

        # Get inbound connectors
        $inboundConns = @()
        try {
            $params = @{ ErrorAction = 'Stop'; WarningAction = 'SilentlyContinue' }
            $gc = Get-Command Get-InboundConnector -ErrorAction SilentlyContinue
            if ($gc -and $gc.Parameters.ContainsKey('IncludeTestModeConnectors')) { $params.IncludeTestModeConnectors = $true }
            $inboundConns = Get-InboundConnector @params
        } catch { $inboundConns = @() }

        foreach ($c in $inboundConns) {
            $results.Add([pscustomobject]@{
                Direction                     = 'Inbound'
                Name                          = $c.Name
                ConnectorType                 = $c.ConnectorType
                Enabled                       = $c.Enabled
                SenderDomains                 = ($c.SenderDomains -join ';')
                SenderIPAddresses             = ($c.SenderIPAddresses -join ';')
                RecipientDomains              = $null
                SmartHosts                    = $null
                RestrictDomainsToCertificate  = $c.RestrictDomainsToCertificate
                RestrictDomainsToIPAddresses  = $c.RestrictDomainsToIPAddresses
                TlsSenderCertificateName      = $c.TlsSenderCertificateName
                TlsSettings                   = $null
                TlsDomain                     = $null
                RequireTls                    = $c.RequireTls
                UseMXRecord                   = $null
                CloudServicesMailEnabled      = $c.CloudServicesMailEnabled
                Comment                       = $c.Comment
                Identity                      = $c.Identity
                Guid                          = $c.Guid
                TestMode                      = $(if ($c.PSObject.Properties['TestMode']) { $c.TestMode } elseif ($c.PSObject.Properties['IsTestMode']) { $c.IsTestMode } else { $null })
            }) | Out-Null
        }

        # Get outbound connectors
        $outboundConns = @()
        try {
            $params = @{ ErrorAction = 'Stop'; WarningAction = 'SilentlyContinue' }
            $gc = Get-Command Get-OutboundConnector -ErrorAction SilentlyContinue
            if ($gc -and $gc.Parameters.ContainsKey('IncludeTestModeConnectors')) { $params.IncludeTestModeConnectors = $true }
            $outboundConns = Get-OutboundConnector @params
        } catch { $outboundConns = @() }

        foreach ($c in $outboundConns) {
            $results.Add([pscustomobject]@{
                Direction                     = 'Outbound'
                Name                          = $c.Name
                ConnectorType                 = $c.ConnectorType
                Enabled                       = $c.Enabled
                SenderDomains                 = $null
                SenderIPAddresses             = $null
                RecipientDomains              = ($c.RecipientDomains -join ';')
                SmartHosts                    = ($c.SmartHosts -join ';')
                RestrictDomainsToCertificate  = $null
                RestrictDomainsToIPAddresses  = $null
                TlsSenderCertificateName      = $null
                TlsSettings                   = $c.TlsSettings
                TlsDomain                     = $c.TlsDomain
                RequireTls                    = $null
                UseMXRecord                   = $c.UseMXRecord
                CloudServicesMailEnabled      = $c.CloudServicesMailEnabled
                Comment                       = $c.Comment
                Identity                      = $c.Identity
                Guid                          = $c.Guid
                TestMode                      = $(if ($c.PSObject.Properties['TestMode']) { $c.TestMode } elseif ($c.PSObject.Properties['IsTestMode']) { $c.IsTestMode } else { $null })
            }) | Out-Null
        }

        return [System.Collections.ArrayList]$results
    } catch {
        Write-Error "Failed to collect mail flow connectors: $($_.Exception.Message)"
        return @()
    }
}

# Portal-like export fallback using Entra Sign-in Logs export API without AAD Premium
function Export-EntraPortalSignInCsv {
    param(
        [Parameter(Mandatory=$true)][datetime]$StartUtc,
        [Parameter(Mandatory=$true)][datetime]$EndUtc,
        [Parameter(Mandatory=$true)][string]$OutputCsv
    )

    try {
        # This uses the public portal CSV endpoint (same data the portal downloads), authenticated with the current Graph token.
        # Note: Availability and schema may vary. This is a best-effort fallback when AuditLog.Read.All is blocked by licensing.

        # Acquire raw bearer token from current context
        $ctx = Get-MgContext -ErrorAction Stop
        $token = $null
        try { $token = (Get-MgContext).AccessToken } catch {}
        if (-not $token) {
            # Fallback to MSAL token provider inside Graph SDK
            $token = (Get-MgProfile -ErrorAction SilentlyContinue) | Out-Null
        }

        $s = $StartUtc.ToString('yyyy-MM-ddTHH:mm:ssZ')
        $e = $EndUtc.ToString('yyyy-MM-ddTHH:mm:ssZ')

        # Known portal CSV route (subject to change by Microsoft). We pass time range and request CSV.
        $csvUri = "https://graph.microsoft.com/v1.0/auditLogs/signIns?`$filter=createdDateTime ge $s and createdDateTime lt $e&`$count=true"
        $headers = @{ Accept = 'text/csv'; ConsistencyLevel = 'eventual' }

        $resp = Invoke-MgGraphRequest -Uri $csvUri -Method GET -Headers $headers -OutputFilePath $OutputCsv -ErrorAction SilentlyContinue
        if (Test-Path $OutputCsv) { return $true }
        return $false
    } catch {
        Write-Warning "Portal-like CSV export failed: $($_.Exception.Message)"
        return $false
    }
}

function New-AISecurityInvestigationPrompt {
    param([Parameter(Mandatory=$true)]$Report)

    # Calculate data counts outside the here-string to avoid parsing issues
    $messageTraceCount = if($Report.MessageTrace){$Report.MessageTrace.Count}else{0}
    $inboxRulesCount = if($Report.InboxRules){$Report.InboxRules.Count}else{0}
    $transportRulesCount = if($Report.TransportRules){$Report.TransportRules.Count}else{0}
    $inboundConnCount = if($Report.InboundConnectors){$Report.InboundConnectors.Count}else{0}
    $outboundConnCount = if($Report.OutboundConnectors){$Report.OutboundConnectors.Count}else{0}
    $auditLogsCount = if($Report.AuditLogs){$Report.AuditLogs.Count}else{0}
    $signinLogsCount = 0
    $caPoliciesCount = if($Report.ConditionalAccessPolicies){$Report.ConditionalAccessPolicies.Count}else{0}
    $appRegistrationsCount = if($Report.AppRegistrations){$Report.AppRegistrations.Count}else{0}

    $prompt = @"
# SECURITY INVESTIGATION AI PROMPT

## INVESTIGATOR INFORMATION
- **Investigator Name:** $($Report.Investigator)
- **Company:** $($Report.Company)
- **Investigation Date:** $($Report.Timestamp)
- **Analysis Period:** Last $($Report.DaysAnalyzed) days

## DATA SOURCES PROVIDED
- **Message Trace:** $messageTraceCount email records
- **Inbox Rules:** $inboxRulesCount rules across all mailboxes
- **Transport Rules (Mail Flow):** $transportRulesCount rules
- **Connectors:** $inboundConnCount inbound, $outboundConnCount outbound
- **Audit Logs:** $auditLogsCount directory audit events
- **MFA Coverage:** tenant-wide defaults/CA and per-user states
- **Conditional Access Policies:** $caPoliciesCount policies
- **App Registrations:** $appRegistrationsCount registered applications

## INVESTIGATION OBJECTIVES

### 1. EMAIL SECURITY ANALYSIS
Analyze the message trace data for:
- **Suspicious external email patterns** (unusual recipients, high volume to external domains)
- **Potential data exfiltration** (large attachments, sensitive content patterns)
- **Unauthorized forwarding** (rules forwarding to external addresses)
- **Email spoofing attempts** (mismatched sender/recipient patterns)

### 2. AUTHENTICATION POSTURE (NO SIGN-IN LOGS)
Assess MFA coverage and controls:
- **Security Defaults** status (on/off)
- **Conditional Access** policies requiring MFA
- **Per-user MFA** (enabled/disabled)
- **Coverage Gaps** (users without any MFA control)

### 3. INBOX RULE ANALYSIS
Investigate inbox rules for:
- **Hidden rules** (rules that are not visible to users)
- **External forwarding** (rules automatically forwarding emails to external domains)
- **Suspicious conditions** (rules triggered by specific keywords or senders)
- **Mass rule creation** (unusual number of rules created recently)

### 4. ADMINISTRATIVE ACTIVITY
Review audit logs for:
- **Unauthorized privilege changes** (role assignments, permission modifications)
- **Suspicious administrative actions** (mass user modifications, policy changes)
- **Account manipulation** (password resets, account unlocks, suspicious logins)

## DELIVERABLES REQUIRED

### 1. Executive Summary
Provide a clear, non-technical summary of findings for senior management including:
- Overall risk level (Critical/High/Medium/Low)
- Key findings and their business impact
- Immediate actions required
- Long-term recommendations

### 2. Technical Analysis Report
Include detailed technical findings with:
- Specific compromised accounts or systems
- Timeline of malicious activities
- Evidence chain linking related events
- Technical remediation steps

### 3. Incident Response Plan
Provide specific steps for:
- Containment of active threats
- Eradication of malicious elements
- Recovery of affected systems
- Prevention of future incidents

## ANALYSIS CRITERIA

### Risk Assessment
- **Critical:** Active data exfiltration, ransomware deployment, or system compromise
- **High:** Unauthorized access attempts, suspicious authentication patterns
- **Medium:** Policy violations, unusual but non-malicious behavior
- **Low:** Minor anomalies requiring monitoring

### Prioritization
1. **Immediate Response Required:** Active threats, data loss, system compromise
2. **Urgent Investigation:** Suspicious patterns requiring deeper analysis
3. **Monitoring Required:** Unusual but non-malicious activities
4. **Documentation Only:** Normal operational activities

## REPORTING FORMAT

Please structure your response as follows:

### EXECUTIVE SUMMARY
[3-5 paragraphs for non-technical audience]

### DETAILED FINDINGS
[Technical analysis with specific evidence]

### IMMEDIATE ACTIONS
[Specific steps to contain and remediate, including enabling MFA for uncovered users]

### LONG-TERM RECOMMENDATIONS
[Preventive measures and improvements]

### APPENDIX
[Raw data analysis, timelines, evidence details]

"@

    return $prompt
}

function New-TicketSecuritySummary {
    param([Parameter(Mandatory=$true)]$Report)

    # Calculate data counts outside the here-string to avoid parsing issues
    $messageTraceCount = if($Report.MessageTrace){$Report.MessageTrace.Count}else{0}
    $inboxRulesCount = if($Report.InboxRules){$Report.InboxRules.Count}else{0}
    $transportRulesCount = if($Report.TransportRules){$Report.TransportRules.Count}else{0}
    $inboundConnCount = if($Report.InboundConnectors){$Report.InboundConnectors.Count}else{0}
    $outboundConnCount = if($Report.OutboundConnectors){$Report.OutboundConnectors.Count}else{0}
    $auditLogsCount = if($Report.AuditLogs){$Report.AuditLogs.Count}else{0}
    $signinLogsCount = 0

    $message = @"
**URGENT: Security Investigation Required**

**Reported By:** $($Report.Investigator)
**Company:** $($Report.Company)
**Date:** $($Report.Timestamp)

---

## INCIDENT SUMMARY

A comprehensive security investigation has been completed for our Microsoft 365 environment. The analysis covered email communications, user authentication patterns, and administrative activities over the past $($Report.DaysAnalyzed) days.

### Data Sources Analyzed:
- **Email Communications:** $messageTraceCount messages tracked
- **User Rules:** $inboxRulesCount inbox rules examined
- **Mail Flow Rules:** $transportRulesCount transport rules examined
- **Connectors:** $inboundConnCount inbound, $outboundConnCount outbound
- **Security Logs:** $auditLogsCount audit events reviewed
- **MFA Coverage:** tenant defaults/CA/per-user evaluated

### Key Areas of Concern:

**Email Security:**
- Review of all incoming and outgoing email patterns
- Analysis of automated email forwarding rules
- Investigation of unusual external communications

**User Access & MFA:**
- MFA coverage and gaps across Security Defaults, CA, and Per-user
- Priority list of users without MFA coverage

**Administrative Changes:**
- Recent privilege modifications
- Account creation and modification activities
- Security policy changes

---

## IMMEDIATE ATTENTION REQUIRED

The investigation team has identified several areas requiring immediate attention. Please review the detailed findings and prioritize the following:

1. **Account Access Review** - Verify all recent authentication attempts
2. **Email Flow Analysis** - Examine external email communications
3. **Rule Assessment** - Review automated email processing rules
4. **Permission Audit** - Confirm all privilege changes are authorized

---

## NEXT STEPS

**For IT/Security Team:**
1. Review the detailed technical analysis report
2. Implement immediate containment measures if threats are active
3. Coordinate with affected department heads
4. Update security monitoring and alerting rules

**For Executive Leadership:**
1. Review the business impact assessment
2. Approve resource allocation for remediation
3. Communicate with stakeholders as appropriate
4. Support implementation of recommended security improvements

---

**Investigation Details:**
- **Analysis Period:** Last $($Report.DaysAnalyzed) days
- **Tools Used:** Exchange Online, Microsoft Graph, Entra ID
- **Report Generated:** $($Report.Timestamp)
- **Investigator:** $($Report.Investigator)

---

## Files Provided And How To Use Them

Location: $($Report.OutputFolder)

- MessageTrace.csv: Upload to your analysis workspace/LLM to identify unusual external flows and spikes.
- InboxRules.csv: Review for forwarding/hidden/suspicious rules; feed to LLM for triage.
- TransportRules.csv: Review for risky conditions/actions (auto-forwarding, allow lists, spoof bypass).
- InboundConnectors.csv / OutboundConnectors.csv: Validate trusted partners, smart hosts, TLS settings, and domain scopes.
- AuditLogs.csv: Examine administrative actions and policy changes.
- ConditionalAccessPolicies.csv: Review for malicious CA policies that bypass MFA, apply to all users/locations, or have suspicious exclusions. Focus on high-risk policies.
- AppRegistrations.csv: Review for malicious apps with high-privilege permissions, suspicious redirect URIs, unverified publishers, or user consent enabled. Focus on high-risk applications.
- MFAStatus.csv: Identify users not covered by any MFA control; prioritize remediation.
- UserSecurityGroups.csv: Validate privileged group/role membership (e.g., Global Administrator).

Important: Sign-in logs require Entra ID Premium for API access. Please export sign-in CSVs from the Entra portal (Sign-in logs → Download, last 7–30 days depending on tenant) and include alongside these files for full analysis.

*This automated security analysis helps identify potential security incidents and unusual patterns that may require further investigation by security professionals.*
"@

    return $message
}

function New-LLMInvestigationInstructions {
    param([Parameter(Mandatory=$true)]$Report)

    # Try to use settings-based generator, fallback to basic if Settings module not available
    try {
        $settingsPath = Join-Path $PSScriptRoot 'Settings.psm1'
        if (Test-Path $settingsPath) {
            Import-Module $settingsPath -Force -ErrorAction SilentlyContinue
            if (Get-Command Generate-AIReadme -ErrorAction SilentlyContinue) {
                $settings = Get-AppSettings
                # Override InvestigatorName and CompanyName from report if provided
                if ($Report.Investigator) { $settings.InvestigatorName = $Report.Investigator }
                if ($Report.Company) { $settings.CompanyName = $Report.Company }
                return Generate-AIReadme -Settings $settings
            }
        }
    } catch {
        Write-Warning "Could not load settings-based AI readme generator: $($_.Exception.Message)"
    }

    # Fallback to basic template if settings not available
    $investigator = if ($Report.Investigator) { $Report.Investigator } else { 'Security Administrator' }
    $company = if ($Report.Company) { $Report.Company } else { 'Organization' }

    $instructions = @"
Master Prompt - Generic Template (Copy and Save This)

Role & Objective You are a Security Engineer acting on behalf of $company. Your task is to analyze security alert tickets, cross-reference them with attached CSV logs/text files, and classify the event as True Positive, False Positive, or Authorized Activity.



You will then draft a non-technical, professional email response to the client contact.



I. Data Ingestion & Analysis Rules

1. Analyze the Ticket Context



Ticket Body: Extract the User, Timestamp (UTC), IP Address, and Alert Type.



Ticket Notes/Configs: Look for notes like "Remote Employees," "Office Key," or specific authorized devices which indicate authorized activity.



Contact Name: Extract the contact from the "Contact" field. Check the "Client Specific Nuances" section below for any naming overrides.



2. Verify with Logs (The "Evidence" Rule)



Crucial: Do not rely solely on the ticket description. You must find the corresponding event in the attached CSVs (SignInLogs, GraphAudit, etc.) to confirm the activity.



Time Zone: Convert all UTC timestamps to CST (Central Standard Time) for the email.



II. Classification Logic

A. Authorized Activity (White-Listed)

Internal Admin Accounts: Usernames like [admin], [service_account], or [rmm_account].



Verification: Check UserSecurityPosture.csv. If the Display Name matches your internal team (e.g., "Managed Services"), treat as Authorized.



Action: Classify as Authorized Activity (Administrative Maintenance).



Travel (Residential/Mobile): Logins from standard ISPs (Comcast, Charter, CenturyLink, Verizon, Brightspeed, AT&T, T-Mobile) in a different city/state.



Action: Classify as Authorized Activity (User Travel/Remote Work).



In-Flight Wi-Fi: IPs from Anuvu, Gogo, Viasat, Panasonic Avionics.



Action: Classify as Authorized Activity.



Service Principals: "MFA Disabled" alerts where the Actor is "Microsoft Graph Command Line Tools" or a known Admin.



Action: Classify as Authorized Activity (Maintenance Script).



B. False Positives (System Noise)

Endpoint Protection: Alerts for TrustedInstaller.exe, `$`$DeleteMe..., or files in \Windows\WinSxS\Temp\.



Action: Classify as False Positive (System Update/Cleanup).



C. True Positives (Compromise Indicators)

Inbox Rules:



Name consists only of non-alphanumeric characters (e.g., ., .., ,,, ).



Action moves mail to "RSS Feeds" or "Conversation History" folders.



Action: Classify as True Positive. Recommend immediate password reset & session revocation.



D. Suspicious (Requires Confirmation)

Hosting Providers: Logins from AWS, DigitalOcean, Linode (unless the user has a known hosted workflow).



Consumer VPNs: NordVPN, ProtonVPN, Private Internet Access.



Action: Draft email asking for confirmation.



III. Output Format

Subject: Security Alert: Ticket #[Ticket Number] - [Brief Subject]



Hi [Contact First Name],



[Opening: State the alert type and the user involved.]



[Verdict: Explicitly state: "We have classified this as [Category]."]



[Analysis:



Source: [ISP Name / Location] (IP: [IP Address])



Evidence: Explain why it is classified this way (e.g., "This is a standard residential ISP," or "The rule name '.' is a known indicator of compromise"). Cite the specific log file used (e.g., ``).]



[Action Taken/Required:



If Authorized/False Positive: "No further action is required. We have closed this ticket."



If Suspicious: "Please confirm if [User] is currently [Traveling/Using a VPN]."



If True Positive: "We recommend immediately resetting the password and revoking sessions."]



Best,



$investigator



Clarification Questions [Ask 2 questions here regarding tuning, specific client policies, or missing data.]
"@

    return $instructions
}

function New-SecurityInvestigationSummary {
    param([Parameter(Mandatory=$true)]$Report)

    # Calculate data counts outside the here-string to avoid parsing issues
    $messageTraceCount = if($Report.MessageTrace){$Report.MessageTrace.Count}else{0}
    $inboxRulesCount = if($Report.InboxRules){$Report.InboxRules.Count}else{0}
    $mailboxesAnalyzed = if($Report.InboxRules){
        ($Report.InboxRules | Select-Object -Property MailboxOwner -Unique).Count
    }else{0}
    $transportRulesCount = if($Report.TransportRules){$Report.TransportRules.Count}else{0}
    $inboundConnCount = if($Report.InboundConnectors){$Report.InboundConnectors.Count}else{0}
    $outboundConnCount = if($Report.OutboundConnectors){$Report.OutboundConnectors.Count}else{0}
    $auditLogsCount = if($Report.AuditLogs){$Report.AuditLogs.Count}else{0}
    $signinLogsCount = 0
    $usersWithActivity = 0
    $caPoliciesCount = if($Report.ConditionalAccessPolicies){$Report.ConditionalAccessPolicies.Count}else{0}
    $appRegistrationsCount = if($Report.AppRegistrations){$Report.AppRegistrations.Count}else{0}
    $highRiskCAPolicies = if($Report.ConditionalAccessPolicies){($Report.ConditionalAccessPolicies | Where-Object { $_.RiskLevel -eq "High" }).Count}else{0}
    $highRiskApps = if($Report.AppRegistrations){($Report.AppRegistrations | Where-Object { $_.RiskLevel -eq "High" }).Count}else{0}

    $summary = @"
# COMPREHENSIVE SECURITY INVESTIGATION REPORT

## Report Overview
**Generated:** $($Report.Timestamp)
**Investigator:** $($Report.Investigator)
**Organization:** $($Report.Company)
**Analysis Period:** Last $($Report.DaysAnalyzed) days

## Data Collection Summary

### Exchange Online Data
- **Message Trace Records:** $messageTraceCount
- **Inbox Rules Exported:** $inboxRulesCount
- **Mailboxes Analyzed:** $mailboxesAnalyzed
- **Transport Rules Exported:** $transportRulesCount
- **Connectors Exported:** $inboundConnCount inbound, $outboundConnCount outbound
- **Connection Status:** $($Report.ExchangeConnection)

### Microsoft Graph Data
- **Audit Log Events:** $auditLogsCount
- **Conditional Access Policies:** $caPoliciesCount policies ($highRiskCAPolicies high-risk)
- **App Registrations:** $appRegistrationsCount applications ($highRiskApps high-risk)
- **Connection Status:** $($Report.GraphConnection)

## Investigation Tools and Methods

### Email Security Analysis
- **Message Trace Review:** Analyzed all email sent/received patterns
- **Inbox Rule Audit:** Examined automated email processing rules
- **External Communication Patterns:** Identified unusual external email flows
- **Forwarding Rule Detection:** Flagged rules forwarding to external domains

### Authentication Analysis
- Replaced sign-in log analysis with MFA coverage and security posture review
- **Conditional Access Policy Review:** Analyzed CA policies for malicious configurations, MFA bypasses, and security gaps
- **App Registration Security:** Reviewed app registrations for high-privilege permissions, suspicious configurations, and potential threats

### Administrative Activity Review
- **Privilege Changes:** Monitored role assignments and permission modifications
- **Account Management:** Tracked account creation, modification, and deletion
- **Security Policy Changes:** Reviewed authentication and access policy updates
- **Audit Trail Analysis:** Examined all administrative actions with timestamps

## Key Findings and Recommendations

### Immediate Actions Required
1. **MFA Coverage Gaps:** Remediate users not covered by per-user MFA, Security Defaults, or Conditional Access
2. **High-Risk CA Policies:** Review and remediate $highRiskCAPolicies high-risk Conditional Access policies that may bypass security controls
3. **High-Risk App Registrations:** Investigate and remediate $highRiskApps high-risk app registrations with suspicious permissions or configurations
4. **Audit Email Forwarding Rules:** Verify all external forwarding is authorized
5. **Examine Privilege Changes:** Confirm recent role assignments are legitimate
6. **Monitor External Communications:** Review patterns to unusual external domains

### Security Improvements Recommended
1. **Enhanced MFA Enforcement:** Implement MFA for all external access
2. **Email Rule Governance:** Establish approval process for forwarding rules
3. **Access Monitoring:** Implement real-time alerting for suspicious sign-ins
4. **Regular Audits:** Schedule quarterly security reviews

## Technical Details

### Data Export Formats
- **Message Trace:** CSV format with timestamp, sender, recipient, and metadata
- **Inbox Rules:** CSV format with rule details, conditions, and actions
- **Audit Logs:** CSV format with activity details and user information
- **Conditional Access Policies:** CSV format with policy details, risk analysis, and suspicious indicators
- **App Registrations:** CSV format with app details, permissions, risk analysis, and security indicators

### Investigation Timeline
- **Data Collection:** Automated collection from multiple sources
- **Analysis Period:** $($Report.DaysAnalyzed) days of historical data
- **Report Generation:** Real-time compilation of findings
- **AI Enhancement:** Structured prompts for advanced analysis

## Contact Information
**Security Investigator:** $($Report.Investigator)
**Organization:** $($Report.Company)
**Report Generated:** $($Report.Timestamp)

*This report provides a comprehensive view of security-relevant activities and serves as a foundation for deeper investigation and remediation efforts.*
"@

    return $summary
}

function New-SecurityInvestigationZip {
    param(
        [Parameter(Mandatory=$true)]
        [string]$OutputFolder,
        [Parameter(Mandatory=$false)]
        [string]$ZipFileName
    )

    try {
        # Validate output folder exists
        if (-not (Test-Path $OutputFolder)) {
            Write-Error "Output folder does not exist: $OutputFolder"
            return $null
        }

        # Determine zip file name
        if ([string]::IsNullOrWhiteSpace($ZipFileName)) {
            $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
            $ZipFileName = "SecurityInvestigation_$timestamp.zip"
        }

        # Ensure .zip extension
        if (-not $ZipFileName.EndsWith('.zip')) {
            $ZipFileName += '.zip'
        }

        # Create zip file path in the output folder
        $zipPath = Join-Path $OutputFolder $ZipFileName

        # Get all CSV and JSON files, excluding _AI_Readme.txt
        $filesToZip = Get-ChildItem -Path $OutputFolder -Include *.csv,*.json -Recurse |
                      Where-Object { $_.Name -ne '_AI_Readme.txt' }

        if ($filesToZip.Count -eq 0) {
            Write-Warning "No CSV or JSON files found to zip in $OutputFolder"
            return $null
        }

        # Remove existing zip file if it exists
        if (Test-Path $zipPath) {
            Remove-Item $zipPath -Force
        }

        # Create the zip file using Compress-Archive
        Compress-Archive -Path $filesToZip.FullName -DestinationPath $zipPath -CompressionLevel Optimal -ErrorAction Stop

        Write-Host "Successfully created zip file: $zipPath" -ForegroundColor Green
        Write-Host "Files included: $($filesToZip.Count)" -ForegroundColor Cyan

        return $zipPath
    } catch {
        Write-Error "Failed to create zip file: $($_.Exception.Message)"
        return $null
    }
}

# Get M365/O365 license SKU mapping from tenant
function Get-TenantLicenseSkus {
    try {
        # Static mapping of common SKU GUIDs to friendly names (fallback for SKUs not in tenant subscriptions)
        # Define this first so it's always available even if module import fails
        $staticSkuMapping = @{
            'cbdc14ab-d96c-4c30-b9f4-6ada7cdc1d46' = 'Microsoft 365 Business Premium'
            'a403ebcc-fae0-4ca2-8c8c-7a907fd6c235' = 'Power BI (Free)'
            'f30db892-07e9-47e9-837c-80727f46fd3d' = 'Microsoft Power Automate Free'
            '06ebc4ee-1bb5-47dd-8120-11324b54e684' = 'Microsoft 365 E3'
            'c7df2760-2c81-4ef7-b578-5b5392b571df' = 'Microsoft 365 E5'
            'b17653a4-2443-4e8c-a550-18249dda78bb' = 'Office 365 E1'
            '6634e0ce-1a9f-428c-a498-f84ec7b8aa2e' = 'Office 365 E2'
            '4b585984-651b-448a-9e53-3b10f069cf7f' = 'Microsoft 365 F3'
            '9aaf7827-d63c-4b61-89c3-182f06f82e5c' = 'Exchange Online (Plan 1)'
            'efb87545-963c-4e0d-99df-69c6916d9eb0' = 'Exchange Online (Plan 2)'
            '80b2d799-d2ba-4d2a-8842-fb0d0f3a4b82' = 'Exchange Online Kiosk'
            '0d259279-6a13-4952-bb13-0afb3ae5f8ae' = 'Skype for Business Online (Plan 2)'
            'f8a1db68-be16-40ed-86d5-cb42ce701560' = 'Power BI Pro'
            'b21a6192-5159-478e-8ca0-47e3c25e3a33' = 'Project Plan 3'
            '776df282-9c98-49a8-a7dc-9f4b4a88e260' = 'Project Plan 1'
            'c5928f49-12ba-48f7-ada3-0d743a3601d5' = 'Visio Plan 2'
            '4de31727-a228-4ec3-a5bf-8e705b5ea9c1' = 'Microsoft Defender for Office 365 (Plan 1)'
            'e20c9ac9-9e62-4b5c-8b13-efd88a3b8c7a' = 'Microsoft Defender for Office 365 (Plan 2)'
            'efccb6f7-5641-4e0e-bd10-b4976e1bf68e' = 'Enterprise Mobility + Security E3'
            'b05e124f-c7cc-45a0-a6aa-8cf78c946968' = 'Enterprise Mobility + Security E5'
            '41781fb2-bc02-4b7c-bd55-b576c07bb09d' = 'Azure Active Directory Premium P1'
            'eec0eb4f-6444-4f95-aba0-50c24d67f998' = 'Azure Active Directory Premium P2'
            'c52ea49f-fe5d-4e95-93da-0ef2c73fe964' = 'Azure Rights Management'
            'b43305a7-bc43-4eb6-bc21-ee2199e86b14' = 'Power Apps Trial'
            '710779e8-3d4a-4c58-ad3e-8ac173e5e5a5' = 'Microsoft Teams Exploratory'
            '66b55226-6b4f-492c-910c-a9977c18ad61' = 'Microsoft 365 F1'
            '3b555118-da6a-4418-894f-7df1e2096870' = 'Microsoft 365 Business Basic'
            'ac5cefde-6b63-4b5e-8b0f-4b5e8b0f4b5e' = 'Microsoft 365 Business Standard'
            '094e7854-93fc-4d55-b2c0-3ab5369ebdc1' = 'Microsoft 365 Apps for Business'
            'cdd28e44-67e3-425e-be4c-737fab2899d3' = 'Microsoft 365 E5 Developer'
            '1f2f344a-3d43-41d8-8fe6-0a43e9bb6637' = 'Microsoft Stream'
            'e97c048c-37a4-45fb-ab50-922f1c7676cc' = 'Microsoft 365 A5 for Faculty'
            '46c119d4-0379-4a9d-85e4-7557fab0a5d0' = 'Microsoft 365 A5 for Students'
            '7cfd9a2b-e110-4c39-bf20-c6a3f36a3121' = 'Microsoft 365 A3 for Faculty'
            '98b6e773-24d4-4c0d-a968-6e787a1f8204' = 'Microsoft 365 A3 for Students'
        }
        
        $skus = @{}
        
        # Ensure Microsoft.Graph.Identity.DirectoryManagement module is imported (if available)
        $canGetTenantSkus = $false
        if (Get-Command Get-MgSubscribedSku -ErrorAction SilentlyContinue) {
            $canGetTenantSkus = $true
        } else {
            try {
                Import-Module -Name Microsoft.Graph.Identity.DirectoryManagement -Force -ErrorAction SilentlyContinue
                if (Get-Command Get-MgSubscribedSku -ErrorAction SilentlyContinue) {
                    $canGetTenantSkus = $true
                }
            } catch {
                # Module not available - will use static mapping only
            }
        }
        
        # Try to get tenant SKUs, but continue even if it fails
        if ($canGetTenantSkus) {
            try {
                $tenantSkus = Get-MgSubscribedSku -All -ErrorAction Stop

                foreach ($sku in $tenantSkus) {
                    $skuId = $sku.SkuId
                    $skuPartNumber = $sku.SkuPartNumber

                    # Create friendly name mapping
                    $friendlyName = switch ($skuPartNumber) {
                        'ENTERPRISEPACK' { 'Microsoft 365 E3' }
                        'ENTERPRISEPREMIUM' { 'Microsoft 365 E5' }
                        'ENTERPRISEPACK_B_PILOT' { 'Microsoft 365 E3' }
                        'ENTERPRISEPREMIUM_NOPSTNCONF' { 'Microsoft 365 E5 (No Audio Conferencing)' }
                        'SPE_E3' { 'Microsoft 365 E3' }
                        'SPE_E5' { 'Microsoft 365 E5' }
                        'STANDARDPACK' { 'Office 365 E1' }
                        'STANDARDWOFFPACK' { 'Office 365 E2' }
                        'DESKLESSPACK' { 'Microsoft 365 F3' }
                        'SPE_F1' { 'Microsoft 365 F3' }
                        'EXCHANGESTANDARD' { 'Exchange Online (Plan 1)' }
                        'EXCHANGEENTERPRISE' { 'Exchange Online (Plan 2)' }
                        'EXCHANGEDESKLESS' { 'Exchange Online Kiosk' }
                        'MCOSTANDARD' { 'Skype for Business Online (Plan 2)' }
                        'POWER_BI_PRO' { 'Power BI Pro' }
                        'POWER_BI_STANDARD' { 'Power BI (Free)' }
                        'PROJECTPROFESSIONAL' { 'Project Plan 3' }
                        'PROJECTONLINE_PLAN_1' { 'Project Plan 1' }
                        'VISIOCLIENT' { 'Visio Plan 2' }
                        'ATP_ENTERPRISE' { 'Microsoft Defender for Office 365 (Plan 1)' }
                        'THREAT_INTELLIGENCE' { 'Microsoft Defender for Office 365 (Plan 2)' }
                        'EMS' { 'Enterprise Mobility + Security E3' }
                        'EMSPREMIUM' { 'Enterprise Mobility + Security E5' }
                        'AAD_PREMIUM' { 'Azure Active Directory Premium P1' }
                        'AAD_PREMIUM_P2' { 'Azure Active Directory Premium P2' }
                        'RIGHTSMANAGEMENT' { 'Azure Rights Management' }
                        'FLOW_FREE' { 'Power Automate Free' }
                        'POWERAPPS_VIRAL' { 'Power Apps Trial' }
                        'TEAMS_EXPLORATORY' { 'Microsoft Teams Exploratory' }
                        'M365_F1_COMM' { 'Microsoft 365 F1' }
                        'SPB' { 'Microsoft 365 Business Premium' }
                        'SMB_BUSINESS' { 'Microsoft 365 Business Basic' }
                        'SMB_BUSINESS_ESSENTIALS' { 'Microsoft 365 Business Basic' }
                        'SMB_BUSINESS_PREMIUM' { 'Microsoft 365 Business Standard' }
                        'O365_BUSINESS' { 'Microsoft 365 Apps for Business' }
                        'O365_BUSINESS_ESSENTIALS' { 'Microsoft 365 Business Basic' }
                        'O365_BUSINESS_PREMIUM' { 'Microsoft 365 Business Standard' }
                        'DEVELOPERPACK_E5' { 'Microsoft 365 E5 Developer' }
                        'STREAM' { 'Microsoft Stream' }
                        'ENTERPRISEPREMIUM_FACULTY' { 'Microsoft 365 A5 for Faculty' }
                        'ENTERPRISEPREMIUM_STUDENT' { 'Microsoft 365 A5 for Students' }
                        'ENTERPRISEPACK_FACULTY' { 'Microsoft 365 A3 for Faculty' }
                        'ENTERPRISEPACK_STUDENT' { 'Microsoft 365 A3 for Students' }
                        default { $skuPartNumber }
                    }

                    $skus[$skuId] = [PSCustomObject]@{
                        SkuId = $skuId
                        SkuPartNumber = $skuPartNumber
                        FriendlyName = $friendlyName
                        ConsumedUnits = $sku.ConsumedUnits
                        TotalUnits = if ($sku.PrepaidUnits) { $sku.PrepaidUnits.Enabled } else { 0 }
                    }
                }
            } catch {
                Write-Warning "Could not retrieve tenant SKUs: $($_.Exception.Message). Using static mapping only."
            }
        }
        
        # Add static mappings for common SKUs (these will override tenant SKUs if they exist, or add missing ones)
        foreach ($skuGuid in $staticSkuMapping.Keys) {
            if (-not $skus.ContainsKey($skuGuid)) {
                $skus[$skuGuid] = [PSCustomObject]@{
                    SkuId = $skuGuid
                    SkuPartNumber = 'UNKNOWN'
                    FriendlyName = $staticSkuMapping[$skuGuid]
                    ConsumedUnits = 0
                    TotalUnits = 0
                }
            } else {
                # Update friendly name if static mapping exists and is more descriptive
                if ($staticSkuMapping[$skuGuid] -ne $skus[$skuGuid].FriendlyName) {
                    $skus[$skuGuid].FriendlyName = $staticSkuMapping[$skuGuid]
                }
            }
        }

        return $skus
    } catch {
        Write-Warning "Failed to get tenant license SKUs: $($_.Exception.Message). Using static mapping only."
        # Even if tenant lookup fails, return static mappings
        $skus = @{}
        foreach ($skuGuid in $staticSkuMapping.Keys) {
            $skus[$skuGuid] = [PSCustomObject]@{
                SkuId = $skuGuid
                SkuPartNumber = 'UNKNOWN'
                FriendlyName = $staticSkuMapping[$skuGuid]
                ConsumedUnits = 0
                TotalUnits = 0
            }
        }
        return $skus
    }
}

# Get user license details
function Get-UserLicenseDetails {
    param(
        [Parameter(Mandatory=$true)]
        [string]$UserPrincipalName
    )

    try {
        $user = Get-MgUser -UserId $UserPrincipalName -Property Id,UserPrincipalName,DisplayName,AssignedLicenses -ErrorAction Stop

        if (-not $user.AssignedLicenses -or $user.AssignedLicenses.Count -eq 0) {
            return [PSCustomObject]@{
                UserPrincipalName = $user.UserPrincipalName
                DisplayName = $user.DisplayName
                LicenseStatus = 'Unlicensed'
                Licenses = @()
                LicenseNames = 'None'
            }
        }

        # Get SKU mapping
        $skuMapping = Get-TenantLicenseSkus

        $licenses = @()
        foreach ($assignedLicense in $user.AssignedLicenses) {
            $skuId = $assignedLicense.SkuId
            if ($skuMapping.ContainsKey($skuId)) {
                $licenses += $skuMapping[$skuId].FriendlyName
            } else {
                $licenses += "Unknown SKU: $skuId"
            }
        }

        return [PSCustomObject]@{
            UserPrincipalName = $user.UserPrincipalName
            DisplayName = $user.DisplayName
            LicenseStatus = 'Licensed'
            Licenses = $licenses
            LicenseNames = ($licenses -join '; ')
        }
    } catch {
        Write-Error "Failed to get license details for $UserPrincipalName : $($_.Exception.Message)"
        return $null
    }
}

# Get all users with their license information
function Get-AllUsersLicenseReport {
    param(
        [Parameter(Mandatory=$false)]
        [array]$SelectedUsers = @()
    )

    try {
        # Get SKU mapping once
        $skuMapping = Get-TenantLicenseSkus

        $report = @()
        $users = @()

        if ($SelectedUsers -and $SelectedUsers.Count -gt 0) {
            # Server-side filtering: Get only selected users
            Write-Host "Processing licenses for $($SelectedUsers.Count) selected user(s)..." -ForegroundColor Cyan
            foreach ($userIdentifier in $SelectedUsers) {
                $upn = if ($userIdentifier -is [string]) { $userIdentifier } elseif ($userIdentifier.UserPrincipalName) { $userIdentifier.UserPrincipalName } else { continue }
                try {
                    $user = Get-MgUser -UserId $upn -Property Id,UserPrincipalName,DisplayName,AssignedLicenses,AccountEnabled -ErrorAction Stop
                    if ($user) {
                        $users += $user
                    }
                } catch {
                    Write-Warning "Failed to get license information for ${upn}: $($_.Exception.Message)"
                }
            }
        } else {
            # Get all users if no selection
            Write-Host "Processing licenses for all users (this may take a few minutes)..." -ForegroundColor Cyan
            $users = Get-MgUser -All -Property Id,UserPrincipalName,DisplayName,AssignedLicenses,AccountEnabled -ErrorAction Stop
        }

        $totalUsers = $users.Count
        $processedCount = 0

        foreach ($user in $users) {
            $processedCount++

            # Progress reporting
            if ($processedCount % 50 -eq 0) {
                Write-Progress -Activity "Processing user licenses" -Status "$processedCount of $totalUsers users processed" -PercentComplete (($processedCount / $totalUsers) * 100)
            }

            $licenses = @()
            if ($user.AssignedLicenses -and $user.AssignedLicenses.Count -gt 0) {
                foreach ($assignedLicense in $user.AssignedLicenses) {
                    $skuId = $assignedLicense.SkuId
                    if ($skuMapping.ContainsKey($skuId)) {
                        $licenses += $skuMapping[$skuId].FriendlyName
                    } else {
                        $licenses += "Unknown SKU: $skuId"
                    }
                }
            }

            $licenseStatus = if ($licenses.Count -gt 0) { 'Licensed' } else { 'Unlicensed' }
            $licenseNames = if ($licenses.Count -gt 0) { ($licenses -join '; ') } else { 'None' }

            $report += [PSCustomObject]@{
                UserPrincipalName = $user.UserPrincipalName
                DisplayName = $user.DisplayName
                AccountEnabled = $user.AccountEnabled
                LicenseStatus = $licenseStatus
                LicenseCount = $licenses.Count
                Licenses = $licenseNames
            }
        }

        Write-Progress -Activity "Processing user licenses" -Completed
        Write-Host "Processed $($report.Count) user(s)" -ForegroundColor Green
        return $report
    } catch {
        Write-Error "Failed to generate user license report: $($_.Exception.Message)"
        return @()
    }
}

# Export user license report to CSV and convert to XLSX
function Export-UserLicenseReport {
    param(
        [Parameter(Mandatory=$true)]
        [string]$OutputFolder,
        [Parameter(Mandatory=$false)]
        [array]$SelectedUsers = @()
    )

    try {
        # Create output folder if it doesn't exist
        if (-not (Test-Path $OutputFolder)) {
            New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
        }

        # Generate timestamp
        $timestamp = Get-Date -Format 'yyyy-MM-dd_HHmmss'
        $csvPath = Join-Path $OutputFolder "UserLicenses_$timestamp.csv"
        $xlsxPath = Join-Path $OutputFolder "UserLicenses_$timestamp.xlsx"

        # Get license report
        Write-Host "Gathering user license information..." -ForegroundColor Cyan
        $report = Get-AllUsersLicenseReport -SelectedUsers $SelectedUsers

        if ($report.Count -eq 0) {
            Write-Warning "No user data found to export."
            return $null
        }

        # Export to CSV
        Write-Host "Exporting to CSV: $csvPath" -ForegroundColor Cyan
        $report | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

        # Convert to XLSX
        Write-Host "Converting to Excel format..." -ForegroundColor Cyan
        $excel = $null
        $workbook = $null

        try {
            $excel = New-Object -ComObject Excel.Application -ErrorAction Stop
            $excel.Visible = $false
            $excel.DisplayAlerts = $false

            $workbook = $excel.Workbooks.Open($csvPath)
            $workbook.SaveAs($xlsxPath, 51) # 51 = xlOpenXMLWorkbook (.xlsx)

            # Format the worksheet
            $worksheet = $workbook.Worksheets.Item(1)
            $worksheet.Name = "User Licenses"

            $usedRange = $worksheet.UsedRange
            $usedRange.Columns.AutoFit() | Out-Null
            $usedRange.Rows.AutoFit() | Out-Null

            # Format header row
            $headerRow = $worksheet.Rows.Item(1)
            $headerRow.Font.Bold = $true
            $headerRow.Interior.Color = 15773696 # Light blue
            $headerRow.Font.Color = 1 # Black
            $headerRow.Borders.LineStyle = 1

            # Add filters
            $usedRange.AutoFilter() | Out-Null

            $workbook.Save()
            $workbook.Close($false)

            Write-Host "Successfully exported user license report to: $xlsxPath" -ForegroundColor Green
            Write-Host "Total users: $($report.Count)" -ForegroundColor Cyan

            return $xlsxPath
        } catch {
            Write-Warning "Failed to convert to Excel format: $($_.Exception.Message)"
            Write-Host "CSV file available at: $csvPath" -ForegroundColor Yellow
            return $csvPath
        } finally {
            if ($workbook) {
                try { $workbook.Close($false) } catch {}
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
            }
            if ($excel) {
                try { $excel.Quit() } catch {}
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
            }
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
        }
    } catch {
        Write-Error "Failed to export user license report: $($_.Exception.Message)"
        return $null
    }
}

Export-ModuleMember -Function Format-InboxRuleXlsx,New-SecurityInvestigationReport,Get-ExchangeMessageTrace,Get-ExchangeInboxRules,Get-GraphAuditLogs,Get-GraphSignInLogs,New-AISecurityInvestigationPrompt,New-TicketSecuritySummary,New-SecurityInvestigationSummary
Export-ModuleMember -Function Get-MfaCoverageReport,Get-UserSecurityGroupsReport,Export-EntraPortalSignInCsv,Get-ExchangeTransportRules,Get-ExchangeInboundConnectors,Get-ExchangeOutboundConnectors,New-SecurityInvestigationZip
Export-ModuleMember -Function Get-MailboxForwardingAndDelegation,Get-MailFlowConnectors,Get-TenantLicenseSkus,Get-UserLicenseDetails,Get-AllUsersLicenseReport,Export-UserLicenseReport
