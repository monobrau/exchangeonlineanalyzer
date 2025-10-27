# Returns:
#   @{ SecurityDefaultsEnabled = <bool>; CAPoliciesRequireMfa = <bool>; Users = <list of user objects> }
function Get-MfaCoverageReport {
    param([switch]$Parallel,[int]$ThrottleLimit = 4)
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
                # Treat any authenticationStrength grant as MFA requirement
                if (-not $requiresMfa -and $grant.authenticationStrength) { $requiresMfa = $true }
            }
            if ($requiresMfa) { $mfaPolicies += $p }
        }

        # 3) Users and per-user evaluation (use REST paging to avoid parameter-set issues)
        $users = @()
        try {
            $uri = 'https://graph.microsoft.com/v1.0/users?$select=id,displayName,userPrincipalName,accountEnabled&$top=999'
            do {
                $page = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
                if ($page.value) { $users += ($page.value | Where-Object { $_.accountEnabled -ne $false }) }
                $uri = $page.'@odata.nextLink'
            } while ($uri)

            # Directory roles map (for policy role assignment evaluation) via REST
            $roles = @(); $roleIdToName = @{}
            try {
                $ruri = 'https://graph.microsoft.com/v1.0/directoryRoles?$select=id,displayName&$top=999'
                do {
                    $rpage = Invoke-MgGraphRequest -Method GET -Uri $ruri -ErrorAction SilentlyContinue
                    if ($rpage.value) { $roles += $rpage.value }
                    $ruri = $rpage.'@odata.nextLink'
                } while ($ruri)
            } catch {}
            foreach ($r in $roles) { if ($r.Id) { $roleIdToName[$r.Id] = $r.DisplayName } }

        # Preload user registration details as fallback per-user MFA signal
        $regMap = @{}
        try {
            $ruri = 'https://graph.microsoft.com/v1.0/reports/authenticationMethods/userRegistrationDetails?$select=id,userPrincipalName,isMfaRegistered,isMfaCapable&$top=999'
            do {
                $rpage = Invoke-MgGraphRequest -Method GET -Uri $ruri -ErrorAction SilentlyContinue
                if ($rpage.value) {
                    foreach ($row in $rpage.value) {
                        if ($row.id) { $regMap[$row.id] = @($row.isMfaRegistered,$row.isMfaCapable) -contains $true }
                    }
                }
                $ruri = $rpage.'@odata.nextLink'
            } while ($ruri)
        } catch {}

        $acc = [System.Collections.Concurrent.ConcurrentBag[object]]::new()
        if ($Parallel -and $PSVersionTable.PSVersion.Major -ge 7) {
            $sec = $secDefaultsEnabled
            $pols = $mfaPolicies
            $reg = $regMap
            $computed = $users | ForEach-Object -Parallel {
                param($u,$sec,$pols,$reg)
                $directMfa = $false
                try {
                    $methods = Invoke-MgGraphRequest -Method GET -Uri ("https://graph.microsoft.com/v1.0/users/{0}/authentication/methods" -f $u.Id) -ErrorAction SilentlyContinue
                    if ($methods.value) {
                        foreach ($m in $methods.value) {
                            $otype = $m.'@odata.type'
                            if ($otype -eq '#microsoft.graph.microsoftAuthenticatorAuthenticationMethod' -or
                                $otype -eq '#microsoft.graph.phoneAuthenticationMethod' -or
                                $otype -eq '#microsoft.graph.softwareOathAuthenticationMethod' -or
                                $otype -eq '#microsoft.graph.fido2AuthenticationMethod' -or
                                $otype -eq '#microsoft.graph.temporaryAccessPassAuthenticationMethod') { $directMfa = $true; break }
                        }
                    }
                } catch {}
                if (-not $directMfa -and $reg.ContainsKey($u.Id)) { $directMfa = [bool]$reg[$u.Id] }
                $userGroups = @(); $userRoles = @()
                try {
                    $mem = @()
                    $mUri = ("https://graph.microsoft.com/v1.0/users/{0}/memberOf?$select=id,displayName&$top=999" -f $u.Id)
                    do {
                        $mResp = Invoke-MgGraphRequest -Method GET -Uri $mUri -ErrorAction SilentlyContinue
                        if ($mResp.value) { $mem += $mResp.value }
                        $mUri = $mResp.'@odata.nextLink'
                    } while ($mUri)
                    foreach ($m in $mem) {
                        if ($m.'@odata.type' -eq '#microsoft.graph.group') { $userGroups += $m.Id }
                        elseif ($m.'@odata.type' -eq '#microsoft.graph.directoryRole') { $userRoles += $m.Id }
                    }
                } catch {}
                $userCaRequiresMfa = $false
                foreach ($p in $pols) {
                    $conds = $p.conditions; if (-not $conds) { continue }
                    $usersCond = $conds.users
                    $incAll = $false; $incUser = $false; $excluded = $false
                    if ($usersCond) {
                        if ($usersCond.includeUsers -and ($usersCond.includeUsers -contains 'All' -or $usersCond.includeUsers -contains $u.Id)) { $incAll = $usersCond.includeUsers -contains 'All'; if (-not $incAll) { $incUser = $true } }
                        if (-not $incUser -and $usersCond.includeGroups) { if (@($usersCond.includeGroups) -ne $null) { if ($usersCond.includeGroups | Where-Object { $userGroups -contains $_ }) { $incUser = $true } } }
                        if (-not $incUser -and $usersCond.includeRoles)  { if (@($usersCond.includeRoles)  -ne $null) { if ($usersCond.includeRoles  | Where-Object { $userRoles  -contains $_ }) { $incUser = $true } } }
                        if ($usersCond.excludeUsers  -and ($usersCond.excludeUsers  -contains $u.Id)) { $excluded = $true }
                        if ($usersCond.excludeGroups -and (@($usersCond.excludeGroups) -ne $null)) { if ($usersCond.excludeGroups | Where-Object { $userGroups -contains $_ }) { $excluded = $true } }
                        if ($usersCond.excludeRoles  -and (@($usersCond.excludeRoles)  -ne $null)) { if ($usersCond.excludeRoles  | Where-Object { $userRoles  -contains $_ }) { $excluded = $true } }
                    }
                    if (($incAll -or $incUser) -and -not $excluded) { $userCaRequiresMfa = $true; break }
                }
                $covered = ($directMfa -or $sec -or $userCaRequiresMfa)
                [pscustomobject]@{
                    DisplayName       = $u.displayName
                    UserPrincipalName = $u.userPrincipalName
                    PerUserMfaEnabled = $directMfa
                    SecurityDefaults  = $sec
                    CARequiresMfa     = $userCaRequiresMfa
                    MfaCovered        = $covered
                }
            } -ThrottleLimit $ThrottleLimit -ArgumentList $sec,$pols,$reg
            if ($computed) { foreach($o in $computed){ $acc.Add($o) } }
        } else {
            foreach ($u in $users) {
                # sequential path reuse same logic as above (inline for clarity)
                $directMfa = $false
                try {
                    $methods = Invoke-MgGraphRequest -Method GET -Uri ("https://graph.microsoft.com/v1.0/users/{0}/authentication/methods" -f $u.Id) -ErrorAction SilentlyContinue
                    if ($methods.value) {
                        foreach ($m in $methods.value) {
                            $otype = $m.'@odata.type'
                            if ($otype -eq '#microsoft.graph.microsoftAuthenticatorAuthenticationMethod' -or
                                $otype -eq '#microsoft.graph.phoneAuthenticationMethod' -or
                                $otype -eq '#microsoft.graph.softwareOathAuthenticationMethod' -or
                                $otype -eq '#microsoft.graph.fido2AuthenticationMethod' -or
                                $otype -eq '#microsoft.graph.temporaryAccessPassAuthenticationMethod') { $directMfa = $true; break }
                        }
                    }
                } catch {}
                if (-not $directMfa -and $regMap.ContainsKey($u.Id)) { $directMfa = [bool]$regMap[$u.Id] }
                $userGroups = @(); $userRoles = @()
                try {
                    $mem = @()
                    $mUri = ("https://graph.microsoft.com/v1.0/users/{0}/memberOf?$select=id,displayName&$top=999" -f $u.Id)
                    do {
                        $mResp = Invoke-MgGraphRequest -Method GET -Uri $mUri -ErrorAction SilentlyContinue
                        if ($mResp.value) { $mem += $mResp.value }
                        $mUri = $mResp.'@odata.nextLink'
                    } while ($mUri)
                    foreach ($m in $mem) {
                        if ($m.'@odata.type' -eq '#microsoft.graph.group') { $userGroups += $m.Id }
                        elseif ($m.'@odata.type' -eq '#microsoft.graph.directoryRole') { $userRoles += $m.Id }
                    }
                } catch {}
                $userCaRequiresMfa = $false
                foreach ($p in $mfaPolicies) {
                    $conds = $p.conditions; if (-not $conds) { continue }
                    $usersCond = $conds.users
                    $incAll = $false; $incUser = $false; $excluded = $false
                    if ($usersCond) {
                        if ($usersCond.includeUsers -and ($usersCond.includeUsers -contains 'All' -or $usersCond.includeUsers -contains $u.Id)) { $incAll = $usersCond.includeUsers -contains 'All'; if (-not $incAll) { $incUser = $true } }
                        if (-not $incUser -and $usersCond.includeGroups) { if (@($usersCond.includeGroups) -ne $null) { if ($usersCond.includeGroups | Where-Object { $userGroups -contains $_ }) { $incUser = $true } } }
                        if (-not $incUser -and $usersCond.includeRoles) { if (@($usersCond.includeRoles) -ne $null) { if ($usersCond.includeRoles | Where-Object { $userRoles -contains $_ }) { $incUser = $true } } }
                        if ($usersCond.excludeUsers -and ($usersCond.excludeUsers -contains $u.Id)) { $excluded = $true }
                        if ($usersCond.excludeGroups) { if (@($usersCond.excludeGroups) -ne $null) { if ($usersCond.excludeGroups | Where-Object { $userGroups -contains $_ }) { $excluded = $true } } }
                        if ($usersCond.excludeRoles) { if (@($usersCond.excludeRoles) -ne $null) { if ($usersCond.excludeRoles | Where-Object { $userRoles -contains $_ }) { $excluded = $true } } }
                    }
                    $applies = ($incAll -or $incUser)
                    if ($applies -and -not $excluded) { $userCaRequiresMfa = $true; break }
                }
                $covered = ($directMfa -or $secDefaultsEnabled -or $userCaRequiresMfa)
                $acc.Add([pscustomobject]@{
                    DisplayName       = $u.displayName
                    UserPrincipalName = $u.userPrincipalName
                    PerUserMfaEnabled = $directMfa
                    SecurityDefaults  = $secDefaultsEnabled
                    CARequiresMfa     = $userCaRequiresMfa
                    MfaCovered        = $covered
                }) | Out-Null
            }
        }
        $users = [System.Collections.ArrayList]$acc
        } catch {}

        $tenantLevelCaMfa = ($mfaPolicies.Count -gt 0)
        return @{ SecurityDefaultsEnabled = $secDefaultsEnabled; CAPoliciesRequireMfa = $tenantLevelCaMfa; Users = $users }
    } catch {
        Write-Error "Get-MfaCoverageReport failed: $($_.Exception.Message)"; return @{ SecurityDefaultsEnabled=$false; CAPoliciesRequireMfa=$false; Users=@() }
    }
}

# Flattens user membership in directory roles and security groups
function Get-UserSecurityGroupsReport {
    param([switch]$Parallel,[int]$ThrottleLimit = 6)
    try {
        $results = New-Object System.Collections.Generic.List[object]

        function Invoke-GraphApiInternal {
            param([Parameter(Mandatory=$true)][string]$Uri)
            try {
                $token = $null
                try { $ctx = Get-MgContext -ErrorAction SilentlyContinue; if ($ctx -and $ctx.AccessToken) { $token = $ctx.AccessToken } } catch {}
                if ($token) {
                    $headers = @{ Authorization = ("Bearer {0}" -f $token) }
                    return Invoke-RestMethod -Method Get -Uri $Uri -Headers $headers -ErrorAction Stop
                } else {
                    return Invoke-MgGraphRequest -Method GET -Uri $Uri -ErrorAction Stop
                }
            } catch { throw }
        }

        # Directory roles (e.g., Global Administrator) via REST to avoid parameter set issues
        $roles = @(); $roleIdToName = @{}
        try {
            $uri = 'https://graph.microsoft.com/v1.0/directoryRoles?$select=id,displayName&$top=999'
            do {
                $resp = Invoke-GraphApiInternal -Uri $uri
                if ($resp.value) { $roles += $resp.value }
                $uri = $resp.'@odata.nextLink'
            } while ($uri)
        } catch { $roles = @() }
        foreach ($r in $roles) { if ($r.Id) { $roleIdToName[$r.Id] = $r.DisplayName } }

        # Elevated/privileged role names (include legacy names)
        $highPrivilegeRoles = @(
            'Global Administrator','Company Administrator','Privileged Role Administrator','Privileged Authentication Administrator',
            'Exchange Administrator','SharePoint Administrator','Security Administrator','Compliance Administrator',
            'User Administrator','Billing Administrator','Helpdesk Administrator','Service Support Administrator',
            'Teams Administrator','Intune Administrator','Application Administrator','Cloud Application Administrator',
            'Power Platform Administrator'
        )

        # Users (REST paging to avoid -All parameter set issues)
        $users = @()
        try {
            $uri = 'https://graph.microsoft.com/v1.0/users?$select=id,displayName,userPrincipalName,accountEnabled&$top=999'
            do {
                $resp = Invoke-GraphApiInternal -Uri $uri
                if ($resp.value) { $users += ($resp.value | Where-Object { $_.accountEnabled -ne $false }) }
                $uri = $resp.'@odata.nextLink'
            } while ($uri)
        } catch { $users = @() }

        $processUser = {
            param($u,$roleIdToName,$highPrivilegeRoles,$results)
            $roleNames = @(); $groupNames = @()
            try {
                # Fetch memberOf via REST with paging; includes groups and directory roles
                $mem = @()
                $mUri = ("https://graph.microsoft.com/v1.0/users/{0}/memberOf?$select=id,displayName&$top=999" -f $u.Id)
                do {
                $mResp = $null; try { $mResp = Invoke-GraphApiInternal -Uri $mUri } catch {}
                    if ($mResp.value) { $mem += $mResp.value }
                    $mUri = $mResp.'@odata.nextLink'
                } while ($mUri)
                foreach ($m in $mem) {
                    if ($m.'@odata.type' -eq '#microsoft.graph.group') { if ($m.DisplayName) { $groupNames += $m.DisplayName } }
                    elseif ($m.'@odata.type' -eq '#microsoft.graph.directoryRole') {
                        $rname = $null
                        if ($roleIdToName.ContainsKey($m.Id)) { $rname = $roleIdToName[$m.Id] }
                        if (-not $rname -and $m.DisplayName) { $rname = $m.DisplayName }
                        if ($rname) { $roleNames += $rname }
                    }
                }
            } catch {}
            $roleNames = $roleNames | Sort-Object -Unique
            $groupNames = $groupNames | Sort-Object -Unique
            $elevated = @($roleNames | Where-Object { $highPrivilegeRoles -contains $_ })
            $results.Add([pscustomobject]@{
                DisplayName        = $u.DisplayName
                UserPrincipalName  = $u.UserPrincipalName
                Roles              = ($roleNames -join '; ')
                Groups             = ($groupNames -join '; ')
                ElevatedRoles      = ($elevated -join '; ')
                IsElevated         = [bool]($elevated -and $elevated.Count -gt 0)
            }) | Out-Null
        }

        # Use sequential processing for compatibility across module versions
        foreach ($u in $users) { & $processUser $u $roleIdToName $highPrivilegeRoles $results }

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
        [object]$ProgressBar
    )

    # Local helper: update progress UI if provided
    function Set-ReportProgress {
        param([int]$Percent,[string]$Text)
        try {
            if ($StatusLabel -and $StatusLabel.GetType().Name -eq "Label" -and $Text) { $StatusLabel.Text = $Text }
            if ($ProgressBar -and $ProgressBar.GetType().Name -eq "ProgressBar") {
                try {
                    if ($ProgressBar.Style -ne [System.Windows.Forms.ProgressBarStyle]::Blocks) { $ProgressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Blocks }
                    if ($ProgressBar.Maximum -ne 100) { $ProgressBar.Maximum = 100 }
                    if (-not $ProgressBar.Visible) { $ProgressBar.Visible = $true }
                    if ($Percent -lt 0) { $Percent = 0 } elseif ($Percent -gt 100) { $Percent = 100 }
                    $ProgressBar.Value = $Percent
                } catch {}
            }
            if ($MainForm -and $MainForm.GetType().Name -eq "Form") { [System.Windows.Forms.Application]::DoEvents() }
        } catch {}
    }

    try {
        if ($StatusLabel -and $StatusLabel.GetType().Name -eq "Label") {
            $StatusLabel.Text = "Starting comprehensive security investigation..."
        }
        if ($MainForm -and $MainForm.GetType().Name -eq "Form") {
            $MainForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        }
        Set-ReportProgress -Percent 2 -Text "Initializing..."
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
    $report.SourceStatus = @{ ExchangeConnected = $false; GraphConnected = $false }

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
    Set-ReportProgress -Percent 5 -Text "Resolving output paths..."

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
    $report.SourceStatus.ExchangeConnected = $exchangeConnected
    $report.SourceStatus.GraphConnected    = $graphConnected

    # Collect data from Exchange Online
    if ($exchangeConnected) {
        try {
            Set-ReportProgress -Percent 10 -Text "Collecting message trace (last $DaysBack days)..."
            $report.MessageTrace = Get-ExchangeMessageTrace -DaysBack 10 -Parallel -ThrottleLimit 3 # always 10 days per requirement

            Set-ReportProgress -Percent 40 -Text "Exporting inbox rules..."
            $report.InboxRules = Get-ExchangeInboxRules -Parallel -ThrottleLimit 6
            try {
                $ri = $report.InboxRules
                $rc = if ($ri) { $ri.Count } else { 0 }
                $mbCount = if ($ri) { ($ri | Select-Object -ExpandProperty MailboxOwner -Unique).Count } else { 0 }
                Write-Host ("Inbox rules collected via Exchange Online: {0} rules across {1} mailbox(es)" -f $rc, $mbCount) -ForegroundColor Green
            } catch {}

            Set-ReportProgress -Percent 50 -Text "Collecting transport rules..."
            $report.TransportRules = Get-ExchangeTransportRules

            Set-ReportProgress -Percent 60 -Text "Collecting mail flow connectors..."
            $report.InboundConnectors = Get-ExchangeInboundConnectors
            $report.OutboundConnectors = Get-ExchangeOutboundConnectors
        } catch {
            Write-Warning "Failed to collect Exchange Online data: $($_.Exception.Message)"
            $report.ExchangeDataError = $_.Exception.Message
        }
    }

    # Collect data from Microsoft Graph (audit logs only)
    if ($graphConnected) {
        try {
            Set-ReportProgress -Percent 70 -Text "Collecting audit logs from Microsoft Graph..."
            $report.AuditLogs = Get-GraphAuditLogs -DaysBack $DaysBack
        } catch {
            Write-Warning "Failed to collect Microsoft Graph data: $($_.Exception.Message)"
            $report.GraphDataError = $_.Exception.Message
        }
    }

    # MFA Coverage and User Security Groups
    if ($graphConnected) {
        try {
            Set-ReportProgress -Percent 78 -Text "Evaluating MFA coverage..."
            try {
                $report.MfaCoverage = Get-MfaCoverageReport -Parallel -ThrottleLimit 4
            } catch {
                $report.MfaCoverage = Get-MfaCoverageReport
            }
            try {
                $uc = if ($report.MfaCoverage -and $report.MfaCoverage.Users) { $report.MfaCoverage.Users.Count } else { 0 }
                Write-Host ("MFA analysis via Microsoft Graph: evaluated {0} user(s)" -f $uc) -ForegroundColor Green
            } catch {}

            Set-ReportProgress -Percent 84 -Text "Collecting user security groups and roles..."
            # Capability check: enable parallel only if Graph supports required switches cleanly
            $canParallelGroups = $false
            try {
                $cmd = Get-Command Get-MgUserMemberOf -ErrorAction SilentlyContinue
                if ($cmd -and $cmd.Parameters.ContainsKey('All')) { $canParallelGroups = $true }
            } catch {}
            if ($canParallelGroups -and $PSVersionTable.PSVersion.Major -ge 7) {
                try { $report.UserSecurityGroups = Get-UserSecurityGroupsReport -Parallel -ThrottleLimit 6 }
                catch { $report.UserSecurityGroups = Get-UserSecurityGroupsReport }
            } else {
                $report.UserSecurityGroups = Get-UserSecurityGroupsReport
            }
            try {
                $ugc = if ($report.UserSecurityGroups) { $report.UserSecurityGroups.Count } else { 0 }
                Write-Host ("User security groups via Microsoft Graph: {0} user(s)" -f $ugc) -ForegroundColor Green
            } catch {}
        } catch {
            Write-Warning "Failed to build MFA/Groups reports: $($_.Exception.Message)"
        }
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
    if ($report.OutputFolder) {
        $exportError = $null
        try {
            $out = $report.OutputFolder
            $exportItems = @(
                @{ Data=$report.MessageTrace;       Csv='MessageTrace.csv';        Json='MessageTrace.json';        Depth=8 },
                @{ Data=$report.InboxRules;         Csv='InboxRules.csv';          Json='InboxRules.json';          Depth=6 },
                @{ Data=$report.TransportRules;     Csv='TransportRules.csv';      Json='TransportRules.json';      Depth=8 },
                @{ Data=$report.InboundConnectors;  Csv='InboundConnectors.csv';   Json='InboundConnectors.json';   Depth=8 },
                @{ Data=$report.OutboundConnectors; Csv='OutboundConnectors.csv';  Json='OutboundConnectors.json';  Depth=8 },
                @{ Data=$report.AuditLogs;          Csv='GraphAuditLogs.csv';      Json='GraphAuditLogs.json';      Depth=8 }
            )

            $exportAction = {
                param($item,$out)
            $csvPath  = Join-Path $out $item.Csv
            $jsonPath = Join-Path $out $item.Json
            if ($item.Data -and $item.Data.Count -gt 0) {
                try { $item.Data | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8 }
                catch { $item.Data | ConvertTo-Json -Depth $item.Depth | Out-File -FilePath $jsonPath -Encoding utf8 }
            } else {
                # Ensure at least a JSON file exists for empty datasets
                '[]' | Out-File -FilePath $jsonPath -Encoding utf8
            }
            }

            if ($PSVersionTable.PSVersion.Major -ge 7) {
                $exportItems | ForEach-Object -Parallel {
                    $item = $_
                    $csvPath  = Join-Path $using:out $item.Csv
                    $jsonPath = Join-Path $using:out $item.Json
                    if ($item -and $item.Data -and $item.Data.Count -gt 0) {
                        try { $item.Data | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8 }
                        catch { $item.Data | ConvertTo-Json -Depth $item.Depth | Out-File -FilePath $jsonPath -Encoding utf8 }
                    } else {
                        '[]' | Out-File -FilePath $jsonPath -Encoding utf8
                    }
                } -ThrottleLimit 4
            } else {
                foreach ($it in $exportItems) { & $exportAction $it $out }
            }

            # MFA Coverage export (ensure file exists even if empty)
            $mfaCsv = Join-Path $report.OutputFolder "MFAStatus.csv"
            if ($report.MfaCoverage -and $report.MfaCoverage.Users -and $report.MfaCoverage.Users.Count -gt 0) {
                try { $report.MfaCoverage.Users | Export-Csv -Path $mfaCsv -NoTypeInformation -Encoding UTF8; $report.FilePaths.MFAStatusCsv = $mfaCsv } catch {}
            } else {
                $mfaHeader = 'DisplayName,UserPrincipalName,PerUserMfaEnabled,SecurityDefaults,CARequiresMfa,MfaCovered'
                try { Set-Content -Path $mfaCsv -Value $mfaHeader -Encoding utf8; $report.FilePaths.MFAStatusCsv = $mfaCsv } catch {}
            }

            # User Security Groups export (ensure file exists even if empty)
            $usgCsv = Join-Path $report.OutputFolder "UserSecurityGroups.csv"
            if ($report.UserSecurityGroups -and $report.UserSecurityGroups.Count -gt 0) {
                try { $report.UserSecurityGroups | Export-Csv -Path $usgCsv -NoTypeInformation -Encoding UTF8; $report.FilePaths.UserSecurityGroupsCsv = $usgCsv } catch {}
            } else {
                $usgHeader = 'DisplayName,UserPrincipalName,Roles,Groups,ElevatedRoles,IsElevated'
                try { Set-Content -Path $usgCsv -Value $usgHeader -Encoding utf8; $report.FilePaths.UserSecurityGroupsCsv = $usgCsv } catch {}
            }
        } catch { $exportError = $_ }

        # Ensure InboxRules.csv exists even if no data
        try {
            $inboxCsv = Join-Path $report.OutputFolder "InboxRules.csv"
            if (-not (Test-Path $inboxCsv)) {
                $inboxHeader = 'MailboxOwner,Name,Enabled,Priority,FromAddressContains,SubjectContains,SentTo,RedirectTo,ForwardTo,ForwardAsAttachment,DeleteMessage,StopProcessing,IsHidden,Description'
                Set-Content -Path $inboxCsv -Value $inboxHeader -Encoding utf8
            }
        } catch {}

        # Save only LLM instructions as TXT (no other text files on disk)
        try {
            $report.LLMInstructions = New-LLMInvestigationInstructions -Report $report
            $llmPath = Join-Path $report.OutputFolder "LLM_Instructions.txt"
            if ($report.LLMInstructions) { $report.LLMInstructions | Out-File -FilePath $llmPath -Encoding utf8 }
            $report.FilePaths.LLMInstructionsTxt = $llmPath
        } catch {}

        # Create a ZIP bundle of all exported files EXCEPT the LLM instructions, for easy ticket upload
        try {
            $zipPath = Join-Path $report.OutputFolder "SecurityInvestigation_ForensicBundle.zip"
            # Collect files to include: everything in the folder except LLM_Instructions.txt and existing zip(s)
            $includeFiles = Get-ChildItem -Path $report.OutputFolder -File -ErrorAction SilentlyContinue |
                Where-Object { $_.Name -ne 'LLM_Instructions.txt' -and $_.Extension -ne '.zip' }
            if ($includeFiles -and $includeFiles.Count -gt 0) {
                # Use Compress-Archive to create the zip; pass explicit file paths
                if (Test-Path $zipPath) { Remove-Item $zipPath -Force -ErrorAction SilentlyContinue }
                Compress-Archive -Path ($includeFiles | ForEach-Object { $_.FullName }) -DestinationPath $zipPath -Force -ErrorAction Stop
                $report.FilePaths.ForensicZip = $zipPath
            }
        } catch {
            Write-Warning ("Failed to create forensic ZIP bundle: {0}" -f $_.Exception.Message)
        }
    }

    return $report
}

function Get-ExchangeMessageTrace {
	param([int]$DaysBack = 10,[switch]$Parallel,[int]$ThrottleLimit = 3)

    try {
        Write-Host "Collecting message trace data..." -ForegroundColor Yellow
		$utcNow = (Get-Date).ToUniversalTime()
		# Include today's partial day plus prior full days
		$windowCount = [Math]::Max(1,$DaysBack)
		$firstDay = $utcNow.Date.AddDays(-($windowCount - 1))

		$results = New-Object System.Collections.Generic.List[object]

		$hasV2 = $null -ne (Get-Command Get-MessageTraceV2 -ErrorAction SilentlyContinue)

		# Sequential per-day windows for prior days
		for ($i = 0; $i -lt ($windowCount - 1); $i++) {
			$winStart = $firstDay.AddDays($i)
			$winEnd = $winStart.AddDays(1)
			try {
				if ($hasV2) {
					# Seek-like pagination via StartingRecipientAddress; emit only strictly-greater recipients to avoid duplicates
					$cursor = $null
					$iterations = 0
					do {
						$params = @{ StartDate = $winStart; EndDate = $winEnd; ErrorAction = 'Stop'; ResultSize = 1000 }
						if ($cursor) { $params.StartingRecipientAddress = $cursor }
						$chunk = Get-MessageTraceV2 @params
						if ($chunk) {
							$emit = if ($cursor) { $chunk | Where-Object { $_.RecipientAddress -gt $cursor } } else { $chunk }
							if ($emit) { [void]$results.AddRange($emit) }
							$nextCursor = ($chunk | Where-Object { -not $cursor -or $_.RecipientAddress -gt $cursor } | Select-Object -Last 1).RecipientAddress
							if (-not $nextCursor -or ($cursor -and $nextCursor -le $cursor)) { break }
							$cursor = $nextCursor
						}
						$iterations++
					} while ($chunk -and $iterations -lt 500)
				} else {
					$batch = Get-MessageTrace -StartDate $winStart -EndDate $winEnd -ErrorAction Stop
					if ($batch) { [void]$results.AddRange($batch) }
				}
			} catch {}
		}

		# Final window: today's partial day up to current time
		$winStart = $utcNow.Date
		$winEnd = $utcNow
		try {
			if ($hasV2) {
				$cursor = $null
				$iterations = 0
				do {
					$params = @{ StartDate = $winStart; EndDate = $winEnd; ErrorAction = 'Stop'; ResultSize = 1000 }
					if ($cursor) { $params.StartingRecipientAddress = $cursor }
					$chunk = Get-MessageTraceV2 @params
					if ($chunk) {
						$emit = if ($cursor) { $chunk | Where-Object { $_.RecipientAddress -gt $cursor } } else { $chunk }
						if ($emit) { [void]$results.AddRange($emit) }
						$nextCursor = ($chunk | Where-Object { -not $cursor -or $_.RecipientAddress -gt $cursor } | Select-Object -Last 1).RecipientAddress
						if (-not $nextCursor -or ($cursor -and $nextCursor -le $cursor)) { break }
						$cursor = $nextCursor
					}
					$iterations++
				} while ($chunk -and $iterations -lt 500)
			} else {
				$batch = Get-MessageTrace -StartDate $winStart -EndDate $winEnd -ErrorAction Stop
				if ($batch) { [void]$results.AddRange($batch) }
			}
		} catch {}

		return [System.Collections.ArrayList]$results
    } catch {
        Write-Error "Failed to collect message trace: $($_.Exception.Message)"
        return @()
    }
}

function Get-ExchangeInboxRules {
    param([switch]$Parallel,[int]$ThrottleLimit = 6)
    try {
        Write-Host "Exporting inbox rules..." -ForegroundColor Yellow

        $mailboxes = @()
        try {
            $mailboxes = Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox,SharedMailbox -ErrorAction Stop
        } catch {
            # Fallbacks if Get-Mailbox is restricted or not available in current session
            try { $mailboxes = Get-Mailbox -ResultSize 2000 -ErrorAction Stop }
            catch {
                try { $mailboxes = Get-EXOMailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox,SharedMailbox -ErrorAction Stop }
                catch { try { $mailboxes = Get-EXOMailbox -ResultSize 2000 -ErrorAction Stop } catch { $mailboxes = @() } }
            }
        }

        $allRules = New-Object System.Collections.Generic.List[object]

        if ($Parallel -and $PSVersionTable.PSVersion.Major -ge 7) {
            $computed = $mailboxes | ForEach-Object -Parallel {
                param($mbx)
                $output = @()
                $upn = if ($mbx.UserPrincipalName) { $mbx.UserPrincipalName } else { $mbx.PrimarySmtpAddress }
                try {
                    $rules = Get-InboxRule -Mailbox $upn -ErrorAction Stop
                    foreach ($r in $rules) {
                        $output += [pscustomobject]@{
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
                    }
                } catch {}
                $output
            } -ThrottleLimit $ThrottleLimit
            if ($computed) {
                foreach ($o in $computed) {
                    if ($o -is [System.Array]) { foreach ($e in $o) { if ($e) { [void]$allRules.Add($e) } } }
                    elseif ($o) { [void]$allRules.Add($o) }
                }
            }
        } else {
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
    param([int]$DaysBack = 10)

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
        $page = Get-MgAuditLogDirectoryAudit -All -Filter "activityDateTime ge $startIso" -ErrorAction Stop
        if ($page) { [void]$raw.AddRange($page) }

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

        return [System.Collections.ArrayList]$flattened
    } catch {
        Write-Error "Failed to collect audit logs: $($_.Exception.Message)"
        return @()
    }
}

function Get-GraphSignInLogs { param([int]$DaysBack = 10,[switch]$MaxAvailable) return @() }

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
- MFAStatus.csv: Identify users not covered by any MFA control; prioritize remediation.
- UserSecurityGroups.csv: Validate privileged group/role membership (e.g., Global Administrator).

Important: Sign-in logs require Entra ID Premium for API access. Please export sign-in CSVs from the Entra portal (Sign-in logs → Download, last 7–30 days depending on tenant) and include alongside these files for full analysis.

*This automated security analysis helps identify potential security incidents and unusual patterns that may require further investigation by security professionals.*
"@

    return $message
}

function New-LLMInvestigationInstructions {
    param([Parameter(Mandatory=$true)]$Report)

    $investigator = $Report.Investigator
    $company = $Report.Company
    $days = $Report.DaysAnalyzed

    $instructions = @"
You are an incident responder assisting $investigator at $company.

Goal: Produce a concise investigation report for a non-technical audience, suitable as a message to the client’s technical contact in our ticketing system.

Input files (provided separately):
- MessageTrace.csv (last $days days)
- InboxRules.csv
- AuditLogs.csv
- MFAStatus.csv
- UserSecurityGroups.csv
- Optional: Sign-in logs CSV exported from the Entra portal (if provided)

Required output:
1) Executive Investigation Summary
   - Brief description of the suspected compromise and current status
   - Key evidence cited from the provided files
   - Timeline of events (chronological) using exact timestamps and sources

2) Findings (Non-Technical)
   - Clear list of findings with minimal jargon
   - Avoid assumptions; only state what evidence supports

3) Recommendations (Minimal)
   - Only immediate, essential actions
   - Defer broader hardening guidance for a separate follow-up

Rules:
- Do not speculate; do not fill gaps without explicit evidence
- Reference evidence by file and row attributes when possible
- Keep the message ready to paste into a ticketing system
- No code blocks unless quoting short data lines for clarity

Format:
Subject: Investigation Update – $company (Timeline + Key Findings)

Body:
1. Executive Summary
2. Timeline of Events
3. Key Findings (Evidence-Backed)
4. Immediate Next Steps (Minimal)

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
- **Connection Status:** $($Report.GraphConnection)

## Investigation Tools and Methods

### Email Security Analysis
- **Message Trace Review:** Analyzed all email sent/received patterns
- **Inbox Rule Audit:** Examined automated email processing rules
- **External Communication Patterns:** Identified unusual external email flows
- **Forwarding Rule Detection:** Flagged rules forwarding to external domains

### Authentication Analysis
- Replaced sign-in log analysis with MFA coverage and security posture review

### Administrative Activity Review
- **Privilege Changes:** Monitored role assignments and permission modifications
- **Account Management:** Tracked account creation, modification, and deletion
- **Security Policy Changes:** Reviewed authentication and access policy updates
- **Audit Trail Analysis:** Examined all administrative actions with timestamps

## Key Findings and Recommendations

### Immediate Actions Required
1. **MFA Coverage Gaps:** Remediate users not covered by per-user MFA, Security Defaults, or Conditional Access
2. **Audit Email Forwarding Rules:** Verify all external forwarding is authorized
3. **Examine Privilege Changes:** Confirm recent role assignments are legitimate
4. **Monitor External Communications:** Review patterns to unusual external domains

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
- **Sign-in Logs:** CSV format with authentication details and risk assessments

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

Export-ModuleMember -Function Format-InboxRuleXlsx,New-SecurityInvestigationReport,Get-ExchangeMessageTrace,Get-ExchangeInboxRules,Get-GraphAuditLogs,Get-GraphSignInLogs,New-AISecurityInvestigationPrompt,New-TicketSecuritySummary,New-SecurityInvestigationSummary
Export-ModuleMember -Function Get-MfaCoverageReport,Get-UserSecurityGroupsReport,Export-EntraPortalSignInCsv,Get-ExchangeTransportRules,Get-ExchangeInboundConnectors,Get-ExchangeOutboundConnectors
