# Returns:
#   @{ SecurityDefaultsEnabled = <bool>; CAPoliciesRequireMfa = <bool>; Users = <list of user objects> }
function Get-MfaCoverageReport {
    try {
        Write-Host "Analyzing tenant MFA coverage..." -ForegroundColor Yellow

        # 1) Security Defaults status (authoritative)
        $secDefaultsEnabled = $false
        try {
            $secDefaults = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/policies/identitySecurityDefaultsEnforcementPolicy' -ErrorAction Stop
            if ($secDefaults -and $secDefaults.isEnabled -ne $null) { $secDefaultsEnabled = [bool]$secDefaults.isEnabled }
        } catch {}

        # Check for federated domains (third-party MFA)
        $federatedDomains = @()
        $federatedDomainsDetails = @{}
        try {
            $allDomains = Get-MgDomain -All -ErrorAction SilentlyContinue
            foreach ($domain in $allDomains) {
                if ($domain.AuthenticationType -eq "Federated") {
                    $federatedDomains += $domain.Id

                    # Try to identify the federation provider
                    try {
                        $fedSettings = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/domains/$($domain.Id)/federationConfiguration" -ErrorAction SilentlyContinue
                        if ($fedSettings.value -and $fedSettings.value.Count -gt 0) {
                            $issuerUri = $fedSettings.value[0].issuerUri
                            $provider = "Unknown"

                            if ($issuerUri -match "duosecurity\.com") {
                                $provider = "Duo Security"
                            } elseif ($issuerUri -match "okta\.com") {
                                $provider = "Okta"
                            } elseif ($issuerUri -match "pingidentity\.com|pingone\.com") {
                                $provider = "Ping Identity"
                            } elseif ($issuerUri -match "auth0\.com") {
                                $provider = "Auth0"
                            } elseif ($issuerUri -match "onelogin\.com") {
                                $provider = "OneLogin"
                            } elseif ($issuerUri -match "adfs") {
                                $provider = "AD FS"
                            } else {
                                $provider = "External IdP"
                            }

                            $federatedDomainsDetails[$domain.Id] = $provider
                        } else {
                            $federatedDomainsDetails[$domain.Id] = "Federated (Unknown Provider)"
                        }
                    } catch {
                        $federatedDomainsDetails[$domain.Id] = "Federated"
                    }
                }
            }

            if ($federatedDomains.Count -gt 0) {
                Write-Host "Found $($federatedDomains.Count) federated domain(s) - third-party MFA may be in use" -ForegroundColor Cyan
                foreach ($fd in $federatedDomains) {
                    $provider = $federatedDomainsDetails[$fd]
                    Write-Host "  - $fd : $provider" -ForegroundColor Gray
                }
            }
        } catch {}

        # Cache enterprise applications to detect third-party MFA providers (Duo, Okta, etc.)
        $thirdPartyMfaApps = @{}
        try {
            Write-Host "Scanning for third-party MFA applications..." -ForegroundColor Yellow
            $enterpriseApps = Get-MgServicePrincipal -All -Property "DisplayName,AppId,ServicePrincipalType" -ErrorAction SilentlyContinue
            if ($enterpriseApps) {
                foreach ($app in $enterpriseApps) {
                    $displayName = $app.DisplayName.ToLower()
                    # Detect common third-party MFA providers
                    if ($displayName -match "duo|okta|ping|auth0|onelogin|rsa|symantec|thales") {
                        $provider = "Unknown"
                        if ($displayName -match "duo") { $provider = "Duo Security" }
                        elseif ($displayName -match "okta") { $provider = "Okta" }
                        elseif ($displayName -match "ping") { $provider = "Ping Identity" }
                        elseif ($displayName -match "auth0") { $provider = "Auth0" }
                        elseif ($displayName -match "onelogin") { $provider = "OneLogin" }
                        elseif ($displayName -match "rsa") { $provider = "RSA SecurID" }
                        elseif ($displayName -match "symantec") { $provider = "Symantec VIP" }
                        elseif ($displayName -match "thales") { $provider = "Thales" }

                        $thirdPartyMfaApps[$app.AppId] = @{
                            DisplayName = $app.DisplayName
                            Provider = $provider
                            AppId = $app.AppId
                        }
                    }
                }
            }
            if ($thirdPartyMfaApps.Count -gt 0) {
                $providerNames = ($thirdPartyMfaApps.Values | Select-Object -ExpandProperty Provider -Unique) -join ', '
                Write-Host "Found $($thirdPartyMfaApps.Count) third-party MFA application(s): $providerNames" -ForegroundColor Cyan
            }
        } catch {
            # If we can't get enterprise apps, continue with other checks
        }

        # 2) Conditional Access policies requiring MFA (tenant-wide set)
        $caPolicies = @()
        try {
            $resp = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies?$top=999' -ErrorAction SilentlyContinue
            if ($resp.value) { $caPolicies = $resp.value }
        } catch {}

        # Filter enabled policies that require MFA
        $mfaPolicies = @()
        $authStrengthCache = @{}

        foreach ($p in $caPolicies) {
            $enabled = ($p.state -eq 'enabled')
            if (-not $enabled) { continue }
            $grant = $p.grantControls
            $requiresMfa = $false
            $hasThirdPartyMfaIndicator = $false

            if ($grant) {
                # Check built-in MFA control
                if ($grant.builtInControls -contains 'mfa') { $requiresMfa = $true }

                # Check authentication strength (newer method)
                if ($grant.authenticationStrength -and $grant.authenticationStrength.id) {
                    $requiresMfa = $true
                    # Cache authentication strength details
                    if (-not $authStrengthCache.ContainsKey($grant.authenticationStrength.id)) {
                        try {
                            $strength = Get-MgPolicyAuthenticationStrengthPolicy -AuthenticationStrengthPolicyId $grant.authenticationStrength.id -ErrorAction SilentlyContinue
                            if ($strength) {
                                $authStrengthCache[$grant.authenticationStrength.id] = $strength

                                # Check if authentication strength name indicates third-party MFA
                                $strengthName = $strength.DisplayName.ToLower()
                                if ($strengthName -match "duo|okta|ping|auth0|onelogin|rsa|symantec|thales") {
                                    $hasThirdPartyMfaIndicator = $true
                                }
                            }
                        } catch {}
                    } else {
                        # Use cached strength to check for third-party MFA
                        $cachedStrength = $authStrengthCache[$grant.authenticationStrength.id]
                        if ($cachedStrength) {
                            $strengthName = $cachedStrength.DisplayName.ToLower()
                            if ($strengthName -match "duo|okta|ping|auth0|onelogin|rsa|symantec|thales") {
                                $hasThirdPartyMfaIndicator = $true
                            }
                        }
                    }
                }

                # Check for custom controls (third-party MFA like Duo) - deprecated but still check
                if ($grant.customAuthenticationFactors -and $grant.customAuthenticationFactors.Count -gt 0) {
                    $hasThirdPartyMfaIndicator = $true
                }

                # Check for Terms of Use (sometimes used with third-party MFA)
                if ($grant.termsOfUse -and $grant.termsOfUse.Count -gt 0) {
                    $hasThirdPartyMfaIndicator = $true
                }
            }

            # Check for third-party MFA application targeting in CA policy conditions
            if ($p.conditions.applications -and $thirdPartyMfaApps.Count -gt 0) {
                $includeApps = $p.conditions.applications.includeApplications
                if ($includeApps) {
                    foreach ($appId in $includeApps) {
                        if ($thirdPartyMfaApps.ContainsKey($appId)) {
                            $hasThirdPartyMfaIndicator = $true
                            break
                        }
                    }
                }
            }

            # Check for authentication context (newer method for third-party MFA)
            if ($p.conditions.authenticationContextClassReferences -and $p.conditions.authenticationContextClassReferences.Count -gt 0) {
                if ($requiresMfa -and $thirdPartyMfaApps.Count -gt 0) {
                    $hasThirdPartyMfaIndicator = $true
                }
            }

            # Add metadata about third-party MFA detection
            if ($hasThirdPartyMfaIndicator) {
                Add-Member -InputObject $p -MemberType NoteProperty -Name 'HasThirdPartyMfaIndicator' -Value $true -Force
            }

            if ($requiresMfa) { $mfaPolicies += $p }
        }

        Write-Host "Found $($mfaPolicies.Count) CA policies requiring MFA" -ForegroundColor Cyan

        # 3) Load all users with optimized batching
        $users = @()
        try {
            Write-Host "Loading all users..." -ForegroundColor Yellow
            $userPage = Get-MgUser -All -Property 'id,displayName,userPrincipalName' -ConsistencyLevel eventual -ErrorAction Stop
            Write-Host "Loaded $($userPage.Count) users" -ForegroundColor Cyan

            # Directory roles map (for policy role assignment evaluation)
            $roles = @(); $roleIdToName = @{}
            try {
                $roles = Get-MgDirectoryRole -All -ErrorAction SilentlyContinue
                foreach ($r in $roles) { $roleIdToName[$r.Id] = $r.DisplayName }
            } catch {}

            # Cache privileged role IDs for warning detection
            $privilegedRoles = @("Global Administrator", "Privileged Role Administrator", "Security Administrator", "Exchange Administrator", "User Administrator")
            $privilegedRoleIds = @()
            foreach ($role in $roles) {
                if ($privilegedRoles -contains $role.DisplayName) {
                    $privilegedRoleIds += $role.Id
                }
            }

            # Process users with progress indicator
            $userCount = 0
            $totalUsers = ($userPage | Measure-Object).Count

            foreach ($u in $userPage) {
                $userCount++
                if ($userCount % 100 -eq 0) {
                    Write-Host "Processing user $userCount of $totalUsers..." -ForegroundColor Gray
                }

                # Check if user's domain is federated (third-party MFA)
                $isFederated = $false
                $federatedProvider = ""
                try {
                    $userDomain = $u.userPrincipalName.Split('@')[1]
                    if ($federatedDomains -contains $userDomain) {
                        $isFederated = $true
                        $federatedProvider = $federatedDomainsDetails[$userDomain]
                    }
                } catch {}

                # Check for registered MFA methods
                $directMfa = $false
                $mfaEnforced = $false
                $weakMfaOnly = $false
                $hasStrongMfa = $false
                $methodTypes = @()

                try {
                    $methods = Invoke-MgGraphRequest -Method GET -Uri ("https://graph.microsoft.com/v1.0/users/{0}/authentication/methods" -f $u.Id) -ErrorAction SilentlyContinue
                    if ($methods.value) {
                        foreach ($m in $methods.value) {
                            $otype = $m.'@odata.type'

                            # Categorize method strength
                            $isStrongMethod = $false
                            $isWeakMethod = $false

                            switch ($otype) {
                                '#microsoft.graph.fido2AuthenticationMethod' {
                                    $directMfa = $true
                                    $isStrongMethod = $true
                                    $methodTypes += "FIDO2"
                                }
                                '#microsoft.graph.windowsHelloForBusinessAuthenticationMethod' {
                                    $directMfa = $true
                                    $isStrongMethod = $true
                                    $methodTypes += "WindowsHello"
                                }
                                '#microsoft.graph.microsoftAuthenticatorAuthenticationMethod' {
                                    $directMfa = $true
                                    $isStrongMethod = $true
                                    $methodTypes += "Authenticator"
                                }
                                '#microsoft.graph.softwareOathAuthenticationMethod' {
                                    $directMfa = $true
                                    $isStrongMethod = $true
                                    $methodTypes += "OATH"
                                }
                                '#microsoft.graph.phoneAuthenticationMethod' {
                                    $directMfa = $true
                                    $isWeakMethod = $true
                                    $methodTypes += "Phone"
                                }
                                '#microsoft.graph.emailAuthenticationMethod' {
                                    $directMfa = $true
                                    $isWeakMethod = $true
                                    $methodTypes += "Email"
                                }
                                '#microsoft.graph.temporaryAccessPassAuthenticationMethod' {
                                    $directMfa = $true
                                    $methodTypes += "TAP"
                                }
                            }

                            if ($isStrongMethod) { $hasStrongMfa = $true }
                        }

                        # Determine if only weak methods are registered
                        if ($directMfa -and -not $hasStrongMfa) { $weakMfaOnly = $true }
                    }
                } catch {}

                # Check per-user MFA enforcement state
                try {
                    $authReq = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$($u.Id)?`$select=strongAuthenticationRequirements" -ErrorAction SilentlyContinue
                    if ($authReq.strongAuthenticationRequirements -and $authReq.strongAuthenticationRequirements.Count -gt 0) {
                        $state = $authReq.strongAuthenticationRequirements[0].state
                        if ($state -eq "Enforced" -or $state -eq "Enabled") {
                            $mfaEnforced = $true
                        }
                    }
                } catch {}

                # Get user group and role memberships for CA evaluation (including transitive/nested groups)
                $userGroups = @()
                $userRoles = @()
                $isPrivileged = $false

                try {
                    # Use transitive membership to get all groups including nested ones
                    $transitiveMembers = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$($u.Id)/transitiveMemberOf?`$select=id" -ErrorAction SilentlyContinue
                    if ($transitiveMembers -and $transitiveMembers.value) {
                        foreach ($m in $transitiveMembers.value) {
                            if ($m.'@odata.type' -eq '#microsoft.graph.group') {
                                $userGroups += $m.id
                            } elseif ($m.'@odata.type' -eq '#microsoft.graph.directoryRole') {
                                $userRoles += $m.id
                                if ($privilegedRoleIds -contains $m.id) {
                                    $isPrivileged = $true
                                }
                            }
                        }
                    }

                    # If transitive query failed or returned nothing, fall back to direct membership
                    if ($userGroups.Count -eq 0 -and $userRoles.Count -eq 0) {
                        $mem = Get-MgUserMemberOf -UserId $u.Id -All -ErrorAction SilentlyContinue
                        foreach ($m in $mem) {
                            if ($m.'@odata.type' -eq '#microsoft.graph.group') {
                                $userGroups += $m.Id
                            }
                            elseif ($m.'@odata.type' -eq '#microsoft.graph.directoryRole') {
                                $userRoles += $m.Id
                                if ($privilegedRoleIds -contains $m.Id) {
                                    $isPrivileged = $true
                                }
                            }
                        }
                    }
                } catch {}

                # Evaluate CA policy applicability
                $userCaRequiresMfa = $false
                $userCaIsConditional = $false
                $applicablePolicyCount = 0

                foreach ($p in $mfaPolicies) {
                    $conds = $p.conditions
                    if (-not $conds) { continue }
                    $usersCond = $conds.users
                    $included = $false
                    $excluded = $false

                    if ($usersCond) {
                        # Check includes
                        if ($usersCond.includeUsers -contains 'All') {
                            $included = $true
                        } elseif ($usersCond.includeUsers -contains $u.Id) {
                            $included = $true
                        } elseif ($usersCond.includeGroups) {
                            foreach ($groupId in $usersCond.includeGroups) {
                                if ($userGroups -contains $groupId) {
                                    $included = $true
                                    break
                                }
                            }
                        } elseif ($usersCond.includeRoles) {
                            foreach ($roleId in $usersCond.includeRoles) {
                                if ($userRoles -contains $roleId) {
                                    $included = $true
                                    break
                                }
                            }
                        }

                        # Check exclusions
                        if ($usersCond.excludeUsers -contains $u.Id) {
                            $excluded = $true
                        } elseif ($usersCond.excludeGroups) {
                            foreach ($groupId in $usersCond.excludeGroups) {
                                if ($userGroups -contains $groupId) {
                                    $excluded = $true
                                    break
                                }
                            }
                        } elseif ($usersCond.excludeRoles) {
                            foreach ($roleId in $usersCond.excludeRoles) {
                                if ($userRoles -contains $roleId) {
                                    $excluded = $true
                                    break
                                }
                            }
                        }
                    }

                    # If policy applies, check if it's conditional
                    if ($included -and -not $excluded) {
                        $userCaRequiresMfa = $true
                        $applicablePolicyCount++

                        # Check if policy has conditions that make it conditional
                        if ($conds.applications -and $conds.applications.includeApplications -notcontains "All") {
                            $userCaIsConditional = $true
                        }
                        if ($conds.platforms -and $conds.platforms.includePlatforms -notcontains "all") {
                            $userCaIsConditional = $true
                        }
                        if ($conds.locations) {
                            $userCaIsConditional = $true
                        }
                        if ($conds.signInRiskLevels -or $conds.userRiskLevels) {
                            $userCaIsConditional = $true
                        }

                        # If any policy is unconditional, user has full MFA coverage
                        if (-not $userCaIsConditional) {
                            break
                        }
                    }
                }

                # Determine overall coverage status
                $covered = $false
                $coverageType = "None"
                $warnings = @()

                if ($isFederated) {
                    # Federated users - MFA handled by external IdP
                    $covered = $true
                    $coverageType = "ThirdParty-Federated"
                    $warnings += "Federated to $federatedProvider - MFA depends on IdP config"
                } elseif ($secDefaultsEnabled) {
                    $covered = $true
                    $coverageType = "SecurityDefaults"
                } elseif ($userCaRequiresMfa -and -not $userCaIsConditional) {
                    $covered = $true
                    $coverageType = "ConditionalAccess-Full"
                } elseif ($userCaRequiresMfa -and $userCaIsConditional) {
                    $covered = $true
                    $coverageType = "ConditionalAccess-Partial"
                    $warnings += "Conditional MFA only"
                } elseif ($mfaEnforced) {
                    $covered = $true
                    $coverageType = "PerUserMFA-Enforced"
                } elseif ($directMfa) {
                    $covered = $false
                    $coverageType = "MethodsOnly-NoEnforcement"
                    $warnings += "Methods registered but not enforced"
                }

                # Add warnings
                if ($weakMfaOnly) {
                    $warnings += "Weak MFA methods only"
                }
                if ($isPrivileged -and -not $hasStrongMfa) {
                    $warnings += "PRIVILEGED USER WITHOUT STRONG MFA"
                }
                if ($isPrivileged -and -not $covered) {
                    $warnings += "CRITICAL: PRIVILEGED USER NOT COVERED"
                }

                $users += [pscustomobject]@{
                    DisplayName           = $u.displayName
                    UserPrincipalName     = $u.userPrincipalName
                    IsFederated           = $isFederated
                    FederatedProvider     = $federatedProvider
                    PerUserMfaEnabled     = $directMfa
                    PerUserMfaEnforced    = $mfaEnforced
                    MfaMethods            = ($methodTypes -join ',')
                    HasStrongMfa          = $hasStrongMfa
                    WeakMfaOnly           = $weakMfaOnly
                    SecurityDefaults      = $secDefaultsEnabled
                    CARequiresMfa         = $userCaRequiresMfa
                    CAIsConditional       = $userCaIsConditional
                    CAPolicyCount         = $applicablePolicyCount
                    IsPrivileged          = $isPrivileged
                    MfaCovered            = $covered
                    CoverageType          = $coverageType
                    Warnings              = ($warnings -join '; ')
                }
            }
        } catch {
            Write-Warning "Error processing users: $($_.Exception.Message)"
        }

        # Calculate summary statistics
        $totalUsers = ($users | Measure-Object).Count
        $coveredUsers = ($users | Where-Object { $_.MfaCovered }).Count
        $uncoveredUsers = $totalUsers - $coveredUsers
        $privilegedUncovered = ($users | Where-Object { $_.IsPrivileged -and -not $_.MfaCovered }).Count
        $weakMethodsCount = ($users | Where-Object { $_.WeakMfaOnly }).Count
        $federatedUsersCount = ($users | Where-Object { $_.IsFederated }).Count

        Write-Host "`nMFA Coverage Summary:" -ForegroundColor Green
        Write-Host "  Total Users: $totalUsers" -ForegroundColor White
        Write-Host "  Covered: $coveredUsers ($([math]::Round($coveredUsers/$totalUsers*100,1))%)" -ForegroundColor Green
        Write-Host "  Uncovered: $uncoveredUsers ($([math]::Round($uncoveredUsers/$totalUsers*100,1))%)" -ForegroundColor $(if($uncoveredUsers -gt 0){"Red"}else{"Green"})
        if ($federatedUsersCount -gt 0) {
            Write-Host "  Federated Users (Third-Party MFA): $federatedUsersCount" -ForegroundColor Cyan
        }
        if ($privilegedUncovered -gt 0) {
            Write-Host "  ⚠️ Privileged users uncovered: $privilegedUncovered" -ForegroundColor Red
        }
        if ($weakMethodsCount -gt 0) {
            Write-Host "  ⚠️ Users with weak MFA only: $weakMethodsCount" -ForegroundColor Yellow
        }

        $tenantLevelCaMfa = ($mfaPolicies.Count -gt 0)
        return @{
            SecurityDefaultsEnabled = $secDefaultsEnabled
            CAPoliciesRequireMfa = $tenantLevelCaMfa
            TotalCAPolicies = $mfaPolicies.Count
            FederatedDomains = $federatedDomains
            FederatedDomainsDetails = $federatedDomainsDetails
            Users = $users
            Summary = @{
                TotalUsers = $totalUsers
                CoveredUsers = $coveredUsers
                UncoveredUsers = $uncoveredUsers
                FederatedUsers = $federatedUsersCount
                PrivilegedUncovered = $privilegedUncovered
                WeakMethodsOnly = $weakMethodsCount
            }
        }
    } catch {
        Write-Error "Get-MfaCoverageReport failed: $($_.Exception.Message)"
        return @{
            SecurityDefaultsEnabled = $false
            CAPoliciesRequireMfa = $false
            TotalCAPolicies = 0
            FederatedDomains = @()
            FederatedDomainsDetails = @{}
            Users = @()
            Summary = @{}
        }
    }
}

# Flattens user membership in directory roles and security groups
function Get-UserSecurityGroupsReport {
    try {
        $results = New-Object System.Collections.Generic.List[object]

        # Directory roles (e.g., Global Administrator)
        $roles = @()
        try { $roles = Get-MgDirectoryRole -All -ErrorAction SilentlyContinue } catch {}
        $roleIdToName = @{}
        foreach ($r in $roles) { $roleIdToName[$r.Id] = $r.DisplayName }

        # Users
        $users = @()
        try { $users = Get-MgUser -All -Property 'id,displayName,userPrincipalName' -ErrorAction Stop } catch {}

        foreach ($u in $users) {
            $groups = @()
            try {
                # Use transitive membership to get all groups including nested ones
                $transitiveMembers = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$($u.Id)/transitiveMemberOf" -ErrorAction SilentlyContinue
                if ($transitiveMembers -and $transitiveMembers.value) {
                    foreach ($m in $transitiveMembers.value) {
                        $name = $null
                        if ($m.'@odata.type' -eq '#microsoft.graph.group') { $name = $m.displayName }
                        elseif ($m.'@odata.type' -eq '#microsoft.graph.directoryRole') { $name = if ($roleIdToName.ContainsKey($m.id)) { $roleIdToName[$m.id] } else { 'Directory Role' } }
                        if ($name) { $groups += $name }
                    }
                } else {
                    # Fall back to direct membership
                    $mem = Get-MgUserMemberOf -UserId $u.Id -All -ErrorAction SilentlyContinue
                    foreach ($m in $mem) {
                        $name = $null
                        if ($m.'@odata.type' -eq '#microsoft.graph.group') { $name = $m.DisplayName }
                        elseif ($m.'@odata.type' -eq '#microsoft.graph.directoryRole') { $name = if ($roleIdToName.ContainsKey($m.Id)) { $roleIdToName[$m.Id] } else { 'Directory Role' } }
                        if ($name) { $groups += $name }
                    }
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
        [bool]$IncludeConnectors = $true,
        [Parameter(Mandatory=$false)]
        [bool]$IncludeAuditLogs = $true,
        [Parameter(Mandatory=$false)]
        [bool]$IncludeMfaCoverage = $true,
        [Parameter(Mandatory=$false)]
        [bool]$IncludeUserSecurityGroups = $true
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
                if ($StatusLabel -and $StatusLabel.GetType().Name -eq "Label") { $StatusLabel.Text = "Collecting message trace data (last 10 days)..." }
                $report.MessageTrace = Get-ExchangeMessageTrace -DaysBack 10 # always 10 days per requirement
            }

            if ($IncludeInboxRules) {
                if ($StatusLabel -and $StatusLabel.GetType().Name -eq "Label") { $StatusLabel.Text = "Exporting all inbox rules for tenant..." }
                $report.InboxRules = Get-ExchangeInboxRules
            }

            if ($IncludeTransportRules) {
                if ($StatusLabel -and $StatusLabel.GetType().Name -eq "Label") { $StatusLabel.Text = "Collecting transport rules..." }
                $report.TransportRules = Get-ExchangeTransportRules
            }

            if ($IncludeConnectors) {
                if ($StatusLabel -and $StatusLabel.GetType().Name -eq "Label") { $StatusLabel.Text = "Collecting mail flow connectors..." }
                $report.InboundConnectors = Get-ExchangeInboundConnectors
                $report.OutboundConnectors = Get-ExchangeOutboundConnectors
            }
        } catch {
            Write-Warning "Failed to collect Exchange Online data: $($_.Exception.Message)"
            $report.ExchangeDataError = $_.Exception.Message
        }
    }

    # Collect data from Microsoft Graph
    if ($graphConnected) {
        try {
            if ($IncludeAuditLogs) {
                if ($StatusLabel -and $StatusLabel.GetType().Name -eq "Label") { $StatusLabel.Text = "Collecting audit logs from Microsoft Graph..." }
                $report.AuditLogs = Get-GraphAuditLogs -DaysBack $DaysBack
            }
        } catch {
            Write-Warning "Failed to collect Microsoft Graph data: $($_.Exception.Message)"
            $report.GraphDataError = $_.Exception.Message
        }
    }

    # MFA Coverage and User Security Groups
    if ($graphConnected) {
        try {
            if ($IncludeMfaCoverage) {
                if ($StatusLabel -and $StatusLabel.GetType().Name -eq "Label") { $StatusLabel.Text = "Evaluating MFA coverage (Security Defaults / CA / Per-user)..." }
                $report.MfaCoverage = Get-MfaCoverageReport
            }

            if ($IncludeUserSecurityGroups) {
                if ($StatusLabel -and $StatusLabel.GetType().Name -eq "Label") { $StatusLabel.Text = "Collecting user security groups and roles..." }
                $report.UserSecurityGroups = Get-UserSecurityGroupsReport
            }
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

            # Connectors export
            $csv = Join-Path $report.OutputFolder "InboundConnectors.csv"
            $json = Join-Path $report.OutputFolder "InboundConnectors.json"
            if ($report.InboundConnectors -and $report.InboundConnectors.Count -gt 0) {
                try { $report.InboundConnectors | Export-Csv -Path $csv -NoTypeInformation -Encoding UTF8; $report.FilePaths.InboundConnectorsCsv = $csv }
                catch { $report.InboundConnectors | ConvertTo-Json -Depth 8 | Out-File -FilePath $json -Encoding utf8; $report.FilePaths.InboundConnectorsJson = $json }
            }

            $csv = Join-Path $report.OutputFolder "OutboundConnectors.csv"
            $json = Join-Path $report.OutputFolder "OutboundConnectors.json"
            if ($report.OutboundConnectors -and $report.OutboundConnectors.Count -gt 0) {
                try { $report.OutboundConnectors | Export-Csv -Path $csv -NoTypeInformation -Encoding UTF8; $report.FilePaths.OutboundConnectorsCsv = $csv }
                catch { $report.OutboundConnectors | ConvertTo-Json -Depth 8 | Out-File -FilePath $json -Encoding utf8; $report.FilePaths.OutboundConnectorsJson = $json }
            }

            $csv = Join-Path $report.OutputFolder "GraphAuditLogs.csv"
            $json = Join-Path $report.OutputFolder "GraphAuditLogs.json"
            if ($report.AuditLogs -and $report.AuditLogs.Count -gt 0) {
                try { $report.AuditLogs | Export-Csv -Path $csv -NoTypeInformation -Encoding UTF8; $report.FilePaths.AuditLogsCsv = $csv }
                catch { $report.AuditLogs | ConvertTo-Json -Depth 8 | Out-File -FilePath $json -Encoding utf8; $report.FilePaths.AuditLogsJson = $json }
            }

            # MFA Coverage export
            if ($report.MfaCoverage -and $report.MfaCoverage.Users) {
                $csv = Join-Path $report.OutputFolder "MFAStatus.csv"
                try { $report.MfaCoverage.Users | Export-Csv -Path $csv -NoTypeInformation -Encoding UTF8; $report.FilePaths.MFAStatusCsv = $csv } catch {}
            }

            # User Security Groups export
            if ($report.UserSecurityGroups) {
                $csv = Join-Path $report.OutputFolder "UserSecurityGroups.csv"
                try { $report.UserSecurityGroups | Export-Csv -Path $csv -NoTypeInformation -Encoding UTF8; $report.FilePaths.UserSecurityGroupsCsv = $csv } catch {}
            }
        } catch { $exportError = $_ }

        # Save only LLM instructions as TXT (no other text files on disk)
        try {
            $report.LLMInstructions = New-LLMInvestigationInstructions -Report $report
            $llmPath = Join-Path $report.OutputFolder "LLM_Instructions.txt"
            if ($report.LLMInstructions) { $report.LLMInstructions | Out-File -FilePath $llmPath -Encoding utf8 }
            $report.FilePaths.LLMInstructionsTxt = $llmPath
        } catch {}
    }

    return $report
}

function Get-ExchangeMessageTrace {
    param([int]$DaysBack = 10)

    try {
        Write-Host "Collecting message trace data..." -ForegroundColor Yellow
        $end = (Get-Date).ToUniversalTime()
        $start = $end.AddDays(-10).Date # always 10 full days; start at 00:00Z

        $results = New-Object System.Collections.Generic.List[object]

        $hasV2 = $null -ne (Get-Command Get-MessageTraceV2 -ErrorAction SilentlyContinue)

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

        return [System.Collections.ArrayList]$results
    } catch {
        Write-Error "Failed to collect message trace: $($_.Exception.Message)"
        return @()
    }
}

function Get-ExchangeInboxRules {
    try {
        Write-Host "Exporting inbox rules..." -ForegroundColor Yellow

        $mailboxes = @()
        try {
            $mailboxes = Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox,SharedMailbox -ErrorAction Stop
        } catch {
            # Fallback narrower call if needed
            $mailboxes = Get-Mailbox -ResultSize 2000 -ErrorAction Stop
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
