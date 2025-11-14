# EntraInvestigator.psm1 - Clean Rebuild
# Essential Entra ID/Graph functions for modular use

$script:requiredModules = @(
    "Microsoft.Graph.Users",
    "Microsoft.Graph.Reports",
    "Microsoft.Graph.Identity.DirectoryManagement",
    "Microsoft.Graph.Identity.SignIns",
    "Microsoft.Graph.Security"
)
$script:requiredScopes = @(
    "User.Read.All", "AuditLog.Read.All", "Organization.Read.All", "Directory.Read.All", "Policy.Read.All", "UserAuthenticationMethod.Read.All", "SecurityEvents.ReadWrite.All"
)

function Test-EntraModules {
    [CmdletBinding()]
    param()
    $missing = @()
    foreach ($m in $script:requiredModules) {
        if (-not (Get-Module -ListAvailable -Name $m)) { $missing += $m }
    }
    return $missing
}

function Install-EntraModules {
    [CmdletBinding()]
    param([string[]]$Modules)
    foreach ($m in $Modules) {
        try { Install-Module -Name $m -Scope CurrentUser -Repository PSGallery -Force -AllowClobber -ErrorAction Stop }
        catch { Write-Error "Failed to install module: $m. $_" }
    }
}

function Connect-EntraGraph {
    [CmdletBinding()]
    param()
    try {
        Disconnect-MgGraph -ErrorAction SilentlyContinue
        Connect-MgGraph -Scopes $script:requiredScopes -ErrorAction Stop
        return $true
    } catch {
        # Check if this is a user cancellation
        $errorMessage = $_.Exception.Message
        $isUserCancellation = $errorMessage -match "User cancelled|Operation cancelled|User canceled|Authentication cancelled|Authentication canceled" -or 
                             $errorMessage -match "AADSTS50020|AADSTS50076|AADSTS50079" -or
                             $errorMessage -match "The user cancelled the authentication"
        
        if ($isUserCancellation) {
            # User cancelled - return false without writing error
            return $false
        } else {
            # Real error - write error message
            Write-Error "Failed to connect to Microsoft Graph: $_"
            return $false
        }
    }
}

function Show-DebugTextBox {
    param(
        [string]$Text,
        [string]$Title = "Debug Output"
    )
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $Title
    $form.Size = New-Object System.Drawing.Size(700, 600)
    $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Multiline = $true
    $textBox.ScrollBars = 'Both'
    $textBox.Dock = 'Fill'
    $textBox.ReadOnly = $false
    $textBox.Text = $Text
    $textBox.Font = New-Object System.Drawing.Font('Consolas', 10)
    $form.Controls.Add($textBox)
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "OK"
    $okButton.Dock = 'Bottom'
    $okButton.Add_Click({ $form.Close() })
    $form.Controls.Add($okButton)
    $form.Topmost = $true
    [void]$form.ShowDialog()
}

function Get-EntraUsers {
    [CmdletBinding()]
    param(
        [int]$MaxUsers = 5000,
        [switch]$LoadAll
    )
    try {
        if ($LoadAll) {
            $result = Get-MgUser -All -Property UserPrincipalName,DisplayName,Id -ConsistencyLevel eventual | Sort-Object UserPrincipalName
        } else {
            # Load users in batches for better performance
            $result = Get-MgUser -Top $MaxUsers -Property UserPrincipalName,DisplayName,Id -ConsistencyLevel eventual | Sort-Object UserPrincipalName
        }
        return $result
    } catch {
        Write-Error "Failed to fetch users: $_"
        return @()
    }
}

function Get-EntraSignInLogs {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string[]]$UserPrincipalNames,
        [Parameter(Mandatory)] [int]$Days
    )
    $allLogs = @()
    $startDate = (Get-Date).AddDays(-$Days).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
    foreach ($upn in $UserPrincipalNames) {
        try {
            $userId = (Get-MgUser -UserId $upn -Property Id).Id
            $filter = "userId eq '$userId' and createdDateTime ge $startDate"
            $logs = Get-MgAuditLogSignIn -Filter $filter -All -ErrorAction Stop
            if ($logs) {
                $logs | ForEach-Object { $_ | Add-Member -MemberType NoteProperty -Name 'UserPrincipalName' -Value $upn -Force }
                $allLogs += $logs
            }
        } catch { Write-Warning ('Could not get logs for {0}: {1}' -f $upn, $_) }
    }
    return $allLogs
}

function Get-EntraUserDetailsAndRoles {
    [CmdletBinding()]
    param([Parameter(Mandatory)] [string]$UserPrincipalName)
    $result = @{User=$null; Roles=@(); Groups=@(); Error=$null}
    try {
        $user = Get-MgUser -UserId $UserPrincipalName -Property Id,DisplayName,AccountEnabled,LastPasswordChangeDateTime,UserPrincipalName
        $result.User = $user
        $memberOf = Get-MgUserMemberOf -UserId $user.Id -All -ErrorAction SilentlyContinue
        if ($memberOf) {
            $result.Groups = $memberOf | Where-Object { $_.'@odata.type' -eq '#microsoft.graph.group' } | Select-Object -ExpandProperty DisplayName
        }
        # Enumerate all directory roles and check if user is a member
        $allRoles = Get-MgDirectoryRole -All -ErrorAction SilentlyContinue
        $userRoles = @()
        foreach ($role in $allRoles) {
            $roleMembers = Get-MgDirectoryRoleMember -DirectoryRoleId $role.Id -All -ErrorAction SilentlyContinue
            if ($roleMembers | Where-Object { $_.Id -eq $user.Id }) {
                $userRoles += $role.DisplayName
            }
        }
        $result.Roles = $userRoles
    } catch { $result.Error = $_.Exception.Message }
    return $result
}

function Get-EntraUserAuditLogs {
    [CmdletBinding()]
    param([Parameter(Mandatory)] [string]$UserPrincipalName, [Parameter(Mandatory)] [int]$Days)
    try {
        $userId = (Get-MgUser -UserId $UserPrincipalName -Property Id).Id
        $startDate = (Get-Date).AddDays(-$Days).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
        $filter = "(initiatedBy/user/id eq '$userId') and (activityDateTime ge $startDate)"
        $logs = Get-MgAuditLogDirectoryAudit -Filter $filter -All -ErrorAction Stop
        return $logs
    } catch {
        Write-Error "Failed to fetch audit logs: $_"
        return @()
    }
}

function Get-EntraUserMfaStatus {
    [CmdletBinding()]
    param([Parameter(Mandatory)] [string]$UserPrincipalName)
    $results = @{
        PerUserMfa = @{ Enabled = $false; Enforced = $false; Methods = @(); MethodDetails = @(); Details = "Not configured"; Warnings = @() }
        SecurityDefaults = @{ Enabled = $false; Details = "Unknown" }
        ConditionalAccess = @{ Policies = @(); RequiresMfa = $false; ConditionalMfa = $false; Details = "No applicable policies" }
        ThirdPartyMfa = @{ Detected = $false; Type = "None"; Details = "No third-party MFA detected" }
        OverallStatus = "Unknown"
        Summary = ""
        Warnings = @()
    }
    try {
        # Get user with expanded properties
        $user = Get-MgUser -UserId $UserPrincipalName -Property Id,UserPrincipalName,DisplayName -ErrorAction Stop

        # Check for federated/third-party authentication
        try {
            # Extract domain from UPN
            $userDomain = $UserPrincipalName.Split('@')[1]
            if ($userDomain) {
                $domain = Get-MgDomain -DomainId $userDomain -ErrorAction SilentlyContinue
                if ($domain) {
                    # Check authentication type
                    if ($domain.AuthenticationType -eq "Federated") {
                        $results.ThirdPartyMfa.Detected = $true
                        $results.ThirdPartyMfa.Type = "Federated Identity Provider"

                        # Try to get federation settings for more details
                        try {
                            $fedSettings = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/domains/$userDomain/federationConfiguration" -ErrorAction SilentlyContinue
                            if ($fedSettings.value -and $fedSettings.value.Count -gt 0) {
                                $issuerUri = $fedSettings.value[0].issuerUri
                                $fedDetails = "Domain is federated"

                                # Detect common providers by issuer URI
                                if ($issuerUri -match "duosecurity\.com") {
                                    $fedDetails = "Federated to Duo Security"
                                } elseif ($issuerUri -match "okta\.com") {
                                    $fedDetails = "Federated to Okta"
                                } elseif ($issuerUri -match "pingidentity\.com|pingone\.com") {
                                    $fedDetails = "Federated to Ping Identity"
                                } elseif ($issuerUri -match "auth0\.com") {
                                    $fedDetails = "Federated to Auth0"
                                } elseif ($issuerUri -match "onelogin\.com") {
                                    $fedDetails = "Federated to OneLogin"
                                } elseif ($issuerUri -match "adfs") {
                                    $fedDetails = "Federated to on-premises AD FS"
                                } else {
                                    $fedDetails = "Federated to external IdP: $issuerUri"
                                }

                                $results.ThirdPartyMfa.Details = "$fedDetails - MFA likely enforced by IdP (cannot verify from Azure AD)"
                            } else {
                                $results.ThirdPartyMfa.Details = "Domain is federated - MFA likely enforced by external identity provider"
                            }
                        } catch {
                            $results.ThirdPartyMfa.Details = "Domain is federated - MFA likely enforced by external identity provider"
                        }
                    }
                }
            }
        } catch {
            # Domain check failed, continue with other checks
        }

        # Get user's group and role memberships for CA policy evaluation
        $userGroups = @()
        $userRoles = @()
        try {
            $memberOf = Get-MgUserMemberOf -UserId $user.Id -All -ErrorAction SilentlyContinue
            if ($memberOf) {
                $userGroups = $memberOf | Where-Object { $_.'@odata.type' -eq '#microsoft.graph.group' } | Select-Object -ExpandProperty Id
                $userRoles = $memberOf | Where-Object { $_.'@odata.type' -eq '#microsoft.graph.directoryRole' } | Select-Object -ExpandProperty Id
            }
        } catch {
            $results.Warnings += "Could not retrieve group/role memberships - CA policy evaluation may be incomplete"
        }

        # Check authentication methods with quality assessment
        $authMethods = Get-MgUserAuthenticationMethod -UserId $user.Id -ErrorAction SilentlyContinue
        if ($authMethods) {
            $mfaMethods = $authMethods | Where-Object { $_.'@odata.type' -ne '#microsoft.graph.passwordAuthenticationMethod' }
            if ($mfaMethods) {
                $results.PerUserMfa.Enabled = $true
                $methodList = @()
                $weakMethods = @()

                foreach ($method in $mfaMethods) {
                    $methodType = $method.'@odata.type' -replace '#microsoft.graph.', '' -replace 'AuthenticationMethod', ''
                    $methodList += $methodType

                    # Detailed method info with security assessment
                    $methodDetail = @{ Type = $methodType; Strong = $true; Details = "" }

                    switch ($method.'@odata.type') {
                        '#microsoft.graph.phoneAuthenticationMethod' {
                            $methodDetail.Strong = $false
                            $methodDetail.Details = "Phone: $($method.PhoneNumber) - ⚠️ Vulnerable to SIM swapping"
                            $weakMethods += "SMS/Voice (SIM swapping risk)"
                        }
                        '#microsoft.graph.emailAuthenticationMethod' {
                            $methodDetail.Strong = $false
                            $methodDetail.Details = "Email: $($method.EmailAddress) - ⚠️ Not recommended for MFA"
                            $weakMethods += "Email (not secure)"
                        }
                        '#microsoft.graph.fido2AuthenticationMethod' {
                            $methodDetail.Details = "FIDO2: $($method.Model) - ✓ Phishing-resistant"
                        }
                        '#microsoft.graph.microsoftAuthenticatorAuthenticationMethod' {
                            $methodDetail.Details = "Microsoft Authenticator - ✓ Strong"
                        }
                        '#microsoft.graph.softwareOathAuthenticationMethod' {
                            $methodDetail.Details = "Software OATH Token - ✓ Strong"
                        }
                        '#microsoft.graph.windowsHelloForBusinessAuthenticationMethod' {
                            $methodDetail.Details = "Windows Hello for Business - ✓ Phishing-resistant"
                        }
                        '#microsoft.graph.temporaryAccessPassAuthenticationMethod' {
                            $tapExpiry = $method.LifetimeInMinutes
                            $tapCreated = $method.CreatedDateTime
                            if ($tapCreated) {
                                $expiryDate = ([DateTime]$tapCreated).AddMinutes($tapExpiry)
                                if ($expiryDate -lt (Get-Date)) {
                                    $methodDetail.Details = "TAP: EXPIRED on $expiryDate"
                                    $results.Warnings += "Temporary Access Pass is expired"
                                } else {
                                    $methodDetail.Details = "TAP: Expires $expiryDate"
                                }
                            }
                        }
                    }

                    $results.PerUserMfa.MethodDetails += $methodDetail
                }

                $results.PerUserMfa.Methods = $methodList
                $results.PerUserMfa.Details = "Methods: $($methodList -join ', ')"

                if ($weakMethods.Count -gt 0) {
                    $results.PerUserMfa.Warnings = $weakMethods
                    $results.Warnings += "Weak MFA methods detected: $($weakMethods -join ', ')"
                }
            } else {
                $results.PerUserMfa.Details = "No MFA methods registered"
            }
        }

        # Check per-user MFA enforcement state (legacy)
        try {
            # Try to get strong authentication requirements
            $authReq = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$($user.Id)?`$select=strongAuthenticationRequirements" -ErrorAction SilentlyContinue
            if ($authReq.strongAuthenticationRequirements -and $authReq.strongAuthenticationRequirements.Count -gt 0) {
                $state = $authReq.strongAuthenticationRequirements[0].state
                if ($state -eq "Enforced" -or $state -eq "Enabled") {
                    $results.PerUserMfa.Enforced = $true
                    $results.PerUserMfa.Details += " | Per-User MFA State: $state"
                }
            }
        } catch {
            # Property not available in all tenants/licenses
        }

        # Check Security Defaults
        $securityDefaults = Get-MgPolicyIdentitySecurityDefaultEnforcementPolicy -ErrorAction SilentlyContinue
        if ($securityDefaults) {
            $results.SecurityDefaults.Enabled = $securityDefaults.IsEnabled
            $results.SecurityDefaults.Details = if ($securityDefaults.IsEnabled) { "Enabled (requires MFA for all users)" } else { "Disabled" }
        }

        # Evaluate Conditional Access policies with complete logic
        $caPolicies = Get-MgIdentityConditionalAccessPolicy -All -ErrorAction SilentlyContinue
        if ($caPolicies) {
            $applicablePolicies = @()

            foreach ($policy in $caPolicies) {
                if ($policy.State -ne "enabled") { continue }

                # Evaluate user assignment (include/exclude logic)
                $included = $false
                $excluded = $false

                # Check includes
                if ($policy.Conditions.Users.IncludeUsers -contains "All") {
                    $included = $true
                } elseif ($policy.Conditions.Users.IncludeUsers -contains $user.Id) {
                    $included = $true
                } elseif ($policy.Conditions.Users.IncludeGroups) {
                    foreach ($groupId in $policy.Conditions.Users.IncludeGroups) {
                        if ($userGroups -contains $groupId) {
                            $included = $true
                            break
                        }
                    }
                } elseif ($policy.Conditions.Users.IncludeRoles) {
                    foreach ($roleId in $policy.Conditions.Users.IncludeRoles) {
                        if ($userRoles -contains $roleId) {
                            $included = $true
                            break
                        }
                    }
                }

                # Check exclusions (take precedence)
                if ($policy.Conditions.Users.ExcludeUsers -contains $user.Id) {
                    $excluded = $true
                } elseif ($policy.Conditions.Users.ExcludeGroups) {
                    foreach ($groupId in $policy.Conditions.Users.ExcludeGroups) {
                        if ($userGroups -contains $groupId) {
                            $excluded = $true
                            break
                        }
                    }
                } elseif ($policy.Conditions.Users.ExcludeRoles) {
                    foreach ($roleId in $policy.Conditions.Users.ExcludeRoles) {
                        if ($userRoles -contains $roleId) {
                            $excluded = $true
                            break
                        }
                    }
                }

                if ($included -and -not $excluded) {
                    # Check grant controls for MFA requirement
                    $requiresMfa = $false
                    $isConditional = $false
                    $controlDetails = ""

                    if ($policy.GrantControls) {
                        # Check built-in controls
                        if ($policy.GrantControls.BuiltInControls -contains "mfa") {
                            $requiresMfa = $true
                            $controlDetails = "Requires MFA"
                        }

                        # Check authentication strength (newer, more granular)
                        if ($policy.GrantControls.AuthenticationStrength) {
                            $requiresMfa = $true
                            $strengthId = $policy.GrantControls.AuthenticationStrength.Id
                            try {
                                $strength = Get-MgPolicyAuthenticationStrengthPolicy -AuthenticationStrengthPolicyId $strengthId -ErrorAction SilentlyContinue
                                if ($strength) {
                                    $controlDetails = "Requires Authentication Strength: $($strength.DisplayName)"
                                    $allowedMethods = $strength.AllowedCombinations -join ', '
                                    $controlDetails += " (Methods: $allowedMethods)"
                                }
                            } catch {
                                $controlDetails = "Requires Authentication Strength (ID: $strengthId)"
                            }
                        }

                        # Check for custom controls (third-party MFA like Duo)
                        if ($policy.GrantControls.CustomAuthenticationFactors) {
                            foreach ($customControl in $policy.GrantControls.CustomAuthenticationFactors) {
                                if (-not $results.ThirdPartyMfa.Detected) {
                                    $results.ThirdPartyMfa.Detected = $true
                                    $results.ThirdPartyMfa.Type = "Custom Authentication Control"
                                    $results.ThirdPartyMfa.Details = "CA policy uses custom authentication control (likely third-party MFA like Duo)"
                                }
                                $controlDetails += " [Includes custom auth control]"
                            }
                        }

                        # Check operator (all vs any)
                        if ($policy.GrantControls.Operator -eq "OR") {
                            $controlDetails += " [One of multiple controls required]"
                        }
                    }

                    # Check session controls (can indicate third-party integrations)
                    if ($policy.SessionControls) {
                        if ($policy.SessionControls.ApplicationEnforcedRestrictions -or
                            $policy.SessionControls.CloudAppSecurity -or
                            $policy.SessionControls.PersistentBrowser) {
                            # Some session controls with MFA can indicate third-party solutions
                            if ($requiresMfa -and $policy.SessionControls.CloudAppSecurity) {
                                if (-not $results.ThirdPartyMfa.Detected) {
                                    $results.ThirdPartyMfa.Detected = $true
                                    $results.ThirdPartyMfa.Type = "Conditional Access App Control"
                                    $results.ThirdPartyMfa.Details = "CA policy uses Cloud App Security/Defender for Cloud Apps controls (may include third-party MFA)"
                                }
                            }
                        }
                    }

                    # Evaluate conditions that make MFA conditional
                    $conditionSummary = @()

                    # Application conditions
                    if ($policy.Conditions.Applications) {
                        $apps = $policy.Conditions.Applications
                        if ($apps.IncludeApplications -notcontains "All") {
                            $isConditional = $true
                            $appCount = ($apps.IncludeApplications | Measure-Object).Count
                            $conditionSummary += "Specific apps only ($appCount apps)"
                        }
                    }

                    # Platform conditions
                    if ($policy.Conditions.Platforms -and $policy.Conditions.Platforms.IncludePlatforms) {
                        if ($policy.Conditions.Platforms.IncludePlatforms -notcontains "all") {
                            $isConditional = $true
                            $platforms = $policy.Conditions.Platforms.IncludePlatforms -join ', '
                            $conditionSummary += "Platforms: $platforms"
                        }
                    }

                    # Location conditions
                    if ($policy.Conditions.Locations -and ($policy.Conditions.Locations.IncludeLocations -or $policy.Conditions.Locations.ExcludeLocations)) {
                        $isConditional = $true
                        if ($policy.Conditions.Locations.ExcludeLocations -contains "AllTrusted") {
                            $conditionSummary += "Excludes trusted locations"
                        } else {
                            $conditionSummary += "Location-based"
                        }
                    }

                    # Risk conditions
                    if ($policy.Conditions.SignInRiskLevels -or $policy.Conditions.UserRiskLevels) {
                        $isConditional = $true
                        $conditionSummary += "Risk-based"
                    }

                    # Device conditions
                    if ($policy.Conditions.Devices) {
                        $isConditional = $true
                        $conditionSummary += "Device-based"
                    }

                    $policyInfo = @{
                        Name = $policy.DisplayName
                        State = $policy.State
                        RequiresMfa = $requiresMfa
                        IsConditional = $isConditional
                        Controls = $controlDetails
                        Conditions = if ($conditionSummary.Count -gt 0) { $conditionSummary -join "; " } else { "Applies to all scenarios" }
                    }

                    $applicablePolicies += $policyInfo

                    if ($requiresMfa) {
                        $results.ConditionalAccess.RequiresMfa = $true
                        if ($isConditional) {
                            $results.ConditionalAccess.ConditionalMfa = $true
                        }
                    }
                }
            }

            $results.ConditionalAccess.Policies = $applicablePolicies
            if ($applicablePolicies.Count -gt 0) {
                $mfaPoliciesCount = ($applicablePolicies | Where-Object { $_.RequiresMfa }).Count
                $conditionalCount = ($applicablePolicies | Where-Object { $_.IsConditional }).Count
                $results.ConditionalAccess.Details = "Found $($applicablePolicies.Count) applicable policies ($mfaPoliciesCount require MFA, $conditionalCount are conditional)"
            }
        }

        # Determine overall status with improved categorization
        if ($results.ThirdPartyMfa.Detected) {
            # Third-party/federated MFA takes precedence in status
            $results.OverallStatus = "✓ PROTECTED (Third-Party MFA)"
            $results.Summary = "$($results.ThirdPartyMfa.Details). Note: MFA enforcement cannot be verified from Azure AD when using external identity providers."
            $results.Warnings += "Third-party MFA detected - actual enforcement depends on external provider configuration"
        } elseif ($results.SecurityDefaults.Enabled) {
            $results.OverallStatus = "✓ PROTECTED (Security Defaults)"
            $results.Summary = "MFA required via Security Defaults for all users in all scenarios."
        } elseif ($results.ConditionalAccess.RequiresMfa -and -not $results.ConditionalAccess.ConditionalMfa) {
            $results.OverallStatus = "✓ PROTECTED (Conditional Access)"
            $results.Summary = "MFA required via Conditional Access policy for all scenarios."
        } elseif ($results.ConditionalAccess.RequiresMfa -and $results.ConditionalAccess.ConditionalMfa) {
            $results.OverallStatus = "⚠️ CONDITIONALLY PROTECTED (Conditional Access)"
            $results.Summary = "MFA required by CA policy, but only for specific apps, platforms, or locations. May not cover all sign-in scenarios."
            $results.Warnings += "MFA is conditional - user may be able to sign in without MFA in some scenarios"
        } elseif ($results.PerUserMfa.Enforced) {
            $results.OverallStatus = "✓ PROTECTED (Per-User MFA)"
            $results.Summary = "Per-user MFA is enforced for this account."
        } elseif ($results.PerUserMfa.Enabled) {
            $results.OverallStatus = "⚠️ METHODS REGISTERED (NO ENFORCEMENT)"
            $results.Summary = "MFA methods are registered, but no policy enforces their use. User can still sign in without MFA."
            $results.Warnings += "MFA methods exist but no enforcement policy detected - user can bypass MFA"
        } else {
            $results.OverallStatus = "❌ NOT PROTECTED"
            $results.Summary = "No MFA methods registered and no enforcement policy detected."
        }

        # Additional warnings for privileged users
        if ($userRoles.Count -gt 0) {
            try {
                $allRoles = Get-MgDirectoryRole -All -ErrorAction SilentlyContinue
                $privilegedRoles = @("Global Administrator", "Privileged Role Administrator", "Security Administrator", "Exchange Administrator")
                foreach ($role in $allRoles) {
                    if ($userRoles -contains $role.Id -and $privilegedRoles -contains $role.DisplayName) {
                        $hasPhishingResistant = $results.PerUserMfa.MethodDetails | Where-Object {
                            $_.Type -eq "fido2" -or $_.Type -eq "windowsHelloForBusiness"
                        }
                        if (-not $hasPhishingResistant) {
                            $results.Warnings += "⚠️ CRITICAL: Privileged admin role ($($role.DisplayName)) without phishing-resistant MFA"
                        }
                        break
                    }
                }
            } catch {}
        }

    } catch {
        $results.OverallStatus = "Error"
        $results.Summary = "Failed to analyze MFA status: $_"
    }
    return $results
}

Export-ModuleMember -Function Test-EntraModules,Install-EntraModules,Connect-EntraGraph,Get-EntraUsers,Get-EntraSignInLogs,Get-EntraUserDetailsAndRoles,Get-EntraUserAuditLogs,Get-EntraUserMfaStatus 