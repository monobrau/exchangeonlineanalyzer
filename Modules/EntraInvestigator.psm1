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
    
    # Check if ExchangeOnlineManagement module is loaded (causes version conflicts)
    $exoModuleLoaded = $false
    try {
        $exoModule = Get-Module -Name ExchangeOnlineManagement -ErrorAction SilentlyContinue
        if ($exoModule) {
            $exoModuleLoaded = $true
            Write-Warning "ExchangeOnlineManagement module is already loaded. This may cause authentication conflicts."
            Write-Warning "Attempting to unload ExchangeOnlineManagement module..."
            try {
                Remove-Module -Name ExchangeOnlineManagement -Force -ErrorAction SilentlyContinue
                Write-Host "Unloaded ExchangeOnlineManagement module." -ForegroundColor Yellow
                Start-Sleep -Milliseconds 500  # Give time for assemblies to release
            } catch {
                Write-Warning "Could not unload ExchangeOnlineManagement module: $_"
                Write-Warning "You may need to restart PowerShell and connect to Entra FIRST before Exchange Online."
            }
        }
    } catch {}
    
    try {
        Disconnect-MgGraph -ErrorAction SilentlyContinue
        
        # If Exchange Online was connected, disconnect it
        try {
            if (Get-Command Disconnect-ExchangeOnline -ErrorAction SilentlyContinue) {
                Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
                Write-Host "Disconnected from Exchange Online to prevent conflicts." -ForegroundColor Yellow
            }
        } catch {}
        
        Connect-MgGraph -Scopes $script:requiredScopes -ErrorAction Stop
        
        # Import required Microsoft Graph modules after connection
        Write-Host "Importing required Microsoft Graph modules..." -ForegroundColor Cyan
        
        # Define which modules are core vs optional
        $coreModules = @('Microsoft.Graph.Users', 'Microsoft.Graph.Reports', 'Microsoft.Graph.Identity.SignIns')
        $optionalModules = @('Microsoft.Graph.Identity.DirectoryManagement', 'Microsoft.Graph.Security')
        
        $missingOptional = @()
        foreach ($moduleName in $script:requiredModules) {
            try {
                if (Get-Module -ListAvailable -Name $moduleName -ErrorAction SilentlyContinue) {
                    Import-Module -Name $moduleName -Force -ErrorAction SilentlyContinue
                    Write-Host "  Imported: $moduleName" -ForegroundColor Gray
                } else {
                    if ($optionalModules -contains $moduleName) {
                        $missingOptional += $moduleName
                        # Don't show anything for optional modules - they're truly optional
                    } else {
                        Write-Warning "Core module $moduleName not available. Some features may not work."
                    }
                }
            } catch {
                if ($optionalModules -contains $moduleName) {
                    $missingOptional += $moduleName
                    # Don't show anything for optional modules
                } else {
                    Write-Warning "Could not import core module $moduleName : $_"
                }
            }
        }
        
        # Only show a brief note if optional modules are missing (and only once, not every time)
        if ($missingOptional.Count -gt 0 -and -not $script:optionalModulesNoted) {
            Write-Host "  Note: Optional modules not installed (license info, security features). Core features work fine." -ForegroundColor DarkGray
            $script:optionalModulesNoted = $true
        }
        
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
            # Check for Microsoft.Identity.Client version conflict
            if ($errorMessage -match "Method not found.*WithLogging|BaseAbstractApplicationBuilder.*WithLogging|Microsoft.Identity.Client|InteractiveBrowserCredential") {
                $errorMsg = @"
CRITICAL: Microsoft.Identity.Client Version Conflict

ExchangeOnlineManagement module has loaded an incompatible version of Microsoft.Identity.Client that conflicts with Microsoft Graph modules.

SOLUTION:
1. Close this PowerShell window completely
2. Open a NEW PowerShell window
3. Connect to Entra/Graph FIRST (before Exchange Online)
4. Then connect to Exchange Online if needed

Original error: $($_.Exception.Message)
"@
                Write-Error $errorMsg
                
                # Show MessageBox if running in GUI context
                try {
                    if ([System.Windows.Forms.MessageBox] -as [type]) {
                        [System.Windows.Forms.MessageBox]::Show(
                            $errorMsg,
                            "Microsoft Graph Connection Failed",
                            [System.Windows.Forms.MessageBoxButtons]::OK,
                            [System.Windows.Forms.MessageBoxIcon]::Error
                        ) | Out-Null
                    }
                } catch {}
                
                return $false
            } else {
                # Real error - write error message
                Write-Error "Failed to connect to Microsoft Graph: $_"
                return $false
            }
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
    $results = @{ PerUserMfa = @{ Enabled = $false; Methods = @(); Details = "Not configured" }; SecurityDefaults = @{ Enabled = $false; Details = "Unknown" }; ConditionalAccess = @{ Policies = @(); RequiresMfa = $false; Details = "No applicable policies" }; OverallStatus = "Unknown"; Summary = "" }
    try {
        $user = Get-MgUser -UserId $UserPrincipalName -Property Id
        $authMethods = Get-MgUserAuthenticationMethod -UserId $user.Id -ErrorAction SilentlyContinue
        if ($authMethods) {
            # Handle both single object and collection
            $methodsList = @()
            if ($authMethods -is [Array]) {
                $methodsList = $authMethods
            } elseif ($authMethods.GetType().Name -eq 'PSObject' -or $authMethods.GetType().Name -eq 'Object') {
                # Check if it has a .value property (paged result)
                if ($authMethods.PSObject.Properties.Name -contains 'value') {
                    $methodsList = $authMethods.value
                } else {
                    $methodsList = @($authMethods)
                }
            } else {
                $methodsList = @($authMethods)
            }
            
            $mfaMethods = $methodsList | Where-Object { 
                if ($null -eq $_) { return $false }
                $methodType = $null
                # Try to get @odata.type from AdditionalProperties
                if ($_.AdditionalProperties) {
                    if ($_.AdditionalProperties.ContainsKey('@odata.type')) {
                        $methodType = $_.AdditionalProperties['@odata.type']
                    } elseif ($_.AdditionalProperties['@odata.type']) {
                        $methodType = $_.AdditionalProperties['@odata.type']
                    }
                }
                # Exclude password and email methods (email is not MFA)
                if ($methodType) {
                    return $methodType -ne '#microsoft.graph.passwordAuthenticationMethod' -and 
                           $methodType -ne '#microsoft.graph.emailAuthenticationMethod'
                }
                # If we can't determine type, include it (might be an MFA method)
                return $true
            }
            
            if ($mfaMethods) {
                # If we found any non-password authentication methods, MFA is enabled
                $results.PerUserMfa.Enabled = $true
                
                $methodNames = $mfaMethods | ForEach-Object { 
                    $method = $_
                    $methodType = $null
                    # Get @odata.type from AdditionalProperties
                    if ($method.AdditionalProperties) {
                        if ($method.AdditionalProperties.ContainsKey('@odata.type')) {
                            $methodType = $method.AdditionalProperties['@odata.type']
                        } elseif ($method.AdditionalProperties['@odata.type']) {
                            $methodType = $method.AdditionalProperties['@odata.type']
                        }
                    }
                    
                    if (-not $methodType) { 
                        # If we can't get the type, try to infer from available properties
                        $phoneNumber = $null
                        $phoneType = $null
                        $deviceTag = $null
                        if ($method.AdditionalProperties) {
                            if ($method.AdditionalProperties.ContainsKey('phoneNumber')) { $phoneNumber = $method.AdditionalProperties['phoneNumber'] }
                            if ($method.AdditionalProperties.ContainsKey('phoneType')) { $phoneType = $method.AdditionalProperties['phoneType'] }
                            if ($method.AdditionalProperties.ContainsKey('deviceTag')) { $deviceTag = $method.AdditionalProperties['deviceTag'] }
                        }
                        
                        if ($phoneNumber -or $phoneType) {
                            $methodType = '#microsoft.graph.phoneAuthenticationMethod'
                        } elseif ($deviceTag) {
                            $methodType = '#microsoft.graph.microsoftAuthenticatorAuthenticationMethod'
                        } else {
                            # Can't determine type, but method exists - return generic indicator
                            return 'MFA Method (Type Unknown)'
                        }
                    }
                    
                    # Helper to get property from AdditionalProperties
                    $getProp = {
                        param($propName)
                        if ($method.AdditionalProperties) {
                            if ($method.AdditionalProperties.ContainsKey($propName)) {
                                return $method.AdditionalProperties[$propName]
                            }
                        }
                        return $null
                    }
                    
                    # Map to readable names
                    switch ($methodType) {
                        '#microsoft.graph.microsoftAuthenticatorAuthenticationMethod' { 
                            $displayName = & $getProp 'displayName'
                            $deviceTag = & $getProp 'deviceTag'
                            $parts = @('Microsoft Authenticator')
                            if ($displayName) { $parts += "($displayName)" }
                            if ($deviceTag) { $parts += "[$deviceTag]" }
                            $parts -join ' '
                        }
                        '#microsoft.graph.phoneAuthenticationMethod' { 
                            $phoneNumber = & $getProp 'phoneNumber'
                            $phoneType = & $getProp 'phoneType'
                            $parts = @('Phone')
                            if ($phoneType) {
                                if ($phoneType -eq 'mobile') { $parts += '(Mobile)' }
                                elseif ($phoneType -eq 'alternateMobile') { $parts += '(Alternate Mobile)' }
                                else { $parts += "($phoneType)" }
                            }
                            if ($phoneNumber) { $parts += "[$phoneNumber]" }
                            $parts -join ' '
                        }
                        '#microsoft.graph.softwareOathAuthenticationMethod' { 
                            $displayName = & $getProp 'displayName'
                            if ($displayName) { "Software OATH Token ($displayName)" } else { 'Software OATH Token' }
                        }
                        '#microsoft.graph.fido2AuthenticationMethod' { 
                            $displayName = & $getProp 'displayName'
                            if ($displayName) { "FIDO2 Security Key ($displayName)" } else { 'FIDO2 Security Key' }
                        }
                        '#microsoft.graph.temporaryAccessPassAuthenticationMethod' { 
                            $displayName = & $getProp 'displayName'
                            if ($displayName) { "Temporary Access Pass ($displayName)" } else { 'Temporary Access Pass' }
                        }
                        '#microsoft.graph.windowsHelloForBusinessAuthenticationMethod' { 
                            $displayName = & $getProp 'displayName'
                            if ($displayName) { "Windows Hello ($displayName)" } else { 'Windows Hello' }
                        }
                        default { 
                            # Fallback: extract readable name from type
                            $methodName = $methodType -replace '#microsoft.graph.', '' -replace 'AuthenticationMethod', ''
                            if ($methodName -and $methodName.Trim() -ne '') { 
                                # Convert camelCase to Title Case
                                $methodName = $methodName -creplace '([a-z])([A-Z])', '$1 $2'
                                # Try to get displayName if available
                                $displayName = & $getProp 'displayName'
                                if ($displayName) { "$methodName ($displayName)" } else { $methodName }
                            } else { 
                                'MFA Method (Type Unknown)'
                            }
                        }
                    }
                } | Where-Object { $_ -ne $null -and $_ -ne '' }
                $results.PerUserMfa.Methods = $methodNames
                if ($methodNames.Count -gt 0) {
                    $results.PerUserMfa.Details = "Methods: $($methodNames -join ', ')"
                } else {
                    # Methods exist but we couldn't parse them - still show as enabled
                    $results.PerUserMfa.Details = "MFA methods registered (unable to determine specific types)"
                }
            } else {
                $results.PerUserMfa.Enabled = $false
                $results.PerUserMfa.Details = "No MFA methods registered"
            }
        }
        $securityDefaults = Get-MgPolicyIdentitySecurityDefaultEnforcementPolicy -ErrorAction SilentlyContinue
        if ($securityDefaults) {
            $results.SecurityDefaults.Enabled = $securityDefaults.IsEnabled
            $results.SecurityDefaults.Details = if ($securityDefaults.IsEnabled) { "Enabled (requires MFA for all users)" } else { "Disabled" }
            # If Security Defaults is enabled, MFA is required for this user
            if ($securityDefaults.IsEnabled) {
                $results.PerUserMfa.Enabled = $true
                if ($results.PerUserMfa.Details -eq "Not configured" -or $results.PerUserMfa.Details -eq "No MFA methods registered") {
                    $results.PerUserMfa.Details = "MFA required via Security Defaults"
                }
            }
        }
        $caPolicies = Get-MgIdentityConditionalAccessPolicy -All -ErrorAction SilentlyContinue
        if ($caPolicies) {
            $applicablePolicies = @()
            foreach ($policy in $caPolicies) {
                if ($policy.State -eq "enabled") {
                    $appliesToUser = $false
                    if (($policy.Conditions.Users.IncludeUsers -contains "All") -or ($policy.Conditions.Users.IncludeUsers -contains $user.Id)) { $appliesToUser = $true }
                    if ($policy.Conditions.Users.ExcludeUsers -contains $user.Id) { $appliesToUser = $false }
                    if ($appliesToUser) {
                        $requiresMfa = $false
                        if ($policy.GrantControls.BuiltInControls -contains "mfa") { $requiresMfa = $true }
                        $policyInfo = @{ Name = $policy.DisplayName; State = $policy.State; Controls = $policy.GrantControls.BuiltInControls -join ", "; Conditions = $policy.Conditions | Out-String; RequiresMfa = $requiresMfa }
                        $applicablePolicies += $policyInfo
                        if ($requiresMfa) { 
                            $results.ConditionalAccess.RequiresMfa = $true
                            # If CA requires MFA, MFA is required for this user
                            $results.PerUserMfa.Enabled = $true
                            if ($results.PerUserMfa.Details -eq "Not configured" -or $results.PerUserMfa.Details -eq "No MFA methods registered") {
                                $results.PerUserMfa.Details = "MFA required via Conditional Access policy"
                            }
                        }
                    }
                }
            }
            $results.ConditionalAccess.Policies = $applicablePolicies
            if ($applicablePolicies.Count -gt 0) {
                $mfaPoliciesCount = ($applicablePolicies | Where-Object { $_.RequiresMfa }).Count
                $results.ConditionalAccess.Details = "Found $($applicablePolicies.Count) applicable policies ($($mfaPoliciesCount) require MFA)"
            }
        }
        if ($results.SecurityDefaults.Enabled) {
            $results.OverallStatus = "Protected (Security Defaults)"
            $results.Summary = "MFA required via Security Defaults."
        } elseif ($results.ConditionalAccess.RequiresMfa) {
            $results.OverallStatus = "Protected (Conditional Access)"
            $results.Summary = "MFA required via one or more Conditional Access policies."
        } elseif ($results.PerUserMfa.Enabled) {
            $results.OverallStatus = "Protected (Per-User MFA)"
            $results.Summary = "MFA methods are registered, but protection may not be enforced by policy."
        } else {
            $results.OverallStatus = "⚠️ NOT PROTECTED"
            $results.Summary = "No MFA enforcement method detected."
        }
    } catch {
        $results.OverallStatus = "Error"
        $results.Summary = "Failed to analyze MFA status: $_"
    }
    return $results
}

Export-ModuleMember -Function Test-EntraModules,Install-EntraModules,Connect-EntraGraph,Get-EntraUsers,Get-EntraSignInLogs,Get-EntraUserDetailsAndRoles,Get-EntraUserAuditLogs,Get-EntraUserMfaStatus 