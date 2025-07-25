# EntraInvestigator.psm1 - Clean Rebuild
# Essential Entra ID/Graph functions for modular use

$script:requiredModules = @(
    "Microsoft.Graph.Users",
    "Microsoft.Graph.Reports",
    "Microsoft.Graph.Identity.DirectoryManagement",
    "Microsoft.Graph.Identity.SignIns"
)
$script:requiredScopes = @(
    "User.Read.All", "AuditLog.Read.All", "Organization.Read.All", "Directory.Read.All", "Policy.Read.All", "UserAuthenticationMethod.Read.All"
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
        Write-Error "Failed to connect to Microsoft Graph: $_"
        return $false
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
    param()
    try {
        $result = Get-MgUser -All -Property UserPrincipalName,DisplayName,Id -ConsistencyLevel eventual | Sort-Object UserPrincipalName
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
            $mfaMethods = $authMethods | Where-Object { $_.'@odata.type' -ne '#microsoft.graph.passwordAuthenticationMethod' }
            if ($mfaMethods) {
                $results.PerUserMfa.Enabled = $true
                $results.PerUserMfa.Methods = $mfaMethods | ForEach-Object { $_.'@odata.type' -replace '#microsoft.graph.', '' -replace 'AuthenticationMethod', '' }
                $results.PerUserMfa.Details = "Methods: $($results.PerUserMfa.Methods -join ', ')"
            } else {
                $results.PerUserMfa.Details = "No MFA methods registered"
            }
        }
        $securityDefaults = Get-MgPolicyIdentitySecurityDefaultEnforcementPolicy -ErrorAction SilentlyContinue
        if ($securityDefaults) {
            $results.SecurityDefaults.Enabled = $securityDefaults.IsEnabled
            $results.SecurityDefaults.Details = if ($securityDefaults.IsEnabled) { "Enabled (requires MFA for all users)" } else { "Disabled" }
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
                        if ($requiresMfa) { $results.ConditionalAccess.RequiresMfa = $true }
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