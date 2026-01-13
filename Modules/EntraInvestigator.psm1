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
            # Get sign-in logs - appliedConditionalAccessPolicies should be included by default with AuditLog.Read.All permission
            $logs = Get-MgAuditLogSignIn -Filter $filter -All -ErrorAction Stop
            if ($logs) {
                foreach ($log in $logs) {
                    try {
                        # Extract Status information
                        $resultStatusCode = $null
                        $resultStatus = "Unknown"
                        if ($log.Status) {
                            $resultStatusCode = $log.Status.ErrorCode
                            if ($log.Status.ErrorCode -eq 0) {
                                $resultStatus = "Success"
                            } elseif ($log.Status.FailureReason) {
                                $resultStatus = $log.Status.FailureReason
                            } elseif ($log.Status.AdditionalDetails) {
                                $resultStatus = $log.Status.AdditionalDetails
                            } else {
                                $resultStatus = "Failure (Code: $($log.Status.ErrorCode))"
                            }
                        }
                        
                        # Extract DeviceDetail information
                        $deviceId = $null
                        $deviceDetailJSON = $null
                        $deviceIsCompliant = $null
                        $deviceIsManaged = $null
                        if ($log.DeviceDetail) {
                            $deviceId = $log.DeviceDetail.DeviceId
                            $deviceIsCompliant = $log.DeviceDetail.IsCompliant
                            $deviceIsManaged = $log.DeviceDetail.IsManaged
                            try {
                                $deviceDetailJSON = $log.DeviceDetail | ConvertTo-Json -Depth 10 -Compress
                            } catch {
                                $deviceDetailJSON = "Error serializing DeviceDetail"
                            }
                        }
                        
                        # Extract Conditional Access Policies
                        $conditionalAccessPoliciesJSON = $null
                        $caPolicyNames = @()
                        $caPolicyResults = @()
                        $caPolicyDetails = @()  # For detailed policy info
                        
                        # Helper function to safely get property value
                        $getProperty = {
                            param($obj, $propNames)
                            foreach ($propName in $propNames) {
                                if ($obj.PSObject.Properties[$propName]) {
                                    return $obj.PSObject.Properties[$propName].Value
                                }
                            }
                            return $null
                        }
                        
                        # Check for ConditionalAccessPolicies in multiple ways
                        # Try appliedConditionalAccessPolicies first (the correct property name per Microsoft Graph API)
                        $caPolicies = $null
                        
                        # Debug: Check what properties are available (only for first log entry to avoid spam)
                        if ($allLogs.Count -eq 0 -and $log.ConditionalAccessStatus -eq "failure") {
                            $availableProps = $log.PSObject.Properties.Name | Where-Object { $_ -like "*onditional*" -or $_ -like "*pplied*" }
                            if ($availableProps) {
                                Write-Host "  Debug: Available CA-related properties: $($availableProps -join ', ')" -ForegroundColor Gray
                            }
                            if ($log.AdditionalProperties) {
                                $additionalCAProps = $log.AdditionalProperties.Keys | Where-Object { $_ -like "*onditional*" -or $_ -like "*pplied*" }
                                if ($additionalCAProps) {
                                    Write-Host "  Debug: CA properties in AdditionalProperties: $($additionalCAProps -join ', ')" -ForegroundColor Gray
                                }
                            }
                        }
                        
                        # Try appliedConditionalAccessPolicies first (correct property name)
                        if ($log.AppliedConditionalAccessPolicies) {
                            $caPolicies = $log.AppliedConditionalAccessPolicies
                        } elseif ($log.appliedConditionalAccessPolicies) {
                            $caPolicies = $log.appliedConditionalAccessPolicies
                        } elseif ($log.PSObject.Properties['AppliedConditionalAccessPolicies']) {
                            $caPolicies = $log.PSObject.Properties['AppliedConditionalAccessPolicies'].Value
                        } elseif ($log.PSObject.Properties['appliedConditionalAccessPolicies']) {
                            $caPolicies = $log.PSObject.Properties['appliedConditionalAccessPolicies'].Value
                        } elseif ($log.AdditionalProperties -and $log.AdditionalProperties['appliedConditionalAccessPolicies']) {
                            $caPolicies = $log.AdditionalProperties['appliedConditionalAccessPolicies']
                        } elseif ($log.AdditionalProperties -and $log.AdditionalProperties.ContainsKey('appliedConditionalAccessPolicies')) {
                            $caPolicies = $log.AdditionalProperties['appliedConditionalAccessPolicies']
                        } elseif ($log.ConditionalAccessPolicies) {
                            $caPolicies = $log.ConditionalAccessPolicies
                        } elseif ($log.conditionalAccessPolicies) {
                            $caPolicies = $log.conditionalAccessPolicies
                        } elseif ($log.PSObject.Properties['ConditionalAccessPolicies']) {
                            $caPolicies = $log.PSObject.Properties['ConditionalAccessPolicies'].Value
                        } elseif ($log.PSObject.Properties['conditionalAccessPolicies']) {
                            $caPolicies = $log.PSObject.Properties['conditionalAccessPolicies'].Value
                        } elseif ($log.AdditionalProperties -and $log.AdditionalProperties['conditionalAccessPolicies']) {
                            $caPolicies = $log.AdditionalProperties['conditionalAccessPolicies']
                        }
                        
                        if ($caPolicies) {
                            # Handle array or single object
                            if ($caPolicies -is [Array]) {
                                $policyArray = $caPolicies
                            } elseif ($caPolicies.Count) {
                                $policyArray = @($caPolicies)
                            } else {
                                $policyArray = @($caPolicies)
                            }
                            
                            if ($policyArray.Count -gt 0) {
                                try {
                                    $conditionalAccessPoliciesJSON = $policyArray | ConvertTo-Json -Depth 10 -Compress
                                    foreach ($policy in $policyArray) {
                                        # Get policy name - appliedConditionalAccessPolicies uses 'id' and 'displayName'
                                        $policyName = & $getProperty $policy @('DisplayName', 'displayName', 'name', 'Name')
                                        
                                        # Get policy result - appliedConditionalAccessPolicies uses 'result'
                                        $policyResult = & $getProperty $policy @('Result', 'result', 'outcome', 'Outcome')
                                        
                                        # Get policy ID - appliedConditionalAccessPolicies uses 'id'
                                        $policyId = & $getProperty $policy @('Id', 'id', 'policyId', 'PolicyId')
                                        
                                        # Get applied grant controls (what was enforced)
                                        $appliedGrantControls = & $getProperty $policy @('AppliedGrantControls', 'appliedGrantControls', 'GrantControls', 'grantControls')
                                        
                                        # Get conditions that triggered (if available)
                                        $conditions = & $getProperty $policy @('Conditions', 'conditions')
                                        
                                        # If still no name, try to get from Id by looking up policy
                                        if (-not $policyName -and $policyId) {
                                            try {
                                                # Try to get policy name from Graph API if we have the ID
                                                $policyObj = Get-MgIdentityConditionalAccessPolicy -ConditionalAccessPolicyId $policyId -ErrorAction SilentlyContinue
                                                if ($policyObj -and $policyObj.DisplayName) {
                                                    $policyName = $policyObj.DisplayName
                                                }
                                            } catch {
                                                # Ignore lookup errors
                                            }
                                        }
                                        
                                        # If we have a policy ID but no name, use the ID as fallback
                                        if (-not $policyName -and $policyId) {
                                            $policyName = "Policy ID: $policyId"
                                        }
                                        
                                        # Only add if we have a policy identifier (name or ID)
                                        if ($policyName -or $policyId) {
                                            if ($policyName) {
                                                $caPolicyNames += $policyName
                                            } else {
                                                $caPolicyNames += $policyId
                                            }
                                            
                                            # Format result for readability
                                            if ($policyResult) {
                                                $resultText = $policyResult.ToString()
                                                # Normalize result values
                                                if ($resultText -eq "failure" -or $resultText -eq "Failure" -or $resultText -eq "0") {
                                                    $resultText = "Failed"
                                                } elseif ($resultText -eq "success" -or $resultText -eq "Success" -or $resultText -eq "1") {
                                                    $resultText = "Success"
                                                } elseif ($resultText -eq "notApplied" -or $resultText -eq "NotApplied" -or $resultText -eq "2") {
                                                    $resultText = "Not Applied"
                                                }
                                                $caPolicyResults += $resultText
                                                
                                                # Create detailed entry for blocked policies
                                                $detailParts = @($policyName)
                                                if ($resultText -eq "Failed") {
                                                    $detailParts += "(BLOCKED)"
                                                } else {
                                                    $detailParts += "($resultText)"
                                                }
                                                
                                                # Add grant controls info if available
                                                if ($appliedGrantControls) {
                                                    $grantControlsStr = $null
                                                    if ($appliedGrantControls -is [Array]) {
                                                        $grantControlsStr = $appliedGrantControls -join ", "
                                                    } elseif ($appliedGrantControls.PSObject.Properties['builtInControls']) {
                                                        $grantControlsStr = $appliedGrantControls.builtInControls -join ", "
                                                    } elseif ($appliedGrantControls.PSObject.Properties['BuiltInControls']) {
                                                        $grantControlsStr = $appliedGrantControls.BuiltInControls -join ", "
                                                    }
                                                    if ($grantControlsStr) {
                                                        $detailParts += "[$grantControlsStr]"
                                                    }
                                                }
                                                
                                                $caPolicyDetails += ($detailParts -join " ")
                                            } else {
                                                $caPolicyResults += "Unknown"
                                                $caPolicyDetails += $policyName
                                            }
                                        }
                                    }
                                } catch {
                                    $conditionalAccessPoliciesJSON = "Error serializing ConditionalAccessPolicies: $($_.Exception.Message)"
                                }
                            }
                        }
                        
                        # Extract Authentication Details
                        $authenticationDetailsJSON = $null
                        $authMethods = @()
                        if ($log.AuthenticationDetails -and $log.AuthenticationDetails.Count -gt 0) {
                            try {
                                $authenticationDetailsJSON = $log.AuthenticationDetails | ConvertTo-Json -Depth 10 -Compress
                                foreach ($authDetail in $log.AuthenticationDetails) {
                                    if ($authDetail.AuthenticationMethod) {
                                        $authMethods += $authDetail.AuthenticationMethod
                                    } elseif ($authDetail.AuthenticationMethodDetail) {
                                        $authMethods += $authDetail.AuthenticationMethodDetail
                                    }
                                }
                            } catch {
                                $authenticationDetailsJSON = "Error serializing AuthenticationDetails"
                            }
                        }
                        
                        # Create enhanced log object with all fields
                        $enhancedLog = [PSCustomObject]@{
                            UserPrincipalName = $upn
                            CreatedDateTime = $log.CreatedDateTime
                            UserId = $log.UserId
                            AppDisplayName = $log.AppDisplayName
                            ClientAppUsed = $log.ClientAppUsed
                            IPAddress = $log.IpAddress
                            Location = if ($log.Location) {
                                $locParts = @()
                                if ($log.Location.City) { $locParts += $log.Location.City }
                                if ($log.Location.State) { $locParts += $log.Location.State }
                                if ($log.Location.CountryOrRegion) { $locParts += $log.Location.CountryOrRegion }
                                if ($locParts.Count -gt 0) { $locParts -join ", " } else { "Unknown" }
                            } else { "Unknown" }
                            CountryOrRegion = if ($log.Location) { $log.Location.CountryOrRegion } else { $null }
                            ResultStatusCode = $resultStatusCode
                            ResultStatus = $resultStatus
                            DeviceId = $deviceId
                            DeviceDetailJSON = $deviceDetailJSON
                            DeviceIsCompliant = $deviceIsCompliant
                            DeviceIsManaged = $deviceIsManaged
                            DeviceDetail = if ($log.DeviceDetail) {
                                $deviceParts = @()
                                if ($log.DeviceDetail.Browser) { $deviceParts += $log.DeviceDetail.Browser }
                                if ($log.DeviceDetail.OperatingSystem) { $deviceParts += $log.DeviceDetail.OperatingSystem }
                                if ($deviceParts.Count -gt 0) { $deviceParts -join " / " } else { "Unknown" }
                            } else { "Unknown" }
                            ConditionalAccessStatus = if ($log.ConditionalAccessStatus) { $log.ConditionalAccessStatus } else { "Not Applied" }
                            ConditionalAccessPoliciesJSON = $conditionalAccessPoliciesJSON
                            CAPolicyNames = $caPolicyNames -join "; "
                            CAPolicyResults = $caPolicyResults -join "; "
                            CAPolicyDetails = $caPolicyDetails -join "; "
                            AuthenticationDetailsJSON = $authenticationDetailsJSON
                            AuthMethods = $authMethods -join "; "
                            RiskLevelAggregated = $log.RiskLevelAggregated
                            RiskLevelDuringSignIn = $log.RiskLevelDuringSignIn
                            RiskState = $log.RiskState
                            ResourceDisplayName = $log.ResourceDisplayName
                            ResourceId = $log.ResourceId
                        }
                        
                        # Preserve all original properties by copying them
                        $log.PSObject.Properties | ForEach-Object {
                            if (-not $enhancedLog.PSObject.Properties[$_.Name]) {
                                try {
                                    $enhancedLog | Add-Member -MemberType NoteProperty -Name $_.Name -Value $_.Value -Force
                                } catch {
                                    # Skip properties that can't be added
                                }
                            }
                        }
                        
                        $allLogs += $enhancedLog
                    } catch {
                        Write-Warning "Error processing sign-in log entry for $upn : $($_.Exception.Message)"
                        # Fallback: add log with UserPrincipalName only
                        $log | Add-Member -MemberType NoteProperty -Name 'UserPrincipalName' -Value $upn -Force
                        $allLogs += $log
                    }
                }
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

function Get-IntuneDeviceRecords {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$false)]
        [string[]]$DeviceIds = @(),
        [Parameter(Mandatory=$false)]
        [string[]]$UserPrincipalNames = @(),
        [Parameter(Mandatory=$false)]
        [int]$DaysBack = 30
    )
    
    $allDevices = @()
    
    try {
        # Check if Device Management module is available
        $deviceMgmtAvailable = $false
        $hasCmdlet = $false
        $hasGraphAPI = $false
        try {
            if (Get-Command Get-MgDeviceManagementManagedDevice -ErrorAction SilentlyContinue) {
                $deviceMgmtAvailable = $true
                $hasCmdlet = $true
                Write-Host "  Device Management cmdlet available: Get-MgDeviceManagementManagedDevice" -ForegroundColor Gray
            }
            if (Get-Command Invoke-MgGraphRequest -ErrorAction SilentlyContinue) {
                $deviceMgmtAvailable = $true
                $hasGraphAPI = $true
                Write-Host "  Graph API available: Invoke-MgGraphRequest" -ForegroundColor Gray
            }
        } catch {
            Write-Warning "Error checking for Device Management APIs: $_"
        }
        
        if (-not $deviceMgmtAvailable) {
            Write-Warning "Device Management APIs not available. Intune device records require Microsoft.Graph.DeviceManagement module or Graph API access."
            Write-Host "  Attempting to check Graph connection..." -ForegroundColor Yellow
            try {
                $context = Get-MgContext -ErrorAction SilentlyContinue
                if ($context) {
                    Write-Host "  Graph context found: $($context.TenantId)" -ForegroundColor Gray
                } else {
                    Write-Warning "  No Graph context found - may need to connect to Microsoft Graph"
                }
            } catch {
                Write-Warning "  Could not check Graph context: $_"
            }
            return @()
        }
        
        # Get user IDs if UserPrincipalNames provided
        $userIds = @()
        if ($UserPrincipalNames -and $UserPrincipalNames.Count -gt 0) {
            foreach ($upn in $UserPrincipalNames) {
                try {
                    $user = Get-MgUser -UserId $upn -Property Id -ErrorAction SilentlyContinue
                    if ($user) { $userIds += $user.Id }
                } catch {
                    Write-Warning "Could not resolve user ID for $upn"
                }
            }
        }
        
        # Retrieve managed devices
        $managedDevices = @()
        
        # First, try Intune managed devices API
        if (Get-Command Get-MgDeviceManagementManagedDevice -ErrorAction SilentlyContinue) {
            # Use cmdlet if available
            if ($DeviceIds -and $DeviceIds.Count -gt 0) {
                foreach ($deviceId in $DeviceIds) {
                    try {
                        $device = Get-MgDeviceManagementManagedDevice -ManagedDeviceId $deviceId -ErrorAction SilentlyContinue
                        if ($device) { $managedDevices += $device }
                    } catch {
                        Write-Warning "Could not retrieve Intune managed device $deviceId : $_"
                    }
                }
            } else {
                # Get all devices, optionally filtered by user
                try {
                    Write-Host "  Querying Intune managed devices..." -ForegroundColor Gray
                    Write-Host "  Calling Get-MgDeviceManagementManagedDevice -All..." -ForegroundColor DarkGray
                    # Request UserId and UserPrincipalName properties explicitly
                    $allDevicesRaw = Get-MgDeviceManagementManagedDevice -All -Property Id,DeviceName,UserId,UserPrincipalName,ComplianceState,LastSyncDateTime,EnrolledDateTime,OperatingSystem,OSVersion,ManagementState,DeviceType,Ownership -ErrorAction Stop
                    Write-Host "  Retrieved $($allDevicesRaw.Count) total Intune managed device(s) from API" -ForegroundColor Green
                    if ($allDevicesRaw.Count -eq 0) {
                        Write-Warning "  Warning: API returned 0 devices, but Intune admin center shows devices exist"
                        Write-Host "  This may indicate a permission issue or API scope problem" -ForegroundColor Yellow
                    }
                    if ($userIds.Count -gt 0) {
                        $managedDevices = $allDevicesRaw | Where-Object { $userIds -contains $_.UserId }
                        Write-Host "  Filtered to $($managedDevices.Count) device(s) for selected user(s)" -ForegroundColor Gray
                    } else {
                        $managedDevices = $allDevicesRaw
                    }
                    Write-Host "  Found $($managedDevices.Count) Intune managed device(s) to process" -ForegroundColor Green
                } catch {
                    $errorMsg = $_.Exception.Message
                    Write-Warning "Could not retrieve Intune managed devices: $errorMsg"
                    if ($errorMsg -like "*permission*" -or $errorMsg -like "*Forbidden*" -or $errorMsg -like "*access denied*") {
                        Write-Warning "  Permission issue detected - requires DeviceManagementManagedDevices.Read.All permission"
                    } elseif ($errorMsg -like "*not found*" -or $errorMsg -like "*does not exist*") {
                        Write-Warning "  Device Management API may not be available - check if Microsoft.Graph.DeviceManagement module is installed"
                    } else {
                        Write-Warning "  Error details: $errorMsg"
                        Write-Warning "  Exception type: $($_.Exception.GetType().FullName)"
                    }
                }
            }
        } else {
            # Fallback to Graph API request for Intune
            try {
                $uri = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices"
                if ($DeviceIds -and $DeviceIds.Count -gt 0) {
                    # Get specific devices
                    foreach ($deviceId in $DeviceIds) {
                        try {
                            $deviceUri = "$uri/$deviceId"
                            $device = Invoke-MgGraphRequest -Method GET -Uri $deviceUri -ErrorAction SilentlyContinue
                            if ($device) { $managedDevices += $device }
                        } catch {
                            Write-Warning "Could not retrieve Intune device $deviceId via API: $_"
                        }
                    }
                } else {
                    # Get all devices
                    Write-Host "  Querying Intune managed devices via Graph API..." -ForegroundColor Gray
                    Write-Host "  Calling: GET $uri" -ForegroundColor DarkGray
                    $response = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
                    if ($response.value) {
                        Write-Host "  Retrieved $($response.value.Count) device(s) from Graph API" -ForegroundColor Green
                        if ($userIds.Count -gt 0) {
                            $managedDevices = $response.value | Where-Object { $userIds -contains $_.userId }
                            Write-Host "  Filtered to $($managedDevices.Count) device(s) for selected user(s)" -ForegroundColor Gray
                        } else {
                            $managedDevices = $response.value
                        }
                        Write-Host "  Found $($managedDevices.Count) Intune managed device(s)" -ForegroundColor Green
                    } else {
                        Write-Warning "  Graph API returned response but no 'value' property found"
                        Write-Host "  Response keys: $($response.PSObject.Properties.Name -join ', ')" -ForegroundColor Yellow
                    }
                }
            } catch {
                $errorMsg = $_.Exception.Message
                Write-Warning "Could not retrieve Intune managed devices via Graph API: $errorMsg"
                if ($errorMsg -like "*permission*" -or $errorMsg -like "*Forbidden*" -or $errorMsg -like "*403*") {
                    Write-Warning "  Permission denied - requires DeviceManagementManagedDevices.Read.All permission"
                    Write-Host "  Current Graph scopes may not include device management permissions" -ForegroundColor Yellow
                } elseif ($errorMsg -like "*404*" -or $errorMsg -like "*not found*") {
                    Write-Warning "  API endpoint not found - Intune may not be licensed or configured"
                } else {
                    Write-Warning "  Full error: $($_.Exception | Out-String)"
                }
            }
        }
        
        # Also try Azure AD devices API (for Azure AD joined devices that may not be Intune enrolled)
        # Query Azure AD devices if:
        # 1. DeviceIds provided (from sign-in logs) - query those specific devices
        # 2. No Intune devices found and user filtering - query Azure AD devices for those users
        # 3. No devices found at all - query all Azure AD devices as fallback
        $shouldQueryAzureAD = $false
        if ($DeviceIds -and $DeviceIds.Count -gt 0) {
            # Always query Azure AD devices when DeviceIds are provided (from sign-in logs)
            $shouldQueryAzureAD = $true
        } elseif ($managedDevices.Count -eq 0 -and $userIds.Count -gt 0) {
            # Query Azure AD devices for selected users if no Intune devices found
            $shouldQueryAzureAD = $true
        } elseif ($managedDevices.Count -eq 0) {
            # Query all Azure AD devices as fallback if nothing found
            $shouldQueryAzureAD = $true
        }
        
        if ($shouldQueryAzureAD) {
            Write-Host "  Querying Azure AD devices..." -ForegroundColor Gray
            try {
                if ($DeviceIds -and $DeviceIds.Count -gt 0) {
                    # Query specific devices by DeviceId
                    foreach ($deviceId in $DeviceIds) {
                        try {
                            # Try Azure AD devices API
                            $azureDevice = Get-MgDevice -DeviceId $deviceId -ErrorAction SilentlyContinue
                            if ($azureDevice) {
                                # Convert Azure AD device to a compatible format
                                $azureDeviceRecord = [PSCustomObject]@{
                                    Id = $azureDevice.Id
                                    DeviceId = $azureDevice.DeviceId
                                    DeviceName = if ($azureDevice.DisplayName) { $azureDevice.DisplayName } else { "Unknown" }
                                    UserId = if ($azureDevice.RegisteredUsers) { ($azureDevice.RegisteredUsers | Select-Object -First 1).Id } else { $null }
                                    OperatingSystem = if ($azureDevice.OperatingSystem) { $azureDevice.OperatingSystem } else { "Unknown" }
                                    OSVersion = if ($azureDevice.OperatingSystemVersion) { $azureDevice.OperatingSystemVersion } else { $null }
                                    IsCompliant = if ($azureDevice.IsCompliant) { $azureDevice.IsCompliant } else { $null }
                                    IsManaged = if ($azureDevice.IsManaged) { $azureDevice.IsManaged } else { $null }
                                    TrustType = if ($azureDevice.TrustType) { $azureDevice.TrustType } else { "Unknown" }
                                    IsAzureADJoined = ($azureDevice.TrustType -eq "AzureAd")
                                    Source = "AzureAD"
                                }
                                $managedDevices += $azureDeviceRecord
                            }
                        } catch {
                            # Try via Graph API if cmdlet fails
                            try {
                                $deviceUri = "https://graph.microsoft.com/v1.0/devices/$deviceId"
                                $azureDevice = Invoke-MgGraphRequest -Method GET -Uri $deviceUri -ErrorAction SilentlyContinue
                                if ($azureDevice) {
                                    $azureDeviceRecord = [PSCustomObject]@{
                                        Id = $azureDevice.id
                                        DeviceId = $azureDevice.deviceId
                                        DeviceName = if ($azureDevice.displayName) { $azureDevice.displayName } else { "Unknown" }
                                        UserId = if ($azureDevice.registeredUsers) { ($azureDevice.registeredUsers | Select-Object -First 1).id } else { $null }
                                        OperatingSystem = if ($azureDevice.operatingSystem) { $azureDevice.operatingSystem } else { "Unknown" }
                                        OSVersion = if ($azureDevice.operatingSystemVersion) { $azureDevice.operatingSystemVersion } else { $null }
                                        IsCompliant = if ($azureDevice.isCompliant) { $azureDevice.isCompliant } else { $null }
                                        IsManaged = if ($azureDevice.isManaged) { $azureDevice.isManaged } else { $null }
                                        TrustType = if ($azureDevice.trustType) { $azureDevice.trustType } else { "Unknown" }
                                        IsAzureADJoined = ($azureDevice.trustType -eq "AzureAd")
                                        Source = "AzureAD"
                                    }
                                    $managedDevices += $azureDeviceRecord
                                }
                            } catch {
                                Write-Warning "Could not retrieve Azure AD device $deviceId : $_"
                            }
                        }
                    }
                } else {
                    # Query all Azure AD devices, optionally filtered by user
                    try {
                        if (Get-Command Get-MgDevice -ErrorAction SilentlyContinue) {
                            $allAzureDevices = Get-MgDevice -All -ErrorAction Stop
                            if ($userIds.Count -gt 0) {
                                # Filter by registered users
                                $filteredAzureDevices = @()
                                foreach ($azureDevice in $allAzureDevices) {
                                    if ($azureDevice.RegisteredUsers) {
                                        foreach ($registeredUser in $azureDevice.RegisteredUsers) {
                                            if ($userIds -contains $registeredUser.Id) {
                                                $filteredAzureDevices += $azureDevice
                                                break
                                            }
                                        }
                                    }
                                }
                                $allAzureDevices = $filteredAzureDevices
                            }
                            
                            foreach ($azureDevice in $allAzureDevices) {
                                $azureDeviceRecord = [PSCustomObject]@{
                                    Id = $azureDevice.Id
                                    DeviceId = $azureDevice.DeviceId
                                    DeviceName = if ($azureDevice.DisplayName) { $azureDevice.DisplayName } else { "Unknown" }
                                    UserId = if ($azureDevice.RegisteredUsers) { ($azureDevice.RegisteredUsers | Select-Object -First 1).Id } else { $null }
                                    OperatingSystem = if ($azureDevice.OperatingSystem) { $azureDevice.OperatingSystem } else { "Unknown" }
                                    OSVersion = if ($azureDevice.OperatingSystemVersion) { $azureDevice.OperatingSystemVersion } else { $null }
                                    IsCompliant = if ($azureDevice.IsCompliant) { $azureDevice.IsCompliant } else { $null }
                                    IsManaged = if ($azureDevice.IsManaged) { $azureDevice.IsManaged } else { $null }
                                    TrustType = if ($azureDevice.TrustType) { $azureDevice.TrustType } else { "Unknown" }
                                    IsAzureADJoined = ($azureDevice.TrustType -eq "AzureAd")
                                    Source = "AzureAD"
                                }
                                $managedDevices += $azureDeviceRecord
                            }
                        } else {
                            # Fallback to Graph API
                            $uri = "https://graph.microsoft.com/v1.0/devices"
                            $response = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
                            if ($response.value) {
                                $azureDevicesToProcess = $response.value
                                if ($userIds.Count -gt 0) {
                                    # Filter by registered users
                                    $filteredAzureDevices = @()
                                    foreach ($azureDevice in $azureDevicesToProcess) {
                                        if ($azureDevice.registeredUsers) {
                                            foreach ($registeredUser in $azureDevice.registeredUsers) {
                                                if ($userIds -contains $registeredUser.id) {
                                                    $filteredAzureDevices += $azureDevice
                                                    break
                                                }
                                            }
                                        }
                                    }
                                    $azureDevicesToProcess = $filteredAzureDevices
                                }
                                
                                foreach ($azureDevice in $azureDevicesToProcess) {
                                    $azureDeviceRecord = [PSCustomObject]@{
                                        Id = $azureDevice.id
                                        DeviceId = $azureDevice.deviceId
                                        DeviceName = if ($azureDevice.displayName) { $azureDevice.displayName } else { "Unknown" }
                                        UserId = if ($azureDevice.registeredUsers) { ($azureDevice.registeredUsers | Select-Object -First 1).id } else { $null }
                                        OperatingSystem = if ($azureDevice.operatingSystem) { $azureDevice.operatingSystem } else { "Unknown" }
                                        OSVersion = if ($azureDevice.operatingSystemVersion) { $azureDevice.operatingSystemVersion } else { $null }
                                        IsCompliant = if ($azureDevice.isCompliant) { $azureDevice.isCompliant } else { $null }
                                        IsManaged = if ($azureDevice.isManaged) { $azureDevice.isManaged } else { $null }
                                        TrustType = if ($azureDevice.trustType) { $azureDevice.trustType } else { "Unknown" }
                                        IsAzureADJoined = ($azureDevice.trustType -eq "AzureAd")
                                        Source = "AzureAD"
                                    }
                                    $managedDevices += $azureDeviceRecord
                                }
                            }
                        }
                    } catch {
                        Write-Warning "Could not query Azure AD devices: $_"
                    }
                }
                if ($managedDevices.Count -gt 0) {
                    Write-Host "  Found $($managedDevices.Count) Azure AD device(s)" -ForegroundColor Green
                }
            } catch {
                Write-Warning "Could not query Azure AD devices: $_"
            }
        }
        
        Write-Host "  Processing $($managedDevices.Count) device(s)..." -ForegroundColor Gray
        
        # Pre-fetch user IDs to UPN mapping for better performance
        $userIdToUpnMap = @{}
        $uniqueUserIds = @()
        foreach ($device in $managedDevices) {
            # Get UserId from device
            $devUserId = if ($device.UserId) { $device.UserId } elseif ($device.userId) { $device.userId } elseif ($device.EnrolledUserId) { $device.EnrolledUserId } elseif ($device.enrolledUserId) { $device.enrolledUserId } else { $null }
            if ($devUserId -and $uniqueUserIds -notcontains $devUserId) {
                $uniqueUserIds += $devUserId
            }
            
            # Also check RegisteredUsers for Azure AD devices
            if ($device.RegisteredUsers) {
                foreach ($regUser in $device.RegisteredUsers) {
                    $regUserId = if ($regUser.Id) { $regUser.Id } elseif ($regUser.id) { $regUser.id } else { $null }
                    if ($regUserId -and $uniqueUserIds -notcontains $regUserId) {
                        $uniqueUserIds += $regUserId
                    }
                }
            }
        }
        
        if ($uniqueUserIds.Count -gt 0) {
            Write-Host "  Resolving UserPrincipalNames for $($uniqueUserIds.Count) unique user(s)..." -ForegroundColor Gray
            # Batch user lookup with progress indicator for large counts
            $batchSize = 50
            $totalBatches = [Math]::Ceiling($uniqueUserIds.Count / $batchSize)
            $batchCount = 0
            for ($i = 0; $i -lt $uniqueUserIds.Count; $i += $batchSize) {
                $batch = $uniqueUserIds[$i..([Math]::Min($i + $batchSize - 1, $uniqueUserIds.Count - 1))]
                $batchCount++
                if ($uniqueUserIds.Count -gt 100) {
                    Write-Host "    Processing batch $batchCount of $totalBatches ($($batch.Count) users)..." -ForegroundColor DarkGray
                }
                foreach ($uid in $batch) {
                    try {
                        $user = Get-MgUser -UserId $uid -Property UserPrincipalName -ErrorAction SilentlyContinue
                        if ($user -and $user.UserPrincipalName) {
                            $userIdToUpnMap[$uid] = $user.UserPrincipalName
                        }
                    } catch {
                        # Silently fail - user might not exist
                    }
                }
            }
            Write-Host "  Resolved $($userIdToUpnMap.Count) UserPrincipalName(s)" -ForegroundColor Gray
        }
        
        # Process each device to extract compliance and ownership info
        $processedCount = 0
        $errorCount = 0
        foreach ($device in $managedDevices) {
            try {
                # Check if this is an Azure AD device (already processed) or Intune device (needs processing)
                $isAzureADDevice = ($device.Source -eq "AzureAD")
                
                if ($processedCount -eq 0) {
                    Write-Host "  First device type: $($device.GetType().FullName), Source: $($device.Source)" -ForegroundColor DarkGray
                    Write-Host "  First device properties: $($device.PSObject.Properties.Name -join ', ')" -ForegroundColor DarkGray
                    # Debug: Show user-related properties
                    $userProps = $device.PSObject.Properties.Name | Where-Object { $_ -like "*user*" -or $_ -like "*User*" -or $_ -like "*principal*" -or $_ -like "*Principal*" }
                    if ($userProps) {
                        Write-Host "  User-related properties found: $($userProps -join ', ')" -ForegroundColor Yellow
                        foreach ($prop in $userProps) {
                            $val = $device.$prop
                            if ($val) {
                                Write-Host "    $prop = $val" -ForegroundColor DarkYellow
                            }
                        }
                    } else {
                        Write-Host "  WARNING: No user-related properties found on device object" -ForegroundColor Red
                    }
                }
                
                if ($isAzureADDevice) {
                    # Azure AD device - already in compatible format, just need to map fields
                    $deviceId = $device.DeviceId
                    $deviceName = $device.DeviceName
                    $userId = $device.UserId
                    $userPrincipalName = $null
                    
                    # Try to resolve UPN for Azure AD devices using pre-fetched map
                    if ($userId) {
                        # Check pre-fetched map first
                        if ($userIdToUpnMap.ContainsKey($userId)) {
                            $userPrincipalName = $userIdToUpnMap[$userId]
                        } else {
                            # Fallback to direct lookup
                            try {
                                $user = Get-MgUser -UserId $userId -Property UserPrincipalName -ErrorAction SilentlyContinue
                                if ($user) { 
                                    $userPrincipalName = $user.UserPrincipalName
                                    $userIdToUpnMap[$userId] = $userPrincipalName  # Cache it
                                }
                            } catch {}
                        }
                    }
                    
                    # Azure AD devices might have RegisteredUsers array - check that too
                    if (-not $userPrincipalName -and $device.RegisteredUsers) {
                        $firstUser = $device.RegisteredUsers | Select-Object -First 1
                        if ($firstUser) {
                            $regUserId = if ($firstUser.Id) { $firstUser.Id } elseif ($firstUser.id) { $firstUser.id } else { $null }
                            if ($regUserId) {
                                if ($userIdToUpnMap.ContainsKey($regUserId)) {
                                    $userPrincipalName = $userIdToUpnMap[$regUserId]
                                } else {
                                    try {
                                        $user = Get-MgUser -UserId $regUserId -Property UserPrincipalName -ErrorAction SilentlyContinue
                                        if ($user) { 
                                            $userPrincipalName = $user.UserPrincipalName
                                            $userIdToUpnMap[$regUserId] = $userPrincipalName
                                            if (-not $userId) { $userId = $regUserId }  # Set userId if not already set
                                        }
                                    } catch {}
                                }
                            }
                        }
                    }
                    $complianceState = if ($device.IsCompliant -eq $true) { "Compliant" } elseif ($device.IsCompliant -eq $false) { "NonCompliant" } else { "Unknown" }
                    $isCompliant = $device.IsCompliant
                    $lastSyncDateTime = $null  # Azure AD devices don't have sync time
                    $enrolledDateTime = $null  # Azure AD devices don't have enrollment date in same format
                    $operatingSystem = $device.OperatingSystem
                    $osVersion = $device.OSVersion
                    $managementState = if ($device.IsManaged) { "Managed" } else { "Unmanaged" }
                    $isManaged = $device.IsManaged
                    $deviceType = $operatingSystem  # Use OS as device type for Azure AD devices
                    $ownership = if ($device.IsAzureADJoined) { "Corporate" } else { "Personal" }
                } else {
                    # Intune managed device - extract from Intune format
                    $deviceId = if ($device.Id) { $device.Id } elseif ($device.id) { $device.id } else { "Unknown" }
                    $deviceName = if ($device.DeviceName) { $device.DeviceName } elseif ($device.deviceName) { $device.deviceName } else { "Unknown" }
                    
                    # Get UserId - check multiple property names
                    $userId = $null
                    if ($device.UserId) { $userId = $device.UserId }
                    elseif ($device.userId) { $userId = $device.userId }
                    elseif ($device.EnrolledUserId) { $userId = $device.EnrolledUserId }
                    elseif ($device.enrolledUserId) { $userId = $device.enrolledUserId }
                    elseif ($device.EnrolledBy) { $userId = $device.EnrolledBy }
                    elseif ($device.enrolledBy) { $userId = $device.enrolledBy }
                    elseif ($device.UserPrincipalName) { 
                        # If we have UPN but not ID, try to resolve
                        try {
                            $user = Get-MgUser -UserId $device.UserPrincipalName -Property Id -ErrorAction SilentlyContinue
                            if ($user) { $userId = $user.Id }
                        } catch {}
                    }
                    elseif ($device.userPrincipalName) {
                        try {
                            $user = Get-MgUser -UserId $device.userPrincipalName -Property Id -ErrorAction SilentlyContinue
                            if ($user) { $userId = $user.Id }
                        } catch {}
                    }
                    
                    # Debug logging for first few devices
                    if ($processedCount -lt 3) {
                        Write-Host "  Device $processedCount - DeviceName: $deviceName, UserId found: $(if ($userId) { $userId } else { 'NULL' })" -ForegroundColor DarkGray
                    }
                    
                    # Resolve UPN - check device object first, then lookup from UserId
                    $userPrincipalName = $null
                    # First, check if UPN is directly on device object (most reliable if API returns it)
                    if ($device.UserPrincipalName) { 
                        $userPrincipalName = $device.UserPrincipalName
                        if ($processedCount -lt 3) {
                            Write-Host "  Device $processedCount - Found UPN directly on device: $userPrincipalName" -ForegroundColor Green
                        }
                    }
                    elseif ($device.userPrincipalName) { 
                        $userPrincipalName = $device.userPrincipalName
                        if ($processedCount -lt 3) {
                            Write-Host "  Device $processedCount - Found UPN directly on device (lowercase): $userPrincipalName" -ForegroundColor Green
                        }
                    }
                    # If UPN not on device, try to resolve from UserId
                    elseif ($userId) {
                        # Check pre-fetched map first
                        if ($userIdToUpnMap.ContainsKey($userId)) {
                            $userPrincipalName = $userIdToUpnMap[$userId]
                            if ($processedCount -lt 3) {
                                Write-Host "  Device $processedCount - Found UPN from cached map: $userPrincipalName" -ForegroundColor Green
                            }
                        } else {
                            # Skip individual lookups during processing to avoid slowdown
                            # User should have been resolved in batch lookup above
                            # If not in cache, it means the user doesn't exist or lookup failed
                            if ($processedCount -lt 3) {
                                Write-Host "  Device $processedCount - UserId $userId not in cache (skipping individual lookup for performance)" -ForegroundColor DarkYellow
                            }
                        }
                    } else {
                        if ($processedCount -lt 3) {
                            Write-Host "  Device $processedCount - No UserId found, cannot resolve UPN" -ForegroundColor DarkYellow
                        }
                    }
                    
                    # Debug: Log first few devices with their UPN status
                    if ($processedCount -lt 3) {
                        Write-Host "  Device $processedCount - Final UPN: $(if ($userPrincipalName) { $userPrincipalName } else { 'EMPTY' })" -ForegroundColor $(if ($userPrincipalName) { 'Green' } else { 'Red' })
                    }
                    
                    # Get compliance status
                    $complianceState = if ($device.ComplianceState) { $device.ComplianceState } elseif ($device.complianceState) { $device.complianceState } else { "Unknown" }
                    $isCompliant = $complianceState -eq "Compliant"
                    
                    # Get last sync time (check-in) - check multiple property names
                    $lastSyncDateTime = $null
                    if ($device.LastSyncDateTime) { $lastSyncDateTime = $device.LastSyncDateTime }
                    elseif ($device.lastSyncDateTime) { $lastSyncDateTime = $device.lastSyncDateTime }
                    elseif ($device.LastSyncTime) { $lastSyncDateTime = $device.LastSyncTime }
                    elseif ($device.lastSyncTime) { $lastSyncDateTime = $device.lastSyncTime }
                    elseif ($device.SyncDateTime) { $lastSyncDateTime = $device.SyncDateTime }
                    elseif ($device.syncDateTime) { $lastSyncDateTime = $device.syncDateTime }
                    elseif ($device.LastCheckInDateTime) { $lastSyncDateTime = $device.LastCheckInDateTime }
                    elseif ($device.lastCheckInDateTime) { $lastSyncDateTime = $device.lastCheckInDateTime }
                    
                    # Get enrollment date - check multiple property names
                    $enrolledDateTime = $null
                    if ($device.EnrolledDateTime) { $enrolledDateTime = $device.EnrolledDateTime }
                    elseif ($device.enrolledDateTime) { $enrolledDateTime = $device.enrolledDateTime }
                    elseif ($device.EnrollmentDateTime) { $enrolledDateTime = $device.EnrollmentDateTime }
                    elseif ($device.enrollmentDateTime) { $enrolledDateTime = $device.enrollmentDateTime }
                    elseif ($device.EnrolledDate) { $enrolledDateTime = $device.EnrolledDate }
                    elseif ($device.enrolledDate) { $enrolledDateTime = $device.enrolledDate }
                    
                    # Get OS information
                    $operatingSystem = if ($device.OperatingSystem) { $device.OperatingSystem } elseif ($device.operatingSystem) { $device.operatingSystem } else { "Unknown" }
                    $osVersion = if ($device.OSVersion) { $device.OSVersion } elseif ($device.osVersion) { $device.osVersion } else { $null }
                    
                    # Get management state
                    $managementState = if ($device.ManagementState) { $device.ManagementState } elseif ($device.managementState) { $device.managementState } else { "Unknown" }
                    $isManaged = $managementState -ne "Unmanaged"
                    
                    # Get device type
                    $deviceType = if ($device.DeviceType) { $device.DeviceType } elseif ($device.deviceType) { $device.deviceType } else { "Unknown" }
                    
                    # Get ownership
                    $ownership = if ($device.Ownership) { $device.Ownership } elseif ($device.ownership) { $device.ownership } else { "Unknown" }
                }
                
                # Try to get compliance policy evaluation details if available
                # NOTE: Per-device API calls are disabled for performance - only use data already on device object
                $compliancePolicyJSON = $null
                try {
                    # Check multiple property names for compliance policies
                    $compliancePolicies = $null
                    if ($device.CompliancePolicies) { $compliancePolicies = $device.CompliancePolicies }
                    elseif ($device.compliancePolicies) { $compliancePolicies = $device.compliancePolicies }
                    elseif ($device.DeviceCompliancePolicyStates) { $compliancePolicies = $device.DeviceCompliancePolicyStates }
                    elseif ($device.deviceCompliancePolicyStates) { $compliancePolicies = $device.deviceCompliancePolicyStates }
                    elseif ($device.PSObject.Properties['CompliancePolicies']) { $compliancePolicies = $device.PSObject.Properties['CompliancePolicies'].Value }
                    elseif ($device.PSObject.Properties['compliancePolicies']) { $compliancePolicies = $device.PSObject.Properties['compliancePolicies'].Value }
                    
                    if ($compliancePolicies) {
                        $compliancePolicyJSON = $compliancePolicies | ConvertTo-Json -Depth 10 -Compress
                    }
                    # Removed per-device API call for compliance policies - too slow for large device counts
                    # If needed, can be fetched separately via batch API call
                } catch {
                    # Silently fail - compliance policies may not be available
                }
                
                $deviceRecord = [PSCustomObject]@{
                    DeviceId = $deviceId
                    DeviceName = $deviceName
                    UserId = $userId
                    UserPrincipalName = $userPrincipalName
                    ComplianceState = $complianceState
                    IsCompliant = $isCompliant
                    LastSyncDateTime = $lastSyncDateTime
                    LastCheckInTime = $lastSyncDateTime  # Alias for clarity
                    EnrolledDateTime = $enrolledDateTime
                    OperatingSystem = $operatingSystem
                    OSVersion = $osVersion
                    ManagementState = $managementState
                    IsManaged = $isManaged
                    DeviceType = $deviceType
                    Ownership = $ownership
                    CompliancePoliciesJSON = $compliancePolicyJSON
                    # Include all original properties
                    RawDeviceData = $device
                }
                
                $allDevices += $deviceRecord
                $processedCount++
            } catch {
                $errorCount++
                Write-Warning "Error processing device record: $($_.Exception.Message)"
                if ($errorCount -le 3) {
                    Write-Warning "  Device object type: $($device.GetType().FullName)"
                    Write-Warning "  Device properties: $($device.PSObject.Properties.Name -join ', ')"
                }
            }
        }
        
        Write-Host "  Successfully processed $processedCount device(s)" -ForegroundColor Green
        if ($errorCount -gt 0) {
            Write-Warning "  Encountered errors processing $errorCount device(s)"
        }
        Write-Host "  Returning $($allDevices.Count) Intune device record(s) total" -ForegroundColor Green
        
        if ($allDevices.Count -eq 0 -and $managedDevices.Count -gt 0) {
            Write-Warning "  WARNING: Had $($managedDevices.Count) device(s) from API but processed 0 - check device processing logic"
        }
        
    } catch {
        Write-Error "Failed to retrieve Intune device records: $($_.Exception.Message)"
        Write-Host "  Exception type: $($_.Exception.GetType().FullName)" -ForegroundColor Red
        if ($_.ScriptStackTrace) {
            Write-Host "  Stack trace: $($_.ScriptStackTrace)" -ForegroundColor DarkRed
        }
        return @()
    }
    
    return $allDevices
}

Export-ModuleMember -Function Test-EntraModules,Install-EntraModules,Connect-EntraGraph,Get-EntraUsers,Get-EntraSignInLogs,Get-EntraUserDetailsAndRoles,Get-EntraUserAuditLogs,Get-EntraUserMfaStatus,Get-IntuneDeviceRecords 