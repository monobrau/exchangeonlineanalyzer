# SecurityAnalysis.psm1
# Risky Users, Conditional Access Policies, and App Registration Analysis

$script:requiredModules = @(
    "Microsoft.Graph.Identity.Protection",
    "Microsoft.Graph.Identity.ConditionalAccess",
    "Microsoft.Graph.Applications",
    "Microsoft.Graph.ServicePrincipals"
)
$script:requiredScopes = @(
    "IdentityRiskEvent.Read.All",
    "IdentityRiskyUser.Read.All",
    "Policy.Read.All",
    "Policy.ReadWrite.ConditionalAccess",
    "Application.Read.All",
    "Directory.Read.All"
)

function Test-SecurityAnalysisModules {
    [CmdletBinding()]
    param()
    $missing = @()
    foreach ($m in $script:requiredModules) {
        if (-not (Get-Module -ListAvailable -Name $m)) { 
            $missing += $m 
        }
    }
    return $missing
}

function Install-SecurityAnalysisModules {
    [CmdletBinding()]
    param()
    $missing = Test-SecurityAnalysisModules
    if ($missing.Count -gt 0) {
        Write-Host "Installing missing modules: $($missing -join ', ')" -ForegroundColor Yellow
        foreach ($m in $missing) {
            try {
                Install-Module -Name $m -Scope CurrentUser -Repository PSGallery -Force -AllowClobber -ErrorAction Stop
                Write-Host "✓ Installed $m" -ForegroundColor Green
            } catch {
                Write-Warning "Failed to install $m`: $($_.Exception.Message)"
            }
        }
    }
}

function Import-SecurityAnalysisModules {
    [CmdletBinding()]
    param()
    foreach ($m in $script:requiredModules) {
        try {
            # Check if cmdlets are available first (might be through umbrella module)
            $cmdletAvailable = $false
            switch ($m) {
                "Microsoft.Graph.Identity.Protection" {
                    $cmdletAvailable = (Get-Command Get-MgIdentityRiskyUser -ErrorAction SilentlyContinue) -ne $null
                }
                "Microsoft.Graph.Identity.ConditionalAccess" {
                    $cmdletAvailable = (Get-Command Get-MgIdentityConditionalAccessPolicy -ErrorAction SilentlyContinue) -ne $null
                }
                "Microsoft.Graph.Applications" {
                    $cmdletAvailable = (Get-Command Get-MgApplication -ErrorAction SilentlyContinue) -ne $null
                }
                "Microsoft.Graph.ServicePrincipals" {
                    $cmdletAvailable = (Get-Command Get-MgServicePrincipal -ErrorAction SilentlyContinue) -ne $null
                }
            }
            
            # If cmdlets are available, skip module import (they're available through umbrella module)
            if ($cmdletAvailable) {
                # Cmdlets are available, no need to import specific module
                continue
            }
            
            # Check if module is available
            if (Get-Module -ListAvailable -Name $m) {
                Import-Module $m -ErrorAction Stop -Force -WarningAction SilentlyContinue
            } else {
                # Try to install automatically if not available
                Write-Host "Module $m not found. Attempting to install..." -ForegroundColor Yellow
                try {
                    Install-Module -Name $m -Scope CurrentUser -Repository PSGallery -Force -AllowClobber -ErrorAction Stop
                    Import-Module $m -ErrorAction Stop -Force -WarningAction SilentlyContinue
                    Write-Host "✓ Successfully installed and imported $m" -ForegroundColor Green
                } catch {
                    Write-Warning "Module $m not available and cmdlets not found. Some features may not work. Run Install-SecurityAnalysisModules to install required modules."
                }
            }
        } catch {
            # Check if cmdlets are available through umbrella module even if import failed
            $cmdletAvailable = $false
            switch ($m) {
                "Microsoft.Graph.Identity.Protection" {
                    $cmdletAvailable = (Get-Command Get-MgIdentityRiskyUser -ErrorAction SilentlyContinue) -ne $null
                }
                "Microsoft.Graph.Identity.ConditionalAccess" {
                    $cmdletAvailable = (Get-Command Get-MgIdentityConditionalAccessPolicy -ErrorAction SilentlyContinue) -ne $null
                }
                "Microsoft.Graph.Applications" {
                    $cmdletAvailable = (Get-Command Get-MgApplication -ErrorAction SilentlyContinue) -ne $null
                }
                "Microsoft.Graph.ServicePrincipals" {
                    $cmdletAvailable = (Get-Command Get-MgServicePrincipal -ErrorAction SilentlyContinue) -ne $null
                }
            }
            if (-not $cmdletAvailable) {
                Write-Warning "Failed to import module $m`: $($_.Exception.Message)"
            }
        }
    }
}

function Get-RiskyUsers {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$false)]
        [ValidateSet("Low", "Medium", "High", "All")]
        [string]$RiskLevel = "All"
    )
    
    try {
        Import-SecurityAnalysisModules
        
        $riskyUsers = @()
        
        # Get risky users
        try {
            $allRiskyUsers = Get-MgIdentityRiskyUser -ErrorAction Stop
            Write-Host "Found $($allRiskyUsers.Count) risky users" -ForegroundColor Cyan
            
            foreach ($user in $allRiskyUsers) {
                $riskLevelValue = $user.RiskLevel
                if ($RiskLevel -eq "All" -or $riskLevelValue -eq $RiskLevel) {
                    $userDetails = @{
                        UserId = $user.Id
                        UserPrincipalName = $user.UserPrincipalName
                        RiskLevel = $riskLevelValue
                        RiskState = $user.RiskState
                        RiskDetail = $user.RiskDetail
                        LastUpdatedDateTime = $user.RiskLastUpdatedDateTime
                    }
                    
                    # Get risk detections for this user
                    try {
                        $riskDetections = Get-MgIdentityRiskDetection -Filter "userId eq '$($user.Id)'" -ErrorAction SilentlyContinue
                        if ($riskDetections) {
                            $userDetails.RiskDetections = $riskDetections | ForEach-Object {
                                @{
                                    Activity = $_.Activity
                                    ActivityDateTime = $_.ActivityDateTime
                                    RiskType = $_.RiskType
                                    RiskLevel = $_.RiskLevel
                                    Location = $_.Location
                                    IpAddress = $_.IpAddress
                                    UserAgent = $_.UserAgent
                                }
                            }
                        }
                    } catch {
                        Write-Warning "Could not retrieve risk detections for user $($user.UserPrincipalName)"
                    }
                    
                    # Get user details
                    try {
                        $mgUser = Get-MgUser -UserId $user.Id -Property DisplayName, Mail, Department, JobTitle -ErrorAction SilentlyContinue
                        if ($mgUser) {
                            $userDetails.DisplayName = $mgUser.DisplayName
                            $userDetails.Mail = $mgUser.Mail
                            $userDetails.Department = $mgUser.Department
                            $userDetails.JobTitle = $mgUser.JobTitle
                        }
                    } catch {
                        Write-Warning "Could not retrieve user details for $($user.UserPrincipalName)"
                    }
                    
                    $riskyUsers += [PSCustomObject]$userDetails
                }
            }
        } catch {
            if ($_.Exception.Message -like "*insufficient privileges*" -or $_.Exception.Message -like "*permission*" -or $_.Exception.Message -like "*access denied*") {
                Write-Warning "Insufficient permissions to read risky users. Requires 'IdentityRiskyUser.Read.All' permission."
                return @()
            } elseif ($_.Exception.Message -like "*license*" -or $_.Exception.Message -like "*subscription*") {
                Write-Warning "Azure AD Identity Protection requires Azure AD Premium P2 license."
                return @()
            } else {
                Write-Warning "Error retrieving risky users: $($_.Exception.Message)"
                return @()
            }
        }
        
        return $riskyUsers
        
    } catch {
        Write-Error "Failed to get risky users: $($_.Exception.Message)"
        return @()
    }
}

function Get-ConditionalAccessPolicies {
    [CmdletBinding()]
    param()
    
    try {
        Import-SecurityAnalysisModules
        
        $policies = @()
        
        try {
            $caPolicies = Get-MgIdentityConditionalAccessPolicy -ErrorAction Stop
            Write-Host "Found $($caPolicies.Count) Conditional Access policies" -ForegroundColor Cyan
            
            foreach ($policy in $caPolicies) {
                # Analyze policy for potential security issues
                $analysis = @{
                    IsEnabled = $policy.State -eq "enabled"
                    HasSuspiciousConditions = $false
                    HasSuspiciousControls = $false
                    SuspiciousIndicators = @()
                    RiskScore = 0
                }
                
                # Check for suspicious conditions
                if ($policy.Conditions) {
                    # Check for overly broad user assignments
                    if ($policy.Conditions.Users -and $policy.Conditions.Users.IncludeUsers -contains "All") {
                        $analysis.HasSuspiciousConditions = $true
                        $analysis.SuspiciousIndicators += "Policy applies to ALL users"
                        $analysis.RiskScore += 3
                    }
                    
                    # Check for guest/external user exclusions that might be suspicious
                    if ($policy.Conditions.Users -and $policy.Conditions.Users.ExcludeUsers) {
                        $excludeCount = $policy.Conditions.Users.ExcludeUsers.Count
                        if ($excludeCount -gt 5) {
                            $analysis.SuspiciousIndicators += "Many excluded users ($excludeCount) - potential bypass"
                            $analysis.RiskScore += 2
                        }
                    }
                    
                    # Check for suspicious locations (trusted IPs that might be compromised)
                    if ($policy.Conditions.Locations -and $policy.Conditions.Locations.IncludeLocations) {
                        $includeLocations = $policy.Conditions.Locations.IncludeLocations
                        if ($includeLocations -contains "All") {
                            $analysis.SuspiciousIndicators += "Policy applies to ALL locations"
                            $analysis.RiskScore += 2
                        }
                    }
                    
                    # Check for device conditions that might bypass MFA
                    if ($policy.Conditions.Devices -and $policy.Conditions.Devices.ExcludeDevices) {
                        $analysis.SuspiciousIndicators += "Device exclusions may bypass security"
                        $analysis.RiskScore += 1
                    }
                }
                
                # Check for suspicious controls
                if ($policy.GrantControls) {
                    # Check if MFA is required
                    $requiresMfa = $policy.GrantControls.BuiltInControls -contains "mfa"
                    if (-not $requiresMfa) {
                        $analysis.HasSuspiciousControls = $true
                        $analysis.SuspiciousIndicators += "Policy does NOT require MFA"
                        $analysis.RiskScore += 5
                    }
                    
                    # Check for session controls that might be too permissive
                    if ($policy.SessionControls) {
                        if ($policy.SessionControls.SignInFrequency -and $policy.SessionControls.SignInFrequency.Value -gt 24) {
                            $analysis.SuspiciousIndicators += "Long sign-in frequency ($($policy.SessionControls.SignInFrequency.Value) hours) - reduced security"
                            $analysis.RiskScore += 2
                        }
                    }
                    
                    # Check for "Require compliant device" or "Require hybrid Azure AD joined device"
                    $requiresCompliantDevice = $policy.GrantControls.BuiltInControls -contains "compliantDevice"
                    $requiresHybridDevice = $policy.GrantControls.BuiltInControls -contains "domainJoinedDevice"
                    if (-not $requiresCompliantDevice -and -not $requiresHybridDevice) {
                        $analysis.SuspiciousIndicators += "No device compliance requirement"
                        $analysis.RiskScore += 1
                    }
                }
                
                # Determine overall risk level
                $riskLevel = "Low"
                if ($analysis.RiskScore -ge 7) {
                    $riskLevel = "High"
                } elseif ($analysis.RiskScore -ge 4) {
                    $riskLevel = "Medium"
                }
                
                $policyDetails = @{
                    Id = $policy.Id
                    DisplayName = $policy.DisplayName
                    State = $policy.State
                    CreatedDateTime = $policy.CreatedDateTime
                    ModifiedDateTime = $policy.ModifiedDateTime
                    Conditions = $policy.Conditions
                    GrantControls = $policy.GrantControls
                    SessionControls = $policy.SessionControls
                    Analysis = $analysis
                    RiskLevel = $riskLevel
                }
                
                $policies += [PSCustomObject]$policyDetails
            }
            
        } catch {
            if ($_.Exception.Message -like "*insufficient privileges*" -or $_.Exception.Message -like "*permission*" -or $_.Exception.Message -like "*access denied*") {
                Write-Warning "Insufficient permissions to read Conditional Access policies. Requires 'Policy.Read.All' permission."
                return @()
            } elseif ($_.Exception.Message -like "*license*" -or $_.Exception.Message -like "*subscription*") {
                Write-Warning "Conditional Access requires Azure AD Premium P1 license."
                return @()
            } else {
                Write-Warning "Error retrieving CA policies: $($_.Exception.Message)"
                return @()
            }
        }
        
        return $policies
        
    } catch {
        Write-Error "Failed to get Conditional Access policies: $($_.Exception.Message)"
        return @()
    }
}

function Get-AppRegistrations {
    [CmdletBinding()]
    param()
    
    try {
        Import-SecurityAnalysisModules
        
        $appRegistrations = @()
        
        try {
            $apps = Get-MgApplication -All -ErrorAction Stop
            Write-Host "Found $($apps.Count) app registrations" -ForegroundColor Cyan
            
            # Pre-fetch all service principals to avoid repeated calls (performance optimization)
            Write-Host "Pre-fetching service principals for permission resolution..." -ForegroundColor Cyan
            $allServicePrincipals = @{}
            try {
                $sps = Get-MgServicePrincipal -All -ErrorAction SilentlyContinue
                foreach ($sp in $sps) {
                    $allServicePrincipals[$sp.AppId] = $sp
                }
                Write-Host "Cached $($allServicePrincipals.Count) service principals" -ForegroundColor Green
            } catch {
                Write-Warning "Could not pre-fetch service principals: $($_.Exception.Message)"
            }
            
            $processedCount = 0
            foreach ($app in $apps) {
                $processedCount++
                if ($processedCount % 50 -eq 0) {
                    Write-Host "Processing app $processedCount of $($apps.Count)..." -ForegroundColor Yellow
                }
                
                # Get service principal for this app
                $servicePrincipal = $null
                if ($allServicePrincipals.ContainsKey($app.AppId)) {
                    $servicePrincipal = $allServicePrincipals[$app.AppId]
                }
                
                # Analyze app for potential security issues
                $analysis = @{
                    HasHighPrivilegePermissions = $false
                    HasSuspiciousPermissions = $false
                    HasUserConsent = $false
                    SuspiciousIndicators = @()
                    RiskScore = 0
                }
                
                # Check required permissions
                $highPrivilegeScopes = @(
                    "User.ReadWrite.All",
                    "Directory.ReadWrite.All",
                    "Mail.ReadWrite",
                    "Mail.Send",
                    "Mailbox.ReadWrite",
                    "Files.ReadWrite.All",
                    "Sites.ReadWrite.All",
                    "Calendars.ReadWrite",
                    "Contacts.ReadWrite",
                    "Notes.ReadWrite.All",
                    "Tasks.ReadWrite",
                    "Group.ReadWrite.All",
                    "RoleManagement.ReadWrite.Directory",
                    "Application.ReadWrite.All",
                    "Policy.ReadWrite.All"
                )
                
                $requiredPermissions = @()
                if ($app.RequiredResourceAccess) {
                    foreach ($resourceAccess in $app.RequiredResourceAccess) {
                        # Use cached service principal instead of making new calls
                        $sp = $null
                        if ($allServicePrincipals.ContainsKey($resourceAccess.ResourceAppId)) {
                            $sp = $allServicePrincipals[$resourceAccess.ResourceAppId]
                        }
                        
                        if ($sp) {
                            foreach ($permission in $resourceAccess.ResourceAccess) {
                                $appRole = $sp.AppRoles | Where-Object { $_.Id -eq $permission.Id }
                                if ($appRole) {
                                    $permName = $appRole.Value
                                    $requiredPermissions += $permName
                                    
                                    # Check for high privilege permissions
                                    if ($highPrivilegeScopes -contains $permName) {
                                        $analysis.HasHighPrivilegePermissions = $true
                                        $analysis.SuspiciousIndicators += "High privilege permission: $permName"
                                        $analysis.RiskScore += 3
                                    }
                                    
                                    # Check for suspicious permission combinations
                                    if ($permName -like "*Write*" -or $permName -like "*ReadWrite*") {
                                        $analysis.HasSuspiciousPermissions = $true
                                    }
                                }
                            }
                        }
                    }
                }
                
                # Check for user consent
                if ($app.Api -and $app.Api.Oauth2PermissionScopes) {
                    foreach ($scope in $app.Api.Oauth2PermissionScopes) {
                        if ($scope.UserConsentDisplayName) {
                            $analysis.HasUserConsent = $true
                            $analysis.SuspiciousIndicators += "Allows user consent: $($scope.Value)"
                            $analysis.RiskScore += 2
                        }
                    }
                }
                
                # Check publisher
                if ($app.PublisherDomain) {
                    if ($app.PublisherDomain -notlike "*.onmicrosoft.com" -and $app.PublisherDomain -notlike "*.microsoft.com") {
                        # External publisher - might be suspicious
                        $analysis.SuspiciousIndicators += "External publisher: $($app.PublisherDomain)"
                        $analysis.RiskScore += 1
                    }
                } else {
                    $analysis.SuspiciousIndicators += "No verified publisher domain"
                    $analysis.RiskScore += 2
                }
                
                # Check if app is disabled or deleted
                if ($app.DeletedDateTime) {
                    $analysis.SuspiciousIndicators += "App is deleted"
                    $analysis.RiskScore += 1
                }
                
                # Check for certificate credentials (more secure than secrets)
                $hasCertificates = $false
                if ($app.KeyCredentials) {
                    $hasCertificates = $true
                }
                if (-not $hasCertificates -and $app.PasswordCredentials) {
                    $analysis.SuspiciousIndicators += "Uses password credentials (less secure than certificates)"
                    $analysis.RiskScore += 1
                }
                
                # Check for redirect URIs that might be suspicious
                if ($app.Web -and $app.Web.RedirectUris) {
                    foreach ($uri in $app.Web.RedirectUris) {
                        if ($uri -like "*localhost*" -or $uri -like "*127.0.0.1*" -or $uri -notlike "https://*") {
                            $analysis.SuspiciousIndicators += "Suspicious redirect URI: $uri"
                            $analysis.RiskScore += 2
                        }
                    }
                }
                
                # Determine overall risk level
                $riskLevel = "Low"
                if ($analysis.RiskScore -ge 7) {
                    $riskLevel = "High"
                } elseif ($analysis.RiskScore -ge 4) {
                    $riskLevel = "Medium"
                }
                
                # Get service principal details
                $spDetails = $null
                if ($servicePrincipal) {
                    $spDetails = @{
                        Id = $servicePrincipal.Id
                        DisplayName = $servicePrincipal.DisplayName
                        ServicePrincipalType = $servicePrincipal.ServicePrincipalType
                        AppOwnerOrganizationId = $servicePrincipal.AppOwnerOrganizationId
                    }
                }
                
                $appDetails = @{
                    Id = $app.Id
                    AppId = $app.AppId
                    DisplayName = $app.DisplayName
                    PublisherDomain = $app.PublisherDomain
                    CreatedDateTime = $app.CreatedDateTime
                    RequiredPermissions = $requiredPermissions
                    Analysis = $analysis
                    RiskLevel = $riskLevel
                    ServicePrincipal = $spDetails
                    WebRedirectUris = if ($app.Web) { $app.Web.RedirectUris } else { @() }
                    HasCertificates = $hasCertificates
                    HasPasswordCredentials = if ($app.PasswordCredentials) { $true } else { $false }
                }
                
                $appRegistrations += [PSCustomObject]$appDetails
            }
            
        } catch {
            if ($_.Exception.Message -like "*insufficient privileges*" -or $_.Exception.Message -like "*permission*" -or $_.Exception.Message -like "*access denied*") {
                Write-Warning "Insufficient permissions to read app registrations. Requires 'Application.Read.All' permission."
                return @()
            } else {
                Write-Warning "Error retrieving app registrations: $($_.Exception.Message)"
                return @()
            }
        }
        
        return $appRegistrations
        
    } catch {
        Write-Error "Failed to get app registrations: $($_.Exception.Message)"
        return @()
    }
}

# Export functions
Export-ModuleMember -Function Get-RiskyUsers, Get-ConditionalAccessPolicies, Get-AppRegistrations, Test-SecurityAnalysisModules, Install-SecurityAnalysisModules, Import-SecurityAnalysisModules

