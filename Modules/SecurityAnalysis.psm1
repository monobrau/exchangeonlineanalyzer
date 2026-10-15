# SecurityAnalysis.psm1
# Risky Users, Conditional Access Policies, and App Registration Analysis

# Required cmdlets and their source modules:
# - Get-MgRiskyUser: Available through Microsoft.Graph.Identity.SignIns (requires Azure AD Premium P2 license)
# - Get-MgIdentityConditionalAccessPolicy: Available through Microsoft.Graph.Identity.SignIns
# - Get-MgServicePrincipal: Available through Microsoft.Graph.Applications
# - Get-MgApplication: Available through Microsoft.Graph.Applications
# Note: Identity.Protection, Identity.ConditionalAccess, and ServicePrincipals don't exist as separate installable modules
# The cmdlets are available through Microsoft.Graph.Identity.SignIns and Microsoft.Graph.Applications which we install
$script:requiredModules = @(
    "Microsoft.Graph.Applications"  # Provides Get-MgApplication and Get-MgServicePrincipal
    # Microsoft.Graph.Identity.SignIns is installed separately and provides Get-MgIdentityConditionalAccessPolicy and Get-MgRiskyUser
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
    
    # Check if required cmdlets are available (they might be through other modules)
    $cmdletsNeeded = @{
        "Get-MgRiskyUser" = "Microsoft.Graph.Identity.SignIns"
        "Get-MgIdentityConditionalAccessPolicy" = "Microsoft.Graph.Identity.SignIns"
        "Get-MgServicePrincipal" = "Microsoft.Graph.Applications"
        "Get-MgApplication" = "Microsoft.Graph.Applications"
    }
    
    $missingCmdlets = @()
    foreach ($cmdlet in $cmdletsNeeded.Keys) {
        if (-not (Get-Command $cmdlet -ErrorAction SilentlyContinue)) {
            $missingCmdlets += "$cmdlet (from $($cmdletsNeeded[$cmdlet]))"
        }
    }
    
    # If cmdlets are missing, we still need the modules
    if ($missingCmdlets.Count -gt 0) {
        Write-Host "Note: Some cmdlets may not be available: $($missingCmdlets -join ', ')" -ForegroundColor Yellow
    }
    
    return $missing
}

function Install-SecurityAnalysisModules {
    [CmdletBinding()]
    param()
    $missing = Test-SecurityAnalysisModules
    if ($missing.Count -gt 0) {
        # Filter out modules that don't exist in the gallery
        $modulesToInstall = @()
        $nonExistentModules = @()
        foreach ($m in $missing) {
            try {
                $moduleExists = Find-Module -Name $m -Repository PSGallery -ErrorAction Stop
                if ($moduleExists) {
                    $modulesToInstall += $m
                } else {
                    $nonExistentModules += $m
                }
            } catch {
                # Module doesn't exist in gallery
                $nonExistentModules += $m
            }
        }
        
        if ($modulesToInstall.Count -gt 0) {
            Write-Host "Installing missing modules: $($modulesToInstall -join ', ')" -ForegroundColor Yellow
            foreach ($m in $modulesToInstall) {
                try {
                    Install-Module -Name $m -Scope CurrentUser -Repository PSGallery -Force -AllowClobber -ErrorAction Stop
                    Write-Host "✓ Installed $m" -ForegroundColor Green
                } catch {
                    Write-Warning "Failed to install $m`: $($_.Exception.Message)"
                }
            }
        }
        
        # Inform about non-existent modules - the cmdlets are available through other modules
        if ($nonExistentModules.Count -gt 0) {
            Write-Host "Note: These modules don't exist as separate packages, but the required cmdlets are available through:" -ForegroundColor Yellow
            foreach ($m in $nonExistentModules) {
                switch ($m) {
                    "Microsoft.Graph.Identity.Protection" {
                        Write-Host "  - Get-MgRiskyUser is available through Microsoft.Graph.Identity.SignIns (already installed)" -ForegroundColor DarkGray
                    }
                    "Microsoft.Graph.Identity.ConditionalAccess" {
                        Write-Host "  - Get-MgIdentityConditionalAccessPolicy is available through Microsoft.Graph.Identity.SignIns (already installed)" -ForegroundColor DarkGray
                    }
                    "Microsoft.Graph.ServicePrincipals" {
                        Write-Host "  - Get-MgServicePrincipal is available through Microsoft.Graph.Applications (already installed)" -ForegroundColor DarkGray
                    }
                }
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
                    $cmdletAvailable = (Get-Command Get-MgRiskyUser -ErrorAction SilentlyContinue) -ne $null
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
                    $cmdletAvailable = (Get-Command Get-MgRiskyUser -ErrorAction SilentlyContinue) -ne $null
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
            $allRiskyUsers = Get-MgRiskyUser -ErrorAction Stop
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
                        [void]$analysis.SuspiciousIndicators.Add("Policy applies to ALL users")
                        $analysis.RiskScore += 3
                    }
                    
                    # Check for guest/external user exclusions that might be suspicious
                    if ($policy.Conditions.Users -and $policy.Conditions.Users.ExcludeUsers) {
                        $excludeCount = $policy.Conditions.Users.ExcludeUsers.Count
                        if ($excludeCount -gt 5) {
                            [void]$analysis.SuspiciousIndicators.Add("Many excluded users ($excludeCount) - potential bypass")
                            $analysis.RiskScore += 2
                        }
                    }
                    
                    # Check for suspicious locations (trusted IPs that might be compromised)
                    if ($policy.Conditions.Locations -and $policy.Conditions.Locations.IncludeLocations) {
                        $includeLocations = $policy.Conditions.Locations.IncludeLocations
                        if ($includeLocations -contains "All") {
                            [void]$analysis.SuspiciousIndicators.Add("Policy applies to ALL locations")
                            $analysis.RiskScore += 2
                        }
                    }
                    
                    # Check for device conditions that might bypass MFA
                    if ($policy.Conditions.Devices -and $policy.Conditions.Devices.ExcludeDevices) {
                        [void]$analysis.SuspiciousIndicators.Add("Device exclusions may bypass security")
                        $analysis.RiskScore += 1
                    }
                }
                
                # Check for suspicious controls
                if ($policy.GrantControls) {
                    # Check if MFA is required
                    $requiresMfa = $policy.GrantControls.BuiltInControls -contains "mfa"
                    if (-not $requiresMfa) {
                        $analysis.HasSuspiciousControls = $true
                        [void]$analysis.SuspiciousIndicators.Add("Policy does NOT require MFA")
                        $analysis.RiskScore += 5
                    }
                    
                    # Check for session controls that might be too permissive
                    if ($policy.SessionControls) {
                        if ($policy.SessionControls.SignInFrequency -and $policy.SessionControls.SignInFrequency.Value -gt 24) {
                            [void]$analysis.SuspiciousIndicators.Add("Long sign-in frequency ($($policy.SessionControls.SignInFrequency.Value) hours) - reduced security")
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
    param(
        [Parameter(Mandatory=$false)]
        [switch]$IncludeOAuthConsents
    )
    
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
                    HasDelegatedConsents = $false
                    SuspiciousIndicators = [System.Collections.ArrayList]::new()
                    RiskScore = 0
                }
                
                # Get OAuth2 permission grants (user/admin consents) if requested
                $oauthGrants = @()
                if ($IncludeOAuthConsents -and $servicePrincipal) {
                    try {
                        $oauthGrants = Get-MgOauth2PermissionGrant -Filter "clientId eq '$($servicePrincipal.Id)'" -ErrorAction SilentlyContinue
                        if ($oauthGrants -and $oauthGrants.Count -gt 0) {
                            $analysis.HasDelegatedConsents = $true
                            foreach ($grant in $oauthGrants) {
                                # ConsentType: AllPrincipals (admin consent) or Principal (user consent)
                                if ($grant.ConsentType -eq 'Principal') {
                                    [void]$analysis.SuspiciousIndicators.Add("User-level OAuth consent granted (ConsentType: Principal)")
                                    $analysis.RiskScore += 2
                                }
                            }
                        }
                    } catch {
                        Write-Verbose "Could not retrieve OAuth2 grants for app $($app.DisplayName): $($_.Exception.Message)"
                    }
                }
                
                # Check required permissions
                # Use hashtable for O(1) lookups instead of O(n) array search
                $highPrivilegeScopes = @{
                    "User.ReadWrite.All" = $true
                    "Directory.ReadWrite.All" = $true
                    "Mail.ReadWrite" = $true
                    "Mail.Send" = $true
                    "Mailbox.ReadWrite" = $true
                    "Files.ReadWrite.All" = $true
                    "Sites.ReadWrite.All" = $true
                    "Calendars.ReadWrite" = $true
                    "Contacts.ReadWrite" = $true
                    "Notes.ReadWrite.All" = $true
                    "Tasks.ReadWrite" = $true
                    "Group.ReadWrite.All" = $true
                    "RoleManagement.ReadWrite.Directory" = $true
                    "Application.ReadWrite.All" = $true
                    "Policy.ReadWrite.All" = $true
                }
                
                $requiredPermissions = [System.Collections.ArrayList]::new()
                if ($app.RequiredResourceAccess) {
                    foreach ($resourceAccess in $app.RequiredResourceAccess) {
                        # Use cached service principal instead of making new calls
                        $sp = $null
                        if ($allServicePrincipals.ContainsKey($resourceAccess.ResourceAppId)) {
                            $sp = $allServicePrincipals[$resourceAccess.ResourceAppId]
                        }
                        
                        if ($sp) {
                            # Pre-build hashtable lookup for AppRoles by ID to avoid O(n) Where-Object in nested loop
                            $appRolesById = @{}
                            if ($sp.AppRoles) {
                                foreach ($role in $sp.AppRoles) {
                                    $appRolesById[$role.Id] = $role
                                }
                            }
                            
                            foreach ($permission in $resourceAccess.ResourceAccess) {
                                # O(1) lookup instead of O(n) Where-Object
                                if ($appRolesById.ContainsKey($permission.Id)) {
                                    $appRole = $appRolesById[$permission.Id]
                                    $permName = $appRole.Value
                                    [void]$requiredPermissions.Add($permName)
                                    
                                    # Check for high privilege permissions - O(1) lookup
                                    if ($highPrivilegeScopes.ContainsKey($permName)) {
                                        $analysis.HasHighPrivilegePermissions = $true
                                        [void]$analysis.SuspiciousIndicators.Add("High privilege permission: $permName")
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
                            [void]$analysis.SuspiciousIndicators.Add("Allows user consent: $($scope.Value)")
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
                    [void]$analysis.SuspiciousIndicators.Add("No verified publisher domain")
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
                    [void]$analysis.SuspiciousIndicators.Add("Uses password credentials (less secure than certificates)")
                    $analysis.RiskScore += 1
                }
                
                # Check for redirect URIs that might be suspicious
                if ($app.Web -and $app.Web.RedirectUris) {
                    foreach ($uri in $app.Web.RedirectUris) {
                        if ($uri -like "*localhost*" -or $uri -like "*127.0.0.1*" -or $uri -notlike "https://*") {
                            [void]$analysis.SuspiciousIndicators.Add("Suspicious redirect URI: $uri")
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
                    OAuthGrants = $oauthGrants
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

function Get-OAuthPermissionGrants {
    <#
    .SYNOPSIS
        Gets OAuth2 permission grants (user and admin consents) to detect suspicious app access.
    .PARAMETER DaysBack
        Number of days to look back for recent grants (default: 30)
    .PARAMETER IncludeAll
        Include all grants, not just recent ones
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$false)]
        [int]$DaysBack = 30,
        
        [Parameter(Mandatory=$false)]
        [switch]$IncludeAll
    )
    
    try {
        Import-SecurityAnalysisModules
        
        $grants = @()
        $suspiciousGrants = @()
        
        try {
            # Get all OAuth2 permission grants
            $allGrants = Get-MgOauth2PermissionGrant -All -ErrorAction Stop
            Write-Host "Found $($allGrants.Count) OAuth2 permission grants" -ForegroundColor Cyan
            
            $cutoffDate = if (-not $IncludeAll) { (Get-Date).AddDays(-$DaysBack) } else { [DateTime]::MinValue }
            
            foreach ($grant in $allGrants) {
                # Filter by date if not including all
                $grantDate = if ($grant.StartTime) { $grant.StartTime } else { [DateTime]::MinValue }
                if (-not $IncludeAll -and $grantDate -lt $cutoffDate) {
                    continue
                }
                
                # Get service principal (app) details
                $spDisplayName = "Unknown"
                $spAppId = "Unknown"
                try {
                    $sp = Get-MgServicePrincipal -ServicePrincipalId $grant.ClientId -ErrorAction SilentlyContinue
                    if ($sp) {
                        $spDisplayName = $sp.DisplayName
                        $spAppId = $sp.AppId
                    }
                } catch {
                    Write-Verbose "Could not retrieve service principal for grant client ID $($grant.ClientId)"
                }
                
                # Get resource (API) details
                $resourceDisplayName = "Unknown"
                try {
                    $resource = Get-MgServicePrincipal -ServicePrincipalId $grant.ResourceId -ErrorAction SilentlyContinue
                    if ($resource) {
                        $resourceDisplayName = $resource.DisplayName
                    }
                } catch {
                    Write-Verbose "Could not retrieve resource service principal for grant resource ID $($grant.ResourceId)"
                }
                
                # Determine if suspicious
                $isSuspicious = $false
                $suspiciousReasons = [System.Collections.ArrayList]::new()
                
                # User-level consent (not admin) is suspicious
                if ($grant.ConsentType -eq 'Principal') {
                    $isSuspicious = $true
                    [void]$suspiciousReasons.Add("User-level consent (not admin-approved)")
                }
                
                # Check for high-risk scopes
                $highRiskScopes = @('Mail.ReadWrite', 'Mail.Send', 'Files.ReadWrite.All', 'Sites.ReadWrite.All', 
                                   'User.ReadWrite.All', 'Directory.ReadWrite.All', 'Mailbox.ReadWrite')
                if ($grant.Scope) {
                    $grantedScopes = $grant.Scope -split ' '
                    foreach ($scope in $grantedScopes) {
                        if ($highRiskScopes -contains $scope.Trim()) {
                            $isSuspicious = $true
                            [void]$suspiciousReasons.Add("High-risk permission: $scope")
                        }
                    }
                }
                
                # External apps (not Microsoft)
                if ($spDisplayName -notlike "*Microsoft*" -and $resourceDisplayName -notlike "*Microsoft*") {
                    [void]$suspiciousReasons.Add("External (non-Microsoft) application")
                }
                
                $grantDetails = [PSCustomObject]@{
                    Id = $grant.Id
                    ClientId = $grant.ClientId
                    ClientDisplayName = $spDisplayName
                    ClientAppId = $spAppId
                    ResourceId = $grant.ResourceId
                    ResourceDisplayName = $resourceDisplayName
                    ConsentType = $grant.ConsentType
                    PrincipalId = $grant.PrincipalId
                    Scope = $grant.Scope
                    StartTime = $grant.StartTime
                    IsSuspicious = $isSuspicious
                    SuspiciousReasons = @($suspiciousReasons)
                }
                
                $grants += $grantDetails
                if ($isSuspicious) {
                    $suspiciousGrants += $grantDetails
                }
            }
            
            Write-Host "Found $($suspiciousGrants.Count) suspicious OAuth2 grants out of $($grants.Count) total" -ForegroundColor $(if ($suspiciousGrants.Count -gt 0) { 'Yellow' } else { 'Green' })
            
        } catch {
            if ($_.Exception.Message -like "*insufficient privileges*" -or $_.Exception.Message -like "*permission*" -or $_.Exception.Message -like "*access denied*") {
                Write-Warning "Insufficient permissions to read OAuth2 grants. Requires 'Directory.Read.All' or 'DelegatedPermissionGrant.ReadWrite.All' permission."
                return @{ AllGrants = @(); SuspiciousGrants = @() }
            } else {
                Write-Warning "Error retrieving OAuth2 grants: $($_.Exception.Message)"
                return @{ AllGrants = @(); SuspiciousGrants = @() }
            }
        }
        
        return @{
            AllGrants = $grants
            SuspiciousGrants = $suspiciousGrants
        }
        
    } catch {
        Write-Error "Failed to get OAuth2 permission grants: $($_.Exception.Message)"
        return @{ AllGrants = @(); SuspiciousGrants = @() }
    }
}

function Get-UserAuthenticationMethodChanges {
    <#
    .SYNOPSIS
        Detects suspicious authentication method changes (MFA/phone/email changes) from audit logs.
    .PARAMETER DaysBack
        Number of days to look back (default: 30)
    .PARAMETER SelectedUsers
        Filter to specific users (optional)
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$false)]
        [int]$DaysBack = 30,
        
        [Parameter(Mandatory=$false)]
        [array]$SelectedUsers = @()
    )
    
    try {
        Import-SecurityAnalysisModules
        
        $changes = @()
        $suspiciousChanges = @()
        
        try {
            $startDate = (Get-Date).AddDays(-$DaysBack)
            $startIso = $startDate.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
            
            # Look for authentication method registration/change activities
            $authMethodActivities = @(
                'User registered security info',
                'User changed default security info',
                'User deleted security info',
                'Admin updated security info',
                'Update user',
                'Change user password'
            )
            
            $allChanges = @()
            
            foreach ($activity in $authMethodActivities) {
                try {
                    $filter = "activityDateTime ge $startIso and activityDisplayName eq '$activity'"
                    $auditLogs = Get-MgAuditLogDirectoryAudit -Filter $filter -All -ErrorAction SilentlyContinue
                    
                    if ($auditLogs) {
                        foreach ($log in $auditLogs) {
                            # Filter by selected users if specified
                            if ($SelectedUsers.Count -gt 0) {
                                $targetUpn = $null
                                if ($log.TargetResources -and $log.TargetResources.Count -gt 0) {
                                    $targetUpn = $log.TargetResources[0].UserPrincipalName
                                }
                                
                                if (-not $targetUpn -or $targetUpn -notin $SelectedUsers) {
                                    continue
                                }
                            }
                            
                            $allChanges += $log
                        }
                    }
                } catch {
                    Write-Verbose "Could not retrieve audit logs for activity '$activity': $($_.Exception.Message)"
                }
            }
            
            Write-Host "Found $($allChanges.Count) authentication method change events" -ForegroundColor Cyan
            
            foreach ($log in $allChanges) {
                $isSuspicious = $false
                $suspiciousReasons = [System.Collections.ArrayList]::new()
                
                # Extract user info
                $targetUpn = if ($log.TargetResources -and $log.TargetResources.Count -gt 0) { 
                    $log.TargetResources[0].UserPrincipalName 
                } else { 
                    "Unknown" 
                }
                
                $initiatedByUpn = if ($log.InitiatedBy.User) {
                    $log.InitiatedBy.User.UserPrincipalName
                } elseif ($log.InitiatedBy.App) {
                    "App: $($log.InitiatedBy.App.DisplayName)"
                } else {
                    "Unknown"
                }
                
                # Suspicious patterns
                # 1. Changes initiated by different user (not self-service)
                if ($targetUpn -ne $initiatedByUpn -and $initiatedByUpn -ne "Unknown" -and -not $initiatedByUpn.StartsWith("App:")) {
                    $isSuspicious = $true
                    [void]$suspiciousReasons.Add("Changed by different user: $initiatedByUpn")
                }
                
                # 2. Deletion of security info
                if ($log.ActivityDisplayName -like "*deleted*") {
                    $isSuspicious = $true
                    [void]$suspiciousReasons.Add("Security info deleted")
                }
                
                # 3. Multiple changes in short time (check if there are more than 3 changes for this user in past hour)
                $recentChangesForUser = $allChanges | Where-Object {
                    $_.TargetResources -and $_.TargetResources[0].UserPrincipalName -eq $targetUpn -and
                    $_.ActivityDateTime -gt $log.ActivityDateTime.AddHours(-1)
                }
                if ($recentChangesForUser.Count -gt 3) {
                    $isSuspicious = $true
                    [void]$suspiciousReasons.Add("Multiple changes ($($recentChangesForUser.Count)) in past hour")
                }
                
                # 4. Failed result
                if ($log.Result -and $log.Result -ne 'success') {
                    [void]$suspiciousReasons.Add("Failed attempt: $($log.Result)")
                }
                
                $changeDetails = [PSCustomObject]@{
                    ActivityDateTime = $log.ActivityDateTime
                    ActivityDisplayName = $log.ActivityDisplayName
                    TargetUserPrincipalName = $targetUpn
                    InitiatedBy = $initiatedByUpn
                    Result = $log.Result
                    CorrelationId = $log.CorrelationId
                    IpAddress = if ($log.InitiatedBy.User) { $log.InitiatedBy.User.IpAddress } else { $null }
                    IsSuspicious = $isSuspicious
                    SuspiciousReasons = @($suspiciousReasons)
                }
                
                $changes += $changeDetails
                if ($isSuspicious) {
                    $suspiciousChanges += $changeDetails
                }
            }
            
            Write-Host "Found $($suspiciousChanges.Count) suspicious authentication method changes out of $($changes.Count) total" -ForegroundColor $(if ($suspiciousChanges.Count -gt 0) { 'Yellow' } else { 'Green' })
            
        } catch {
            if ($_.Exception.Message -like "*insufficient privileges*" -or $_.Exception.Message -like "*permission*" -or $_.Exception.Message -like "*access denied*") {
                Write-Warning "Insufficient permissions to read audit logs. Requires 'AuditLog.Read.All' or 'Directory.Read.All' permission."
                return @{ AllChanges = @(); SuspiciousChanges = @() }
            } else {
                Write-Warning "Error retrieving authentication method changes: $($_.Exception.Message)"
                return @{ AllChanges = @(); SuspiciousChanges = @() }
            }
        }
        
        return @{
            AllChanges = $changes
            SuspiciousChanges = $suspiciousChanges
        }
        
    } catch {
        Write-Error "Failed to get authentication method changes: $($_.Exception.Message)"
        return @{ AllChanges = @(); SuspiciousChanges = @() }
    }
}

# Export functions
Export-ModuleMember -Function Get-RiskyUsers, Get-ConditionalAccessPolicies, Get-AppRegistrations, Test-SecurityAnalysisModules, Install-SecurityAnalysisModules, Import-SecurityAnalysisModules, Get-OAuthPermissionGrants, Get-UserAuthenticationMethodChanges

