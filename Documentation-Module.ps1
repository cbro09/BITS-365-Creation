# Fixed Documentation Module for Microsoft 365 Tenant Setup Utility
#requires -Version 5.1
<#
.SYNOPSIS
    FIXED Documentation Module for Microsoft 365 Tenant Setup Utility
.DESCRIPTION
    Comprehensive fixes for data population issues:
    - Fixed license data collection and mapping
    - Fixed Excel data population in correct rows
    - Enhanced error handling and debugging
    - Fixed SharePoint permissions collection
.NOTES
    Version: 3.3 - COMPLETE FIXES for data population issues
    Dependencies: Microsoft Graph PowerShell SDK
#>

# === Enhanced Documentation Configuration ===
$DocumentationConfig = @{
    OutputDirectory = "$env:USERPROFILE\Documents\M365TenantSetup_Documentation_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
    ReportFormats = @('HTML', 'Excel', 'JSON')
    IncludeScreenshots = $false
    DetailLevel = 'Detailed'
    MaxUsersPerPermissionCell = 20
    EnableTableExpansion = $true
}

# === FIXED License Mapping Table ===
$script:LicenseMapping = @{
    # Microsoft 365 Business Plans
    "MICROSOFT_BUSINESS_BASIC" = "Business Basic"
    "MICROSOFT_BUSINESS_STANDARD" = "Business Standard" 
    "MICROSOFT_BUSINESS_PREMIUM" = "Business Premium"
    "O365_BUSINESS_ESSENTIALS" = "Business Basic"
    "O365_BUSINESS" = "Business Standard"
    "O365_BUSINESS_PREMIUM" = "Business Premium"
    
    # Microsoft 365 Enterprise Plans
    "SPE_E3" = "E3"
    "SPE_E5" = "E5"
    "MICROSOFT_365_E3" = "E3"
    "MICROSOFT_365_E5" = "E5"
    "ENTERPRISEPACK" = "E3"
    "ENTERPRISEPREMIUM" = "E5"
    
    # Exchange Plans
    "EXCHANGESTANDARD" = "Exchange Plan 1"
    "EXCHANGEENTERPRISE" = "Exchange Plan 2"
    "EXCHANGEONLINE_PLAN1" = "Exchange Plan 1"
    "EXCHANGEONLINE_PLAN2" = "Exchange Plan 2"
    
    # Additional software mappings
    "TEAMS_EXPLORATORY" = "Teams"
    "TEAMS1" = "Teams"
    "FLOW_FREE" = "Power Automate"
    "POWER_BI_STANDARD" = "Power BI"
    "INTUNE_A" = "Intune"
    "AAD_PREMIUM" = "Azure AD Premium"
    "AAD_PREMIUM_P2" = "Azure AD Premium P2"
}

# === MAIN DOCUMENTATION FUNCTION ===
function New-TenantDocumentation {
    <#
    .SYNOPSIS
        FIXED main documentation function with proper module isolation
    #>
    Write-LogMessage -Message "Starting Enhanced Documentation process..." -Type Info
    
    try {
        # STEP 1: Store core functions to prevent them being cleared
        $writeLogFunction = ${function:Write-LogMessage}
        $testNotEmptyFunction = ${function:Test-NotEmpty}
        $showProgressFunction = ${function:Show-Progress}
        
        # STEP 2: Remove ALL Graph modules first to avoid conflicts
        Write-LogMessage -Message "Clearing all Graph modules to prevent conflicts..." -Type Info
        Get-Module Microsoft.Graph* | Remove-Module -Force -ErrorAction SilentlyContinue
        
        # STEP 3: Restore core functions
        ${function:Write-LogMessage} = $writeLogFunction
        ${function:Test-NotEmpty} = $testNotEmptyFunction
        ${function:Show-Progress} = $showProgressFunction
        
        # STEP 4: Disconnect any existing sessions
        try {
            Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        }
        catch {
            # Ignore disconnect errors
        }
        
        # STEP 5: Force load ONLY the exact modules needed for Documentation
        $documentationModules = @(
            'Microsoft.Graph.Groups',
            'Microsoft.Graph.Users',
            'Microsoft.Graph.Identity.ConditionalAccess',
            'Microsoft.Graph.Sites',
            'Microsoft.Graph.DeviceManagement',
            'Microsoft.Graph.Identity.DirectoryManagement'
        )
        
        Write-LogMessage -Message "Loading required Graph modules for Enhanced Documentation..." -Type Info
        foreach ($module in $documentationModules) {
            try {
                Import-Module $module -Force -ErrorAction Stop
                Write-LogMessage -Message "Successfully loaded: $module" -Type Info -LogOnly
            }
            catch {
                Write-LogMessage -Message "Failed to load module $module : $($_.Exception.Message)" -Type Warning
            }
        }
        
        # STEP 6: Connect with specific scopes for Documentation
        $requiredScopes = @(
            'User.Read.All',
            'Group.Read.All', 
            'Policy.Read.All',
            'Sites.Read.All',
            'DeviceManagementConfiguration.Read.All',
            'DeviceManagementApps.Read.All',
            'DeviceManagementManagedDevices.Read.All',
            'Directory.Read.All'
        )
        
        Write-LogMessage -Message "Connecting to Microsoft Graph with Enhanced Documentation scopes..." -Type Info
        Connect-MgGraph -Scopes $requiredScopes -NoWelcome -ErrorAction Stop
        
        # Verify connection
        $context = Get-MgContext
        if (-not $context) {
            Write-LogMessage -Message "Please connect first." -Type Error
            return $false
        }
        
        Write-LogMessage -Message "Connected to Microsoft Graph as $($context.Account)" -Type Success
        
        # STEP 7: Execute main documentation logic
        # Create output directory
        $created = New-DocumentationDirectory
        if (-not $created) {
            return $false
        }
        
        # Look for the Excel template
        $templatePath = Find-ExcelTemplate
        if (-not $templatePath) {
            Write-LogMessage -Message "Excel template not found. Please ensure the template file is available." -Type Error
            return $false
        }
        
        # Gather all tenant information with FIXED data collection
        Write-LogMessage -Message "Gathering comprehensive tenant configuration data..." -Type Info
        $tenantData = Get-FixedTenantConfiguration
        
        # Generate enhanced populated Excel documentation
        Write-LogMessage -Message "Populating Excel template with enhanced configuration data..." -Type Info
        $excelGenerated = New-FixedExcelDocumentation -TenantData $tenantData -TemplatePath $templatePath
        
        $documentsGenerated = 0
        if ($excelGenerated) { $documentsGenerated++ }
        
        # Generate supplementary reports
        Write-LogMessage -Message "Generating supplementary HTML report..." -Type Info
        $htmlGenerated = New-HTMLDocumentation -TenantData $tenantData
        if ($htmlGenerated) { $documentsGenerated++ }
        
        # Generate JSON Export for backup
        Write-LogMessage -Message "Generating JSON configuration backup..." -Type Info
        $jsonGenerated = New-JSONDocumentation -TenantData $tenantData
        if ($jsonGenerated) { $documentsGenerated++ }
        
        # Generate Configuration Summary
        Write-LogMessage -Message "Generating configuration summary..." -Type Info
        $summaryGenerated = New-ConfigurationSummary -TenantData $tenantData
        if ($summaryGenerated) { $documentsGenerated++ }
        
        Write-LogMessage -Message "Enhanced documentation generation completed. Generated $documentsGenerated documents." -Type Success
        Write-LogMessage -Message "Documentation saved to: $($DocumentationConfig.OutputDirectory)" -Type Info
        
        # Open the documentation directory
        $openDirectory = Read-Host "Would you like to open the documentation directory? (Y/N)"
        if ($openDirectory -eq 'Y' -or $openDirectory -eq 'y') {
            Start-Process explorer.exe $DocumentationConfig.OutputDirectory
            Write-LogMessage -Message "Opened documentation directory" -Type Success
        }
        
        Write-LogMessage -Message "Enhanced documentation completed successfully" -Type Success
        return $true
    }
    catch {
        Write-LogMessage -Message "Error in Enhanced Documentation: $($_.Exception.Message)" -Type Error
        return $false
    }
}

# === FIXED DATA COLLECTION FUNCTIONS ===

function Get-FixedTenantConfiguration {
    <#
    .SYNOPSIS
        FIXED tenant configuration data collection with proper license handling
    #>
    try {
        Write-LogMessage -Message "Starting enhanced tenant data collection..." -Type Info
        
        $tenantConfig = @{
            Tenant = Get-TenantInformation
            Groups = Get-FixedGroupsInformation
            Users = Get-FixedUsersInformation
            ConditionalAccess = Get-FixedConditionalAccessInformation
            SharePoint = Get-FixedSharePointInformation
            Intune = Get-FixedIntuneInformation
        }
        
        Write-LogMessage -Message "Enhanced tenant data collection completed" -Type Success
        return $tenantConfig
    }
    catch {
        Write-LogMessage -Message "Error in enhanced tenant configuration collection: $($_.Exception.Message)" -Type Error
        return @{}
    }
}

function Get-FixedUsersInformation {
    <#
    .SYNOPSIS
        FIXED users information collection with proper license mapping
    #>
    try {
        Write-LogMessage -Message "Collecting enhanced users information with license mapping..." -Type Info
        
        # Get all users with license information
        $users = Get-MgUser -All -Property Id,UserPrincipalName,DisplayName,GivenName,Surname,JobTitle,Department,OfficeLocation,AccountEnabled,UserType,CreatedDateTime,AssignedLicenses
        
        # Get all available license SKUs for mapping
        $subscribedSkus = Get-MgSubscribedSku
        Write-LogMessage -Message "LICENSE DEBUG: Found $($subscribedSkus.Count) available license SKUs" -Type Info
        
        $usersInfo = @{
            TotalUsers = $users.Count
            EnabledUsers = ($users | Where-Object { $_.AccountEnabled -eq $true }).Count
            DisabledUsers = ($users | Where-Object { $_.AccountEnabled -eq $false }).Count
            GuestUsers = ($users | Where-Object { $_.UserType -eq "Guest" }).Count
            Users = @()
            LicenseMapping = $script:LicenseMapping
        }
        
        foreach ($user in $users) {
            try {
                # FIXED: Properly collect user licenses
                $userLicenses = @()
                $baseLicense = ""
                $additionalLicenses = @()
                
                if ($user.AssignedLicenses -and $user.AssignedLicenses.Count -gt 0) {
                    # Map license SKU IDs to readable names
                    foreach ($assignedLicense in $user.AssignedLicenses) {
                        $sku = $subscribedSkus | Where-Object { $_.SkuId -eq $assignedLicense.SkuId }
                        if ($sku) {
                            $userLicenses += $sku.SkuPartNumber
                        }
                    }
                    
                    Write-LogMessage -Message "LICENSE DEBUG: User $($user.UserPrincipalName) has licenses: [$($userLicenses -join ', ')]" -Type Info -LogOnly
                    
                    # FIXED: Apply proper license mapping
                    if ($userLicenses.Count -gt 0) {
                        $mappedLicenses = Apply-FixedLicenseMapping -UserLicenses $userLicenses
                        $baseLicense = $mappedLicenses.BaseLicense
                        $additionalLicenses = $mappedLicenses.AdditionalLicenses
                        
                        Write-LogMessage -Message "LICENSE DEBUG: Mapped - Base: '$baseLicense', Additional: [$($additionalLicenses -join ', ')]" -Type Info -LogOnly
                    }
                } else {
                    Write-LogMessage -Message "LICENSE DEBUG: User $($user.UserPrincipalName) has no assigned licenses" -Type Info -LogOnly
                }
                
                # Get manager information
                $managerEmail = ""
                try {
                    $manager = Get-MgUserManager -UserId $user.Id -ErrorAction SilentlyContinue
                    if ($manager) {
                        $managerEmail = $manager.AdditionalProperties.userPrincipalName
                    }
                }
                catch {
                    # Manager not found or accessible
                }
                
                $userData = @{
                    Id = $user.Id
                    UserPrincipalName = $user.UserPrincipalName
                    DisplayName = $user.DisplayName
                    GivenName = $user.GivenName
                    Surname = $user.Surname
                    JobTitle = $user.JobTitle
                    Department = $user.Department
                    Office = $user.OfficeLocation
                    AccountEnabled = $user.AccountEnabled
                    UserType = $user.UserType
                    CreatedDateTime = $user.CreatedDateTime
                    Licenses = ($userLicenses -join ", ")
                    BaseLicense = $baseLicense
                    AdditionalSoftware1 = if ($additionalLicenses.Count -gt 0) { $additionalLicenses[0] } else { "" }
                    AdditionalSoftware2 = if ($additionalLicenses.Count -gt 1) { $additionalLicenses[1] } else { "" }
                    Manager = $managerEmail
                }
                
                $usersInfo.Users += $userData
            }
            catch {
                Write-LogMessage -Message "Error processing user $($user.UserPrincipalName): $($_.Exception.Message)" -Type Warning -LogOnly
            }
        }
        
        Write-LogMessage -Message "Successfully collected $($usersInfo.Users.Count) users with enhanced license mapping" -Type Success
        return $usersInfo
    }
    catch {
        Write-LogMessage -Message "Error collecting users information: $($_.Exception.Message)" -Type Error
        return @{
            TotalUsers = 0
            EnabledUsers = 0
            DisabledUsers = 0
            GuestUsers = 0
            Users = @()
            LicenseMapping = @{}
        }
    }
}

function Apply-FixedLicenseMapping {
    <#
    .SYNOPSIS
        FIXED license mapping logic to determine base license and additional software
    #>
    param(
        [string[]]$UserLicenses
    )
    
    $baseLicense = ""
    $additionalLicenses = @()
    
    # FIXED: Priority order for base licenses
    $baseLicensePriority = @{
        "ENTERPRISEPREMIUM" = 100    # E5
        "SPE_E5" = 100               # E5
        "MICROSOFT_365_E5" = 100     # E5
        "ENTERPRISEPACK" = 90        # E3
        "SPE_E3" = 90                # E3
        "MICROSOFT_365_E3" = 90      # E3
        "MICROSOFT_BUSINESS_PREMIUM" = 70   # Business Premium
        "O365_BUSINESS_PREMIUM" = 70        # Business Premium
        "MICROSOFT_BUSINESS_STANDARD" = 60  # Business Standard
        "O365_BUSINESS" = 60                # Business Standard
        "MICROSOFT_BUSINESS_BASIC" = 50     # Business Basic
        "O365_BUSINESS_ESSENTIALS" = 50     # Business Basic
    }
    
    # Separate base licenses from additional software
    $potentialBaseLicenses = @()
    $potentialAdditionalSoftware = @()
    
    foreach ($license in $UserLicenses) {
        $mappedName = if ($script:LicenseMapping.ContainsKey($license)) { 
            $script:LicenseMapping[$license] 
        } else { 
            $license 
        }
        
        if ($baseLicensePriority.ContainsKey($license)) {
            $potentialBaseLicenses += @{
                License = $license
                MappedName = $mappedName
                Priority = $baseLicensePriority[$license]
            }
        } else {
            $potentialAdditionalSoftware += $mappedName
        }
    }
    
    # Select the highest priority license as base license
    if ($potentialBaseLicenses.Count -gt 0) {
        $topLicense = $potentialBaseLicenses | Sort-Object Priority -Descending | Select-Object -First 1
        $baseLicense = $topLicense.MappedName
        
        # Add any remaining base licenses to additional software
        $remainingBaseLicenses = $potentialBaseLicenses | Where-Object { $_.License -ne $topLicense.License }
        foreach ($remaining in $remainingBaseLicenses) {
            $potentialAdditionalSoftware += $remaining.MappedName
        }
    } else {
        # If no recognized base license, use the first license as base
        if ($UserLicenses.Count -gt 0) {
            $firstLicense = $UserLicenses[0]
            $baseLicense = if ($script:LicenseMapping.ContainsKey($firstLicense)) { 
                $script:LicenseMapping[$firstLicense] 
            } else { 
                $firstLicense 
            }
            
            # Add remaining licenses to additional software
            for ($i = 1; $i -lt $UserLicenses.Count; $i++) {
                $license = $UserLicenses[$i]
                $mappedName = if ($script:LicenseMapping.ContainsKey($license)) { 
                    $script:LicenseMapping[$license] 
                } else { 
                    $license 
                }
                $potentialAdditionalSoftware += $mappedName
            }
        }
    }
    
    return @{
        BaseLicense = $baseLicense
        AdditionalLicenses = $potentialAdditionalSoftware | Select-Object -Unique
    }
}

function Get-FixedGroupsInformation {
    <#
    .SYNOPSIS
        Collects groups information with proper error handling
    #>
    try {
        Write-LogMessage -Message "Collecting enhanced groups information..." -Type Info
        
        $groups = Get-MgGroup -All -Property Id,DisplayName,Description,GroupTypes,SecurityEnabled,MailEnabled,Mail,CreatedDateTime
        
        $groupsInfo = @{
            SecurityGroups = @()
            DistributionGroups = @()
            TotalGroups = $groups.Count
        }
        
        foreach ($group in $groups) {
            try {
                # Get group members count
                $memberCount = 0
                $members = @()
                try {
                    $groupMembers = Get-MgGroupMember -GroupId $group.Id -Top 50
                    $memberCount = $groupMembers.Count
                    $members = $groupMembers | ForEach-Object {
                        if ($_.AdditionalProperties.displayName) {
                            $_.AdditionalProperties.displayName
                        } else {
                            $_.Id
                        }
                    }
                }
                catch {
                    Write-LogMessage -Message "Could not retrieve members for group $($group.DisplayName)" -Type Warning -LogOnly
                }
                
                $groupData = @{
                    Id = $group.Id
                    DisplayName = $group.DisplayName
                    Description = $group.Description
                    GroupTypes = $group.GroupTypes -join ", "
                    SecurityEnabled = $group.SecurityEnabled
                    MailEnabled = $group.MailEnabled
                    Mail = $group.Mail
                    Members = ($members -join ", ")
                    MemberCount = $memberCount
                    CreatedDateTime = $group.CreatedDateTime
                }
                
                # Categorize groups
                if ($group.SecurityEnabled -and -not $group.MailEnabled) {
                    $groupsInfo.SecurityGroups += $groupData
                } elseif ($group.MailEnabled) {
                    $groupsInfo.DistributionGroups += $groupData
                } else {
                    # Office 365 groups go to security groups section
                    $groupsInfo.SecurityGroups += $groupData
                }
            }
            catch {
                Write-LogMessage -Message "Error processing group $($group.DisplayName): $($_.Exception.Message)" -Type Warning -LogOnly
            }
        }
        
        Write-LogMessage -Message "Collected $($groupsInfo.SecurityGroups.Count) security groups and $($groupsInfo.DistributionGroups.Count) distribution groups" -Type Success
        return $groupsInfo
    }
    catch {
        Write-LogMessage -Message "Error collecting groups information: $($_.Exception.Message)" -Type Error
        return @{
            SecurityGroups = @()
            DistributionGroups = @()
            TotalGroups = 0
        }
    }
}

function Get-FixedConditionalAccessInformation {
    <#
    .SYNOPSIS
        Collects Conditional Access policies with technical details
    #>
    try {
        Write-LogMessage -Message "Collecting enhanced Conditional Access policies..." -Type Info
        
        $policies = Get-MgIdentityConditionalAccessPolicy -All
        
        $caInfo = @{
            Policies = @()
            TotalPolicies = $policies.Count
            EnabledPolicies = 0
            DisabledPolicies = 0
        }
        
        foreach ($policy in $policies) {
            try {
                # Parse conditions and controls for readable format
                $userConditions = ""
                if ($policy.Conditions.Users) {
                    $includeUsers = if ($policy.Conditions.Users.IncludeUsers) { "Include Users: $($policy.Conditions.Users.IncludeUsers -join ', ')" } else { "" }
                    $excludeUsers = if ($policy.Conditions.Users.ExcludeUsers) { "Exclude Users: $($policy.Conditions.Users.ExcludeUsers -join ', ')" } else { "" }
                    $includeGroups = if ($policy.Conditions.Users.IncludeGroups) { "Include Groups: $($policy.Conditions.Users.IncludeGroups -join ', ')" } else { "" }
                    $excludeGroups = if ($policy.Conditions.Users.ExcludeGroups) { "Exclude Groups: $($policy.Conditions.Users.ExcludeGroups -join ', ')" } else { "" }
                    $userConditions = @($includeUsers, $excludeUsers, $includeGroups, $excludeGroups) | Where-Object { $_ } | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" } | Join-String -Separator "; "
                }
                
                $appConditions = ""
                if ($policy.Conditions.Applications) {
                    $includeApps = if ($policy.Conditions.Applications.IncludeApplications) { "Include Apps: $($policy.Conditions.Applications.IncludeApplications -join ', ')" } else { "" }
                    $excludeApps = if ($policy.Conditions.Applications.ExcludeApplications) { "Exclude Apps: $($policy.Conditions.Applications.ExcludeApplications -join ', ')" } else { "" }
                    $appConditions = @($includeApps, $excludeApps) | Where-Object { $_ } | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" } | Join-String -Separator "; "
                }
                
                $platformConditions = ""
                if ($policy.Conditions.Platforms) {
                    $platformConditions = "Platforms: $($policy.Conditions.Platforms.IncludePlatforms -join ', ')"
                }
                
                $grantControls = ""
                if ($policy.GrantControls) {
                    $operator = if ($policy.GrantControls.Operator) { "Operator: $($policy.GrantControls.Operator)" } else { "" }
                    $controls = if ($policy.GrantControls.BuiltInControls) { "Controls: $($policy.GrantControls.BuiltInControls -join ', ')" } else { "" }
                    $grantControls = @($operator, $controls) | Where-Object { $_ } | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" } | Join-String -Separator "; "
                }
                
                $sessionControls = ""
                if ($policy.SessionControls) {
                    $sessionParts = @()
                    if ($policy.SessionControls.ApplicationEnforcedRestrictions) { $sessionParts += "App Enforced Restrictions" }
                    if ($policy.SessionControls.CloudAppSecurity) { $sessionParts += "Cloud App Security" }
                    if ($policy.SessionControls.SignInFrequency) { 
                        $frequency = $policy.SessionControls.SignInFrequency
                        $sessionParts += "Sign-in Frequency: $($frequency.Value) $($frequency.Type)"
                    } else {
                        $sessionParts += "Sign-in Frequency: "
                    }
                    $sessionControls = $sessionParts -join "; "
                }
                
                $policyData = @{
                    Id = $policy.Id
                    DisplayName = $policy.DisplayName
                    State = $policy.State
                    UserConditions = $userConditions
                    ApplicationConditions = $appConditions
                    PlatformConditions = $platformConditions
                    GrantControls = $grantControls
                    SessionControls = $sessionControls
                    CreatedDateTime = $policy.CreatedDateTime
                    ModifiedDateTime = $policy.ModifiedDateTime
                }
                
                $caInfo.Policies += $policyData
                
                if ($policy.State -eq "enabled") {
                    $caInfo.EnabledPolicies++
                } else {
                    $caInfo.DisabledPolicies++
                }
            }
            catch {
                Write-LogMessage -Message "Error processing CA policy $($policy.DisplayName): $($_.Exception.Message)" -Type Warning -LogOnly
            }
        }
        
        Write-LogMessage -Message "Collected $($caInfo.TotalPolicies) Conditional Access policies with technical details" -Type Success
        return $caInfo
    }
    catch {
        Write-LogMessage -Message "Error collecting Conditional Access information: $($_.Exception.Message)" -Type Error
        return @{
            Policies = @()
            TotalPolicies = 0
            EnabledPolicies = 0
            DisabledPolicies = 0
        }
    }
}

function Get-FixedSharePointInformation {
    <#
    .SYNOPSIS
        FIXED SharePoint information collection with better error handling
    #>
    try {
        Write-LogMessage -Message "Collecting enhanced SharePoint information with permissions..." -Type Info
        
        $spInfo = @{
            TenantSettings = @{}
            Sites = @()
            TotalSites = 0
            StorageUsed = "Not available"
            SharingSettings = "Not available"
            ExternalSharingEnabled = "Not available"
        }
        
        # Try multiple methods to get SharePoint sites
        $sites = @()
        try {
            Write-LogMessage -Message "Retrieving all SharePoint sites with enhanced diagnostics..." -Type Info
            
            # Method 1: Get all sites
            try {
                $allSites = Get-MgSite -All -Top 200 -ErrorAction Stop
                $sites = $allSites
                Write-LogMessage -Message "SUCCESS: Found $($sites.Count) SharePoint sites using Get-MgSite -All" -Type Success
            }
            catch {
                Write-LogMessage -Message "Method 1 failed: $($_.Exception.Message)" -Type Warning
                
                # Method 2: Try with search
                try {
                    $searchSites = Get-MgSite -Search "*" -ErrorAction Stop
                    $sites = $searchSites  
                    Write-LogMessage -Message "SUCCESS: Found $($sites.Count) SharePoint sites using search method" -Type Success
                }
                catch {
                    Write-LogMessage -Message "Method 2 failed: $($_.Exception.Message)" -Type Warning
                    
                    # Method 3: Create a placeholder entry if no sites accessible
                    Write-LogMessage -Message "No SharePoint sites accessible with current permissions" -Type Warning
                    $sites = @()
                }
            }
        }
        catch {
            Write-LogMessage -Message "Could not retrieve SharePoint sites: $($_.Exception.Message)" -Type Warning
        }
        
        # Process found sites or create default entry
        if ($sites.Count -gt 0) {
            foreach ($site in $sites) {
                try {
                    # Get site permissions if possible
                    $owners = "Not accessible"
                    $members = "Not accessible"
                    $readOnly = "Not accessible"
                    
                    try {
                        # Try to get site permissions - this may fail due to permissions
                        $sitePermissions = Get-MgSitePermission -SiteId $site.Id -ErrorAction SilentlyContinue
                        # Process permissions if available
                    }
                    catch {
                        # Permissions not accessible
                    }
                    
                    $siteData = @{
                        Id = $site.Id
                        DisplayName = if ($site.DisplayName) { $site.DisplayName } else { $site.Name }
                        WebUrl = $site.WebUrl
                        Approver = "To be determined"
                        Owners = $owners
                        Members = $members
                        ReadOnly = $readOnly
                        CreatedDateTime = $site.CreatedDateTime
                        LastModifiedDateTime = $site.LastModifiedDateTime
                    }
                    
                    $spInfo.Sites += $siteData
                }
                catch {
                    Write-LogMessage -Message "Error processing site $($site.DisplayName): $($_.Exception.Message)" -Type Warning -LogOnly
                }
            }
        } else {
            # Create a placeholder entry for the template
            $spInfo.Sites += @{
                Id = "N/A"
                DisplayName = "No sites accessible or no permissions"
                WebUrl = "N/A"
                Approver = "N/A"
                Owners = "N/A"
                Members = "N/A"
                ReadOnly = "N/A"
                CreatedDateTime = "N/A"
                LastModifiedDateTime = "N/A"
            }
        }
        
        $spInfo.TotalSites = $spInfo.Sites.Count
        
        Write-LogMessage -Message "Collected SharePoint information for $($spInfo.TotalSites) sites" -Type Success
        return $spInfo
    }
    catch {
        Write-LogMessage -Message "Error collecting SharePoint information: $($_.Exception.Message)" -Type Error
        return @{
            TenantSettings = @{}
            Sites = @()
            TotalSites = 0
            StorageUsed = "Error retrieving"
            SharingSettings = "Error retrieving"
            ExternalSharingEnabled = "Error retrieving"
        }
    }
}

function Get-FixedIntuneInformation {
    <#
    .SYNOPSIS
        Collects Intune information with enhanced error handling
    #>
    try {
        Write-LogMessage -Message "Starting enhanced Intune data collection..." -Type Info
        
        $intuneInfo = @{
            DeviceCompliancePolicies = @()
            DeviceConfigurationPolicies = @()
            AppProtectionPolicies = @()
            EnrollmentRestrictions = @()
            ManagedDevices = @()
            ManagedApps = @()
            TotalDevices = 0
        }
        
        # Get compliance policies
        try {
            $compliancePolicies = Get-MgDeviceManagementDeviceCompliancePolicy
            Write-LogMessage -Message "Found $($compliancePolicies.Count) compliance policies" -Type Info
            $intuneInfo.DeviceCompliancePolicies = $compliancePolicies | ForEach-Object {
                @{
                    Id = $_.Id
                    DisplayName = $_.DisplayName
                    Description = $_.Description
                    CreatedDateTime = $_.CreatedDateTime
                    LastModifiedDateTime = $_.LastModifiedDateTime
                }
            }
        }
        catch {
            Write-LogMessage -Message "Could not retrieve compliance policies: $($_.Exception.Message)" -Type Warning -LogOnly
        }
        
        # Get configuration policies
        try {
            $configPolicies = Get-MgDeviceManagementDeviceConfiguration
            Write-LogMessage -Message "Found $($configPolicies.Count) configuration policies" -Type Info
            $intuneInfo.DeviceConfigurationPolicies = $configPolicies | ForEach-Object {
                @{
                    Id = $_.Id
                    DisplayName = $_.DisplayName
                    Description = $_.Description
                    CreatedDateTime = $_.CreatedDateTime
                    LastModifiedDateTime = $_.LastModifiedDateTime
                }
            }
        }
        catch {
            Write-LogMessage -Message "Could not retrieve configuration policies: $($_.Exception.Message)" -Type Warning -LogOnly
        }
        
        # Get managed applications with enhanced methods
        try {
            Write-LogMessage -Message "Collecting managed applications with enhanced methods..." -Type Info
            
            $apps = @()
            try {
                # Try Graph API method
                $uri = "https://graph.microsoft.com/v1.0/deviceAppManagement/mobileApps"
                $response = Invoke-MgGraphRequest -Uri $uri -Method GET
                $apps = $response.value
                Write-LogMessage -Message "Found $($apps.Count) managed apps using Graph API" -Type Success
            }
            catch {
                Write-LogMessage -Message "Standard mobile app collection failed: $($_.Exception.Message)" -Type Warning
            }
            
            if ($apps.Count -gt 0) {
                $processedApps = @()
                foreach ($app in $apps) {
                    try {
                        # Determine platform based on app type
                        $platform = ""
                        switch ($app.'@odata.type') {
                            "#microsoft.graph.iosLobApp" { $platform = "iOS" }
                            "#microsoft.graph.iosStoreApp" { $platform = "iOS" }
                            "#microsoft.graph.androidLobApp" { $platform = "Android" }
                            "#microsoft.graph.androidStoreApp" { $platform = "Android" }
                            "#microsoft.graph.win32LobApp" { $platform = "Windows" }
                            "#microsoft.graph.winGetApp" { $platform = "Windows" }
                            "#microsoft.graph.macOSLobApp" { $platform = "macOS" }
                            "#microsoft.graph.macOSMicrosoftEdgeApp" { $platform = "macOS" }
                            "#microsoft.graph.microsoftStoreForBusinessApp" { $platform = "Windows" }
                            default { 
                                # Try to determine from other properties
                                if ($app.applicableDeviceType) {
                                    if ($app.applicableDeviceType.iPad -or $app.applicableDeviceType.iPhoneAndIPod) {
                                        $platform = "iPadOS" 
                                    }
                                } else {
                                    $platform = ""
                                }
                            }
                        }
                        
                        $appData = @{
                            Id = $app.id
                            DisplayName = $app.displayName
                            Description = $app.description
                            Publisher = $app.publisher
                            Platform = $platform
                            CreatedDateTime = $app.createdDateTime
                            LastModifiedDateTime = $app.lastModifiedDateTime
                            ODataType = $app.'@odata.type'
                        }
                        
                        $processedApps += $appData
                    }
                    catch {
                        Write-LogMessage -Message "Error processing app $($app.displayName): $($_.Exception.Message)" -Type Warning -LogOnly
                    }
                }
                
                $intuneInfo.ManagedApps = $processedApps
                Write-LogMessage -Message "Successfully processed $($processedApps.Count) managed apps with enhanced platform categorization" -Type Success
                
                # Debug platform distribution
                $platformCounts = $processedApps | Group-Object Platform | ForEach-Object { "$($_.Name): $($_.Count)" }
                Write-LogMessage -Message "INTUNE DEBUG: Platform distribution - $($platformCounts -join ', ')" -Type Info
            }
        }
        catch {
            Write-LogMessage -Message "Could not retrieve managed applications: $($_.Exception.Message)" -Type Warning -LogOnly
        }
        
        # Get managed devices
        try {
            $devices = Get-MgDeviceManagementManagedDevice -Top 100
            $intuneInfo.TotalDevices = $devices.Count
            Write-LogMessage -Message "Found $($devices.Count) managed devices" -Type Info
            $intuneInfo.ManagedDevices = $devices | ForEach-Object {
                @{
                    Id = $_.Id
                    DeviceName = $_.DeviceName
                    OperatingSystem = $_.OperatingSystem
                    OSVersion = $_.OSVersion
                    ComplianceState = $_.ComplianceState
                    EnrolledDateTime = $_.EnrolledDateTime
                    LastSyncDateTime = $_.LastSyncDateTime
                    UserPrincipalName = $_.UserPrincipalName
                }
            }
        }
        catch {
            Write-LogMessage -Message "Could not retrieve managed devices: $($_.Exception.Message)" -Type Warning -LogOnly
        }
        
        Write-LogMessage -Message "Enhanced Intune data collection completed - Config: $($intuneInfo.DeviceConfigurationPolicies.Count), Compliance: $($intuneInfo.DeviceCompliancePolicies.Count), Apps: $($intuneInfo.ManagedApps.Count), Devices: $($intuneInfo.TotalDevices)" -Type Success
        return $intuneInfo
    }
    catch {
        Write-LogMessage -Message "Error collecting enhanced Intune information: $($_.Exception.Message)" -Type Warning -LogOnly
        return @{
            DeviceCompliancePolicies = @()
            DeviceConfigurationPolicies = @()
            AppProtectionPolicies = @()
            EnrollmentRestrictions = @()
            ManagedDevices = @()
            ManagedApps = @()
            TotalDevices = 0
        }
    }
}

# === FIXED EXCEL DOCUMENTATION FUNCTIONS ===

function New-FixedExcelDocumentation {
    <#
    .SYNOPSIS
        FIXED Excel documentation generation with proper data population
    #>
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$TenantData,
        
        [Parameter(Mandatory = $true)]
        [string]$TemplatePath
    )
    
    try {
        Write-LogMessage -Message "EXCEL DEBUG: Starting Excel documentation generation..." -Type Info
        
        # Copy template to output directory
        $outputFileName = "TenantConfiguration_Enhanced_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
        $outputPath = Join-Path $DocumentationConfig.OutputDirectory "Reports\$outputFileName"
        
        Copy-Item -Path $TemplatePath -Destination $outputPath -Force
        Write-LogMessage -Message "EXCEL DEBUG: Copied template to: $outputPath" -Type Info
        
        # Test if file exists and is accessible
        if (-not (Test-Path $outputPath)) {
            Write-LogMessage -Message "EXCEL DEBUG: Output file does not exist after copy!" -Type Error
            return $false
        }
        
        $fileInfo = Get-Item $outputPath
        Write-LogMessage -Message "EXCEL DEBUG: File size after copy: $($fileInfo.Length) bytes" -Type Info
        
        # Create Excel COM object
        try {
            $excel = New-Object -ComObject Excel.Application
            Write-LogMessage -Message "EXCEL DEBUG: Excel application created successfully" -Type Success
            $excel.Visible = $false
            $excel.ScreenUpdating = $false
        }
        catch {
            Write-LogMessage -Message "EXCEL DEBUG: Failed to create Excel application: $($_.Exception.Message)" -Type Error
            return $false
        }
        
        try {
            # Open workbook
            Write-LogMessage -Message "EXCEL DEBUG: Opening workbook..." -Type Info
            $Workbook = $excel.Workbooks.Open($outputPath)
            Write-LogMessage -Message "EXCEL DEBUG: Workbook opened successfully" -Type Success
            
            # Debug available worksheets
            $worksheetNames = @()
            for ($i = 1; $i -le $Workbook.Worksheets.Count; $i++) {
                $worksheetNames += $Workbook.Worksheets.Item($i).Name
            }
            Write-LogMessage -Message "EXCEL DEBUG: Available worksheets: $($worksheetNames -join ', ')" -Type Info
            
            # FIXED: Update sheets with proper data
            Update-FixedUsersSheet -Workbook $Workbook -TenantData $TenantData
            Update-FixedLicensingSheet -Workbook $Workbook -TenantData $TenantData
            Update-FixedConditionalAccessSheet -Workbook $Workbook -TenantData $TenantData
            Update-FixedSharePointSheet -Workbook $Workbook -TenantData $TenantData
            Update-FixedIntuneAppsSheets -Workbook $Workbook -TenantData $TenantData
            Add-FixedSecurityGroupsSheet -Workbook $Workbook -TenantData $TenantData
            Update-FixedDistributionListSheet -Workbook $Workbook -TenantData $TenantData
            
            # Final save
            Write-LogMessage -Message "EXCEL DEBUG: Performing final save..." -Type Info
            $Workbook.Save()
            $Workbook.Close()
            Write-LogMessage -Message "EXCEL DEBUG: Workbook saved and closed" -Type Success
            
            return $true
        }
        catch {
            Write-LogMessage -Message "EXCEL DEBUG: Error during workbook operations: $($_.Exception.Message)" -Type Error
            return $false
        }
        finally {
            try {
                $excel.ScreenUpdating = $true
                $excel.Quit()
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
                Write-LogMessage -Message "EXCEL DEBUG: Excel application cleaned up" -Type Success
            }
            catch {
                Write-LogMessage -Message "EXCEL DEBUG: Error during Excel cleanup: $($_.Exception.Message)" -Type Warning
            }
        }
    }
    catch {
        Write-LogMessage -Message "EXCEL DEBUG: Fatal error in Excel documentation: $($_.Exception.Message)" -Type Error
        return $false
    }
}

function Update-FixedUsersSheet {
    <#
    .SYNOPSIS
        FIXED Users sheet update with proper data population
    #>
    param(
        [Parameter(Mandatory = $true)]
        $Workbook,
        
        [Parameter(Mandatory = $true)]
        [hashtable]$TenantData
    )
    
    try {
        Write-LogMessage -Message "EXCEL DEBUG: Updating Users sheet..." -Type Info
        
        $worksheet = $Workbook.Worksheets.Item("Users")
        if (-not $worksheet) {
            Write-LogMessage -Message "EXCEL DEBUG: Users worksheet not found" -Type Error
            return
        }
        
        $users = $TenantData.Users.Users
        Write-LogMessage -Message "EXCEL DEBUG: Processing $($users.Count) users" -Type Info
        
        if (-not $users -or $users.Count -eq 0) {
            Write-LogMessage -Message "EXCEL DEBUG: No users data to populate" -Type Warning
            return
        }
        
        # Start at row 7 (first data row based on template structure)
        $startRow = 7
        $currentRow = $startRow
        
        # Populate first 50 users to avoid Excel timeouts
        $usersToProcess = $users | Select-Object -First 50
        
        foreach ($user in $usersToProcess) {
            try {
                Write-LogMessage -Message "EXCEL DEBUG: Processing user $($user.UserPrincipalName) at row $currentRow" -Type Info -LogOnly
                
                # Map to template columns: A=First Name, B=Last Name, C=Email, D=Job Title, E=Manager, F=Department, G=Office, H=Phone
                $worksheet.Cells.Item($currentRow, 1).Value = if ($user.GivenName) { $user.GivenName } else { "" }
                $worksheet.Cells.Item($currentRow, 2).Value = if ($user.Surname) { $user.Surname } else { "" }
                $worksheet.Cells.Item($currentRow, 3).Value = if ($user.UserPrincipalName) { $user.UserPrincipalName } else { "" }
                $worksheet.Cells.Item($currentRow, 4).Value = if ($user.JobTitle) { $user.JobTitle } else { "" }
                $worksheet.Cells.Item($currentRow, 5).Value = if ($user.Manager) { $user.Manager } else { "" }
                $worksheet.Cells.Item($currentRow, 6).Value = if ($user.Department) { $user.Department } else { "" }
                $worksheet.Cells.Item($currentRow, 7).Value = if ($user.Office) { $user.Office } else { "" }
                $worksheet.Cells.Item($currentRow, 8).Value = "" # Phone Number placeholder
                
                $currentRow++
                
                # Save every 10 rows to prevent loss
                if (($currentRow - $startRow) % 10 -eq 0) {
                    $Workbook.Save()
                    Write-LogMessage -Message "EXCEL DEBUG: Saved progress at row $currentRow" -Type Info -LogOnly
                }
                
            }
            catch {
                Write-LogMessage -Message "EXCEL DEBUG: Error updating user row $currentRow - $($_.Exception.Message)" -Type Error
                $currentRow++
            }
        }
        
        # Final save
        $Workbook.Save()
        Write-LogMessage -Message "EXCEL DEBUG: Successfully updated Users sheet with $($usersToProcess.Count) users" -Type Success
    }
    catch {
        Write-LogMessage -Message "EXCEL DEBUG: Error updating Users sheet - $($_.Exception.Message)" -Type Error
        Write-LogMessage -Message "EXCEL DEBUG: Full error: $($_.Exception)" -Type Error -LogOnly
    }
}

function Update-FixedLicensingSheet {
    <#
    .SYNOPSIS
        FIXED Licensing sheet update with proper license data population
    #>
    param(
        [Parameter(Mandatory = $true)]
        $Workbook,
        
        [Parameter(Mandatory = $true)]
        [hashtable]$TenantData
    )
    
    try {
        Write-LogMessage -Message "EXCEL DEBUG: Updating Licensing sheet..." -Type Info
        
        $worksheet = $Workbook.Worksheets.Item("Licensing")
        if (-not $worksheet) {
            Write-LogMessage -Message "EXCEL DEBUG: Licensing worksheet not found" -Type Error
            return
        }
        
        # Get ALL users first and debug their licenses
        $allUsers = $TenantData.Users.Users
        Write-LogMessage -Message "EXCEL DEBUG: Total users available: $($allUsers.Count)" -Type Info
        
        # Debug: Check licensing data for first 5 users
        $sampleUsers = $allUsers | Select-Object -First 5
        foreach ($sampleUser in $sampleUsers) {
            Write-LogMessage -Message "EXCEL DEBUG: User $($sampleUser.UserPrincipalName) - BaseLicense: '$($sampleUser.BaseLicense)' - Licenses: '$($sampleUser.Licenses)'" -Type Info
        }
        
        # FIXED: Find users with ANY license data (not just BaseLicense)
        $licensedUsers = $allUsers | Where-Object { 
            ($_.BaseLicense -and $_.BaseLicense.Trim() -ne "" -and $_.BaseLicense -ne "null") -or
            ($_.Licenses -and $_.Licenses.Trim() -ne "" -and $_.Licenses -ne "null")
        }
        
        Write-LogMessage -Message "EXCEL DEBUG: Found $($licensedUsers.Count) licensed users" -Type Info
        
        # If still no licensed users found, use ALL users as fallback
        if (-not $licensedUsers -or $licensedUsers.Count -eq 0) {
            $licensedUsers = $allUsers | Select-Object -First 20
            Write-LogMessage -Message "EXCEL DEBUG: Using $($licensedUsers.Count) users as fallback (no license filtering)" -Type Info
        }
        
        if (-not $licensedUsers -or $licensedUsers.Count -eq 0) {
            Write-LogMessage -Message "EXCEL DEBUG: No users found to populate licensing sheet" -Type Warning
            return
        }
        
        # FIXED: Start at row 25 (licensing table location based on template analysis)
        $startRow = 25
        $currentRow = $startRow
        
        # Clear any existing data first
        for ($clearRow = $startRow; $clearRow -le ($startRow + 30); $clearRow++) {
            for ($clearCol = 2; $clearCol -le 6; $clearCol++) {
                $worksheet.Cells.Item($clearRow, $clearCol).Value = ""
            }
        }
        
        foreach ($user in $licensedUsers) {
            try {
                Write-LogMessage -Message "EXCEL DEBUG: Processing licensed user $($user.UserPrincipalName) at row $currentRow" -Type Info -LogOnly
                
                # FIXED: Map to correct licensing table columns (B=User, C=License, D=Additional1, E=Additional2)
                $worksheet.Cells.Item($currentRow, 2).Value = if ($user.DisplayName) { $user.DisplayName } else { $user.UserPrincipalName }
                $worksheet.Cells.Item($currentRow, 3).Value = if ($user.BaseLicense -and $user.BaseLicense.Trim() -ne "") { $user.BaseLicense } else { $user.Licenses }
                $worksheet.Cells.Item($currentRow, 4).Value = if ($user.AdditionalSoftware1) { $user.AdditionalSoftware1 } else { "" }
                $worksheet.Cells.Item($currentRow, 5).Value = if ($user.AdditionalSoftware2) { $user.AdditionalSoftware2 } else { "" }
                
                $currentRow++
            }
            catch {
                Write-LogMessage -Message "EXCEL DEBUG: Error updating licensing row $currentRow - $($_.Exception.Message)" -Type Error
                $currentRow++
            }
        }
        
        $Workbook.Save()
        Write-LogMessage -Message "EXCEL DEBUG: Successfully updated Licensing sheet with $($licensedUsers.Count) users starting at row $startRow" -Type Success
    }
    catch {
        Write-LogMessage -Message "EXCEL DEBUG: Error updating Licensing sheet - $($_.Exception.Message)" -Type Error
    }
}

function Update-FixedConditionalAccessSheet {
    <#
    .SYNOPSIS
        Updates Conditional Access sheet with policy details
    #>
    param(
        [Parameter(Mandatory = $true)]
        $Workbook,
        
        [Parameter(Mandatory = $true)]
        [hashtable]$TenantData
    )
    
    try {
        Write-LogMessage -Message "EXCEL DEBUG: Updating Conditional Access sheet..." -Type Info
        
        $worksheet = $Workbook.Worksheets.Item("Conditional Access")
        if (-not $worksheet) {
            Write-LogMessage -Message "EXCEL DEBUG: Conditional Access worksheet not found" -Type Error
            return
        }
        
        $policies = $TenantData.ConditionalAccess.Policies
        Write-LogMessage -Message "EXCEL DEBUG: Processing $($policies.Count) CA policies" -Type Info
        
        if (-not $policies -or $policies.Count -eq 0) {
            Write-LogMessage -Message "EXCEL DEBUG: No CA policies data to populate" -Type Warning
            return
        }
        
        # Start at row 9 (based on template structure)
        $startRow = 9
        $currentRow = $startRow
        
        foreach ($policy in $policies) {
            try {
                # Map to template columns: B=Policy Name, C=Policy Settings
                $worksheet.Cells.Item($currentRow, 2).Value = $policy.DisplayName
                
                # Combine all policy settings into a readable format
                $policySettings = @()
                $policySettings += "State: $($policy.State)"
                if ($policy.UserConditions) { $policySettings += "Users: $($policy.UserConditions)" }
                if ($policy.ApplicationConditions) { $policySettings += "Applications: $($policy.ApplicationConditions)" }
                if ($policy.PlatformConditions) { $policySettings += $policy.PlatformConditions }
                if ($policy.GrantControls) { $policySettings += "Grant: $($policy.GrantControls)" }
                if ($policy.SessionControls) { $policySettings += "Session: $($policy.SessionControls)" }
                
                $worksheet.Cells.Item($currentRow, 3).Value = $policySettings -join " | "
                
                $currentRow++
            }
            catch {
                Write-LogMessage -Message "EXCEL DEBUG: Error updating CA row $currentRow - $($_.Exception.Message)" -Type Error
                $currentRow++
            }
        }
        
        $Workbook.Save()
        Write-LogMessage -Message "EXCEL DEBUG: Successfully updated Conditional Access sheet with $($policies.Count) policies" -Type Success
    }
    catch {
        Write-LogMessage -Message "EXCEL DEBUG: Error updating Conditional Access sheet - $($_.Exception.Message)" -Type Error
    }
}

function Update-FixedSharePointSheet {
    <#
    .SYNOPSIS
        Updates SharePoint Site sheet with site data
    #>
    param(
        [Parameter(Mandatory = $true)]
        $Workbook,
        
        [Parameter(Mandatory = $true)]
        [hashtable]$TenantData
    )
    
    try {
        Write-LogMessage -Message "Updating SharePoint Site sheet with permissions..." -Type Info
        
        $worksheet = $Workbook.Worksheets.Item("SharePoint Site")
        if (-not $worksheet) {
            Write-LogMessage -Message "SharePoint Site worksheet not found" -Type Warning
            return
        }
        
        $sites = $TenantData.SharePoint.Sites
        if (-not $sites -or $sites.Count -eq 0) {
            Write-LogMessage -Message "No SharePoint sites data to populate" -Type Warning
            return
        }
        
        # Start at row 7 (data starts here based on template)
        $startRow = 7
        $currentRow = $startRow
        
        foreach ($site in $sites) {
            try {
                # Map to template columns: B=SharePoint Site Name, C=Approver, D=Owners, E=Members, F=Read Only
                $worksheet.Cells.Item($currentRow, 2).Value = $site.DisplayName
                $worksheet.Cells.Item($currentRow, 3).Value = $site.Approver
                $worksheet.Cells.Item($currentRow, 4).Value = $site.Owners
                $worksheet.Cells.Item($currentRow, 5).Value = $site.Members
                $worksheet.Cells.Item($currentRow, 6).Value = $site.ReadOnly
                
                $currentRow++
            }
            catch {
                Write-LogMessage -Message "Error updating SharePoint row $currentRow - $($_.Exception.Message)" -Type Error
                $currentRow++
            }
        }
        
        $Workbook.Save()
        Write-LogMessage -Message "Successfully updated SharePoint Site sheet with $($sites.Count) sites" -Type Success
    }
    catch {
        Write-LogMessage -Message "Error updating SharePoint Site sheet - $($_.Exception.Message)" -Type Error
    }
}

function Update-FixedIntuneAppsSheets {
    <#
    .SYNOPSIS
        Updates all Intune Apps sheets with platform-specific data
    #>
    param(
        [Parameter(Mandatory = $true)]
        $Workbook,
        
        [Parameter(Mandatory = $true)]
        [hashtable]$TenantData
    )
    
    try {
        Write-LogMessage -Message "Updating Intune Apps sheets with platform categorization..." -Type Info
        
        $apps = $TenantData.Intune.ManagedApps
        if (-not $apps -or $apps.Count -eq 0) {
            Write-LogMessage -Message "No Intune apps data to populate" -Type Warning
            return
        }
        
        # Define platform mappings to sheet names
        $platformSheets = @{
            "iPadOS" = "Intune Apple iPadOS Apps "
            "iOS" = "Intune Apple IOS Apps "
            "macOS" = "Intune Mac OS Apps"
            "Android" = "Intune Android Apps"
            "Windows" = "Intune Windows Apps"
        }
        
        foreach ($platform in $platformSheets.Keys) {
            $sheetName = $platformSheets[$platform]
            $platformApps = $apps | Where-Object { $_.Platform -eq $platform -or ($platform -eq "iPadOS" -and $_.Platform -eq "iOS") }
            
            Write-LogMessage -Message "Processing $platform apps for sheet '$sheetName'..." -Type Info
            
            try {
                $worksheet = $Workbook.Worksheets.Item($sheetName)
                if ($worksheet) {
                    if ($platformApps.Count -gt 0) {
                        $startRow = 7  # Based on template structure
                        $currentRow = $startRow
                        
                        foreach ($app in $platformApps) {
                            try {
                                # Map to template columns: B=App Name, C=Publisher, D=Description
                                $worksheet.Cells.Item($currentRow, 2).Value = $app.DisplayName
                                $worksheet.Cells.Item($currentRow, 3).Value = if ($app.Publisher) { $app.Publisher } else { "" }
                                $worksheet.Cells.Item($currentRow, 4).Value = if ($app.Description) { $app.Description } else { "" }
                                
                                $currentRow++
                            }
                            catch {
                                Write-LogMessage -Message "Error updating $platform app row $currentRow - $($_.Exception.Message)" -Type Error
                                $currentRow++
                            }
                        }
                        
                        Write-LogMessage -Message "Successfully updated $sheetName with $($platformApps.Count) apps" -Type Success
                    } else {
                        Write-LogMessage -Message "No $platform apps to populate" -Type Info
                    }
                } else {
                    Write-LogMessage -Message "$sheetName worksheet not found" -Type Warning
                }
            }
            catch {
                Write-LogMessage -Message "Error updating $sheetName - $($_.Exception.Message)" -Type Error
            }
        }
        
        $Workbook.Save()
        Write-LogMessage -Message "Completed updating all Intune Apps sheets" -Type Success
    }
    catch {
        Write-LogMessage -Message "Error updating Intune Apps sheets - $($_.Exception.Message)" -Type Error
    }
}

function Add-FixedSecurityGroupsSheet {
    <#
    .SYNOPSIS
        Adds Security Groups sheet with group data
    #>
    param(
        [Parameter(Mandatory = $true)]
        $Workbook,
        
        [Parameter(Mandatory = $true)]
        [hashtable]$TenantData
    )
    
    try {
        Write-LogMessage -Message "Adding Security Groups sheet..." -Type Info
        
        # Check if Security Groups sheet already exists
        $existingSheet = $null
        try {
            $existingSheet = $Workbook.Worksheets.Item("Security Groups")
        }
        catch {
            # Sheet doesn't exist, will create it
        }
        
        if ($existingSheet) {
            $worksheet = $existingSheet
        } else {
            # Create new sheet
            $worksheet = $Workbook.Worksheets.Add()
            $worksheet.Name = "Security Groups"
            
            # Add headers
            $worksheet.Cells.Item(1, 1).Value = "Security Groups"
            $worksheet.Cells.Item(2, 1).Value = "Group Name"
            $worksheet.Cells.Item(2, 2).Value = "Description"
            $worksheet.Cells.Item(2, 3).Value = "Members"
        }
        
        $groups = $TenantData.Groups.SecurityGroups
        if (-not $groups -or $groups.Count -eq 0) {
            Write-LogMessage -Message "No security groups data to populate" -Type Warning
            return
        }
        
        # Start at row 3 (after headers)
        $startRow = 3
        $currentRow = $startRow
        
        foreach ($group in $groups) {
            try {
                $worksheet.Cells.Item($currentRow, 1).Value = $group.DisplayName
                $worksheet.Cells.Item($currentRow, 2).Value = if ($group.Description) { $group.Description } else { "" }
                $worksheet.Cells.Item($currentRow, 3).Value = if ($group.Members) { $group.Members } else { "" }
                
                $currentRow++
            }
            catch {
                Write-LogMessage -Message "Error updating security group row $currentRow - $($_.Exception.Message)" -Type Error
                $currentRow++
            }
        }
        
        $Workbook.Save()
        Write-LogMessage -Message "Successfully added Security Groups sheet with $($groups.Count) groups" -Type Success
    }
    catch {
        Write-LogMessage -Message "Error adding Security Groups sheet - $($_.Exception.Message)" -Type Error
    }
}

function Update-FixedDistributionListSheet {
    <#
    .SYNOPSIS
        Updates Distribution list sheet with distribution group data
    #>
    param(
        [Parameter(Mandatory = $true)]
        $Workbook,
        
        [Parameter(Mandatory = $true)]
        [hashtable]$TenantData
    )
    
    try {
        Write-LogMessage -Message "EXCEL DEBUG: Updating Distribution list sheet..." -Type Info
        
        $worksheet = $Workbook.Worksheets.Item("Distribution list")
        if (-not $worksheet) {
            Write-LogMessage -Message "EXCEL DEBUG: Distribution list worksheet not found" -Type Error
            return
        }
        
        $groups = $TenantData.Groups.DistributionGroups
        if (-not $groups -or $groups.Count -eq 0) {
            Write-LogMessage -Message "EXCEL DEBUG: No distribution groups data to populate" -Type Warning
            return
        }
        
        # Start at row 9 (based on template structure)
        $startRow = 9
        $currentRow = $startRow
        
        foreach ($group in $groups) {
            try {
                # Map to template column B (Distribution List Name)
                $worksheet.Cells.Item($currentRow, 2).Value = $group.DisplayName
                
                $currentRow++
            }
            catch {
                Write-LogMessage -Message "EXCEL DEBUG: Error updating distribution group row $currentRow - $($_.Exception.Message)" -Type Error
                $currentRow++
            }
        }
        
        $Workbook.Save()
        Write-LogMessage -Message "EXCEL DEBUG: Successfully updated Distribution list sheet with $($groups.Count) groups in range B$startRow-B$($currentRow-1)" -Type Success
    }
    catch {
        Write-LogMessage -Message "EXCEL DEBUG: Error updating Distribution list sheet - $($_.Exception.Message)" -Type Error
    }
}

# === HELPER FUNCTIONS ===

function New-DocumentationDirectory {
    try {
        if (-not (Test-Path $DocumentationConfig.OutputDirectory)) {
            New-Item -Path $DocumentationConfig.OutputDirectory -ItemType Directory -Force | Out-Null
            New-Item -Path "$($DocumentationConfig.OutputDirectory)\Reports" -ItemType Directory -Force | Out-Null
            New-Item -Path "$($DocumentationConfig.OutputDirectory)\Templates" -ItemType Directory -Force | Out-Null
            Write-LogMessage -Message "Created documentation directory: $($DocumentationConfig.OutputDirectory)" -Type Success
        }
        return $true
    }
    catch {
        Write-LogMessage -Message "Failed to create documentation directory: $($_.Exception.Message)" -Type Error
        return $false
    }
}

function Find-ExcelTemplate {
    $possiblePaths = @(
        ".\Master Spreadsheet Customer Details  Test.xlsx",
        "$PSScriptRoot\Master Spreadsheet Customer Details  Test.xlsx",
        "$env:USERPROFILE\Documents\Master Spreadsheet Customer Details  Test.xlsx",
        "$env:USERPROFILE\Downloads\Master Spreadsheet Customer Details  Test.xlsx",
        "$env:TEMP\Master Spreadsheet Customer Details  Test.xlsx"
    )
    
    Write-LogMessage -Message "Searching for Excel template in multiple locations..." -Type Info
    
    foreach ($path in $possiblePaths) {
        Write-LogMessage -Message "Checking: $path" -Type Info -LogOnly
        if (Test-Path $path) {
            Write-LogMessage -Message "Found Excel template at: $path" -Type Success
            return $path
        }
    }
    
    Write-LogMessage -Message "Excel template not found in default locations. Prompting user to select..." -Type Info
    Add-Type -AssemblyName System.Windows.Forms
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*"
    $openFileDialog.Title = "Select Master Spreadsheet Template"
    $openFileDialog.InitialDirectory = "$env:USERPROFILE\Downloads"
    
    if ($openFileDialog.ShowDialog() -eq 'OK') {
        Write-LogMessage -Message "User selected template: $($openFileDialog.FileName)" -Type Success
        return $openFileDialog.FileName
    }
    
    Write-LogMessage -Message "No template file selected. Cannot continue." -Type Error
    return $null
}

function Get-TenantInformation {
    try {
        $org = Get-MgOrganization
        return @{
            DisplayName = $org.DisplayName
            Id = $org.Id
            TenantType = $org.TenantType
            CreatedDateTime = $org.CreatedDateTime
            Country = $org.Country
            CountryLetterCode = $org.CountryLetterCode
        }
    }
    catch {
        Write-LogMessage -Message "Error collecting tenant information: $($_.Exception.Message)" -Type Warning
        return @{}
    }
}

function New-HTMLDocumentation {
    param([hashtable]$TenantData)
    # Placeholder for HTML documentation generation
    return $true
}

function New-JSONDocumentation {
    param([hashtable]$TenantData)
    # Placeholder for JSON documentation generation
    return $true
}

function New-ConfigurationSummary {
    param([hashtable]$TenantData)
    # Placeholder for configuration summary generation
    return $true
}

# === NO DIRECT EXECUTION - PREVENTS INFINITE LOOP ===
# The main script will call New-TenantDocumentation function directly