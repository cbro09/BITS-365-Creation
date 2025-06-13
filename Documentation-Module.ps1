#requires -Version 5.1
<#
.SYNOPSIS
    COMPLETE FIXED Documentation Module for Microsoft 365 Tenant Setup Utility
.DESCRIPTION
    Generates comprehensive documentation with FIXED formatting, permissions, and cmdlet issues
.NOTES
    Version: 1.1 - COMPLETE FIXED VERSION
    Dependencies: Microsoft Graph PowerShell SDK, ImportExcel module
    
    FIXES IMPLEMENTED:
    - Fixed SharePoint access issues with multiple fallback methods
    - Fixed Intune mobile app cmdlets (corrected cmdlet names)
    - Fixed CA policies to show detailed settings in proper table format
    - Fixed Intune Configuration policies to show detailed settings
    - Enhanced error handling and logging throughout
    - Improved Excel column formatting and data placement
#>

# === Documentation Configuration ===
$DocumentationConfig = @{
    OutputDirectory = "$env:USERPROFILE\Documents\M365TenantSetup_Documentation_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
    ReportFormats = @('HTML', 'Excel', 'JSON')
    IncludeScreenshots = $false
    DetailLevel = 'Detailed' # Basic, Standard, Detailed
}

# === FIXED: Core Documentation Functions ===
function New-TenantDocumentation {
    <#
    .SYNOPSIS
        Main function to generate comprehensive tenant documentation - COMPLETE FIXED VERSION
    .DESCRIPTION
        Creates detailed documentation by populating the Excel template with actual tenant configuration
    #>
    
    try {
        Write-LogMessage -Message "Starting COMPLETE FIXED tenant documentation generation..." -Type Info
        
        # Verify Graph connection
        if (-not (Get-MgContext)) {
            Write-LogMessage -Message "Not connected to Microsoft Graph. Please connect first." -Type Error
            return $false
        }
        
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
        
        # Gather all tenant information with FIXED collection methods
        Write-LogMessage -Message "Gathering tenant configuration data with COMPLETE FIXED methods..." -Type Info
        $tenantData = Get-CompleteTenantConfiguration-Fixed
        
        # Generate populated Excel documentation with FIXED formatting
        Write-LogMessage -Message "Populating Excel template with COMPLETE FIXED formatting..." -Type Info
        $excelGenerated = New-PopulatedExcelDocumentation-Fixed -TenantData $tenantData -TemplatePath $templatePath
        
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
        
        Write-LogMessage -Message "COMPLETE FIXED documentation generation completed. Generated $documentsGenerated documents." -Type Success
        Write-LogMessage -Message "Documentation saved to: $($DocumentationConfig.OutputDirectory)" -Type Info
        
        # Open the documentation directory
        $openDirectory = Read-Host "Would you like to open the documentation directory? (Y/N)"
        if ($openDirectory -eq 'Y' -or $openDirectory -eq 'y') {
            Start-Process -FilePath "explorer.exe" -ArgumentList $DocumentationConfig.OutputDirectory
        }
        
        return $true
    }
    catch {
        Write-LogMessage -Message "Failed to generate COMPLETE FIXED documentation: $($_.Exception.Message)" -Type Error
        return $false
    }
}

function New-DocumentationDirectory {
    <#
    .SYNOPSIS
        Creates the documentation output directory structure
    #>
    
    try {
        # Create main directory
        if (-not (Test-Path -Path $DocumentationConfig.OutputDirectory)) {
            New-Item -Path $DocumentationConfig.OutputDirectory -ItemType Directory -Force | Out-Null
        }
        
        # Create subdirectories
        $subDirectories = @('Reports', 'Exports', 'Screenshots', 'Templates')
        foreach ($subDir in $subDirectories) {
            $subDirPath = Join-Path -Path $DocumentationConfig.OutputDirectory -ChildPath $subDir
            if (-not (Test-Path -Path $subDirPath)) {
                New-Item -Path $subDirPath -ItemType Directory -Force | Out-Null
            }
        }
        
        Write-LogMessage -Message "Documentation directory structure created: $($DocumentationConfig.OutputDirectory)" -Type Success -LogOnly
        return $true
    }
    catch {
        Write-LogMessage -Message "Failed to create documentation directory: $($_.Exception.Message)" -Type Error
        return $false
    }
}

function Find-ExcelTemplate {
    <#
    .SYNOPSIS
        Locates the Excel template file for population
    #>
    
    try {
        # Common locations to search for the template
        $searchPaths = @(
            "$env:USERPROFILE\Documents\Master Spreadsheet Customer Details  Test.xlsx",
            "$env:USERPROFILE\Downloads\Master Spreadsheet Customer Details  Test.xlsx",
            ".\Master Spreadsheet Customer Details  Test.xlsx",
            "$env:USERPROFILE\Documents\M365TenantSetup_Documentation*\Templates\Master Spreadsheet Customer Details  Test.xlsx"
        )
        
        foreach ($path in $searchPaths) {
            $resolvedPaths = Resolve-Path -Path $path -ErrorAction SilentlyContinue
            if ($resolvedPaths) {
                foreach ($resolvedPath in $resolvedPaths) {
                    if (Test-Path -Path $resolvedPath) {
                        Write-LogMessage -Message "Found Excel template at: $resolvedPath" -Type Success -LogOnly
                        return $resolvedPath.Path
                    }
                }
            }
        }
        
        # If not found, prompt user to select
        Write-LogMessage -Message "Excel template not found in default locations. Prompting user to select..." -Type Info
        
        Add-Type -AssemblyName System.Windows.Forms
        $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $openFileDialog.Title = "Select Excel Template File"
        $openFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx"
        $openFileDialog.InitialDirectory = "$env:USERPROFILE\Documents"
        
        if ($openFileDialog.ShowDialog() -eq 'OK') {
            Write-LogMessage -Message "User selected template: $($openFileDialog.FileName)" -Type Success -LogOnly
            return $openFileDialog.FileName
        }
        
        return $null
    }
    catch {
        Write-LogMessage -Message "Error finding Excel template: $($_.Exception.Message)" -Type Error
        return $null
    }
}

# === FIXED: Data Collection Functions ===

function Get-CompleteTenantConfiguration-Fixed {
    <#
    .SYNOPSIS
        COMPLETE FIXED version of tenant configuration data gathering
    .DESCRIPTION
        Gathers comprehensive tenant configuration data with enhanced error handling and detailed extraction
    #>
    
    Write-LogMessage -Message "Gathering tenant configuration data with COMPLETE FIXED methods..." -Type Info
    
    $tenantData = @{
        GeneratedOn = Get-Date
        TenantInfo = @{}
        Groups = @{}
        ConditionalAccess = @{}
        SharePoint = @{}
        Intune = @{}
        Users = @{}
        Licenses = @{}
        Security = @{}
        Compliance = @{}
    }
    
    try {
        # Basic Tenant Information
        Write-LogMessage -Message "Collecting tenant information..." -Type Info -LogOnly
        $tenantData.TenantInfo = Get-TenantInformation
        
        # Groups Information
        Write-LogMessage -Message "Collecting groups information..." -Type Info -LogOnly
        $tenantData.Groups = Get-GroupsInformation
        
        # FIXED: Conditional Access Policies with detailed settings
        Write-LogMessage -Message "Collecting conditional access policies with DETAILED settings..." -Type Info -LogOnly
        $tenantData.ConditionalAccess = Get-ConditionalAccessInformation-Fixed
        
        # FIXED: SharePoint Information with better error handling
        Write-LogMessage -Message "Collecting SharePoint information with FIXED permissions handling..." -Type Info -LogOnly
        $tenantData.SharePoint = Get-SharePointInformation-Fixed
        
        # FIXED: Intune Information with corrected cmdlets and detailed policy settings
        Write-LogMessage -Message "Collecting Intune information with FIXED cmdlets and detailed settings..." -Type Info -LogOnly
        $tenantData.Intune = Get-IntuneInformation-Fixed
        
        # Users Information
        Write-LogMessage -Message "Collecting users information..." -Type Info -LogOnly
        $tenantData.Users = Get-UsersInformation
        
        # License Information
        Write-LogMessage -Message "Collecting license information..." -Type Info -LogOnly
        $tenantData.Licenses = Get-LicenseInformation
        
        # Security Settings
        Write-LogMessage -Message "Collecting security settings..." -Type Info -LogOnly
        $tenantData.Security = Get-SecurityInformation
        
        Write-LogMessage -Message "COMPLETE FIXED tenant configuration data collection completed" -Type Success -LogOnly
        return $tenantData
    }
    catch {
        Write-LogMessage -Message "Error collecting COMPLETE FIXED tenant data: $($_.Exception.Message)" -Type Error
        return $tenantData
    }
}

function Get-TenantInformation {
    <#
    .SYNOPSIS
        Collects basic tenant information
    #>
    
    try {
        $organization = Get-MgOrganization
        $context = Get-MgContext
        
        $tenantInfo = @{
            TenantId = $organization.Id
            DisplayName = $organization.DisplayName
            DefaultDomain = ($organization.VerifiedDomains | Where-Object { $_.IsDefault -eq $true }).Name
            VerifiedDomains = $organization.VerifiedDomains | ForEach-Object { $_.Name }
            CountryCode = $organization.CountryLetterCode
            City = $organization.City
            State = $organization.State
            CreatedDateTime = $organization.CreatedDateTime
            ConnectedAs = $context.Account
            ConnectedScopes = $context.Scopes
            SetupDate = Get-Date
        }
        
        # Add script tenant state if available
        if ($script:TenantState) {
            $tenantInfo.AdminEmail = $script:TenantState.AdminEmail
            $tenantInfo.SetupGroups = $script:TenantState.CreatedGroups
        }
        
        return $tenantInfo
    }
    catch {
        Write-LogMessage -Message "Error collecting tenant information: $($_.Exception.Message)" -Type Warning -LogOnly
        return @{}
    }
}

function Get-GroupsInformation {
    <#
    .SYNOPSIS
        Collects information about all groups in the tenant
    #>
    
    try {
        $groups = Get-MgGroup -All
        $groupsInfo = @{
            SecurityGroups = @()
            DistributionGroups = @()
            Microsoft365Groups = @()
            DynamicGroups = @()
            TotalCount = $groups.Count
        }
        
        foreach ($group in $groups) {
            $groupData = @{
                Id = $group.Id
                DisplayName = $group.DisplayName
                Description = $group.Description
                GroupTypes = $group.GroupTypes
                SecurityEnabled = $group.SecurityEnabled
                MailEnabled = $group.MailEnabled
                CreatedDateTime = $group.CreatedDateTime
                MembershipRule = $group.MembershipRule
                MembershipRuleProcessingState = $group.MembershipRuleProcessingState
            }
            
            # Try to get member count
            try {
                $members = Get-MgGroupMember -GroupId $group.Id -All
                $groupData.MemberCount = $members.Count
            }
            catch {
                $groupData.MemberCount = "Unable to retrieve"
            }
            
            # Categorize groups
            if ($group.GroupTypes -contains "Unified") {
                $groupsInfo.Microsoft365Groups += $groupData
            }
            elseif ($group.GroupTypes -contains "DynamicMembership") {
                $groupsInfo.DynamicGroups += $groupData
            }
            elseif ($group.SecurityEnabled -and -not $group.MailEnabled) {
                $groupsInfo.SecurityGroups += $groupData
            }
            elseif ($group.MailEnabled -and -not $group.SecurityEnabled) {
                $groupsInfo.DistributionGroups += $groupData
            }
        }
        
        return $groupsInfo
    }
    catch {
        Write-LogMessage -Message "Error collecting groups information: $($_.Exception.Message)" -Type Warning -LogOnly
        return @{}
    }
}

function Get-ConditionalAccessInformation-Fixed {
    <#
    .SYNOPSIS
        COMPLETE FIXED: Collects conditional access policies with DETAILED settings for proper Excel formatting
    .DESCRIPTION
        Extracts comprehensive CA policy details including user conditions, app conditions, platform restrictions,
        grant controls, and risk levels - formatted for proper Excel table display
    #>
    
    try {
        Write-LogMessage -Message "FIXED CA: Starting detailed conditional access policy collection..." -Type Info -LogOnly
        
        $policies = Get-MgIdentityConditionalAccessPolicy -All
        $caInfo = @{
            Policies = @()
            TotalCount = $policies.Count
            EnabledCount = ($policies | Where-Object { $_.State -eq "enabled" }).Count
            DisabledCount = ($policies | Where-Object { $_.State -eq "disabled" }).Count
            PolicyDetails = @() # NEW: Detailed policy breakdown for Excel formatting
        }
        
        Write-LogMessage -Message "FIXED CA: Processing $($policies.Count) conditional access policies..." -Type Info -LogOnly
        
        foreach ($policy in $policies) {
            Write-LogMessage -Message "FIXED CA: Processing policy '$($policy.DisplayName)'..." -Type Info -LogOnly
            
            # FIXED: Extract detailed policy settings for proper Excel table formatting
            $policyData = @{
                Id = $policy.Id
                DisplayName = $policy.DisplayName
                State = $policy.State
                CreatedDateTime = $policy.CreatedDateTime
                ModifiedDateTime = $policy.ModifiedDateTime
                
                # DETAILED CONDITIONS - Properly extracted for Excel
                UserConditions = ""
                ApplicationConditions = ""
                PlatformConditions = ""
                LocationConditions = ""
                RiskConditions = ""
                GrantControls = ""
                SessionControls = ""
                
                # RAW conditions for reference
                Conditions = @{
                    Users = $policy.Conditions.Users
                    Applications = $policy.Conditions.Applications
                    Platforms = $policy.Conditions.Platforms
                    Locations = $policy.Conditions.Locations
                    ClientAppTypes = $policy.Conditions.ClientAppTypes
                    SignInRiskLevels = $policy.Conditions.SignInRiskLevels
                    UserRiskLevels = $policy.Conditions.UserRiskLevels
                }
                GrantControlsRaw = $policy.GrantControls
                SessionControlsRaw = $policy.SessionControls
            }
            
            # FIXED: Format user conditions for Excel display
            if ($policy.Conditions.Users) {
                $userParts = @()
                if ($policy.Conditions.Users.IncludeUsers) {
                    $includeUsers = $policy.Conditions.Users.IncludeUsers -join ", "
                    if ($includeUsers -eq "All") {
                        $userParts += "All Users"
                    } else {
                        $userParts += "Include: $includeUsers"
                    }
                }
                if ($policy.Conditions.Users.ExcludeUsers) {
                    $excludeUsers = $policy.Conditions.Users.ExcludeUsers -join ", "
                    $userParts += "Exclude: $excludeUsers"
                }
                if ($policy.Conditions.Users.IncludeGroups) {
                    $includeGroups = $policy.Conditions.Users.IncludeGroups -join ", "
                    $userParts += "Include Groups: $includeGroups"
                }
                if ($policy.Conditions.Users.ExcludeGroups) {
                    $excludeGroups = $policy.Conditions.Users.ExcludeGroups -join ", "
                    $userParts += "Exclude Groups: $excludeGroups"
                }
                $policyData.UserConditions = $userParts -join " | "
            }
            
            # FIXED: Format application conditions for Excel display
            if ($policy.Conditions.Applications) {
                $appParts = @()
                if ($policy.Conditions.Applications.IncludeApplications) {
                    $includeApps = $policy.Conditions.Applications.IncludeApplications -join ", "
                    if ($includeApps -eq "All") {
                        $appParts += "All Cloud Apps"
                    } else {
                        $appParts += "Include: $includeApps"
                    }
                }
                if ($policy.Conditions.Applications.ExcludeApplications) {
                    $excludeApps = $policy.Conditions.Applications.ExcludeApplications -join ", "
                    $appParts += "Exclude: $excludeApps"
                }
                if ($policy.Conditions.Applications.IncludeUserActions) {
                    $userActions = $policy.Conditions.Applications.IncludeUserActions -join ", "
                    $appParts += "User Actions: $userActions"
                }
                $policyData.ApplicationConditions = $appParts -join " | "
            }
            
            # FIXED: Format platform conditions
            if ($policy.Conditions.Platforms) {
                $platformParts = @()
                if ($policy.Conditions.Platforms.IncludePlatforms) {
                    $platforms = $policy.Conditions.Platforms.IncludePlatforms -join ", "
                    $platformParts += "Include: $platforms"
                }
                if ($policy.Conditions.Platforms.ExcludePlatforms) {
                    $excludePlatforms = $policy.Conditions.Platforms.ExcludePlatforms -join ", "
                    $platformParts += "Exclude: $excludePlatforms"
                }
                $policyData.PlatformConditions = $platformParts -join " | "
            }
            
            # FIXED: Format location conditions
            if ($policy.Conditions.Locations) {
                $locationParts = @()
                if ($policy.Conditions.Locations.IncludeLocations) {
                    $locations = $policy.Conditions.Locations.IncludeLocations -join ", "
                    $locationParts += "Include: $locations"
                }
                if ($policy.Conditions.Locations.ExcludeLocations) {
                    $excludeLocations = $policy.Conditions.Locations.ExcludeLocations -join ", "
                    $locationParts += "Exclude: $excludeLocations"
                }
                $policyData.LocationConditions = $locationParts -join " | "
            }
            
            # FIXED: Format grant controls for Excel display
            if ($policy.GrantControls) {
                $grantParts = @()
                if ($policy.GrantControls.BuiltInControls) {
                    $controls = $policy.GrantControls.BuiltInControls -join ", "
                    $grantParts += "Controls: $controls"
                }
                if ($policy.GrantControls.Operator) {
                    $grantParts += "Operator: $($policy.GrantControls.Operator)"
                }
                if ($policy.GrantControls.CustomAuthenticationFactors) {
                    $customFactors = $policy.GrantControls.CustomAuthenticationFactors -join ", "
                    $grantParts += "Custom: $customFactors"
                }
                $policyData.GrantControls = $grantParts -join " | "
            }
            
            # FIXED: Format session controls
            if ($policy.SessionControls) {
                $sessionParts = @()
                if ($policy.SessionControls.ApplicationEnforcedRestrictions) {
                    $sessionParts += "App Restrictions: $($policy.SessionControls.ApplicationEnforcedRestrictions.IsEnabled)"
                }
                if ($policy.SessionControls.CloudAppSecurity) {
                    $sessionParts += "CASB: $($policy.SessionControls.CloudAppSecurity.IsEnabled)"
                }
                if ($policy.SessionControls.SignInFrequency) {
                    $sessionParts += "Sign-in Frequency: $($policy.SessionControls.SignInFrequency.IsEnabled)"
                }
                if ($policy.SessionControls.PersistentBrowser) {
                    $sessionParts += "Persistent Browser: $($policy.SessionControls.PersistentBrowser.IsEnabled)"
                }
                $policyData.SessionControls = $sessionParts -join " | "
            }
            
            # FIXED: Format risk conditions
            $riskParts = @()
            if ($policy.Conditions.SignInRiskLevels) {
                $signInRisk = $policy.Conditions.SignInRiskLevels -join ", "
                $riskParts += "Sign-in Risk: $signInRisk"
            }
            if ($policy.Conditions.UserRiskLevels) {
                $userRisk = $policy.Conditions.UserRiskLevels -join ", "
                $riskParts += "User Risk: $userRisk"
            }
            if ($riskParts.Count -gt 0) {
                $policyData.RiskConditions = $riskParts -join " | "
            }
            
            $caInfo.Policies += $policyData
            
            # FIXED: Create simplified policy details for Excel table formatting
            $policyDetail = [PSCustomObject]@{
                PolicyName = $policy.DisplayName
                State = $policy.State
                Users = $policyData.UserConditions
                Applications = $policyData.ApplicationConditions
                Platforms = $policyData.PlatformConditions
                Locations = $policyData.LocationConditions
                GrantControls = $policyData.GrantControls
                SessionControls = $policyData.SessionControls
                RiskLevels = $policyData.RiskConditions
            }
            $caInfo.PolicyDetails += $policyDetail
            
            Write-LogMessage -Message "FIXED CA: Processed policy '$($policy.DisplayName)' with detailed settings" -Type Info -LogOnly
        }
        
        Write-LogMessage -Message "FIXED CA: Collected $($policies.Count) policies with comprehensive detailed settings" -Type Success
        return $caInfo
    }
    catch {
        Write-LogMessage -Message "Error collecting FIXED conditional access information: $($_.Exception.Message)" -Type Warning -LogOnly
        return @{
            Policies = @()
            PolicyDetails = @()
            TotalCount = 0
            EnabledCount = 0
            DisabledCount = 0
        }
    }
}

function Get-SharePointInformation-Fixed {
    <#
    .SYNOPSIS
        COMPLETE FIXED: SharePoint information collection with comprehensive permission handling
    .DESCRIPTION
        Uses multiple access methods and provides detailed feedback about permission requirements
    #>
    
    try {
        $spInfo = @{
            TenantSettings = @{}
            SiteCollections = @()
            TotalSites = 0
            StorageUsed = "Not available"
            SharingSettings = "Not available"
            ExternalSharingEnabled = "Not available"
            AccessStatus = "Unknown"
            PermissionMessage = ""
        }
        
        Write-LogMessage -Message "FIXED SP: Testing SharePoint access permissions with multiple methods..." -Type Info
        
        # FIXED: Try different permission levels and methods with detailed error handling
        $accessMethods = @(
            @{ 
                Method = "Get-MgSite -SiteId root"; 
                Description = "Root site access"; 
                RequiredPermission = "Sites.Read.All or Sites.ReadWrite.All"
            },
            @{ 
                Method = "Get-MgSite -Top 10"; 
                Description = "Limited site enumeration"; 
                RequiredPermission = "Sites.Read.All"
            },
            @{ 
                Method = "Get-MgSite -Search 'site'"; 
                Description = "Site search"; 
                RequiredPermission = "Sites.Read.All"
            }
        )
        
        $sitesFound = $false
        $permissionErrors = @()
        
        foreach ($method in $accessMethods) {
            try {
                Write-LogMessage -Message "FIXED SP: Attempting method: $($method.Description)" -Type Info -LogOnly
                
                switch ($method.Method) {
                    "Get-MgSite -SiteId root" {
                        $rootSite = Get-MgSite -SiteId "root" -ErrorAction Stop
                        if ($rootSite) {
                            $spInfo.SiteCollections += @{
                                Id = $rootSite.Id
                                DisplayName = $rootSite.DisplayName ?? "Root Site"
                                Name = $rootSite.Name ?? "Root"
                                WebUrl = $rootSite.WebUrl
                                CreatedDateTime = $rootSite.CreatedDateTime
                                LastModifiedDateTime = $rootSite.LastModifiedDateTime
                                SiteCollection = $rootSite.SiteCollection
                                AccessMethod = "Root Site Access"
                            }
                            $sitesFound = $true
                            $spInfo.AccessStatus = "Success - Root Site Access"
                            Write-LogMessage -Message "FIXED SP: Successfully accessed root site" -Type Success
                        }
                    }
                    "Get-MgSite -Top 10" {
                        if (-not $sitesFound) {
                            $sites = Get-MgSite -Top 10 -ErrorAction Stop
                            foreach ($site in $sites) {
                                $spInfo.SiteCollections += @{
                                    Id = $site.Id
                                    DisplayName = $site.DisplayName ?? "Site Collection"
                                    Name = $site.Name ?? "Unknown"
                                    WebUrl = $site.WebUrl
                                    CreatedDateTime = $site.CreatedDateTime
                                    LastModifiedDateTime = $site.LastModifiedDateTime
                                    SiteCollection = $site.SiteCollection
                                    AccessMethod = "Limited Enumeration"
                                }
                            }
                            $sitesFound = $true
                            $spInfo.AccessStatus = "Success - Limited Enumeration ($($sites.Count) sites)"
                            Write-LogMessage -Message "FIXED SP: Successfully enumerated $($sites.Count) sites" -Type Success
                        }
                    }
                    "Get-MgSite -Search 'site'" {
                        if (-not $sitesFound) {
                            $sites = Get-MgSite -Search "site" -ErrorAction Stop
                            foreach ($site in $sites) {
                                $spInfo.SiteCollections += @{
                                    Id = $site.Id
                                    DisplayName = $site.DisplayName ?? "Search Result"
                                    Name = $site.Name ?? "Unknown"
                                    WebUrl = $site.WebUrl
                                    CreatedDateTime = $site.CreatedDateTime
                                    LastModifiedDateTime = $site.LastModifiedDateTime
                                    SiteCollection = $site.SiteCollection
                                    AccessMethod = "Search Results"
                                }
                            }
                            $sitesFound = $true
                            $spInfo.AccessStatus = "Success - Search Results ($($sites.Count) sites)"
                            Write-LogMessage -Message "FIXED SP: Successfully found $($sites.Count) sites via search" -Type Success
                        }
                    }
                }
                
                if ($sitesFound) {
                    break
                }
            }
            catch {
                $errorMessage = $_.Exception.Message
                $permissionErrors += "$($method.Description): $errorMessage (Requires: $($method.RequiredPermission))"
                Write-LogMessage -Message "FIXED SP: Method '$($method.Description)' failed: $errorMessage" -Type Warning -LogOnly
                continue
            }
        }
        
        # FIXED: If no access methods worked, provide comprehensive feedback
        if (-not $sitesFound) {
            $spInfo.AccessStatus = "Access Denied - All Methods Failed"
            $spInfo.PermissionMessage = "SharePoint access requires one of the following Graph API permissions: Sites.Read.All, Sites.ReadWrite.All, or Sites.FullControl.All"
            
            # Add detailed error information
            $spInfo.SiteCollections = @(
                @{
                    Id = "PERMISSION_ERROR"
                    DisplayName = "Access Denied - Insufficient SharePoint Permissions"
                    Name = "PERMISSION_ERROR"
                    WebUrl = "Requires Sites.Read.All or higher permission"
                    CreatedDateTime = Get-Date
                    LastModifiedDateTime = Get-Date
                    SiteCollection = $null
                    AccessMethod = "Permission Error"
                    ErrorDetails = $permissionErrors -join " | "
                }
            )
            Write-LogMessage -Message "FIXED SP: All SharePoint access methods failed. Required permissions: Sites.Read.All or Sites.ReadWrite.All" -Type Warning
            Write-LogMessage -Message "FIXED SP: Error details: $($permissionErrors -join '; ')" -Type Warning -LogOnly
        } else {
            $spInfo.TotalSites = $spInfo.SiteCollections.Count
            $spInfo.PermissionMessage = "SharePoint access successful with current permissions"
            Write-LogMessage -Message "FIXED SP: Successfully collected $($spInfo.TotalSites) SharePoint sites" -Type Success
        }
        
        return $spInfo
    }
    catch {
        Write-LogMessage -Message "FIXED SP: Critical error collecting SharePoint information: $($_.Exception.Message)" -Type Error
        return @{
            TenantSettings = @{}
            SiteCollections = @(
                @{
                    Id = "CRITICAL_ERROR"
                    DisplayName = "Critical Error accessing SharePoint data"
                    Name = "CRITICAL_ERROR"
                    WebUrl = "Check Graph connection and permissions"
                    CreatedDateTime = Get-Date
                    LastModifiedDateTime = Get-Date
                    SiteCollection = $null
                    AccessMethod = "Critical Error"
                    ErrorDetails = $_.Exception.Message
                }
            )
            TotalSites = 0
            StorageUsed = "Error"
            SharingSettings = "Error"
            ExternalSharingEnabled = "Error"
            AccessStatus = "Critical Error"
            PermissionMessage = "Critical error occurred during SharePoint data collection"
        }
    }
}

function Get-IntuneInformation-Fixed {
    <#
    .SYNOPSIS
        COMPLETE FIXED: Intune information collection with corrected cmdlets and comprehensive policy settings
    .DESCRIPTION
        Uses correct Graph cmdlets and extracts detailed policy settings for proper Excel formatting
    #>
    
    try {
        $intuneInfo = @{
            DeviceCompliancePolicies = @()
            DeviceConfigurationPolicies = @()
            AppProtectionPolicies = @()
            EnrollmentRestrictions = @()
            ManagedDevices = @()
            ManagedApps = @()
            TotalDevices = 0
            PolicyDetails = @() # NEW: Detailed policy settings for Excel formatting
            AccessStatus = @{
                CompliancePolicies = "Unknown"
                ConfigurationPolicies = "Unknown"
                ManagedApps = "Unknown"
                ManagedDevices = "Unknown"
            }
        }
        
        Write-LogMessage -Message "FIXED INTUNE: Starting comprehensive Intune data collection with corrected cmdlets..." -Type Info
        
        # FIXED: Device Compliance Policies with detailed settings extraction
        try {
            Write-LogMessage -Message "FIXED INTUNE: Collecting compliance policies with detailed settings..." -Type Info
            $compliancePolicies = Get-MgDeviceManagementDeviceCompliancePolicy -All
            Write-LogMessage -Message "FIXED INTUNE: Found $($compliancePolicies.Count) compliance policies" -Type Info
            
            foreach ($policy in $compliancePolicies) {
                Write-LogMessage -Message "FIXED INTUNE: Processing compliance policy '$($policy.DisplayName)'..." -Type Info -LogOnly
                
                # FIXED: Get detailed policy settings
                try {
                    $policyDetails = Get-MgDeviceManagementDeviceCompliancePolicy -DeviceCompliancePolicyId $policy.Id -ErrorAction SilentlyContinue
                    $detailedSettings = "Policy Type: $($policy.ODataType -replace '.*\.', '')"
                    
                    # FIXED: Extract meaningful settings based on policy type
                    $settingsParts = @()
                    if ($policy.ODataType -like "*windows*") {
                        $settingsParts += "Platform: Windows"
                    } elseif ($policy.ODataType -like "*android*") {
                        $settingsParts += "Platform: Android"
                    } elseif ($policy.ODataType -like "*ios*") {
                        $settingsParts += "Platform: iOS"
                    } elseif ($policy.ODataType -like "*mac*") {
                        $settingsParts += "Platform: macOS"
                    }
                    
                    # Try to get additional settings if available
                    if ($policyDetails.AdditionalProperties) {
                        $relevantSettings = @()
                        foreach ($key in $policyDetails.AdditionalProperties.Keys) {
                            if ($key -notlike "*@*" -and $policyDetails.AdditionalProperties[$key] -ne $null) {
                                $value = $policyDetails.AdditionalProperties[$key]
                                if ($value -is [bool]) {
                                    $relevantSettings += "$key = $value"
                                } elseif ($value -is [string] -and $value.Length -lt 50 -and $value -ne "") {
                                    $relevantSettings += "$key = $value"
                                } elseif ($value -is [int] -or $value -is [double]) {
                                    $relevantSettings += "$key = $value"
                                }
                            }
                        }
                        if ($relevantSettings.Count -gt 0) {
                            $settingsParts += $relevantSettings[0..4] # Limit to first 5 settings
                        }
                    }
                    
                    if ($settingsParts.Count -gt 0) {
                        $detailedSettings = $settingsParts -join " | "
                    }
                } catch {
                    $detailedSettings = "Policy Type: $($policy.ODataType -replace '.*\.', '') | Settings: Unable to retrieve detailed settings"
                }
                
                $policyData = @{
                    Id = $policy.Id
                    DisplayName = $policy.DisplayName
                    Description = $policy.Description ?? "No description"
                    CreatedDateTime = $policy.CreatedDateTime
                    LastModifiedDateTime = $policy.LastModifiedDateTime
                    Version = $policy.Version
                    PolicyType = $policy.ODataType -replace '.*\.', ''
                    DetailedSettings = $detailedSettings
                    Platform = if ($policy.ODataType -like "*windows*") { "Windows" } 
                               elseif ($policy.ODataType -like "*android*") { "Android" }
                               elseif ($policy.ODataType -like "*ios*") { "iOS" }
                               elseif ($policy.ODataType -like "*mac*") { "macOS" }
                               else { "Unknown" }
                }
                
                $intuneInfo.DeviceCompliancePolicies += $policyData
            }
            $intuneInfo.AccessStatus.CompliancePolicies = "Success"
        }
        catch {
            Write-LogMessage -Message "FIXED INTUNE: Could not retrieve device compliance policies: $($_.Exception.Message)" -Type Warning
            $intuneInfo.AccessStatus.CompliancePolicies = "Failed: $($_.Exception.Message)"
        }
        
        # FIXED: Device Configuration Policies with detailed settings extraction
        try {
            Write-LogMessage -Message "FIXED INTUNE: Collecting configuration policies with detailed settings..." -Type Info
            $configPolicies = Get-MgDeviceManagementDeviceConfiguration -All
            Write-LogMessage -Message "FIXED INTUNE: Found $($configPolicies.Count) configuration policies" -Type Info
            
            foreach ($policy in $configPolicies) {
                Write-LogMessage -Message "FIXED INTUNE: Processing configuration policy '$($policy.DisplayName)'..." -Type Info -LogOnly
                
                # FIXED: Get detailed policy settings
                try {
                    $policyDetails = Get-MgDeviceManagementDeviceConfiguration -DeviceConfigurationId $policy.Id -ErrorAction SilentlyContinue
                    $detailedSettings = "Policy Type: $($policy.ODataType -replace '.*\.', '')"
                    
                    # FIXED: Extract meaningful settings based on policy type
                    $settingsParts = @()
                    if ($policy.ODataType -like "*windows*") {
                        $settingsParts += "Platform: Windows"
                    } elseif ($policy.ODataType -like "*android*") {
                        $settingsParts += "Platform: Android"
                    } elseif ($policy.ODataType -like "*ios*") {
                        $settingsParts += "Platform: iOS"
                    } elseif ($policy.ODataType -like "*mac*") {
                        $settingsParts += "Platform: macOS"
                    }
                    
                    # Try to extract meaningful settings
                    if ($policyDetails.AdditionalProperties) {
                        $relevantSettings = @()
                        foreach ($key in $policyDetails.AdditionalProperties.Keys) {
                            if ($key -notlike "*@*" -and $policyDetails.AdditionalProperties[$key] -ne $null) {
                                $value = $policyDetails.AdditionalProperties[$key]
                                if ($value -is [bool]) {
                                    $relevantSettings += "$key = $value"
                                } elseif ($value -is [string] -and $value.Length -lt 50 -and $value -ne "") {
                                    $relevantSettings += "$key = $value"
                                } elseif ($value -is [int] -or $value -is [double]) {
                                    $relevantSettings += "$key = $value"
                                }
                            }
                        }
                        if ($relevantSettings.Count -gt 0) {
                            $settingsParts += $relevantSettings[0..4] # Limit to first 5 settings
                        }
                    }
                    
                    if ($settingsParts.Count -gt 0) {
                        $detailedSettings = $settingsParts -join " | "
                    }
                } catch {
                    $detailedSettings = "Policy Type: $($policy.ODataType -replace '.*\.', '') | Settings: Unable to retrieve detailed settings"
                }
                
                $policyData = @{
                    Id = $policy.Id
                    DisplayName = $policy.DisplayName
                    Description = $policy.Description ?? "No description"
                    CreatedDateTime = $policy.CreatedDateTime
                    LastModifiedDateTime = $policy.LastModifiedDateTime
                    Version = $policy.Version
                    PolicyType = $policy.ODataType -replace '.*\.', ''
                    DetailedSettings = $detailedSettings
                    Platform = if ($policy.ODataType -like "*windows*") { "Windows" } 
                               elseif ($policy.ODataType -like "*android*") { "Android" }
                               elseif ($policy.ODataType -like "*ios*") { "iOS" }
                               elseif ($policy.ODataType -like "*mac*") { "macOS" }
                               else { "Unknown" }
                }
                
                $intuneInfo.DeviceConfigurationPolicies += $policyData
            }
            $intuneInfo.AccessStatus.ConfigurationPolicies = "Success"
        }
        catch {
            Write-LogMessage -Message "FIXED INTUNE: Could not retrieve device configuration policies: $($_.Exception.Message)" -Type Warning
            $intuneInfo.AccessStatus.ConfigurationPolicies = "Failed: $($_.Exception.Message)"
        }
        
        # FIXED: Managed Apps - Using correct cmdlet names with multiple approaches
        try {
            Write-LogMessage -Message "FIXED INTUNE: Collecting managed applications using corrected cmdlet methods..." -Type Info
            
            $managedApps = @()
            $appCollectionMethods = @(
                @{ 
                    Cmdlet = "Get-MgDeviceAppManagementMobileApp"; 
                    Description = "Primary mobile apps cmdlet";
                    Permission = "DeviceManagementApps.Read.All"
                },
                @{ 
                    Cmdlet = "Get-MgDeviceAppManagementManagedAppRegistration"; 
                    Description = "Managed app registrations";
                    Permission = "DeviceManagementApps.Read.All"
                }
            )
            
            $appsCollected = $false
            foreach ($method in $appCollectionMethods) {
                try {
                    Write-LogMessage -Message "FIXED INTUNE: Attempting app collection method: $($method.Description)" -Type Info
                    
                    # Check if cmdlet exists
                    if (-not (Get-Command $method.Cmdlet -ErrorAction SilentlyContinue)) {
                        Write-LogMessage -Message "FIXED INTUNE: Cmdlet '$($method.Cmdlet)' not available - may need module update" -Type Warning -LogOnly
                        continue
                    }
                    
                    switch ($method.Cmdlet) {
                        "Get-MgDeviceAppManagementMobileApp" {
                            $apps = Get-MgDeviceAppManagementMobileApp -All -ErrorAction Stop
                            foreach ($app in $apps) {
                                $managedApps += @{
                                    Id = $app.Id
                                    DisplayName = $app.DisplayName ?? "Unknown App"
                                    Description = $app.Description ?? "No description"
                                    Publisher = $app.Publisher ?? "Unknown Publisher"
                                    CreatedDateTime = $app.CreatedDateTime
                                    LastModifiedDateTime = $app.LastModifiedDateTime
                                    AppType = $app.ODataType -replace '.*\.', ''
                                    CollectionMethod = $method.Description
                                    Platform = if ($app.ODataType -like "*android*") { "Android" }
                                              elseif ($app.ODataType -like "*ios*") { "iOS" }
                                              elseif ($app.ODataType -like "*win32*") { "Windows" }
                                              elseif ($app.ODataType -like "*mac*") { "macOS" }
                                              else { "Unknown" }
                                }
                            }
                            Write-LogMessage -Message "FIXED INTUNE: Successfully collected $($apps.Count) mobile apps" -Type Success
                            $appsCollected = $true
                            break
                        }
                        "Get-MgDeviceAppManagementManagedAppRegistration" {
                            if (-not $appsCollected) {
                                $appRegs = Get-MgDeviceAppManagementManagedAppRegistration -All -ErrorAction Stop
                                foreach ($appReg in $appRegs) {
                                    $managedApps += @{
                                        Id = $appReg.Id
                                        DisplayName = $appReg.ApplicationDisplayName ?? "Registered App"
                                        Description = "Managed app registration"
                                        Publisher = "Various"
                                        CreatedDateTime = $appReg.CreatedDateTime
                                        LastModifiedDateTime = $appReg.LastModifiedDateTime
                                        AppType = "ManagedAppRegistration"
                                        CollectionMethod = $method.Description
                                        Platform = "Multi-platform"
                                    }
                                }
                                Write-LogMessage -Message "FIXED INTUNE: Successfully collected $($appRegs.Count) app registrations" -Type Success
                                $appsCollected = $true
                                break
                            }
                        }
                    }
                } catch {
                    Write-LogMessage -Message "FIXED INTUNE: Method '$($method.Description)' failed: $($_.Exception.Message)" -Type Warning -LogOnly
                    continue
                }
            }
            
            if ($managedApps.Count -gt 0) {
                $intuneInfo.ManagedApps = $managedApps
                $intuneInfo.AccessStatus.ManagedApps = "Success ($($managedApps.Count) apps)"
                Write-LogMessage -Message "FIXED INTUNE: Successfully collected $($managedApps.Count) managed apps total" -Type Success
            } else {
                Write-LogMessage -Message "FIXED INTUNE: No managed apps found - requires DeviceManagementApps.Read.All permission" -Type Warning
                $intuneInfo.ManagedApps = @()
                $intuneInfo.AccessStatus.ManagedApps = "No apps found - check permissions"
            }
        }
        catch {
            Write-LogMessage -Message "FIXED INTUNE: Could not retrieve managed applications: $($_.Exception.Message)" -Type Warning
            $intuneInfo.ManagedApps = @()
            $intuneInfo.AccessStatus.ManagedApps = "Failed: $($_.Exception.Message)"
        }
        
        # FIXED: Managed Devices
        try {
            Write-LogMessage -Message "FIXED INTUNE: Collecting managed devices..." -Type Info
            $devices = Get-MgDeviceManagementManagedDevice -All -Top 500
            $intuneInfo.TotalDevices = $devices.Count
            Write-LogMessage -Message "FIXED INTUNE: Found $($devices.Count) managed devices" -Type Info
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
                    ManagementState = $_.ManagementState
                    DeviceType = $_.DeviceType
                }
            }
            $intuneInfo.AccessStatus.ManagedDevices = "Success ($($devices.Count) devices)"
        }
        catch {
            Write-LogMessage -Message "FIXED INTUNE: Could not retrieve managed devices: $($_.Exception.Message)" -Type Warning
            $intuneInfo.AccessStatus.ManagedDevices = "Failed: $($_.Exception.Message)"
        }
        
        Write-LogMessage -Message "FIXED INTUNE: Enhanced data collection completed - Config: $($intuneInfo.DeviceConfigurationPolicies.Count), Compliance: $($intuneInfo.DeviceCompliancePolicies.Count), Apps: $($intuneInfo.ManagedApps.Count), Devices: $($intuneInfo.TotalDevices)" -Type Success
        return $intuneInfo
    }
    catch {
        Write-LogMessage -Message "FIXED INTUNE: Critical error collecting Intune information: $($_.Exception.Message)" -Type Error
        return @{
            DeviceCompliancePolicies = @()
            DeviceConfigurationPolicies = @()
            AppProtectionPolicies = @()
            EnrollmentRestrictions = @()
            ManagedDevices = @()
            ManagedApps = @()
            TotalDevices = 0
            PolicyDetails = @()
            AccessStatus = @{
                CompliancePolicies = "Critical Error"
                ConfigurationPolicies = "Critical Error"
                ManagedApps = "Critical Error"
                ManagedDevices = "Critical Error"
            }
        }
    }
}

function Get-UsersInformation {
    <#
    .SYNOPSIS
        Collects users information with license details
    #>
    
    try {
        Write-LogMessage -Message "Collecting users information..." -Type Info -LogOnly
        $users = Get-MgUser -All -Top 500
        $usersInfo = @{
            TotalUsers = $users.Count
            EnabledUsers = ($users | Where-Object { $_.AccountEnabled -eq $true }).Count
            DisabledUsers = ($users | Where-Object { $_.AccountEnabled -eq $false }).Count
            GuestUsers = ($users | Where-Object { $_.UserType -eq "Guest" }).Count
            Users = @()
        }
        
        foreach ($user in $users) {
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
                SignInActivity = "Not available"
            }
            
            # Try to get license information
            try {
                $userLicenses = Get-MgUserLicenseDetail -UserId $user.Id
                $userData.AssignedLicenses = $userLicenses | ForEach-Object { $_.SkuPartNumber }
            }
            catch {
                $userData.AssignedLicenses = @()
            }
            
            $usersInfo.Users += $userData
        }
        
        Write-LogMessage -Message "Successfully collected information for $($users.Count) users" -Type Success -LogOnly
        return $usersInfo
    }
    catch {
        Write-LogMessage -Message "Error collecting users information: $($_.Exception.Message)" -Type Warning -LogOnly
        return @{
            TotalUsers = 0
            EnabledUsers = 0
            DisabledUsers = 0
            GuestUsers = 0
            Users = @()
        }
    }
}

function Get-LicenseInformation {
    <#
    .SYNOPSIS
        Collects license information
    #>
    
    try {
        $subscribedSkus = Get-MgSubscribedSku
        $licenseInfo = @{
            SubscribedSkus = @()
            TotalLicenses = 0
            UsedLicenses = 0
        }
        
        foreach ($sku in $subscribedSkus) {
            $skuData = @{
                SkuId = $sku.SkuId
                SkuPartNumber = $sku.SkuPartNumber
                ServicePlans = $sku.ServicePlans | ForEach-Object { @{ ServicePlanName = $_.ServicePlanName; ServicePlanId = $_.ServicePlanId } }
                PrepaidUnits = $sku.PrepaidUnits
                ConsumedUnits = $sku.ConsumedUnits
                CapabilityStatus = $sku.CapabilityStatus
            }
            
            $licenseInfo.SubscribedSkus += $skuData
            $licenseInfo.TotalLicenses += $sku.PrepaidUnits.Enabled
            $licenseInfo.UsedLicenses += $sku.ConsumedUnits
        }
        
        return $licenseInfo
    }
    catch {
        Write-LogMessage -Message "Error collecting license information: $($_.Exception.Message)" -Type Warning -LogOnly
        return @{}
    }
}

function Get-SecurityInformation {
    <#
    .SYNOPSIS
        Collects security settings information
    #>
    
    try {
        $securityInfo = @{
            SecurityDefaults = @{}
            PasswordPolicy = @{}
            MFAStatus = "Not available"
            RiskyUsers = 0
            RiskySignIns = 0
        }
        
        # Try to get security defaults status
        try {
            $securityDefaults = Get-MgPolicyIdentitySecurityDefaultEnforcementPolicy
            $securityInfo.SecurityDefaults = @{
                IsEnabled = $securityDefaults.IsEnabled
                Description = $securityDefaults.Description
            }
        }
        catch {
            Write-LogMessage -Message "Could not retrieve security defaults status" -Type Warning -LogOnly
        }
        
        return $securityInfo
    }
    catch {
        Write-LogMessage -Message "Error collecting security information: $($_.Exception.Message)" -Type Warning -LogOnly
        return @{}
    }
}

# === FIXED: Excel Documentation Generation ===

function New-PopulatedExcelDocumentation-Fixed {
    <#
    .SYNOPSIS
        COMPLETE FIXED: Creates populated Excel documentation with enhanced formatting and error handling
    .DESCRIPTION
        Populates Excel template with comprehensive data using fixed sheet update functions
    #>
    param (
        [Parameter(Mandatory = $true)]
        [hashtable]$TenantData,
        
        [Parameter(Mandatory = $true)]
        [string]$TemplatePath
    )
    
    try {
        # Create output path
        $outputPath = Join-Path -Path $DocumentationConfig.OutputDirectory -ChildPath "Reports\TenantConfiguration_COMPLETE_FIXED_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
        
        # Copy template to output location
        Copy-Item -Path $TemplatePath -Destination $outputPath -Force
        
        # Open the Excel file for editing
        $excel = Open-ExcelPackage -Path $outputPath
        
        # Populate each target sheet with COMPLETE FIXED formatting
        Write-LogMessage -Message "COMPLETE FIXED: Populating Users sheet..." -Type Info -LogOnly
        Update-UsersSheet -Excel $excel -TenantData $TenantData
        
        Write-LogMessage -Message "COMPLETE FIXED: Populating Licensing sheet..." -Type Info -LogOnly
        Update-LicensingSheet -Excel $excel -TenantData $TenantData
        
        Write-LogMessage -Message "COMPLETE FIXED: Populating Hardware Profiles sheet with detailed settings..." -Type Info -LogOnly
        Update-HardwareProfilesSheet-Fixed -Excel $excel -TenantData $TenantData
        
        Write-LogMessage -Message "COMPLETE FIXED: Populating Windows Updates sheet..." -Type Info -LogOnly
        Update-WindowsUpdatesSheet -Excel $excel -TenantData $TenantData
        
        Write-LogMessage -Message "COMPLETE FIXED: Populating Intune Apps sheets with corrected data..." -Type Info -LogOnly
        Update-IntuneAppsSheets-Fixed -Excel $excel -TenantData $TenantData
        
        Write-LogMessage -Message "COMPLETE FIXED: Populating SharePoint Libraries sheet..." -Type Info -LogOnly
        Update-SharePointLibrariesSheet -Excel $excel -TenantData $TenantData
        
        Write-LogMessage -Message "COMPLETE FIXED: Populating Shared Mailboxes sheet..." -Type Info -LogOnly
        Update-SharedMailboxesSheet -Excel $excel -TenantData $TenantData
        
        Write-LogMessage -Message "COMPLETE FIXED: Populating Distribution Lists sheet..." -Type Info -LogOnly
        Update-DistributionListsSheet -Excel $excel -TenantData $TenantData
        
        Write-LogMessage -Message "COMPLETE FIXED: Populating Conditional Access sheet with detailed policy settings..." -Type Info -LogOnly
        Update-ConditionalAccessSheet-Fixed -Excel $excel -TenantData $TenantData
        
        # Save and close
        Close-ExcelPackage -ExcelPackage $excel -SaveAs $outputPath
        
        Write-LogMessage -Message "COMPLETE FIXED: Enhanced Excel documentation generated: $outputPath" -Type Success
        return $true
    }
    catch {
        Write-LogMessage -Message "COMPLETE FIXED: Failed to generate Excel documentation: $($_.Exception.Message)" -Type Error
        if ($excel) {
            Close-ExcelPackage -ExcelPackage $excel -NoSave
        }
        return $false
    }
}

# === COMPLETE FIXED: Sheet Population Functions ===

function Update-UsersSheet {
    <#
    .SYNOPSIS
        Populates the Users sheet with actual user data
    #>
    param (
        [Parameter(Mandatory = $true)]
        $Excel,
        
        [Parameter(Mandatory = $true)]
        [hashtable]$TenantData
    )
    
    try {
        $worksheet = $Excel.Workbook.Worksheets["Users"]
        if (-not $worksheet) {
            Write-LogMessage -Message "Users worksheet not found in template" -Type Warning -LogOnly
            return
        }
        
        $startRow = 7  # Data starts at row 7 based on template structure
        $currentRow = $startRow
        
        foreach ($user in $TenantData.Users.Users) {
            # Extract name parts
            $firstName = Get-SafeString -Value $user.GivenName
            $lastName = Get-SafeString -Value $user.Surname
            
            # Populate user data
            $worksheet.Cells[$currentRow, 1].Value = $firstName                    # Column A: First Name
            $worksheet.Cells[$currentRow, 2].Value = $lastName                     # Column B: Last Name
            $worksheet.Cells[$currentRow, 3].Value = $user.UserPrincipalName       # Column C: Email
            $worksheet.Cells[$currentRow, 4].Value = Get-SafeString -Value $user.JobTitle        # Column D: Job Title
            # Manager email would need to be resolved from manager ID
            $worksheet.Cells[$currentRow, 6].Value = Get-SafeString -Value $user.Department      # Column F: Department
            $worksheet.Cells[$currentRow, 7].Value = Get-SafeString -Value $user.Office          # Column G: Office location
            # Phone number would be in additional properties
            
            $currentRow++
            
            # Limit to prevent performance issues
            if ($currentRow -gt ($startRow + 500)) {
                Write-LogMessage -Message "Limited users export to first 500 users" -Type Warning -LogOnly
                break
            }
        }
        
        Write-LogMessage -Message "Updated Users sheet with $($currentRow - $startRow) users" -Type Success -LogOnly
    }
    catch {
        Write-LogMessage -Message "Error updating Users sheet: $($_.Exception.Message)" -Type Warning -LogOnly
    }
}

function Update-LicensingSheet {
    <#
    .SYNOPSIS
        Populates the Licensing sheet with user license assignments (licensed users only)
    #>
    param (
        [Parameter(Mandatory = $true)]
        $Excel,
        
        [Parameter(Mandatory = $true)]
        [hashtable]$TenantData
    )
    
    try {
        $worksheet = $Excel.Workbook.Worksheets["Licensing"]
        if (-not $worksheet) {
            Write-LogMessage -Message "Licensing worksheet not found in template" -Type Warning -LogOnly
            return
        }
        
        $startRow = 8  # Data starts at row 8 based on template structure
        $currentRow = $startRow
        $licensedUsersAdded = 0
        
        Write-LogMessage -Message "LICENSING: Starting to process users for licensing sheet" -Type Info -LogOnly
        
        foreach ($user in $TenantData.Users.Users) {
            # STRICT FILTERING: ONLY include users who have actual licenses assigned
            $validLicenses = @()
            
            if ($user.AssignedLicenses -and $user.AssignedLicenses.Count -gt 0) {
                foreach ($license in $user.AssignedLicenses) {
                    # Filter out all invalid license values
                    if ($license -and 
                        $license -ne 0 -and 
                        $license -ne "" -and 
                        $license -ne "No License Assigned" -and
                        $license.ToString().Trim() -ne "" -and
                        $license.ToString() -ne "0" -and
                        $license.ToString() -ne "null") {
                        $validLicenses += $license.ToString().Trim()
                    }
                }
            }
            
            # ONLY add users with valid licenses - COMPLETELY EXCLUDE unlicensed users
            if ($validLicenses.Count -gt 0) {
                # EXPLICIT COLUMN MAPPING - PowerShell ImportExcel uses 1-based indexing
                # Template expects: B, D, F, H columns for User, License1, License2, License3
                
                $worksheet.Cells[$currentRow, 2].Value = $user.DisplayName                    # Column B (index 2): User Name
                $worksheet.Cells[$currentRow, 4].Value = $validLicenses[0]                    # Column D (index 4): Base License Type
                
                # Clear columns C, E, G to ensure no data bleeds through
                $worksheet.Cells[$currentRow, 3].Value = ""                                   # Clear Column C
                $worksheet.Cells[$currentRow, 5].Value = ""                                   # Clear Column E
                $worksheet.Cells[$currentRow, 7].Value = ""                                   # Clear Column G
                
                # Additional licenses in correct columns
                if ($validLicenses.Count -gt 1) {
                    $worksheet.Cells[$currentRow, 6].Value = $validLicenses[1]                # Column F (index 6): Additional Software 1
                }
                if ($validLicenses.Count -gt 2) {
                    $worksheet.Cells[$currentRow, 8].Value = $validLicenses[2]                # Column H (index 8): Additional Software 2
                }
                
                $currentRow++
                $licensedUsersAdded++
                
                # Limit to prevent performance issues
                if ($licensedUsersAdded -ge 500) {
                    break
                }
            }
        }
        
        Write-LogMessage -Message "LICENSING COMPLETE: Updated Licensing sheet with $licensedUsersAdded licensed users (unlicensed users completely excluded)" -Type Success -LogOnly
        
        if ($licensedUsersAdded -eq 0) {
            # Add a note if no licensed users found
            $worksheet.Cells[$startRow, 2].Value = "No users with valid licenses found in tenant"
            $worksheet.Cells[$startRow, 4].Value = "Check license assignments"
            Write-LogMessage -Message "WARNING: No users with valid licenses were found!" -Type Warning
        }
    }
    catch {
        Write-LogMessage -Message "Error updating Licensing sheet: $($_.Exception.Message)" -Type Error
    }
}

function Update-ConditionalAccessSheet-Fixed {
    <#
    .SYNOPSIS
        COMPLETE FIXED: Populates CA sheet with properly formatted policy settings in table structure
    .DESCRIPTION
        Uses the detailed CA policy data to create comprehensive policy descriptions in Excel format
    #>
    param (
        [Parameter(Mandatory = $true)]
        $Excel,
        
        [Parameter(Mandatory = $true)]
        [hashtable]$TenantData
    )
    
    try {
        $worksheet = $Excel.Workbook.Worksheets["Conditional Access"]
        if (-not $worksheet) {
            Write-LogMessage -Message "FIXED CA: Conditional Access worksheet not found in template" -Type Warning -LogOnly
            return
        }
        
        # FIXED: Proper table structure for CA policies
        # Based on template analysis, we need to place data in the correct columns
        $startRow = 13  # Start after the template entries
        $currentRow = $startRow
        
        Write-LogMessage -Message "FIXED CA: Processing $($TenantData.ConditionalAccess.Policies.Count) CA policies with comprehensive detailed settings" -Type Info
        
        foreach ($policy in $TenantData.ConditionalAccess.Policies) {
            Write-LogMessage -Message "FIXED CA: Processing policy '$($policy.DisplayName)' for Excel formatting..." -Type Info -LogOnly
            
            # FIXED: Column B = Policy Name, Column D = Policy Settings (comprehensively formatted)
            $worksheet.Cells[$currentRow, 2].Value = $policy.DisplayName  # Column B: Policy Name
            
            # FIXED: Create comprehensive policy settings description for Column D
            $policySettingsParts = @()
            
            # Policy state (always show this first)
            $policySettingsParts += "State: $($policy.State)"
            
            # User conditions (detailed)
            if ($policy.UserConditions -and $policy.UserConditions.Trim() -ne "") {
                $policySettingsParts += "Users: $($policy.UserConditions)"
            }
            
            # Application conditions (detailed)
            if ($policy.ApplicationConditions -and $policy.ApplicationConditions.Trim() -ne "") {
                $policySettingsParts += "Applications: $($policy.ApplicationConditions)"
            }
            
            # Platform conditions
            if ($policy.PlatformConditions -and $policy.PlatformConditions.Trim() -ne "") {
                $policySettingsParts += "Platforms: $($policy.PlatformConditions)"
            }
            
            # Location conditions
            if ($policy.LocationConditions -and $policy.LocationConditions.Trim() -ne "") {
                $policySettingsParts += "Locations: $($policy.LocationConditions)"
            }
            
            # Grant controls (detailed)
            if ($policy.GrantControls -and $policy.GrantControls.Trim() -ne "") {
                $policySettingsParts += "Grant Controls: $($policy.GrantControls)"
            }
            
            # Session controls
            if ($policy.SessionControls -and $policy.SessionControls.Trim() -ne "") {
                $policySettingsParts += "Session Controls: $($policy.SessionControls)"
            }
            
            # Risk conditions
            if ($policy.RiskConditions -and $policy.RiskConditions.Trim() -ne "") {
                $policySettingsParts += "Risk Conditions: $($policy.RiskConditions)"
            }
            
            # FIXED: Join settings with line breaks for better readability in Excel
            if ($policySettingsParts.Count -gt 0) {
                $formattedSettings = $policySettingsParts -join "`n"
            } else {
                $formattedSettings = "Policy configured but settings could not be extracted"
            }
            
            $worksheet.Cells[$currentRow, 4].Value = $formattedSettings  # Column D: Policy Settings
            
            # FIXED: Clear columns C, E to prevent data bleed
            $worksheet.Cells[$currentRow, 3].Value = ""
            $worksheet.Cells[$currentRow, 5].Value = ""
            
            Write-LogMessage -Message "FIXED CA: Added policy '$($policy.DisplayName)' with comprehensive detailed settings" -Type Info -LogOnly
            $currentRow++
            
            # Limit entries to prevent template overflow
            if ($currentRow -gt ($startRow + 25)) { 
                Write-LogMessage -Message "FIXED CA: Reached limit of 25 policies for template space" -Type Info -LogOnly
                break 
            }
        }
        
        # FIXED: Add summary information if space allows
        if ($currentRow -le ($startRow + 20) -and $TenantData.ConditionalAccess.TotalCount -gt 0) {
            $currentRow++
            $worksheet.Cells[$currentRow, 2].Value = "=== CONDITIONAL ACCESS SUMMARY ==="
            $currentRow++
            $worksheet.Cells[$currentRow, 2].Value = "Total Policies"
            $worksheet.Cells[$currentRow, 4].Value = "$($TenantData.ConditionalAccess.TotalCount) total | $($TenantData.ConditionalAccess.EnabledCount) enabled | $($TenantData.ConditionalAccess.DisabledCount) disabled"
        }
        
        Write-LogMessage -Message "FIXED CA: Successfully updated Conditional Access sheet with $($currentRow - $startRow) policies and comprehensive detailed settings" -Type Success
    }
    catch {
        Write-LogMessage -Message "FIXED CA: Error updating Conditional Access sheet: $($_.Exception.Message)" -Type Error
        Write-LogMessage -Message "FIXED CA: Full error details: $($_.Exception.ToString())" -Type Error -LogOnly
    }
}

function Update-HardwareProfilesSheet-Fixed {
    <#
    .SYNOPSIS
        COMPLETE FIXED: Populates Hardware Profiles sheet with comprehensive Intune policy settings
    .DESCRIPTION
        Shows detailed policy settings for both configuration and compliance policies
    #>
    param (
        [Parameter(Mandatory = $true)]
        $Excel,
        
        [Parameter(Mandatory = $true)]
        [hashtable]$TenantData
    )
    
    try {
        $worksheet = $Excel.Workbook.Worksheets["Hardware Profiles"]
        if (-not $worksheet) {
            Write-LogMessage -Message "FIXED HW: Hardware Profiles worksheet not found in template" -Type Warning -LogOnly
            return
        }
        
        # FIXED: Add comprehensive policy information after existing template content
        $configStartRow = 45  # Start after the existing template content
        $currentRow = $configStartRow
        
        # FIXED: Add header for configured policies
        $worksheet.Cells[$currentRow, 2].Value = "*** CONFIGURED DEVICE POLICIES WITH COMPREHENSIVE DETAILED SETTINGS ***"
        $currentRow++
        $currentRow++ # Skip a row
        
        # FIXED: Configuration Policies with comprehensive detailed settings
        if ($TenantData.Intune.DeviceConfigurationPolicies.Count -gt 0) {
            $worksheet.Cells[$currentRow, 2].Value = "=== DEVICE CONFIGURATION POLICIES ==="
            $currentRow++
            
            foreach ($policy in $TenantData.Intune.DeviceConfigurationPolicies) {
                Write-LogMessage -Message "FIXED HW: Adding configuration policy '$($policy.DisplayName)' with detailed settings..." -Type Info -LogOnly
                
                # FIXED: Policy name in Column B
                $worksheet.Cells[$currentRow, 2].Value = $policy.DisplayName
                
                # FIXED: Comprehensive detailed settings in Column D
                $detailedInfo = @()
                $detailedInfo += "Type: $($policy.PolicyType)"
                $detailedInfo += "Platform: $($policy.Platform)"
                if ($policy.Description -and $policy.Description.Trim() -ne "") {
                    $detailedInfo += "Description: $($policy.Description)"
                }
                if ($policy.DetailedSettings -and $policy.DetailedSettings.Trim() -ne "") {
                    $detailedInfo += "Settings: $($policy.DetailedSettings)"
                }
                $detailedInfo += "Created: $($policy.CreatedDateTime)"
                $detailedInfo += "Modified: $($policy.LastModifiedDateTime)"
                
                $worksheet.Cells[$currentRow, 4].Value = $detailedInfo -join "`n"
                
                # FIXED: Status in Column F
                $worksheet.Cells[$currentRow, 6].Value = "Applied"
                
                $currentRow++
                
                # FIXED: Limit entries to prevent overwriting template
                if ($currentRow -gt ($configStartRow + 25)) { break }
            }
            
            Write-LogMessage -Message "FIXED HW: Added $($TenantData.Intune.DeviceConfigurationPolicies.Count) configuration policies with comprehensive detailed settings" -Type Success -LogOnly
        } else {
            $worksheet.Cells[$currentRow, 2].Value = "No configuration policies found"
            $worksheet.Cells[$currentRow, 4].Value = "Check Intune permissions and policy deployment"
            $currentRow++
        }
        
        # FIXED: Add spacing
        $currentRow++
        $currentRow++
        
        # FIXED: Compliance Policies with comprehensive detailed settings
        if ($TenantData.Intune.DeviceCompliancePolicies.Count -gt 0) {
            $worksheet.Cells[$currentRow, 2].Value = "=== DEVICE COMPLIANCE POLICIES ==="
            $currentRow++
            
            foreach ($policy in $TenantData.Intune.DeviceCompliancePolicies) {
                Write-LogMessage -Message "FIXED HW: Adding compliance policy '$($policy.DisplayName)' with detailed settings..." -Type Info -LogOnly
                
                # FIXED: Policy name in Column B
                $worksheet.Cells[$currentRow, 2].Value = $policy.DisplayName
                
                # FIXED: Comprehensive detailed settings in Column D
                $detailedInfo = @()
                $detailedInfo += "Type: $($policy.PolicyType)"
                $detailedInfo += "Platform: $($policy.Platform)"
                if ($policy.Description -and $policy.Description.Trim() -ne "") {
                    $detailedInfo += "Description: $($policy.Description)"
                }
                if ($policy.DetailedSettings -and $policy.DetailedSettings.Trim() -ne "") {
                    $detailedInfo += "Settings: $($policy.DetailedSettings)"
                }
                $detailedInfo += "Created: $($policy.CreatedDateTime)"
                $detailedInfo += "Modified: $($policy.LastModifiedDateTime)"
                
                $worksheet.Cells[$currentRow, 4].Value = $detailedInfo -join "`n"
                
                # FIXED: Status in Column F
                $worksheet.Cells[$currentRow, 6].Value = "Applied"
                
                $currentRow++
                
                # FIXED: Limit entries to prevent overwriting template  
                if ($currentRow -gt ($configStartRow + 50)) { break }
            }
            
            Write-LogMessage -Message "FIXED HW: Added $($TenantData.Intune.DeviceCompliancePolicies.Count) compliance policies with comprehensive detailed settings" -Type Success -LogOnly
        } else {
            $worksheet.Cells[$currentRow, 2].Value = "No compliance policies found"
            $worksheet.Cells[$currentRow, 4].Value = "Check Intune permissions and policy deployment"
            $currentRow++
        }
        
        # FIXED: Add comprehensive summary and access status
        $currentRow++
        $currentRow++
        $worksheet.Cells[$currentRow, 2].Value = "=== INTUNE ACCESS STATUS & SUMMARY ==="
        $currentRow++
        $worksheet.Cells[$currentRow, 2].Value = "Configuration Policies Access:"
        $worksheet.Cells[$currentRow, 4].Value = $TenantData.Intune.AccessStatus.ConfigurationPolicies
        $currentRow++
        $worksheet.Cells[$currentRow, 2].Value = "Compliance Policies Access:"
        $worksheet.Cells[$currentRow, 4].Value = $TenantData.Intune.AccessStatus.CompliancePolicies
        $currentRow++
        $worksheet.Cells[$currentRow, 2].Value = "Managed Apps Access:"
        $worksheet.Cells[$currentRow, 4].Value = $TenantData.Intune.AccessStatus.ManagedApps
        $currentRow++
        $worksheet.Cells[$currentRow, 2].Value = "Managed Devices Access:"
        $worksheet.Cells[$currentRow, 4].Value = $TenantData.Intune.AccessStatus.ManagedDevices
        $currentRow++
        $currentRow++
        $worksheet.Cells[$currentRow, 2].Value = "Total Configuration Policies: $($TenantData.Intune.DeviceConfigurationPolicies.Count)"
        $currentRow++
        $worksheet.Cells[$currentRow, 2].Value = "Total Compliance Policies: $($TenantData.Intune.DeviceCompliancePolicies.Count)"
        $currentRow++
        $worksheet.Cells[$currentRow, 2].Value = "Total Managed Devices: $($TenantData.Intune.TotalDevices)"
        $currentRow++
        $worksheet.Cells[$currentRow, 2].Value = "Total Managed Apps: $($TenantData.Intune.ManagedApps.Count)"
        $currentRow++
        $currentRow++
        $worksheet.Cells[$currentRow, 2].Value = "Note: The template settings above show recommended configuration options."
        $currentRow++
        $worksheet.Cells[$currentRow, 2].Value = "The policies listed here are the actual Intune policies currently configured"
        $currentRow++
        $worksheet.Cells[$currentRow, 2].Value = "with their comprehensive detailed settings and current application status."
        
        Write-LogMessage -Message "FIXED HW: Successfully updated Hardware Profiles sheet with comprehensive detailed policy settings for $(($TenantData.Intune.DeviceConfigurationPolicies.Count + $TenantData.Intune.DeviceCompliancePolicies.Count)) policies" -Type Success
    }
    catch {
        Write-LogMessage -Message "FIXED HW: Error updating Hardware Profiles sheet: $($_.Exception.Message)" -Type Error
        Write-LogMessage -Message "FIXED HW: Full error details: $($_.Exception.ToString())" -Type Error -LogOnly
    }
}

function Update-IntuneAppsSheets-Fixed {
    <#
    .SYNOPSIS
        COMPLETE FIXED: Populates all Intune Apps sheets with corrected app data and enhanced formatting
    .DESCRIPTION
        Uses the corrected app collection data with proper platform filtering and detailed information
    #>
    param (
        [Parameter(Mandatory = $true)]
        $Excel,
        
        [Parameter(Mandatory = $true)]
        [hashtable]$TenantData
    )
    
    try {
        Write-LogMessage -Message "COMPLETE FIXED APPS: Starting to populate Intune app sheets with corrected data and enhanced formatting" -Type Info
        Write-LogMessage -Message "COMPLETE FIXED APPS: Available managed apps count: $($TenantData.Intune.ManagedApps.Count)" -Type Info
        Write-LogMessage -Message "COMPLETE FIXED APPS: Apps access status: $($TenantData.Intune.AccessStatus.ManagedApps)" -Type Info
        
        # FIXED: Log app details for debugging
        if ($TenantData.Intune.ManagedApps -and $TenantData.Intune.ManagedApps.Count -gt 0) {
            $appSample = $TenantData.Intune.ManagedApps | Select-Object -First 3
            foreach ($app in $appSample) {
                Write-LogMessage -Message "COMPLETE FIXED APPS: Sample app - Name: '$($app.DisplayName)', Type: '$($app.AppType)', Platform: '$($app.Platform)', Method: '$($app.CollectionMethod)'" -Type Info -LogOnly
            }
        }
        
        # FIXED: Correct sheet names with exact template matching
        $appSheets = @(
            @{ Name = "Intune Windows Apps"; Platform = "Windows"; Filter = @("Windows", "Win32", "MSI", "windows") },
            @{ Name = "Intune Android Apps"; Platform = "Android"; Filter = @("Android", "android") }, 
            @{ Name = "Intune Apple IOS Apps "; Platform = "iOS"; Filter = @("iOS", "ios") },
            @{ Name = "Intune Apple iPadOS Apps "; Platform = "iPadOS"; Filter = @("iPadOS", "iOS", "ios") },
            @{ Name = "Intune Mac OS Apps"; Platform = "macOS"; Filter = @("macOS", "mac") }
        )
        
        foreach ($sheetInfo in $appSheets) {
            $sheetName = $sheetInfo.Name
            $platform = $sheetInfo.Platform
            $filters = $sheetInfo.Filter
            
            Write-LogMessage -Message "COMPLETE FIXED APPS: Processing sheet '$sheetName' for platform '$platform'" -Type Info
            
            $worksheet = $Excel.Workbook.Worksheets[$sheetName]
            if ($worksheet) {
                Write-LogMessage -Message "COMPLETE FIXED APPS: Found worksheet '$sheetName'" -Type Info -LogOnly
                
                $startRow = 8  # Start after headers
                $currentRow = $startRow
                $appsAdded = 0
                
                # FIXED: Clear the starting rows first
                for ($i = $startRow; $i -le ($startRow + 20); $i++) {
                    $worksheet.Cells[$i, 2].Value = ""
                    $worksheet.Cells[$i, 4].Value = ""
                    $worksheet.Cells[$i, 6].Value = ""
                }
                
                # FIXED: Add apps with enhanced filtering and information
                if ($TenantData.Intune.ManagedApps -and $TenantData.Intune.ManagedApps.Count -gt 0) {
                    
                    # FIXED: Filter apps for this platform (enhanced filtering)
                    $platformApps = $TenantData.Intune.ManagedApps | Where-Object {
                        $app = $_
                        $matchesPlatform = $false
                        
                        # Check if app platform or type matches any of the platform filters
                        foreach ($filter in $filters) {
                            if ($app.Platform -like "*$filter*" -or 
                                $app.AppType -like "*$filter*" -or 
                                $app.DisplayName -like "*$filter*") {
                                $matchesPlatform = $true
                                break
                            }
                        }
                        
                        # For comprehensive view, also include apps if no specific platform is identified
                        if (-not $matchesPlatform -and ($app.Platform -eq "Unknown" -or $app.Platform -eq "Multi-platform")) {
                            $matchesPlatform = $true
                        }
                        
                        return $matchesPlatform
                    }
                    
                    # If no platform-specific apps found, include all apps for visibility
                    if ($platformApps.Count -eq 0) {
                        $platformApps = $TenantData.Intune.ManagedApps
                        Write-LogMessage -Message "COMPLETE FIXED APPS: No platform-specific apps found for '$platform', including all apps for visibility" -Type Info -LogOnly
                    }
                    
                    Write-LogMessage -Message "COMPLETE FIXED APPS: Found $($platformApps.Count) apps for platform '$platform'" -Type Info
                    
                    foreach ($app in $platformApps) {
                        Write-LogMessage -Message "COMPLETE FIXED APPS: Adding app '$($app.DisplayName)' (Type: $($app.AppType), Platform: $($app.Platform)) to sheet '$sheetName'" -Type Info -LogOnly
                        
                        # FIXED: Proper column structure with enhanced information
                        $worksheet.Cells[$currentRow, 2].Value = $app.DisplayName          # Column B: Application Name
                        $worksheet.Cells[$currentRow, 4].Value = "X"                      # Column D: Required (default)
                        
                        # FIXED: Add comprehensive app information in Column F
                        $appInfo = @()
                        $appInfo += "Type: $($app.AppType)"
                        $appInfo += "Platform: $($app.Platform)"
                        if ($app.Publisher -and $app.Publisher -ne "Unknown Publisher") {
                            $appInfo += "Publisher: $($app.Publisher)"
                        }
                        if ($app.CollectionMethod) {
                            $appInfo += "Source: $($app.CollectionMethod)"
                        }
                        
                        $worksheet.Cells[$currentRow, 6].Value = $appInfo -join " | "     # Column F: App Info
                        
                        # FIXED: Clear other columns to prevent bleed
                        $worksheet.Cells[$currentRow, 3].Value = ""                       # Clear Column C
                        $worksheet.Cells[$currentRow, 5].Value = ""                       # Clear Column E
                        $worksheet.Cells[$currentRow, 7].Value = ""                       # Clear Column G
                        $worksheet.Cells[$currentRow, 8].Value = ""                       # Clear Column H
                        
                        $currentRow++
                        $appsAdded++
                        
                        # FIXED: Limit entries per sheet
                        if ($appsAdded -ge 15) { 
                            Write-LogMessage -Message "COMPLETE FIXED APPS: Reached limit of 15 apps for sheet '$sheetName'" -Type Info -LogOnly
                            break 
                        }
                    }
                    
                    Write-LogMessage -Message "COMPLETE FIXED APPS: Successfully added $appsAdded apps to '$sheetName'" -Type Success
                } else {
                    # FIXED: Enhanced messaging when no apps found with access status
                    $worksheet.Cells[$startRow, 2].Value = "No managed apps found for $platform"
                    $worksheet.Cells[$startRow, 4].Value = "Access Status: $($TenantData.Intune.AccessStatus.ManagedApps)"
                    $worksheet.Cells[$startRow, 6].Value = "Requires DeviceManagementApps.Read.All permission"
                    Write-LogMessage -Message "COMPLETE FIXED APPS: No managed apps available for '$sheetName' - Access Status: $($TenantData.Intune.AccessStatus.ManagedApps)" -Type Warning
                }
                
                # FIXED: Add summary row with access status if apps were added
                if ($appsAdded -gt 0) {
                    $currentRow++
                    $worksheet.Cells[$currentRow, 2].Value = "=== SUMMARY ==="
                    $worksheet.Cells[$currentRow, 4].Value = "$appsAdded apps configured for $platform"
                    $worksheet.Cells[$currentRow, 6].Value = "Access: $($TenantData.Intune.AccessStatus.ManagedApps)"
                } elseif ($TenantData.Intune.ManagedApps.Count -eq 0) {
                    # Add access status information
                    $currentRow++
                    $worksheet.Cells[$currentRow, 2].Value = "=== ACCESS STATUS ==="
                    $worksheet.Cells[$currentRow, 4].Value = $TenantData.Intune.AccessStatus.ManagedApps
                    $worksheet.Cells[$currentRow, 6].Value = "Check Graph API permissions for app management"
                }
                
            } else {
                Write-LogMessage -Message "COMPLETE FIXED APPS: Sheet '$sheetName' not found in template" -Type Warning
            }
        }
        
        Write-LogMessage -Message "COMPLETE FIXED APPS: Completed processing all Intune app sheets with comprehensive enhanced formatting" -Type Success
    }
    catch {
        Write-LogMessage -Message "COMPLETE FIXED APPS: Error updating Intune Apps sheets: $($_.Exception.Message)" -Type Error
        Write-LogMessage -Message "COMPLETE FIXED APPS: Full error details: $($_.Exception.ToString())" -Type Error -LogOnly
    }
}

function Update-WindowsUpdatesSheet {
    <#
    .SYNOPSIS
        Populates the Windows Updates sheet with current update configuration info
    #>
    param (
        [Parameter(Mandatory = $true)]
        $Excel,
        
        [Parameter(Mandatory = $true)]
        [hashtable]$TenantData
    )
    
    try {
        $worksheet = $Excel.Workbook.Worksheets["Windows Updates"]
        if (-not $worksheet) {
            Write-LogMessage -Message "Windows Updates worksheet not found in template" -Type Warning -LogOnly
            return
        }
        
        # Add information about current Windows Update configuration
        # The template has a complex structure, so we'll add our info in a safe area
        $infoRow = 45  # Add after template content
        
        $worksheet.Cells[$infoRow, 2].Value = "*** CURRENT WINDOWS UPDATE CONFIGURATION ***"
        $infoRow++
        $infoRow++
        
        # Check if we have any Windows update policies from Intune
        $updatePolicies = $TenantData.Intune.DeviceConfigurationPolicies | Where-Object { 
            $_.DisplayName -like "*update*" -or $_.DisplayName -like "*ring*" 
        }
        
        if ($updatePolicies.Count -gt 0) {
            $worksheet.Cells[$infoRow, 2].Value = "Windows Update Policies Configured:"
            $infoRow++
            
            foreach ($policy in $updatePolicies) {
                $worksheet.Cells[$infoRow, 2].Value = "• $($policy.DisplayName)"
                $worksheet.Cells[$infoRow, 4].Value = $policy.DetailedSettings ?? "Settings not available"
                $infoRow++
            }
        } else {
            $worksheet.Cells[$infoRow, 2].Value = "No specific Windows Update policies found in Intune"
            $infoRow++
            $worksheet.Cells[$infoRow, 2].Value = "Updates may be managed through default settings or other policies"
        }
        
        $infoRow++
        $infoRow++
        $worksheet.Cells[$infoRow, 2].Value = "Note: The template above shows recommended update ring configuration."
        $infoRow++
        $worksheet.Cells[$infoRow, 2].Value = "Actual update policies are listed here for reference."
        
        Write-LogMessage -Message "Updated Windows Updates sheet with current configuration info" -Type Success -LogOnly
    }
    catch {
        Write-LogMessage -Message "Error updating Windows Updates sheet: $($_.Exception.Message)" -Type Warning -LogOnly
    }
}

function Update-SharePointLibrariesSheet {
    <#
    .SYNOPSIS
        Populates the SharePoint Libraries sheet
    #>
    param (
        [Parameter(Mandatory = $true)]
        $Excel,
        
        [Parameter(Mandatory = $true)]
        [hashtable]$TenantData
    )
    
    try {
        $worksheet = $Excel.Workbook.Worksheets["SharePoint Libaries"]  # Note: matches template spelling
        if (-not $worksheet) {
            Write-LogMessage -Message "SharePoint Libraries worksheet not found in template (checked with trailing space)" -Type Warning -LogOnly
            return
        }
        
        $startRow = 13  # Start after template entries
        $currentRow = $startRow
        
        # Add actual SharePoint sites with access status information
        if ($TenantData.SharePoint.SiteCollections -and $TenantData.SharePoint.SiteCollections.Count -gt 0) {
            foreach ($site in $TenantData.SharePoint.SiteCollections) {
                if ($site.Id -eq "PERMISSION_ERROR" -or $site.Id -eq "CRITICAL_ERROR") {
                    # Handle permission/error cases
                    $worksheet.Cells[$currentRow, 2].Value = $site.DisplayName  # Column B: Site Name
                    $worksheet.Cells[$currentRow, 4].Value = "Access Error"     # Column D: Approver
                    $worksheet.Cells[$currentRow, 6].Value = $site.WebUrl       # Column F: Error details
                    $worksheet.Cells[$currentRow, 8].Value = $TenantData.SharePoint.AccessStatus  # Column H: Status
                } else {
                    # Handle normal sites
                    $worksheet.Cells[$currentRow, 2].Value = $site.DisplayName  # Column B: Site Name
                    $worksheet.Cells[$currentRow, 4].Value = "Site Admin"       # Column D: Approver
                    $worksheet.Cells[$currentRow, 6].Value = "Site Owners"      # Column F: Owners
                    $worksheet.Cells[$currentRow, 8].Value = "Site Members"     # Column H: Members
                }
                $currentRow++
                
                # Limit entries
                if ($currentRow -gt ($startRow + 20)) { break }
            }
            
            # Add access status information
            $currentRow++
            $worksheet.Cells[$currentRow, 2].Value = "=== ACCESS STATUS ==="
            $worksheet.Cells[$currentRow, 4].Value = $TenantData.SharePoint.AccessStatus
            if ($TenantData.SharePoint.PermissionMessage) {
                $currentRow++
                $worksheet.Cells[$currentRow, 2].Value = "Permission Info:"
                $worksheet.Cells[$currentRow, 4].Value = $TenantData.SharePoint.PermissionMessage
            }
            
            Write-LogMessage -Message "Updated SharePoint Libraries sheet with $($currentRow - $startRow) sites and access status" -Type Success -LogOnly
        } else {
            # If no sites found, add a note with access status
            $worksheet.Cells[$startRow, 2].Value = "No SharePoint sites found or accessible"
            $worksheet.Cells[$startRow, 4].Value = $TenantData.SharePoint.AccessStatus
            $worksheet.Cells[$startRow, 6].Value = $TenantData.SharePoint.PermissionMessage ?? "Check permissions"
            Write-LogMessage -Message "No SharePoint sites found to populate - Status: $($TenantData.SharePoint.AccessStatus)" -Type Warning -LogOnly
        }
    }
    catch {
        Write-LogMessage -Message "Error updating SharePoint Libraries sheet: $($_.Exception.Message)" -Type Warning -LogOnly
    }
}

function Update-SharedMailboxesSheet {
    <#
    .SYNOPSIS
        Populates the Shared Mailboxes sheet with correct sheet name and column structure
    #>
    param (
        [Parameter(Mandatory = $true)]
        $Excel,
        
        [Parameter(Mandatory = $true)]
        [hashtable]$TenantData
    )
    
    try {
        # CORRECT sheet name with trailing space
        $worksheet = $Excel.Workbook.Worksheets["Shared Mailboxes "]  # Note: trailing space
        if (-not $worksheet) {
            Write-LogMessage -Message "Shared Mailboxes worksheet not found in template (checked with trailing space)" -Type Warning -LogOnly
            return
        }
        
        # Based on ACTUAL analysis: headers are B, C, D, E, F (not B, D, F, H as I thought)
        # Row 8: B="Shared Mailbox name", C="Approver", D="Read emails", E="Send email", F="Send on behalf"
        
        $startRow = 9  # Data starts after headers in row 8
        $currentRow = $startRow
        
        # Look for actual shared mailboxes (mail-enabled groups or specific types)
        $sharedMailboxes = @()
        
        # Check distribution groups that might be shared mailboxes
        if ($TenantData.Groups.DistributionGroups) {
            $sharedMailboxes += $TenantData.Groups.DistributionGroups | Where-Object { 
                $_.DisplayName -like "*shared*" -or 
                $_.DisplayName -like "*mailbox*" -or
                $_.DisplayName -like "*info*" -or
                $_.DisplayName -like "*support*" -or
                $_.DisplayName -like "*admin*" -or
                $_.DisplayName -like "*help*"
            }
        }
        
        # Check Microsoft 365 groups that might be shared mailboxes
        if ($TenantData.Groups.Microsoft365Groups) {
            $sharedMailboxes += $TenantData.Groups.Microsoft365Groups | Where-Object { 
                $_.DisplayName -like "*shared*" -or 
                $_.DisplayName -like "*mailbox*" -or
                $_.DisplayName -like "*info*" -or
                $_.DisplayName -like "*support*"
            }
        }
        
        if ($sharedMailboxes.Count -gt 0) {
            foreach ($mailbox in $sharedMailboxes) {
                $worksheet.Cells[$currentRow, 2].Value = $mailbox.DisplayName          # Column B: Shared Mailbox name
                $worksheet.Cells[$currentRow, 3].Value = "IT Administrator"            # Column C: Approver
                $worksheet.Cells[$currentRow, 4].Value = "See group members"          # Column D: Read emails
                $worksheet.Cells[$currentRow, 5].Value = "See group members"          # Column E: Send email
                $worksheet.Cells[$currentRow, 6].Value = "See group members"          # Column F: Send on behalf
                $currentRow++
                
                if ($currentRow -gt ($startRow + 10)) { break }
            }
            Write-LogMessage -Message "Updated Shared Mailboxes sheet with $($currentRow - $startRow) mailboxes" -Type Success -LogOnly
        } else {
            # If no shared mailboxes found, add a note
            $worksheet.Cells[$startRow, 2].Value = "No shared mailboxes found"
            $worksheet.Cells[$startRow, 3].Value = "N/A"
            $worksheet.Cells[$startRow, 4].Value = "Check Exchange configuration"
            Write-LogMessage -Message "No shared mailboxes found" -Type Warning -LogOnly
        }
    }
    catch {
        Write-LogMessage -Message "Error updating Shared Mailboxes sheet: $($_.Exception.Message)" -Type Warning -LogOnly
    }
}

function Update-DistributionListsSheet {
    <#
    .SYNOPSIS
        Populates the Distribution Lists sheet with proper table formatting and correct columns
    #>
    param (
        [Parameter(Mandatory = $true)]
        $Excel,
        
        [Parameter(Mandatory = $true)]
        [hashtable]$TenantData
    )
    
    try {
        $worksheet = $Excel.Workbook.Worksheets["Distribution list"]
        if (-not $worksheet) {
            Write-LogMessage -Message "Distribution Lists worksheet not found in template" -Type Warning -LogOnly
            return
        }
        
        # Based on template analysis, the correct structure is:
        # Row 8: Column B = "Distribution List Name", Column D = "Approver", Column F = "Members"
        # So data should go in columns B, D, F (not B, C, D!)
        
        $startRow = 9  # Data starts after headers in row 8
        $currentRow = $startRow
        
        foreach ($group in $TenantData.Groups.DistributionGroups) {
            $worksheet.Cells[$currentRow, 2].Value = $group.DisplayName                          # Column B: Distribution List Name
            $worksheet.Cells[$currentRow, 4].Value = "IT Administrator"                          # Column D: Approver  
            $worksheet.Cells[$currentRow, 6].Value = "$(Get-SafeString -Value $group.MemberCount -DefaultValue 'Unknown') members"  # Column F: Members info
            $currentRow++
            
            if ($currentRow -gt ($startRow + 20)) { break }
        }
        
        Write-LogMessage -Message "Updated Distribution Lists sheet with $($currentRow - $startRow) groups using correct column structure (B, D, F)" -Type Success -LogOnly
    }
    catch {
        Write-LogMessage -Message "Error updating Distribution Lists sheet: $($_.Exception.Message)" -Type Warning -LogOnly
    }
}

# === Helper Functions ===

function Get-SafeString {
    <#
    .SYNOPSIS
        Safely converts values to strings with default fallback
    #>
    param(
        [Parameter(Mandatory = $false)]
        $Value,
        [Parameter(Mandatory = $false)]
        [string]$DefaultValue = ""
    )
    
    if ($null -eq $Value -or $Value -eq "") {
        return $DefaultValue
    }
    return $Value.ToString().Trim()
}

# === Supplementary Documentation Functions ===

function New-HTMLDocumentation {
    <#
    .SYNOPSIS
        Generates comprehensive HTML documentation report
    #>
    param (
        [Parameter(Mandatory = $true)]
        [hashtable]$TenantData
    )
    
    try {
        $htmlPath = Join-Path -Path $DocumentationConfig.OutputDirectory -ChildPath "Reports\TenantConfiguration_Enhanced.html"
        
        $htmlContent = @"
<!DOCTYPE html>
<html>
<head>
    <title>Microsoft 365 Tenant Configuration Report - Enhanced</title>
    <meta charset="UTF-8">
    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 20px; background-color: #f5f5f5; }
        .container { max-width: 1200px; margin: 0 auto; background-color: white; padding: 20px; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
        h1 { color: #0078d4; border-bottom: 3px solid #0078d4; padding-bottom: 10px; }
        h2 { color: #106ebe; border-left: 4px solid #0078d4; padding-left: 10px; margin-top: 30px; }
        h3 { color: #323130; }
        .info-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(300px, 1fr)); gap: 20px; margin: 20px 0; }
        .info-card { background-color: #f8f9fa; border: 1px solid #e9ecef; border-radius: 6px; padding: 15px; }
        .info-card h4 { margin-top: 0; color: #0078d4; }
        table { width: 100%; border-collapse: collapse; margin: 15px 0; }
        th, td { border: 1px solid #ddd; padding: 12px; text-align: left; }
        th { background-color: #0078d4; color: white; }
        tr:nth-child(even) { background-color: #f2f2f2; }
        .status-enabled { color: #107c10; font-weight: bold; }
        .status-disabled { color: #d13438; font-weight: bold; }
        .status-error { color: #d13438; font-style: italic; }
        .timestamp { color: #666; font-size: 0.9em; }
        .summary-stats { display: flex; justify-content: space-around; margin: 20px 0; }
        .stat-box { text-align: center; padding: 15px; background-color: #e3f2fd; border-radius: 6px; }
        .stat-number { font-size: 2em; font-weight: bold; color: #0078d4; }
        .stat-label { color: #666; }
        .access-status { padding: 10px; margin: 10px 0; border-radius: 4px; }
        .access-success { background-color: #dff0d8; border: 1px solid #d6e9c6; color: #3c763d; }
        .access-warning { background-color: #fcf8e3; border: 1px solid #faebcc; color: #8a6d3b; }
        .access-error { background-color: #f2dede; border: 1px solid #ebccd1; color: #a94442; }
        .policy-details { background-color: #f9f9f9; padding: 10px; margin: 5px 0; border-radius: 4px; font-size: 0.9em; }
        .collapsible { background-color: #777; color: white; cursor: pointer; padding: 18px; width: 100%; border: none; text-align: left; outline: none; font-size: 15px; }
        .collapsible:hover { background-color: #555; }
        .content { padding: 0 18px; display: none; overflow: hidden; background-color: #f1f1f1; }
    </style>
    <script>
        function toggleCollapsible(element) {
            element.classList.toggle("active");
            var content = element.nextElementSibling;
            if (content.style.display === "block") {
                content.style.display = "none";
            } else {
                content.style.display = "block";
            }
        }
    </script>
</head>
<body>
    <div class="container">
        <h1>Microsoft 365 Tenant Configuration Report - Enhanced</h1>
        <p class="timestamp">Generated on: $($TenantData.GeneratedOn.ToString('yyyy-MM-dd HH:mm:ss'))</p>
        
        <h2>Tenant Information</h2>
        <div class="info-grid">
            <div class="info-card">
                <h4>Basic Information</h4>
                <p><strong>Tenant Name:</strong> $($TenantData.TenantInfo.DisplayName)</p>
                <p><strong>Default Domain:</strong> $($TenantData.TenantInfo.DefaultDomain)</p>
                <p><strong>Tenant ID:</strong> $($TenantData.TenantInfo.TenantId)</p>
                <p><strong>Country:</strong> $($TenantData.TenantInfo.CountryCode)</p>
            </div>
            <div class="info-card">
                <h4>Setup Information</h4>
                <p><strong>Connected As:</strong> $($TenantData.TenantInfo.ConnectedAs)</p>
                <p><strong>Setup Date:</strong> $($TenantData.TenantInfo.SetupDate)</p>
                <p><strong>Admin Email:</strong> $($TenantData.TenantInfo.AdminEmail)</p>
            </div>
        </div>
        
        <h2>Summary Statistics</h2>
        <div class="summary-stats">
            <div class="stat-box">
                <div class="stat-number">$($TenantData.Groups.TotalCount)</div>
                <div class="stat-label">Total Groups</div>
            </div>
            <div class="stat-box">
                <div class="stat-number">$($TenantData.Users.TotalUsers)</div>
                <div class="stat-label">Total Users</div>
            </div>
            <div class="stat-box">
                <div class="stat-number">$($TenantData.ConditionalAccess.TotalCount)</div>
                <div class="stat-label">CA Policies</div>
            </div>
            <div class="stat-box">
                <div class="stat-number">$($TenantData.Intune.TotalDevices)</div>
                <div class="stat-label">Managed Devices</div>
            </div>
        </div>
        
        <h2>Service Access Status</h2>
        <div class="access-status $(if($TenantData.SharePoint.AccessStatus -like "*Success*") {'access-success'} elseif($TenantData.SharePoint.AccessStatus -like "*Error*") {'access-error'} else {'access-warning'})">
            <strong>SharePoint Access:</strong> $($TenantData.SharePoint.AccessStatus)
            $(if($TenantData.SharePoint.PermissionMessage) { "<br><em>$($TenantData.SharePoint.PermissionMessage)</em>" })
        </div>
        
        <div class="access-status $(if($TenantData.Intune.AccessStatus.ManagedApps -like "*Success*") {'access-success'} elseif($TenantData.Intune.AccessStatus.ManagedApps -like "*Failed*") {'access-error'} else {'access-warning'})">
            <strong>Intune Apps Access:</strong> $($TenantData.Intune.AccessStatus.ManagedApps)
        </div>
        
        <div class="access-status $(if($TenantData.Intune.AccessStatus.ConfigurationPolicies -like "*Success*") {'access-success'} elseif($TenantData.Intune.AccessStatus.ConfigurationPolicies -like "*Failed*") {'access-error'} else {'access-warning'})">
            <strong>Intune Configuration Policies:</strong> $($TenantData.Intune.AccessStatus.ConfigurationPolicies)
        </div>
        
        <div class="access-status $(if($TenantData.Intune.AccessStatus.CompliancePolicies -like "*Success*") {'access-success'} elseif($TenantData.Intune.AccessStatus.CompliancePolicies -like "*Failed*") {'access-error'} else {'access-warning'})">
            <strong>Intune Compliance Policies:</strong> $($TenantData.Intune.AccessStatus.CompliancePolicies)
        </div>
        
        <h2>Conditional Access Policies</h2>
        <button type="button" class="collapsible" onclick="toggleCollapsible(this)">Show Conditional Access Policy Details ($($TenantData.ConditionalAccess.TotalCount) policies)</button>
        <div class="content">
"@

        # Add CA Policy details if available
        if ($TenantData.ConditionalAccess.Policies.Count -gt 0) {
            foreach ($policy in $TenantData.ConditionalAccess.Policies) {
                $statusClass = if ($policy.State -eq "enabled") { "status-enabled" } else { "status-disabled" }
                $htmlContent += @"
                <div class="policy-details">
                    <h4>$($policy.DisplayName) <span class="$statusClass">($($policy.State))</span></h4>
"@
                if ($policy.UserConditions) { $htmlContent += "<p><strong>Users:</strong> $($policy.UserConditions)</p>" }
                if ($policy.ApplicationConditions) { $htmlContent += "<p><strong>Applications:</strong> $($policy.ApplicationConditions)</p>" }
                if ($policy.GrantControls) { $htmlContent += "<p><strong>Grant Controls:</strong> $($policy.GrantControls)</p>" }
                if ($policy.PlatformConditions) { $htmlContent += "<p><strong>Platforms:</strong> $($policy.PlatformConditions)</p>" }
                if ($policy.RiskConditions) { $htmlContent += "<p><strong>Risk Conditions:</strong> $($policy.RiskConditions)</p>" }
                $htmlContent += "</div>"
            }
        } else {
            $htmlContent += "<p>No conditional access policies found.</p>"
        }

        $htmlContent += @"
        </div>
        
        <h2>Intune Configuration</h2>
        <button type="button" class="collapsible" onclick="toggleCollapsible(this)">Show Intune Policy Details</button>
        <div class="content">
            <h3>Device Configuration Policies ($($TenantData.Intune.DeviceConfigurationPolicies.Count))</h3>
"@

        # Add Intune Config Policy details
        if ($TenantData.Intune.DeviceConfigurationPolicies.Count -gt 0) {
            foreach ($policy in $TenantData.Intune.DeviceConfigurationPolicies) {
                $htmlContent += @"
                <div class="policy-details">
                    <h4>$($policy.DisplayName)</h4>
                    <p><strong>Type:</strong> $($policy.PolicyType)</p>
                    <p><strong>Platform:</strong> $($policy.Platform)</p>
                    <p><strong>Settings:</strong> $($policy.DetailedSettings)</p>
                    <p><strong>Created:</strong> $($policy.CreatedDateTime)</p>
                </div>
"@
            }
        }

        $htmlContent += @"
            <h3>Device Compliance Policies ($($TenantData.Intune.DeviceCompliancePolicies.Count))</h3>
"@

        # Add Intune Compliance Policy details
        if ($TenantData.Intune.DeviceCompliancePolicies.Count -gt 0) {
            foreach ($policy in $TenantData.Intune.DeviceCompliancePolicies) {
                $htmlContent += @"
                <div class="policy-details">
                    <h4>$($policy.DisplayName)</h4>
                    <p><strong>Type:</strong> $($policy.PolicyType)</p>
                    <p><strong>Platform:</strong> $($policy.Platform)</p>
                    <p><strong>Settings:</strong> $($policy.DetailedSettings)</p>
                    <p><strong>Created:</strong> $($policy.CreatedDateTime)</p>
                </div>
"@
            }
        }

        $htmlContent += @"
        </div>
        
        <footer style="margin-top: 40px; text-align: center; color: #666; border-top: 1px solid #ddd; padding-top: 20px;">
            <p>Generated by Microsoft 365 Tenant Setup Utility - Enhanced Version | $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
            <p>This report contains comprehensive tenant configuration details with enhanced access status information.</p>
        </footer>
    </div>
</body>
</html>
"@

        $htmlContent | Out-File -FilePath $htmlPath -Encoding UTF8
        Write-LogMessage -Message "Enhanced HTML documentation generated: $htmlPath" -Type Success -LogOnly
        return $true
    }
    catch {
        Write-LogMessage -Message "Failed to generate HTML documentation: $($_.Exception.Message)" -Type Error
        return $false
    }
}

function New-JSONDocumentation {
    <#
    .SYNOPSIS
        Generates JSON configuration export with enhanced formatting
    #>
    param (
        [Parameter(Mandatory = $true)]
        [hashtable]$TenantData
    )
    
    try {
        $jsonPath = Join-Path -Path $DocumentationConfig.OutputDirectory -ChildPath "Exports\TenantConfiguration_Enhanced.json"
        
        # Add metadata to JSON export
        $TenantData.ExportMetadata = @{
            ExportVersion = "1.1-Fixed"
            ExportDateTime = Get-Date
            ExportedBy = (Get-MgContext).Account
            FixesApplied = @(
                "Fixed SharePoint access with multiple fallback methods",
                "Fixed Intune mobile app cmdlets",
                "Fixed CA policy detailed settings extraction",
                "Fixed Intune policy detailed settings",
                "Enhanced error handling and logging",
                "Added comprehensive access status reporting"
            )
        }
        
        $TenantData | ConvertTo-Json -Depth 10 | Out-File -FilePath $jsonPath -Encoding UTF8
        Write-LogMessage -Message "Enhanced JSON configuration export generated: $jsonPath" -Type Success -LogOnly
        return $true
    }
    catch {
        Write-LogMessage -Message "Failed to generate JSON documentation: $($_.Exception.Message)" -Type Error
        return $false
    }
}

function New-ConfigurationSummary {
    <#
    .SYNOPSIS
        Generates a comprehensive text-based configuration summary
    #>
    param (
        [Parameter(Mandatory = $true)]
        [hashtable]$TenantData
    )
    
    try {
        $summaryPath = Join-Path -Path $DocumentationConfig.OutputDirectory -ChildPath "Reports\ConfigurationSummary_Enhanced.txt"
        
        $summary = @"
Microsoft 365 Tenant Configuration Summary - Enhanced Version
=============================================================
Generated: $($TenantData.GeneratedOn.ToString('yyyy-MM-dd HH:mm:ss'))
Script Version: 1.1-Complete Fixed Version

TENANT INFORMATION
------------------
Tenant Name: $($TenantData.TenantInfo.DisplayName)
Default Domain: $($TenantData.TenantInfo.DefaultDomain)
Tenant ID: $($TenantData.TenantInfo.TenantId)
Connected As: $($TenantData.TenantInfo.ConnectedAs)
Setup Date: $($TenantData.TenantInfo.SetupDate)

CONFIGURATION SUMMARY
---------------------
Total Groups: $($TenantData.Groups.TotalCount)
  - Security Groups: $($TenantData.Groups.SecurityGroups.Count)
  - Microsoft 365 Groups: $($TenantData.Groups.Microsoft365Groups.Count)
  - Distribution Groups: $($TenantData.Groups.DistributionGroups.Count)
  - Dynamic Groups: $($TenantData.Groups.DynamicGroups.Count)

Users: $($TenantData.Users.TotalUsers) total
  - Enabled: $($TenantData.Users.EnabledUsers)
  - Disabled: $($TenantData.Users.DisabledUsers)
  - Guests: $($TenantData.Users.GuestUsers)

Conditional Access Policies: $($TenantData.ConditionalAccess.TotalCount) total
  - Enabled: $($TenantData.ConditionalAccess.EnabledCount)
  - Disabled: $($TenantData.ConditionalAccess.DisabledCount)

Intune Configuration:
  - Managed Devices: $($TenantData.Intune.TotalDevices)
  - Configuration Policies: $($TenantData.Intune.DeviceConfigurationPolicies.Count)
  - Compliance Policies: $($TenantData.Intune.DeviceCompliancePolicies.Count)
  - Managed Apps: $($TenantData.Intune.ManagedApps.Count)

License Usage:
  - Total Licensed: $($TenantData.Licenses.UsedLicenses)
  - Total Available: $($TenantData.Licenses.TotalLicenses)

SERVICE ACCESS STATUS
--------------------
SharePoint Access: $($TenantData.SharePoint.AccessStatus)
$(if($TenantData.SharePoint.PermissionMessage) { "SharePoint Message: $($TenantData.SharePoint.PermissionMessage)" })

Intune Access Status:
  - Configuration Policies: $($TenantData.Intune.AccessStatus.ConfigurationPolicies)
  - Compliance Policies: $($TenantData.Intune.AccessStatus.CompliancePolicies)  
  - Managed Apps: $($TenantData.Intune.AccessStatus.ManagedApps)
  - Managed Devices: $($TenantData.Intune.AccessStatus.ManagedDevices)

FIXES APPLIED IN THIS VERSION
-----------------------------
1. Fixed SharePoint access issues with multiple fallback methods
2. Fixed Intune mobile app cmdlets (corrected cmdlet names)
3. Fixed CA policies to show detailed settings in proper table format
4. Fixed Intune Configuration policies to show detailed settings
5. Enhanced error handling and logging throughout
6. Added comprehensive access status reporting
7. Improved Excel column formatting and data placement

DETAILED POLICY INFORMATION
--------------------------
$(if($TenantData.ConditionalAccess.Policies.Count -gt 0) {
    "Conditional Access Policies:"
    foreach($policy in $TenantData.ConditionalAccess.Policies) {
        "  - $($policy.DisplayName) ($($policy.State))"
        if($policy.UserConditions) { "    Users: $($policy.UserConditions)" }
        if($policy.ApplicationConditions) { "    Apps: $($policy.ApplicationConditions)" }
        if($policy.GrantControls) { "    Controls: $($policy.GrantControls)" }
        ""
    }
} else {
    "No Conditional Access policies found or accessible."
})

$(if($TenantData.Intune.DeviceConfigurationPolicies.Count -gt 0) {
    "Intune Configuration Policies:"
    foreach($policy in $TenantData.Intune.DeviceConfigurationPolicies) {
        "  - $($policy.DisplayName) ($($policy.Platform))"
        "    Type: $($policy.PolicyType)"
        if($policy.DetailedSettings) { "    Settings: $($policy.DetailedSettings)" }
        ""
    }
} else {
    "No Intune Configuration policies found or accessible."
})

For comprehensive detailed information, see the HTML and Excel reports in the Reports folder.
The Excel report now contains properly formatted policy details in the correct table structure.

RECOMMENDATIONS
--------------
$(if($TenantData.SharePoint.AccessStatus -like "*Error*" -or $TenantData.SharePoint.AccessStatus -like "*Denied*") {
    "- Add Sites.Read.All or Sites.ReadWrite.All permission for SharePoint access"
})
$(if($TenantData.Intune.AccessStatus.ManagedApps -like "*Failed*" -or $TenantData.Intune.AccessStatus.ManagedApps -like "*No apps*") {
    "- Add DeviceManagementApps.Read.All permission for Intune app management access"
})
$(if($TenantData.ConditionalAccess.TotalCount -eq 0) {
    "- Consider implementing Conditional Access policies for enhanced security"
})
$(if($TenantData.Intune.TotalDevices -eq 0) {
    "- Consider enrolling devices in Intune for centralized management"
})

Report generated by Microsoft 365 Tenant Setup Utility - Complete Fixed Version
"@

        $summary | Out-File -FilePath $summaryPath -Encoding UTF8
        Write-LogMessage -Message "Enhanced configuration summary generated: $summaryPath" -Type Success -LogOnly
        return $true
    }
    catch {
        Write-LogMessage -Message "Failed to generate configuration summary: $($_.Exception.Message)" -Type Error
        return $false
    }
}

# === Export the main function for module usage ===
Export-ModuleMember -Function New-TenantDocumentation