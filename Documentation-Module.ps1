# Enhanced Documentation Module for Microsoft 365 Tenant Setup Utility
#requires -Version 5.1
<#
.SYNOPSIS
    Enhanced Documentation Module for Microsoft 365 Tenant Setup Utility
.DESCRIPTION
    Generates comprehensive documentation with SharePoint permissions, enhanced license mapping, 
    Security Groups sheet, table expansion, and detailed Conditional Access policies
.NOTES
    Version: 3.1 - Complete Enhancement with all requested features - FIXED MODULE LOADING
    Dependencies: Microsoft Graph PowerShell SDK, ImportExcel module
#>

# === Enhanced Documentation Configuration ===
$DocumentationConfig = @{
    OutputDirectory = "$env:USERPROFILE\Documents\M365TenantSetup_Documentation_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
    ReportFormats = @('HTML', 'Excel', 'JSON')
    IncludeScreenshots = $false
    DetailLevel = 'Detailed'
    MaxUsersPerPermissionCell = 20  # Truncate if more users in SharePoint permissions
    EnableTableExpansion = $true
}

# === Enhanced License Mapping Table ===
$LicenseMapping = @{
    # Map to existing template license types FIRST
    "Microsoft_365_Business_Basic" = "Business Basic"
    "Microsoft_365_Business_Standard" = "Business Standard" 
    "Microsoft_365_Business_Premium" = "Business Premium"
    
    # E-series licenses (consistent short names)
    "Microsoft_365_E1" = "E1"
    "Microsoft_365_E3" = "E3"
    "Microsoft_365_E5" = "E5"
    "Microsoft_365_E5_(no_Teams)" = "E5 (no Teams)"
    
    # Exchange plans
    "ExchangeOnline_PLAN1" = "Exchange Plan 1"
    "ExchangeOnline_PLAN2" = "Exchange Plan 2"
    
    # Additional software mappings (for Additional Software columns)
    "Microsoft_Teams_Enterprise_New" = "Teams"
    "FLOW_FREE" = "Power Automate"
    "PowerBI_Standard" = "Power BI"
    "Microsoft_Intune" = "Intune"
    "Azure_Active_Directory_Premium_P1" = "Azure AD P1"
    "Azure_Active_Directory_Premium_P2" = "Azure AD P2"
    "Microsoft_Defender_for_Office_365_Plan_1" = "Defender for Office 365"
    
    # Everything else keeps full name if not in mapping table
}

# === CONVERTED TO DIRECT EXECUTION METHOD ===
function New-TenantDocumentation {
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
        
        # STEP 5: Force load ONLY the exact modules needed for Enhanced Documentation - FIXED
        $documentationModules = @(
            'Microsoft.Graph.Groups',
            'Microsoft.Graph.Identity.DirectoryManagement', 
            'Microsoft.Graph.Users',
            'Microsoft.Graph.Sites',
            'Microsoft.Graph.DeviceManagement',
            'Microsoft.Graph.Identity.SignIns',
            'Microsoft.Graph.Applications',
            'Microsoft.Graph.Teams'
        )
        
        Write-LogMessage -Message "Loading required Graph modules for Enhanced Documentation..." -Type Info
        foreach ($module in $documentationModules) {
            try {
                Import-Module $module -Force -ErrorAction Stop
                $moduleInfo = Get-Module $module
                Write-LogMessage -Message "Loaded $module version $($moduleInfo.Version)" -Type Success -LogOnly
            }
            catch {
                Write-LogMessage -Message "Failed to load $module module - $($_.Exception.Message)" -Type Error
                return $false
            }
        }
        
        # STEP 6: Connect with COMPREHENSIVE scopes needed for Enhanced Documentation
        $documentationScopes = @(
            "Group.ReadWrite.All",
            "Directory.ReadWrite.All",
            "User.ReadWrite.All",
            "Sites.ReadWrite.All",
            "Sites.Manage.All",
            "DeviceManagementManagedDevices.ReadWrite.All",
            "DeviceManagementConfiguration.ReadWrite.All",
            "DeviceManagementApps.ReadWrite.All",
            "Policy.ReadWrite.ConditionalAccess",
            "Application.ReadWrite.All",
            "Team.ReadBasic.All"
        )
        
        Write-LogMessage -Message "Connecting to Microsoft Graph with Enhanced Documentation scopes..." -Type Info
        Connect-MgGraph -Scopes $documentationScopes -ErrorAction Stop | Out-Null
        
        $context = Get-MgContext
        Write-LogMessage -Message "Connected to Microsoft Graph as $($context.Account)" -Type Success
        
        # STEP 7: Enhanced Documentation logic starts here
        
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
        
        # Gather all tenant information with enhancements
        Write-LogMessage -Message "Gathering comprehensive tenant configuration data..." -Type Info
        $tenantData = Get-EnhancedTenantConfiguration
        
        # Generate enhanced populated Excel documentation
        Write-LogMessage -Message "Populating Excel template with enhanced configuration data..." -Type Info
        $excelGenerated = New-EnhancedExcelDocumentation -TenantData $tenantData -TemplatePath $templatePath
        
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
            if (Test-Path $DocumentationConfig.OutputDirectory) {
                Start-Process explorer.exe -ArgumentList $DocumentationConfig.OutputDirectory
                Write-LogMessage -Message "Opened documentation directory" -Type Success
            }
        }
        
        Write-LogMessage -Message "Enhanced documentation completed successfully" -Type Success
        return $true
        
    }
    catch {
        Write-LogMessage -Message "Enhanced documentation process failed: $($_.Exception.Message)" -Type Error
        Write-LogMessage -Message "Error Details: $($_.Exception)" -Type Error -LogOnly
        return $false
    }
}

# === Enhanced Data Collection Functions ===

function Get-EnhancedTenantConfiguration {
    <#
    .SYNOPSIS
        Collects comprehensive tenant configuration with enhancements
    #>
    
    Write-LogMessage -Message "Starting enhanced tenant data collection..." -Type Info
    
    $tenantData = @{
        Tenant = Get-TenantInformation
        Groups = Get-EnhancedGroupsInformation
        Users = Get-EnhancedUsersInformation
        ConditionalAccess = Get-EnhancedConditionalAccessInformation
        SharePoint = Get-EnhancedSharePointInformation
        Intune = Get-EnhancedIntuneInformation
        GeneratedDate = Get-Date
    }
    
    Write-LogMessage -Message "Enhanced tenant data collection completed" -Type Success
    return $tenantData
}

function Get-EnhancedGroupsInformation {
    <#
    .SYNOPSIS
        Collects comprehensive groups information including security groups and distribution lists
    #>
    
    try {
        Write-LogMessage -Message "Collecting enhanced groups information..." -Type Info
        
        $allGroups = Get-MgGroup -All -Top 1000
        
        $groupsInfo = @{
            SecurityGroups = @()
            DistributionGroups = @()
            TotalGroups = $allGroups.Count
        }
        
        foreach ($group in $allGroups) {
            try {
                # Get group members
                $members = @()
                try {
                    $groupMembers = Get-MgGroupMember -GroupId $group.Id -All
                    $members = $groupMembers | ForEach-Object {
                        if ($_.AdditionalProperties.userPrincipalName) {
                            $_.AdditionalProperties.userPrincipalName
                        } elseif ($_.AdditionalProperties.displayName) {
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
                    MemberCount = $members.Count
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

function Get-EnhancedUsersInformation {
    <#
    .SYNOPSIS
        Collects users information with enhanced license mapping
    #>
    
    try {
        Write-LogMessage -Message "Collecting enhanced users information with license mapping..." -Type Info
        
        $users = Get-MgUser -All -Top 500
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
                # Get user licenses
                $userLicenses = @()
                $baseLicense = ""
                $additionalLicenses = @()
                
                if ($user.AssignedLicenses) {
                    # Get license details
                    $subscribedSkus = Get-MgSubscribedSku
                    
                    foreach ($assignedLicense in $user.AssignedLicenses) {
                        $sku = $subscribedSkus | Where-Object { $_.SkuId -eq $assignedLicense.SkuId }
                        if ($sku) {
                            $licenseDisplayName = $sku.SkuPartNumber
                            $userLicenses += $licenseDisplayName
                        }
                    }
                    
                    # Apply license mapping logic
                    if ($userLicenses.Count -gt 0) {
                        $mappedLicenses = Apply-LicenseMapping -UserLicenses $userLicenses
                        $baseLicense = $mappedLicenses.BaseLicense
                        $additionalLicenses = $mappedLicenses.AdditionalLicenses
                        Write-LogMessage -Message "User $($user.UserPrincipalName) - Base: $baseLicense, Additional: $($additionalLicenses -join ',')" -Type Info -LogOnly
                    }
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
                    Manager = ""
                }
                
                # Get manager information if available
                try {
                    $manager = Get-MgUserManager -UserId $user.Id -ErrorAction SilentlyContinue
                    if ($manager) {
                        $userData.Manager = $manager.AdditionalProperties.userPrincipalName
                    }
                }
                catch {
                    # Manager not found or accessible
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

function Apply-LicenseMapping {
    <#
    .SYNOPSIS
        Applies license mapping logic to determine base license and additional software
    #>
    param(
        [string[]]$UserLicenses
    )
    
    $baseLicense = ""
    $additionalLicenses = @()
    
    # Define priority order for base licenses (higher priority = more likely to be base license)
    $baseLicensePriority = @{
        "Microsoft_365_E5" = 100
        "Microsoft_365_E5_(no_Teams)" = 99
        "Microsoft_365_E3" = 90
        "Microsoft_365_E1" = 80
        "Microsoft_365_Business_Premium" = 70
        "Microsoft_365_Business_Standard" = 60
        "Microsoft_365_Business_Basic" = 50
    }
    
    # Separate base licenses from additional software
    $potentialBaseLicenses = @()
    $potentialAdditionalSoftware = @()
    
    foreach ($license in $UserLicenses) {
        $mappedName = if ($script:LicenseMapping.ContainsKey($license)) { $script:LicenseMapping[$license] } else { $license }
        
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
            $baseLicense = if ($script:LicenseMapping.ContainsKey($firstLicense)) { $script:LicenseMapping[$firstLicense] } else { $firstLicense }
            
            # Add remaining licenses to additional software
            for ($i = 1; $i -lt $UserLicenses.Count; $i++) {
                $license = $UserLicenses[$i]
                $mappedName = if ($script:LicenseMapping.ContainsKey($license)) { $script:LicenseMapping[$license] } else { $license }
                $potentialAdditionalSoftware += $mappedName
            }
        }
    }
    
    return @{
        BaseLicense = $baseLicense
        AdditionalLicenses = $potentialAdditionalSoftware | Select-Object -Unique
    }
}

function Get-EnhancedSharePointInformation {
    <#
    .SYNOPSIS
        Collects SharePoint sites with comprehensive permissions information
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
        
        # Get all SharePoint sites
        $sites = @()
        try {
            Write-LogMessage -Message "Retrieving all SharePoint sites..." -Type Info
            $allSites = Get-MgSite -All -Top 200
            $sites = $allSites
            Write-LogMessage -Message "Found $($sites.Count) SharePoint sites" -Type Success
        }
        catch {
            Write-LogMessage -Message "Error retrieving SharePoint sites: $($_.Exception.Message)" -Type Warning
            Write-LogMessage -Message "This may be due to insufficient SharePoint permissions" -Type Warning
        }
        
        if ($sites.Count -gt 0) {
            $spInfo.TotalSites = $sites.Count
            $processedCount = 0
            
            foreach ($site in $sites) {
                $processedCount++
                Write-Progress -Activity "Processing SharePoint Sites" -Status "Site $processedCount of $($sites.Count): $($site.DisplayName)" -PercentComplete (($processedCount / $sites.Count) * 100)
                
                try {
                    # Get site permissions
                    $permissions = Get-SharePointSitePermissions -SiteId $site.Id -SiteName $site.DisplayName
                    
                    $siteData = @{
                        Id = $site.Id
                        DisplayName = $site.DisplayName
                        Name = $site.Name
                        WebUrl = $site.WebUrl
                        CreatedDateTime = $site.CreatedDateTime
                        LastModifiedDateTime = $site.LastModifiedDateTime
                        SiteCollection = $site.SiteCollection
                        Owners = $permissions.Owners
                        Members = $permissions.Members
                        ReadOnly = $permissions.ReadOnly
                        Approver = "" # This would need to be manually filled or configured
                    }
                    
                    $spInfo.Sites += $siteData
                }
                catch {
                    Write-LogMessage -Message "Error processing site $($site.DisplayName): $($_.Exception.Message)" -Type Warning -LogOnly
                    
                    # Add site with basic info only if permissions fail
                    $siteData = @{
                        Id = $site.Id
                        DisplayName = $site.DisplayName
                        Name = $site.Name
                        WebUrl = $site.WebUrl
                        CreatedDateTime = $site.CreatedDateTime
                        LastModifiedDateTime = $site.LastModifiedDateTime
                        SiteCollection = $site.SiteCollection
                        Owners = "Permission access failed"
                        Members = "Permission access failed"
                        ReadOnly = "Permission access failed"
                        Approver = ""
                    }
                    
                    $spInfo.Sites += $siteData
                }
            }
            
            Write-Progress -Activity "Processing SharePoint Sites" -Completed
            Write-LogMessage -Message "Successfully processed $($spInfo.Sites.Count) SharePoint sites with permissions" -Type Success
        } else {
            Write-LogMessage -Message "No SharePoint sites accessible with current permissions" -Type Warning
            $spInfo.TotalSites = 0
            $spInfo.Sites = @(
                @{
                    Id = "No Access"
                    DisplayName = "SharePoint sites not accessible"
                    Name = "Insufficient permissions"
                    WebUrl = "Check SharePoint Administrator role"
                    CreatedDateTime = Get-Date
                    LastModifiedDateTime = Get-Date
                    SiteCollection = $null
                    Owners = "N/A"
                    Members = "N/A"
                    ReadOnly = "N/A"
                    Approver = "N/A"
                }
            )
        }
        
        return $spInfo
    }
    catch {
        Write-LogMessage -Message "Error in enhanced SharePoint information collection: $($_.Exception.Message)" -Type Error
        return @{
            TenantSettings = @{}
            Sites = @()
            TotalSites = 0
            StorageUsed = "Error collecting data"
            SharingSettings = "Error collecting data"
            ExternalSharingEnabled = "Error collecting data"
        }
    }
}

function Get-SharePointSitePermissions {
    <#
    .SYNOPSIS
        Gets permissions for a specific SharePoint site
    #>
    param(
        [string]$SiteId,
        [string]$SiteName
    )
    
    $permissions = @{
        Owners = ""
        Members = ""
        ReadOnly = ""
    }
    
    try {
        Write-LogMessage -Message "Getting permissions for site: $SiteName" -Type Info -LogOnly
        
        # Try to get site permissions through Graph API
        # This is a complex operation as SharePoint permissions can be at multiple levels
        
        # Method 1: Try to get the associated Office 365 group (for modern team sites)
        try {
            $siteDetails = Get-MgSite -SiteId $SiteId
            
            # Check if this site has an associated Office 365 group
            if ($siteDetails.SiteCollection.Hostname -like "*.sharepoint.com") {
                # Try to find associated group by looking for groups with matching SharePoint site
                $groups = Get-MgGroup -Filter "resourceProvisioningOptions/any(x:x eq 'Team')" -Select "id,displayName,mail" -Top 50
                
                foreach ($group in $groups) {
                    try {
                        # Get group members and owners
                        $groupOwners = Get-MgGroupOwner -GroupId $group.Id
                        $groupMembers = Get-MgGroupMember -GroupId $group.Id
                        
                        if ($groupOwners -or $groupMembers) {
                            # Extract owner names/emails
                            $ownerNames = @()
                            foreach ($owner in $groupOwners) {
                                if ($owner.AdditionalProperties.userPrincipalName) {
                                    $ownerNames += $owner.AdditionalProperties.userPrincipalName
                                } elseif ($owner.AdditionalProperties.displayName) {
                                    $ownerNames += $owner.AdditionalProperties.displayName
                                }
                            }
                            
                            # Extract member names/emails (excluding owners)
                            $memberNames = @()
                            $ownerIds = $groupOwners | ForEach-Object { $_.Id }
                            foreach ($member in $groupMembers) {
                                if ($member.Id -notin $ownerIds) {
                                    if ($member.AdditionalProperties.userPrincipalName) {
                                        $memberNames += $member.AdditionalProperties.userPrincipalName
                                    } elseif ($member.AdditionalProperties.displayName) {
                                        $memberNames += $member.AdditionalProperties.displayName
                                    }
                                }
                            }
                            
                            # Limit the number of users shown to prevent overwhelming the Excel cell
                            if ($ownerNames.Count -gt $DocumentationConfig.MaxUsersPerPermissionCell) {
                                $permissions.Owners = ($ownerNames | Select-Object -First $DocumentationConfig.MaxUsersPerPermissionCell) -join ", " + ", ... and $($ownerNames.Count - $DocumentationConfig.MaxUsersPerPermissionCell) more"
                            } else {
                                $permissions.Owners = $ownerNames -join ", "
                            }
                            
                            if ($memberNames.Count -gt $DocumentationConfig.MaxUsersPerPermissionCell) {
                                $permissions.Members = ($memberNames | Select-Object -First $DocumentationConfig.MaxUsersPerPermissionCell) -join ", " + ", ... and $($memberNames.Count - $DocumentationConfig.MaxUsersPerPermissionCell) more"
                            } else {
                                $permissions.Members = $memberNames -join ", "
                            }
                            
                            # For ReadOnly, we might need to check site visitors specifically
                            $permissions.ReadOnly = "Visitors group (detailed permissions require SharePoint admin)"
                            
                            break  # Found a matching group, stop looking
                        }
                    }
                    catch {
                        # Continue to next group
                        continue
                    }
                }
            }
        }
        catch {
            Write-LogMessage -Message "Could not retrieve Office 365 group permissions for $SiteName" -Type Warning -LogOnly
        }
        
        # Method 2: Try direct site permissions API (if available)
        if ([string]::IsNullOrEmpty($permissions.Owners)) {
            try {
                # Try to get site permissions directly
                $uri = "https://graph.microsoft.com/v1.0/sites/$SiteId/permissions"
                $sitePermissions = Invoke-MgGraphRequest -Uri $uri -Method GET -ErrorAction SilentlyContinue
                
                if ($sitePermissions -and $sitePermissions.value) {
                    $allPermissions = @()
                    foreach ($perm in $sitePermissions.value) {
                        if ($perm.grantedToIdentitiesV2) {
                            foreach ($identity in $perm.grantedToIdentitiesV2) {
                                $allPermissions += "$($identity.user.displayName) ($($perm.roles -join ', '))"
                            }
                        }
                    }
                    
                    if ($allPermissions.Count -gt 0) {
                        $permissions.Owners = "Mixed permissions"
                        $permissions.Members = $allPermissions -join ", "
                        $permissions.ReadOnly = "See members list"
                    }
                }
            }
            catch {
                # Direct permissions API not available or insufficient permissions
                Write-LogMessage -Message "Direct site permissions API not accessible for $SiteName" -Type Warning -LogOnly
            }
        }
        
        # Fallback: Indicate that manual configuration is needed
        if ([string]::IsNullOrEmpty($permissions.Owners) -and [string]::IsNullOrEmpty($permissions.Members)) {
            $permissions.Owners = "Manual configuration required"
            $permissions.Members = "Manual configuration required"
            $permissions.ReadOnly = "Manual configuration required"
        }
        
    }
    catch {
        Write-LogMessage -Message "Error getting permissions for site $SiteName`: $($_.Exception.Message)" -Type Warning -LogOnly
        $permissions.Owners = "Error retrieving permissions"
        $permissions.Members = "Error retrieving permissions"
        $permissions.ReadOnly = "Error retrieving permissions"
    }
    
    return $permissions
}

function Get-EnhancedConditionalAccessInformation {
    <#
    .SYNOPSIS
        Collects Conditional Access policies with detailed technical information
    #>
    
    try {
        Write-LogMessage -Message "Collecting enhanced Conditional Access policies..." -Type Info
        
        $caInfo = @{
            Policies = @()
            TotalPolicies = 0
        }
        
        $policies = Get-MgIdentityConditionalAccessPolicy -All
        $caInfo.TotalPolicies = $policies.Count
        
        foreach ($policy in $policies) {
            try {
                # Create detailed technical description
                $technicalDetails = Build-ConditionalAccessTechnicalDetails -Policy $policy
                
                $policyData = @{
                    Id = $policy.Id
                    DisplayName = $policy.DisplayName
                    State = $policy.State
                    CreatedDateTime = $policy.CreatedDateTime
                    ModifiedDateTime = $policy.ModifiedDateTime
                    TechnicalDetails = $technicalDetails
                    Conditions = @{
                        Users = $policy.Conditions.Users
                        Applications = $policy.Conditions.Applications
                        Platforms = $policy.Conditions.Platforms
                        Locations = $policy.Conditions.Locations
                        ClientAppTypes = $policy.Conditions.ClientAppTypes
                        SignInRiskLevels = $policy.Conditions.SignInRiskLevels
                        UserRiskLevels = $policy.Conditions.UserRiskLevels
                    }
                    GrantControls = $policy.GrantControls
                    SessionControls = $policy.SessionControls
                }
                
                $caInfo.Policies += $policyData
            }
            catch {
                Write-LogMessage -Message "Error processing CA policy $($policy.DisplayName): $($_.Exception.Message)" -Type Warning -LogOnly
            }
        }
        
        Write-LogMessage -Message "Collected $($caInfo.TotalPolicies) Conditional Access policies with technical details" -Type Success
        return $caInfo
    }
    catch {
        Write-LogMessage -Message "Error collecting Conditional Access information: $($_.Exception.Message)" -Type Warning
        return @{
            Policies = @()
            TotalPolicies = 0
        }
    }
}

function Build-ConditionalAccessTechnicalDetails {
    <#
    .SYNOPSIS
        Builds technical details string for Conditional Access policy
    #>
    param($Policy)
    
    $details = @()
    
    # State
    $details += "State: $($Policy.State)"
    
    # Users/Groups conditions
    if ($Policy.Conditions.Users) {
        $userConditions = @()
        if ($Policy.Conditions.Users.IncludeUsers) {
            $userConditions += "Include Users: $($Policy.Conditions.Users.IncludeUsers -join ', ')"
        }
        if ($Policy.Conditions.Users.ExcludeUsers) {
            $userConditions += "Exclude Users: $($Policy.Conditions.Users.ExcludeUsers -join ', ')"
        }
        if ($Policy.Conditions.Users.IncludeGroups) {
            $userConditions += "Include Groups: $($Policy.Conditions.Users.IncludeGroups -join ', ')"
        }
        if ($Policy.Conditions.Users.ExcludeGroups) {
            $userConditions += "Exclude Groups: $($Policy.Conditions.Users.ExcludeGroups -join ', ')"
        }
        if ($userConditions.Count -gt 0) {
            $details += "Users: $($userConditions -join '; ')"
        }
    }
    
    # Applications conditions
    if ($Policy.Conditions.Applications) {
        $appConditions = @()
        if ($Policy.Conditions.Applications.IncludeApplications) {
            $appConditions += "Include Apps: $($Policy.Conditions.Applications.IncludeApplications -join ', ')"
        }
        if ($Policy.Conditions.Applications.ExcludeApplications) {
            $appConditions += "Exclude Apps: $($Policy.Conditions.Applications.ExcludeApplications -join ', ')"
        }
        if ($appConditions.Count -gt 0) {
            $details += "Applications: $($appConditions -join '; ')"
        }
    }
    
    # Platform conditions
    if ($Policy.Conditions.Platforms -and $Policy.Conditions.Platforms.IncludePlatforms) {
        $details += "Platforms: $($Policy.Conditions.Platforms.IncludePlatforms -join ', ')"
    }
    
    # Location conditions
    if ($Policy.Conditions.Locations) {
        $locationConditions = @()
        if ($Policy.Conditions.Locations.IncludeLocations) {
            $locationConditions += "Include: $($Policy.Conditions.Locations.IncludeLocations -join ', ')"
        }
        if ($Policy.Conditions.Locations.ExcludeLocations) {
            $locationConditions += "Exclude: $($Policy.Conditions.Locations.ExcludeLocations -join ', ')"
        }
        if ($locationConditions.Count -gt 0) {
            $details += "Locations: $($locationConditions -join '; ')"
        }
    }
    
    # Grant controls
    if ($Policy.GrantControls) {
        $grantDetails = @()
        if ($Policy.GrantControls.Operator) {
            $grantDetails += "Operator: $($Policy.GrantControls.Operator)"
        }
        if ($Policy.GrantControls.BuiltInControls) {
            $grantDetails += "Controls: $($Policy.GrantControls.BuiltInControls -join ', ')"
        }
        if ($grantDetails.Count -gt 0) {
            $details += "Grant: $($grantDetails -join '; ')"
        }
    }
    
    # Session controls
    if ($Policy.SessionControls) {
        $sessionDetails = @()
        if ($Policy.SessionControls.ApplicationEnforcedRestrictions) {
            $sessionDetails += "App Enforced Restrictions"
        }
        if ($Policy.SessionControls.CloudAppSecurity) {
            $sessionDetails += "Cloud App Security"
        }
        if ($Policy.SessionControls.SignInFrequency) {
            $sessionDetails += "Sign-in Frequency: $($Policy.SessionControls.SignInFrequency.Value) $($Policy.SessionControls.SignInFrequency.Type)"
        }
        if ($sessionDetails.Count -gt 0) {
            $details += "Session: $($sessionDetails -join '; ')"
        }
    }
    
    return $details -join " | "
}

function Get-EnhancedIntuneInformation {
    <#
    .SYNOPSIS
        Collects Intune configuration information with enhanced mobile app collection
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
        
        # Compliance Policies
        try {
            $compliancePolicies = Get-MgDeviceManagementDeviceCompliancePolicy -All
            Write-LogMessage -Message "Found $($compliancePolicies.Count) compliance policies" -Type Info
            $intuneInfo.DeviceCompliancePolicies = $compliancePolicies | ForEach-Object {
                @{
                    Id = $_.Id
                    DisplayName = $_.DisplayName
                    Description = $_.Description
                    CreatedDateTime = $_.CreatedDateTime
                    LastModifiedDateTime = $_.LastModifiedDateTime
                    Version = $_.Version
                }
            }
        }
        catch {
            Write-LogMessage -Message "Could not retrieve compliance policies: $($_.Exception.Message)" -Type Warning
        }
        
        # Configuration Policies
        try {
            $configPolicies = Get-MgDeviceManagementDeviceConfiguration -All
            Write-LogMessage -Message "Found $($configPolicies.Count) configuration policies" -Type Info
            $intuneInfo.DeviceConfigurationPolicies = $configPolicies | ForEach-Object {
                @{
                    Id = $_.Id
                    DisplayName = $_.DisplayName
                    Description = $_.Description
                    CreatedDateTime = $_.CreatedDateTime
                    LastModifiedDateTime = $_.LastModifiedDateTime
                    Version = $_.Version
                }
            }
        }
        catch {
            Write-LogMessage -Message "Could not retrieve device configuration policies: $($_.Exception.Message)" -Type Warning
        }
        
        # Enhanced Managed Apps Collection with multiple methods
        try {
            Write-LogMessage -Message "Collecting managed applications with enhanced methods..." -Type Info
            
            $managedApps = @()
            
            # Method 1: Standard Get-MgDeviceManagementMobileApp
            try {
                $managedApps = Get-MgDeviceManagementMobileApp -All -ErrorAction Stop
                Write-LogMessage -Message "Found $($managedApps.Count) managed apps using standard method" -Type Success
            }
            catch {
                Write-LogMessage -Message "Standard mobile app collection failed: $($_.Exception.Message)" -Type Warning
                
                # Method 2: Try with Graph API directly
                try {
                    $uri = "https://graph.microsoft.com/v1.0/deviceAppManagement/mobileApps"
                    $response = Invoke-MgGraphRequest -Uri $uri -Method GET
                    $managedApps = $response.value
                    Write-LogMessage -Message "Found $($managedApps.Count) managed apps using Graph API" -Type Success
                }
                catch {
                    Write-LogMessage -Message "Graph API mobile app collection also failed: $($_.Exception.Message)" -Type Warning
                }
            }
            
            if ($managedApps -and $managedApps.Count -gt 0) {
                $intuneInfo.ManagedApps = $managedApps | ForEach-Object {
                    # Determine platform based on odata type
                    $platform = "Unknown"
                    if ($_.'@odata.type') {
                        switch ($_.'@odata.type') {
                            "#microsoft.graph.win32LobApp" { $platform = "Windows" }
                            "#microsoft.graph.winGetApp" { $platform = "Windows" }
                            "#microsoft.graph.officeSuiteApp" { $platform = "Windows" }
                            "#microsoft.graph.androidLobApp" { $platform = "Android" }
                            "#microsoft.graph.androidManagedStoreApp" { $platform = "Android" }
                            "#microsoft.graph.iosLobApp" { $platform = "iOS" }
                            "#microsoft.graph.iosStoreApp" { $platform = "iOS" }
                            "#microsoft.graph.iosVppApp" { $platform = "iOS" }
                            "#microsoft.graph.macOSLobApp" { $platform = "macOS" }
                            "#microsoft.graph.macOSOfficeApp" { $platform = "macOS" }
                            "#microsoft.graph.webApp" { $platform = "Web" }
                            default { $platform = "Unknown" }
                        }
                    }
                    
                    @{
                        Id = $_.Id
                        DisplayName = $_.DisplayName
                        Description = $_.Description
                        Publisher = $_.Publisher
                        CreatedDateTime = $_.CreatedDateTime
                        LastModifiedDateTime = $_.LastModifiedDateTime
                        '@odata.type' = $_.'@odata.type'
                        Platform = $platform
                    }
                }
                Write-LogMessage -Message "Successfully processed $($intuneInfo.ManagedApps.Count) managed apps with platform categorization" -Type Success
            } else {
                Write-LogMessage -Message "No managed apps found or accessible" -Type Warning
                $intuneInfo.ManagedApps = @()
            }
        }
        catch {
            Write-LogMessage -Message "Could not retrieve managed applications: $($_.Exception.Message)" -Type Warning
            Write-LogMessage -Message "This might be due to insufficient permissions for app management" -Type Info
            $intuneInfo.ManagedApps = @()
        }
        
        # Managed Devices
        try {
            Write-LogMessage -Message "Collecting managed devices..." -Type Info -LogOnly
            $devices = Get-MgDeviceManagementManagedDevice -All -Top 500
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

# === Enhanced Excel Documentation Functions ===

function New-EnhancedExcelDocumentation {
    <#
    .SYNOPSIS
        Generates enhanced Excel documentation with robust error handling
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
        
        Write-LogMessage -Message "EXCEL DEBUG: File size after copy: $((Get-Item $outputPath).Length) bytes" -Type Info
        
        # Load Excel application with error handling
        try {
            $excel = New-Object -ComObject Excel.Application
            $excel.Visible = $false
            $excel.DisplayAlerts = $false
            $excel.ScreenUpdating = $false  # Improve performance
            Write-LogMessage -Message "EXCEL DEBUG: Excel application created successfully" -Type Success
        }
        catch {
            Write-LogMessage -Message "EXCEL DEBUG: Failed to create Excel application: $($_.Exception.Message)" -Type Error
            return $false
        }
        
        try {
            Write-LogMessage -Message "EXCEL DEBUG: Opening workbook..." -Type Info
            $workbook = $excel.Workbooks.Open($outputPath)
            Write-LogMessage -Message "EXCEL DEBUG: Workbook opened successfully" -Type Success
            
            # List all worksheets for debugging
            Write-LogMessage -Message "EXCEL DEBUG: Available worksheets:" -Type Info
            for ($i = 1; $i -le $workbook.Worksheets.Count; $i++) {
                $sheetName = $workbook.Worksheets.Item($i).Name
                Write-LogMessage -Message "EXCEL DEBUG: Sheet $i`: $sheetName" -Type Info -LogOnly
            }
            
            # Update sheets one by one with robust error handling
            try {
                Update-EnhancedUsersSheet -Workbook $workbook -TenantData $TenantData
                Update-EnhancedLicensingSheet -Workbook $workbook -TenantData $TenantData
                Update-EnhancedConditionalAccessSheet -Workbook $workbook -TenantData $TenantData
                Update-EnhancedSharePointSheet -Workbook $workbook -TenantData $TenantData
                Update-EnhancedIntuneAppsSheets -Workbook $workbook -TenantData $TenantData
                Add-SecurityGroupsSheet -Workbook $workbook -TenantData $TenantData
                Update-DistributionListSheet -Workbook $workbook -TenantData $TenantData
            }
            catch {
                Write-LogMessage -Message "EXCEL DEBUG: Error during sheet updates: $($_.Exception.Message)" -Type Error
            }
            
            # Final save and close
            Write-LogMessage -Message "EXCEL DEBUG: Performing final save..." -Type Info
            $workbook.Save()
            $workbook.Close()
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

# === Enhanced Excel Update Functions ===

function Update-EnhancedUsersSheet {
    <#
    .SYNOPSIS
        Updates Users sheet with enhanced license mapping and robust error handling
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
        
        # Start at row 7 (first data row)
        $startRow = 7
        $currentRow = $startRow
        
        # Populate first 50 users to avoid Excel timeouts
        $usersToProcess = $users | Select-Object -First 50
        
        foreach ($user in $usersToProcess) {
            try {
                Write-LogMessage -Message "EXCEL DEBUG: Processing user $($user.UserPrincipalName) at row $currentRow" -Type Info -LogOnly
                
                # Use .Value instead of .Value2 for better compatibility
                $worksheet.Cells.Item($currentRow, 1).Value = if ($user.GivenName) { $user.GivenName } else { "" }
                $worksheet.Cells.Item($currentRow, 2).Value = if ($user.Surname) { $user.Surname } else { "" }
                $worksheet.Cells.Item($currentRow, 3).Value = if ($user.UserPrincipalName) { $user.UserPrincipalName } else { "" }
                $worksheet.Cells.Item($currentRow, 4).Value = if ($user.JobTitle) { $user.JobTitle } else { "" }
                $worksheet.Cells.Item($currentRow, 5).Value = if ($user.Manager) { $user.Manager } else { "" }
                $worksheet.Cells.Item($currentRow, 6).Value = if ($user.Department) { $user.Department } else { "" }
                $worksheet.Cells.Item($currentRow, 7).Value = if ($user.Office) { $user.Office } else { "" }
                $worksheet.Cells.Item($currentRow, 8).Value = "" # Phone Number
                
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

function Update-EnhancedLicensingSheet {
    <#
    .SYNOPSIS
        Updates Licensing sheet with enhanced license mapping
    #>
    param(
        [Parameter(Mandatory = $true)]
        $Workbook,
        
        [Parameter(Mandatory = $true)]
        [hashtable]$TenantData
    )
    
    try {
        Write-LogMessage -Message "Updating Licensing sheet with enhanced license mapping..." -Type Info
        
        $worksheet = $Workbook.Worksheets["Licensing"]
        if (-not $worksheet) {
            Write-LogMessage -Message "Licensing worksheet not found" -Type Warning
            return
        }
        
        $users = $TenantData.Users.Users | Where-Object { $_.BaseLicense -and $_.BaseLicense.Trim() -ne "" }
        if (-not $users -or $users.Count -eq 0) {
            Write-LogMessage -Message "No licensed users found. Checking all users for licensing data..." -Type Warning
            
            # Debug: Show first few users and their license info
            $allUsers = $TenantData.Users.Users | Select-Object -First 5
            foreach ($debugUser in $allUsers) {
                Write-LogMessage -Message "Debug User: $($debugUser.UserPrincipalName) - BaseLicense: '$($debugUser.BaseLicense)' - Licenses: '$($debugUser.Licenses)'" -Type Info -LogOnly
            }
            
            # If no licensed users, populate with all users but mark as unlicensed
            $users = $TenantData.Users.Users | Select-Object -First 20  # Limit to first 20 for testing
            Write-LogMessage -Message "Using first 20 users for licensing sheet as fallback" -Type Info
        }
        
        # Check if we need to expand the licensing table
        $availableRows = 48  # From template analysis
        $neededRows = $users.Count
        
        if ($neededRows -gt $availableRows) {
            Write-LogMessage -Message "Expanding Licensing table: need $neededRows rows, have $availableRows" -Type Info
            Expand-ExcelTable -Worksheet $worksheet -CurrentRows $availableRows -NeededRows $neededRows -StartRow 8
        }
        
        # Populate licensing data
        $startRow = 8  # Data starts at row 8 based on analysis
        $currentRow = $startRow
        
        foreach ($user in $users) {
            try {
                # Map to template columns: B=User Name, C=Base License Type, D=Additional Software 1, E=Additional Software 2
                $worksheet.Cells.Item($currentRow, 2).Value2 = $user.DisplayName
                $worksheet.Cells.Item($currentRow, 3).Value2 = $user.BaseLicense
                $worksheet.Cells.Item($currentRow, 4).Value2 = $user.AdditionalSoftware1
                $worksheet.Cells.Item($currentRow, 5).Value2 = $user.AdditionalSoftware2
                
                $currentRow++
            }
            catch {
                Write-LogMessage -Message "Error updating licensing row $currentRow`: $($_.Exception.Message)" -Type Warning -LogOnly
                $currentRow++
            }
        }
        
        Write-LogMessage -Message "Successfully updated Licensing sheet with $($users.Count) licensed users" -Type Success
    }
    catch {
        Write-LogMessage -Message "Error updating Licensing sheet: $($_.Exception.Message)" -Type Error
    }
}

function Update-EnhancedSharePointSheet {
    <#
    .SYNOPSIS
        Updates SharePoint Site sheet with permissions data
    #>
    param(
        [Parameter(Mandatory = $true)]
        $Workbook,
        
        [Parameter(Mandatory = $true)]
        [hashtable]$TenantData
    )
    
    try {
        Write-LogMessage -Message "Updating SharePoint Site sheet with permissions..." -Type Info
        
        $worksheet = $Workbook.Worksheets["SharePoint Site"]
        if (-not $worksheet) {
            Write-LogMessage -Message "SharePoint Site worksheet not found" -Type Warning
            return
        }
        
        $sites = $TenantData.SharePoint.Sites
        if (-not $sites -or $sites.Count -eq 0) {
            Write-LogMessage -Message "No SharePoint sites data to populate" -Type Warning
            return
        }
        
        # Check if we need to expand the table
        $availableRows = 27  # From template analysis
        $neededRows = $sites.Count
        
        if ($neededRows -gt $availableRows) {
            Write-LogMessage -Message "Expanding SharePoint table: need $neededRows rows, have $availableRows" -Type Info
            Expand-ExcelTable -Worksheet $worksheet -CurrentRows $availableRows -NeededRows $neededRows -StartRow 7
        }
        
        # Populate SharePoint data
        $startRow = 7  # Data starts at row 7
        $currentRow = $startRow
        
        foreach ($site in $sites) {
            try {
                # Map to template columns: B=SharePoint Site Name, C=Approver, D=Owners, E=Members, F=Read Only
                $worksheet.Cells.Item($currentRow, 2).Value2 = $site.DisplayName
                $worksheet.Cells.Item($currentRow, 3).Value2 = $site.Approver
                $worksheet.Cells.Item($currentRow, 4).Value2 = $site.Owners
                $worksheet.Cells.Item($currentRow, 5).Value2 = $site.Members
                $worksheet.Cells.Item($currentRow, 6).Value2 = $site.ReadOnly
                
                $currentRow++
            }
            catch {
                Write-LogMessage -Message "Error updating SharePoint row $currentRow`: $($_.Exception.Message)" -Type Warning -LogOnly
                $currentRow++
            }
        }
        
        Write-LogMessage -Message "Successfully updated SharePoint Site sheet with $($sites.Count) sites" -Type Success
    }
    catch {
        Write-LogMessage -Message "Error updating SharePoint Site sheet: $($_.Exception.Message)" -Type Error
    }
}

function Update-EnhancedConditionalAccessSheet {
    <#
    .SYNOPSIS
        Updates Conditional Access sheet with technical details and robust error handling
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
            Write-LogMessage -Message "EXCEL DEBUG: No Conditional Access policies to populate" -Type Warning
            return
        }
        
        # Start at row 9 (CA table starts here)
        $startRow = 9
        $currentRow = $startRow
        
        foreach ($policy in $policies) {
            try {
                Write-LogMessage -Message "EXCEL DEBUG: Processing CA policy $($policy.DisplayName) at row $currentRow" -Type Info -LogOnly
                
                $worksheet.Cells.Item($currentRow, 2).Value = if ($policy.DisplayName) { $policy.DisplayName } else { "Unnamed Policy" }
                $worksheet.Cells.Item($currentRow, 3).Value = if ($policy.TechnicalDetails) { $policy.TechnicalDetails } else { "State: $($policy.State)" }
                
                $currentRow++
            }
            catch {
                Write-LogMessage -Message "EXCEL DEBUG: Error updating CA policy row $currentRow - $($_.Exception.Message)" -Type Error
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

function Update-EnhancedIntuneAppsSheets {
    <#
    .SYNOPSIS
        Updates all Intune Apps sheets with platform-specific apps and expansion
    #>
    param(
        [Parameter(Mandatory = $true)]
        $Workbook,
        
        [Parameter(Mandatory = $true)]
        [hashtable]$TenantData
    )
    
    try {
        Write-LogMessage -Message "Updating Intune Apps sheets with platform categorization..." -Type Info
        
        $managedApps = $TenantData.Intune.ManagedApps
        if (-not $managedApps -or $managedApps.Count -eq 0) {
            Write-LogMessage -Message "No managed apps found to populate" -Type Warning
            return
        }
        
        # Group apps by platform
        $appsByPlatform = @{
            "Windows" = $managedApps | Where-Object { $_.Platform -eq "Windows" }
            "Android" = $managedApps | Where-Object { $_.Platform -eq "Android" }
            "iOS" = $managedApps | Where-Object { $_.Platform -eq "iOS" }
            "iPadOS" = $managedApps | Where-Object { $_.Platform -eq "iOS" }  # iOS apps work on iPadOS
            "macOS" = $managedApps | Where-Object { $_.Platform -eq "macOS" }
        }
        
        # Sheet mapping
        $sheetMapping = @{
            "Windows" = "Intune Windows Apps"
            "Android" = "Intune Android Apps"
            "iOS" = "Intune Apple IOS Apps "
            "iPadOS" = "Intune Apple iPadOS Apps "
            "macOS" = "Intune Mac OS Apps"
        }
        
        foreach ($platform in $appsByPlatform.Keys) {
            $sheetName = $sheetMapping[$platform]
            $platformApps = $appsByPlatform[$platform]
            
            Write-LogMessage -Message "Processing $platform apps for sheet '$sheetName'..." -Type Info
            
            $worksheet = $Workbook.Worksheets[$sheetName]
            if (-not $worksheet) {
                Write-LogMessage -Message "Worksheet '$sheetName' not found" -Type Warning
                continue
            }
            
            if (-not $platformApps -or $platformApps.Count -eq 0) {
                Write-LogMessage -Message "No $platform apps to populate" -Type Info
                continue
            }
            
            # Check if we need to expand the table
            $availableRows = 26  # From template analysis (27 total - 1 header)
            $neededRows = $platformApps.Count
            
            if ($neededRows -gt $availableRows) {
                Write-LogMessage -Message "Expanding $platform apps table: need $neededRows rows, have $availableRows" -Type Info
                Expand-ExcelTable -Worksheet $worksheet -CurrentRows $availableRows -NeededRows $neededRows -StartRow 8
            }
            
            # Populate app data
            $startRow = 8  # Data starts at row 8
            $currentRow = $startRow
            
            foreach ($app in $platformApps) {
                try {
                    # Map to template columns: B=Application Name, C=Required, D=Optional, E=Selected users only
                    $worksheet.Cells.Item($currentRow, 2).Value2 = $app.DisplayName
                    $worksheet.Cells.Item($currentRow, 3).Value2 = "" # Required - to be configured manually
                    $worksheet.Cells.Item($currentRow, 4).Value2 = "" # Optional - to be configured manually
                    $worksheet.Cells.Item($currentRow, 5).Value2 = "" # Selected users only - to be configured manually
                    
                    $currentRow++
                }
                catch {
                    Write-LogMessage -Message "Error updating $platform app row $currentRow`: $($_.Exception.Message)" -Type Warning -LogOnly
                    $currentRow++
                }
            }
            
            Write-LogMessage -Message "Successfully updated $sheetName with $($platformApps.Count) apps" -Type Success
        }
        
        Write-LogMessage -Message "Completed updating all Intune Apps sheets" -Type Success
    }
    catch {
        Write-LogMessage -Message "Error updating Intune Apps sheets: $($_.Exception.Message)" -Type Error
    }
}

function Add-SecurityGroupsSheet {
    <#
    .SYNOPSIS
        Adds a new Security Groups sheet to the workbook
    #>
    param(
        [Parameter(Mandatory = $true)]
        $Workbook,
        
        [Parameter(Mandatory = $true)]
        [hashtable]$TenantData
    )
    
    try {
        Write-LogMessage -Message "Adding Security Groups sheet..." -Type Info
        
        $securityGroups = $TenantData.Groups.SecurityGroups
        if (-not $securityGroups -or $securityGroups.Count -eq 0) {
            Write-LogMessage -Message "No security groups to add" -Type Warning
            return
        }
        
        # Add new worksheet
        $newWorksheet = $Workbook.Worksheets.Add()
        $newWorksheet.Name = "Security Groups"
        
        # Set up headers
        $newWorksheet.Cells.Item(1, 1).Value2 = "Security Groups"
        $newWorksheet.Cells.Item(1, 1).Font.Bold = $true
        $newWorksheet.Cells.Item(1, 1).Font.Size = 14
        
        $newWorksheet.Cells.Item(3, 1).Value2 = "Security groups are used for access control, license assignment, and conditional access policies."
        
        # Column headers
        $newWorksheet.Cells.Item(5, 2).Value2 = "Group Name"
        $newWorksheet.Cells.Item(5, 3).Value2 = "Description"
        $newWorksheet.Cells.Item(5, 4).Value2 = "Type"
        $newWorksheet.Cells.Item(5, 5).Value2 = "Members"
        $newWorksheet.Cells.Item(5, 6).Value2 = "Member Count"
        
        # Make headers bold
        $headerRange = $newWorksheet.Range("B5:F5")
        $headerRange.Font.Bold = $true
        $headerRange.Interior.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::LightGray)
        
        # Populate security groups data
        $startRow = 6
        $currentRow = $startRow
        
        foreach ($group in $securityGroups) {
            try {
                $newWorksheet.Cells.Item($currentRow, 2).Value2 = $group.DisplayName
                $newWorksheet.Cells.Item($currentRow, 3).Value2 = $group.Description
                
                # Determine group type
                $groupType = "Security Group"
                if ($group.GroupTypes -and $group.GroupTypes -like "*Unified*") {
                    $groupType = "Microsoft 365 Group"
                } elseif ($group.MailEnabled) {
                    $groupType = "Mail-Enabled Security Group"
                }
                
                $newWorksheet.Cells.Item($currentRow, 4).Value2 = $groupType
                $newWorksheet.Cells.Item($currentRow, 5).Value2 = $group.Members
                $newWorksheet.Cells.Item($currentRow, 6).Value2 = $group.MemberCount
                
                $currentRow++
            }
            catch {
                Write-LogMessage -Message "Error adding security group row $currentRow`: $($_.Exception.Message)" -Type Warning -LogOnly
                $currentRow++
            }
        }
        
        # Apply table formatting
        $tableRange = $newWorksheet.Range("B5:F$($currentRow - 1)")
        $tableRange.Borders.LineStyle = 1
        $tableRange.Borders.Weight = 2
        
        # Auto-fit columns
        $newWorksheet.Columns("B:F").AutoFit()
        
        Write-LogMessage -Message "Successfully added Security Groups sheet with $($securityGroups.Count) groups" -Type Success
    }
    catch {
        Write-LogMessage -Message "Error adding Security Groups sheet: $($_.Exception.Message)" -Type Error
    }
}

function Update-DistributionListSheet {
    <#
    .SYNOPSIS
        Updates the Distribution list sheet with distribution groups
    #>
    param(
        [Parameter(Mandatory = $true)]
        $Workbook,
        
        [Parameter(Mandatory = $true)]
        [hashtable]$TenantData
    )
    
    try {
        Write-LogMessage -Message "Updating Distribution list sheet..." -Type Info
        
        $worksheet = $Workbook.Worksheets["Distribution list"]
        if (-not $worksheet) {
            Write-LogMessage -Message "Distribution list worksheet not found" -Type Warning
            return
        }
        
        $distributionGroups = $TenantData.Groups.DistributionGroups
        if (-not $distributionGroups -or $distributionGroups.Count -eq 0) {
            Write-LogMessage -Message "No distribution groups to populate" -Type Warning
            return
        }
        
        # Find the data start row (look for existing structure)
        $startRow = 6  # Estimated based on template structure
        $currentRow = $startRow
        
        foreach ($group in $distributionGroups) {
            try {
                # Map to template structure - this may need adjustment based on actual template
                $worksheet.Cells.Item($currentRow, 2).Value2 = $group.DisplayName
                $worksheet.Cells.Item($currentRow, 3).Value2 = $group.Members
                
                $currentRow++
            }
            catch {
                Write-LogMessage -Message "Error updating distribution group row $currentRow`: $($_.Exception.Message)" -Type Warning -LogOnly
                $currentRow++
            }
        }
        
        Write-LogMessage -Message "Successfully updated Distribution list sheet with $($distributionGroups.Count) groups" -Type Success
    }
    catch {
        Write-LogMessage -Message "Error updating Distribution list sheet: $($_.Exception.Message)" -Type Error
    }
}

function Expand-ExcelTable {
    <#
    .SYNOPSIS
        Expands an Excel table by inserting additional rows while maintaining formatting
    #>
    param(
        [Parameter(Mandatory = $true)]
        $Worksheet,
        
        [Parameter(Mandatory = $true)]
        [int]$CurrentRows,
        
        [Parameter(Mandatory = $true)]
        [int]$NeededRows,
        
        [Parameter(Mandatory = $true)]
        [int]$StartRow
    )
    
    try {
        if ($NeededRows -le $CurrentRows) {
            return  # No expansion needed
        }
        
        $rowsToAdd = $NeededRows - $CurrentRows
        $insertionPoint = $StartRow + $CurrentRows
        
        Write-LogMessage -Message "Expanding table: inserting $rowsToAdd rows at row $insertionPoint" -Type Info -LogOnly
        
        # Insert rows
        $insertRange = $Worksheet.Range("$insertionPoint`:$insertionPoint")
        for ($i = 0; $i -lt $rowsToAdd; $i++) {
            $insertRange.Insert([Microsoft.Office.Interop.Excel.XlInsertShiftDirection]::xlShiftDown)
        }
        
        # Copy formatting from the last row of the original table
        $sourceRow = $StartRow + $CurrentRows - 1
        $targetStartRow = $insertionPoint
        $targetEndRow = $insertionPoint + $rowsToAdd - 1
        
        $sourceRange = $Worksheet.Range("$sourceRow`:$sourceRow")
        $targetRange = $Worksheet.Range("$targetStartRow`:$targetEndRow")
        
        $sourceRange.Copy()
        $targetRange.PasteSpecial([Microsoft.Office.Interop.Excel.XlPasteType]::xlPasteFormats)
        
        # Clear clipboard
        $Worksheet.Application.CutCopyMode = $false
        
        Write-LogMessage -Message "Successfully expanded table with $rowsToAdd additional rows" -Type Success -LogOnly
    }
    catch {
        Write-LogMessage -Message "Error expanding Excel table: $($_.Exception.Message)" -Type Warning -LogOnly
    }
}

# === Additional Helper Functions ===

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
