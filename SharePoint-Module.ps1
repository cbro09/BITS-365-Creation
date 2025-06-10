# === SharePoint.ps1 ===
# SharePoint Online configuration with Root Site as Hub and automated security group assignment

# SharePoint configuration
$SharePointConfig = @{
    SiteTemplate = "SITEPAGEPUBLISHING#0"
    DefaultSites = @("HR", "Processes", "Templates", "Documents", "Policies")
    StorageQuota = 1024
}

function New-TenantSharePoint {
    Write-LogMessage -Message "Starting SharePoint configuration..." -Type Info
    
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
            Disconnect-SPOService -ErrorAction SilentlyContinue | Out-Null
        }
        catch {
            # Ignore disconnect errors
        }
        
        # STEP 5: Force load modules including PnP PowerShell for better group management
        $sharePointModules = @(
            'Microsoft.Graph.Groups',
            'Microsoft.Graph.Identity.DirectoryManagement',
            'Microsoft.Online.SharePoint.PowerShell'
        )
        
        Write-LogMessage -Message "Loading SharePoint modules..." -Type Info
        foreach ($module in $sharePointModules) {
            try {
                Get-Module $module | Remove-Module -Force -ErrorAction SilentlyContinue
                Import-Module -Name $module -Force -ErrorAction Stop
                $moduleInfo = Get-Module $module
                Write-LogMessage -Message "Loaded $module version $($moduleInfo.Version)" -Type Success -LogOnly
            }
            catch {
                Write-LogMessage -Message "Failed to load $module module - $($_.Exception.Message)" -Type Error
                return $false
            }
        }
        
        # Check if PnP PowerShell is available for better group management
        $pnpAvailable = $false
        try {
            Import-Module PnP.PowerShell -Force -ErrorAction Stop
            $pnpAvailable = $true
            Write-LogMessage -Message "PnP PowerShell module loaded - will use for security group assignment" -Type Success
        }
        catch {
            Write-LogMessage -Message "PnP PowerShell not available - will provide manual instructions" -Type Warning
        }
        
        # STEP 6: Connect to Microsoft Graph
        $sharePointScopes = @(
            "Group.ReadWrite.All",
            "Directory.ReadWrite.All"
        )
        
        Write-LogMessage -Message "Connecting to Microsoft Graph with SharePoint scopes..." -Type Info
        Connect-MgGraph -Scopes $sharePointScopes -ErrorAction Stop | Out-Null
        
        $context = Get-MgContext
        Write-LogMessage -Message "Connected to Microsoft Graph as $($context.Account)" -Type Success
        
        # STEP 7: Validate prerequisites
        Write-LogMessage -Message "Validating SharePoint prerequisites..." -Type Info
        
        if (-not $script:TenantState) {
            Write-LogMessage -Message "ERROR: TenantState not initialized. Please run 'Connect to Microsoft Graph and Verify Tenant' first." -Type Error
            return $false
        }
        
        if ([string]::IsNullOrWhiteSpace($script:TenantState.TenantName)) {
            Write-LogMessage -Message "ERROR: Tenant name not found in TenantState" -Type Error
            return $false
        }
        
        Write-LogMessage -Message "Prerequisites validation completed" -Type Success
        
        # Get SharePoint URLs and email configuration
        $customerName = $script:TenantState.TenantName
        Write-Host "SharePoint URL Configuration" -ForegroundColor Yellow
        Write-Host "Example: If your tenant is 'm365x36060197.sharepoint.com', enter 'm365x36060197'" -ForegroundColor Cyan
        $tenantName = Read-Host "Enter your SharePoint tenant name (without .sharepoint.com)"
        
        # Construct URLs automatically - Root site as hub
        $adminUrl = "https://$tenantName-admin.sharepoint.com"
        $tenantUrl = "https://$tenantName.sharepoint.com"
        $hubSiteUrl = $tenantUrl  # Root site as hub
        
        # Get and validate email addresses
        $ownerEmail = $script:TenantState.AdminEmail
        $currentUserEmail = $context.Account
        
        Write-LogMessage -Message "Owner Email from TenantState: '$ownerEmail'" -Type Info
        Write-LogMessage -Message "Current User Email from Context: '$currentUserEmail'" -Type Info
        
        if ([string]::IsNullOrWhiteSpace($ownerEmail)) {
            Write-LogMessage -Message "Owner email is empty, prompting for input..." -Type Warning
            $ownerEmail = Read-Host "Enter the site owner email address"
        }
        
        if ([string]::IsNullOrWhiteSpace($currentUserEmail)) {
            Write-LogMessage -Message "Current user email is empty, using owner email..." -Type Warning
            $currentUserEmail = $ownerEmail
        }
        
        Write-LogMessage -Message "Using Owner Email: '$ownerEmail'" -Type Info
        Write-LogMessage -Message "Using Current User Email: '$currentUserEmail'" -Type Info
        Write-LogMessage -Message "SharePoint Admin URL: $adminUrl" -Type Info
        Write-LogMessage -Message "SharePoint Tenant URL: $tenantUrl" -Type Info
        Write-LogMessage -Message "Hub Site (Root): $hubSiteUrl" -Type Info
        
        # Connect to SharePoint Online
        Write-LogMessage -Message "Connecting to SharePoint Online Admin Center..." -Type Info
        try {
            Connect-SPOService -Url $adminUrl -ModernAuth $true
            Write-LogMessage -Message "Successfully connected to SharePoint Online" -Type Success
            
            try {
                $tenantInfo = Get-SPOTenant -ErrorAction Stop
                Write-LogMessage -Message "SharePoint Administrator permissions verified" -Type Success
            }
            catch {
                Write-LogMessage -Message "WARNING: Connected but may not have SharePoint Administrator permissions" -Type Warning
            }
        }
        catch {
            Write-LogMessage -Message "Failed to connect - $($_.Exception.Message)" -Type Error
            return $false
        }
        
        # Verify root site exists and get its information
        Write-LogMessage -Message "Verifying root site for hub registration..." -Type Info
        try {
            $rootSiteInfo = Get-SPOSite -Identity $hubSiteUrl -Detailed -ErrorAction Stop
            Write-LogMessage -Message "Root site found - Title: '$($rootSiteInfo.Title)', Template: $($rootSiteInfo.Template)" -Type Success
            
            # Update root site title to reflect hub status
            $hubSiteTitle = "$customerName Hub"
            if ($rootSiteInfo.Title -ne $hubSiteTitle) {
                try {
                    Set-SPOSite -Identity $hubSiteUrl -Title $hubSiteTitle -ErrorAction Stop
                    Write-LogMessage -Message "Root site title updated to: '$hubSiteTitle'" -Type Success
                }
                catch {
                    Write-LogMessage -Message "Could not update root site title - $($_.Exception.Message)" -Type Warning
                }
            }
        }
        catch {
            Write-LogMessage -Message "ERROR: Cannot access root site at $hubSiteUrl - $($_.Exception.Message)" -Type Error
            return $false
        }
        
        # Register root site as Hub Site
        Write-LogMessage -Message "Registering root site as Hub..." -Type Info
        try {
            $existingHubSites = Get-SPOHubSite -ErrorAction SilentlyContinue
            $isAlreadyHub = $existingHubSites | Where-Object { $_.SiteUrl -eq $hubSiteUrl }
            
            if ($isAlreadyHub) {
                Write-LogMessage -Message "Root site is already registered as a Hub site" -Type Warning
                $hubSiteId = $isAlreadyHub.ID
            }
            else {
                Write-LogMessage -Message "Root site template: $($rootSiteInfo.Template), Status: $($rootSiteInfo.Status)" -Type Info
                
                # Check if root site is associated with another hub (unlikely but possible)
                if ($rootSiteInfo.HubSiteId -and $rootSiteInfo.HubSiteId -ne "00000000-0000-0000-0000-000000000000") {
                    Write-LogMessage -Message "Root site is associated with another hub. Removing association first..." -Type Warning
                    Remove-SPOHubSiteAssociation -Site $hubSiteUrl -ErrorAction Stop
                    Start-Sleep -Seconds 10
                }
                
                $registrationSuccess = $false
                
                # Try different registration approaches
                try {
                    Register-SPOHubSite -Site $hubSiteUrl
                    $registrationSuccess = $true
                    Write-LogMessage -Message "Root site registered as hub successfully (no principals)" -Type Success
                }
                catch {
                    try {
                        Register-SPOHubSite -Site $hubSiteUrl -Principals @()
                        $registrationSuccess = $true
                        Write-LogMessage -Message "Root site registered as hub successfully (empty array)" -Type Success
                    }
                    catch {
                        try {
                            Register-SPOHubSite -Site $hubSiteUrl -Principals $ownerEmail
                            $registrationSuccess = $true
                            Write-LogMessage -Message "Root site registered as hub successfully (with owner)" -Type Success
                        }
                        catch {
                            Write-LogMessage -Message "All hub registration approaches failed: $($_.Exception.Message)" -Type Error
                        }
                    }
                }
                
                if ($registrationSuccess) {
                    Start-Sleep -Seconds 15
                    $verifyHub = Get-SPOHubSite | Where-Object { $_.SiteUrl -eq $hubSiteUrl }
                    if ($verifyHub) {
                        Write-LogMessage -Message "Root site hub registration verified successfully" -Type Success
                        $hubSiteId = $verifyHub.ID
                    }
                    else {
                        Write-LogMessage -Message "Root site hub registration verification failed" -Type Error
                        $hubSiteId = $null
                    }
                }
                else {
                    $hubSiteId = $null
                }
            }
        }
        catch {
            Write-LogMessage -Message "Root site hub registration failed - $($_.Exception.Message)" -Type Error
            $hubSiteId = $null
        }
        
        # Create spoke sites configuration
        $spokeSites = @()
        foreach ($siteName in $SharePointConfig.DefaultSites) {
            $spokeSites += @{ 
                Name = $siteName
                URL = "$tenantUrl/sites/$($siteName.ToLower())" 
            }
        }
        
        # Create security groups for spoke sites
        $securityGroups = @{}
        Write-LogMessage -Message "Creating security groups for spoke sites..." -Type Info
        
        foreach ($site in $spokeSites) {
            $siteName = $site.Name
            Write-LogMessage -Message "Creating security groups for site: $siteName" -Type Info
            
            foreach ($groupType in @("Members", "Owners", "Visitors")) {
                $groupName = "$siteName SharePoint $groupType"
                $existingGroup = Get-MgGroup -Filter "DisplayName eq '$groupName'" -ErrorAction SilentlyContinue
                
                if ($existingGroup -eq $null) {
                    try {
                        $newGroup = New-MgGroup -DisplayName $groupName -MailEnabled:$false -MailNickname "$siteName-SPO-$groupType" -SecurityEnabled:$true
                        $securityGroups["$siteName-$groupType"] = $newGroup.Id
                        Write-LogMessage -Message "Security group created: $groupName" -Type Success
                    } catch {
                        Write-LogMessage -Message "Failed to create security group: $groupName - $($_.Exception.Message)" -Type Error
                        continue
                    }
                } else {
                    $securityGroups["$siteName-$groupType"] = $existingGroup.Id
                    Write-LogMessage -Message "Security group already exists: $groupName" -Type Warning
                }
            }
        }
        
        # Wait for group propagation
        Write-LogMessage -Message "Waiting for security groups to propagate (2 minutes)..." -Type Info
        Start-Sleep -Seconds 120
        
        # Create spoke sites
        $createdSites = @()
        Write-LogMessage -Message "Creating spoke sites..." -Type Info
        
        foreach ($site in $spokeSites) {
            $siteUrl = $site.URL
            $siteName = $site.Name
            
            try {
                $existingSpokeSite = $null
                try {
                    $existingSpokeSite = Get-SPOSite -Identity $siteUrl -ErrorAction Stop
                    Write-LogMessage -Message "$siteName site already exists: $siteUrl" -Type Warning
                    $createdSites += $siteUrl
                }
                catch {
                    Write-LogMessage -Message "Creating $siteName site..." -Type Info
                    New-SPOSite -Url $siteUrl -Owner $ownerEmail -StorageQuota $SharePointConfig.StorageQuota -Title "$siteName" -Template $SharePointConfig.SiteTemplate
                    Write-LogMessage -Message "$siteName site created: $siteUrl" -Type Success
                    $createdSites += $siteUrl
                }
            }
            catch {
                Write-LogMessage -Message "Failed to create $siteName site - $($_.Exception.Message)" -Type Error
                continue
            }
        }
        
        # Wait for site provisioning
        Write-LogMessage -Message "Waiting for SharePoint sites to provision (3 minutes)..." -Type Info
        Start-Sleep -Seconds 180
        
        # Associate spoke sites with root hub
        if ($hubSiteId) {
            Write-LogMessage -Message "Associating spoke sites with root hub..." -Type Info
            foreach ($siteUrl in $createdSites) {
                try {
                    $siteInfo = Get-SPOSite -Identity $siteUrl -Detailed
                    if ($siteInfo.HubSiteId -eq $hubSiteId) {
                        $siteName = ($spokeSites | Where-Object { $_.URL -eq $siteUrl }).Name
                        Write-LogMessage -Message "$siteName already associated with root hub" -Type Warning
                    }
                    else {
                        Add-SPOHubSiteAssociation -Site $siteUrl -HubSite $hubSiteUrl
                        $siteName = ($spokeSites | Where-Object { $_.URL -eq $siteUrl }).Name
                        Write-LogMessage -Message "$siteName associated with root hub" -Type Success
                    }
                }
                catch {
                    $siteName = ($spokeSites | Where-Object { $_.URL -eq $siteUrl }).Name
                    Write-LogMessage -Message "Failed to associate $siteName with root hub - $($_.Exception.Message)" -Type Warning
                }
                Start-Sleep -Seconds 2
            }
        }
        
        # Set site collection administrators for spoke sites
        Write-LogMessage -Message "Configuring spoke site permissions..." -Type Info
        foreach ($siteUrl in $createdSites) {
            try {
                Set-SPOUser -Site $siteUrl -LoginName $ownerEmail -IsSiteCollectionAdmin $true -ErrorAction Stop
                if ($currentUserEmail -ne $ownerEmail) {
                    Set-SPOUser -Site $siteUrl -LoginName $currentUserEmail -IsSiteCollectionAdmin $true -ErrorAction Stop
                }
                $siteName = ($spokeSites | Where-Object { $_.URL -eq $siteUrl }).Name
                Write-LogMessage -Message "Set site collection admins for $siteName" -Type Success
            }
            catch {
                $siteName = ($spokeSites | Where-Object { $_.URL -eq $siteUrl }).Name
                Write-LogMessage -Message "Failed to set site collection admin for $siteName - $($_.Exception.Message)" -Type Warning
            }
        }
        
        # Configure root site permissions
        Write-LogMessage -Message "Configuring root hub site permissions..." -Type Info
        try {
            Set-SPOUser -Site $hubSiteUrl -LoginName $ownerEmail -IsSiteCollectionAdmin $true -ErrorAction Stop
            if ($currentUserEmail -ne $ownerEmail) {
                Set-SPOUser -Site $hubSiteUrl -LoginName $currentUserEmail -IsSiteCollectionAdmin $true -ErrorAction Stop
            }
            Write-LogMessage -Message "Set site collection admins for root hub site" -Type Success
        }
        catch {
            Write-LogMessage -Message "Failed to set site collection admin for root hub - $($_.Exception.Message)" -Type Warning
        }
        
        # AUTOMATED SECURITY GROUP ASSIGNMENT using PnP PowerShell
        if ($pnpAvailable) {
            Write-LogMessage -Message "Starting automated security group assignment using PnP PowerShell..." -Type Info
            $groupAssignmentSuccess = $true
            
            foreach ($site in $spokeSites) {
                $siteUrl = $site.URL
                $siteName = $site.Name
                
                Write-LogMessage -Message "Configuring security groups for $siteName..." -Type Info
                
                try {
                    # Connect to the specific site with PnP
                    Connect-PnPOnline -Url $siteUrl -Interactive -ErrorAction Stop
                    Write-LogMessage -Message "Connected to $siteName with PnP PowerShell" -Type Success
                    
                    foreach ($groupType in @("Members", "Owners", "Visitors")) {
                        $groupKey = "$siteName-$groupType"
                        $groupDisplayName = "$siteName SharePoint $groupType"
                        $sharePointGroupName = "$siteName $groupType"
                        
                        if ($securityGroups.ContainsKey($groupKey)) {
                            try {
                                # Add Azure AD security group to SharePoint group
                                Add-PnPUser -LoginName $groupDisplayName -Group $sharePointGroupName -ErrorAction Stop
                                Write-LogMessage -Message "Added '$groupDisplayName' to $siteName $sharePointGroupName" -Type Success
                            }
                            catch {
                                if ($_.Exception.Message -like "*already exists*" -or $_.Exception.Message -like "*already a member*") {
                                    Write-LogMessage -Message "'$groupDisplayName' already exists in $siteName $sharePointGroupName" -Type Warning
                                }
                                else {
                                    Write-LogMessage -Message "Failed to add '$groupDisplayName' to $siteName: $($_.Exception.Message)" -Type Warning
                                    $groupAssignmentSuccess = $false
                                }
                            }
                        }
                    }
                    
                    # Disconnect from this site
                    Disconnect-PnPOnline
                    
                }
                catch {
                    Write-LogMessage -Message "Failed to connect to $siteName with PnP PowerShell: $($_.Exception.Message)" -Type Error
                    $groupAssignmentSuccess = $false
                }
                
                Start-Sleep -Seconds 2
            }
            
            if ($groupAssignmentSuccess) {
                Write-LogMessage -Message "Automated security group assignment completed successfully!" -Type Success
            }
            else {
                Write-LogMessage -Message "Some security group assignments failed - manual configuration may be needed" -Type Warning
            }
        }
        else {
            Write-LogMessage -Message "PnP PowerShell not available - security groups created but not assigned to sites" -Type Warning
        }
        
        # Final results and guidance
        Write-LogMessage -Message "SharePoint configuration completed!" -Type Success
        Write-LogMessage -Message "Root Hub Site: '$customerName Hub' at $hubSiteUrl" -Type Info
        
        Write-LogMessage -Message "Spoke site URLs created:" -Type Info
        foreach ($siteUrl in $createdSites) {
            $siteName = ($spokeSites | Where-Object { $_.URL -eq $siteUrl }).Name
            Write-LogMessage -Message "   $siteName: $siteUrl" -Type Info
        }
        
        if (-not $pnpAvailable -or -not $groupAssignmentSuccess) {
            Write-LogMessage -Message "Manual security group configuration options:" -Type Info
            Write-LogMessage -Message "1. SharePoint Admin Center: https://$tenantName-admin.sharepoint.com" -Type Info
            Write-LogMessage -Message "2. Install PnP PowerShell: Install-Module PnP.PowerShell -Scope CurrentUser" -Type Info
            Write-LogMessage -Message "3. For each site, add the corresponding security groups to SharePoint groups" -Type Info
        }
        
        Write-LogMessage -Message "Navigation and search integration will be available from the root site hub" -Type Info
        
        # Clean disconnect
        try {
            Disconnect-SPOService
        }
        catch {
            # Ignore disconnect errors
        }
        
        return $true
    }
    catch {
        Write-LogMessage -Message "Error in SharePoint configuration - $($_.Exception.Message)" -Type Error
        return $false
    }
}