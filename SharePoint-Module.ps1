# === SharePoint.ps1 ===
# SharePoint Online configuration and site creation functions - Fixed Hub Site Registration

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
        
        # STEP 5: Force load ONLY the exact modules needed for SharePoint
        $sharePointModules = @(
            'Microsoft.Graph.Groups',
            'Microsoft.Graph.Identity.DirectoryManagement',
            'Microsoft.Online.SharePoint.PowerShell'
        )
        
        Write-LogMessage -Message "Loading ONLY SharePoint modules in exact order..." -Type Info
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
        
        # STEP 6: Connect with EXACT scopes needed for SharePoint
        $sharePointScopes = @(
            "Group.ReadWrite.All",
            "Directory.ReadWrite.All"
        )
        
        Write-LogMessage -Message "Connecting to Microsoft Graph with SharePoint scopes..." -Type Info
        Connect-MgGraph -Scopes $sharePointScopes -ErrorAction Stop | Out-Null
        
        $context = Get-MgContext
        Write-LogMessage -Message "Connected to Microsoft Graph as $($context.Account)" -Type Success
        
        # STEP 7: SharePoint configuration with improved hub site handling
        # Get SharePoint URLs - simplified input
        $customerName = $script:TenantState.TenantName
        Write-Host "SharePoint URL Configuration" -ForegroundColor Yellow
        Write-Host "Example: If your tenant is 'm365x36060197.sharepoint.com', enter 'm365x36060197'" -ForegroundColor Cyan
        $tenantName = Read-Host "Enter your SharePoint tenant name (without .sharepoint.com)"
        
        # Construct URLs automatically
        $adminUrl = "https://$tenantName-admin.sharepoint.com"
        $tenantUrl = "https://$tenantName.sharepoint.com"
        
        # Get and validate email addresses
        $ownerEmail = $script:TenantState.AdminEmail
        $currentUserEmail = $context.Account
        
        # Debug email addresses
        Write-LogMessage -Message "Owner Email from TenantState: '$ownerEmail'" -Type Info
        Write-LogMessage -Message "Current User Email from Context: '$currentUserEmail'" -Type Info
        
        # Validate and get proper email if needed
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
        
        # Connect to SharePoint Online Admin Center with modern authentication
        Write-LogMessage -Message "Connecting to SharePoint Online Admin Center with modern authentication..." -Type Info
        try {
            Connect-SPOService -Url $adminUrl -ModernAuth $true
            Write-LogMessage -Message "Successfully connected to SharePoint Online" -Type Success
        }
        catch {
            Write-LogMessage -Message "Failed to connect with modern auth, trying standard connection..." -Type Warning
            Connect-SPOService -Url $adminUrl
        }
        
        # Create a Hub Site
        $hubSiteTitle = "$customerName Hub"
        $hubSiteUrl = "$tenantUrl/sites/corporatehub"
        
        # Check if the Hub Site already exists, if not, create one
        Write-LogMessage -Message "Creating hub site: $hubSiteUrl" -Type Info
        try {
            # Use more reliable method to check if site exists
            $existingHubSite = $null
            try {
                $existingHubSite = Get-SPOSite -Identity $hubSiteUrl -ErrorAction Stop
                Write-LogMessage -Message "Hub site already exists: $hubSiteUrl" -Type Warning
            }
            catch {
                # Site doesn't exist, proceed with creation
                Write-LogMessage -Message "Creating new hub site..." -Type Info
                New-SPOSite -Url $hubSiteUrl -Owner $ownerEmail -StorageQuota $SharePointConfig.StorageQuota -Title $hubSiteTitle -Template $SharePointConfig.SiteTemplate
                Write-LogMessage -Message "Hub site created successfully: $hubSiteUrl" -Type Success
                
                # Wait for site to be fully provisioned
                Write-LogMessage -Message "Waiting for hub site to be fully provisioned..." -Type Info
                Start-Sleep -Seconds 60
            }
        }
        catch {
            Write-LogMessage -Message "Error with hub site creation - $($_.Exception.Message)" -Type Error
            return $false
        }
        
        # IMPROVED Hub Site Registration with proper error handling
        Write-LogMessage -Message "Registering site as Hub..." -Type Info
        try {
            # Step 1: Check current hub site status
            $existingHubSites = Get-SPOHubSite -ErrorAction SilentlyContinue
            $isAlreadyHub = $existingHubSites | Where-Object { $_.SiteUrl -eq $hubSiteUrl }
            
            if ($isAlreadyHub) {
                Write-LogMessage -Message "Site is already registered as a Hub site" -Type Warning
                $hubSiteId = $isAlreadyHub.ID
            }
            else {
                # Step 2: Verify site exists and get detailed info
                $siteInfo = Get-SPOSite -Identity $hubSiteUrl -Detailed -ErrorAction Stop
                Write-LogMessage -Message "Site template: $($siteInfo.Template), Status: $($siteInfo.Status)" -Type Info
                
                # Step 3: Check if site is associated with another hub
                if ($siteInfo.HubSiteId -and $siteInfo.HubSiteId -ne "00000000-0000-0000-0000-000000000000") {
                    Write-LogMessage -Message "Site is associated with another hub. Removing association first..." -Type Warning
                    Remove-SPOHubSiteAssociation -Site $hubSiteUrl -ErrorAction Stop
                    Start-Sleep -Seconds 10
                }
                
                # Step 4: Register as hub site using the method from research
                Write-LogMessage -Message "Registering as hub site with null principals..." -Type Info
                try {
                    # Use the exact method from research - pass $null for Principals
                    Register-SPOHubSite -Site $hubSiteUrl -Principals $null
                    Write-LogMessage -Message "Hub site registered successfully using null principals" -Type Success
                }
                catch {
                    # Fallback: Try with owner email if null fails
                    if ($_.Exception.Message -like "*Cannot bind argument to parameter 'Principals'*") {
                        Write-LogMessage -Message "Null principals failed, trying with owner email..." -Type Warning
                        Register-SPOHubSite -Site $hubSiteUrl -Principals @($ownerEmail)
                        Write-LogMessage -Message "Hub site registered successfully with owner email" -Type Success
                    }
                    else {
                        throw $_
                    }
                }
                
                # Step 5: Verify registration was successful
                Start-Sleep -Seconds 15
                $verifyHub = Get-SPOHubSite | Where-Object { $_.SiteUrl -eq $hubSiteUrl }
                if ($verifyHub) {
                    Write-LogMessage -Message "Hub site registration verified successfully" -Type Success
                    $hubSiteId = $verifyHub.ID
                }
                else {
                    Write-LogMessage -Message "Hub site registration verification failed" -Type Error
                    $hubSiteId = $null
                }
            }
        }
        catch {
            Write-LogMessage -Message "Hub site registration failed - $($_.Exception.Message)" -Type Error
            Write-LogMessage -Message "Continuing without hub site functionality..." -Type Warning
            $hubSiteId = $null
        }
        
        # Create spokes sites array
        $spokeSites = @()
        foreach ($siteName in $SharePointConfig.DefaultSites) {
            $spokeSites += @{ 
                Name = $siteName
                URL = "$tenantUrl/sites/$($siteName.ToLower())" 
            }
        }
        
        # Create security groups for each site with progress tracking
        $securityGroups = @{}
        $totalGroups = $spokeSites.Count * 3 # 3 groups per site (Members, Owners, Visitors)
        $currentGroup = 0
        
        Write-Progress -Id 1 -Activity "Creating Security Groups" -Status "Starting group creation..." -PercentComplete 0
        
        foreach ($site in $spokeSites) {
            $siteName = $site.Name
            Write-LogMessage -Message "Creating security groups for site: $siteName" -Type Info
            
            foreach ($groupType in @("Members", "Owners", "Visitors")) {
                $currentGroup++
                $percentComplete = [math]::Round(($currentGroup / $totalGroups) * 100, 2)
                
                $groupName = "$siteName SharePoint $groupType"
                Write-Progress -Id 1 -Activity "Creating Security Groups" -Status "Creating: $groupName" -PercentComplete $percentComplete -CurrentOperation "Group $currentGroup of $totalGroups"
                
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
        
        # Complete security groups progress
        Write-Progress -Id 1 -Activity "Creating Security Groups" -Completed
        
        # Wait for security groups to propagate
        Write-LogMessage -Message "Waiting for security groups to propagate (2 minutes)..." -Type Info
        $propagationTime = 120
        for ($i = 1; $i -le $propagationTime; $i++) {
            $percentComplete = [math]::Round(($i / $propagationTime) * 100, 2)
            $remaining = $propagationTime - $i
            Write-Progress -Id 2 -Activity "Group Propagation" -Status "Waiting for Azure AD synchronization..." -PercentComplete $percentComplete -SecondsRemaining $remaining
            Start-Sleep -Seconds 1
        }
        Write-Progress -Id 2 -Activity "Group Propagation" -Completed
        
        # Create spoke sites with progress tracking
        $totalSites = $spokeSites.Count
        $currentSite = 0
        $createdSites = @()
        
        Write-Progress -Id 3 -Activity "Creating SharePoint Sites" -Status "Starting site creation..." -PercentComplete 0
        
        foreach ($site in $spokeSites) {
            $currentSite++
            $percentComplete = [math]::Round(($currentSite / $totalSites) * 100, 2)
            
            $siteUrl = $site.URL
            $siteName = $site.Name
            
            Write-Progress -Id 3 -Activity "Creating SharePoint Sites" -Status "Creating: $siteName" -PercentComplete $percentComplete -CurrentOperation "Site $currentSite of $totalSites"
            
            # Check if the Spoke Site already exists, if not, create one
            try {
                $existingSpokeSite = $null
                try {
                    $existingSpokeSite = Get-SPOSite -Identity $siteUrl -ErrorAction Stop
                    Write-LogMessage -Message "$siteName site already exists: $siteUrl" -Type Warning
                    $createdSites += $siteUrl
                }
                catch {
                    # Site doesn't exist, create it
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
        
        # Complete site creation progress
        Write-Progress -Id 3 -Activity "Creating SharePoint Sites" -Completed
        
        # Wait for sites to provision
        Write-LogMessage -Message "Waiting for SharePoint sites to provision (3 minutes)..." -Type Info
        $provisionTime = 180
        for ($i = 1; $i -le $provisionTime; $i++) {
            $percentComplete = [math]::Round(($i / $provisionTime) * 100, 2)
            $remaining = $provisionTime - $i
            Write-Progress -Id 4 -Activity "Site Provisioning" -Status "Waiting for SharePoint sites to fully provision..." -PercentComplete $percentComplete -SecondsRemaining $remaining
            Start-Sleep -Seconds 1
        }
        Write-Progress -Id 4 -Activity "Site Provisioning" -Completed
        
        # Associate sites with hub (only if hub was registered successfully)
        if ($hubSiteId) {
            Write-LogMessage -Message "Associating sites with hub..." -Type Info
            $totalAssociations = $createdSites.Count
            $currentAssociation = 0
            
            Write-Progress -Id 5 -Activity "Hub Site Associations" -Status "Starting hub associations..." -PercentComplete 0
            
            foreach ($siteUrl in $createdSites) {
                $currentAssociation++
                $percentComplete = [math]::Round(($currentAssociation / $totalAssociations) * 100, 2)
                $siteName = ($spokeSites | Where-Object { $_.URL -eq $siteUrl }).Name
                
                Write-Progress -Id 5 -Activity "Hub Site Associations" -Status "Associating: $siteName" -PercentComplete $percentComplete -CurrentOperation "Association $currentAssociation of $totalAssociations"
                
                try {
                    # Check if site is already associated
                    $siteInfo = Get-SPOSite -Identity $siteUrl -Detailed
                    if ($siteInfo.HubSiteId -eq $hubSiteId) {
                        Write-LogMessage -Message "$siteName already associated with hub" -Type Warning
                    }
                    elseif ($siteInfo.HubSiteId -and $siteInfo.HubSiteId -ne "00000000-0000-0000-0000-000000000000") {
                        Write-LogMessage -Message "$siteName is associated with different hub, removing first..." -Type Warning
                        Remove-SPOHubSiteAssociation -Site $siteUrl
                        Start-Sleep -Seconds 5
                        Add-SPOHubSiteAssociation -Site $siteUrl -HubSite $hubSiteUrl
                        Write-LogMessage -Message "$siteName re-associated with hub" -Type Success
                    }
                    else {
                        Add-SPOHubSiteAssociation -Site $siteUrl -HubSite $hubSiteUrl
                        Write-LogMessage -Message "$siteName associated with hub" -Type Success
                    }
                }
                catch {
                    if ($_.Exception.Message -like "*already associated*") {
                        Write-LogMessage -Message "$siteName already associated with hub" -Type Warning
                    }
                    else {
                        Write-LogMessage -Message "Failed to associate $siteName with hub - $($_.Exception.Message)" -Type Warning
                    }
                }
                
                # Small delay to prevent throttling
                Start-Sleep -Seconds 2
            }
            
            Write-Progress -Id 5 -Activity "Hub Site Associations" -Completed
        }
        else {
            Write-LogMessage -Message "Skipping hub associations - hub site not registered" -Type Warning
        }
        
        # Configure site permissions (simplified approach)
        Write-LogMessage -Message "Configuring basic site permissions..." -Type Info
        
        # Set current user as site collection admin for all sites
        $currentUserEmail = $context.Account
        
        foreach ($siteUrl in $createdSites) {
            try {
                # Ensure both owner and current user have site collection admin rights
                Set-SPOUser -Site $siteUrl -LoginName $ownerEmail -IsSiteCollectionAdmin $true
                if ($currentUserEmail -ne $ownerEmail) {
                    Set-SPOUser -Site $siteUrl -LoginName $currentUserEmail -IsSiteCollectionAdmin $true
                }
                Write-LogMessage -Message "Set site collection admins for $siteUrl" -Type Success
            }
            catch {
                Write-LogMessage -Message "Failed to set site collection admin for $siteUrl - $($_.Exception.Message)" -Type Warning
            }
        }
        
        # Provide guidance for manual security group configuration
        Write-LogMessage -Message "SharePoint sites created successfully!" -Type Success
        Write-LogMessage -Message "Manual security group configuration may be needed:" -Type Info
        Write-LogMessage -Message "1. SharePoint Admin Center: https://$tenantName-admin.sharepoint.com" -Type Info
        Write-LogMessage -Message "2. For each site, go to Permissions and add the corresponding security groups:" -Type Info
        
        foreach ($site in $spokeSites) {
            $siteName = $site.Name
            Write-LogMessage -Message "   $siteName site: Add '$siteName SharePoint Members/Owners/Visitors' groups" -Type Info
        }
        
        Write-LogMessage -Message "3. Alternative: Use PnP PowerShell for automated group assignment" -Type Info
        Write-LogMessage -Message "   Install-Module PnP.PowerShell -Scope CurrentUser" -Type Info
        Write-LogMessage -Message "   Connect-PnPOnline [SiteURL] -Interactive" -Type Info
        Write-LogMessage -Message "   Add-PnPGroupMember -LoginName '[GroupName]' -Identity 'Site Members'" -Type Info
        
        Write-LogMessage -Message "SharePoint configuration completed successfully" -Type Success
        
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
        
        # Clean up any progress bars on error
        Write-Progress -Id 1 -Activity "Creating Security Groups" -Completed
        Write-Progress -Id 2 -Activity "Group Propagation" -Completed
        Write-Progress -Id 3 -Activity "Creating SharePoint Sites" -Completed
        Write-Progress -Id 4 -Activity "Site Provisioning" -Completed
        Write-Progress -Id 5 -Activity "Hub Site Associations" -Completed
        
        return $false
    }
}