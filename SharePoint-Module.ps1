# === SharePoint.ps1 ===
# SharePoint Online configuration and site creation functions - Converted to Direct Execution

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
        
        # STEP 7: Original SharePoint logic continues with fixes
        # Get SharePoint URLs - simplified input
        $customerName = $script:TenantState.TenantName
        Write-Host "SharePoint URL Configuration" -ForegroundColor Yellow
        Write-Host "Example: If your tenant is 'm365x36060197.sharepoint.com', enter 'm365x36060197'" -ForegroundColor Cyan
        $tenantName = Read-Host "Enter your SharePoint tenant name (without .sharepoint.com)"
        
        # Construct URLs automatically
        $adminUrl = "https://$tenantName-admin.sharepoint.com"
        $tenantUrl = "https://$tenantName.sharepoint.com"
        $ownerEmail = $script:TenantState.AdminEmail
        
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
        try {
            $existingHubSite = Get-SPOSite | Where-Object { $_.Url -eq $hubSiteUrl }
            if ($existingHubSite) {
                Write-LogMessage -Message "Hub site already exists: $hubSiteUrl" -Type Warning
            } else {
                New-SPOSite -Url $hubSiteUrl -Owner $ownerEmail -StorageQuota $SharePointConfig.StorageQuota -Title $hubSiteTitle -Template $SharePointConfig.SiteTemplate
                Write-LogMessage -Message "Hub site created: $hubSiteUrl" -Type Success
            }
        }
        catch {
            Write-LogMessage -Message "Hub site may already exist or creation failed - $($_.Exception.Message)" -Type Warning
        }
        
        # Set the site as a Hub Site with proper verification
        Write-LogMessage -Message "Registering Hub site..." -Type Info
        try {
            # First check if it's already a hub site
            $existingHubSites = Get-SPOHubSite -ErrorAction SilentlyContinue
            $isAlreadyHub = $existingHubSites | Where-Object { $_.SiteUrl -eq $hubSiteUrl }
            
            if ($isAlreadyHub) {
                Write-LogMessage -Message "Site is already registered as a Hub site: $hubSiteUrl" -Type Warning
            }
            else {
                # Wait for site to be fully provisioned before registering as hub
                Write-LogMessage -Message "Waiting for hub site to be fully provisioned before registration..." -Type Info
                Start-Sleep -Seconds 30
                
                # Register as hub site
                $principals = @($ownerEmail)
                Register-SPOHubSite -Site $hubSiteUrl -Principals $principals
                Write-LogMessage -Message "Hub site registered successfully: $hubSiteUrl" -Type Success
                
                # Verify registration was successful
                Start-Sleep -Seconds 10
                $verifyHub = Get-SPOHubSite | Where-Object { $_.SiteUrl -eq $hubSiteUrl }
                if ($verifyHub) {
                    Write-LogMessage -Message "Hub site registration verified" -Type Success
                }
                else {
                    Write-LogMessage -Message "Hub site registration may not have completed - will retry association later" -Type Warning
                }
            }
        }
        catch {
            Write-LogMessage -Message "Hub site registration failed - $($_.Exception.Message)" -Type Error
            Write-LogMessage -Message "Continuing without hub site functionality..." -Type Warning
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
        
        # Wait for security groups to propagate with proper progress bar
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
        
        Write-Progress -Id 3 -Activity "Creating SharePoint Sites" -Status "Starting site creation..." -PercentComplete 0
        
        foreach ($site in $spokeSites) {
            $currentSite++
            $percentComplete = [math]::Round(($currentSite / $totalSites) * 100, 2)
            
            $siteUrl = $site.URL
            $siteName = $site.Name
            
            Write-Progress -Id 3 -Activity "Creating SharePoint Sites" -Status "Creating: $siteName" -PercentComplete $percentComplete -CurrentOperation "Site $currentSite of $totalSites"
            
            # Check if the Spoke Site already exists, if not, create one
            try {
                $existingSpokeSite = Get-SPOSite | Where-Object { $_.Url -eq $siteUrl }
                if ($existingSpokeSite) {
                    Write-LogMessage -Message "$siteName site already exists: $siteUrl" -Type Warning
                } else {
                    New-SPOSite -Url $siteUrl -Owner $ownerEmail -StorageQuota $SharePointConfig.StorageQuota -Title "$siteName" -Template $SharePointConfig.SiteTemplate
                    Write-LogMessage -Message "$siteName site created: $siteUrl" -Type Success
                }
            }
            catch {
                # If checking fails, try to create anyway - might be a permissions issue with Get-SPOSite
                try {
                    New-SPOSite -Url $siteUrl -Owner $ownerEmail -StorageQuota $SharePointConfig.StorageQuota -Title "$siteName" -Template $SharePointConfig.SiteTemplate
                    Write-LogMessage -Message "$siteName site created: $siteUrl" -Type Success
                }
                catch {
                    if ($_.Exception.Message -like "*already exists*" -or $_.Exception.Message -like "*site collection*already*") {
                        Write-LogMessage -Message "$siteName site already exists: $siteUrl" -Type Warning
                    } else {
                        Write-LogMessage -Message "Failed to create $siteName site - $($_.Exception.Message)" -Type Error
                        continue
                    }
                }
            }
        
            # Register Spoke Site to the Hub with proper verification
            try {
                # Verify hub site exists before association
                $hubSiteExists = Get-SPOHubSite | Where-Object { $_.SiteUrl -eq $hubSiteUrl }
                if ($hubSiteExists) {
                    Add-SPOHubSiteAssociation -Site $siteUrl -HubSite $hubSiteUrl
                    Write-LogMessage -Message "$siteName site associated with Hub site" -Type Success
                }
                else {
                    Write-LogMessage -Message "Hub site not found or not registered properly - skipping association for $siteName" -Type Warning
                }
            }
            catch {
                if ($_.Exception.Message -like "*already associated*") {
                    Write-LogMessage -Message "$siteName site already associated with hub" -Type Warning
                }
                else {
                    Write-LogMessage -Message "Failed to associate $siteName site with hub - $($_.Exception.Message)" -Type Warning
                }
            }
        }
        
        # Complete site creation progress
        Write-Progress -Id 3 -Activity "Creating SharePoint Sites" -Completed
        
        # Wait for sites to provision with proper progress bar
        Write-LogMessage -Message "Waiting for SharePoint sites to provision (3 minutes)..." -Type Info
        $provisionTime = 180
        for ($i = 1; $i -le $provisionTime; $i++) {
            $percentComplete = [math]::Round(($i / $provisionTime) * 100, 2)
            $remaining = $provisionTime - $i
            Write-Progress -Id 4 -Activity "Site Provisioning" -Status "Waiting for SharePoint sites to fully provision..." -PercentComplete $percentComplete -SecondsRemaining $remaining
            Start-Sleep -Seconds 1
        }
        Write-Progress -Id 4 -Activity "Site Provisioning" -Completed
        
        # Ensure current user has site collection admin rights before adding groups
        Write-LogMessage -Message "Verifying SharePoint Administrator permissions..." -Type Info
        $currentUserEmail = $context.Account
        
        # Verify SharePoint admin permissions first
        try {
            $adminCheck = Get-SPOTenant -ErrorAction Stop
            Write-LogMessage -Message "SharePoint Administrator permissions verified" -Type Success
        }
        catch {
            Write-LogMessage -Message "WARNING: May not have SharePoint Administrator permissions. Security group assignment may fail." -Type Warning
            Write-LogMessage -Message "Please ensure your account has SharePoint Administrator role in Microsoft 365 admin center" -Type Warning
        }
        
        # Add security groups to sites with enhanced permissions handling
        $totalOperations = $spokeSites.Count * 3
        $currentOperation = 0
        
        Write-Progress -Id 5 -Activity "Configuring Site Permissions" -Status "Starting permission configuration..." -PercentComplete 0
        
        foreach ($site in $spokeSites) {
            $siteUrl = $site.URL
            $siteName = $site.Name
            
            # Set site collection administrator permissions for both accounts
            Write-LogMessage -Message "Setting site collection admin permissions for $siteUrl..." -Type Info
            try {
                # Add the owner email as site collection admin
                Set-SPOUser -Site $siteUrl -LoginName $ownerEmail -IsSiteCollectionAdmin $true
                Write-LogMessage -Message "Set $ownerEmail as site collection admin for $siteUrl" -Type Success
                
                # Also add current user if different
                if ($currentUserEmail -ne $ownerEmail) {
                    Set-SPOUser -Site $siteUrl -LoginName $currentUserEmail -IsSiteCollectionAdmin $true
                    Write-LogMessage -Message "Set $currentUserEmail as site collection admin for $siteUrl" -Type Success
                }
                
                # Wait for permission changes to propagate
                Start-Sleep -Seconds 5
            }
            catch {
                Write-LogMessage -Message "Failed to set site collection admin for $siteUrl - $($_.Exception.Message)" -Type Error
                Write-LogMessage -Message "This may cause security group assignment to fail for this site" -Type Warning
            }
            
            foreach ($groupType in @("Members", "Owners", "Visitors")) {
                $currentOperation++
                $percentComplete = [math]::Round(($currentOperation / $totalOperations) * 100, 2)
                
                $groupKey = "$siteName-$groupType"
                $spoGroupName = "$siteName $groupType"
                
                Write-Progress -Id 5 -Activity "Configuring Site Permissions" -Status "Adding $groupType to $siteName" -PercentComplete $percentComplete -CurrentOperation "Operation $currentOperation of $totalOperations"
                
                if ($securityGroups.ContainsKey($groupKey)) {
                    $groupId = $securityGroups[$groupKey]
                    # Use proper Azure AD security group claim format
                    $claimFormat = "c:0t.c|tenant|$groupId"
                    
                    # Multiple attempt strategies
                    $success = $false
                    
                    # Strategy 1: Standard Add-SPOUser approach
                    try {
                        Add-SPOUser -Site $siteUrl -Group $spoGroupName -LoginName $claimFormat
                        Write-LogMessage -Message "Added security group '$groupType' to $siteName using standard method" -Type Success
                        $success = $true
                    }
                    catch {
                        if ($_.Exception.Message -like "*already exists*" -or $_.Exception.Message -like "*already a member*") {
                            Write-LogMessage -Message "Security group already exists as $groupType in $siteName" -Type Warning
                            $success = $true
                        }
                        else {
                            Write-LogMessage -Message "Standard method failed for $siteName $groupType - $($_.Exception.Message)" -Type Warning
                        }
                    }
                    
                    # Strategy 2: Alternative with explicit site collection context
                    if (-not $success) {
                        try {
                            # Wait and retry with different approach
                            Start-Sleep -Seconds 3
                            
                            # Try with explicit permission level
                            $permissionLevel = switch ($groupType) {
                                "Owners" { "Full Control" }
                                "Members" { "Edit" }
                                "Visitors" { "Read" }
                                default { "Read" }
                            }
                            
                            Add-SPOUser -Site $siteUrl -LoginName $claimFormat -Group $spoGroupName
                            Write-LogMessage -Message "Added security group '$groupType' to $siteName using alternative method" -Type Success
                            $success = $true
                        }
                        catch {
                            Write-LogMessage -Message "Alternative method also failed for $siteName $groupType - $($_.Exception.Message)" -Type Warning
                        }
                    }
                    
                    # Strategy 3: Try with specific permission levels directly
                    if (-not $success) {
                        try {
                            Start-Sleep -Seconds 2
                            $permissionLevel = switch ($groupType) {
                                "Owners" { "Full Control" }
                                "Members" { "Edit" }
                                "Visitors" { "Read" }
                                default { "Read" }
                            }
                            
                            # Alternative: Try adding to site directly with permission level
                            $siteGroups = Get-SPOSiteGroup -Site $siteUrl
                            $targetGroup = $siteGroups | Where-Object { $_.Title -eq $spoGroupName }
                            
                            if ($targetGroup) {
                                Add-SPOUser -Site $siteUrl -LoginName $claimFormat -Group $targetGroup.Title
                                Write-LogMessage -Message "Added security group to $siteName $groupType using direct group reference" -Type Success
                                $success = $true
                            }
                            else {
                                Write-LogMessage -Message "Could not find SharePoint group '$spoGroupName' in site $siteUrl" -Type Warning
                            }
                        }
                        catch {
                            Write-LogMessage -Message "Direct group reference method also failed for $siteName $groupType - $($_.Exception.Message)" -Type Warning
                        }
                    }
                    
                    if (-not $success) {
                        Write-LogMessage -Message "All methods failed to add security group to $siteName as $groupType" -Type Error
                        Write-LogMessage -Message "Manual configuration may be required for this site" -Type Warning
                        Write-LogMessage -Message "Group ID: $groupId, Claim Format: $claimFormat" -Type Info
                    }
                } else {
                    Write-LogMessage -Message "No security group found for $siteName $groupType. Skipping..." -Type Warning
                }
            }
            
            # Add small delay between sites to prevent throttling
            Start-Sleep -Seconds 2
        }
        
        # Complete permissions configuration progress
        Write-Progress -Id 5 -Activity "Configuring Site Permissions" -Completed
        
        # Provide manual configuration guidance
        Write-LogMessage -Message "If any security group assignments failed, you can manually configure them:" -Type Info
        Write-LogMessage -Message "1. SharePoint Admin Center: https://$tenantName-admin.sharepoint.com" -Type Info
        Write-LogMessage -Message "   → Active Sites → Select site → Permissions → Add security groups" -Type Info
        Write-LogMessage -Message "2. Alternative with PnP PowerShell (more reliable):" -Type Info
        Write-LogMessage -Message "   Install-Module PnP.PowerShell" -Type Info
        Write-LogMessage -Message "   Connect-PnPOnline https://$tenantName.sharepoint.com/sites/sitename -Interactive" -Type Info
        Write-LogMessage -Message "   Add-PnPGroupMember -LoginName 'GroupName' -Identity 'Site Members'" -Type Info
        Write-LogMessage -Message "3. Security group names created: HR/Processes/Templates/Documents/Policies SharePoint Members/Owners/Visitors" -Type Info
        
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
        Write-Progress -Id 5 -Activity "Configuring Site Permissions" -Completed
        
        return $false
    }
}