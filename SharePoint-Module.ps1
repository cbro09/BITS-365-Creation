# === SharePoint.ps1 ===
# SharePoint Online configuration and site creation functions

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
        
        # STEP 2: Remove ALL Graph modules first to avoid conflicts
        Write-LogMessage -Message "Clearing all Graph modules to prevent conflicts..." -Type Info
        Get-Module Microsoft.Graph* | Remove-Module -Force -ErrorAction SilentlyContinue
        
        # STEP 3: Restore core functions
        ${function:Write-LogMessage} = $writeLogFunction
        ${function:Test-NotEmpty} = $testNotEmptyFunction
        
        # STEP 4: Disconnect any existing sessions
        try {
            Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        }
        catch {
            # Ignore disconnect errors
        }
        
        # STEP 5: Force load ONLY the exact modules needed for SharePoint
        $sharePointModules = @(
            'Microsoft.Graph.Groups',
            'Microsoft.Graph.Identity.DirectoryManagement'
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
        
        # STEP 7: Original script logic continues unchanged but with fixes
        # Clear SharePoint authentication cache to prevent conflicts
        Write-LogMessage -Message "Clearing SharePoint authentication cache..." -Type Info
        try {
            Disconnect-SPOService -ErrorAction SilentlyContinue | Out-Null
            Remove-Item "$env:USERPROFILE\.mg" -Recurse -Force -ErrorAction SilentlyContinue
        }
        catch {
            # Ignore cleanup errors
        }
        
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
            Write-LogMessage -Message "Hub site may already exist or creation failed: $($_.Exception.Message)" -Type Warning
        }
        
        # Set the site as a Hub Site
        try {
            $principals = @($ownerEmail)
            Register-SPOHubSite -Site $hubSiteUrl -Principals $principals -ErrorAction SilentlyContinue
            Write-LogMessage -Message "Hub site registered: $hubSiteUrl" -Type Success
        }
        catch {
            Write-LogMessage -Message "Hub site registration failed (may already be registered): $($_.Exception.Message)" -Type Warning
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
                        Write-LogMessage -Message "Failed to create $siteName site: $($_.Exception.Message)" -Type Error
                        continue
                    }
                }
            }
        
            # Register Spoke Site to the Hub
            try {
                Add-SPOHubSiteAssociation -Site $siteUrl -HubSite $hubSiteUrl -ErrorAction SilentlyContinue
                Write-LogMessage -Message "$siteName site associated with Hub site" -Type Success
            }
            catch {
                Write-LogMessage -Message "Failed to associate $siteName site with hub (may already be associated): $($_.Exception.Message)" -Type Warning
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
        Write-LogMessage -Message "Verifying and setting site collection administrator permissions..." -Type Info
        $currentUserEmail = $context.Account
        
        # Add security groups to sites with improved permissions handling
        $totalOperations = $spokeSites.Count * 3
        $currentOperation = 0
        
        Write-Progress -Id 5 -Activity "Configuring Site Permissions" -Status "Starting permission configuration..." -PercentComplete 0
        
        foreach ($site in $spokeSites) {
            $siteUrl = $site.URL
            $siteName = $site.Name
            
            # Ensure current user is site collection admin for this site
            try {
                Set-SPOUser -Site $siteUrl -LoginName $currentUserEmail -IsSiteCollectionAdmin $true -ErrorAction SilentlyContinue
                Write-LogMessage -Message "Set site collection admin permissions for $siteUrl" -Type Info -LogOnly
            }
            catch {
                Write-LogMessage -Message "Could not set site collection admin for $siteUrl - permissions may be limited" -Type Warning
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
                    
                    # Retry logic for group addition with exponential backoff
                    $maxRetries = 3
                    $retryCount = 0
                    $success = $false
                    
                    while ($retryCount -lt $maxRetries -and -not $success) {
                        try {
                            Add-SPOUser -Site $siteUrl -Group $spoGroupName -LoginName $claimFormat -ErrorAction Stop
                            Write-LogMessage -Message "Added security group as $groupType to $siteUrl" -Type Success
                            $success = $true
                        } 
                        catch {
                            $retryCount++
                            if ($_.Exception.Message -like "*already exists*" -or $_.Exception.Message -like "*already a member*") {
                                Write-LogMessage -Message "Security group already exists as $groupType in $siteUrl" -Type Warning
                                $success = $true
                            } 
                            elseif ($_.Exception.Message -like "*unauthorized*" -or $_.Exception.Message -like "*access denied*") {
                                Write-LogMessage -Message "Access denied adding group to $siteUrl - trying to set admin permissions..." -Type Warning
                                try {
                                    Set-SPOUser -Site $siteUrl -LoginName $ownerEmail -IsSiteCollectionAdmin $true
                                    Start-Sleep -Seconds (2 * $retryCount) # Exponential backoff
                                }
                                catch {
                                    Write-LogMessage -Message "Could not set admin permissions for $siteUrl" -Type Error
                                    break
                                }
                            }
                            else {
                                Write-LogMessage -Message "Attempt $retryCount failed to add security group to $siteUrl - $($_.Exception.Message)" -Type Warning
                                if ($retryCount -lt $maxRetries) {
                                    Start-Sleep -Seconds (2 * $retryCount) # Exponential backoff
                                }
                            }
                        }
                    }
                    
                    if (-not $success) {
                        Write-LogMessage -Message "Failed to add security group after $maxRetries attempts to $siteUrl" -Type Error
                    }
                } else {
                    Write-LogMessage -Message "No security group found for $siteName $groupType. Skipping..." -Type Warning
                }
            }
        }
        
        # Complete permissions configuration progress
        Write-Progress -Id 5 -Activity "Configuring Site Permissions" -Completed
        
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