# === SharePoint.ps1 ===
# SharePoint Online configuration and site creation functions - FIXED VERSION

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
            Disconnect-SPOService -ErrorAction SilentlyContinue | Out-Null
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
        
        # STEP 7: Get SharePoint configuration with validation
        $customerName = $script:TenantState.TenantName
        Write-Host "`nSharePoint URL Configuration" -ForegroundColor Yellow
        Write-Host "Example: If your tenant is 'm365x36060197.sharepoint.com', enter 'm365x36060197'" -ForegroundColor Cyan
        
        # Auto-detect tenant name from default domain if possible
        $suggestedTenant = ""
        if ($script:TenantState.DefaultDomain) {
            $domain = $script:TenantState.DefaultDomain
            if ($domain -like "*.onmicrosoft.com") {
                $suggestedTenant = $domain.Replace(".onmicrosoft.com", "")
                Write-Host "Suggested tenant name based on your domain: $suggestedTenant" -ForegroundColor Green
            }
        }
        
        $tenantName = Read-Host "Enter your SharePoint tenant name (without .sharepoint.com)"
        if ([string]::IsNullOrWhiteSpace($tenantName)) {
            if (-not [string]::IsNullOrWhiteSpace($suggestedTenant)) {
                $tenantName = $suggestedTenant
                Write-LogMessage -Message "Using suggested tenant name: $tenantName" -Type Info
            } else {
                Write-LogMessage -Message "Tenant name is required" -Type Error
                return $false
            }
        }
        
        # Construct URLs automatically
        $adminUrl = "https://$tenantName-admin.sharepoint.com"
        $tenantUrl = "https://$tenantName.sharepoint.com"
        $ownerEmail = $script:TenantState.AdminEmail
        
        Write-LogMessage -Message "SharePoint Admin URL: $adminUrl" -Type Info
        Write-LogMessage -Message "SharePoint Tenant URL: $tenantUrl" -Type Info
        Write-LogMessage -Message "Site Owner: $ownerEmail" -Type Info
        
        # CRITICAL: Check SharePoint Administrator permissions BEFORE proceeding
        Write-LogMessage -Message "Verifying SharePoint Administrator permissions..." -Type Info
        Write-Host "`nCRITICAL: This script requires SharePoint Administrator or Global Administrator role." -ForegroundColor Red
        Write-Host "Please confirm you have one of these roles in Microsoft 365 Admin Center:" -ForegroundColor Yellow
        Write-Host "- SharePoint Administrator" -ForegroundColor Cyan
        Write-Host "- Global Administrator" -ForegroundColor Cyan
        
        $permissionConfirm = Read-Host "`nDo you have SharePoint Administrator or Global Administrator role? (Y/N)"
        if ($permissionConfirm -ne 'Y' -and $permissionConfirm -ne 'y') {
            Write-LogMessage -Message "SharePoint Administrator role is required. Please assign the role first." -Type Error
            Write-LogMessage -Message "Go to https://admin.microsoft.com > Users > Active Users > Select User > Manage Roles" -Type Info
            return $false
        }
        
        # Connect to SharePoint Online Admin Center with modern authentication
        Write-LogMessage -Message "Connecting to SharePoint Online Admin Center..." -Type Info
        try {
            Connect-SPOService -Url $adminUrl -ModernAuth $true
            Write-LogMessage -Message "Successfully connected to SharePoint Online" -Type Success
            
            # Test connection by trying to get tenant information
            $tenantInfo = Get-SPOTenant -ErrorAction Stop
            Write-LogMessage -Message "Connection verified - SharePoint Administrator permissions confirmed" -Type Success
        }
        catch {
            Write-LogMessage -Message "Failed to connect or verify permissions: $($_.Exception.Message)" -Type Error
            Write-LogMessage -Message "This usually means insufficient permissions or incorrect tenant name" -Type Error
            return $false
        }
        
        # Create security groups FIRST (they need time to propagate)
        Write-LogMessage -Message "Creating security groups (these need time to propagate)..." -Type Info
        $spokeSites = @()
        foreach ($siteName in $SharePointConfig.DefaultSites) {
            $spokeSites += @{ 
                Name = $siteName
                URL = "$tenantUrl/sites/$($siteName.ToLower())" 
            }
        }
        
        $securityGroups = @{}
        $totalGroups = $spokeSites.Count * 3
        $currentGroup = 0
        
        foreach ($site in $spokeSites) {
            $siteName = $site.Name
            Write-LogMessage -Message "Creating security groups for site: $siteName" -Type Info
            
            foreach ($groupType in @("Members", "Owners", "Visitors")) {
                $currentGroup++
                $groupName = "$siteName SharePoint $groupType"
                
                $existingGroup = Get-MgGroup -Filter "DisplayName eq '$groupName'" -ErrorAction SilentlyContinue
                
                if (-not $existingGroup) {
                    try {
                        $mailNick = "$($siteName.ToLower())spo$($groupType.ToLower())"
                        $newGroup = New-MgGroup -DisplayName $groupName -MailEnabled:$false -MailNickname $mailNick -SecurityEnabled:$true -Description "Security group for $siteName SharePoint $groupType access"
                        $securityGroups["$siteName-$groupType"] = $newGroup.Id
                        Write-LogMessage -Message "Created security group: $groupName" -Type Success
                        
                        # Add a small delay to prevent throttling
                        Start-Sleep -Seconds 1
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
        
        # EXTENDED propagation wait for security groups (critical for SharePoint)
        Write-LogMessage -Message "Waiting for security groups to propagate in Azure AD (5 minutes)..." -Type Info
        Write-LogMessage -Message "This extended wait is necessary for SharePoint to recognize the groups." -Type Info
        
        $propagationTime = 300 # 5 minutes for SharePoint
        for ($i = 1; $i -le $propagationTime; $i++) {
            $remaining = $propagationTime - $i
            $minutes = [math]::Floor($remaining / 60)
            $seconds = $remaining % 60
            Write-Progress -Activity "Security Group Propagation" -Status "Waiting for Azure AD sync (required for SharePoint)..." -PercentComplete (($i / $propagationTime) * 100) -CurrentOperation "Time remaining: $minutes min $seconds sec"
            Start-Sleep -Seconds 1
        }
        Write-Progress -Activity "Security Group Propagation" -Completed
        
        # Create Hub Site
        $hubSiteTitle = "$customerName Hub"
        $hubSiteUrl = "$tenantUrl/sites/corporatehub"
        
        Write-LogMessage -Message "Creating or verifying hub site: $hubSiteTitle" -Type Info
        
        # Check if hub site exists and create if needed
        try {
            $existingHubSite = Get-SPOSite -Identity $hubSiteUrl -ErrorAction SilentlyContinue
            if ($existingHubSite) {
                Write-LogMessage -Message "Hub site already exists: $hubSiteUrl" -Type Warning
            } else {
                Write-LogMessage -Message "Creating new hub site..." -Type Info
                New-SPOSite -Url $hubSiteUrl -Owner $ownerEmail -StorageQuota $SharePointConfig.StorageQuota -Title $hubSiteTitle -Template $SharePointConfig.SiteTemplate -Wait
                Write-LogMessage -Message "Hub site created: $hubSiteUrl" -Type Success
                
                # Wait for site to be fully provisioned
                Write-LogMessage -Message "Waiting for hub site to be fully provisioned..." -Type Info
                Start-Sleep -Seconds 60
            }
        }
        catch {
            Write-LogMessage -Message "Error with hub site creation: $($_.Exception.Message)" -Type Error
            if ($_.Exception.Message -like "*already exists*") {
                Write-LogMessage -Message "Hub site already exists (creation failed but site exists)" -Type Warning
            } else {
                Write-LogMessage -Message "Hub site creation failed - continuing without hub functionality" -Type Warning
                $hubSiteUrl = $null
            }
        }
        
        # Register as Hub Site with proper verification and retry logic
        if ($hubSiteUrl) {
            Write-LogMessage -Message "Registering site as Hub Site..." -Type Info
            $hubRegistered = $false
            $maxAttempts = 3
            
            for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
                try {
                    # Check if already registered
                    $existingHubSites = Get-SPOHubSite -ErrorAction SilentlyContinue
                    $isAlreadyHub = $existingHubSites | Where-Object { $_.SiteUrl -eq $hubSiteUrl }
                    
                    if ($isAlreadyHub) {
                        Write-LogMessage -Message "Site is already registered as Hub Site" -Type Success
                        $hubRegistered = $true
                        break
                    }
                    
                    # Register as hub site
                    Write-LogMessage -Message "Attempt $attempt to register hub site..." -Type Info
                    Register-SPOHubSite -Site $hubSiteUrl -Principals $ownerEmail
                    
                    # Verify registration
                    Start-Sleep -Seconds 30
                    $verifyHub = Get-SPOHubSite | Where-Object { $_.SiteUrl -eq $hubSiteUrl }
                    if ($verifyHub) {
                        Write-LogMessage -Message "Hub site registration successful!" -Type Success
                        $hubRegistered = $true
                        break
                    } else {
                        Write-LogMessage -Message "Hub registration verification failed on attempt $attempt" -Type Warning
                        if ($attempt -lt $maxAttempts) {
                            Write-LogMessage -Message "Waiting before retry..." -Type Info
                            Start-Sleep -Seconds 30
                        }
                    }
                }
                catch {
                    Write-LogMessage -Message "Hub registration attempt $attempt failed: $($_.Exception.Message)" -Type Warning
                    if ($attempt -eq $maxAttempts) {
                        Write-LogMessage -Message "All hub registration attempts failed" -Type Error
                    } else {
                        Start-Sleep -Seconds 30
                    }
                }
            }
            
            if (-not $hubRegistered) {
                Write-LogMessage -Message "Hub site registration failed - sites will be created without hub association" -Type Warning
                $hubSiteUrl = $null
            }
        }
        
        # Create spoke sites
        Write-LogMessage -Message "Creating spoke sites..." -Type Info
        $createdSites = @()
        
        foreach ($site in $spokeSites) {
            $siteUrl = $site.URL
            $siteName = $site.Name
            
            Write-LogMessage -Message "Processing site: $siteName" -Type Info
            
            try {
                $existingSite = Get-SPOSite -Identity $siteUrl -ErrorAction SilentlyContinue
                if ($existingSite) {
                    Write-LogMessage -Message "$siteName site already exists: $siteUrl" -Type Warning
                    $createdSites += $siteUrl
                } else {
                    Write-LogMessage -Message "Creating $siteName site..." -Type Info
                    New-SPOSite -Url $siteUrl -Owner $ownerEmail -StorageQuota $SharePointConfig.StorageQuota -Title $siteName -Template $SharePointConfig.SiteTemplate -Wait
                    Write-LogMessage -Message "$siteName site created: $siteUrl" -Type Success
                    $createdSites += $siteUrl
                    
                    # Small delay between site creations
                    Start-Sleep -Seconds 10
                }
            }
            catch {
                if ($_.Exception.Message -like "*already exists*") {
                    Write-LogMessage -Message "$siteName site already exists" -Type Warning
                    $createdSites += $siteUrl
                } else {
                    Write-LogMessage -Message "Failed to create $siteName site: $($_.Exception.Message)" -Type Error
                    continue
                }
            }
            
            # Associate with hub site if hub is available
            if ($hubRegistered -and $hubSiteUrl) {
                try {
                    Add-SPOHubSiteAssociation -Site $siteUrl -HubSite $hubSiteUrl
                    Write-LogMessage -Message "$siteName site associated with Hub site" -Type Success
                }
                catch {
                    if ($_.Exception.Message -like "*already associated*") {
                        Write-LogMessage -Message "$siteName site already associated with hub" -Type Warning
                    } else {
                        Write-LogMessage -Message "Failed to associate $siteName with hub: $($_.Exception.Message)" -Type Warning
                    }
                }
            }
        }
        
        # Configure site permissions with proper error handling
        Write-LogMessage -Message "Configuring site permissions..." -Type Info
        Write-LogMessage -Message "Note: Permission configuration may show warnings - this is normal in some tenants" -Type Info
        
        $successfulPermissions = 0
        $totalPermissionOperations = $createdSites.Count * 3
        
        foreach ($siteUrl in $createdSites) {
            $siteName = ($spokeSites | Where-Object { $_.URL -eq $siteUrl }).Name
            
            Write-LogMessage -Message "Configuring permissions for: $siteName" -Type Info
            
            # Set site collection admin first
            try {
                Set-SPOUser -Site $siteUrl -LoginName $ownerEmail -IsSiteCollectionAdmin $true
                Write-LogMessage -Message "Set site collection admin for $siteName" -Type Success
                Start-Sleep -Seconds 5
            }
            catch {
                Write-LogMessage -Message "Failed to set site collection admin for $siteName : $($_.Exception.Message)" -Type Warning
            }
            
            # Try to add security groups with multiple strategies
            foreach ($groupType in @("Members", "Owners", "Visitors")) {
                $groupKey = "$siteName-$groupType"
                
                if ($securityGroups.ContainsKey($groupKey)) {
                    $groupId = $securityGroups[$groupKey]
                    $success = $false
                    
                    # Strategy 1: Modern approach with group ID
                    try {
                        $claimFormat = "c:0t.c|tenant|$groupId"
                        $spoGroupName = switch ($groupType) {
                            "Owners" { "$siteName Owners" }
                            "Members" { "$siteName Members" }
                            "Visitors" { "$siteName Visitors" }
                        }
                        
                        Add-SPOUser -Site $siteUrl -LoginName $claimFormat -Group $spoGroupName
                        Write-LogMessage -Message "Added $groupType security group to $siteName" -Type Success
                        $successfulPermissions++
                        $success = $true
                    }
                    catch {
                        # Strategy 2: Try with default SharePoint groups
                        try {
                            $defaultGroup = switch ($groupType) {
                                "Owners" { "$siteName Owners" }
                                "Members" { "$siteName Members" }  
                                "Visitors" { "$siteName Visitors" }
                            }
                            
                            Add-SPOUser -Site $siteUrl -LoginName $claimFormat -Group $defaultGroup
                            Write-LogMessage -Message "Added $groupType security group to $siteName (alternative method)" -Type Success
                            $successfulPermissions++
                            $success = $true
                        }
                        catch {
                            Write-LogMessage -Message "Could not add $groupType group to $siteName - manual configuration needed" -Type Warning
                        }
                    }
                } else {
                    Write-LogMessage -Message "No security group found for $siteName $groupType" -Type Warning
                }
            }
        }
        
        # Provide summary and manual configuration guidance
        Write-LogMessage -Message "SharePoint configuration summary:" -Type Info
        Write-LogMessage -Message "- Hub site: $($hubRegistered ? 'Successfully created' : 'Creation failed')" -Type Info
        Write-LogMessage -Message "- Spoke sites: $($createdSites.Count) sites processed" -Type Info
        Write-LogMessage -Message "- Security group assignments: $successfulPermissions of $totalPermissionOperations successful" -Type Info
        
        if ($successfulPermissions -lt $totalPermissionOperations) {
            Write-LogMessage -Message "Some security group assignments failed. Manual configuration options:" -Type Warning
            Write-LogMessage -Message "Option 1 - SharePoint Admin Center:" -Type Info
            Write-LogMessage -Message "  1. Go to https://$tenantName-admin.sharepoint.com" -Type Info
            Write-LogMessage -Message "  2. Sites → Active Sites → Select site → Permissions" -Type Info
            Write-LogMessage -Message "  3. Add the security groups manually" -Type Info
            Write-LogMessage -Message "" -Type Info
            Write-LogMessage -Message "Option 2 - PnP PowerShell (recommended for bulk operations):" -Type Info
            Write-LogMessage -Message "  Install-Module PnP.PowerShell -Scope CurrentUser" -Type Info
            Write-LogMessage -Message "  Connect-PnPOnline https://$tenantName.sharepoint.com/sites/[sitename] -Interactive" -Type Info
            Write-LogMessage -Message "  Add-PnPGroupMember -LoginName '[GroupName]' -Identity '[Site Group]'" -Type Info
            Write-LogMessage -Message "" -Type Info
            Write-LogMessage -Message "Security groups created:" -Type Info
            foreach ($site in $spokeSites) {
                Write-LogMessage -Message "  - $($site.Name) SharePoint Members/Owners/Visitors" -Type Info
            }
        }
        
        Write-LogMessage -Message "SharePoint configuration completed!" -Type Success
        
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
        Write-LogMessage -Message "Critical error in SharePoint configuration: $($_.Exception.Message)" -Type Error
        Write-LogMessage -Message "Full error details: $($_.Exception.ToString())" -Type Error -LogOnly
        
        # Clean up any progress bars on error
        Write-Progress -Activity "Security Group Propagation" -Completed
        
        # Clean disconnect on error
        try {
            Disconnect-SPOService -ErrorAction SilentlyContinue
        }
        catch {
            # Ignore disconnect errors
        }
        
        return $false
    }
}