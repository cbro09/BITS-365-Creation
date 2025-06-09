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
        
        # STEP 7: Original script logic continues unchanged
        # Clear SharePoint authentication cache to prevent conflicts
        Write-LogMessage -Message "Clearing SharePoint authentication cache..." -Type Info
        try {
            Disconnect-MsgGraph -ErrorAction SilentlyContinue | Out-Null
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
        
        # Connect to SharePoint Online Admin Center
        Write-LogMessage -Message "Connecting to SharePoint Online Admin Center..." -Type Info
        Connect-SPOService -Url $adminUrl
        
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
        
        # Create security groups for each site
        $securityGroups = @{}
        
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
        
        # Wait for security groups to propagate
        Write-LogMessage -Message "Waiting for security groups to propagate (2 minutes)..." -Type Info
        for ($i = 0; $i -lt 120; $i++) {
            Start-Sleep -Seconds 1
            Show-Progress -Current ($i + 1) -Total 120 -Status "Waiting for groups to propagate..."
        }
        Write-Host ""
        
        # Create spoke sites
        foreach ($site in $spokeSites) {
            $siteUrl = $site.URL
            $siteName = $site.Name
            
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
        
        # Wait for sites to provision
        Write-LogMessage -Message "Waiting for SharePoint sites to provision (3 minutes)..." -Type Info
        for ($i = 0; $i -lt 180; $i++) {
            Start-Sleep -Seconds 1
            Show-Progress -Current ($i + 1) -Total 180 -Status "Waiting for sites to provision..."
        }
        Write-Host ""
        
        # Add security groups to sites
        foreach ($site in $spokeSites) {
            $siteUrl = $site.URL
            $siteName = $site.Name
            
            foreach ($groupType in @("Members", "Owners", "Visitors")) {
                $groupKey = "$siteName-$groupType"
                $spoGroupName = "$siteName $groupType"
                
                if ($securityGroups.ContainsKey($groupKey)) {
                    $groupId = $securityGroups[$groupKey]
                    try {
                        Add-SPOUser -Site $siteUrl -Group $spoGroupName -LoginName "c:0t.c|tenant|$groupId" -ErrorAction Stop
                        Write-LogMessage -Message "Added security group as $groupType to $siteUrl" -Type Success
                    } catch {
                        # Check if it's already added
                        if ($_.Exception.Message -like "*already exists*" -or $_.Exception.Message -like "*already a member*") {
                            Write-LogMessage -Message "Security group already exists as $groupType in $siteUrl" -Type Warning
                        } else {
                            Write-LogMessage -Message "Failed to add security group to $siteUrl - $($_.Exception.Message)" -Type Error
                        }
                    }
                } else {
                    Write-LogMessage -Message "No security group found for $siteName $groupType. Skipping..." -Type Warning
                }
            }
        }
        
        Write-LogMessage -Message "SharePoint configuration completed successfully" -Type Success
        Disconnect-SPOService
        return $true
    }
    catch {
        Write-LogMessage -Message "Error in SharePoint configuration - $($_.Exception.Message)" -Type Error
        return $false
    }
}