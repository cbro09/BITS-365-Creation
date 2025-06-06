# === Groups.ps1 ===
# Group creation and management functions - Direct Execution Method

# Default group configuration
$DefaultGroups = @{
    Security = @("BITS Admin", "SSPR Enabled", "NoMFA Exemption")
    License = @("BusinessBasic", "BusinessStandard", "BusinessPremium", "ExchangeOnline1", "ExchangeOnline2")
}

function New-TenantGroups {
    Write-LogMessage -Message "Starting group creation process..." -Type Info
    
    try {
        # Store core functions to prevent them being cleared
        $writeLogFunction = ${function:Write-LogMessage}
        $testNotEmptyFunction = ${function:Test-NotEmpty}
        $showProgressFunction = ${function:Show-Progress}
        
        # Remove ALL Graph modules first to avoid conflicts
        Write-LogMessage -Message "Clearing all Graph modules to prevent conflicts..." -Type Info
        Get-Module Microsoft.Graph* | Remove-Module -Force -ErrorAction SilentlyContinue
        
        # Restore core functions
        ${function:Write-LogMessage} = $writeLogFunction
        ${function:Test-NotEmpty} = $testNotEmptyFunction
        ${function:Show-Progress} = $showProgressFunction
        
        # Disconnect any existing sessions
        try {
            Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        }
        catch {
            # Ignore disconnect errors
        }
        
        # Force load ONLY the exact modules needed for Groups
        $groupModules = @(
            'Microsoft.Graph.Identity.DirectoryManagement',
            'Microsoft.Graph.Groups'
        )
        
        Write-LogMessage -Message "Loading ONLY Groups modules in exact order..." -Type Info
        foreach ($module in $groupModules) {
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
        
        # Connect with EXACT scopes needed for Groups
        $groupScopes = @(
            "Group.ReadWrite.All",
            "Directory.ReadWrite.All"
        )
        
        Write-LogMessage -Message "Connecting to Microsoft Graph with Groups scopes..." -Type Info
        Connect-MgGraph -Scopes $groupScopes -ErrorAction Stop | Out-Null
        
        $context = Get-MgContext
        Write-LogMessage -Message "Connected to Microsoft Graph as $($context.Account)" -Type Success
        
        $tenantName = $script:TenantState.TenantName
        Write-LogMessage -Message "Creating groups for tenant: $tenantName" -Type Info
        
        # Create license groups
        foreach ($license in $DefaultGroups.License) {
            $displayName = "Microsoft 365 $license Users"
            
            # Check if already exists
            $existingGroup = Get-MgGroup -Filter "displayName eq '$displayName'" -ErrorAction SilentlyContinue
            if ($existingGroup) {
                Write-LogMessage -Message "Group '$displayName' already exists" -Type Warning
                $script:TenantState.CreatedGroups[$displayName] = $existingGroup.Id
                continue
            }
            
            # Create using direct API
            $body = @{
                displayName = $displayName
                description = "Dynamic license group for $license"
                groupTypes = @("DynamicMembership")
                mailEnabled = $false
                mailNickname = "$($license)Users"
                membershipRule = "user.extensionAttribute1 eq `"$license`""
                membershipRuleProcessingState = "On"
                securityEnabled = $true
            }

            try {
                $result = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/groups" -Body $body
                Write-LogMessage -Message "Created dynamic license group: $displayName" -Type Success
                $script:TenantState.CreatedGroups[$displayName] = $result.id
            }
            catch {
                Write-LogMessage -Message "Failed to create $displayName - $($_.Exception.Message)" -Type Error
            }
        }

        # Create domain users group
        if (-not [string]::IsNullOrEmpty($tenantName)) {
            $domainGroupName = "$tenantName Users"
            # Check if already exists
            $existingGroup = Get-MgGroup -Filter "displayName eq '$domainGroupName'" -ErrorAction SilentlyContinue
            
            if (-not $existingGroup) {
                $body = @{
                    displayName = $domainGroupName
                    description = "All users in $tenantName tenant"
                    groupTypes = @("DynamicMembership")
                    mailEnabled = $false
                    mailNickname = "DomainUsers"
                    membershipRule = "user.userType -ne `"Guest`""
                    membershipRuleProcessingState = "On"
                    securityEnabled = $true
                }

                try {
                    $result = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/groups" -Body $body
                    Write-LogMessage -Message "Created domain users group: $domainGroupName" -Type Success
                    $script:TenantState.CreatedGroups[$domainGroupName] = $result.id
                }
                catch {
                    Write-LogMessage -Message "Failed to create domain users group - $($_.Exception.Message)" -Type Error
                }
            }
            else {
                Write-LogMessage -Message "Domain users group already exists" -Type Warning
                $script:TenantState.CreatedGroups[$domainGroupName] = $existingGroup.Id
            }
        }

        # Create regular security groups
        foreach ($name in $DefaultGroups.Security) {
            # Check if already exists
            $existingGroup = Get-MgGroup -Filter "displayName eq '$name'" -ErrorAction SilentlyContinue
            if ($existingGroup) {
                Write-LogMessage -Message "Group '$name' already exists" -Type Warning
                $script:TenantState.CreatedGroups[$name] = $existingGroup.Id
                continue
            }
            
            $mailNick = $name -replace '\s', ''
            
            try {
                $newGroup = New-MgGroup -DisplayName $name -Description "Security group" -MailEnabled:$false -MailNickname $mailNick -SecurityEnabled:$true
                Write-LogMessage -Message "Created security group: $name" -Type Success
                $script:TenantState.CreatedGroups[$name] = $newGroup.Id
            }
            catch {
                Write-LogMessage -Message "Failed to create group $name - $($_.Exception.Message)" -Type Error
            }
        }

        Write-LogMessage -Message "Group creation completed" -Type Success
        return $true
    }
    catch {
        Write-LogMessage -Message "Error in group creation process - $($_.Exception.Message)" -Type Error
        return $false
    }
}