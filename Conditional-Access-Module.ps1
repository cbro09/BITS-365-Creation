# === ConditionalAccess.ps1 ===
# Conditional Access policy creation and management functions

function New-TenantCAPolices {
    Write-LogMessage -Message "Starting CA policy creation process..." -Type Info
    
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
        
        # STEP 5: Force load ONLY the exact modules needed for ConditionalAccess
        $conditionalAccessModules = @(
            'Microsoft.Graph.Identity.DirectoryManagement',
            'Microsoft.Graph.Groups',
            'Microsoft.Graph.Identity.SignIns'
        )
        
        Write-LogMessage -Message "Loading ONLY ConditionalAccess modules in exact order..." -Type Info
        foreach ($module in $conditionalAccessModules) {
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
        
        # STEP 6: Connect with EXACT scopes needed for ConditionalAccess
        $conditionalAccessScopes = @(
            "Policy.ReadWrite.ConditionalAccess",
            "Group.ReadWrite.All",
            "Directory.ReadWrite.All"
        )
        
        Write-LogMessage -Message "Connecting to Microsoft Graph with ConditionalAccess scopes..." -Type Info
        Connect-MgGraph -Scopes $conditionalAccessScopes -ErrorAction Stop | Out-Null
        
        $context = Get-MgContext
        Write-LogMessage -Message "Connected to Microsoft Graph as $($context.Account)" -Type Success
        # Query for NoMFA Exemption group directly using Graph API
        Write-LogMessage -Message "Querying for NoMFA Exemption group..." -Type Info
        $noMfaGroupId = $null
        try {
            $noMfaGroup = Get-MgGroup -Filter "displayName eq 'NoMFA Exemption'" -ErrorAction Stop
            if ($noMfaGroup) {
                $noMfaGroupId = $noMfaGroup.Id
                Write-LogMessage -Message "Found NoMFA Exemption group: $($noMfaGroup.Id)" -Type Success
            }
            else {
                Write-LogMessage -Message "NoMFA Exemption group not found. Some policies may not be correctly configured." -Type Warning
            }
        }
        catch {
            Write-LogMessage -Message "Error querying for NoMFA Exemption group: $($_.Exception.Message)" -Type Warning
        }
        
        # Function to check if policy exists using Graph cmdlets
function Test-PolicyExists {
    param ([string]$PolicyName)
    
    try {
        Write-LogMessage -Message "Checking if policy '$PolicyName' exists..." -Type Info -LogOnly
        $existingPolicy = Get-MgIdentityConditionalAccessPolicy -Filter "displayName eq '$PolicyName'" -ErrorAction Stop
        
        if ($existingPolicy) {
            Write-LogMessage -Message "Policy '$PolicyName' already exists with ID: $($existingPolicy.Id)" -Type Info -LogOnly
            return $true
        }
        else {
            Write-LogMessage -Message "Policy '$PolicyName' does not exist, will create" -Type Info -LogOnly
            return $false
        }
    }
    catch {
        Write-LogMessage -Message "Error checking if policy exists: $($_.Exception.Message)" -Type Warning
        return $false
    }
}

        # Create C001 - Block Legacy Authentication
        $policyName = "C001 - Block Legacy Authentication All Apps"
        if (Test-PolicyExists -PolicyName $policyName) {
            Write-LogMessage -Message "Policy '$policyName' already exists" -Type Warning
        }
        else {
            $body = @{
                displayName = $policyName
                state = "disabled"
                conditions = @{
                    users = @{
                        includeUsers = @("All")
                    }
                    applications = @{
                        includeApplications = @("All")
                    }
                    clientAppTypes = @("exchangeActiveSync", "other")
                }
                grantControls = @{
                    operator = "OR"
                    builtInControls = @("block")
                }
            }
            
            try {
                Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies" -Body $body -ErrorAction Stop
                Write-LogMessage -Message "Created policy: $policyName" -Type Success
            }
            catch {
                Write-LogMessage -Message "Failed to create policy $policyName - $($_.Exception.Message)" -Type Error
            }
        }

        # Create C002 - MFA Required for All Users
        $policyName = "C002 - MFA Required for All Users"
        if (Test-PolicyExists -PolicyName $policyName) {
            Write-LogMessage -Message "Policy '$policyName' already exists" -Type Warning
        }
        else {
            $body = @{
                displayName = $policyName
                state = "enabled"
                conditions = @{
                    users = @{
                        includeUsers = @("All")
                    }
                    applications = @{
                        includeApplications = @("All")
                    }
                    clientAppTypes = @("browser", "mobileAppsAndDesktopClients")
                }
                grantControls = @{
                    operator = "OR"
                    builtInControls = @("mfa")
                }
            }
            
            # Add NoMFA group exclusion if available
            if ($noMfaGroupId) {
                $body.conditions.users.excludeGroups = @($noMfaGroupId)
                Write-LogMessage -Message "Added NoMFA Exemption group to exclusions" -Type Info
            }
            
            try {
                Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies" -Body $body -ErrorAction Stop
                Write-LogMessage -Message "Created policy: $policyName" -Type Success
            }
            catch {
                Write-LogMessage -Message "Failed to create policy $policyName - $($_.Exception.Message)" -Type Error
            }
        }

        # Create C003 - Block Non Corporate Devices
        $policyName = "C003 - Block Non Corporate Devices"
        if (Test-PolicyExists -PolicyName $policyName) {
            Write-LogMessage -Message "Policy '$policyName' already exists" -Type Warning
        }
        else {
            $body = @{
                displayName = $policyName
                state = "enabledForReportingButNotEnforced"
                conditions = @{
                    users = @{
                        includeUsers = @("All")
                        excludeRoles = @("d29b2b05-8046-44ba-8758-1e26182fcf32")  # Global Admin role
                    }
                    applications = @{
                        includeApplications = @("All")
                    }
                    clientAppTypes = @("all")
                    platforms = @{
                        includePlatforms = @("all")
                    }
                }
                grantControls = @{
                    operator = "OR"
                    builtInControls = @("compliantDevice", "domainJoinedDevice")
                }
            }
            
            try {
                Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies" -Body $body -ErrorAction Stop
                Write-LogMessage -Message "Created policy: $policyName" -Type Success
            }
            catch {
                Write-LogMessage -Message "Failed to create policy $policyName - $($_.Exception.Message)" -Type Error
            }
        }

        # Create C004 - Require Password Change and MFA for High Risk Users
        $policyName = "C004 - Require Password Change and MFA for High Risk Users"
        if (Test-PolicyExists -PolicyName $policyName) {
            Write-LogMessage -Message "Policy '$policyName' already exists" -Type Warning
        }
        else {
            $body = @{
                displayName = $policyName
                state = "enabled"
                conditions = @{
                    users = @{
                        includeUsers = @("All")
                    }
                    applications = @{
                        includeApplications = @("All")
                    }
                    userRiskLevels = @("high")
                }
                grantControls = @{
                    operator = "AND"
                    builtInControls = @("mfa", "passwordChange")
                }
            }
            
            try {
                Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies" -Body $body -ErrorAction Stop
                Write-LogMessage -Message "Created policy: $policyName" -Type Success
            }
            catch {
                Write-LogMessage -Message "Failed to create policy $policyName - $($_.Exception.Message)" -Type Error
            }
        }

        # Create C005 - Require MFA for Risky Sign-Ins
        $policyName = "C005 - Require MFA for Risky Sign-Ins"
        if (Test-PolicyExists -PolicyName $policyName) {
            Write-LogMessage -Message "Policy '$policyName' already exists" -Type Warning
        }
        else {
            $body = @{
                displayName = $policyName
                state = "enabled"
                conditions = @{
                    users = @{
                        includeUsers = @("All")
                    }
                    applications = @{
                        includeApplications = @("All")
                    }
                    signInRiskLevels = @("high", "medium")
                }
                grantControls = @{
                    operator = "OR"
                    builtInControls = @("mfa")
                }
            }
            
            try {
                Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies" -Body $body -ErrorAction Stop
                Write-LogMessage -Message "Created policy: $policyName" -Type Success
            }
            catch {
                Write-LogMessage -Message "Failed to create policy $policyName - $($_.Exception.Message)" -Type Error
            }
        }

        Write-LogMessage -Message "CA Policy setup completed" -Type Success
        return $true
    }
    catch {
        Write-LogMessage -Message "Error in CA policy creation process - $($_.Exception.Message)" -Type Error
        return $false
    }
}