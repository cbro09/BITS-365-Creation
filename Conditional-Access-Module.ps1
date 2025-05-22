# === ConditionalAccess.ps1 ===
# Conditional Access policy creation and management functions

function New-TenantCAPolices {
    Write-LogMessage -Message "Starting CA policy creation process..." -Type Info
    Import-RequiredGraphModules
    
    try {
        # Check for NoMFA Exemption group ID
        $noMfaGroupId = $script:TenantState.CreatedGroups["NoMFA Exemption"]
        if (-not $noMfaGroupId) {
            Write-LogMessage -Message "NoMFA Exemption group not found. Some policies may not be correctly configured." -Type Warning
        }
        
        # Function to check if policy exists using direct API
        function Test-PolicyExists {
            param ([string]$PolicyName)
            
            try {
                $response = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies" -ErrorAction Stop
                
                if ($response.PSObject.Properties.Name -contains "value") {
                    $policies = $response.value
                } else {
                    $policies = @($response)
                }
                
                foreach ($p in $policies) {
                    if ($p.displayName -eq $PolicyName) {
                        return $true
                    }
                }
                return $false
            }
            catch {
                Write-LogMessage -Message "Error checking policies - $($_.Exception.Message)" -Type Error
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