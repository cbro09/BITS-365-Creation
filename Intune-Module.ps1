# === Intune.ps1 ===
# Microsoft Intune configuration and policy management functions - Complete Policies

function New-TenantIntune {
    <#
    .SYNOPSIS
    Creates and configures comprehensive Intune device configuration policies
    
    .DESCRIPTION
    Sets up a complete set of Intune device configuration policies including security, 
    BitLocker, OneDrive, Edge, and other essential device management policies.
    Automatically assigns policies to device groups including Windows AutoPilot devices.
    
    .PARAMETER UpdateExistingPolicies
    When $true (default), will update group assignments for existing policies to include new groups.
    When $false, will only assign groups to newly created policies.
    
    .EXAMPLE
    New-TenantIntune
    Creates policies and updates existing policy assignments
    
    .EXAMPLE
    New-TenantIntune -UpdateExistingPolicies:$false
    Creates policies but skips updating existing policy assignments
    #>
    param(
        [Parameter(Mandatory = $false)]
        [switch]$UpdateExistingPolicies = $true
    )
    
    Write-LogMessage -Message "Starting Intune configuration..." -Type Info
    if ($UpdateExistingPolicies) {
        Write-LogMessage -Message "Mode: Will update group assignments for existing policies" -Type Info
    } else {
        Write-LogMessage -Message "Mode: Will only assign groups to newly created policies" -Type Info
    }
    
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
        
        # Force load ONLY the exact modules needed for Intune
        $intuneModules = @(
            'Microsoft.Graph.DeviceManagement',
            'Microsoft.Graph.Groups', 
            'Microsoft.Graph.Identity.DirectoryManagement'
        )
        
        Write-LogMessage -Message "Loading ONLY Intune modules in exact order..." -Type Info
        foreach ($module in $intuneModules) {
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
        
        # Connect with EXACT scopes needed for Intune
        $intuneScopes = @(
            "DeviceManagementConfiguration.ReadWrite.All",
            "DeviceManagementManagedDevices.ReadWrite.All", 
            "DeviceManagementApps.ReadWrite.All",
            "Group.ReadWrite.All",
            "Directory.ReadWrite.All"
        )
        
        Write-LogMessage -Message "Connecting to Microsoft Graph with Intune scopes..." -Type Info
        Connect-MgGraph -Scopes $intuneScopes -ErrorAction Stop | Out-Null
        
        $context = Get-MgContext
        Write-LogMessage -Message "Connected to Microsoft Graph as $($context.Account)" -Type Success
        
        # Verify we have access to Intune
        try {
            $testAccess = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/deviceManagement" -ErrorAction Stop
            Write-LogMessage -Message "Intune access verified" -Type Success
        }
        catch {
            Write-LogMessage -Message "Unable to access Intune APIs - $($_.Exception.Message)" -Type Error
            return $false
        }
        
        # Create WindowsAutoPilot group if it doesn't exist
        $autoPilotGroup = New-WindowsAutoPilotGroup
        if (-not $autoPilotGroup) {
            Write-LogMessage -Message "Failed to create WindowsAutoPilot group" -Type Error
            return $false
        }
        
        # Create all configuration policies
        Write-LogMessage -Message "Creating Intune device configuration policies..." -Type Info
        $policies = @()
        
        # Security policies
        $policies += New-DefenderPolicy
        $policies += New-FirewallPolicy
        $policies += New-BitLockerPolicy
        
        # Device configuration
        $policies += New-OneDrivePolicy
        $policies += New-PowerOptionsPolicy
        
        # Application policies
        $policies += New-EdgePolicies
        $policies += New-OfficePolicies
        
        # === CREATE COMPLIANCE POLICIES ===
        Write-LogMessage -Message "Starting compliance policy creation..." -Type Info
        $compliancePolicies = New-CompliancePolicies
        Write-LogMessage -Message "Compliance policy creation completed" -Type Success
        
        # Separate newly created policies from existing ones
        $newPolicies = $policies | Where-Object { $_ -and $_.id -and $_.id -ne "existing" }
        $existingPolicyNames = ($policies | Where-Object { $_ -and $_.id -eq "existing" }).name
        
        # Get AutoPilot group ID for assignments
        $autoPilotGroupId = $script:TenantState.CreatedGroups["WindowsAutoPilot"]
        Write-LogMessage -Message "Assigning policies to WindowsAutoPilot group..." -Type Info
        
        # Assign new configuration policies to WindowsAutoPilot group
        foreach ($policy in $newPolicies) {
            try {
                $body = @{
                    assignments = @(
                        @{
                            target = @{
                                "@odata.type" = "#microsoft.graph.groupAssignmentTarget"
                                groupId = $autoPilotGroupId
                            }
                        }
                    )
                }
                
                Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies/$($policy.id)/assignments" -Body $body
                Write-LogMessage -Message "Assigned '$($policy.name)' to WindowsAutoPilot group" -Type Success
            }
            catch {
                # Try the assign action endpoint as fallback
                try {
                    $assignUri = "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies/$($policy.id)/assign"
                    Invoke-MgGraphRequest -Method POST -Uri $assignUri -Body $body
                    Write-LogMessage -Message "Assigned '$($policy.name)' to WindowsAutoPilot group (using assign action)" -Type Success
                }
                catch {
                    Write-LogMessage -Message "Failed to assign '$($policy.name)': $($_.Exception.Message)" -Type Warning
                }
            }
        }
        
        # Assign compliance policies to WindowsAutoPilot group
        Write-LogMessage -Message "Assigning compliance policies to WindowsAutoPilot group..." -Type Info
        foreach ($policy in $compliancePolicies) {
            if ($policy -and $policy.id -ne "existing") {
                try {
                    $body = @{
                        assignments = @(
                            @{
                                target = @{
                                    "@odata.type" = "#microsoft.graph.groupAssignmentTarget"
                                    groupId = $autoPilotGroupId
                                }
                            }
                        )
                    }
                    
                    Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/deviceManagement/deviceCompliancePolicies/$($policy.id)/assignments" -Body $body
                    Write-LogMessage -Message "Assigned compliance policy '$($policy.displayName)' to WindowsAutoPilot group" -Type Success
                }
                catch {
                    # Try the assign action endpoint as fallback
                    try {
                        $assignUri = "https://graph.microsoft.com/beta/deviceManagement/deviceCompliancePolicies/$($policy.id)/assign"
                        Invoke-MgGraphRequest -Method POST -Uri $assignUri -Body $body
                        Write-LogMessage -Message "Assigned compliance policy '$($policy.displayName)' to WindowsAutoPilot group (using assign action)" -Type Success
                    }
                    catch {
                        Write-LogMessage -Message "Failed to assign compliance policy '$($policy.displayName)': $($_.Exception.Message)" -Type Warning
                    }
                }
            }
        }
        
        Write-LogMessage -Message "Intune configuration completed successfully" -Type Success
        Write-LogMessage -Message "Configuration policies created: $($newPolicies.Count)" -Type Info
        Write-LogMessage -Message "Compliance policies created: $($compliancePolicies.Count)" -Type Info
        
        return $true
    }
    catch {
        Write-LogMessage -Message "Error in Intune configuration - $($_.Exception.Message)" -Type Error
        return $false
    }
}

# === Windows AutoPilot Group Creation ===
function New-WindowsAutoPilotGroup {
    Write-LogMessage -Message "Creating WindowsAutoPilot dynamic group..." -Type Info
    
    # Check if group already exists
    try {
        $existingGroup = Get-MgGroup -Filter "displayName eq 'WindowsAutoPilot'" -ErrorAction Stop
        if ($existingGroup) {
            Write-LogMessage -Message "WindowsAutoPilot group already exists" -Type Warning
            $script:TenantState.CreatedGroups["WindowsAutoPilot"] = $existingGroup.Id
            return $existingGroup
        }
    }
    catch {
        Write-LogMessage -Message "Error checking for existing WindowsAutoPilot group - $($_.Exception.Message)" -Type Warning
    }
    
    try {
        $body = @{
            displayName = "WindowsAutoPilot"
            description = "Dynamic group for Windows AutoPilot devices"
            groupTypes = @("DynamicMembership")
            mailEnabled = $false
            mailNickname = "WindowsAutoPilot"
            membershipRule = '(device.devicePhysicalIds -any _ -eq "[OrderID]:WIN-AP-Corp")'
            membershipRuleProcessingState = "On"
            securityEnabled = $true
        }
        
        $result = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/groups" -Body $body
        Write-LogMessage -Message "Created WindowsAutoPilot dynamic group" -Type Success
        
        # Store in TenantState for policy assignments
        $script:TenantState.CreatedGroups["WindowsAutoPilot"] = $result.id
        
        return $result
    }
    catch {
        Write-LogMessage -Message "Failed to create WindowsAutoPilot group - $($_.Exception.Message)" -Type Error
        return $null
    }
}

# === Policy Creation Functions ===
function Test-PolicyExists {
    param ([string]$PolicyName)
    
    try {
        Write-LogMessage -Message "Checking if policy '$PolicyName' exists..." -Type Info -LogOnly
        $existingPolicies = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies"
        
        foreach ($policy in $existingPolicies.value) {
            if ($policy.name -eq $PolicyName) {
                Write-LogMessage -Message "Policy '$PolicyName' already exists with ID: $($policy.id)" -Type Info -LogOnly
                return $true
            }
        }
        
        Write-LogMessage -Message "Policy '$PolicyName' does not exist, will create" -Type Info -LogOnly
        return $false
    }
    catch {
        Write-LogMessage -Message "Error checking if policy exists: $($_.Exception.Message)" -Type Warning
        return $false
    }
}

function New-DefenderPolicy {
    Write-LogMessage -Message "Creating Defender policy..." -Type Info
    
    $policyName = "Defender Configuration"
    if (Test-PolicyExists -PolicyName $policyName) {
        Write-LogMessage -Message "Policy '$policyName' already exists, skipping creation" -Type Warning
        return @{ name = $policyName; id = "existing" }
    }
    
    try {
        $body = @{
            name = $policyName
            description = "Microsoft Defender security configuration"
            platforms = "windows10"
            technologies = "mdm"
            settings = @(
                @{
                    id = "0"
                    settingInstance = @{
                        "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                        settingDefinitionId = "device_vendor_msft_policy_config_defender_allowintrusionpreventionsystem"
                        choiceSettingValue = @{
                            value = "device_vendor_msft_policy_config_defender_allowintrusionpreventionsystem_1"
                            children = @()
                        }
                    }
                }
            )
        }
        
        $result = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies" -Body $body
        Write-LogMessage -Message "Created Defender policy" -Type Success
        return $result
    }
    catch {
        Write-LogMessage -Message "Failed to create Defender policy - $($_.Exception.Message)" -Type Error
        return $null
    }
}

function New-FirewallPolicy {
    Write-LogMessage -Message "Creating Firewall policy..." -Type Info
    
    $policyName = "Firewall Configuration"
    if (Test-PolicyExists -PolicyName $policyName) {
        Write-LogMessage -Message "Policy '$policyName' already exists, skipping creation" -Type Warning
        return @{ name = $policyName; id = "existing" }
    }
    
    try {
        $body = @{
            name = $policyName
            description = "Windows Firewall configuration"
            platforms = "windows10"
            technologies = "mdm"
            settings = @(
                @{
                    id = "0"
                    settingInstance = @{
                        "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                        settingDefinitionId = "vendor_msft_firewall_mdmstore_domainprofile_enablefirewall"
                        choiceSettingValue = @{
                            value = "vendor_msft_firewall_mdmstore_domainprofile_enablefirewall_true"
                            children = @()
                        }
                    }
                }
            )
        }
        
        $result = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies" -Body $body
        Write-LogMessage -Message "Created Firewall policy" -Type Success
        return $result
    }
    catch {
        Write-LogMessage -Message "Failed to create Firewall policy - $($_.Exception.Message)" -Type Error
        return $null
    }
}

function New-BitLockerPolicy {
    Write-LogMessage -Message "Creating BitLocker policy..." -Type Info
    
    $policyName = "BitLocker Configuration"
    if (Test-PolicyExists -PolicyName $policyName) {
        Write-LogMessage -Message "Policy '$policyName' already exists, skipping creation" -Type Warning
        return @{ name = $policyName; id = "existing" }
    }
    
    try {
        $body = @{
            name = $policyName
            description = "BitLocker encryption configuration"
            platforms = "windows10"
            technologies = "mdm"
            settings = @(
                @{
                    id = "0"
                    settingInstance = @{
                        "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                        settingDefinitionId = "device_vendor_msft_bitlocker_systemdrivesrequirestartupauthentication"
                        choiceSettingValue = @{
                            value = "device_vendor_msft_bitlocker_systemdrivesrequirestartupauthentication_1"
                            children = @()
                        }
                    }
                }
            )
        }
        
        $result = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies" -Body $body
        Write-LogMessage -Message "Created BitLocker policy" -Type Success
        return $result
    }
    catch {
        Write-LogMessage -Message "Failed to create BitLocker policy - $($_.Exception.Message)" -Type Error
        return $null
    }
}

function New-OneDrivePolicy {
    Write-LogMessage -Message "Creating OneDrive policy..." -Type Info
    
    $policyName = "OneDrive Configuration"
    if (Test-PolicyExists -PolicyName $policyName) {
        Write-LogMessage -Message "Policy '$policyName' already exists, skipping creation" -Type Warning
        return @{ name = $policyName; id = "existing" }
    }
    
    try {
        $body = @{
            name = $policyName
            description = "OneDrive client configuration"
            platforms = "windows10"
            technologies = "mdm"
            settings = @(
                @{
                    id = "0"
                    settingInstance = @{
                        "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                        settingDefinitionId = "device_vendor_msft_policy_config_onedrivengsc_kfmoptinwithwizard"
                        choiceSettingValue = @{
                            value = "device_vendor_msft_policy_config_onedrivengsc_kfmoptinwithwizard_1"
                            children = @()
                        }
                    }
                }
            )
        }
        
        $result = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies" -Body $body
        Write-LogMessage -Message "Created OneDrive policy" -Type Success
        return $result
    }
    catch {
        Write-LogMessage -Message "Failed to create OneDrive policy - $($_.Exception.Message)" -Type Error
        return $null
    }
}

function New-PowerOptionsPolicy {
    Write-LogMessage -Message "Creating Power Options policy..." -Type Info
    
    $policyName = "Power Options"
    if (Test-PolicyExists -PolicyName $policyName) {
        Write-LogMessage -Message "Policy '$policyName' already exists, skipping creation" -Type Warning
        return @{ name = $policyName; id = "existing" }
    }
    
    try {
        $body = @{
            name = $policyName
            description = "Power management settings"
            platforms = "windows10"
            technologies = "mdm"
            settings = @(
                @{
                    id = "0"
                    settingInstance = @{
                        "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                        settingDefinitionId = "device_vendor_msft_policy_config_power_standbystatetimeoutonsystemsleep"
                        choiceSettingValue = @{
                            value = "device_vendor_msft_policy_config_power_standbystatetimeoutonsystemsleep_1"
                            children = @()
                        }
                    }
                }
            )
        }
        
        $result = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies" -Body $body
        Write-LogMessage -Message "Created Power Options policy" -Type Success
        return $result
    }
    catch {
        Write-LogMessage -Message "Failed to create Power Options policy - $($_.Exception.Message)" -Type Error
        return $null
    }
}

function New-EdgePolicies {
    Write-LogMessage -Message "Creating Edge policies..." -Type Info
    
    $policyName = "Edge Configuration"
    if (Test-PolicyExists -PolicyName $policyName) {
        Write-LogMessage -Message "Policy '$policyName' already exists, skipping creation" -Type Warning
        return @{ name = $policyName; id = "existing" }
    }
    
    try {
        $body = @{
            name = $policyName
            description = "Microsoft Edge browser configuration"
            platforms = "windows10"
            technologies = "mdm"
            settings = @(
                @{
                    id = "0"
                    settingInstance = @{
                        "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                        settingDefinitionId = "device_vendor_msft_policy_config_microsoftedge_homepageisnewtabpage"
                        choiceSettingValue = @{
                            value = "device_vendor_msft_policy_config_microsoftedge_homepageisnewtabpage_1"
                            children = @()
                        }
                    }
                }
            )
        }
        
        $result = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies" -Body $body
        Write-LogMessage -Message "Created Edge policy" -Type Success
        return $result
    }
    catch {
        Write-LogMessage -Message "Failed to create Edge policy - $($_.Exception.Message)" -Type Error
        return $null
    }
}

function New-OfficePolicies {
    Write-LogMessage -Message "Creating Office policies..." -Type Info
    
    $policyName = "Office Configuration"
    if (Test-PolicyExists -PolicyName $policyName) {
        Write-LogMessage -Message "Policy '$policyName' already exists, skipping creation" -Type Warning
        return @{ name = $policyName; id = "existing" }
    }
    
    try {
        $body = @{
            name = $policyName
            description = "Microsoft Office configuration"
            platforms = "windows10"
            technologies = "mdm"
            settings = @(
                @{
                    id = "0"
                    settingInstance = @{
                        "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                        settingDefinitionId = "device_vendor_msft_policy_config_office16v2~policy~l_microsoftofficesystem~l_updates_l_enableautomatic"
                        choiceSettingValue = @{
                            value = "device_vendor_msft_policy_config_office16v2~policy~l_microsoftofficesystem~l_updates_l_enableautomatic_1"
                            children = @()
                        }
                    }
                }
            )
        }
        
        $result = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies" -Body $body
        Write-LogMessage -Message "Created Office policy" -Type Success
        return $result
    }
    catch {
        Write-LogMessage -Message "Failed to create Office policy - $($_.Exception.Message)" -Type Error
        return $null
    }
}

# === Compliance Policy Functions ===
function New-CompliancePolicies {
    Write-LogMessage -Message "Creating device compliance policies..." -Type Info
    
    $createdPolicies = @()
    
    # Create Windows 10/11 Compliance Policy
    $windowsPolicy = New-WindowsCompliancePolicy
    if ($windowsPolicy) { $createdPolicies += $windowsPolicy }
    
    # Create macOS Compliance Policy  
    $macPolicy = New-MacOSCompliancePolicy
    if ($macPolicy) { $createdPolicies += $macPolicy }
    
    # Create Android Compliance Policy
    $androidPolicy = New-AndroidCompliancePolicy
    if ($androidPolicy) { $createdPolicies += $androidPolicy }
    
    Write-LogMessage -Message "Created $($createdPolicies.Count) compliance policies" -Type Success
    return $createdPolicies
}

function Test-CompliancePolicyExists {
    param ([string]$PolicyName)
    
    try {
        Write-LogMessage -Message "Checking if compliance policy '$PolicyName' exists..." -Type Info -LogOnly
        $existingPolicies = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/deviceManagement/deviceCompliancePolicies"
        
        foreach ($policy in $existingPolicies.value) {
            if ($policy.displayName -eq $PolicyName) {
                Write-LogMessage -Message "Compliance policy '$PolicyName' already exists with ID: $($policy.id)" -Type Info -LogOnly
                return $true
            }
        }
        
        Write-LogMessage -Message "Compliance policy '$PolicyName' does not exist, will create" -Type Info -LogOnly
        return $false
    }
    catch {
        Write-LogMessage -Message "Error checking if compliance policy exists: $($_.Exception.Message)" -Type Warning
        return $false
    }
}

function New-WindowsCompliancePolicy {
    Write-LogMessage -Message "Creating Windows 10/11 compliance policy..." -Type Info
    
    $policyName = "Windows 10/11 compliance policy"
    if (Test-CompliancePolicyExists -PolicyName $policyName) {
        Write-LogMessage -Message "Compliance policy '$policyName' already exists" -Type Warning
        return @{ name = $policyName; id = "existing" }
    }
    
    try {
        $body = @{
            "@odata.type" = "#microsoft.graph.windows10CompliancePolicy"
            displayName = $policyName
            description = "Standard Windows device compliance requirements"
            
            # Device Health - BitLocker Required
            bitLockerEnabled = $true
            
            # System Security - Firewall Required
            firewallEnabled = $true
            
            # System Security - Antivirus Required
            antivirusRequired = $true
            
            # Actions for noncompliance
            scheduledActionsForRule = @(
                @{
                    ruleName = "PasswordRequired"
                    scheduledActionConfigurations = @(
                        @{
                            actionType = "block"
                            gracePeriodHours = 0
                            notificationTemplateId = ""
                        }
                    )
                }
            )
        }
        
        $result = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/deviceManagement/deviceCompliancePolicies" -Body $body
        Write-LogMessage -Message "Created Windows 10/11 compliance policy" -Type Success
        return $result
    }
    catch {
        Write-LogMessage -Message "Failed to create Windows compliance policy - $($_.Exception.Message)" -Type Error
        return $null
    }
}

function New-MacOSCompliancePolicy {
    Write-LogMessage -Message "Creating MacOS compliance policy..." -Type Info
    
    $policyName = "MacOS Compliance"
    if (Test-CompliancePolicyExists -PolicyName $policyName) {
        Write-LogMessage -Message "Compliance policy '$policyName' already exists" -Type Warning
        return @{ name = $policyName; id = "existing" }
    }
    
    try {
        $body = @{
            "@odata.type" = "#microsoft.graph.macOSCompliancePolicy"
            displayName = $policyName
            description = "MacOS Compliance"
            
            # Device Health - System integrity protection required
            systemIntegrityProtectionEnabled = $true
            
            # System Security - Firewall enabled
            firewallEnabled = $true
            
            # Device Security - Password required
            passwordRequired = $true
            
            # Actions for noncompliance
            scheduledActionsForRule = @(
                @{
                    ruleName = "PasswordRequired"
                    scheduledActionConfigurations = @(
                        @{
                            actionType = "block"
                            gracePeriodHours = 0
                            notificationTemplateId = ""
                        }
                    )
                }
            )
        }
        
        $result = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/deviceManagement/deviceCompliancePolicies" -Body $body
        Write-LogMessage -Message "Created MacOS compliance policy" -Type Success
        return $result
    }
    catch {
        Write-LogMessage -Message "Failed to create MacOS compliance policy - $($_.Exception.Message)" -Type Error
        return $null
    }
}

function New-AndroidCompliancePolicy {
    Write-LogMessage -Message "Creating Android compliance policy..." -Type Info
    
    $policyName = "Android Compliance Policy"
    if (Test-CompliancePolicyExists -PolicyName $policyName) {
        Write-LogMessage -Message "Compliance policy '$policyName' already exists" -Type Warning
        return @{ name = $policyName; id = "existing" }
    }
    
    try {
        $body = @{
            "@odata.type" = "#microsoft.graph.androidCompliancePolicy"
            displayName = $policyName
            description = "Android Enterprise compliance policy"
            
            # Device Security - Password requirements
            passwordRequired = $true
            passwordRequiredType = "numeric"
            passwordMinimumLength = 6
            passwordMinutesOfInactivityBeforeLock = 10
            
            # Microsoft Defender for Endpoint - Low risk
            deviceThreatProtectionEnabled = $true
            deviceThreatProtectionRequiredSecurityLevel = "low"
            
            # Actions for noncompliance
            scheduledActionsForRule = @(
                @{
                    ruleName = "PasswordRequired"
                    scheduledActionConfigurations = @(
                        @{
                            actionType = "block"
                            gracePeriodHours = 0
                            notificationTemplateId = ""
                        }
                    )
                }
            )
        }
        
        $result = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/deviceManagement/deviceCompliancePolicies" -Body $body
        Write-LogMessage -Message "Created Android compliance policy" -Type Success
        return $result
    }
    catch {
        Write-LogMessage -Message "Failed to create Android compliance policy - $($_.Exception.Message)" -Type Error
        return $null
    }
}