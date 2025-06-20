# === Intune.ps1 ===
# Microsoft Intune configuration and policy management functions - Complete Policies
# Direct Execution Method - All functionality preserved

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
        
        # ===================================================================
        # HELPER FUNCTIONS (Inline)
        # ===================================================================
        
        # Policy existence checking
        function Test-PolicyExists {
            param ([string]$PolicyName)
            try {
                $existingPolicies = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies" -ErrorAction Stop
                foreach ($policy in $existingPolicies.value) {
                    if ($policy.name -eq $PolicyName) { return $true }
                }
                return $false
            }
            catch { return $false }
        }
        
        # Compliance policy existence checking
        function Test-CompliancePolicyExists {
            param ([string]$PolicyName)
            try {
                $existingPolicies = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/deviceManagement/deviceCompliancePolicies" -ErrorAction Stop
                foreach ($policy in $existingPolicies.value) {
                    if ($policy.displayName -eq $PolicyName) { return $true }
                }
                return $false
            }
            catch { return $false }
        }
        
        # Get SharePoint root site URL
        function Get-SharePointRootSiteUrl {
            try {
                if ($script:TenantState -and $script:TenantState.DefaultDomain) {
                    $domain = $script:TenantState.DefaultDomain
                    $tenantName = $domain.Split('.')[0]
                    return "https://$tenantName.sharepoint.com"
                }
                return "https://www.office.com"
            }
            catch {
                return "https://www.office.com"
            }
        }
        
        # ===================================================================
        # 1. DIRECT EXECUTION: Create All Device Platform Groups
        # ===================================================================
        
        # WindowsAutoPilot Group
        Write-LogMessage -Message "Creating WindowsAutoPilot dynamic group..." -Type Info
        $autopilotGroup = $null
        try {
            $existingGroup = Get-MgGroup -Filter "displayName eq 'WindowsAutoPilot'" -ErrorAction SilentlyContinue
            
            if ($existingGroup) {
                Write-LogMessage -Message "WindowsAutoPilot group already exists" -Type Warning
                $autopilotGroup = $existingGroup
                $script:TenantState.CreatedGroups["WindowsAutoPilot"] = $existingGroup.Id
            }
            else {
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
                
                $autopilotGroup = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/groups" -Body $body
                Write-LogMessage -Message "Created WindowsAutoPilot dynamic group" -Type Success
                $script:TenantState.CreatedGroups["WindowsAutoPilot"] = $autopilotGroup.id
            }
        }
        catch {
            Write-LogMessage -Message "Failed to create WindowsAutoPilot group - $($_.Exception.Message)" -Type Error
            $autopilotGroup = $null
        }
        
        # MacOS Devices Group
        Write-LogMessage -Message "Creating MacOS Devices dynamic group..." -Type Info
        $macosGroup = $null
        try {
            $existingGroup = Get-MgGroup -Filter "displayName eq 'MacOS Devices'" -ErrorAction SilentlyContinue
            
            if ($existingGroup) {
                Write-LogMessage -Message "MacOS Devices group already exists" -Type Warning
                $macosGroup = $existingGroup
                $script:TenantState.CreatedGroups["MacOSDevices"] = $existingGroup.Id
            }
            else {
                $body = @{
                    displayName = "MacOS Devices"
                    description = "Dynamic group for macOS devices"
                    groupTypes = @("DynamicMembership")
                    mailEnabled = $false
                    mailNickname = "MacOSDevices"
                    membershipRule = '(device.deviceOSType -eq "MacMDM")'
                    membershipRuleProcessingState = "On"
                    securityEnabled = $true
                }
                
                $macosGroup = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/groups" -Body $body
                Write-LogMessage -Message "Created MacOS Devices dynamic group" -Type Success
                $script:TenantState.CreatedGroups["MacOSDevices"] = $macosGroup.id
            }
        }
        catch {
            Write-LogMessage -Message "Failed to create MacOS Devices group - $($_.Exception.Message)" -Type Error
            $macosGroup = $null
        }
        
        # Android Devices Group  
        Write-LogMessage -Message "Creating Android Devices dynamic group..." -Type Info
        $androidGroup = $null
        try {
            $existingGroup = Get-MgGroup -Filter "displayName eq 'Android Devices'" -ErrorAction SilentlyContinue
            
            if ($existingGroup) {
                Write-LogMessage -Message "Android Devices group already exists" -Type Warning
                $androidGroup = $existingGroup
                $script:TenantState.CreatedGroups["AndroidDevices"] = $existingGroup.Id
            }
            else {
                $body = @{
                    displayName = "Android Devices"
                    description = "Dynamic group for Android devices"
                    groupTypes = @("DynamicMembership")
                    mailEnabled = $false
                    mailNickname = "AndroidDevices"
                    membershipRule = '(device.deviceOSType -eq "Android")'
                    membershipRuleProcessingState = "On"
                    securityEnabled = $true
                }
                
                $androidGroup = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/groups" -Body $body
                Write-LogMessage -Message "Created Android Devices dynamic group" -Type Success
                $script:TenantState.CreatedGroups["AndroidDevices"] = $androidGroup.id
            }
        }
        catch {
            Write-LogMessage -Message "Failed to create Android Devices group - $($_.Exception.Message)" -Type Error
            $androidGroup = $null
        }
        
        # iOS Devices Group
        Write-LogMessage -Message "Creating iOS Devices dynamic group..." -Type Info
        $iosGroup = $null
        try {
            $existingGroup = Get-MgGroup -Filter "displayName eq 'iOS Devices'" -ErrorAction SilentlyContinue
            
            if ($existingGroup) {
                Write-LogMessage -Message "iOS Devices group already exists" -Type Warning
                $iosGroup = $existingGroup
                $script:TenantState.CreatedGroups["iOSDevices"] = $existingGroup.Id
            }
            else {
                $body = @{
                    displayName = "iOS Devices"
                    description = "Dynamic group for iOS devices"
                    groupTypes = @("DynamicMembership")
                    mailEnabled = $false
                    mailNickname = "iOSDevices"
                    membershipRule = '(device.deviceOSType -eq "iPad") or (device.deviceOSType -eq "iPhone")'
                    membershipRuleProcessingState = "On"
                    securityEnabled = $true
                }
                
                $iosGroup = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/groups" -Body $body
                Write-LogMessage -Message "Created iOS Devices dynamic group" -Type Success
                $script:TenantState.CreatedGroups["iOSDevices"] = $iosGroup.id
            }
        }
        catch {
            Write-LogMessage -Message "Failed to create iOS Devices group - $($_.Exception.Message)" -Type Error
            $iosGroup = $null
        }
        
        # ===================================================================
        # 2. DIRECT EXECUTION: Enable Windows LAPS
        # ===================================================================
        Write-LogMessage -Message "Checking Windows LAPS prerequisite..." -Type Info
        
        $lapsEnabled = $false
        try {
            $lapsSettings = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/policies/deviceRegistrationPolicy" -ErrorAction SilentlyContinue
            
            if ($lapsSettings -and $lapsSettings.localAdminPassword -and $lapsSettings.localAdminPassword.isEnabled) {
                Write-LogMessage -Message "Windows LAPS is already enabled" -Type Info
                $lapsEnabled = $true
            }
            else {
                $body = @{
                    localAdminPassword = @{
                        isEnabled = $true
                    }
                }
                
                Invoke-MgGraphRequest -Method PATCH -Uri "https://graph.microsoft.com/beta/policies/deviceRegistrationPolicy" -Body $body
                Write-LogMessage -Message "Windows LAPS has been enabled" -Type Success
                $lapsEnabled = $true
            }
        }
        catch {
            Write-LogMessage -Message "Failed to enable Windows LAPS - $($_.Exception.Message)" -Type Error
            $lapsEnabled = $false
        }
        
        if (-not $lapsEnabled) {
            Write-LogMessage -Message "LAPS enablement failed - LAPS policies may not work correctly" -Type Warning
        }
        
        # ===================================================================
        # 3. DIRECT EXECUTION: Create All Configuration Policies
        # ===================================================================
        Write-LogMessage -Message "Creating comprehensive configuration policies..." -Type Info
        $policies = @()
        
        # ===================================================================
        # POLICY 1: DEFENDER CONFIGURATION
        # ===================================================================
        Write-LogMessage -Message "Creating comprehensive Defender policy..." -Type Info
        
        $policyName = "Defender Configuration"
        if (Test-PolicyExists -PolicyName $policyName) {
            Write-LogMessage -Message "Policy '$policyName' already exists, skipping creation" -Type Warning
            $policies += @{ name = $policyName; id = "existing" }
        }
        else {
            try {
                $body = @{
                    name = $policyName
                    description = "Microsoft Defender comprehensive security configuration"
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
                $policies += $result
            }
            catch {
                Write-LogMessage -Message "Failed to create Defender policy - $($_.Exception.Message)" -Type Error
            }
        }
        
        # ===================================================================
        # POLICY 2: DEFENDER ANTIVIRUS POLICY (NGP Windows default policy) - ALL 27 SETTINGS
        # ===================================================================
        Write-LogMessage -Message "Creating comprehensive Defender Antivirus policy with 27 settings..." -Type Info
        
        $policyName = "NGP Windows default policy"
        if (Test-PolicyExists -PolicyName $policyName) {
            Write-LogMessage -Message "Policy '$policyName' already exists, skipping creation" -Type Warning
            $policies += @{ name = $policyName; id = "existing" }
        }
        else {
            try {
                $body = @{
                    name = $policyName
                    description = "Default policy sets settings for all endpoints that are not governed by any other policy, ensuring that all your clients are managed as soon as MDE is deployed. The default policy is based on a set of pre-configured recommended settings and can be adjusted by user with admin priviledges."
                    platforms = "windows10"
                    technologies = "mdm,microsoftSense"
                    templateReference = @{
                        templateId = "804339ad-1553-4478-a742-138fb5807418_1"
                        templateFamily = "endpointSecurityAntivirus"
                        templateDisplayName = "Microsoft Defender Antivirus"
                        templateDisplayVersion = "Version 1"
                    }
                    settings = @(
                        @{
                            id = "0"
                            settingInstance = @{
                                "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                settingDefinitionId = "device_vendor_msft_policy_config_defender_puaprotection"
                                settingInstanceTemplateReference = @{
                                    settingInstanceTemplateId = "c0135c2a-f802-44f4-9b71-b0b976411b8c"
                                }
                                choiceSettingValue = @{
                                    value = "device_vendor_msft_policy_config_defender_puaprotection_1"
                                    settingValueTemplateReference = @{
                                        settingValueTemplateId = "2d790211-18cb-4e32-b8cc-97407e2c0b45"
                                        useTemplateDefault = $false
                                    }
                                    children = @()
                                }
                            }
                        },
                        @{
                            id = "1"
                            settingInstance = @{
                                "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                settingDefinitionId = "device_vendor_msft_policy_config_defender_cloudblocklevel"
                                settingInstanceTemplateReference = @{
                                    settingInstanceTemplateId = "c7a37009-c16e-4145-84c8-89a8c121fb15"
                                }
                                choiceSettingValue = @{
                                    value = "device_vendor_msft_policy_config_defender_cloudblocklevel_2"
                                    settingValueTemplateReference = @{
                                        settingValueTemplateId = "517b4e84-e933-42b9-b92f-00e640b1a82d"
                                        useTemplateDefault = $false
                                    }
                                    children = @()
                                }
                            }
                        },
                        @{
                            id = "2"
                            settingInstance = @{
                                "@odata.type" = "#microsoft.graph.deviceManagementConfigurationSimpleSettingInstance"
                                settingDefinitionId = "device_vendor_msft_policy_config_defender_cloudextendedtimeout"
                                settingInstanceTemplateReference = @{
                                    settingInstanceTemplateId = "f61c2788-14e4-4e80-a5a7-bf2ff5052f63"
                                }
                                simpleSettingValue = @{
                                    "@odata.type" = "#microsoft.graph.deviceManagementConfigurationIntegerSettingValue"
                                    value = 50
                                    settingValueTemplateReference = @{
                                        settingValueTemplateId = "608f1561-b603-46bd-bf5f-0b9872002f75"
                                        useTemplateDefault = $false
                                    }
                                }
                            }
                        },
                        @{
                            id = "3"
                            settingInstance = @{
                                "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                settingDefinitionId = "device_vendor_msft_policy_config_defender_enablenetworkprotection"
                                settingInstanceTemplateReference = @{
                                    settingInstanceTemplateId = "f53ab20e-8af6-48f5-9fa1-46863e1e517e"
                                }
                                choiceSettingValue = @{
                                    value = "device_vendor_msft_policy_config_defender_enablenetworkprotection_1"
                                    settingValueTemplateReference = @{
                                        settingValueTemplateId = "ee58fb51-9ae5-408b-9406-b92b643f388a"
                                        useTemplateDefault = $false
                                    }
                                    children = @()
                                }
                            }
                        },
                        @{
                            id = "4"
                            settingInstance = @{
                                "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                settingDefinitionId = "device_vendor_msft_policy_config_defender_allowrealtimemonitoring"
                                settingInstanceTemplateReference = @{
                                    settingInstanceTemplateId = "f0790e28-9231-4d37-8f44-84bb47ca1b3e"
                                }
                                choiceSettingValue = @{
                                    value = "device_vendor_msft_policy_config_defender_allowrealtimemonitoring_1"
                                    settingValueTemplateReference = @{
                                        settingValueTemplateId = "0492c452-1069-4b91-9363-93b8e006ab12"
                                        useTemplateDefault = $false
                                    }
                                    children = @()
                                }
                            }
                        },
                        @{
                            id = "5"
                            settingInstance = @{
                                "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                settingDefinitionId = "device_vendor_msft_policy_config_defender_schedulescanday"
                                settingInstanceTemplateReference = @{
                                    settingInstanceTemplateId = "087d3362-7e78-4983-96bc-1f4ea183f0e4"
                                }
                                choiceSettingValue = @{
                                    value = "device_vendor_msft_policy_config_defender_schedulescanday_2"
                                    settingValueTemplateReference = @{
                                        settingValueTemplateId = "7f4d9dda-6d48-4353-90ca-9fa7164c7215"
                                        useTemplateDefault = $false
                                    }
                                    children = @()
                                }
                            }
                        },
                        @{
                            id = "6"
                            settingInstance = @{
                                "@odata.type" = "#microsoft.graph.deviceManagementConfigurationSimpleSettingInstance"
                                settingDefinitionId = "device_vendor_msft_policy_config_defender_schedulequickscantime"
                                settingInstanceTemplateReference = @{
                                    settingInstanceTemplateId = "784a4af1-33fa-45f2-b945-138b7ff3bcb6"
                                }
                                simpleSettingValue = @{
                                    "@odata.type" = "#microsoft.graph.deviceManagementConfigurationIntegerSettingValue"
                                    value = 720
                                    settingValueTemplateReference = @{
                                        settingValueTemplateId = "5d5c55c8-1a4e-4272-830d-8dc64cd3ac03"
                                        useTemplateDefault = $false
                                    }
                                }
                            }
                        },
                        @{
                            id = "7"
                            settingInstance = @{
                                "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                settingDefinitionId = "device_vendor_msft_policy_config_defender_scanparameter"
                                settingInstanceTemplateReference = @{
                                    settingInstanceTemplateId = "27ca2652-46f3-4cc7-83f2-bf85ff722d84"
                                }
                                choiceSettingValue = @{
                                    value = "device_vendor_msft_policy_config_defender_scanparameter_1"
                                    settingValueTemplateReference = @{
                                        settingValueTemplateId = "70c8f42e-ee6a-4ef1-a070-cb0e9d472581"
                                        useTemplateDefault = $false
                                    }
                                    children = @()
                                }
                            }
                        },
                        @{
                            id = "8"
                            settingInstance = @{
                                "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                settingDefinitionId = "device_vendor_msft_policy_config_defender_allowarchivescanning"
                                settingInstanceTemplateReference = @{
                                    settingInstanceTemplateId = "7c5c9cde-f74d-4d11-904f-de4c27f72d89"
                                }
                                choiceSettingValue = @{
                                    value = "device_vendor_msft_policy_config_defender_allowarchivescanning_1"
                                    settingValueTemplateReference = @{
                                        settingValueTemplateId = "9ead75d4-6f30-4bc5-8cc5-ab0f999d79f0"
                                        useTemplateDefault = $false
                                    }
                                    children = @()
                                }
                            }
                        },
                        @{
                            id = "9"
                            settingInstance = @{
                                "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                settingDefinitionId = "device_vendor_msft_policy_config_defender_enablelowcpupriority"
                                settingInstanceTemplateReference = @{
                                    settingInstanceTemplateId = "cdeb96cf-18f5-4477-a710-0ea9ecc618af"
                                }
                                choiceSettingValue = @{
                                    value = "device_vendor_msft_policy_config_defender_enablelowcpupriority_0"
                                    settingValueTemplateReference = @{
                                        settingValueTemplateId = "045a4a13-deee-4e24-9fe4-985c9357680d"
                                        useTemplateDefault = $false
                                    }
                                    children = @()
                                }
                            }
                        },
                        @{
                            id = "10"
                            settingInstance = @{
                                "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                settingDefinitionId = "device_vendor_msft_policy_config_defender_disablecatchupfullscan"
                                settingInstanceTemplateReference = @{
                                    settingInstanceTemplateId = "f881b08c-f047-40d2-b7d9-3dde7ce9ef64"
                                }
                                choiceSettingValue = @{
                                    value = "device_vendor_msft_policy_config_defender_disablecatchupfullscan_0"
                                    settingValueTemplateReference = @{
                                        settingValueTemplateId = "1b26092f-48c4-447b-99d4-e9c501542f1c"
                                        useTemplateDefault = $false
                                    }
                                    children = @()
                                }
                            }
                        },
                        @{
                            id = "11"
                            settingInstance = @{
                                "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                settingDefinitionId = "device_vendor_msft_policy_config_defender_disablecatchupquickscan"
                                settingInstanceTemplateReference = @{
                                    settingInstanceTemplateId = "dabf6781-9d5d-42da-822a-d4327aa2bdd1"
                                }
                                choiceSettingValue = @{
                                    value = "device_vendor_msft_policy_config_defender_disablecatchupquickscan_0"
                                    settingValueTemplateReference = @{
                                        settingValueTemplateId = "d263ced7-0d23-4095-9326-99c8b3f5d35b"
                                        useTemplateDefault = $false
                                    }
                                    children = @()
                                }
                            }
                        },
                        @{
                            id = "12"
                            settingInstance = @{
                                "@odata.type" = "#microsoft.graph.deviceManagementConfigurationSimpleSettingInstance"
                                settingDefinitionId = "device_vendor_msft_policy_config_defender_avgcpuloadfactor"
                                settingInstanceTemplateReference = @{
                                    settingInstanceTemplateId = "816cc03e-8f96-4cba-b14f-2658d031a79a"
                                }
                                simpleSettingValue = @{
                                    "@odata.type" = "#microsoft.graph.deviceManagementConfigurationIntegerSettingValue"
                                    value = 50
                                    settingValueTemplateReference = @{
                                        settingValueTemplateId = "37195fb1-3743-4c8e-a0ce-b6fae6fa3acd"
                                        useTemplateDefault = $false
                                    }
                                }
                            }
                        },
                        @{
                            id = "13"
                            settingInstance = @{
                                "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                settingDefinitionId = "device_vendor_msft_policy_config_defender_allowuseruiaccess"
                                settingInstanceTemplateReference = @{
                                    settingInstanceTemplateId = "0170a900-b0bc-4ccc-b7ce-dda9be49189b"
                                }
                                choiceSettingValue = @{
                                    value = "device_vendor_msft_policy_config_defender_allowuseruiaccess_1"
                                    settingValueTemplateReference = @{
                                        settingValueTemplateId = "4b6c9739-4449-4006-8e5f-3049136470ea"
                                        useTemplateDefault = $false
                                    }
                                    children = @()
                                }
                            }
                        },
                        @{
                            id = "14"
                            settingInstance = @{
                                "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                settingDefinitionId = "device_vendor_msft_policy_config_defender_allowcloudprotection"
                                settingInstanceTemplateReference = @{
                                    settingInstanceTemplateId = "7da139f1-9b7e-407d-853a-c2e5037cdc70"
                                }
                                choiceSettingValue = @{
                                    value = "device_vendor_msft_policy_config_defender_allowcloudprotection_1"
                                    settingValueTemplateReference = @{
                                        settingValueTemplateId = "16fe8afd-67be-4c50-8619-d535451a500c"
                                        useTemplateDefault = $false
                                    }
                                    children = @()
                                }
                            }
                        },
                        @{
                            id = "15"
                            settingInstance = @{
                                "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                settingDefinitionId = "device_vendor_msft_policy_config_defender_realtimescandirection"
                                settingInstanceTemplateReference = @{
                                    settingInstanceTemplateId = "f5ff00a4-e5c7-44cf-a650-9c7619ff1561"
                                }
                                choiceSettingValue = @{
                                    value = "device_vendor_msft_policy_config_defender_realtimescandirection_0"
                                    settingValueTemplateReference = @{
                                        settingValueTemplateId = "6b4e3497-cfbb-4761-a152-de935bbf3f07"
                                        useTemplateDefault = $false
                                    }
                                    children = @()
                                }
                            }
                        },
                        @{
                            id = "16"
                            settingInstance = @{
                                "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                settingDefinitionId = "device_vendor_msft_policy_config_defender_allowbehaviormonitoring"
                                settingInstanceTemplateReference = @{
                                    settingInstanceTemplateId = "8eef615a-1aa0-46f4-a25a-12cbe65de5ab"
                                }
                                choiceSettingValue = @{
                                    value = "device_vendor_msft_policy_config_defender_allowbehaviormonitoring_1"
                                    settingValueTemplateReference = @{
                                        settingValueTemplateId = "905921da-95e2-4a10-9e30-fe5540002ce1"
                                        useTemplateDefault = $false
                                    }
                                    children = @()
                                }
                            }
                        },
                        @{
                            id = "17"
                            settingInstance = @{
                                "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                settingDefinitionId = "device_vendor_msft_policy_config_defender_allowioavprotection"
                                settingInstanceTemplateReference = @{
                                    settingInstanceTemplateId = "fa06231d-aed4-4601-b631-3a37e85b62a0"
                                }
                                choiceSettingValue = @{
                                    value = "device_vendor_msft_policy_config_defender_allowioavprotection_1"
                                    settingValueTemplateReference = @{
                                        settingValueTemplateId = "df4e6cbd-f7ff-41c8-88cd-fa25264a237e"
                                        useTemplateDefault = $false
                                    }
                                    children = @()
                                }
                            }
                        },
                        @{
                            id = "18"
                            settingInstance = @{
                                "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                settingDefinitionId = "device_vendor_msft_policy_config_defender_allowscriptscanning"
                                settingInstanceTemplateReference = @{
                                    settingInstanceTemplateId = "000cf176-949c-4c08-a5d4-90ed43718db7"
                                }
                                choiceSettingValue = @{
                                    value = "device_vendor_msft_policy_config_defender_allowscriptscanning_1"
                                    settingValueTemplateReference = @{
                                        settingValueTemplateId = "ab9e4320-c953-4067-ac9a-be2becd06b4a"
                                        useTemplateDefault = $false
                                    }
                                    children = @()
                                }
                            }
                        },
                        @{
                            id = "19"
                            settingInstance = @{
                                "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                settingDefinitionId = "device_vendor_msft_policy_config_defender_allowscanningnetworkfiles"
                                settingInstanceTemplateReference = @{
                                    settingInstanceTemplateId = "f8f28442-0a6b-4b52-b42c-d31d9687c1cf"
                                }
                                choiceSettingValue = @{
                                    value = "device_vendor_msft_policy_config_defender_allowscanningnetworkfiles_1"
                                    settingValueTemplateReference = @{
                                        settingValueTemplateId = "7b8c858c-a17d-4623-9e20-f34b851670ce"
                                        useTemplateDefault = $false
                                    }
                                    children = @()
                                }
                            }
                        },
                        @{
                            id = "20"
                            settingInstance = @{
                                "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                settingDefinitionId = "device_vendor_msft_policy_config_defender_allowemailscanning"
                                settingInstanceTemplateReference = @{
                                    settingInstanceTemplateId = "b0d9ee81-de6a-4750-86d7-9397961c9852"
                                }
                                choiceSettingValue = @{
                                    value = "device_vendor_msft_policy_config_defender_allowemailscanning_0"
                                    settingValueTemplateReference = @{
                                        settingValueTemplateId = "fdf107fd-e13b-4507-9d8f-db4d93476af9"
                                        useTemplateDefault = $false
                                    }
                                    children = @()
                                }
                            }
                        },
                        @{
                            id = "21"
                            settingInstance = @{
                                "@odata.type" = "#microsoft.graph.deviceManagementConfigurationSimpleSettingInstance"
                                settingDefinitionId = "device_vendor_msft_policy_config_defender_daystoretaincleanedmalware"
                                settingInstanceTemplateReference = @{
                                    settingInstanceTemplateId = "6f6d741c-1186-42e2-b2f2-8582febcfd60"
                                }
                                simpleSettingValue = @{
                                    "@odata.type" = "#microsoft.graph.deviceManagementConfigurationIntegerSettingValue"
                                    value = 0
                                    settingValueTemplateReference = @{
                                        settingValueTemplateId = "214b6feb-c9b2-4a17-af54-d51c805473e4"
                                        useTemplateDefault = $false
                                    }
                                }
                            }
                        },
                        @{
                            id = "22"
                            settingInstance = @{
                                "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                settingDefinitionId = "device_vendor_msft_policy_config_defender_submitsamplesconsent"
                                settingInstanceTemplateReference = @{
                                    settingInstanceTemplateId = "bc47ce7d-a251-4cae-a8a2-6e8384904ab7"
                                }
                                choiceSettingValue = @{
                                    value = "device_vendor_msft_policy_config_defender_submitsamplesconsent_1"
                                    settingValueTemplateReference = @{
                                        settingValueTemplateId = "826ed4b6-e04f-4975-9d23-6f0904b0d87e"
                                        useTemplateDefault = $false
                                    }
                                    children = @()
                                }
                            }
                        },
                        @{
                            id = "23"
                            settingInstance = @{
                                "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                settingDefinitionId = "device_vendor_msft_policy_config_defender_allowfullscanonmappednetworkdrives"
                                settingInstanceTemplateReference = @{
                                    settingInstanceTemplateId = "dac47505-f072-48d6-9f23-8d93262d58ed"
                                }
                                choiceSettingValue = @{
                                    value = "device_vendor_msft_policy_config_defender_allowfullscanonmappednetworkdrives_0"
                                    settingValueTemplateReference = @{
                                        settingValueTemplateId = "3e920b10-3773-4ac5-957e-e5573aec6d04"
                                        useTemplateDefault = $false
                                    }
                                    children = @()
                                }
                            }
                        },
                        @{
                            id = "24"
                            settingInstance = @{
                                "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                settingDefinitionId = "device_vendor_msft_policy_config_defender_allowfullscanremovabledrivescanning"
                                settingInstanceTemplateReference = @{
                                    settingInstanceTemplateId = "fb36e70b-5bc9-488a-a949-8ea3ac1634d5"
                                }
                                choiceSettingValue = @{
                                    value = "device_vendor_msft_policy_config_defender_allowfullscanremovabledrivescanning_1"
                                    settingValueTemplateReference = @{
                                        settingValueTemplateId = "366c5727-629b-4a81-b50b-52f90282fa2c"
                                        useTemplateDefault = $false
                                    }
                                    children = @()
                                }
                            }
                        },
                        @{
                            id = "25"
                            settingInstance = @{
                                "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                settingDefinitionId = "device_vendor_msft_policy_config_defender_checkforsignaturesbeforerunningscan"
                                settingInstanceTemplateReference = @{
                                    settingInstanceTemplateId = "4fea56e3-7bb6-4ad3-88c6-e364dd2f97b9"
                                }
                                choiceSettingValue = @{
                                    value = "device_vendor_msft_policy_config_defender_checkforsignaturesbeforerunningscan_1"
                                    settingValueTemplateReference = @{
                                        settingValueTemplateId = "010779d1-edd4-441d-8034-89ad57a863fe"
                                        useTemplateDefault = $false
                                    }
                                    children = @()
                                }
                            }
                        },
                        @{
                            id = "26"
                            settingInstance = @{
                                "@odata.type" = "#microsoft.graph.deviceManagementConfigurationSimpleSettingInstance"
                                settingDefinitionId = "device_vendor_msft_policy_config_defender_signatureupdateinterval"
                                settingInstanceTemplateReference = @{
                                    settingInstanceTemplateId = "89879f27-6b7d-44d4-a08e-0a0de3e9663d"
                                }
                                simpleSettingValue = @{
                                    "@odata.type" = "#microsoft.graph.deviceManagementConfigurationIntegerSettingValue"
                                    value = 4
                                    settingValueTemplateReference = @{
                                        settingValueTemplateId = "0af6bbed-a74a-4d08-8587-b16b10b774cb"
                                        useTemplateDefault = $false
                                    }
                                }
                            }
                        }
                    )
                }
                
                $result = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies" -Body $body
                Write-LogMessage -Message "Created comprehensive Defender Antivirus policy with 27 settings" -Type Success
                $policies += $result
            }
            catch {
                Write-LogMessage -Message "Failed to create Defender Antivirus policy - $($_.Exception.Message)" -Type Error
            }
        }
        
        # ===================================================================
        # POLICY 3: FIREWALL POLICY (Windows Firewall)
        # ===================================================================
        Write-LogMessage -Message "Creating Windows Firewall policy with template..." -Type Info
        
        $policyName = "Firewall Windows default policy"
        if (Test-PolicyExists -PolicyName $policyName) {
            Write-LogMessage -Message "Policy '$policyName' already exists, skipping creation" -Type Warning
            $policies += @{ name = $policyName; id = "existing" }
        }
        else {
            try {
                $body = @{
                    name = $policyName
                    description = "Default policy sets settings for all endpoints that are not governed by any other policy, ensuring that all your clients are managed as soon as MDE is deployed. The default policy is based on a set of pre-configured recommended settings and can be adjusted by user with admin priviledges."
                    platforms = "windows10"
                    technologies = "mdm,microsoftSense"
                    templateReference = @{
                        templateId = "6078910e-d808-4a9f-a51d-1b8a7bacb7c0_1"
                        templateFamily = "endpointSecurityFirewall"
                        templateDisplayName = "Windows Firewall"
                        templateDisplayVersion = "Version 1"
                    }
                    settings = @(
                        @{
                            id = "0"
                            settingInstance = @{
                                "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                settingDefinitionId = "vendor_msft_firewall_mdmstore_domainprofile_enablefirewall"
                                settingInstanceTemplateReference = @{
                                    settingInstanceTemplateId = "7714c373-a19a-4b64-ba6d-2e9db04a7684"
                                }
                                choiceSettingValue = @{
                                    value = "vendor_msft_firewall_mdmstore_domainprofile_enablefirewall_true"
                                    settingValueTemplateReference = @{
                                        settingValueTemplateId = "120c5dbe-0c88-46f0-b897-2c996d3e5277"
                                        useTemplateDefault = $false
                                    }
                                    children = @(
                                        @{
                                            "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                            settingDefinitionId = "vendor_msft_firewall_mdmstore_domainprofile_defaultinboundaction"
                                            choiceSettingValue = @{
                                                value = "vendor_msft_firewall_mdmstore_domainprofile_defaultinboundaction_1"
                                                children = @()
                                            }
                                        },
                                        @{
                                            "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                            settingDefinitionId = "vendor_msft_firewall_mdmstore_domainprofile_defaultoutboundaction"
                                            choiceSettingValue = @{
                                                value = "vendor_msft_firewall_mdmstore_domainprofile_defaultoutboundaction_0"
                                                children = @()
                                            }
                                        }
                                    )
                                }
                            }
                        },
                        @{
                            id = "1"
                            settingInstance = @{
                                "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                settingDefinitionId = "vendor_msft_firewall_mdmstore_privateprofile_enablefirewall"
                                settingInstanceTemplateReference = @{
                                    settingInstanceTemplateId = "1c14f914-69bb-49f8-af5b-e29173a6ee95"
                                }
                                choiceSettingValue = @{
                                    value = "vendor_msft_firewall_mdmstore_privateprofile_enablefirewall_true"
                                    settingValueTemplateReference = @{
                                        settingValueTemplateId = "9d55dfae-d55f-4f2a-af03-9a9524f61e76"
                                        useTemplateDefault = $false
                                    }
                                    children = @(
                                        @{
                                            "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                            settingDefinitionId = "vendor_msft_firewall_mdmstore_privateprofile_defaultinboundaction"
                                            choiceSettingValue = @{
                                                value = "vendor_msft_firewall_mdmstore_privateprofile_defaultinboundaction_1"
                                                children = @()
                                            }
                                        },
                                        @{
                                            "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                            settingDefinitionId = "vendor_msft_firewall_mdmstore_privateprofile_defaultoutboundaction"
                                            choiceSettingValue = @{
                                                value = "vendor_msft_firewall_mdmstore_privateprofile_defaultoutboundaction_0"
                                                children = @()
                                            }
                                        }
                                    )
                                }
                            }
                        },
                        @{
                            id = "2"
                            settingInstance = @{
                                "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                settingDefinitionId = "vendor_msft_firewall_mdmstore_publicprofile_enablefirewall"
                                settingInstanceTemplateReference = @{
                                    settingInstanceTemplateId = "e2714734-708e-4286-8ae9-d56821e306a3"
                                }
                                choiceSettingValue = @{
                                    value = "vendor_msft_firewall_mdmstore_publicprofile_enablefirewall_true"
                                    settingValueTemplateReference = @{
                                        settingValueTemplateId = "c38694c7-51a4-4a35-8f64-b10866a04776"
                                        useTemplateDefault = $false
                                    }
                                    children = @(
                                        @{
                                            "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                            settingDefinitionId = "vendor_msft_firewall_mdmstore_publicprofile_defaultinboundaction"
                                            choiceSettingValue = @{
                                                value = "vendor_msft_firewall_mdmstore_publicprofile_defaultinboundaction_1"
                                                children = @()
                                            }
                                        },
                                        @{
                                            "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                            settingDefinitionId = "vendor_msft_firewall_mdmstore_publicprofile_defaultoutboundaction"
                                            choiceSettingValue = @{
                                                value = "vendor_msft_firewall_mdmstore_publicprofile_defaultoutboundaction_0"
                                                children = @()
                                            }
                                        }
                                    )
                                }
                            }
                        }
                    )
                }
                
                $result = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies" -Body $body
                Write-LogMessage -Message "Created Windows Firewall policy with template references" -Type Success
                $policies += $result
            }
            catch {
                Write-LogMessage -Message "Failed to create Windows Firewall policy - $($_.Exception.Message)" -Type Error
            }
        }
        
        # ===================================================================
        # POLICY 4: TAMPER PROTECTION POLICY
        # ===================================================================
        Write-LogMessage -Message "Creating Tamper Protection policy..." -Type Info
        
        $policyName = "Tamper Protection"
        if (Test-PolicyExists -PolicyName $policyName) {
            Write-LogMessage -Message "Policy '$policyName' already exists, skipping creation" -Type Warning
            $policies += @{ name = $policyName; id = "existing" }
        }
        else {
            try {
                $body = @{
                    name = $policyName
                    description = "Windows Security tamper protection configuration"
                    platforms = "windows10"
                    technologies = "mdm,microsoftSense"
                    templateReference = @{
                        templateId = "d948ff9b-99cb-4ee0-8012-1fbc09685377_1"
                        templateFamily = "endpointSecurityAntivirus"
                        templateDisplayName = "Windows Security Experience"
                        templateDisplayVersion = "Version 1"
                    }
                    settings = @(
                        @{
                            id = "0"
                            settingInstance = @{
                                "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                settingDefinitionId = "vendor_msft_defender_configuration_tamperprotection_options"
                                settingInstanceTemplateReference = @{
                                    settingInstanceTemplateId = "5655cab2-7e6b-4c49-9ce2-3865da05f7e6"
                                }
                                choiceSettingValue = @{
                                    value = "vendor_msft_defender_configuration_tamperprotection_options_0"
                                    settingValueTemplateReference = @{
                                        settingValueTemplateId = "fc365da9-2c1b-4f79-aa4b-dedca69e728f"
                                        useTemplateDefault = $false
                                    }
                                    children = @()
                                }
                            }
                        }
                    )
                }
                
                $result = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies" -Body $body
                Write-LogMessage -Message "Created Tamper Protection policy" -Type Success
                $policies += $result
            }
            catch {
                Write-LogMessage -Message "Failed to create Tamper Protection policy - $($_.Exception.Message)" -Type Error
            }
        }
        
        # ===================================================================
        # POLICY 5: BITLOCKER POLICY (Enable Bitlocker) - ALL 13 SETTINGS
        # ===================================================================
        Write-LogMessage -Message "Creating comprehensive BitLocker policy with 13 settings..." -Type Info
        
        $policyName = "Enable Bitlocker"
        if (Test-PolicyExists -PolicyName $policyName) {
            Write-LogMessage -Message "Policy '$policyName' already exists, skipping creation" -Type Warning
            $policies += @{ name = $policyName; id = "existing" }
        }
        else {
            try {
                $body = @{
                    name = $policyName
                    description = "Comprehensive BitLocker drive encryption configuration"
                    platforms = "windows10"
                    technologies = "mdm"
                    settings = @(
                        # Require Device Encryption
                        @{
                            id = "0"
                            settingInstance = @{
                                "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                settingDefinitionId = "device_vendor_msft_bitlocker_requiredeviceencryption"
                                choiceSettingValue = @{
                                    value = "device_vendor_msft_bitlocker_requiredeviceencryption_1"
                                    children = @()
                                }
                            }
                        },
                        # Allow warning for other disk encryption
                        @{
                            id = "1"
                            settingInstance = @{
                                "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                settingDefinitionId = "device_vendor_msft_bitlocker_allowwarningforotherdiskencryption"
                                choiceSettingValue = @{
                                    value = "device_vendor_msft_bitlocker_allowwarningforotherdiskencryption_0"
                                    children = @(
                                        @{
                                            "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                            settingDefinitionId = "device_vendor_msft_bitlocker_allowstandarduserencryption"
                                            choiceSettingValue = @{
                                                value = "device_vendor_msft_bitlocker_allowstandarduserencryption_1"
                                                children = @()
                                            }
                                        }
                                    )
                                }
                            }
                        },
                        # Configure recovery password rotation
                        @{
                            id = "2"
                            settingInstance = @{
                                "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                settingDefinitionId = "device_vendor_msft_bitlocker_configurerecoverypasswordrotation"
                                choiceSettingValue = @{
                                    value = "device_vendor_msft_bitlocker_configurerecoverypasswordrotation_0"
                                    children = @()
                                }
                            }
                        },
                        # Encryption method by drive type
                        @{
                            id = "3"
                            settingInstance = @{
                                "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                settingDefinitionId = "device_vendor_msft_bitlocker_encryptionmethodbydrivetype"
                                choiceSettingValue = @{
                                    value = "device_vendor_msft_bitlocker_encryptionmethodbydrivetype_1"
                                    children = @(
                                        @{
                                            "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                            settingDefinitionId = "device_vendor_msft_bitlocker_encryptionmethodbydrivetype_encryptionmethodwithxtsfdvdropdown_name"
                                            choiceSettingValue = @{
                                                value = "device_vendor_msft_bitlocker_encryptionmethodbydrivetype_encryptionmethodwithxtsfdvdropdown_name_7"
                                                children = @()
                                            }
                                        },
                                        @{
                                            "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                            settingDefinitionId = "device_vendor_msft_bitlocker_encryptionmethodbydrivetype_encryptionmethodwithxtsosdropdown_name"
                                            choiceSettingValue = @{
                                                value = "device_vendor_msft_bitlocker_encryptionmethodbydrivetype_encryptionmethodwithxtsosdropdown_name_7"
                                                children = @()
                                            }
                                        },
                                        @{
                                            "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                            settingDefinitionId = "device_vendor_msft_bitlocker_encryptionmethodbydrivetype_encryptionmethodwithxtsrdvdropdown_name"
                                            choiceSettingValue = @{
                                                value = "device_vendor_msft_bitlocker_encryptionmethodbydrivetype_encryptionmethodwithxtsrdvdropdown_name_4"
                                                children = @()
                                            }
                                        }
                                    )
                                }
                            }
                        },
                        # System drives encryption type
                        @{
                            id = "4"
                            settingInstance = @{
                                "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                settingDefinitionId = "device_vendor_msft_bitlocker_systemdrivesencryptiontype"
                                choiceSettingValue = @{
                                    value = "device_vendor_msft_bitlocker_systemdrivesencryptiontype_1"
                                    children = @(
                                        @{
                                            "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                            settingDefinitionId = "device_vendor_msft_bitlocker_systemdrivesencryptiontype_osencryptiontypedropdown_name"
                                            choiceSettingValue = @{
                                                value = "device_vendor_msft_bitlocker_systemdrivesencryptiontype_osencryptiontypedropdown_name_1"
                                                children = @()
                                            }
                                        }
                                    )
                                }
                            }
                        },
                        # System drives require startup authentication
                        @{
                            id = "5"
                            settingInstance = @{
                                "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                settingDefinitionId = "device_vendor_msft_bitlocker_systemdrivesrequirestartupauthentication"
                                choiceSettingValue = @{
                                    value = "device_vendor_msft_bitlocker_systemdrivesrequirestartupauthentication_1"
                                    children = @(
                                        @{
                                            "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                            settingDefinitionId = "device_vendor_msft_bitlocker_systemdrivesrequirestartupauthentication_configurenontpmstartupkeyusage_name"
                                            choiceSettingValue = @{
                                                value = "device_vendor_msft_bitlocker_systemdrivesrequirestartupauthentication_configurenontpmstartupkeyusage_name_0"
                                                children = @()
                                            }
                                        },
                                        @{
                                            "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                            settingDefinitionId = "device_vendor_msft_bitlocker_systemdrivesrequirestartupauthentication_configuretpmpinkeyusagedropdown_name"
                                            choiceSettingValue = @{
                                                value = "device_vendor_msft_bitlocker_systemdrivesrequirestartupauthentication_configuretpmpinkeyusagedropdown_name_0"
                                                children = @()
                                            }
                                        },
                                        @{
                                            "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                            settingDefinitionId = "device_vendor_msft_bitlocker_systemdrivesrequirestartupauthentication_configuretpmstartupkeyusagedropdown_name"
                                            choiceSettingValue = @{
                                                value = "device_vendor_msft_bitlocker_systemdrivesrequirestartupauthentication_configuretpmstartupkeyusagedropdown_name_0"
                                                children = @()
                                            }
                                        },
                                        @{
                                            "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                            settingDefinitionId = "device_vendor_msft_bitlocker_systemdrivesrequirestartupauthentication_configurepinusagedropdown_name"
                                            choiceSettingValue = @{
                                                value = "device_vendor_msft_bitlocker_systemdrivesrequirestartupauthentication_configurepinusagedropdown_name_0"
                                                children = @()
                                            }
                                        },
                                        @{
                                            "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                            settingDefinitionId = "device_vendor_msft_bitlocker_systemdrivesrequirestartupauthentication_configuretpmusagedropdown_name"
                                            choiceSettingValue = @{
                                                value = "device_vendor_msft_bitlocker_systemdrivesrequirestartupauthentication_configuretpmusagedropdown_name_2"
                                                children = @()
                                            }
                                        }
                                    )
                                }
                            }
                        },
                        # System drives minimum PIN length
                        @{
                            id = "6"
                            settingInstance = @{
                                "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                settingDefinitionId = "device_vendor_msft_bitlocker_systemdrivesminimumpinlength"
                                choiceSettingValue = @{
                                    value = "device_vendor_msft_bitlocker_systemdrivesminimumpinlength_0"
                                    children = @()
                                }
                            }
                        },
                        # System drives enhanced PIN
                        @{
                            id = "7"
                            settingInstance = @{
                                "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                settingDefinitionId = "device_vendor_msft_bitlocker_systemdrivesenhancedpin"
                                choiceSettingValue = @{
                                    value = "device_vendor_msft_bitlocker_systemdrivesenhancedpin_0"
                                    children = @()
                                }
                            }
                        },
                        # System drives recovery options
                        @{
                            id = "8"
                            settingInstance = @{
                                "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                settingDefinitionId = "device_vendor_msft_bitlocker_systemdrivesrecoveryoptions"
                                choiceSettingValue = @{
                                    value = "device_vendor_msft_bitlocker_systemdrivesrecoveryoptions_1"
                                    children = @(
                                        @{
                                            "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                            settingDefinitionId = "device_vendor_msft_bitlocker_systemdrivesrecoveryoptions_osrecoverykeyusagedropdown_name"
                                            choiceSettingValue = @{
                                                value = "device_vendor_msft_bitlocker_systemdrivesrecoveryoptions_osrecoverykeyusagedropdown_name_2"
                                                children = @()
                                            }
                                        },
                                        @{
                                            "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                            settingDefinitionId = "device_vendor_msft_bitlocker_systemdrivesrecoveryoptions_osrecoverypasswordusagedropdown_name"
                                            choiceSettingValue = @{
                                                value = "device_vendor_msft_bitlocker_systemdrivesrecoveryoptions_osrecoverypasswordusagedropdown_name_2"
                                                children = @()
                                            }
                                        },
                                        @{
                                            "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                            settingDefinitionId = "device_vendor_msft_bitlocker_systemdrivesrecoveryoptions_osallowdra_name"
                                            choiceSettingValue = @{
                                                value = "device_vendor_msft_bitlocker_systemdrivesrecoveryoptions_osallowdra_name_0"
                                                children = @()
                                            }
                                        },
                                        @{
                                            "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                            settingDefinitionId = "device_vendor_msft_bitlocker_systemdrivesrecoveryoptions_osactivedirectorybackupdropdown_name"
                                            choiceSettingValue = @{
                                                value = "device_vendor_msft_bitlocker_systemdrivesrecoveryoptions_osactivedirectorybackupdropdown_name_1"
                                                children = @()
                                            }
                                        },
                                        @{
                                            "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                            settingDefinitionId = "device_vendor_msft_bitlocker_systemdrivesrecoveryoptions_osrequireactivedirectorybackup_name"
                                            choiceSettingValue = @{
                                                value = "device_vendor_msft_bitlocker_systemdrivesrecoveryoptions_osrequireactivedirectorybackup_name_0"
                                                children = @()
                                            }
                                        },
                                        @{
                                            "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                            settingDefinitionId = "device_vendor_msft_bitlocker_systemdrivesrecoveryoptions_oshiderecoverypage_name"
                                            choiceSettingValue = @{
                                                value = "device_vendor_msft_bitlocker_systemdrivesrecoveryoptions_oshiderecoverypage_name_1"
                                                children = @()
                                            }
                                        },
                                        @{
                                            "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                            settingDefinitionId = "device_vendor_msft_bitlocker_systemdrivesrecoveryoptions_osactivedirectorybackup_name"
                                            choiceSettingValue = @{
                                                value = "device_vendor_msft_bitlocker_systemdrivesrecoveryoptions_osactivedirectorybackup_name_1"
                                                children = @()
                                            }
                                        }
                                    )
                                }
                            }
                        },
                        # Fixed drives encryption type
                        @{
                            id = "9"
                            settingInstance = @{
                                "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                settingDefinitionId = "device_vendor_msft_bitlocker_fixeddrivesencryptiontype"
                                choiceSettingValue = @{
                                    value = "device_vendor_msft_bitlocker_fixeddrivesencryptiontype_1"
                                    children = @(
                                        @{
                                            "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                            settingDefinitionId = "device_vendor_msft_bitlocker_fixeddrivesencryptiontype_fdvencryptiontypedropdown_name"
                                            choiceSettingValue = @{
                                                value = "device_vendor_msft_bitlocker_fixeddrivesencryptiontype_fdvencryptiontypedropdown_name_1"
                                                children = @()
                                            }
                                        }
                                    )
                                }
                            }
                        },
                        # Fixed drives recovery options
                        @{
                            id = "10"
                            settingInstance = @{
                                "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                settingDefinitionId = "device_vendor_msft_bitlocker_fixeddrivesrecoveryoptions"
                                choiceSettingValue = @{
                                    value = "device_vendor_msft_bitlocker_fixeddrivesrecoveryoptions_1"
                                    children = @(
                                        @{
                                            "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                            settingDefinitionId = "device_vendor_msft_bitlocker_fixeddrivesrecoveryoptions_fdvrecoverykeyusagedropdown_name"
                                            choiceSettingValue = @{
                                                value = "device_vendor_msft_bitlocker_fixeddrivesrecoveryoptions_fdvrecoverykeyusagedropdown_name_2"
                                                children = @()
                                            }
                                        },
                                        @{
                                            "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                            settingDefinitionId = "device_vendor_msft_bitlocker_fixeddrivesrecoveryoptions_fdvrecoverypasswordusagedropdown_name"
                                            choiceSettingValue = @{
                                                value = "device_vendor_msft_bitlocker_fixeddrivesrecoveryoptions_fdvrecoverypasswordusagedropdown_name_2"
                                                children = @()
                                            }
                                        },
                                        @{
                                            "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                            settingDefinitionId = "device_vendor_msft_bitlocker_fixeddrivesrecoveryoptions_fdvallowdra_name"
                                            choiceSettingValue = @{
                                                value = "device_vendor_msft_bitlocker_fixeddrivesrecoveryoptions_fdvallowdra_name_0"
                                                children = @()
                                            }
                                        },
                                        @{
                                            "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                            settingDefinitionId = "device_vendor_msft_bitlocker_fixeddrivesrecoveryoptions_fdvactivedirectorybackupdropdown_name"
                                            choiceSettingValue = @{
                                                value = "device_vendor_msft_bitlocker_fixeddrivesrecoveryoptions_fdvactivedirectorybackupdropdown_name_1"
                                                children = @()
                                            }
                                        },
                                        @{
                                            "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                            settingDefinitionId = "device_vendor_msft_bitlocker_fixeddrivesrecoveryoptions_fdvrequireactivedirectorybackup_name"
                                            choiceSettingValue = @{
                                                value = "device_vendor_msft_bitlocker_fixeddrivesrecoveryoptions_fdvrequireactivedirectorybackup_name_1"
                                                children = @()
                                            }
                                        },
                                        @{
                                            "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                            settingDefinitionId = "device_vendor_msft_bitlocker_fixeddrivesrecoveryoptions_fdvhiderecoverypage_name"
                                            choiceSettingValue = @{
                                                value = "device_vendor_msft_bitlocker_fixeddrivesrecoveryoptions_fdvhiderecoverypage_name_1"
                                                children = @()
                                            }
                                        },
                                        @{
                                            "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                            settingDefinitionId = "device_vendor_msft_bitlocker_fixeddrivesrecoveryoptions_fdvactivedirectorybackup_name"
                                            choiceSettingValue = @{
                                                value = "device_vendor_msft_bitlocker_fixeddrivesrecoveryoptions_fdvactivedirectorybackup_name_1"
                                                children = @()
                                            }
                                        }
                                    )
                                }
                            }
                        },
                        # Removable drives configure BDE
                        @{
                            id = "11"
                            settingInstance = @{
                                "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                settingDefinitionId = "device_vendor_msft_bitlocker_removabledrivesconfigurebde"
                                choiceSettingValue = @{
                                    value = "device_vendor_msft_bitlocker_removabledrivesconfigurebde_1"
                                    children = @(
                                        @{
                                            "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                            settingDefinitionId = "device_vendor_msft_bitlocker_removabledrivesconfigurebde_rdvallowbde_name"
                                            choiceSettingValue = @{
                                                value = "device_vendor_msft_bitlocker_removabledrivesconfigurebde_rdvallowbde_name_1"
                                                children = @(
                                                    @{
                                                        "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                                        settingDefinitionId = "device_vendor_msft_bitlocker_removabledrivesencryptiontype"
                                                        choiceSettingValue = @{
                                                            value = "device_vendor_msft_bitlocker_removabledrivesencryptiontype_1"
                                                            children = @(
                                                                @{
                                                                    "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                                                    settingDefinitionId = "device_vendor_msft_bitlocker_removabledrivesencryptiontype_rdvencryptiontypedropdown_name"
                                                                    choiceSettingValue = @{
                                                                        value = "device_vendor_msft_bitlocker_removabledrivesencryptiontype_rdvencryptiontypedropdown_name_1"
                                                                        children = @()
                                                                    }
                                                                }
                                                            )
                                                        }
                                                    }
                                                )
                                            }
                                        },
                                        @{
                                            "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                            settingDefinitionId = "device_vendor_msft_bitlocker_removabledrivesconfigurebde_rdvdisablebde_name"
                                            choiceSettingValue = @{
                                                value = "device_vendor_msft_bitlocker_removabledrivesconfigurebde_rdvdisablebde_name_1"
                                                children = @()
                                            }
                                        }
                                    )
                                }
                            }
                        },
                        # Removable drives require encryption
                        @{
                            id = "12"
                            settingInstance = @{
                                "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                settingDefinitionId = "device_vendor_msft_bitlocker_removabledrivesrequireencryption"
                                choiceSettingValue = @{
                                    value = "device_vendor_msft_bitlocker_removabledrivesrequireencryption_1"
                                    children = @(
                                        @{
                                            "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                            settingDefinitionId = "device_vendor_msft_bitlocker_removabledrivesrequireencryption_rdvcrossorg"
                                            choiceSettingValue = @{
                                                value = "device_vendor_msft_bitlocker_removabledrivesrequireencryption_rdvcrossorg_1"
                                                children = @()
                                            }
                                        }
                                    )
                                }
                            }
                        }
                    )
                }
                
                $result = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies" -Body $body
                Write-LogMessage -Message "Created comprehensive BitLocker policy with 13 settings" -Type Success
                $policies += $result
            }
            catch {
                Write-LogMessage -Message "Failed to create BitLocker policy - $($_.Exception.Message)" -Type Error
            }
        }
        
        # ===================================================================
        # POLICY 6: LAPS POLICY (Local Admin Password Solution)
        # ===================================================================
        Write-LogMessage -Message "Creating LAPS policy with domain-based admin name..." -Type Info
        
        $policyName = "LAPS"
        if (Test-PolicyExists -PolicyName $policyName) {
            Write-LogMessage -Message "Policy '$policyName' already exists, skipping creation" -Type Warning
            $policies += @{ name = $policyName; id = "existing" }
        }
        else {
            try {
                # Get domain initials from tenant
                $adminAccountName = "localadmin"
                if ($script:TenantState -and $script:TenantState.TenantName) {
                    $tenantName = $script:TenantState.TenantName
                    # Extract initials from tenant name (e.g., "Penneys" -> "P", "BITS Corp" -> "BC")
                    $initials = ($tenantName -split '\s+' | ForEach-Object { $_.Substring(0,1).ToUpper() }) -join ''
                    $adminAccountName = "$($initials)Local"
                }
                
                Write-LogMessage -Message "Setting LAPS admin account name to: $adminAccountName" -Type Info
                
                $body = @{
                    name = $policyName
                    description = "Local Admin Password Solution"
                    platforms = "windows10"
                    technologies = "mdm"
                    templateReference = @{
                        templateId = "adc46e5a-f4aa-4ff6-aeff-4f27bc525796_1"
                        templateFamily = "endpointSecurityAccountProtection"
                        templateDisplayName = "Local admin password solution (Windows LAPS)"
                        templateDisplayVersion = "Version 1"
                    }
                    settings = @(
                        @{
                            id = "0"
                            settingInstance = @{
                                "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                settingDefinitionId = "device_vendor_msft_laps_policies_backupdirectory"
                                settingInstanceTemplateReference = @{
                                    settingInstanceTemplateId = "a3270f64-e493-499d-8900-90290f61ed8a"
                                }
                                choiceSettingValue = @{
                                    value = "device_vendor_msft_laps_policies_backupdirectory_1"
                                    settingValueTemplateReference = @{
                                        settingValueTemplateId = "4d90f03d-e14c-43c4-86da-681da96a2f92"
                                        useTemplateDefault = $false
                                    }
                                    children = @(
                                        @{
                                            "@odata.type" = "#microsoft.graph.deviceManagementConfigurationSimpleSettingInstance"
                                            settingDefinitionId = "device_vendor_msft_laps_policies_passwordagedays_aad"
                                            simpleSettingValue = @{
                                                "@odata.type" = "#microsoft.graph.deviceManagementConfigurationIntegerSettingValue"
                                                value = 7
                                            }
                                        }
                                    )
                                }
                            }
                        },
                        @{
                            id = "1"
                            settingInstance = @{
                                "@odata.type" = "#microsoft.graph.deviceManagementConfigurationSimpleSettingInstance"
                                settingDefinitionId = "device_vendor_msft_laps_policies_administratoraccountname"
                                settingInstanceTemplateReference = @{
                                    settingInstanceTemplateId = "d3d7d492-0019-4f56-96f8-1967f7deabeb"
                                }
                                simpleSettingValue = @{
                                    "@odata.type" = "#microsoft.graph.deviceManagementConfigurationStringSettingValue"
                                    value = $adminAccountName
                                    settingValueTemplateReference = @{
                                        settingValueTemplateId = "992c7fce-f9e4-46ab-ac11-e167398859ea"
                                        useTemplateDefault = $false
                                    }
                                }
                            }
                        },
                        @{
                            id = "2"
                            settingInstance = @{
                                "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                settingDefinitionId = "device_vendor_msft_laps_policies_passwordcomplexity"
                                settingInstanceTemplateReference = @{
                                    settingInstanceTemplateId = "8a7459e8-1d1c-458a-8906-7b27d216de52"
                                }
                                choiceSettingValue = @{
                                    value = "device_vendor_msft_laps_policies_passwordcomplexity_3"
                                    settingValueTemplateReference = @{
                                        settingValueTemplateId = "aa883ab5-625e-4e3b-b830-a37a4bb8ce01"
                                        useTemplateDefault = $false
                                    }
                                    children = @()
                                }
                            }
                        },
                        @{
                            id = "3"
                            settingInstance = @{
                                "@odata.type" = "#microsoft.graph.deviceManagementConfigurationSimpleSettingInstance"
                                settingDefinitionId = "device_vendor_msft_laps_policies_passwordlength"
                                settingInstanceTemplateReference = @{
                                    settingInstanceTemplateId = "da7a1dbd-caf7-4341-ab63-ece6f994ff02"
                                }
                                simpleSettingValue = @{
                                    "@odata.type" = "#microsoft.graph.deviceManagementConfigurationIntegerSettingValue"
                                    value = 20
                                    settingValueTemplateReference = @{
                                        settingValueTemplateId = "d08f1266-5345-4f53-8ae1-4c20e6cb5ec9"
                                        useTemplateDefault = $false
                                    }
                                }
                            }
                        },
                        @{
                            id = "4"
                            settingInstance = @{
                                "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                settingDefinitionId = "device_vendor_msft_laps_policies_postauthenticationactions"
                                settingInstanceTemplateReference = @{
                                    settingInstanceTemplateId = "d9282eb1-d187-42ae-b366-7081f32dcfff"
                                }
                                choiceSettingValue = @{
                                    value = "device_vendor_msft_laps_policies_postauthenticationactions_3"
                                    settingValueTemplateReference = @{
                                        settingValueTemplateId = "68ff4f78-baa8-4b32-bf3d-5ad5566d8142"
                                        useTemplateDefault = $false
                                    }
                                    children = @()
                                }
                            }
                        },
                        @{
                            id = "5"
                            settingInstance = @{
                                "@odata.type" = "#microsoft.graph.deviceManagementConfigurationSimpleSettingInstance"
                                settingDefinitionId = "device_vendor_msft_laps_policies_postauthenticationresetdelay"
                                settingInstanceTemplateReference = @{
                                    settingInstanceTemplateId = "a9e21166-4055-4042-9372-efaf3ef41868"
                                }
                                simpleSettingValue = @{
                                    "@odata.type" = "#microsoft.graph.deviceManagementConfigurationIntegerSettingValue"
                                    value = 1
                                    settingValueTemplateReference = @{
                                        settingValueTemplateId = "0deb6aee-8dac-40c4-a9dd-c3718e5c1d52"
                                        useTemplateDefault = $false
                                    }
                                }
                            }
                        }
                    )
                }
                
                $result = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies" -Body $body
                Write-LogMessage -Message "Created LAPS policy with admin account: $adminAccountName" -Type Success
                $policies += $result
            }
            catch {
                Write-LogMessage -Message "Failed to create LAPS policy - $($_.Exception.Message)" -Type Error
            }
        }
        
        # ===================================================================
        # POLICY 7: ONEDRIVE CONFIGURATION - ALL 7 SETTINGS
        # ===================================================================
        Write-LogMessage -Message "Creating comprehensive OneDrive policy..." -Type Info
        
        $policyName = "OneDrive Configuration"
        if (Test-PolicyExists -PolicyName $policyName) {
            Write-LogMessage -Message "Policy '$policyName' already exists, skipping creation" -Type Warning
            $policies += @{ name = $policyName; id = "existing" }
        }
        else {
            try {
                $body = @{
                    name = $policyName
                    description = "OneDrive for Business configuration with Known Folder Move"
                    platforms = "windows10"
                    technologies = "mdm"
                    settings = @(
                        # Disable pause on metered networks
                        @{
                            id = "0"
                            settingInstance = @{
                                "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                settingDefinitionId = "user_vendor_msft_policy_config_onedrivengscv2~policy~onedrivengsc_disablepauseonmeterednetwork"
                                choiceSettingValue = @{
                                    value = "user_vendor_msft_policy_config_onedrivengscv2~policy~onedrivengsc_disablepauseonmeterednetwork_1"
                                    children = @()
                                }
                            }
                        },
                        # Block opt-out from KFM
                        @{
                            id = "1"
                            settingInstance = @{
                                "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                settingDefinitionId = "device_vendor_msft_policy_config_onedrivengscv2~policy~onedrivengsc_kfmblockoptout"
                                choiceSettingValue = @{
                                    value = "device_vendor_msft_policy_config_onedrivengscv2~policy~onedrivengsc_kfmblockoptout_1"
                                    children = @()
                                }
                            }
                        },
                        # Disable personal sync
                        @{
                            id = "2"
                            settingInstance = @{
                                "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                settingDefinitionId = "user_vendor_msft_policy_config_onedrivengscv2~policy~onedrivengsc_disablepersonalsync"
                                choiceSettingValue = @{
                                    value = "user_vendor_msft_policy_config_onedrivengscv2~policy~onedrivengsc_disablepersonalsync_1"
                                    children = @()
                                }
                            }
                        },
                        # Force local mass delete detection
                        @{
                            id = "3"
                            settingInstance = @{
                                "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                settingDefinitionId = "device_vendor_msft_policy_config_onedrivengscv2~policy~onedrivengsc_forcedlocalmassdeletedetection"
                                choiceSettingValue = @{
                                    value = "device_vendor_msft_policy_config_onedrivengscv2~policy~onedrivengsc_forcedlocalmassdeletedetection_1"
                                    children = @()
                                }
                            }
                        },
                        # KFM Opt-in with Desktop, Documents, Pictures
                        @{
                            id = "4"
                            settingInstance = @{
                                "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                settingDefinitionId = "device_vendor_msft_policy_config_onedrivengscv2.updates~policy~onedrivengsc_kfmoptinnowizard"
                                choiceSettingValue = @{
                                    value = "device_vendor_msft_policy_config_onedrivengscv2.updates~policy~onedrivengsc_kfmoptinnowizard_1"
                                    children = @(
                                        @{
                                            "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                            settingDefinitionId = "device_vendor_msft_policy_config_onedrivengscv2.updates~policy~onedrivengsc_kfmoptinnowizard_kfmoptinnowizard_desktop_checkbox"
                                            choiceSettingValue = @{
                                                value = "device_vendor_msft_policy_config_onedrivengscv2.updates~policy~onedrivengsc_kfmoptinnowizard_kfmoptinnowizard_desktop_checkbox_1"
                                                children = @()
                                            }
                                        },
                                        @{
                                            "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                            settingDefinitionId = "device_vendor_msft_policy_config_onedrivengscv2.updates~policy~onedrivengsc_kfmoptinnowizard_kfmoptinnowizard_documents_checkbox"
                                            choiceSettingValue = @{
                                                value = "device_vendor_msft_policy_config_onedrivengscv2.updates~policy~onedrivengsc_kfmoptinnowizard_kfmoptinnowizard_documents_checkbox_1"
                                                children = @()
                                            }
                                        },
                                        @{
                                            "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                            settingDefinitionId = "device_vendor_msft_policy_config_onedrivengscv2.updates~policy~onedrivengsc_kfmoptinnowizard_kfmoptinnowizard_pictures_checkbox"
                                            choiceSettingValue = @{
                                                value = "device_vendor_msft_policy_config_onedrivengscv2.updates~policy~onedrivengsc_kfmoptinnowizard_kfmoptinnowizard_pictures_checkbox_1"
                                                children = @()
                                            }
                                        },
                                        @{
                                            "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                            settingDefinitionId = "device_vendor_msft_policy_config_onedrivengscv2.updates~policy~onedrivengsc_kfmoptinnowizard_kfmoptinnowizard_dropdown"
                                            choiceSettingValue = @{
                                                value = "device_vendor_msft_policy_config_onedrivengscv2.updates~policy~onedrivengsc_kfmoptinnowizard_kfmoptinnowizard_dropdown_0"
                                                children = @()
                                            }
                                        },
                                        @{
                                            "@odata.type" = "#microsoft.graph.deviceManagementConfigurationSimpleSettingInstance"
                                            settingDefinitionId = "device_vendor_msft_policy_config_onedrivengscv2.updates~policy~onedrivengsc_kfmoptinnowizard_kfmoptinnowizard_textbox"
                                            simpleSettingValue = @{
                                                "@odata.type" = "#microsoft.graph.deviceManagementConfigurationStringSettingValue"
                                                value = ""
                                            }
                                        }
                                    )
                                }
                            }
                        },
                        # Silent Account Config
                        @{
                            id = "5"
                            settingInstance = @{
                                "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                settingDefinitionId = "device_vendor_msft_policy_config_onedrivengscv2~policy~onedrivengsc_silentaccountconfig"
                                choiceSettingValue = @{
                                    value = "device_vendor_msft_policy_config_onedrivengscv2~policy~onedrivengsc_silentaccountconfig_1"
                                    children = @()
                                }
                            }
                        },
                        # Files on Demand
                        @{
                            id = "6"
                            settingInstance = @{
                                "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                settingDefinitionId = "device_vendor_msft_policy_config_onedrivengscv2~policy~onedrivengsc_filesondemandenabled"
                                choiceSettingValue = @{
                                    value = "device_vendor_msft_policy_config_onedrivengscv2~policy~onedrivengsc_filesondemandenabled_1"
                                    children = @()
                                }
                            }
                        }
                    )
                }
                
                $result = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies" -Body $body
                Write-LogMessage -Message "Created comprehensive OneDrive policy with 7 settings" -Type Success
                $policies += $result
            }
            catch {
                Write-LogMessage -Message "Failed to create OneDrive policy - $($_.Exception.Message)" -Type Error
            }
        }
        
        # ===================================================================
        # POLICY 8: POWER OPTIONS POLICY - ALL 6 SETTINGS
        # ===================================================================
        Write-LogMessage -Message "Creating comprehensive Power Options policy..." -Type Info
        
        $policyName = "Power Options Configuration"
        if (Test-PolicyExists -PolicyName $policyName) {
            Write-LogMessage -Message "Policy '$policyName' already exists, skipping creation" -Type Warning
            $policies += @{ name = $policyName; id = "existing" }
        }
        else {
            try {
                $body = @{
                    name = $policyName
                    description = "Comprehensive power management settings for devices"
                    platforms = "windows10"
                    technologies = "mdm"
                    settings = @(
                        # Allow Hibernate
                        @{
                            id = "0"
                            settingInstance = @{
                                "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                settingDefinitionId = "device_vendor_msft_policy_config_power_allowhibernate"
                                choiceSettingValue = @{
                                    value = "device_vendor_msft_policy_config_power_allowhibernate_1"
                                    children = @()
                                }
                            }
                        },
                        # Lid close action on battery (Sleep)
                        @{
                            id = "1"
                            settingInstance = @{
                                "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                settingDefinitionId = "device_vendor_msft_policy_config_power_selectlidcloseactiononbattery"
                                choiceSettingValue = @{
                                    value = "device_vendor_msft_policy_config_power_selectlidcloseactiononbattery_1"
                                    children = @()
                                }
                            }
                        },
                        # Lid close action plugged in (Sleep)
                        @{
                            id = "2"
                            settingInstance = @{
                                "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                settingDefinitionId = "device_vendor_msft_policy_config_power_selectlidcloseactionpluggedin"
                                choiceSettingValue = @{
                                    value = "device_vendor_msft_policy_config_power_selectlidcloseactionpluggedin_1"
                                    children = @()
                                }
                            }
                        },
                        # Power button action on battery (Do nothing)
                        @{
                            id = "3"
                            settingInstance = @{
                                "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                settingDefinitionId = "device_vendor_msft_policy_config_power_selectpowerbuttonactiononbattery"
                                choiceSettingValue = @{
                                    value = "device_vendor_msft_policy_config_power_selectpowerbuttonactiononbattery_0"
                                    children = @()
                                }
                            }
                        },
                        # Power button action plugged in (Do nothing)
                        @{
                            id = "4"
                            settingInstance = @{
                                "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                settingDefinitionId = "device_vendor_msft_policy_config_power_selectpowerbuttonactionpluggedin"
                                choiceSettingValue = @{
                                    value = "device_vendor_msft_policy_config_power_selectpowerbuttonactionpluggedin_0"
                                    children = @()
                                }
                            }
                        },
                        # Unattended sleep timeout plugged in (15 minutes = 900 seconds)
                        @{
                            id = "5"
                            settingInstance = @{
                                "@odata.type" = "#microsoft.graph.deviceManagementConfigurationSimpleSettingInstance"
                                settingDefinitionId = "device_vendor_msft_policy_config_power_unattendedsleeptimeoutpluggedin"
                                simpleSettingValue = @{
                                    "@odata.type" = "#microsoft.graph.deviceManagementConfigurationIntegerSettingValue"
                                    value = 900
                                }
                            }
                        }
                    )
                }
                
                $result = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies" -Body $body
                Write-LogMessage -Message "Created comprehensive Power Options policy" -Type Success
                $policies += $result
            }
            catch {
                Write-LogMessage -Message "Failed to create Power Options policy - $($_.Exception.Message)" -Type Error
            }
        }
        
        # ===================================================================
        # POLICY 9: ADMIN ACCOUNT POLICY - Enable and Rename Built-in Administrator
        # ===================================================================
        Write-LogMessage -Message "Creating Admin Account policy with rename..." -Type Info
        
        $policyName = "Enable Built-in Administrator Account"
        if (Test-PolicyExists -PolicyName $policyName) {
            Write-LogMessage -Message "Policy '$policyName' already exists, skipping creation" -Type Warning
            $policies += @{ name = $policyName; id = "existing" }
        }
        else {
            try {
                $body = @{
                    name = $policyName
                    description = "Enable and configure built-in administrator account for LAPS"
                    platforms = "windows10"
                    technologies = "mdm"
                    settings = @(
                        # Enable Administrator Account
                        @{
                            id = "0"
                            settingInstance = @{
                                "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                settingDefinitionId = "device_vendor_msft_policy_config_localpoliciessecurityoptions_accounts_enableadministratoraccountstatus"
                                choiceSettingValue = @{
                                    value = "device_vendor_msft_policy_config_localpoliciessecurityoptions_accounts_enableadministratoraccountstatus_1"
                                    children = @()
                                }
                            }
                        },
                        # Rename Administrator Account
                        @{
                            id = "1"
                            settingInstance = @{
                                "@odata.type" = "#microsoft.graph.deviceManagementConfigurationSimpleSettingInstance"
                                settingDefinitionId = "device_vendor_msft_policy_config_localpoliciessecurityoptions_accounts_renameadministratoraccount"
                                simpleSettingValue = @{
                                    "@odata.type" = "#microsoft.graph.deviceManagementConfigurationStringSettingValue"
                                    value = "localadmin"
                                }
                            }
                        }
                    )
                }
                
                $result = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies" -Body $body
                Write-LogMessage -Message "Created Admin Account policy with rename setting" -Type Success
                $policies += $result
            }
            catch {
                Write-LogMessage -Message "Failed to create Admin Account policy - $($_.Exception.Message)" -Type Error
            }
        }
        
        # ===================================================================
        # POLICY 10: UNENROLLMENT PREVENTION POLICY
        # ===================================================================
        Write-LogMessage -Message "Creating Device Unenrollment Prevention policy..." -Type Info
        
        $policyName = "Prevent Users From Unenrolling Devices"
        if (Test-PolicyExists -PolicyName $policyName) {
            Write-LogMessage -Message "Policy '$policyName' already exists, skipping creation" -Type Warning
            $policies += @{ name = $policyName; id = "existing" }
        }
        else {
            try {
                $body = @{
                    name = $policyName
                    description = "Prevent users from manually unenrolling devices from Intune"
                    platforms = "windows10"
                    technologies = "mdm"
                    settings = @(
                        @{
                            id = "0"
                            settingInstance = @{
                                "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                settingDefinitionId = "device_vendor_msft_policy_config_experience_allowmanualmdmunenrollment"
                                choiceSettingValue = @{
                                    value = "device_vendor_msft_policy_config_experience_allowmanualmdmunenrollment_0"
                                    children = @()
                                }
                            }
                        }
                    )
                }
                
                $result = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies" -Body $body
                Write-LogMessage -Message "Created Device Unenrollment Prevention policy" -Type Success
                $policies += $result
            }
            catch {
                Write-LogMessage -Message "Failed to create Unenrollment Prevention policy - $($_.Exception.Message)" -Type Error
            }
        }
        
        # ===================================================================
        # POLICY 11: EDGE POLICIES - Default Web Pages with SharePoint Homepage
        # ===================================================================
        Write-LogMessage -Message "Creating Edge policy with SharePoint homepage..." -Type Info
        
        $policyName = "Default Web Pages"
        if (Test-PolicyExists -PolicyName $policyName) {
            Write-LogMessage -Message "Policy '$policyName' already exists, skipping creation" -Type Warning
            $policies += @{ name = $policyName; id = "existing" }
        }
        else {
            try {
                # Get SharePoint root site URL from tenant
                $sharePointUrl = Get-SharePointRootSiteUrl
                
                $body = @{
                    name = $policyName
                    description = "Setting SharePoint home page as default start up page"
                    platforms = "windows10"
                    technologies = "mdm"
                    settings = @(
                        # Restore on startup
                        @{
    id = "0"
    settingInstance = @{
        "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
        settingDefinitionId = "device_vendor_msft_policy_config_microsoft_edgev77.3~policy~microsoft_edge~startup_restoreonstartup"
        choiceSettingValue = @{
            # PARENT: Keep as _1 (this enables the policy)
            value = "device_vendor_msft_policy_config_microsoft_edgev77.3~policy~microsoft_edge~startup_restoreonstartup_1"
            children = @(
                @{
                    "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                    settingDefinitionId = "device_vendor_msft_policy_config_microsoft_edgev77.3~policy~microsoft_edge~startup_restoreonstartup_restoreonstartup"
                    choiceSettingValue = @{
                        # CHILD: Change from _5 to _4 (this sets the behavior)
                        value = "device_vendor_msft_policy_config_microsoft_edgev77.3~policy~microsoft_edge~startup_restoreonstartup_restoreonstartup_4"
                        children = @()
                    }
                }
            )
        }
    }
},
                        # Home page location
                        @{
                            id = "1"
                            settingInstance = @{
                                "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                settingDefinitionId = "device_vendor_msft_policy_config_microsoft_edge~policy~microsoft_edge~startup_homepagelocation"
                                choiceSettingValue = @{
                                    value = "device_vendor_msft_policy_config_microsoft_edge~policy~microsoft_edge~startup_homepagelocation_1"
                                    children = @(
                                        @{
                                            "@odata.type" = "#microsoft.graph.deviceManagementConfigurationSimpleSettingInstance"
                                            settingDefinitionId = "device_vendor_msft_policy_config_microsoft_edge~policy~microsoft_edge~startup_homepagelocation_homepagelocation"
                                            simpleSettingValue = @{
                                                "@odata.type" = "#microsoft.graph.deviceManagementConfigurationStringSettingValue"
                                                value = $sharePointUrl
                                            }
                                        }
                                    )
                                }
                            }
                        },
                        # Restore on startup URLs
                        @{
                            id = "2"
                            settingInstance = @{
                                "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                settingDefinitionId = "device_vendor_msft_policy_config_microsoft_edge~policy~microsoft_edge~startup_restoreonstartupurls"
                                choiceSettingValue = @{
                                    value = "device_vendor_msft_policy_config_microsoft_edge~policy~microsoft_edge~startup_restoreonstartupurls_1"
                                    children = @(
                                        @{
                                            "@odata.type" = "#microsoft.graph.deviceManagementConfigurationSimpleSettingCollectionInstance"
                                            settingDefinitionId = "device_vendor_msft_policy_config_microsoft_edge~policy~microsoft_edge~startup_restoreonstartupurls_restoreonstartupurlsdesc"
                                            simpleSettingCollectionValue = @(
                                                @{
                                                    "@odata.type" = "#microsoft.graph.deviceManagementConfigurationStringSettingValue"
                                                    value = $sharePointUrl
                                                }
                                            )
                                        }
                                    )
                                }
                            }
                        }
                    )
                }
                
                $result = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies" -Body $body
                Write-LogMessage -Message "Created Edge policy with SharePoint home page: $sharePointUrl" -Type Success
                $policies += $result
            }
            catch {
                Write-LogMessage -Message "Failed to create Edge policies - $($_.Exception.Message)" -Type Error
            }
        }
        
        # ===================================================================
        # POLICY 12: EDGE UPDATE POLICY - All 4 Settings
        # ===================================================================
        Write-LogMessage -Message "Creating Edge Update policy..." -Type Info
        
        $policyName = "Edge Update Policy"
        if (Test-PolicyExists -PolicyName $policyName) {
            Write-LogMessage -Message "Policy '$policyName' already exists, skipping creation" -Type Warning
            $policies += @{ name = $policyName; id = "existing" }
        }
        else {
            try {
                $body = @{
                    name = $policyName
                    description = "Microsoft Edge update configuration"
                    platforms = "windows10"
                    technologies = "mdm"
                    settings = @(
                        # Target Channel
                        @{
                            id = "0"
                            settingInstance = @{
                                "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                settingDefinitionId = "device_vendor_msft_policy_config_updatev95~policy~cat_edgeupdate~cat_applications~cat_microsoftedge_pol_targetchannelmicrosoftedge"
                                choiceSettingValue = @{
                                    value = "device_vendor_msft_policy_config_updatev95~policy~cat_edgeupdate~cat_applications~cat_microsoftedge_pol_targetchannelmicrosoftedge_1"
                                    children = @(
                                        @{
                                            "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                            settingDefinitionId = "device_vendor_msft_policy_config_updatev95~policy~cat_edgeupdate~cat_applications~cat_microsoftedge_pol_targetchannelmicrosoftedge_part_targetchannel"
                                            choiceSettingValue = @{
                                                value = "device_vendor_msft_policy_config_updatev95~policy~cat_edgeupdate~cat_applications~cat_microsoftedge_pol_targetchannelmicrosoftedge_part_targetchannel_stable"
                                                children = @()
                                            }
                                        }
                                    )
                                }
                            }
                        },
                        # Update Policy Microsoft Edge
                        @{
                            id = "1"
                            settingInstance = @{
                                "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                settingDefinitionId = "device_vendor_msft_policy_config_update~policy~cat_google~cat_googleupdate~cat_applications~cat_microsoftedge_pol_updatepolicymicrosoftedge"
                                choiceSettingValue = @{
                                    value = "device_vendor_msft_policy_config_update~policy~cat_google~cat_googleupdate~cat_applications~cat_microsoftedge_pol_updatepolicymicrosoftedge_1"
                                    children = @(
                                        @{
                                            "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                            settingDefinitionId = "device_vendor_msft_policy_config_update~policy~cat_google~cat_googleupdate~cat_applications~cat_microsoftedge_pol_updatepolicymicrosoftedge_part_updatepolicy"
                                            choiceSettingValue = @{
                                                value = "device_vendor_msft_policy_config_update~policy~cat_google~cat_googleupdate~cat_applications~cat_microsoftedge_pol_updatepolicymicrosoftedge_part_updatepolicy_1"
                                                children = @()
                                            }
                                        }
                                    )
                                }
                            }
                        },
                        # Default Update Policy
                        @{
                            id = "2"
                            settingInstance = @{
                                "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                settingDefinitionId = "device_vendor_msft_policy_config_update~policy~cat_google~cat_googleupdate~cat_applications_pol_defaultupdatepolicy"
                                choiceSettingValue = @{
                                    value = "device_vendor_msft_policy_config_update~policy~cat_google~cat_googleupdate~cat_applications_pol_defaultupdatepolicy_1"
                                    children = @(
                                        @{
                                            "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                            settingDefinitionId = "device_vendor_msft_policy_config_update~policy~cat_google~cat_googleupdate~cat_applications_pol_defaultupdatepolicy_part_updatepolicy"
                                            choiceSettingValue = @{
                                                value = "device_vendor_msft_policy_config_update~policy~cat_google~cat_googleupdate~cat_applications_pol_defaultupdatepolicy_part_updatepolicy_1"
                                                children = @()
                                            }
                                        }
                                    )
                                }
                            }
                        },
                        # Auto Update Check Period
                        @{
                            id = "3"
                            settingInstance = @{
                                "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                settingDefinitionId = "device_vendor_msft_policy_config_update~policy~cat_google~cat_googleupdate~cat_preferences_pol_autoupdatecheckperiod"
                                choiceSettingValue = @{
                                    value = "device_vendor_msft_policy_config_update~policy~cat_google~cat_googleupdate~cat_preferences_pol_autoupdatecheckperiod_1"
                                    children = @(
                                        @{
                                            "@odata.type" = "#microsoft.graph.deviceManagementConfigurationSimpleSettingInstance"
                                            settingDefinitionId = "device_vendor_msft_policy_config_update~policy~cat_google~cat_googleupdate~cat_preferences_pol_autoupdatecheckperiod_part_autoupdatecheckperiod"
                                            simpleSettingValue = @{
                                                "@odata.type" = "#microsoft.graph.deviceManagementConfigurationIntegerSettingValue"
                                                value = 700
                                            }
                                        }
                                    )
                                }
                            }
                        }
                    )
                }
                
                $result = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies" -Body $body
                Write-LogMessage -Message "Created Edge Update policy" -Type Success
                $policies += $result
            }
            catch {
                Write-LogMessage -Message "Failed to create Edge Update policy - $($_.Exception.Message)" -Type Error
            }
        }
        
        # ===================================================================
        # POLICY 13: OFFICE POLICIES
        # ===================================================================
        Write-LogMessage -Message "Creating Office configuration policies..." -Type Info
        
        $policyName = "Office Updates Configuration"
        if (Test-PolicyExists -PolicyName $policyName) {
            Write-LogMessage -Message "Policy '$policyName' already exists, skipping creation" -Type Warning
            $policies += @{ name = $policyName; id = "existing" }
        }
        else {
            try {
                $body = @{
                    name = $policyName
                    description = "Microsoft Office update settings"
                    platforms = "windows10"
                    technologies = "mdm"
                    settings = @(
                        @{
                            id = "0"
                            settingInstance = @{
                                "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                settingDefinitionId = "device_vendor_msft_policy_config_office16v2~policy~l_microsoftofficemachine~l_updates_l_enableautomaticupdates"
                                choiceSettingValue = @{
                                    value = "device_vendor_msft_policy_config_office16v2~policy~l_microsoftofficemachine~l_updates_l_enableautomaticupdates_1"
                                    children = @()
                                }
                            }
                        }
                    )
                }
                
                $result = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies" -Body $body
                Write-LogMessage -Message "Created Office Updates policy" -Type Success
                $policies += $result
            }
            catch {
                Write-LogMessage -Message "Failed to create Office policies - $($_.Exception.Message)" -Type Error
            }
        }
        
        # ===================================================================
        # POLICY 14: OUTLOOK POLICY
        # ===================================================================
        Write-LogMessage -Message "Creating Outlook configuration policy..." -Type Info
        
        $policyName = "Outlook Configuration"
        if (Test-PolicyExists -PolicyName $policyName) {
            Write-LogMessage -Message "Policy '$policyName' already exists, skipping creation" -Type Warning
            $policies += @{ name = $policyName; id = "existing" }
        }
        else {
            try {
                $body = @{
                    name = $policyName
                    description = "Microsoft Outlook user experience settings"
                    platforms = "windows10"
                    technologies = "mdm"
                    settings = @(
                        @{
                            id = "0"
                            settingInstance = @{
                                "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                settingDefinitionId = "user_vendor_msft_policy_config_office16v2~policy~l_microsoftofficesystem~l_languagesettings~l_other_l_disablecomingsoon"
                                choiceSettingValue = @{
                                    value = "user_vendor_msft_policy_config_office16v2~policy~l_microsoftofficesystem~l_languagesettings~l_other_l_disablecomingsoon_1"
                                    children = @()
                                }
                            }
                        }
                    )
                }
                
                $result = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies" -Body $body
                Write-LogMessage -Message "Created Outlook policy" -Type Success
                $policies += $result
            }
            catch {
                Write-LogMessage -Message "Failed to create Outlook policy - $($_.Exception.Message)" -Type Error
            }
        }
        
        # ===================================================================
        # POLICY 15: DISABLE UAC FOR QUICKASSIST
        # ===================================================================
        Write-LogMessage -Message "Creating Disable UAC for QuickAssist policy..." -Type Info
        
        $policyName = "Disable UAC for Quickassist"
        if (Test-PolicyExists -PolicyName $policyName) {
            Write-LogMessage -Message "Policy '$policyName' already exists, skipping creation" -Type Warning
            $policies += @{ name = $policyName; id = "existing" }
        }
        else {
            try {
                $body = @{
                    name = $policyName
                    description = "Disable UAC secure desktop prompt for QuickAssist"
                    platforms = "windows10"
                    technologies = "mdm"
                    settings = @(
                        @{
                            id = "0"
                            settingInstance = @{
                                "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                settingDefinitionId = "device_vendor_msft_policy_config_localpoliciessecurityoptions_useraccountcontrol_switchtothesecuredesktopwhenpromptingforelevation"
                                choiceSettingValue = @{
                                    value = "device_vendor_msft_policy_config_localpoliciessecurityoptions_useraccountcontrol_switchtothesecuredesktopwhenpromptingforelevation_0"
                                    children = @()
                                }
                            }
                        }
                    )
                }
                
                $result = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies" -Body $body
                Write-LogMessage -Message "Created Disable UAC for QuickAssist policy" -Type Success
                $policies += $result
            }
            catch {
                Write-LogMessage -Message "Failed to create Disable UAC policy - $($_.Exception.Message)" -Type Error
            }
        }
        
        # Show EDR note instead of trying to create
        Write-LogMessage -Message "EDR Policy requires manual enablement:" -Type Warning
        Write-LogMessage -Message "1. Go to https://security.microsoft.com" -Type Info
        Write-LogMessage -Message "2. Navigate to Settings > Endpoints > Device management > Onboarding" -Type Info
        Write-LogMessage -Message "3. Enable Microsoft Defender for Business" -Type Info
        Write-LogMessage -Message "4. Configure the security connector" -Type Info
        
        # ===================================================================
        # 4. DIRECT EXECUTION: CREATE ALL COMPLIANCE POLICIES
        # ===================================================================
        Write-LogMessage -Message "Starting compliance policy creation..." -Type Info
        $compliancePolicies = @()
        
        # ===================================================================
        # COMPLIANCE POLICY 1: WINDOWS 10/11 COMPLIANCE POLICY (WITH REQUIRED ACTIONS)
        # ===================================================================
        Write-LogMessage -Message "Creating Windows 10/11 compliance policy..." -Type Info
        
        $windowsPolicyName = "Windows 10/11 compliance policy"
        if (Test-CompliancePolicyExists -PolicyName $windowsPolicyName) {
            Write-LogMessage -Message "Policy '$windowsPolicyName' already exists, skipping creation" -Type Warning
            $compliancePolicies += @{ displayName = $windowsPolicyName; id = "existing" }
        }
        else {
            try {
                # Windows compliance policy with required scheduledActionsForRule
                $body = @{
                    "@odata.type" = "#microsoft.graph.windows10CompliancePolicy"
                    displayName = $windowsPolicyName
                    description = "Standard Windows device compliance requirements"
                    bitLockerEnabled = $true
                    antivirusRequired = $true
                    deviceThreatProtectionEnabled = $false
                    deviceThreatProtectionRequiredSecurityLevel = "unavailable"
                    scheduledActionsForRule = @(
                        @{
                            ruleName = "PasswordRequired"
                            scheduledActionConfigurations = @(
                                @{
                                    actionType = "block"
                                    gracePeriodHours = 72
                                    notificationTemplateId = ""
                                }
                            )
                        }
                    )
                }
                
                $result = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/deviceManagement/deviceCompliancePolicies" -Body $body
                Write-LogMessage -Message "Created Windows 10/11 compliance policy with BitLocker and antivirus requirements" -Type Success
                $compliancePolicies += $result
            }
            catch {
                Write-LogMessage -Message "Failed to create Windows compliance policy - $($_.Exception.Message)" -Type Error
                Write-LogMessage -Message "Response details: $($_.Exception.Response)" -Type Error
                try {
                    $errorResponse = $_.Exception.Response.Content.ReadAsStringAsync().Result
                    Write-LogMessage -Message "Error response body: $errorResponse" -Type Error
                }
                catch {
                    Write-LogMessage -Message "Could not read error response details" -Type Warning
                }
            }
        }
        
        # ===================================================================
        # COMPLIANCE POLICY 2: ANDROID COMPLIANCE POLICY (FIXED PROPERTIES)
        # ===================================================================
        Write-LogMessage -Message "Creating Android compliance policy..." -Type Info
        
        $androidPolicyName = "Android Compliance Policy"
        if (Test-CompliancePolicyExists -PolicyName $androidPolicyName) {
            Write-LogMessage -Message "Policy '$androidPolicyName' already exists, skipping creation" -Type Warning
            $compliancePolicies += @{ displayName = $androidPolicyName; id = "existing" }
        }
        else {
            try {
                # Simplified Android compliance policy - remove problematic properties
                $body = @{
                    "@odata.type" = "#microsoft.graph.androidCompliancePolicy"
                    displayName = $androidPolicyName
                    description = "Android Enterprise compliance policy"
                    passwordRequired = $true
                    passwordMinimumLength = 6
                    securityBlockJailbrokenDevices = $true
                    storageRequireEncryption = $true
                    deviceThreatProtectionEnabled = $false
                    deviceThreatProtectionRequiredSecurityLevel = "unavailable"
                    scheduledActionsForRule = @(
                        @{
                            ruleName = "PasswordRequired"
                            scheduledActionConfigurations = @(
                                @{
                                    actionType = "block"
                                    gracePeriodHours = 72
                                    notificationTemplateId = ""
                                }
                            )
                        }
                    )
                }
                
                $result = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/deviceManagement/deviceCompliancePolicies" -Body $body
                Write-LogMessage -Message "Created Android compliance policy with password and encryption requirements" -Type Success
                $compliancePolicies += $result
            }
            catch {
                Write-LogMessage -Message "Failed to create Android compliance policy - $($_.Exception.Message)" -Type Error
                
                # Try with even more basic Android policy
                try {
                    Write-LogMessage -Message "Trying basic Android compliance policy..." -Type Info
                    $basicBody = @{
                        "@odata.type" = "#microsoft.graph.androidCompliancePolicy"
                        displayName = $androidPolicyName
                        description = "Basic Android compliance policy"
                        passwordRequired = $true
                        securityBlockJailbrokenDevices = $true
                        deviceThreatProtectionEnabled = $false
                        deviceThreatProtectionRequiredSecurityLevel = "unavailable"
                        scheduledActionsForRule = @(
                            @{
                                ruleName = "PasswordRequired"
                                scheduledActionConfigurations = @(
                                    @{
                                        actionType = "block"
                                        gracePeriodHours = 72
                                        notificationTemplateId = ""
                                    }
                                )
                            }
                        )
                    }
                    
                    $result = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/deviceManagement/deviceCompliancePolicies" -Body $basicBody
                    Write-LogMessage -Message "Created basic Android compliance policy" -Type Success
                    $compliancePolicies += $result
                }
                catch {
                    Write-LogMessage -Message "Failed to create basic Android compliance policy - $($_.Exception.Message)" -Type Error
                }
            }
        }
        
        # ===================================================================
        # COMPLIANCE POLICY 3: IOS COMPLIANCE POLICY (WITH REQUIRED ACTIONS)
        # ===================================================================
        Write-LogMessage -Message "Creating iOS compliance policy..." -Type Info
        
        $iosPolicyName = "iOS Compliance Policy"
        if (Test-CompliancePolicyExists -PolicyName $iosPolicyName) {
            Write-LogMessage -Message "Policy '$iosPolicyName' already exists, skipping creation" -Type Warning
            $compliancePolicies += @{ displayName = $iosPolicyName; id = "existing" }
        }
        else {
            try {
                # iOS compliance policy with required scheduledActionsForRule
                $body = @{
                    "@odata.type" = "#microsoft.graph.iosCompliancePolicy"
                    displayName = $iosPolicyName
                    description = "iOS compliance policy"
                    passcodeRequired = $true
                    passcodeMinimumLength = 6
                    passcodeRequiredType = "numeric"
                    deviceThreatProtectionEnabled = $false
                    deviceThreatProtectionRequiredSecurityLevel = "unavailable"
                    securityBlockJailbrokenDevices = $true
                    scheduledActionsForRule = @(
                        @{
                            ruleName = "PasswordRequired"
                            scheduledActionConfigurations = @(
                                @{
                                    actionType = "block"
                                    gracePeriodHours = 72
                                    notificationTemplateId = ""
                                }
                            )
                        }
                    )
                }
                
                $result = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/deviceManagement/deviceCompliancePolicies" -Body $body
                Write-LogMessage -Message "Created iOS compliance policy with passcode and jailbreak protection" -Type Success
                $compliancePolicies += $result
            }
            catch {
                Write-LogMessage -Message "Failed to create iOS compliance policy - $($_.Exception.Message)" -Type Error
            }
        }
        
        # ===================================================================
        # COMPLIANCE POLICY 4: MACOS COMPLIANCE POLICY (WITH REQUIRED ACTIONS)
        # ===================================================================
        Write-LogMessage -Message "Creating macOS compliance policy..." -Type Info
        
        $macosPolicyName = "MacOS Compliance"
        if (Test-CompliancePolicyExists -PolicyName $macosPolicyName) {
            Write-LogMessage -Message "Policy '$macosPolicyName' already exists, skipping creation" -Type Warning
            $compliancePolicies += @{ displayName = $macosPolicyName; id = "existing" }
        }
        else {
            try {
                # macOS compliance policy with required scheduledActionsForRule
                $body = @{
                    "@odata.type" = "#microsoft.graph.macOSCompliancePolicy"
                    displayName = $macosPolicyName
                    description = "MacOS compliance policy"
                    passwordRequired = $true
                    systemIntegrityProtectionEnabled = $true
                    firewallEnabled = $true
                    deviceThreatProtectionEnabled = $false
                    deviceThreatProtectionRequiredSecurityLevel = "unavailable"
                    scheduledActionsForRule = @(
                        @{
                            ruleName = "PasswordRequired"
                            scheduledActionConfigurations = @(
                                @{
                                    actionType = "block"
                                    gracePeriodHours = 72
                                    notificationTemplateId = ""
                                }
                            )
                        }
                    )
                }
                
                $result = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/deviceManagement/deviceCompliancePolicies" -Body $body
                Write-LogMessage -Message "Created macOS compliance policy with system integrity and firewall requirements" -Type Success
                $compliancePolicies += $result
            }
            catch {
                Write-LogMessage -Message "Failed to create macOS compliance policy - $($_.Exception.Message)" -Type Error
            }
        }
        
        Write-LogMessage -Message "Compliance policy creation completed" -Type Success

        # ===================================================================
        # 5. DIRECT EXECUTION: POLICY ASSIGNMENT TO AUTOPILOT GROUP
        # ===================================================================
        
        # Separate newly created policies from existing ones
        $newPolicies = $policies | Where-Object { $_ -and $_.id -and $_.id -ne "existing" }
        $existingPolicyNames = ($policies | Where-Object { $_ -and $_.id -eq "existing" }).name
        
        # Get WindowsAutoPilot group ID
        if (-not $script:TenantState.CreatedGroups.ContainsKey("WindowsAutoPilot")) {
            Write-LogMessage -Message "WindowsAutoPilot group not found, cannot assign policies" -Type Error
            return $false
        }

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

        Write-LogMessage -Message "Assigning compliance policies to platform-specific groups..." -Type Info
        
        # Use the SAME assignment method that works for configuration policies (with fallback)
        
        # Assign Windows compliance policy to WindowsAutoPilot group
        $windowsCompliancePolicy = $compliancePolicies | Where-Object { $_.displayName -eq "Windows 10/11 compliance policy" -and $_.id -ne "existing" }
        if ($windowsCompliancePolicy -and $script:TenantState.CreatedGroups.ContainsKey("WindowsAutoPilot")) {
            $autoPilotGroupId = $script:TenantState.CreatedGroups["WindowsAutoPilot"]
            try {
                # Try primary method first (same as configuration policies)
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
                
                Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/deviceManagement/deviceCompliancePolicies/$($windowsCompliancePolicy.id)/assignments" -Body $body
                Write-LogMessage -Message "Assigned Windows compliance policy to WindowsAutoPilot group" -Type Success
            }
            catch {
                # Try the assign action endpoint as fallback (same as configuration policies)
                try {
                    $assignUri = "https://graph.microsoft.com/beta/deviceManagement/deviceCompliancePolicies/$($windowsCompliancePolicy.id)/assign"
                    Invoke-MgGraphRequest -Method POST -Uri $assignUri -Body $body
                    Write-LogMessage -Message "Assigned Windows compliance policy to WindowsAutoPilot group (using assign action)" -Type Success
                }
                catch {
                    Write-LogMessage -Message "Failed to assign Windows compliance policy: $($_.Exception.Message)" -Type Warning
                }
            }
        }
        
        # Assign Android compliance policy to Android Devices group
        $androidCompliancePolicy = $compliancePolicies | Where-Object { $_.displayName -eq "Android Compliance Policy" -and $_.id -ne "existing" }
        if ($androidCompliancePolicy -and $script:TenantState.CreatedGroups.ContainsKey("AndroidDevices")) {
            $androidGroupId = $script:TenantState.CreatedGroups["AndroidDevices"]
            try {
                # Try primary method first
                $body = @{
                    assignments = @(
                        @{
                            target = @{
                                "@odata.type" = "#microsoft.graph.groupAssignmentTarget"
                                groupId = $androidGroupId
                            }
                        }
                    )
                }
                
                Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/deviceManagement/deviceCompliancePolicies/$($androidCompliancePolicy.id)/assignments" -Body $body
                Write-LogMessage -Message "Assigned Android compliance policy to Android Devices group" -Type Success
            }
            catch {
                # Try the assign action endpoint as fallback
                try {
                    $assignUri = "https://graph.microsoft.com/beta/deviceManagement/deviceCompliancePolicies/$($androidCompliancePolicy.id)/assign"
                    Invoke-MgGraphRequest -Method POST -Uri $assignUri -Body $body
                    Write-LogMessage -Message "Assigned Android compliance policy to Android Devices group (using assign action)" -Type Success
                }
                catch {
                    Write-LogMessage -Message "Failed to assign Android compliance policy: $($_.Exception.Message)" -Type Warning
                }
            }
        }
        
        # Assign macOS compliance policy to MacOS Devices group  
        $macosCompliancePolicy = $compliancePolicies | Where-Object { $_.displayName -eq "MacOS Compliance" -and $_.id -ne "existing" }
        if ($macosCompliancePolicy -and $script:TenantState.CreatedGroups.ContainsKey("MacOSDevices")) {
            $macosGroupId = $script:TenantState.CreatedGroups["MacOSDevices"]
            try {
                # Try primary method first
                $body = @{
                    assignments = @(
                        @{
                            target = @{
                                "@odata.type" = "#microsoft.graph.groupAssignmentTarget"
                                groupId = $macosGroupId
                            }
                        }
                    )
                }
                
                Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/deviceManagement/deviceCompliancePolicies/$($macosCompliancePolicy.id)/assignments" -Body $body
                Write-LogMessage -Message "Assigned macOS compliance policy to MacOS Devices group" -Type Success
            }
            catch {
                # Try the assign action endpoint as fallback
                try {
                    $assignUri = "https://graph.microsoft.com/beta/deviceManagement/deviceCompliancePolicies/$($macosCompliancePolicy.id)/assign"
                    Invoke-MgGraphRequest -Method POST -Uri $assignUri -Body $body
                    Write-LogMessage -Message "Assigned macOS compliance policy to MacOS Devices group (using assign action)" -Type Success
                }
                catch {
                    Write-LogMessage -Message "Failed to assign macOS compliance policy: $($_.Exception.Message)" -Type Warning
                }
            }
        }
        
        # Assign iOS compliance policy to iOS Devices group
        $iosCompliancePolicy = $compliancePolicies | Where-Object { $_.displayName -eq "iOS Compliance Policy" -and $_.id -ne "existing" }
        if ($iosCompliancePolicy -and $script:TenantState.CreatedGroups.ContainsKey("iOSDevices")) {
            $iosGroupId = $script:TenantState.CreatedGroups["iOSDevices"]
            try {
                # Try primary method first
                $body = @{
                    assignments = @(
                        @{
                            target = @{
                                "@odata.type" = "#microsoft.graph.groupAssignmentTarget"
                                groupId = $iosGroupId
                            }
                        }
                    )
                }
                
                Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/deviceManagement/deviceCompliancePolicies/$($iosCompliancePolicy.id)/assignments" -Body $body
                Write-LogMessage -Message "Assigned iOS compliance policy to iOS Devices group" -Type Success
            }
            catch {
                # Try the assign action endpoint as fallback
                try {
                    $assignUri = "https://graph.microsoft.com/beta/deviceManagement/deviceCompliancePolicies/$($iosCompliancePolicy.id)/assign"
                    Invoke-MgGraphRequest -Method POST -Uri $assignUri -Body $body
                    Write-LogMessage -Message "Assigned iOS compliance policy to iOS Devices group (using assign action)" -Type Success
                }
                catch {
                    Write-LogMessage -Message "Failed to assign iOS compliance policy: $($_.Exception.Message)" -Type Warning
                }
            }
        }

        # Handle existing policies if update mode is enabled
        if ($existingPolicyNames.Count -gt 0 -and $UpdateExistingPolicies) {
            Write-LogMessage -Message "Updating assignments for $($existingPolicyNames.Count) existing policies..." -Type Info
            Write-LogMessage -Message "Existing policies: $($existingPolicyNames -join ', ')" -Type Info
            
            # Get all existing configuration policies
            $allPoliciesResponse = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies"
            $existingPolicies = $allPoliciesResponse.value | Where-Object { $_.name -in $existingPolicyNames }
            
            foreach ($policy in $existingPolicies) {
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
                    Write-LogMessage -Message "Assigned existing policy '$($policy.name)' to WindowsAutoPilot group" -Type Success
                }
                catch {
                    # Try the assign action endpoint as fallback
                    try {
                        $assignUri = "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies/$($policy.id)/assign"
                        Invoke-MgGraphRequest -Method POST -Uri $assignUri -Body $body
                        Write-LogMessage -Message "Assigned existing policy '$($policy.name)' to WindowsAutoPilot group (using assign action)" -Type Success
                    }
                    catch {
                        Write-LogMessage -Message "Failed to assign existing policy '$($policy.name)': $($_.Exception.Message)" -Type Warning
                    }
                }
            }
            
            # Also handle existing compliance policies using the SAME method as configuration policies (with fallback)
            $existingCompliancePolicyNames = ($compliancePolicies | Where-Object { $_ -and $_.id -eq "existing" }).displayName
            if ($existingCompliancePolicyNames.Count -gt 0) {
                Write-LogMessage -Message "Updating assignments for existing compliance policies..." -Type Info
                $allCompliancePoliciesResponse = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/deviceManagement/deviceCompliancePolicies"
                $existingCompliancePolicies = $allCompliancePoliciesResponse.value | Where-Object { $_.displayName -in $existingCompliancePolicyNames }
                
                foreach ($policy in $existingCompliancePolicies) {
                    # Determine which group to assign based on policy name
                    $targetGroupId = $null
                    $targetGroupName = ""
                    
                    switch ($policy.displayName) {
                        "Windows 10/11 compliance policy" {
                            if ($script:TenantState.CreatedGroups.ContainsKey("WindowsAutoPilot")) {
                                $targetGroupId = $script:TenantState.CreatedGroups["WindowsAutoPilot"]
                                $targetGroupName = "WindowsAutoPilot"
                            }
                        }
                        "Android Compliance Policy" {
                            if ($script:TenantState.CreatedGroups.ContainsKey("AndroidDevices")) {
                                $targetGroupId = $script:TenantState.CreatedGroups["AndroidDevices"]
                                $targetGroupName = "Android Devices"
                            }
                        }
                        "MacOS Compliance" {
                            if ($script:TenantState.CreatedGroups.ContainsKey("MacOSDevices")) {
                                $targetGroupId = $script:TenantState.CreatedGroups["MacOSDevices"]
                                $targetGroupName = "MacOS Devices"
                            }
                        }
                        "iOS Compliance Policy" {
                            if ($script:TenantState.CreatedGroups.ContainsKey("iOSDevices")) {
                                $targetGroupId = $script:TenantState.CreatedGroups["iOSDevices"]
                                $targetGroupName = "iOS Devices"
                            }
                        }
                    }
                    
                    if ($targetGroupId) {
                        try {
                            # Try primary method first (same as configuration policies)
                            $body = @{
                                assignments = @(
                                    @{
                                        target = @{
                                            "@odata.type" = "#microsoft.graph.groupAssignmentTarget"
                                            groupId = $targetGroupId
                                        }
                                    }
                                )
                            }
                            
                            Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/deviceManagement/deviceCompliancePolicies/$($policy.id)/assignments" -Body $body
                            Write-LogMessage -Message "Successfully assigned existing compliance policy '$($policy.displayName)' to $targetGroupName group" -Type Success
                        }
                        catch {
                            # Try the assign action endpoint as fallback (same as configuration policies)
                            try {
                                $assignUri = "https://graph.microsoft.com/beta/deviceManagement/deviceCompliancePolicies/$($policy.id)/assign"
                                Invoke-MgGraphRequest -Method POST -Uri $assignUri -Body $body
                                Write-LogMessage -Message "Successfully assigned existing compliance policy '$($policy.displayName)' to $targetGroupName group (using assign action)" -Type Success
                            }
                            catch {
                                Write-LogMessage -Message "Failed to assign existing compliance policy '$($policy.displayName)': $($_.Exception.Message)" -Type Warning
                            }
                        }
                    }
                    else {
                        Write-LogMessage -Message "No target group found for existing compliance policy '$($policy.displayName)'" -Type Warning
                    }
                }
            }
        }
        elseif ($existingPolicyNames.Count -gt 0 -and -not $UpdateExistingPolicies) {
            Write-LogMessage -Message "Found $($existingPolicyNames.Count) existing policies, but update mode is disabled" -Type Warning
            Write-LogMessage -Message "Existing policies: $($existingPolicyNames -join ', ')" -Type Info
        }
        
        Write-LogMessage -Message "Intune configuration completed successfully" -Type Success
        return $true
    }
    catch {
        Write-LogMessage -Message "Error in Intune configuration - $($_.Exception.Message)" -Type Error
        return $false
    }
}